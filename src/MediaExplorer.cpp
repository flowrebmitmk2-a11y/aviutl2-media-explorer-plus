//----------------------------------------------------------------------------------
//  PreviewExplorer - IExplorerBrowser 実装
//  left pane: AviUtl2 preview embedding (Phase 1)
//  right pane: IExplorerBrowser + 2-tier tabs with right-click menus
//----------------------------------------------------------------------------------
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <dwmapi.h>
#include <string>
#include <vector>
#include <algorithm>
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "dwmapi.lib")

#include "plugin2.h"
#include "logger2.h"
#include "config2.h"

//------------------------------------------------------------------------------
// データ構造
//------------------------------------------------------------------------------
struct SmallTab { std::wstring name, path; };
struct BigTab   { std::wstring name; std::vector<SmallTab> smallTabs; int smallTabIdx = 0; };

//------------------------------------------------------------------------------
// グローバル
//------------------------------------------------------------------------------
EDIT_HANDLE*   g_editHandle = nullptr;
LOG_HANDLE*    g_logger     = nullptr;
CONFIG_HANDLE* g_config     = nullptr;

HWND g_hwndMain     = nullptr;  // 登録ウィンドウ
HWND g_hwndLeft     = nullptr;  // 左ペイン（プレビューコンテナ）
HWND g_hwndRight    = nullptr;  // 右ペイン（エクスプローラー）
HWND g_hwndInfo     = nullptr;  // 左ペイン内テキスト（診断ログ表示）
HWND g_hwndEmbedded = nullptr;  // 組み込み試行したプレビューウィンドウ

// 右ペイン内コントロール
HWND g_hwndBigTab   = nullptr;  // ID=101
HWND g_hwndSmallTab = nullptr;  // ID=102

// IExplorerBrowser
static IExplorerBrowser* g_peb = nullptr;

// タブデータ
std::vector<BigTab> g_bigTabs;
int g_bigTabIdx = 0;

// スプリッタ
constexpr int SPLITTER_W  = 5;
constexpr int MIN_LEFT_W  = 100;
constexpr int MIN_RIGHT_W = 200;
int  g_splitX   = 400;
bool g_dragging = false;

// タイマー
constexpr UINT_PTR TIMER_FIND_PREVIEW = 1;

// ウィンドウクラス名
constexpr wchar_t WC_MAIN[]     = L"PreviewExplorer_Main";
constexpr wchar_t WC_RIGHT[]    = L"PreviewExplorer_Right";
constexpr wchar_t WC_SETTINGS[] = L"PreviewExplorer_Settings";
constexpr wchar_t WC_INPUTDLG[] = L"PreviewExplorer_InputDlg";

// レイアウト定数（右ペイン内）
constexpr int BIG_TAB_H   = 36;
constexpr int SMALL_TAB_H = 24;

// フォント
static HFONT g_hFontBig    = nullptr;
static HFONT g_hFontNormal = nullptr;

// テーマカラー (CONFIG_HANDLE から取得)
static COLORREF g_clrBg   = RGB(0x30, 0x30, 0x30);  // デフォルト: ダークグレー
static COLORREF g_clrText = RGB(0xFF, 0xFF, 0xFF);  // デフォルト: 白
static HBRUSH   g_hBrushBg = nullptr;

// 設定ウィンドウ用前方宣言
static HWND g_hwndSettings = nullptr;

//------------------------------------------------------------------------------
// ログユーティリティ
//------------------------------------------------------------------------------
static void Log(const wchar_t* msg) {
    if (g_logger) g_logger->log(g_logger, msg);
}

//------------------------------------------------------------------------------
// テーマ初期化 (InitializeConfig → RegisterPlugin の順で呼ばれるため
//               RegisterPlugin 冒頭で一度だけ呼ぶ)
//------------------------------------------------------------------------------
static void InitTheme() {
    // フォント情報 (style.conf [Font] の "Control" キー)
    LPCWSTR fontName = L"Yu Gothic UI";
    float   fontSize = 13.0f;
    if (g_config) {
        FONT_INFO* fi = g_config->get_font_info(g_config, "Control");
        if (fi && fi->name && fi->name[0]) {
            fontName = fi->name;
            if (fi->size > 0.0f) fontSize = fi->size;
        }
    }
    int logH = -MulDiv((int)fontSize, 96, 72);  // pt → logical units (96 DPI)

    // 色情報 (style.conf [Color] の各キー)
    // API は COLORREF (0x00BBGGRR) 形式で返す想定、0 は未取得
    if (g_config) {
        int raw = g_config->get_color_code(g_config, "Background");
        if (raw) g_clrBg = (COLORREF)raw;
        raw = g_config->get_color_code(g_config, "Text");
        if (raw) g_clrText = (COLORREF)raw;
    }

    // ブラシ再作成
    if (g_hBrushBg) { DeleteObject(g_hBrushBg); g_hBrushBg = nullptr; }
    g_hBrushBg = CreateSolidBrush(g_clrBg);

    // フォント再作成
    if (g_hFontBig)    { DeleteObject(g_hFontBig);    g_hFontBig    = nullptr; }
    if (g_hFontNormal) { DeleteObject(g_hFontNormal); g_hFontNormal = nullptr; }
    g_hFontBig    = CreateFontW(logH - 1, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName);
    g_hFontNormal = CreateFontW(logH - 3, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName);
}

// ダークテーマかどうか判定 (輝度 < 128 ならダーク)
static bool IsDarkTheme() {
    BYTE r = GetRValue(g_clrBg), g2 = GetGValue(g_clrBg), b = GetBValue(g_clrBg);
    return (299 * r + 587 * g2 + 114 * b) / 1000 < 128;
}

// ウィンドウにダークモードを適用 (Windows 10 1809+)
static void TrySetDarkMode(HWND hwnd) {
    BOOL dark = IsDarkTheme() ? TRUE : FALSE;
    DwmSetWindowAttribute(hwnd, 20 /* DWMWA_USE_IMMERSIVE_DARK_MODE */, &dark, sizeof(dark));
}

//------------------------------------------------------------------------------
// INI 状態保存
//------------------------------------------------------------------------------
static std::wstring GetIniPath() {
    std::wstring base = g_config ? g_config->app_data_path : L".";
    return base + L"\\Plugin\\PreviewExplorer\\state.ini";
}

static void EnsureDir(const std::wstring& path) {
    auto pos = path.rfind(L'\\');
    if (pos == std::wstring::npos) return;
    std::wstring parent = path.substr(0, pos);
    EnsureDir(parent);
    CreateDirectoryW(parent.c_str(), nullptr);
}

static void SaveState();
static void LoadState();

//------------------------------------------------------------------------------
// IExplorerBrowser ナビゲーション
//------------------------------------------------------------------------------
static void NavigateExplorerToCurrentTab() {
    if (!g_peb) return;
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    auto& bt = g_bigTabs[g_bigTabIdx];
    if (bt.smallTabIdx < 0 || bt.smallTabIdx >= (int)bt.smallTabs.size()) return;
    const std::wstring& path = bt.smallTabs[bt.smallTabIdx].path;

    if (path.empty() || path == L"QuickAccess") {
        // Windows クイックアクセスへ移動
        PIDLIST_ABSOLUTE pidl = nullptr;
        // FOLDERID_QuickAccess (Windows 10+) を手動定義
        const GUID FOLDERID_QA = {0x679f85cb, 0x0220, 0x4080,
            {0xb2, 0x9b, 0x55, 0x40, 0xcc, 0x05, 0xaa, 0xb6}};
        if (SUCCEEDED(SHGetKnownFolderIDList(FOLDERID_QA, 0, nullptr, &pidl))) {
            g_peb->BrowseToIDList(pidl, SBSP_ABSOLUTE);
            CoTaskMemFree(pidl);
        }
    } else {
        IShellItem* psi = nullptr;
        if (SUCCEEDED(SHCreateItemFromParsingName(path.c_str(), nullptr, IID_PPV_ARGS(&psi)))) {
            g_peb->BrowseToObject(psi, SBSP_ABSOLUTE);
            psi->Release();
        }
    }
}

//------------------------------------------------------------------------------
// 現在の IExplorerBrowser の場所をパス文字列として取得
//------------------------------------------------------------------------------
static std::wstring GetCurrentExplorerPath() {
    if (!g_peb) return L"";
    std::wstring result;
    IShellView* psv = nullptr;
    if (SUCCEEDED(g_peb->GetCurrentView(IID_PPV_ARGS(&psv)))) {
        IFolderView2* pfv2 = nullptr;
        if (SUCCEEDED(psv->QueryInterface(IID_PPV_ARGS(&pfv2)))) {
            IPersistFolder2* ppf2 = nullptr;
            if (SUCCEEDED(pfv2->GetFolder(IID_PPV_ARGS(&ppf2)))) {
                PIDLIST_ABSOLUTE pidl = nullptr;
                if (SUCCEEDED(ppf2->GetCurFolder(&pidl))) {
                    wchar_t path[MAX_PATH]{};
                    SHGetPathFromIDListW(pidl, path);
                    result = path;
                    CoTaskMemFree(pidl);
                }
                ppf2->Release();
            }
            pfv2->Release();
        }
        psv->Release();
    }
    return result;
}

//------------------------------------------------------------------------------
// スモールタブコントロールを再構築
//------------------------------------------------------------------------------
static void RebuildSmallTabs() {
    if (!g_hwndSmallTab) return;
    TabCtrl_DeleteAllItems(g_hwndSmallTab);
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    auto& bt = g_bigTabs[g_bigTabIdx];
    for (int i = 0; i < (int)bt.smallTabs.size(); i++) {
        TCITEMW tci{};
        tci.mask    = TCIF_TEXT;
        tci.pszText = const_cast<wchar_t*>(bt.smallTabs[i].name.c_str());
        TabCtrl_InsertItem(g_hwndSmallTab, i, &tci);
    }
    if (bt.smallTabIdx >= (int)bt.smallTabs.size())
        bt.smallTabIdx = (int)bt.smallTabs.size() - 1;
    if (bt.smallTabIdx < 0) bt.smallTabIdx = 0;
    if (!bt.smallTabs.empty())
        TabCtrl_SetCurSel(g_hwndSmallTab, bt.smallTabIdx);
}

//------------------------------------------------------------------------------
// 右ペインのレイアウト
//------------------------------------------------------------------------------
static void DoRightLayout(HWND hwnd) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    int W = rc.right, H = rc.bottom;
    int y = 0;

    if (g_hwndBigTab) {
        SetWindowPos(g_hwndBigTab, nullptr, 0, y, W, BIG_TAB_H, SWP_NOZORDER | SWP_NOACTIVATE);
        y += BIG_TAB_H;
    }
    if (g_hwndSmallTab) {
        SetWindowPos(g_hwndSmallTab, nullptr, 0, y, W, SMALL_TAB_H, SWP_NOZORDER | SWP_NOACTIVATE);
        y += SMALL_TAB_H;
    }

    if (g_peb) {
        RECT ebRect = {0, y, W, H};
        g_peb->SetRect(nullptr, ebRect);
    }
}

//==============================================================================
// 入力ダイアログ（インライン実装）
//==============================================================================
struct InputDlgData {
    const wchar_t* title;
    const wchar_t* prompt;
    std::wstring   initial;
    std::wstring   result;
    bool           ok = false;
    HWND           hwndEdit   = nullptr;
    HWND           hwndOK     = nullptr;
    HWND           hwndCancel = nullptr;
};

static LRESULT CALLBACK InputDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lp);
        auto* d  = reinterpret_cast<InputDlgData*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(d));

        HINSTANCE hi = GetModuleHandleW(nullptr);
        // プロンプト
        CreateWindowExW(0, L"STATIC", d->prompt,
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            8, 8, 280, 20, hwnd, nullptr, hi, nullptr);
        // エディット
        d->hwndEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", d->initial.c_str(),
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            8, 32, 280, 22, hwnd, reinterpret_cast<HMENU>(1), hi, nullptr);
        // OK / Cancel
        d->hwndOK = CreateWindowExW(0, L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
            66, 62, 70, 24, hwnd, reinterpret_cast<HMENU>(IDOK), hi, nullptr);
        d->hwndCancel = CreateWindowExW(0, L"BUTTON", L"キャンセル",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            148, 62, 80, 24, hwnd, reinterpret_cast<HMENU>(IDCANCEL), hi, nullptr);

        SendMessageW(d->hwndEdit, EM_SETSEL, 0, -1);
        SetFocus(d->hwndEdit);
        return 0;
    }
    case WM_COMMAND: {
        auto* d = reinterpret_cast<InputDlgData*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        if (!d) break;
        int id = LOWORD(wp);
        if (id == IDOK) {
            wchar_t buf[512]{};
            GetWindowTextW(d->hwndEdit, buf, 512);
            d->result = buf;
            d->ok = true;
            DestroyWindow(hwnd);
        } else if (id == IDCANCEL) {
            d->ok = false;
            DestroyWindow(hwnd);
        }
        return 0;
    }
    case WM_KEYDOWN:
        if (wp == VK_ESCAPE) {
            auto* d = reinterpret_cast<InputDlgData*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if (d) { d->ok = false; }
            DestroyWindow(hwnd);
        }
        return 0;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

static bool ShowInputDialog(HWND parent, const wchar_t* title,
                            const wchar_t* prompt, std::wstring& result) {
    InputDlgData d;
    d.title   = title;
    d.prompt  = prompt;
    d.initial = result;

    HWND hdlg = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        WC_INPUTDLG, title,
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, 310, 102,
        parent, nullptr, GetModuleHandleW(nullptr), &d);

    if (!hdlg) return false;

    // ローカルメッセージループ
    MSG m{};
    while (IsWindow(hdlg) && GetMessageW(&m, nullptr, 0, 0)) {
        if (!IsDialogMessageW(hdlg, &m)) {
            TranslateMessage(&m);
            DispatchMessageW(&m);
        }
    }

    if (d.ok) {
        result = d.result;
        return true;
    }
    return false;
}

//==============================================================================
// ビッグタブ右クリック サブクラスプロシージャ
//==============================================================================
static LRESULT CALLBACK BigTabSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                           UINT_PTR, DWORD_PTR) {
    if (msg == WM_RBUTTONDOWN) {
        TCHITTESTINFO hti{};
        hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        int hitIdx = TabCtrl_HitTest(hwnd, &hti);

        HMENU hm = CreatePopupMenu();
        AppendMenuW(hm, MF_STRING, 3001, L"追加");
        if (hitIdx >= 0) {
            AppendMenuW(hm, MF_STRING, 3002, L"名前変更");
            AppendMenuW(hm, MF_STRING, 3003, L"削除");
        }
        POINT pt; GetCursorPos(&pt);
        int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON,
                                 pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(hm);

        if (cmd == 3001) {
            // 追加
            BigTab bt;
            bt.name = L"新規";
            SmallTab st; st.name = L"クイックアクセス"; st.path = L"";
            bt.smallTabs.push_back(st);
            g_bigTabs.push_back(bt);
            g_bigTabIdx = (int)g_bigTabs.size() - 1;

            TCITEMW tci{};
            tci.mask    = TCIF_TEXT;
            tci.pszText = const_cast<wchar_t*>(g_bigTabs[g_bigTabIdx].name.c_str());
            TabCtrl_InsertItem(hwnd, g_bigTabIdx, &tci);
            TabCtrl_SetCurSel(hwnd, g_bigTabIdx);
            RebuildSmallTabs();
            NavigateExplorerToCurrentTab();
            SaveState();

        } else if (cmd == 3002 && hitIdx >= 0) {
            // 名前変更
            std::wstring name = g_bigTabs[hitIdx].name;
            if (ShowInputDialog(GetParent(hwnd), L"名前変更", L"新しい名前:", name)) {
                g_bigTabs[hitIdx].name = name;
                TCITEMW tci{};
                tci.mask    = TCIF_TEXT;
                tci.pszText = const_cast<wchar_t*>(name.c_str());
                TabCtrl_SetItem(hwnd, hitIdx, &tci);
                SaveState();
            }

        } else if (cmd == 3003 && hitIdx >= 0) {
            // 削除
            if (g_bigTabs.size() <= 1) {
                MessageBoxW(hwnd, L"最後のタブは削除できません。", L"削除", MB_OK | MB_ICONWARNING);
                return 0;
            }
            g_bigTabs.erase(g_bigTabs.begin() + hitIdx);
            TabCtrl_DeleteItem(hwnd, hitIdx);
            if (g_bigTabIdx >= (int)g_bigTabs.size())
                g_bigTabIdx = (int)g_bigTabs.size() - 1;
            TabCtrl_SetCurSel(hwnd, g_bigTabIdx);
            RebuildSmallTabs();
            NavigateExplorerToCurrentTab();
            SaveState();
        }
        return 0;
    }
    return DefSubclassProc(hwnd, msg, wp, lp);
}

//==============================================================================
// スモールタブ右クリック サブクラスプロシージャ
//==============================================================================
static LRESULT CALLBACK SmallTabSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                             UINT_PTR, DWORD_PTR) {
    if (msg == WM_RBUTTONDOWN) {
        TCHITTESTINFO hti{};
        hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        int hitIdx = TabCtrl_HitTest(hwnd, &hti);

        if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return 0;
        auto& bt = g_bigTabs[g_bigTabIdx];

        HMENU hm = CreatePopupMenu();
        AppendMenuW(hm, MF_STRING, 4001, L"追加");
        if (hitIdx >= 0) {
            AppendMenuW(hm, MF_STRING, 4002, L"名前変更");
            AppendMenuW(hm, MF_STRING, 4003, L"パス変更...");
            AppendMenuW(hm, MF_STRING, 4004, L"現在の場所をパスに設定");
            AppendMenuW(hm, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(hm, MF_STRING, 4005, L"削除");
        }
        POINT pt; GetCursorPos(&pt);
        int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON,
                                 pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(hm);

        if (cmd == 4001) {
            // 追加
            SmallTab st; st.name = L"新規"; st.path = L"";
            bt.smallTabs.push_back(st);
            int newIdx = (int)bt.smallTabs.size() - 1;
            TCITEMW tci{};
            tci.mask    = TCIF_TEXT;
            tci.pszText = const_cast<wchar_t*>(st.name.c_str());
            TabCtrl_InsertItem(hwnd, newIdx, &tci);
            bt.smallTabIdx = newIdx;
            TabCtrl_SetCurSel(hwnd, newIdx);
            NavigateExplorerToCurrentTab();
            SaveState();

        } else if (cmd == 4002 && hitIdx >= 0) {
            // 名前変更
            std::wstring name = bt.smallTabs[hitIdx].name;
            if (ShowInputDialog(GetParent(hwnd), L"名前変更", L"新しい名前:", name)) {
                bt.smallTabs[hitIdx].name = name;
                TCITEMW tci{};
                tci.mask    = TCIF_TEXT;
                tci.pszText = const_cast<wchar_t*>(name.c_str());
                TabCtrl_SetItem(hwnd, hitIdx, &tci);
                SaveState();
            }

        } else if (cmd == 4003 && hitIdx >= 0) {
            // パス変更（SHBrowseForFolder）
            wchar_t disp[MAX_PATH]{};
            BROWSEINFOW bi{};
            bi.hwndOwner = GetParent(hwnd);
            bi.pszDisplayName = disp;
            bi.lpszTitle = L"フォルダを選択";
            bi.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
            if (pidl) {
                wchar_t path[MAX_PATH]{};
                SHGetPathFromIDListW(pidl, path);
                CoTaskMemFree(pidl);
                bt.smallTabs[hitIdx].path = path;
                // 現在のタブなら移動
                if (hitIdx == bt.smallTabIdx)
                    NavigateExplorerToCurrentTab();
                SaveState();
            }

        } else if (cmd == 4004 && hitIdx >= 0) {
            // 現在の IExplorerBrowser の場所をパスに設定
            std::wstring path = GetCurrentExplorerPath();
            if (!path.empty()) {
                bt.smallTabs[hitIdx].path = path;
                SaveState();
            }

        } else if (cmd == 4005 && hitIdx >= 0) {
            // 削除
            if (bt.smallTabs.size() <= 1) {
                MessageBoxW(hwnd, L"最後のタブは削除できません。", L"削除", MB_OK | MB_ICONWARNING);
                return 0;
            }
            bt.smallTabs.erase(bt.smallTabs.begin() + hitIdx);
            TabCtrl_DeleteItem(hwnd, hitIdx);
            if (bt.smallTabIdx >= (int)bt.smallTabs.size())
                bt.smallTabIdx = (int)bt.smallTabs.size() - 1;
            TabCtrl_SetCurSel(hwnd, bt.smallTabIdx);
            NavigateExplorerToCurrentTab();
            SaveState();
        }
        return 0;
    }
    return DefSubclassProc(hwnd, msg, wp, lp);
}

//==============================================================================
// 設定ウィンドウ
//==============================================================================
struct SettingsControls {
    HWND listBig       = nullptr;
    HWND editBigName   = nullptr;
    HWND btnBigAdd     = nullptr;
    HWND btnBigDel     = nullptr;
    HWND btnBigUp      = nullptr;
    HWND btnBigDown    = nullptr;
    HWND btnBigSave    = nullptr;
    HWND listSmall     = nullptr;
    HWND editSmallName = nullptr;
    HWND editSmallPath = nullptr;
    HWND btnSmallAdd   = nullptr;
    HWND btnSmallDel   = nullptr;
    HWND btnSmallUp    = nullptr;
    HWND btnSmallDown  = nullptr;
    HWND btnSmallBrowse= nullptr;
    HWND btnSmallSave  = nullptr;
    HWND btnClose      = nullptr;
};
static SettingsControls g_sc{};

static std::vector<BigTab> g_settingsTabs;
static int g_settBigIdx = -1;

static void SettingsRefreshBigList() {
    SendMessageW(g_sc.listBig, LB_RESETCONTENT, 0, 0);
    for (auto& bt : g_settingsTabs)
        SendMessageW(g_sc.listBig, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(bt.name.c_str()));
    if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size())
        SendMessageW(g_sc.listBig, LB_SETCURSEL, g_settBigIdx, 0);
}

static void SettingsRefreshSmallList() {
    SendMessageW(g_sc.listSmall, LB_RESETCONTENT, 0, 0);
    if (g_settBigIdx < 0 || g_settBigIdx >= (int)g_settingsTabs.size()) return;
    auto& bt = g_settingsTabs[g_settBigIdx];
    for (auto& st : bt.smallTabs)
        SendMessageW(g_sc.listSmall, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(st.name.c_str()));
}

static void SettingsApplyToMain() {
    g_bigTabs = g_settingsTabs;
    if (g_bigTabIdx >= (int)g_bigTabs.size())
        g_bigTabIdx = (int)g_bigTabs.size() - 1;
    if (g_bigTabIdx < 0) g_bigTabIdx = 0;

    if (g_hwndBigTab) {
        TabCtrl_DeleteAllItems(g_hwndBigTab);
        for (int i = 0; i < (int)g_bigTabs.size(); i++) {
            TCITEMW tci{};
            tci.mask    = TCIF_TEXT;
            tci.pszText = const_cast<wchar_t*>(g_bigTabs[i].name.c_str());
            TabCtrl_InsertItem(g_hwndBigTab, i, &tci);
        }
        if (!g_bigTabs.empty())
            TabCtrl_SetCurSel(g_hwndBigTab, g_bigTabIdx);
    }
    RebuildSmallTabs();
    NavigateExplorerToCurrentTab();
    SaveState();
}

static LRESULT CALLBACK SettingsWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        RECT rc;
        GetClientRect(hwnd, &rc);
        int W = rc.right, H = rc.bottom;
        int half = W / 2 - 8;
        int x1 = 4, x2 = W / 2 + 4;
        int y = 4;

        CreateWindowExW(0, L"STATIC", L"大カテゴリ",
            WS_CHILD | WS_VISIBLE, x1, y, half, 16,
            hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        y += 18;
        g_sc.listBig = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY,
            x1, y, half, 140,
            hwnd, reinterpret_cast<HMENU>(1001), GetModuleHandleW(nullptr), nullptr);
        y += 144;
        g_sc.btnBigAdd  = CreateWindowExW(0, L"BUTTON", L"追加",  WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x1,      y, 40, 22, hwnd, reinterpret_cast<HMENU>(1010), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnBigDel  = CreateWindowExW(0, L"BUTTON", L"削除",  WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x1+44,  y, 40, 22, hwnd, reinterpret_cast<HMENU>(1011), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnBigUp   = CreateWindowExW(0, L"BUTTON", L"↑",    WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x1+88,  y, 30, 22, hwnd, reinterpret_cast<HMENU>(1012), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnBigDown = CreateWindowExW(0, L"BUTTON", L"↓",    WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x1+122, y, 30, 22, hwnd, reinterpret_cast<HMENU>(1013), GetModuleHandleW(nullptr), nullptr);
        y += 26;
        CreateWindowExW(0, L"STATIC", L"名前:", WS_CHILD|WS_VISIBLE, x1, y, 36, 18, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        g_sc.editBigName = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL,
            x1+38, y, half-38-50, 20,
            hwnd, reinterpret_cast<HMENU>(1020), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnBigSave = CreateWindowExW(0, L"BUTTON", L"保存",  WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
            x1+half-46, y, 46, 22, hwnd, reinterpret_cast<HMENU>(1021), GetModuleHandleW(nullptr), nullptr);

        y = 4;
        CreateWindowExW(0, L"STATIC", L"小カテゴリ",
            WS_CHILD|WS_VISIBLE, x2, y, half, 16,
            hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        y += 18;
        g_sc.listSmall = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_NOTIFY,
            x2, y, half, 140,
            hwnd, reinterpret_cast<HMENU>(2001), GetModuleHandleW(nullptr), nullptr);
        y += 144;
        g_sc.btnSmallAdd   = CreateWindowExW(0, L"BUTTON", L"追加",  WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x2,      y, 40, 22, hwnd, reinterpret_cast<HMENU>(2010), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnSmallDel   = CreateWindowExW(0, L"BUTTON", L"削除",  WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x2+44,  y, 40, 22, hwnd, reinterpret_cast<HMENU>(2011), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnSmallUp    = CreateWindowExW(0, L"BUTTON", L"↑",    WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x2+88,  y, 30, 22, hwnd, reinterpret_cast<HMENU>(2012), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnSmallDown  = CreateWindowExW(0, L"BUTTON", L"↓",    WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x2+122, y, 30, 22, hwnd, reinterpret_cast<HMENU>(2013), GetModuleHandleW(nullptr), nullptr);
        y += 26;
        CreateWindowExW(0, L"STATIC", L"名前:", WS_CHILD|WS_VISIBLE, x2, y, 36, 18, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        g_sc.editSmallName = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL,
            x2+38, y, half-38-50, 20,
            hwnd, reinterpret_cast<HMENU>(2020), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnSmallSave = CreateWindowExW(0, L"BUTTON", L"保存",  WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
            x2+half-46, y, 46, 22, hwnd, reinterpret_cast<HMENU>(2021), GetModuleHandleW(nullptr), nullptr);
        y += 26;
        CreateWindowExW(0, L"STATIC", L"パス:", WS_CHILD|WS_VISIBLE, x2, y, 36, 18, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        g_sc.editSmallPath = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL,
            x2+38, y, half-38-56, 20,
            hwnd, reinterpret_cast<HMENU>(2030), GetModuleHandleW(nullptr), nullptr);
        g_sc.btnSmallBrowse = CreateWindowExW(0, L"BUTTON", L"参照...", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
            x2+half-52, y, 52, 22, hwnd, reinterpret_cast<HMENU>(2031), GetModuleHandleW(nullptr), nullptr);

        g_sc.btnClose = CreateWindowExW(0, L"BUTTON", L"閉じる", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
            W-80, H-30, 76, 26, hwnd, reinterpret_cast<HMENU>(9999), GetModuleHandleW(nullptr), nullptr);

        g_settingsTabs = g_bigTabs;
        g_settBigIdx   = g_bigTabIdx;
        SettingsRefreshBigList();
        SettingsRefreshSmallList();
        return 0;
    }

    case WM_COMMAND: {
        int id = LOWORD(wp);
        switch (id) {
        case 1001:
            if (HIWORD(wp) == LBN_SELCHANGE) {
                int sel = (int)SendMessageW(g_sc.listBig, LB_GETCURSEL, 0, 0);
                if (sel != LB_ERR) {
                    g_settBigIdx = sel;
                    SettingsRefreshSmallList();
                    if (g_settBigIdx < (int)g_settingsTabs.size())
                        SetWindowTextW(g_sc.editBigName, g_settingsTabs[g_settBigIdx].name.c_str());
                }
            }
            break;
        case 2001:
            if (HIWORD(wp) == LBN_SELCHANGE) {
                if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                    auto& bt = g_settingsTabs[g_settBigIdx];
                    int sel = (int)SendMessageW(g_sc.listSmall, LB_GETCURSEL, 0, 0);
                    if (sel != LB_ERR && sel < (int)bt.smallTabs.size()) {
                        SetWindowTextW(g_sc.editSmallName, bt.smallTabs[sel].name.c_str());
                        SetWindowTextW(g_sc.editSmallPath, bt.smallTabs[sel].path.c_str());
                    }
                }
            }
            break;
        case 1010: {
            BigTab bt; bt.name = L"新規";
            SmallTab st; st.name = L"クイックアクセス"; st.path = L"";
            bt.smallTabs.push_back(st);
            g_settingsTabs.push_back(bt);
            g_settBigIdx = (int)g_settingsTabs.size() - 1;
            SettingsRefreshBigList();
            SettingsRefreshSmallList();
            break;
        }
        case 1011: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                g_settingsTabs.erase(g_settingsTabs.begin() + g_settBigIdx);
                if (g_settBigIdx >= (int)g_settingsTabs.size()) g_settBigIdx--;
                SettingsRefreshBigList();
                SettingsRefreshSmallList();
            }
            break;
        }
        case 1012: {
            if (g_settBigIdx > 0) {
                std::swap(g_settingsTabs[g_settBigIdx], g_settingsTabs[g_settBigIdx - 1]);
                g_settBigIdx--;
                SettingsRefreshBigList();
            }
            break;
        }
        case 1013: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size() - 1) {
                std::swap(g_settingsTabs[g_settBigIdx], g_settingsTabs[g_settBigIdx + 1]);
                g_settBigIdx++;
                SettingsRefreshBigList();
            }
            break;
        }
        case 1021: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                wchar_t buf[256]{};
                GetWindowTextW(g_sc.editBigName, buf, 256);
                g_settingsTabs[g_settBigIdx].name = buf;
                SettingsRefreshBigList();
            }
            break;
        }
        case 2010: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                SmallTab st; st.name = L"新規"; st.path = L"";
                g_settingsTabs[g_settBigIdx].smallTabs.push_back(st);
                SettingsRefreshSmallList();
            }
            break;
        }
        case 2011: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                auto& bt = g_settingsTabs[g_settBigIdx];
                int sel = (int)SendMessageW(g_sc.listSmall, LB_GETCURSEL, 0, 0);
                if (sel != LB_ERR && sel < (int)bt.smallTabs.size()) {
                    bt.smallTabs.erase(bt.smallTabs.begin() + sel);
                    SettingsRefreshSmallList();
                }
            }
            break;
        }
        case 2012: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                auto& bt = g_settingsTabs[g_settBigIdx];
                int sel = (int)SendMessageW(g_sc.listSmall, LB_GETCURSEL, 0, 0);
                if (sel > 0 && sel < (int)bt.smallTabs.size()) {
                    std::swap(bt.smallTabs[sel], bt.smallTabs[sel - 1]);
                    SendMessageW(g_sc.listSmall, LB_SETCURSEL, sel - 1, 0);
                    SettingsRefreshSmallList();
                }
            }
            break;
        }
        case 2013: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                auto& bt = g_settingsTabs[g_settBigIdx];
                int sel = (int)SendMessageW(g_sc.listSmall, LB_GETCURSEL, 0, 0);
                if (sel >= 0 && sel < (int)bt.smallTabs.size() - 1) {
                    std::swap(bt.smallTabs[sel], bt.smallTabs[sel + 1]);
                    SendMessageW(g_sc.listSmall, LB_SETCURSEL, sel + 1, 0);
                    SettingsRefreshSmallList();
                }
            }
            break;
        }
        case 2021: {
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                auto& bt = g_settingsTabs[g_settBigIdx];
                int sel = (int)SendMessageW(g_sc.listSmall, LB_GETCURSEL, 0, 0);
                if (sel >= 0 && sel < (int)bt.smallTabs.size()) {
                    wchar_t nbuf[256]{}, pbuf[MAX_PATH]{};
                    GetWindowTextW(g_sc.editSmallName, nbuf, 256);
                    GetWindowTextW(g_sc.editSmallPath, pbuf, MAX_PATH);
                    bt.smallTabs[sel].name = nbuf;
                    bt.smallTabs[sel].path = pbuf;
                    SettingsRefreshSmallList();
                }
            }
            break;
        }
        case 2031: {
            wchar_t buf[MAX_PATH]{};
            BROWSEINFOW bi{};
            bi.hwndOwner = hwnd;
            bi.pszDisplayName = buf;
            bi.lpszTitle = L"フォルダを選択";
            bi.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
            if (pidl) {
                wchar_t path[MAX_PATH]{};
                SHGetPathFromIDListW(pidl, path);
                CoTaskMemFree(pidl);
                SetWindowTextW(g_sc.editSmallPath, path);
            }
            break;
        }
        case 9999:
            SettingsApplyToMain();
            DestroyWindow(hwnd);
            g_hwndSettings = nullptr;
            break;
        }
        return 0;
    }

    case WM_CLOSE:
        SettingsApplyToMain();
        DestroyWindow(hwnd);
        g_hwndSettings = nullptr;
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

static void OpenSettingsCallback(HWND hwndParent, HINSTANCE /*dll_hinst*/) {
    if (g_hwndSettings && IsWindow(g_hwndSettings)) {
        SetForegroundWindow(g_hwndSettings);
        return;
    }
    g_hwndSettings = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        WC_SETTINGS, L"PreviewExplorer 設定",
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, 640, 320,
        hwndParent, nullptr, GetModuleHandleW(nullptr), nullptr);
}

//==============================================================================
// 右ペイン プロシージャ
//==============================================================================
static LRESULT CALLBACK RightWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        HINSTANCE hi = GetModuleHandleW(nullptr);

        // ビッグタブ（太字フォント・高さ36px）
        g_hwndBigTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE | TCS_HOTTRACK,
            0, 0, 400, BIG_TAB_H,
            hwnd, reinterpret_cast<HMENU>(101), hi, nullptr);
        if (g_hFontBig)
            SendMessageW(g_hwndBigTab, WM_SETFONT, (WPARAM)g_hFontBig, TRUE);

        // スモールタブ（通常フォント・高さ24px）
        g_hwndSmallTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE,
            0, BIG_TAB_H, 400, SMALL_TAB_H,
            hwnd, reinterpret_cast<HMENU>(102), hi, nullptr);
        if (g_hFontNormal)
            SendMessageW(g_hwndSmallTab, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        // タブ サブクラス化
        SetWindowSubclass(g_hwndBigTab,   BigTabSubclassProc,   101, 0);
        SetWindowSubclass(g_hwndSmallTab, SmallTabSubclassProc, 102, 0);

        // ビッグタブ初期化
        for (int i = 0; i < (int)g_bigTabs.size(); i++) {
            TCITEMW tci{};
            tci.mask    = TCIF_TEXT;
            tci.pszText = const_cast<wchar_t*>(g_bigTabs[i].name.c_str());
            TabCtrl_InsertItem(g_hwndBigTab, i, &tci);
        }
        if (!g_bigTabs.empty())
            TabCtrl_SetCurSel(g_hwndBigTab, g_bigTabIdx);

        RebuildSmallTabs();

        // IExplorerBrowser 初期化
        HRESULT hr = CoCreateInstance(CLSID_ExplorerBrowser, nullptr,
            CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_peb));
        if (SUCCEEDED(hr)) {
            RECT rc = {0, 0, 400, 400};
            FOLDERSETTINGS fs = { FVM_DETAILS, FWF_AUTOARRANGE };
            g_peb->Initialize(hwnd, &rc, &fs);
            g_peb->SetOptions(EBO_NOBORDER | EBO_ALWAYSNAVIGATE);

            // 初期ナビゲーション
            NavigateExplorerToCurrentTab();
        } else {
            Log(L"IExplorerBrowser の作成に失敗しました。");
        }
        return 0;
    }

    case WM_ERASEBKGND: {
        HDC hdc = reinterpret_cast<HDC>(wp);
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_BTNFACE + 1));
        return 1;
    }

    case WM_SIZE:
        DoRightLayout(hwnd);
        return 0;

    case WM_NOTIFY: {
        auto* nmhdr = reinterpret_cast<NMHDR*>(lp);
        // ビッグタブ切り替え
        if (nmhdr->hwndFrom == g_hwndBigTab && nmhdr->code == TCN_SELCHANGE) {
            int sel = TabCtrl_GetCurSel(g_hwndBigTab);
            if (sel >= 0) {
                g_bigTabIdx = sel;
                RebuildSmallTabs();
                NavigateExplorerToCurrentTab();
            }
            return 0;
        }
        // スモールタブ切り替え
        if (nmhdr->hwndFrom == g_hwndSmallTab && nmhdr->code == TCN_SELCHANGE) {
            int sel = TabCtrl_GetCurSel(g_hwndSmallTab);
            if (sel >= 0 && g_bigTabIdx < (int)g_bigTabs.size()) {
                g_bigTabs[g_bigTabIdx].smallTabIdx = sel;
                NavigateExplorerToCurrentTab();
            }
            return 0;
        }
        break;
    }

    case WM_DESTROY:
        if (g_peb) {
            g_peb->Destroy();
            g_peb->Release();
            g_peb = nullptr;
        }
        if (g_hFontBig)    { DeleteObject(g_hFontBig);    g_hFontBig    = nullptr; }
        if (g_hFontNormal) { DeleteObject(g_hFontNormal); g_hFontNormal = nullptr; }
        if (g_hBrushBg)    { DeleteObject(g_hBrushBg);    g_hBrushBg    = nullptr; }
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

//------------------------------------------------------------------------------
// INI 保存・読み込み
//------------------------------------------------------------------------------
static void SaveState() {
    auto ini = GetIniPath();
    EnsureDir(ini);

    WritePrivateProfileStringW(L"Window", L"SplitX",
        std::to_wstring(g_splitX).c_str(), ini.c_str());

    WritePrivateProfileStringW(L"Tabs", L"BigTabCount",
        std::to_wstring(g_bigTabs.size()).c_str(), ini.c_str());
    WritePrivateProfileStringW(L"Tabs", L"BigTabIndex",
        std::to_wstring(g_bigTabIdx).c_str(), ini.c_str());
    for (int i = 0; i < (int)g_bigTabs.size(); i++) {
        WritePrivateProfileStringW(L"Tabs",
            (L"SmallTabIndex_" + std::to_wstring(i)).c_str(),
            std::to_wstring(g_bigTabs[i].smallTabIdx).c_str(), ini.c_str());
    }

    for (int i = 0; i < (int)g_bigTabs.size(); i++) {
        std::wstring sec = L"BigTab_" + std::to_wstring(i);
        WritePrivateProfileStringW(sec.c_str(), L"Name",
            g_bigTabs[i].name.c_str(), ini.c_str());
        WritePrivateProfileStringW(sec.c_str(), L"SmallTabCount",
            std::to_wstring(g_bigTabs[i].smallTabs.size()).c_str(), ini.c_str());
        for (int j = 0; j < (int)g_bigTabs[i].smallTabs.size(); j++) {
            WritePrivateProfileStringW(sec.c_str(),
                (L"SmallTab_" + std::to_wstring(j) + L"_Name").c_str(),
                g_bigTabs[i].smallTabs[j].name.c_str(), ini.c_str());
            WritePrivateProfileStringW(sec.c_str(),
                (L"SmallTab_" + std::to_wstring(j) + L"_Path").c_str(),
                g_bigTabs[i].smallTabs[j].path.c_str(), ini.c_str());
        }
    }
}

static void LoadState() {
    auto ini = GetIniPath();

    g_splitX = GetPrivateProfileIntW(L"Window", L"SplitX", 400, ini.c_str());

    int bigCount = GetPrivateProfileIntW(L"Tabs", L"BigTabCount", 0, ini.c_str());
    g_bigTabIdx  = GetPrivateProfileIntW(L"Tabs", L"BigTabIndex",  0, ini.c_str());

    g_bigTabs.clear();

    if (bigCount == 0) {
        // デフォルト構造
        BigTab bt;
        bt.name = L"動画";
        SmallTab qa;     qa.name = L"クイックアクセス"; qa.path = L"";
        SmallTab sample; sample.name = L"素材フォルダ"; sample.path = L"";
        bt.smallTabs.push_back(qa);
        bt.smallTabs.push_back(sample);
        g_bigTabs.push_back(bt);
        g_bigTabIdx = 0;
    } else {
        for (int i = 0; i < bigCount; i++) {
            std::wstring sec = L"BigTab_" + std::to_wstring(i);
            wchar_t nbuf[256]{};
            GetPrivateProfileStringW(sec.c_str(), L"Name", L"", nbuf, 256, ini.c_str());

            BigTab bt;
            bt.name = nbuf;
            bt.smallTabIdx = GetPrivateProfileIntW(L"Tabs",
                (L"SmallTabIndex_" + std::to_wstring(i)).c_str(), 0, ini.c_str());

            int stCount = GetPrivateProfileIntW(sec.c_str(), L"SmallTabCount", 0, ini.c_str());
            for (int j = 0; j < stCount; j++) {
                wchar_t sn[256]{}, sp[MAX_PATH]{};
                GetPrivateProfileStringW(sec.c_str(),
                    (L"SmallTab_" + std::to_wstring(j) + L"_Name").c_str(),
                    L"", sn, 256, ini.c_str());
                GetPrivateProfileStringW(sec.c_str(),
                    (L"SmallTab_" + std::to_wstring(j) + L"_Path").c_str(),
                    L"", sp, MAX_PATH, ini.c_str());
                SmallTab st; st.name = sn; st.path = sp;
                bt.smallTabs.push_back(st);
            }
            g_bigTabs.push_back(bt);
        }
    }

    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size())
        g_bigTabIdx = 0;
}

//------------------------------------------------------------------------------
// レイアウト（メインウィンドウ）
//------------------------------------------------------------------------------
static void DoLayout(HWND hwnd) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    int W = rc.right, H = rc.bottom;

    g_splitX = max(MIN_LEFT_W, min(g_splitX, W - MIN_RIGHT_W - SPLITTER_W));

    SetWindowPos(g_hwndLeft,  nullptr, 0, 0, g_splitX, H, SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(g_hwndRight, nullptr, g_splitX + SPLITTER_W, 0,
        W - g_splitX - SPLITTER_W, H, SWP_NOZORDER | SWP_NOACTIVATE);

    // 組み込みプレビューをリサイズ
    if (g_hwndEmbedded && IsWindow(g_hwndEmbedded)) {
        RECT lrc;
        GetClientRect(g_hwndLeft, &lrc);
        SetWindowPos(g_hwndEmbedded, nullptr, 0, 0, lrc.right, lrc.bottom,
            SWP_NOZORDER | SWP_NOACTIVATE);
    }
}

//------------------------------------------------------------------------------
// プレビューウィンドウ探索 & 埋め込み（Phase 1 保持）
//------------------------------------------------------------------------------
struct WndEntry {
    HWND    hwnd;
    DWORD   pid;
    wchar_t cls[256];
    wchar_t title[256];
    int     w, h;
};
struct EnumCtx {
    DWORD                 targetPid;
    std::vector<WndEntry> entries;
};
static BOOL CALLBACK EnumTopCb(HWND hwnd, LPARAM lp) {
    auto* ctx = reinterpret_cast<EnumCtx*>(lp);
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (pid != ctx->targetPid) return TRUE;

    WndEntry e{};
    e.hwnd = hwnd; e.pid = pid;
    GetClassNameW(hwnd, e.cls, 256);
    GetWindowTextW(hwnd, e.title, 256);
    RECT rc{};
    GetWindowRect(hwnd, &rc);
    e.w = rc.right - rc.left;
    e.h = rc.bottom - rc.top;
    ctx->entries.push_back(e);
    return TRUE;
}

static void TryEmbedPreview() {
    HWND hostMain = g_editHandle->get_host_app_window();

    std::wstring log;
    if (!hostMain) {
        log = L"get_host_app_window() = null\nAviUtl2 本体ウィンドウを取得できませんでした。";
        if (g_logger) g_logger->warn(g_logger, log.c_str());
        if (g_hwndInfo && IsWindow(g_hwndInfo)) SetWindowTextW(g_hwndInfo, log.c_str());
        return;
    }

    DWORD hostPid = 0;
    GetWindowThreadProcessId(hostMain, &hostPid);

    wchar_t buf[128];
    swprintf_s(buf, L"HostMain HWND=%08X  PID=%u\n", (DWORD)(DWORD_PTR)hostMain, hostPid);
    log = buf;

    EnumCtx ctx{ hostPid, {} };
    EnumWindows(EnumTopCb, reinterpret_cast<LPARAM>(&ctx));

    swprintf_s(buf, L"同プロセス TL ウィンドウ数: %d\n\n", (int)ctx.entries.size());
    log += buf;

    HWND preview = nullptr;
    for (auto& e : ctx.entries) {
        wchar_t line[512];
        swprintf_s(line, L"[%08X] %s \"%s\" %dx%d\n",
            (DWORD)(DWORD_PTR)e.hwnd, e.cls, e.title, e.w, e.h);
        log += line;
        if (!preview && wcsstr(e.title, L"プレビュー")) preview = e.hwnd;
    }

    if (g_logger) g_logger->log(g_logger, log.c_str());
    if (g_hwndInfo && IsWindow(g_hwndInfo)) SetWindowTextW(g_hwndInfo, log.c_str());

    if (preview) {
        swprintf_s(buf, L"プレビュー候補: HWND=%08X -> SetParent 試行", (DWORD)(DWORD_PTR)preview);
        if (g_logger) g_logger->log(g_logger, buf);

        RECT lrc{};
        GetClientRect(g_hwndLeft, &lrc);

        LONG style = GetWindowLongW(preview, GWL_STYLE);
        SetWindowLongW(preview, GWL_STYLE,
            (style & ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME)) | WS_CHILD | WS_VISIBLE);
        SetParent(preview, g_hwndLeft);
        SetWindowPos(preview, HWND_TOP, 0, 0, lrc.right, lrc.bottom,
            SWP_SHOWWINDOW | SWP_FRAMECHANGED);

        g_hwndEmbedded = preview;
        if (g_hwndInfo && IsWindow(g_hwndInfo)) ShowWindow(g_hwndInfo, SW_HIDE);
        if (g_logger) g_logger->log(g_logger, L"SetParent 完了。表示を確認してください。");
    } else {
        log += L"\n[!] タイトルに「プレビュー」を含むウィンドウが見つかりませんでした。\n";
        if (g_hwndInfo && IsWindow(g_hwndInfo)) SetWindowTextW(g_hwndInfo, log.c_str());
        if (g_logger) g_logger->warn(g_logger, L"プレビューウィンドウが見つかりませんでした。");
    }
}

//------------------------------------------------------------------------------
// メインウィンドウ プロシージャ
//------------------------------------------------------------------------------
static LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE:
        LoadState();

        // 左ペイン（プレビューコンテナ）
        g_hwndLeft = CreateWindowExW(0, WC_STATIC, nullptr,
            WS_VISIBLE | WS_CHILD | SS_BLACKRECT,
            0, 0, g_splitX, 400,
            hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        // 左ペイン内：診断テキスト
        g_hwndInfo = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT",
            L"プレビューウィンドウを検索中...",
            WS_VISIBLE | WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            0, 0, g_splitX, 400,
            g_hwndLeft, nullptr, GetModuleHandleW(nullptr), nullptr);

        // 右ペイン（エクスプローラー）
        g_hwndRight = CreateWindowExW(0, WC_RIGHT, nullptr,
            WS_VISIBLE | WS_CHILD,
            g_splitX + SPLITTER_W, 0, 400, 400,
            hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        SetTimer(hwnd, TIMER_FIND_PREVIEW, 2000, nullptr);
        return 0;

    case WM_TIMER:
        if (wp == TIMER_FIND_PREVIEW) {
            KillTimer(hwnd, TIMER_FIND_PREVIEW);
            TryEmbedPreview();
        }
        return 0;

    case WM_SIZE:
        DoLayout(hwnd);
        return 0;

    case WM_SETCURSOR: {
        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(hwnd, &pt);
        if (pt.x >= g_splitX && pt.x < g_splitX + SPLITTER_W) {
            SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
            return TRUE;
        }
        break;
    }
    case WM_LBUTTONDOWN: {
        int x = GET_X_LPARAM(lp);
        if (x >= g_splitX && x < g_splitX + SPLITTER_W) {
            g_dragging = true;
            SetCapture(hwnd);
        }
        return 0;
    }
    case WM_MOUSEMOVE:
        if (g_dragging) {
            g_splitX = GET_X_LPARAM(lp);
            DoLayout(hwnd);
        }
        return 0;
    case WM_LBUTTONUP:
        if (g_dragging) {
            g_dragging = false;
            ReleaseCapture();
            SaveState();
        }
        return 0;

    case WM_ERASEBKGND: {
        // スプリッター領域をテーマ背景色で塗りつぶす
        HDC hdc = reinterpret_cast<HDC>(wp);
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_WINDOW + 1));
        return 1;
    }

    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT: {
        // 診断ログ EDIT をテーマ色で表示
        HDC hdc = reinterpret_cast<HDC>(wp);
        SetTextColor(hdc, g_clrText);
        SetBkColor(hdc, g_clrBg);
        return reinterpret_cast<LRESULT>(g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_WINDOW + 1));
    }

    case WM_DESTROY:
        SaveState();
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

//------------------------------------------------------------------------------
// DLL エクスポート
//------------------------------------------------------------------------------
EXTERN_C __declspec(dllexport) DWORD RequiredVersion()               { return 2003300; }
EXTERN_C __declspec(dllexport) void  InitializeLogger(LOG_HANDLE* h)  { g_logger = h; }
EXTERN_C __declspec(dllexport) void  InitializeConfig(CONFIG_HANDLE* h){ g_config = h; }
EXTERN_C __declspec(dllexport) bool  InitializePlugin(DWORD)           { return true; }
EXTERN_C __declspec(dllexport) void  UninitializePlugin()              {}

EXTERN_C __declspec(dllexport) void RegisterPlugin(HOST_APP_TABLE* host) {
    host->set_plugin_information(L"Preview edit + Explorer version 0.2");

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    // テーマ初期化 (InitializeConfig → RegisterPlugin の順なので g_config は有効)
    InitTheme();

    INITCOMMONCONTROLSEX icex{ sizeof(icex), ICC_WIN95_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&icex);

    HINSTANCE hi = GetModuleHandleW(nullptr);

    // Main ウィンドウクラス (WM_ERASEBKGND で塗るため NULL でも可、念のため設定)
    {
        WNDCLASSEXW wcex{};
        wcex.cbSize        = sizeof(wcex);
        wcex.lpszClassName = WC_MAIN;
        wcex.lpfnWndProc   = MainWndProc;
        wcex.hInstance     = hi;
        wcex.hbrBackground = g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_WINDOW + 1);
        wcex.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wcex);
    }
    // Right ウィンドウクラス (WM_ERASEBKGND で塗るため NULL でも可)
    {
        WNDCLASSEXW wcex{};
        wcex.cbSize        = sizeof(wcex);
        wcex.lpszClassName = WC_RIGHT;
        wcex.lpfnWndProc   = RightWndProc;
        wcex.hInstance     = hi;
        wcex.hbrBackground = g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_BTNFACE + 1);
        wcex.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wcex);
    }
    // Settings ウィンドウクラス
    {
        WNDCLASSEXW wcex{};
        wcex.cbSize        = sizeof(wcex);
        wcex.lpszClassName = WC_SETTINGS;
        wcex.lpfnWndProc   = SettingsWndProc;
        wcex.hInstance     = hi;
        wcex.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wcex.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wcex);
    }
    // InputDlg ウィンドウクラス
    {
        WNDCLASSEXW wcex{};
        wcex.cbSize        = sizeof(wcex);
        wcex.lpszClassName = WC_INPUTDLG;
        wcex.lpfnWndProc   = InputDlgProc;
        wcex.hInstance     = hi;
        wcex.cbWndExtra    = sizeof(LONG_PTR);
        wcex.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wcex.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wcex);
    }

    // WS_POPUP で作成（AviUtl2 が WS_CHILD に変換して管理する）
    g_hwndMain = CreateWindowExW(
        0, WC_MAIN, L"Preview edit + Explorer",
        WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT,
        CW_USEDEFAULT, CW_USEDEFAULT,
        nullptr, nullptr, hi, nullptr);
    if (!g_hwndMain) return;

    // ダークモード適用 (Windows 10 1809+)
    TrySetDarkMode(g_hwndMain);

    host->register_window_client(L"Preview edit + Explorer", g_hwndMain);
    g_editHandle = host->create_edit_handle();
    host->register_config_menu(L"Preview edit + Explorer", OpenSettingsCallback);
}
