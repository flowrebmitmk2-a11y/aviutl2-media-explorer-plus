//----------------------------------------------------------------------------------
//  PreviewExplorer - Phase 1-5 完全実装
//  left pane: AviUtl2 preview embedding
//  right pane: media explorer (2-tier tabs, file list, quick access)
//----------------------------------------------------------------------------------
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <algorithm>
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")

#include "plugin2.h"
#include "logger2.h"
#include "config2.h"

//------------------------------------------------------------------------------
// データ構造
//------------------------------------------------------------------------------
struct SmallTab { std::wstring name, path; };
struct BigTab   { std::wstring name; std::vector<SmallTab> smallTabs; int smallTabIdx = 0; };

struct FileItem {
    std::wstring name, fullPath, typeName;
    bool     isDir      = false;
    bool     supported  = true;
    ULONGLONG size      = 0;
    FILETIME  modified  = {};
    int       iconIdx   = 0;
};

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
HWND g_hwndBigTab    = nullptr;  // ID=101
HWND g_hwndSmallTab  = nullptr;  // ID=102
HWND g_hwndQAToggle  = nullptr;  // ID=103
HWND g_hwndQAPanel   = nullptr;
HWND g_hwndUpBtn     = nullptr;  // ID=104
HWND g_hwndPathLabel = nullptr;  // ID=105
HWND g_hwndListView  = nullptr;  // ID=106

// タブデータ
std::vector<BigTab> g_bigTabs;
int g_bigTabIdx = 0;

// ファイルリスト
std::vector<FileItem> g_fileItems;
std::wstring          g_currentPath;
bool                  g_qaVisible = true;

// クイックアクセス
struct QAItem { std::wstring path; HWND btn = nullptr; };
std::vector<QAItem> g_qaItems;

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

// レイアウト定数（右ペイン内）
constexpr int BIG_TAB_H   = 26;
constexpr int SMALL_TAB_H = 26;
constexpr int QA_BTN_H    = 24;
constexpr int QA_PANEL_H  = 60;  // 折り畳み時 0、展開時この値
constexpr int NAV_H       = 26;
constexpr int UP_BTN_W    = 60;
constexpr int QA_BASE_ID  = 200;

// 設定ウィンドウ用前方宣言
static HWND g_hwndSettings = nullptr;

//------------------------------------------------------------------------------
// ログユーティリティ
//------------------------------------------------------------------------------
static void Log(const wchar_t* msg) {
    if (g_logger) g_logger->log(g_logger, msg);
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
// ファイルサイズ文字列
//------------------------------------------------------------------------------
static std::wstring FormatSize(ULONGLONG sz) {
    wchar_t buf[64];
    if (sz < 1024ULL)            swprintf_s(buf, L"%llu B",        sz);
    else if (sz < 1024ULL*1024)  swprintf_s(buf, L"%.1f KB",       sz / 1024.0);
    else if (sz < 1024ULL*1024*1024) swprintf_s(buf, L"%.1f MB",  sz / (1024.0*1024));
    else                         swprintf_s(buf, L"%.2f GB",       sz / (1024.0*1024*1024));
    return buf;
}

//------------------------------------------------------------------------------
// FILETIME -> 表示文字列
//------------------------------------------------------------------------------
static std::wstring FormatFileTime(const FILETIME& ft) {
    FILETIME local{};
    FileTimeToLocalFileTime(&ft, &local);
    SYSTEMTIME st{};
    FileTimeToSystemTime(&local, &st);
    wchar_t buf[64];
    swprintf_s(buf, L"%04d/%02d/%02d %02d:%02d",
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute);
    return buf;
}

//------------------------------------------------------------------------------
// シェルシステムイメージリスト取得
//------------------------------------------------------------------------------
static HIMAGELIST g_sysImgList = nullptr;

static void EnsureSysImageList() {
    if (g_sysImgList) return;
    SHFILEINFOW sfi{};
    g_sysImgList = reinterpret_cast<HIMAGELIST>(
        SHGetFileInfoW(L"C:\\", 0, &sfi, sizeof(sfi),
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON));
}

//------------------------------------------------------------------------------
// is_support_media_file コールバック用パラメータ
//------------------------------------------------------------------------------
struct SupportCheckParam {
    std::vector<FileItem>* items;
    bool done = false;
};

static void SupportCheckCb(void* param, EDIT_SECTION* edit) {
    auto* p = static_cast<SupportCheckParam*>(param);
    for (auto& f : *p->items) {
        if (!f.isDir) {
            f.supported = edit->is_support_media_file(f.fullPath.c_str(), false);
        }
    }
    p->done = true;
}

//------------------------------------------------------------------------------
// QA ボタン再構築
//------------------------------------------------------------------------------
static void RebuildQAButtons() {
    // 既存ボタン破棄
    for (auto& qa : g_qaItems) {
        if (qa.btn && IsWindow(qa.btn)) DestroyWindow(qa.btn);
        qa.btn = nullptr;
    }
    if (!g_hwndQAPanel) return;

    int x = 2, y = 2;
    int btnW = 120, btnH = 22;
    for (int i = 0; i < (int)g_qaItems.size(); i++) {
        // パスからフォルダ名を取得
        std::wstring label = g_qaItems[i].path;
        auto pos = label.rfind(L'\\');
        if (pos != std::wstring::npos && pos + 1 < label.size())
            label = label.substr(pos + 1);
        if (label.empty()) label = g_qaItems[i].path;

        HWND btn = CreateWindowExW(0, L"BUTTON", label.c_str(),
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            x, y, btnW, btnH,
            g_hwndQAPanel,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(QA_BASE_ID + i)),
            GetModuleHandleW(nullptr), nullptr);
        g_qaItems[i].btn = btn;
        x += btnW + 4;
        if (x + btnW > 500) { x = 2; y += btnH + 2; }
    }
}

//------------------------------------------------------------------------------
// ファイルリスト構築
//------------------------------------------------------------------------------
static void PopulateFileList(const std::wstring& path) {
    g_fileItems.clear();
    if (g_hwndListView) ListView_DeleteAllItems(g_hwndListView);

    if (path.empty()) {
        SetWindowTextW(g_hwndPathLabel, L"");
        return;
    }

    SetWindowTextW(g_hwndPathLabel, path.c_str());
    g_currentPath = path;

    std::wstring pattern = path;
    if (!pattern.empty() && pattern.back() != L'\\') pattern += L'\\';
    pattern += L'*';

    std::vector<FileItem> dirs, files;

    WIN32_FIND_DATAW fd{};
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE) goto done;

    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        FileItem fi;
        fi.name = fd.cFileName;
        fi.fullPath = path + (path.back() == L'\\' ? L"" : L"\\") + fd.cFileName;
        fi.isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        fi.size  = fi.isDir ? 0 : ((ULONGLONG)fd.nFileSizeHigh << 32) | fd.nFileSizeLow;
        fi.modified = fd.ftLastWriteTime;

        // アイコン & 種類名
        SHFILEINFOW sfi{};
        SHGetFileInfoW(fi.fullPath.c_str(), 0, &sfi, sizeof(sfi),
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_TYPENAME);
        fi.iconIdx  = sfi.iIcon;
        fi.typeName = sfi.szTypeName;

        if (fi.isDir) dirs.push_back(std::move(fi));
        else          files.push_back(std::move(fi));
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);

    // 名前順ソート
    auto byName = [](const FileItem& a, const FileItem& b) {
        return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0;
    };
    std::sort(dirs.begin(),  dirs.end(),  byName);
    std::sort(files.begin(), files.end(), byName);

    for (auto& d : dirs)   g_fileItems.push_back(std::move(d));
    for (auto& f : files)  g_fileItems.push_back(std::move(f));

done:
    // サポートチェック
    if (g_editHandle) {
        SupportCheckParam scp{ &g_fileItems };
        g_editHandle->call_edit_section_param(&scp, SupportCheckCb);
        // 失敗 (プロジェクトなし) の場合 done==false のまま → supported=true のまま
    }

    // リストビューへ追加
    if (!g_hwndListView) return;
    EnsureSysImageList();

    for (int i = 0; i < (int)g_fileItems.size(); i++) {
        const auto& fi = g_fileItems[i];
        LVITEMW lvi{};
        lvi.mask    = LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM;
        lvi.iItem   = i;
        lvi.iSubItem = 0;
        lvi.pszText = const_cast<wchar_t*>(fi.name.c_str());
        lvi.iImage  = fi.iconIdx;
        lvi.lParam  = static_cast<LPARAM>(i);
        ListView_InsertItem(g_hwndListView, &lvi);

        // 種類
        ListView_SetItemText(g_hwndListView, i, 1,
            const_cast<wchar_t*>(fi.typeName.c_str()));
        // サイズ
        std::wstring szStr = fi.isDir ? L"" : FormatSize(fi.size);
        ListView_SetItemText(g_hwndListView, i, 2,
            const_cast<wchar_t*>(szStr.c_str()));
        // 更新日時
        std::wstring dtStr = FormatFileTime(fi.modified);
        ListView_SetItemText(g_hwndListView, i, 3,
            const_cast<wchar_t*>(dtStr.c_str()));
    }
}

//------------------------------------------------------------------------------
// 親フォルダへ移動
//------------------------------------------------------------------------------
static void NavigateUp() {
    if (g_currentPath.empty()) return;
    std::wstring parent = g_currentPath;
    // 末尾の \\ を除去
    while (!parent.empty() && parent.back() == L'\\') parent.pop_back();
    auto pos = parent.rfind(L'\\');
    if (pos == std::wstring::npos) {
        // ドライブルートの場合は移動しない
        return;
    }
    std::wstring up = parent.substr(0, pos);
    if (up.size() == 2 && up[1] == L':') up += L'\\';  // "C:" -> "C:\"
    PopulateFileList(up);
}

//------------------------------------------------------------------------------
// 現在のスモールタブのパスに移動
//------------------------------------------------------------------------------
static void NavigateToCurrentTab() {
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    auto& bt = g_bigTabs[g_bigTabIdx];
    if (bt.smallTabIdx < 0 || bt.smallTabIdx >= (int)bt.smallTabs.size()) return;
    const std::wstring& path = bt.smallTabs[bt.smallTabIdx].path;
    PopulateFileList(path);
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
    int W = rc.right;
    int y = 0;

    if (g_hwndBigTab) {
        SetWindowPos(g_hwndBigTab, nullptr, 0, y, W, BIG_TAB_H, SWP_NOZORDER | SWP_NOACTIVATE);
        y += BIG_TAB_H;
    }
    if (g_hwndSmallTab) {
        SetWindowPos(g_hwndSmallTab, nullptr, 0, y, W, SMALL_TAB_H, SWP_NOZORDER | SWP_NOACTIVATE);
        y += SMALL_TAB_H;
    }
    if (g_hwndQAToggle) {
        SetWindowPos(g_hwndQAToggle, nullptr, 0, y, W, QA_BTN_H, SWP_NOZORDER | SWP_NOACTIVATE);
        y += QA_BTN_H;
    }
    if (g_hwndQAPanel) {
        int panelH = g_qaVisible ? QA_PANEL_H : 0;
        SetWindowPos(g_hwndQAPanel, nullptr, 0, y, W, panelH, SWP_NOZORDER | SWP_NOACTIVATE);
        ShowWindow(g_hwndQAPanel, g_qaVisible ? SW_SHOW : SW_HIDE);
        y += panelH;
    }
    // Navigation bar
    if (g_hwndUpBtn) {
        SetWindowPos(g_hwndUpBtn, nullptr, 0, y, UP_BTN_W, NAV_H, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    if (g_hwndPathLabel) {
        SetWindowPos(g_hwndPathLabel, nullptr, UP_BTN_W, y, W - UP_BTN_W, NAV_H, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    y += NAV_H;
    // ListView fills the rest
    if (g_hwndListView) {
        int lvH = rc.bottom - y;
        if (lvH < 0) lvH = 0;
        SetWindowPos(g_hwndListView, nullptr, 0, y, W, lvH, SWP_NOZORDER | SWP_NOACTIVATE);
        // 名前列を残りいっぱいに
        RECT lrc;
        GetClientRect(g_hwndListView, &lrc);
        int nameW = lrc.right - 90 - 80 - 130;
        if (nameW < 50) nameW = 50;
        ListView_SetColumnWidth(g_hwndListView, 0, nameW);
    }
}

//------------------------------------------------------------------------------
// ファイルリスト右クリック処理
//------------------------------------------------------------------------------
static void OnListViewRightClick(HWND hwnd, int iItem) {
    if (iItem < 0 || iItem >= (int)g_fileItems.size()) return;
    const FileItem& fi = g_fileItems[iItem];
    if (!fi.isDir) return;

    HMENU hm = CreatePopupMenu();
    AppendMenuW(hm, MF_STRING, 1, L"クイックアクセスに追加");
    POINT pt;
    GetCursorPos(&pt);
    int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
    DestroyMenu(hm);

    if (cmd == 1) {
        // 重複チェック
        for (auto& qa : g_qaItems) {
            if (_wcsicmp(qa.path.c_str(), fi.fullPath.c_str()) == 0) return;
        }
        QAItem qa;
        qa.path = fi.fullPath;
        g_qaItems.push_back(qa);
        RebuildQAButtons();
        SaveState();
    }
}

//------------------------------------------------------------------------------
// QAボタン右クリック処理
//------------------------------------------------------------------------------
static void OnQAButtonRightClick(int qaIdx) {
    HMENU hm = CreatePopupMenu();
    AppendMenuW(hm, MF_STRING, 1, L"削除");
    POINT pt;
    GetCursorPos(&pt);
    int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0,
        g_hwndQAPanel, nullptr);
    DestroyMenu(hm);

    if (cmd == 1 && qaIdx >= 0 && qaIdx < (int)g_qaItems.size()) {
        if (g_qaItems[qaIdx].btn && IsWindow(g_qaItems[qaIdx].btn))
            DestroyWindow(g_qaItems[qaIdx].btn);
        g_qaItems.erase(g_qaItems.begin() + qaIdx);
        RebuildQAButtons();
        SaveState();
    }
}

//------------------------------------------------------------------------------
// ListView ダブルクリック処理
//------------------------------------------------------------------------------
struct PlaceFileParam {
    std::wstring path;
    bool done = false;
};
static void PlaceFileCb(void* param, EDIT_SECTION* edit) {
    auto* p = static_cast<PlaceFileParam*>(param);
    edit->create_object_from_media_file(p->path.c_str(),
        edit->info->layer, edit->info->frame, 0);
    p->done = true;
}

static void OnListViewDblClick(int iItem) {
    if (iItem < 0 || iItem >= (int)g_fileItems.size()) return;
    const FileItem& fi = g_fileItems[iItem];
    if (fi.isDir) {
        PopulateFileList(fi.fullPath);
    } else {
        if (g_editHandle) {
            PlaceFileParam pfp;
            pfp.path = fi.fullPath;
            g_editHandle->call_edit_section_param(&pfp, PlaceFileCb);
        }
    }
}

//==============================================================================
// 設定ウィンドウ
//==============================================================================
// 設定ウィンドウ内コントロール
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

// 設定ウィンドウ用ローカルデータ（開いている間のコピー）
static std::vector<BigTab> g_settingsTabs;
static int g_settBigIdx   = -1;

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

    // ビッグタブを再構築
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
    NavigateToCurrentTab();
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

        // --- Big tab 側 ---
        CreateWindowExW(0, L"STATIC", L"大カテゴリ",
            WS_CHILD | WS_VISIBLE, x1, y, half, 16,
            hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        y += 18;
        g_sc.listBig = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", nullptr,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY,
            x1, y, half, 140,
            hwnd, reinterpret_cast<HMENU>(1001), GetModuleHandleW(nullptr), nullptr);
        y += 144;
        // ボタン行
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

        // --- Small tab 側 ---
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

        // 閉じるボタン
        g_sc.btnClose = CreateWindowExW(0, L"BUTTON", L"閉じる", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
            W-80, H-30, 76, 26, hwnd, reinterpret_cast<HMENU>(9999), GetModuleHandleW(nullptr), nullptr);

        // 初期データ
        g_settingsTabs = g_bigTabs;
        g_settBigIdx   = g_bigTabIdx;
        SettingsRefreshBigList();
        SettingsRefreshSmallList();
        return 0;
    }

    case WM_COMMAND: {
        int id = LOWORD(wp);
        switch (id) {
        case 1001: // Big listbox 選択変更
            if (HIWORD(wp) == LBN_SELCHANGE) {
                int sel = (int)SendMessageW(g_sc.listBig, LB_GETCURSEL, 0, 0);
                if (sel != LB_ERR) {
                    g_settBigIdx = sel;
                    SettingsRefreshSmallList();
                    // 名前エディットに反映
                    if (g_settBigIdx < (int)g_settingsTabs.size())
                        SetWindowTextW(g_sc.editBigName, g_settingsTabs[g_settBigIdx].name.c_str());
                }
            }
            break;
        case 2001: // Small listbox 選択変更
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

        // Big tab 操作
        case 1010: { // 追加
            BigTab bt;
            bt.name = L"新規";
            g_settingsTabs.push_back(bt);
            g_settBigIdx = (int)g_settingsTabs.size() - 1;
            SettingsRefreshBigList();
            SettingsRefreshSmallList();
            break;
        }
        case 1011: { // 削除
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                g_settingsTabs.erase(g_settingsTabs.begin() + g_settBigIdx);
                if (g_settBigIdx >= (int)g_settingsTabs.size()) g_settBigIdx--;
                SettingsRefreshBigList();
                SettingsRefreshSmallList();
            }
            break;
        }
        case 1012: { // 上へ
            if (g_settBigIdx > 0) {
                std::swap(g_settingsTabs[g_settBigIdx], g_settingsTabs[g_settBigIdx - 1]);
                g_settBigIdx--;
                SettingsRefreshBigList();
            }
            break;
        }
        case 1013: { // 下へ
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size() - 1) {
                std::swap(g_settingsTabs[g_settBigIdx], g_settingsTabs[g_settBigIdx + 1]);
                g_settBigIdx++;
                SettingsRefreshBigList();
            }
            break;
        }
        case 1021: { // Big 名前保存
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                wchar_t buf[256]{};
                GetWindowTextW(g_sc.editBigName, buf, 256);
                g_settingsTabs[g_settBigIdx].name = buf;
                SettingsRefreshBigList();
            }
            break;
        }

        // Small tab 操作
        case 2010: { // 追加
            if (g_settBigIdx >= 0 && g_settBigIdx < (int)g_settingsTabs.size()) {
                SmallTab st; st.name = L"新規"; st.path = L"";
                g_settingsTabs[g_settBigIdx].smallTabs.push_back(st);
                SettingsRefreshSmallList();
            }
            break;
        }
        case 2011: { // 削除
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
        case 2012: { // Small 上へ
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
        case 2013: { // Small 下へ
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
        case 2021: { // Small 名前+パス保存
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
        case 2031: { // 参照...
            // BrowseForFolder
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

        case 9999: // 閉じる
            // 設定を反映
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
// QA パネル プロシージャ
//==============================================================================
static LRESULT CALLBACK QAPanelProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_COMMAND: {
        int id = LOWORD(wp);
        if (id >= QA_BASE_ID && id < QA_BASE_ID + (int)g_qaItems.size()) {
            int idx = id - QA_BASE_ID;
            PopulateFileList(g_qaItems[idx].path);
        }
        return 0;
    }
    case WM_CONTEXTMENU: {
        HWND hBtn = reinterpret_cast<HWND>(wp);
        for (int i = 0; i < (int)g_qaItems.size(); i++) {
            if (g_qaItems[i].btn == hBtn) {
                OnQAButtonRightClick(i);
                break;
            }
        }
        return 0;
    }
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

//==============================================================================
// 右ペイン プロシージャ
//==============================================================================
static LRESULT CALLBACK RightWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        HINSTANCE hi = GetModuleHandleW(nullptr);

        // Big tab
        g_hwndBigTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE | TCS_FIXEDWIDTH,
            0, 0, 400, BIG_TAB_H,
            hwnd, reinterpret_cast<HMENU>(101), hi, nullptr);

        // Small tab
        g_hwndSmallTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE | TCS_FIXEDWIDTH,
            0, BIG_TAB_H, 400, SMALL_TAB_H,
            hwnd, reinterpret_cast<HMENU>(102), hi, nullptr);

        // QA トグルボタン
        g_hwndQAToggle = CreateWindowExW(0, L"BUTTON",
            g_qaVisible ? L"▲ クイックアクセス" : L"▼ クイックアクセス",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, BIG_TAB_H + SMALL_TAB_H, 400, QA_BTN_H,
            hwnd, reinterpret_cast<HMENU>(103), hi, nullptr);

        // QA パネル (子ウィンドウ。QAPanelProc 使用)
        g_hwndQAPanel = CreateWindowExW(0, L"PreviewExplorer_QAPanel", nullptr,
            WS_CHILD | (g_qaVisible ? WS_VISIBLE : 0),
            0, BIG_TAB_H + SMALL_TAB_H + QA_BTN_H, 400, QA_PANEL_H,
            hwnd, nullptr, hi, nullptr);

        // ↑ 上へボタン
        g_hwndUpBtn = CreateWindowExW(0, L"BUTTON", L"↑ 上へ",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, UP_BTN_W, NAV_H,
            hwnd, reinterpret_cast<HMENU>(104), hi, nullptr);

        // パスラベル
        g_hwndPathLabel = CreateWindowExW(WS_EX_CLIENTEDGE, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_LEFTNOWORDWRAP | SS_NOPREFIX,
            UP_BTN_W, 0, 300, NAV_H,
            hwnd, reinterpret_cast<HMENU>(105), hi, nullptr);

        // ListView
        g_hwndListView = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, nullptr,
            WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS,
            0, 0, 400, 300,
            hwnd, reinterpret_cast<HMENU>(106), hi, nullptr);

        // ListView 拡張スタイル
        ListView_SetExtendedListViewStyle(g_hwndListView,
            LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);

        // システムイメージリストをセット
        EnsureSysImageList();
        if (g_sysImgList)
            ListView_SetImageList(g_hwndListView, g_sysImgList, LVSIL_SMALL);

        // 列追加
        LVCOLUMNW col{};
        col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
        col.fmt  = LVCFMT_LEFT;
        col.cx   = 200; col.pszText = const_cast<wchar_t*>(L"名前");
        ListView_InsertColumn(g_hwndListView, 0, &col);
        col.cx   = 90;  col.pszText = const_cast<wchar_t*>(L"種類");
        ListView_InsertColumn(g_hwndListView, 1, &col);
        col.cx   = 80;  col.fmt = LVCFMT_RIGHT;
        col.pszText = const_cast<wchar_t*>(L"サイズ");
        ListView_InsertColumn(g_hwndListView, 2, &col);
        col.cx   = 130; col.fmt = LVCFMT_LEFT;
        col.pszText = const_cast<wchar_t*>(L"更新日時");
        ListView_InsertColumn(g_hwndListView, 3, &col);

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
        RebuildQAButtons();

        return 0;
    }

    case WM_SIZE:
        DoRightLayout(hwnd);
        return 0;

    case WM_COMMAND: {
        int id = LOWORD(wp);
        if (id == 103) {
            // QA トグル
            g_qaVisible = !g_qaVisible;
            SetWindowTextW(g_hwndQAToggle,
                g_qaVisible ? L"▲ クイックアクセス" : L"▼ クイックアクセス");
            DoRightLayout(hwnd);
            SaveState();
        } else if (id == 104) {
            // 上へ
            NavigateUp();
        }
        return 0;
    }

    case WM_NOTIFY: {
        auto* nmhdr = reinterpret_cast<NMHDR*>(lp);
        // Big タブ切り替え
        if (nmhdr->hwndFrom == g_hwndBigTab && nmhdr->code == TCN_SELCHANGE) {
            int sel = TabCtrl_GetCurSel(g_hwndBigTab);
            if (sel >= 0) {
                g_bigTabIdx = sel;
                RebuildSmallTabs();
                NavigateToCurrentTab();
            }
            return 0;
        }
        // Small タブ切り替え
        if (nmhdr->hwndFrom == g_hwndSmallTab && nmhdr->code == TCN_SELCHANGE) {
            int sel = TabCtrl_GetCurSel(g_hwndSmallTab);
            if (sel >= 0 && g_bigTabIdx < (int)g_bigTabs.size()) {
                g_bigTabs[g_bigTabIdx].smallTabIdx = sel;
                NavigateToCurrentTab();
            }
            return 0;
        }
        // ListView 通知
        if (nmhdr->hwndFrom == g_hwndListView) {
            if (nmhdr->code == NM_DBLCLK) {
                auto* nmi = reinterpret_cast<NMITEMACTIVATE*>(lp);
                OnListViewDblClick(nmi->iItem);
                return 0;
            }
            if (nmhdr->code == NM_RCLICK) {
                auto* nmi = reinterpret_cast<NMITEMACTIVATE*>(lp);
                OnListViewRightClick(hwnd, nmi->iItem);
                return 0;
            }
            if (nmhdr->code == NM_CUSTOMDRAW) {
                auto* ncd = reinterpret_cast<NMLVCUSTOMDRAW*>(lp);
                switch (ncd->nmcd.dwDrawStage) {
                case CDDS_PREPAINT:
                    SetWindowLongPtrW(hwnd, DWLP_MSGRESULT, CDRF_NOTIFYITEMDRAW);
                    return CDRF_NOTIFYITEMDRAW;
                case CDDS_ITEMPREPAINT: {
                    int idx = static_cast<int>(ncd->nmcd.lItemlParam);
                    if (idx >= 0 && idx < (int)g_fileItems.size()) {
                        if (!g_fileItems[idx].supported) {
                            ncd->clrText = GetSysColor(COLOR_GRAYTEXT);
                        }
                    }
                    return CDRF_NEWFONT;
                }
                }
                return CDRF_DODEFAULT;
            }
        }
        break;
    }
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
    WritePrivateProfileStringW(L"Window", L"QuickAccessVisible",
        std::to_wstring(g_qaVisible ? 1 : 0).c_str(), ini.c_str());

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

    // Quick Access
    WritePrivateProfileStringW(L"QuickAccess", L"Count",
        std::to_wstring(g_qaItems.size()).c_str(), ini.c_str());
    for (int i = 0; i < (int)g_qaItems.size(); i++) {
        WritePrivateProfileStringW(L"QuickAccess",
            (L"Path_" + std::to_wstring(i)).c_str(),
            g_qaItems[i].path.c_str(), ini.c_str());
    }
}

static void LoadState() {
    auto ini = GetIniPath();

    g_splitX  = GetPrivateProfileIntW(L"Window", L"SplitX", 400, ini.c_str());
    g_qaVisible = GetPrivateProfileIntW(L"Window", L"QuickAccessVisible", 1, ini.c_str()) != 0;

    int bigCount = GetPrivateProfileIntW(L"Tabs", L"BigTabCount", 0, ini.c_str());
    g_bigTabIdx  = GetPrivateProfileIntW(L"Tabs", L"BigTabIndex",  0, ini.c_str());

    g_bigTabs.clear();

    if (bigCount == 0) {
        // デフォルト: 1つ作成
        BigTab bt;
        bt.name = L"動画";
        SmallTab st; st.name = L"サンプル"; st.path = L"";
        bt.smallTabs.push_back(st);
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

    // Quick Access
    int qaCount = GetPrivateProfileIntW(L"QuickAccess", L"Count", 0, ini.c_str());
    g_qaItems.clear();
    for (int i = 0; i < qaCount; i++) {
        wchar_t pbuf[MAX_PATH]{};
        GetPrivateProfileStringW(L"QuickAccess",
            (L"Path_" + std::to_wstring(i)).c_str(),
            L"", pbuf, MAX_PATH, ini.c_str());
        QAItem qa; qa.path = pbuf;
        g_qaItems.push_back(qa);
    }
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
    host->set_plugin_information(L"Preview edit + Explorer version 0.1");

    INITCOMMONCONTROLSEX icex{ sizeof(icex), ICC_WIN95_CLASSES | ICC_LISTVIEW_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&icex);

    HINSTANCE hi = GetModuleHandleW(nullptr);

    // Main ウィンドウクラス
    {
        WNDCLASSEXW wcex{};
        wcex.cbSize        = sizeof(wcex);
        wcex.lpszClassName = WC_MAIN;
        wcex.lpfnWndProc   = MainWndProc;
        wcex.hInstance     = hi;
        wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wcex.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wcex);
    }
    // Right ウィンドウクラス
    {
        WNDCLASSEXW wcex{};
        wcex.cbSize        = sizeof(wcex);
        wcex.lpszClassName = WC_RIGHT;
        wcex.lpfnWndProc   = RightWndProc;
        wcex.hInstance     = hi;
        wcex.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
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
    // QA パネルクラス
    {
        WNDCLASSEXW wcex{};
        wcex.cbSize        = sizeof(wcex);
        wcex.lpszClassName = L"PreviewExplorer_QAPanel";
        wcex.lpfnWndProc   = QAPanelProc;
        wcex.hInstance     = hi;
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

    host->register_window_client(L"Preview edit + Explorer", g_hwndMain);
    g_editHandle = host->create_edit_handle();
    host->register_config_menu(L"Preview edit + Explorer", OpenSettingsCallback);
}
