//----------------------------------------------------------------------------------
//  PreviewExplorer - 右ペイン (IExplorerBrowser + ナビゲーション)
//----------------------------------------------------------------------------------
#include "MediaExplorer.h"

// TabsSubclass.cpp で定義
LRESULT CALLBACK BigTabSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
LRESULT CALLBACK SmallTabSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);

// コントロール ID
constexpr int ID_BTN_BACK = 201;
constexpr int ID_BTN_FWD  = 202;
constexpr int ID_BTN_UP   = 203;
constexpr int ID_ADDR_EDIT = 204;

//------------------------------------------------------------------------------
// ダークモードをウィンドウとその子に適用
//------------------------------------------------------------------------------
static void ApplyDarkToWindow(HWND hwnd) {
    if (!hwnd) return;
    using fnAllow = BOOL(WINAPI*)(HWND, BOOL);
    static HMODULE hUx = []() -> HMODULE {
        HMODULE h = GetModuleHandleW(L"uxtheme.dll");
        return h ? h : LoadLibraryW(L"uxtheme.dll");
    }();
    static auto pAllow = hUx
        ? reinterpret_cast<fnAllow>(GetProcAddress(hUx, MAKEINTRESOURCEA(133))) : nullptr;

    BOOL dark = TRUE;
    DwmSetWindowAttribute(hwnd, 20, &dark, sizeof(dark));  // DWMWA_USE_IMMERSIVE_DARK_MODE
    if (pAllow) pAllow(hwnd, TRUE);
    SetWindowTheme(hwnd, L"DarkMode_Explorer", nullptr);
}

static BOOL CALLBACK ApplyDarkChildCb(HWND h, LPARAM) {
    ApplyDarkToWindow(h);
    return TRUE;
}

//------------------------------------------------------------------------------
// IExplorerBrowserEvents 実装
//  - ナビゲーション完了時にアドレスバーを更新
//  - ビュー生成時にダークモードを適用
//------------------------------------------------------------------------------
class CExplorerEvents : public IExplorerBrowserEvents {
    volatile LONG m_ref = 1;
public:
    DWORD m_cookie = 0;

    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IExplorerBrowserEvents)) {
            *ppv = static_cast<IExplorerBrowserEvents*>(this);
            AddRef(); return S_OK;
        }
        *ppv = nullptr; return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef()  override { return InterlockedIncrement(&m_ref); }
    STDMETHODIMP_(ULONG) Release() override { return InterlockedDecrement(&m_ref); }

    STDMETHODIMP OnNavigationPending(PCIDLIST_ABSOLUTE) override { return S_OK; }

    STDMETHODIMP OnViewCreated(IShellView* psv) override {
        // シェルビューウィンドウにダークテーマ適用
        HWND hwndSV = nullptr;
        if (SUCCEEDED(psv->GetWindow(&hwndSV)) && hwndSV) {
            ApplyDarkToWindow(hwndSV);
            EnumChildWindows(hwndSV, ApplyDarkChildCb, 0);
        }
        return S_OK;
    }

    STDMETHODIMP OnNavigationComplete(PCIDLIST_ABSOLUTE pidl) override {
        if (!g_hwndAddrEdit || !pidl) return S_OK;
        // ファイルシステムパスを取得。仮想フォルダは表示名を使用
        wchar_t path[MAX_PATH]{};
        SHGetPathFromIDListW(pidl, path);
        if (path[0] == L'\0') {
            // 仮想フォルダ: 表示名を取得
            IShellItem* psi = nullptr;
            if (SUCCEEDED(SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psi)))) {
                LPWSTR name = nullptr;
                if (SUCCEEDED(psi->GetDisplayName(SIGDN_NORMALDISPLAY, &name)) && name) {
                    wcsncpy_s(path, name, MAX_PATH - 1);
                    CoTaskMemFree(name);
                }
                psi->Release();
            }
        }
        SetWindowTextW(g_hwndAddrEdit, path);
        return S_OK;
    }

    STDMETHODIMP OnNavigationFailed(PCIDLIST_ABSOLUTE) override { return S_OK; }
};

static CExplorerEvents g_ebEvents;

//------------------------------------------------------------------------------
// キーボードフック (AviUtl の TranslateAccelerator によるCtrl+C/V横取り対策)
//------------------------------------------------------------------------------
static HHOOK g_hKeyHook = nullptr;

// フック内で PostMessage → RightWndProc で実行
constexpr UINT WM_EXEC_SHELL_CMD = WM_APP + 200;

static bool IsFocusInRightPane() {
    if (!g_hwndRight) return false;
    HWND hFg = GetForegroundWindow();
    if (!hFg) return false;
    DWORD fgPid = 0; GetWindowThreadProcessId(hFg, &fgPid);
    if (fgPid != GetCurrentProcessId()) return false;
    HWND hFocus = GetFocus();
    if (!hFocus) return false;
    return hFocus == g_hwndRight || IsChild(g_hwndRight, hFocus);
}

static void ExecShellCmd(ULONG cmdId) {
    if (!g_peb) return;
    IShellView* psv = nullptr;
    if (SUCCEEDED(g_peb->GetCurrentView(IID_PPV_ARGS(&psv)))) {
        IOleCommandTarget* pct = nullptr;
        if (SUCCEEDED(psv->QueryInterface(IID_PPV_ARGS(&pct)))) {
            pct->Exec(nullptr, cmdId, OLECMDEXECOPT_DONTPROMPTUSER, nullptr, nullptr);
            pct->Release();
        }
        psv->Release();
    }
}

static LRESULT CALLBACK LowLevelKeyHook(int nCode, WPARAM wp, LPARAM lp) {
    if (nCode == HC_ACTION && (wp == WM_KEYDOWN || wp == WM_SYSKEYDOWN)) {
        auto* khs = reinterpret_cast<KBDLLHOOKSTRUCT*>(lp);
        bool ctrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        if (ctrl && IsFocusInRightPane()) {
            HWND hFocus = GetFocus();
            UINT vk = khs->vkCode;
            if (hFocus == g_hwndAddrEdit) {
                // アドレスバー: テキスト編集ショートカット (SendMessage は同スレッドなので安全)
                if      (vk == 'C') { SendMessageW(hFocus, WM_COPY,   0, 0); return 1; }
                else if (vk == 'X') { SendMessageW(hFocus, WM_CUT,    0, 0); return 1; }
                else if (vk == 'V') { SendMessageW(hFocus, WM_PASTE,  0, 0); return 1; }
                else if (vk == 'Z') { SendMessageW(hFocus, WM_UNDO,   0, 0); return 1; }
                else if (vk == 'A') { SendMessageW(hFocus, EM_SETSEL, 0, -1); return 1; }
            } else if (g_hwndRight) {
                // シェルビュー: PostMessage でフック外のコンテキストで COM 呼び出し
                if      (vk == 'C') { PostMessageW(g_hwndRight, WM_EXEC_SHELL_CMD, OLECMDID_COPY,      0); return 1; }
                else if (vk == 'X') { PostMessageW(g_hwndRight, WM_EXEC_SHELL_CMD, OLECMDID_CUT,       0); return 1; }
                else if (vk == 'V') { PostMessageW(g_hwndRight, WM_EXEC_SHELL_CMD, OLECMDID_PASTE,     0); return 1; }
                else if (vk == 'A') { PostMessageW(g_hwndRight, WM_EXEC_SHELL_CMD, OLECMDID_SELECTALL, 0); return 1; }
            }
        }
    }
    return CallNextHookEx(g_hKeyHook, nCode, wp, lp);
}

//------------------------------------------------------------------------------
// IExplorerBrowser ナビゲーション
//------------------------------------------------------------------------------
void NavigateExplorerToCurrentTab() {
    if (!g_peb) return;
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    auto& bt = g_bigTabs[g_bigTabIdx];
    if (bt.smallTabIdx < 0 || bt.smallTabIdx >= (int)bt.smallTabs.size()) return;
    const std::wstring& path = bt.smallTabs[bt.smallTabIdx].path;

    if (path.empty() || path == L"QuickAccess") {
        PIDLIST_ABSOLUTE pidl = nullptr;
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
// 現在のパスを取得
//------------------------------------------------------------------------------
std::wstring GetCurrentExplorerPath() {
    if (!g_peb) return L"";
    std::wstring result;
    IShellView* psv = nullptr;
    if (FAILED(g_peb->GetCurrentView(IID_PPV_ARGS(&psv)))) return L"";
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
    return result;
}

//------------------------------------------------------------------------------
// 戻る / 進む / 上へ
//------------------------------------------------------------------------------
static void NavBack() {
    if (!g_peb) return;
    IShellBrowser* psb = nullptr;
    if (SUCCEEDED(g_peb->QueryInterface(IID_PPV_ARGS(&psb)))) {
        psb->BrowseObject(nullptr, SBSP_NAVIGATEBACK);
        psb->Release();
    }
}

static void NavForward() {
    if (!g_peb) return;
    IShellBrowser* psb = nullptr;
    if (SUCCEEDED(g_peb->QueryInterface(IID_PPV_ARGS(&psb)))) {
        psb->BrowseObject(nullptr, SBSP_NAVIGATEFORWARD);
        psb->Release();
    }
}

static void NavUp() {
    if (!g_peb) return;
    std::wstring path = GetCurrentExplorerPath();
    if (path.empty()) return;
    // 末尾の '\' を除去 (ルートドライブ以外)
    while (path.size() > 3 && path.back() == L'\\') path.pop_back();
    auto pos = path.rfind(L'\\');
    if (pos == std::wstring::npos) return;
    std::wstring parent = path.substr(0, pos);
    if (parent.size() == 2 && parent[1] == L':') parent += L'\\';  // "C:" -> "C:\"
    if (parent.empty()) return;
    IShellItem* psi = nullptr;
    if (SUCCEEDED(SHCreateItemFromParsingName(parent.c_str(), nullptr, IID_PPV_ARGS(&psi)))) {
        g_peb->BrowseToObject(psi, SBSP_ABSOLUTE);
        psi->Release();
    }
}

//------------------------------------------------------------------------------
// アドレスバー Enter キー → ナビゲーション
//------------------------------------------------------------------------------
static LRESULT CALLBACK AddrEditSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                              UINT_PTR, DWORD_PTR) {
    if (msg == WM_KEYDOWN && wp == VK_RETURN) {
        wchar_t buf[MAX_PATH]{};
        GetWindowTextW(hwnd, buf, MAX_PATH);
        if (buf[0] && g_peb) {
            IShellItem* psi = nullptr;
            HRESULT hr = SHCreateItemFromParsingName(buf, nullptr, IID_PPV_ARGS(&psi));
            if (SUCCEEDED(hr)) {
                g_peb->BrowseToObject(psi, SBSP_ABSOLUTE);
                psi->Release();
            } else {
                // 入力が無効: 現在のパスに戻す
                SetWindowTextW(hwnd, GetCurrentExplorerPath().c_str());
            }
        }
        // フォーカスをエクスプローラービューへ
        if (g_hwndRight) SetFocus(g_hwndRight);
        return 0;
    }
    return DefSubclassProc(hwnd, msg, wp, lp);
}

//------------------------------------------------------------------------------
// スモールタブ再構築
//------------------------------------------------------------------------------
void RebuildSmallTabs() {
    if (!g_hwndSmallTab) return;
    TabCtrl_DeleteAllItems(g_hwndSmallTab);
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    auto& bt = g_bigTabs[g_bigTabIdx];
    for (int i = 0; i < (int)bt.smallTabs.size(); i++) {
        TCITEMW tci{}; tci.mask = TCIF_TEXT;
        tci.pszText = const_cast<wchar_t*>(bt.smallTabs[i].name.c_str());
        TabCtrl_InsertItem(g_hwndSmallTab, i, &tci);
    }
    if (bt.smallTabIdx >= (int)bt.smallTabs.size()) bt.smallTabIdx = (int)bt.smallTabs.size() - 1;
    if (bt.smallTabIdx < 0) bt.smallTabIdx = 0;
    if (!bt.smallTabs.empty()) TabCtrl_SetCurSel(g_hwndSmallTab, bt.smallTabIdx);
}

//------------------------------------------------------------------------------
// 右ペインのレイアウト
//------------------------------------------------------------------------------
void DoRightLayout(HWND hwnd) {
    RECT rc; GetClientRect(hwnd, &rc);
    int W = rc.right, H = rc.bottom;
    int y = 0;

    if (g_hwndBigTab) {
        SetWindowPos(g_hwndBigTab, nullptr, 0, y, W, g_bigTabH, SWP_NOZORDER | SWP_NOACTIVATE);
        y += g_bigTabH;
    }
    if (g_hwndSmallTab) {
        SetWindowPos(g_hwndSmallTab, nullptr, 0, y, W, g_smallTabH, SWP_NOZORDER | SWP_NOACTIVATE);
        y += g_smallTabH;
    }

    // ナビゲーションバー: [←][→][↑][アドレス入力]
    if (g_hwndBtnBack) {
        constexpr int GAP = 2;
        int x = GAP;
        SetWindowPos(g_hwndBtnBack,  nullptr, x, y + 2, NAV_BTN_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); x += NAV_BTN_W + GAP;
        SetWindowPos(g_hwndBtnFwd,   nullptr, x, y + 2, NAV_BTN_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); x += NAV_BTN_W + GAP;
        SetWindowPos(g_hwndBtnUp,    nullptr, x, y + 2, NAV_BTN_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); x += NAV_BTN_W + GAP;
        SetWindowPos(g_hwndAddrEdit, nullptr, x, y + 3, W - x - GAP, NAV_BAR_H - 6, SWP_NOZORDER | SWP_NOACTIVATE);
        y += NAV_BAR_H;
    }

    if (g_peb) {
        RECT ebRect = {0, y, W, H};
        g_peb->SetRect(nullptr, ebRect);
    }
}

//------------------------------------------------------------------------------
// 右ペイン プロシージャ
//------------------------------------------------------------------------------
LRESULT CALLBACK RightWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE: {
        HINSTANCE hi = GetModuleHandleW(nullptr);

        // ── ビッグタブ ──
        g_hwndBigTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE | TCS_HOTTRACK,
            0, 0, 100, 36, hwnd, reinterpret_cast<HMENU>(101), hi, nullptr);
        if (g_hFontBig) SendMessageW(g_hwndBigTab, WM_SETFONT, (WPARAM)g_hFontBig, TRUE);

        // ── スモールタブ ──
        g_hwndSmallTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE,
            0, 36, 100, 24, hwnd, reinterpret_cast<HMENU>(102), hi, nullptr);
        if (g_hFontNormal) SendMessageW(g_hwndSmallTab, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

        SetWindowSubclass(g_hwndBigTab,   BigTabSubclassProc,   101, 0);
        SetWindowSubclass(g_hwndSmallTab, SmallTabSubclassProc, 102, 0);

        // WM_PAINT で完全自前描画するため visual styles は不問
        // (SetWindowTheme を呼ばない方が NM_CUSTOMDRAW との競合を避けられる)

        // ビッグタブにアイテム追加
        for (int i = 0; i < (int)g_bigTabs.size(); i++) {
            TCITEMW tci{}; tci.mask = TCIF_TEXT;
            tci.pszText = const_cast<wchar_t*>(g_bigTabs[i].name.c_str());
            TabCtrl_InsertItem(g_hwndBigTab, i, &tci);
        }
        if (!g_bigTabs.empty()) TabCtrl_SetCurSel(g_hwndBigTab, g_bigTabIdx);
        RebuildSmallTabs();

        // タブ高さを動的計測 (フォント適用後に実際の高さを取得)
        {
            RECT tr{};
            if (g_hwndBigTab && TabCtrl_GetItemRect(g_hwndBigTab, 0, &tr) && tr.bottom > 0)
                g_bigTabH = tr.bottom + 6;
            if (g_hwndSmallTab && TabCtrl_GetItemRect(g_hwndSmallTab, 0, &tr) && tr.bottom > 0)
                g_smallTabH = tr.bottom + 4;
        }

        // ── ナビゲーションバー ──
        g_hwndBtnBack = CreateWindowExW(0, L"BUTTON", L"←",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, NAV_BTN_W, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_BTN_BACK), hi, nullptr);
        g_hwndBtnFwd = CreateWindowExW(0, L"BUTTON", L"→",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, NAV_BTN_W, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_BTN_FWD), hi, nullptr);
        g_hwndBtnUp = CreateWindowExW(0, L"BUTTON", L"↑",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, NAV_BTN_W, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_BTN_UP), hi, nullptr);
        g_hwndAddrEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", nullptr,
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            0, 0, 200, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_ADDR_EDIT), hi, nullptr);

        // ナビバーは小さいのでシステムデフォルトフォントを使用
        // (g_hFontNormal はタブ専用。アドレスバーに適用すると大きすぎる)
        SetWindowSubclass(g_hwndAddrEdit, AddrEditSubclassProc, ID_ADDR_EDIT, 0);

        // ボタン・アドレスバーにダークテーマを適用
        for (HWND h : {g_hwndBtnBack, g_hwndBtnFwd, g_hwndBtnUp}) {
            ApplyDarkToWindow(h);
            SetWindowTheme(h, L"DarkMode_Explorer", nullptr);
        }
        ApplyDarkToWindow(g_hwndAddrEdit);
        SetWindowTheme(g_hwndAddrEdit, L"DarkMode_Explorer", nullptr);

        // キーボードフック設置 (AviUtl のアクセラレータによる横取り対策)
        if (!g_hKeyHook)
            g_hKeyHook = SetWindowsHookExW(WH_KEYBOARD_LL, LowLevelKeyHook, nullptr, 0);

        // ── IExplorerBrowser 初期化 ──
        HRESULT hr = CoCreateInstance(CLSID_ExplorerBrowser, nullptr,
            CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&g_peb));
        if (SUCCEEDED(hr)) {
            RECT rc = {0, 0, 600, 600};
            FOLDERSETTINGS fs{};
            fs.ViewMode = FVM_DETAILS;
            fs.fFlags   = FWF_AUTOARRANGE | FWF_NOSUBFOLDERS & 0;  // デフォルト
            g_peb->Initialize(hwnd, &rc, &fs);
            g_peb->SetOptions(static_cast<EXPLORER_BROWSER_OPTIONS>(EBO_NOBORDER | EBO_ALWAYSNAVIGATE));

            // イベント登録
            IConnectionPointContainer* pcpc = nullptr;
            if (SUCCEEDED(g_peb->QueryInterface(IID_PPV_ARGS(&pcpc)))) {
                IConnectionPoint* pcp = nullptr;
                if (SUCCEEDED(pcpc->FindConnectionPoint(IID_IExplorerBrowserEvents, &pcp))) {
                    pcp->Advise(&g_ebEvents, &g_ebEvents.m_cookie);
                    pcp->Release();
                }
                pcpc->Release();
            }

            NavigateExplorerToCurrentTab();
        } else {
            Log(L"IExplorerBrowser の作成に失敗しました。");
        }
        return 0;
    }

    case WM_SIZE:
        DoRightLayout(hwnd);
        return 0;

    case WM_COMMAND: {
        int id = LOWORD(wp);
        if (id == ID_BTN_BACK) { NavBack();    return 0; }
        if (id == ID_BTN_FWD)  { NavForward(); return 0; }
        if (id == ID_BTN_UP)   { NavUp();      return 0; }
        break;
    }

    case WM_NOTIFY: {
        auto* nmhdr = reinterpret_cast<NMHDR*>(lp);

        // ── ボタン NM_CUSTOMDRAW : ダーク描画 ──
        if ((nmhdr->hwndFrom == g_hwndBtnBack || nmhdr->hwndFrom == g_hwndBtnFwd ||
             nmhdr->hwndFrom == g_hwndBtnUp)
            && nmhdr->code == NM_CUSTOMDRAW) {
            auto* nmcd = reinterpret_cast<NMCUSTOMDRAW*>(lp);
            if (nmcd->dwDrawStage == CDDS_PREPAINT) {
                auto clamp = [](int v) { return max(0, min(255, v)); };
                bool hot     = (nmcd->uItemState & CDIS_HOT)     != 0;
                bool pressed = (nmcd->uItemState & CDIS_SELECTED) != 0;
                COLORREF btnBg = pressed
                    ? RGB(clamp(GetRValue(g_clrBg) + 0x30), clamp(GetGValue(g_clrBg) + 0x38), clamp(GetBValue(g_clrBg) + 0x50))
                    : hot
                    ? RGB(clamp(GetRValue(g_clrBg) + 0x20), clamp(GetGValue(g_clrBg) + 0x20), clamp(GetBValue(g_clrBg) + 0x28))
                    : RGB(clamp(GetRValue(g_clrBg) + 0x10), clamp(GetGValue(g_clrBg) + 0x10), clamp(GetBValue(g_clrBg) + 0x10));
                SetDCBrushColor(nmcd->hdc, btnBg);
                FillRect(nmcd->hdc, &nmcd->rc, (HBRUSH)GetStockObject(DC_BRUSH));
                // 枠線
                COLORREF border = RGB(clamp(GetRValue(g_clrBg) + 0x40),
                                      clamp(GetGValue(g_clrBg) + 0x40),
                                      clamp(GetBValue(g_clrBg) + 0x40));
                SetDCPenColor(nmcd->hdc, border);
                HGDIOBJ oldPen = SelectObject(nmcd->hdc, GetStockObject(DC_PEN));
                HGDIOBJ oldBrush = SelectObject(nmcd->hdc, GetStockObject(NULL_BRUSH));
                Rectangle(nmcd->hdc, nmcd->rc.left, nmcd->rc.top, nmcd->rc.right, nmcd->rc.bottom);
                SelectObject(nmcd->hdc, oldPen);
                SelectObject(nmcd->hdc, oldBrush);
                // ラベル
                wchar_t text[8]{};
                GetWindowTextW(nmhdr->hwndFrom, text, 8);
                SetTextColor(nmcd->hdc, g_clrText);
                SetBkMode(nmcd->hdc, TRANSPARENT);
                RECT rc = nmcd->rc;
                DrawTextW(nmcd->hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                return CDRF_SKIPDEFAULT;
            }
            return CDRF_DODEFAULT;
        }

        if (nmhdr->hwndFrom == g_hwndBigTab && nmhdr->code == TCN_SELCHANGE) {
            int sel = TabCtrl_GetCurSel(g_hwndBigTab);
            if (sel >= 0) { g_bigTabIdx = sel; RebuildSmallTabs(); NavigateExplorerToCurrentTab(); }
            return 0;
        }
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

    case WM_ERASEBKGND: {
        HDC hdc = reinterpret_cast<HDC>(wp); RECT rc; GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_BTNFACE + 1));
        return 1;
    }

    case WM_EXEC_SHELL_CMD:
        ExecShellCmd(static_cast<ULONG>(wp));
        return 0;

    case WM_CTLCOLOREDIT: {
        // アドレスバーのダーク色
        HDC hdc = reinterpret_cast<HDC>(wp);
        SetTextColor(hdc, g_clrText);
        auto clamp = [](int v) { return max(0, min(255, v)); };
        COLORREF editBg = RGB(clamp(GetRValue(g_clrBg) + 0x18),
                              clamp(GetGValue(g_clrBg) + 0x18),
                              clamp(GetBValue(g_clrBg) + 0x18));
        SetBkColor(hdc, editBg);
        static HBRUSH s_hBrEdit = nullptr;
        static COLORREF s_lastEditBg = 0;
        if (!s_hBrEdit || s_lastEditBg != editBg) {
            if (s_hBrEdit) DeleteObject(s_hBrEdit);
            s_hBrEdit = CreateSolidBrush(editBg);
            s_lastEditBg = editBg;
        }
        return reinterpret_cast<LRESULT>(s_hBrEdit);
    }

    case WM_DESTROY:
        // キーボードフック解除
        if (g_hKeyHook) { UnhookWindowsHookEx(g_hKeyHook); g_hKeyHook = nullptr; }
        // イベント解除
        if (g_peb && g_ebEvents.m_cookie) {
            IConnectionPointContainer* pcpc = nullptr;
            if (SUCCEEDED(g_peb->QueryInterface(IID_PPV_ARGS(&pcpc)))) {
                IConnectionPoint* pcp = nullptr;
                if (SUCCEEDED(pcpc->FindConnectionPoint(IID_IExplorerBrowserEvents, &pcp))) {
                    pcp->Unadvise(g_ebEvents.m_cookie);
                    pcp->Release();
                }
                pcpc->Release();
            }
            g_ebEvents.m_cookie = 0;
        }
        if (g_peb) { g_peb->Destroy(); g_peb->Release(); g_peb = nullptr; }
        if (g_hFontBig)    { DeleteObject(g_hFontBig);    g_hFontBig    = nullptr; }
        if (g_hFontNormal) { DeleteObject(g_hFontNormal); g_hFontNormal = nullptr; }
        if (g_hBrushBg)    { DeleteObject(g_hBrushBg);    g_hBrushBg    = nullptr; }
        g_hwndBigTab = g_hwndSmallTab = nullptr;
        g_hwndBtnBack = g_hwndBtnFwd = g_hwndBtnUp = g_hwndAddrEdit = nullptr;
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}
