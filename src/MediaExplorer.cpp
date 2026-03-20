//----------------------------------------------------------------------------------
//  PreviewExplorer - メインエントリ・プラグインエクスポート
//----------------------------------------------------------------------------------
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")

#include "MediaExplorer.h"

//------------------------------------------------------------------------------
// グローバル定義
//------------------------------------------------------------------------------
EDIT_HANDLE*   g_editHandle = nullptr;
LOG_HANDLE*    g_logger     = nullptr;
CONFIG_HANDLE* g_config     = nullptr;

HWND g_hwndMain     = nullptr;
HWND g_hwndLeft     = nullptr;
HWND g_hwndRight    = nullptr;
HWND g_hwndInfo     = nullptr;
HWND g_hwndEmbedded = nullptr;

HWND g_hwndBigTab   = nullptr;
HWND g_hwndSmallTab = nullptr;
HWND g_hwndBtnBack  = nullptr;
HWND g_hwndBtnFwd   = nullptr;
HWND g_hwndBtnUp    = nullptr;
HWND g_hwndAddrEdit = nullptr;

IExplorerBrowser* g_peb = nullptr;

std::vector<BigTab> g_bigTabs;
int    g_bigTabIdx  = 0;
double g_splitRatio = 0.4;
int    g_splitX     = 400;
bool   g_dragging   = false;

int g_bigTabH   = 36;
int g_smallTabH = 24;
int g_splitterW = 5;

COLORREF g_clrBg   = RGB(0x20, 0x20, 0x20);
COLORREF g_clrText = RGB(0xFF, 0xFF, 0xFF);
HBRUSH   g_hBrushBg    = nullptr;
HFONT    g_hFontBig    = nullptr;
HFONT    g_hFontNormal = nullptr;

constexpr UINT_PTR TIMER_FIND_PREVIEW = 1;

//------------------------------------------------------------------------------
// ログ
//------------------------------------------------------------------------------
void Log(const wchar_t* msg) {
    if (g_logger) g_logger->log(g_logger, msg);
}

//------------------------------------------------------------------------------
// ダークモード (uxtheme 未公開 API)
//------------------------------------------------------------------------------
// ordinal 135: SetPreferredAppMode(mode) ─ 0=Default, 1=AllowDark, 2=ForceDark
// ordinal 133: AllowDarkModeForWindow(hwnd, allow)
// ordinal 104: RefreshImmersiveColorPolicyState()
static void InitDarkMode() {
    HMODULE hUx = GetModuleHandleW(L"uxtheme.dll");
    if (!hUx) hUx = LoadLibraryW(L"uxtheme.dll");
    if (!hUx) return;

    using fnSetMode = int(WINAPI*)(int);
    auto pSetMode = reinterpret_cast<fnSetMode>(GetProcAddress(hUx, MAKEINTRESOURCEA(135)));
    if (pSetMode) pSetMode(2);  // ForceDark

    using fnRefresh = void(WINAPI*)();
    auto pRefresh = reinterpret_cast<fnRefresh>(GetProcAddress(hUx, MAKEINTRESOURCEA(104)));
    if (pRefresh) pRefresh();
}

//------------------------------------------------------------------------------
// テーマ初期化
//------------------------------------------------------------------------------
static void InitTheme() {
    LPCWSTR fontName = L"Yu Gothic UI";
    float   fontSize = 13.0f;
    if (g_config) {
        FONT_INFO* fi = g_config->get_font_info(g_config, "Control");
        if (fi && fi->name && fi->name[0]) {
            fontName = fi->name;
            if (fi->size > 0.0f) fontSize = fi->size;
        }
    }
    int logH = -MulDiv((int)fontSize, 96, 72);

    if (g_config) {
        int raw = g_config->get_color_code(g_config, "Background");
        if (raw) g_clrBg = (COLORREF)raw;
        raw = g_config->get_color_code(g_config, "Text");
        if (raw) g_clrText = (COLORREF)raw;
    }

    if (g_hBrushBg) { DeleteObject(g_hBrushBg); g_hBrushBg = nullptr; }
    g_hBrushBg = CreateSolidBrush(g_clrBg);

    if (g_hFontBig)    { DeleteObject(g_hFontBig);    g_hFontBig    = nullptr; }
    if (g_hFontNormal) { DeleteObject(g_hFontNormal); g_hFontNormal = nullptr; }
    // logH は負値 (character height)。logH - N は絶対値が増えて「大きく」なるため
    // 大タブ: fontSize (13pt), 小タブ: fontSize - 1 (12pt)
    int bigH    = -MulDiv((int)fontSize,           96, 72);
    int normalH = -MulDiv(max(9, (int)fontSize-1), 96, 72);
    g_hFontBig    = CreateFontW(bigH,    0, 0, 0, FW_BOLD,   FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName);
    g_hFontNormal = CreateFontW(normalH, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName);
}

//------------------------------------------------------------------------------
// INI パス・ディレクトリ確保
//------------------------------------------------------------------------------
static std::wstring GetIniPath() {
    std::wstring base = g_config ? g_config->app_data_path : L".";
    return base + L"\\Plugin\\PreviewExplorer\\state.ini";
}

static void EnsureDir(const std::wstring& path) {
    auto pos = path.rfind(L'\\');
    if (pos == std::wstring::npos) return;
    EnsureDir(path.substr(0, pos));
    CreateDirectoryW(path.substr(0, pos).c_str(), nullptr);
}

//------------------------------------------------------------------------------
// 状態保存
//------------------------------------------------------------------------------
void SaveState() {
    auto ini = GetIniPath();
    EnsureDir(ini);

    wchar_t buf[64];
    swprintf_s(buf, L"%.4f", g_splitRatio);
    WritePrivateProfileStringW(L"Window", L"SplitRatio", buf, ini.c_str());

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
        WritePrivateProfileStringW(sec.c_str(), L"Name", g_bigTabs[i].name.c_str(), ini.c_str());
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

//------------------------------------------------------------------------------
// 状態読み込み
//------------------------------------------------------------------------------
void LoadState() {
    auto ini = GetIniPath();

    wchar_t buf[64]{};
    GetPrivateProfileStringW(L"Window", L"SplitRatio", L"0.4000", buf, 64, ini.c_str());
    g_splitRatio = _wtof(buf);
    if (g_splitRatio < 0.05 || g_splitRatio > 0.95) g_splitRatio = 0.4;

    int bigCount = GetPrivateProfileIntW(L"Tabs", L"BigTabCount", 0, ini.c_str());
    g_bigTabIdx  = GetPrivateProfileIntW(L"Tabs", L"BigTabIndex",  0, ini.c_str());
    g_bigTabs.clear();

    if (bigCount == 0) {
        BigTab bt; bt.name = L"動画";
        bt.smallTabs.push_back({ L"クイックアクセス", L"" });
        bt.smallTabs.push_back({ L"素材フォルダ",     L"" });
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
                    (L"SmallTab_" + std::to_wstring(j) + L"_Name").c_str(), L"", sn, 256, ini.c_str());
                GetPrivateProfileStringW(sec.c_str(),
                    (L"SmallTab_" + std::to_wstring(j) + L"_Path").c_str(), L"", sp, MAX_PATH, ini.c_str());
                bt.smallTabs.push_back({ sn, sp });
            }
            g_bigTabs.push_back(bt);
        }
    }
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) g_bigTabIdx = 0;
}

//------------------------------------------------------------------------------
// レイアウト（メインウィンドウ）
//------------------------------------------------------------------------------
void DoLayout(HWND hwnd) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    int W = rc.right, H = rc.bottom;

    g_splitX = max(MIN_LEFT_W, min(g_splitX, W - MIN_RIGHT_W - g_splitterW));

    SetWindowPos(g_hwndLeft,  nullptr, 0, 0, g_splitX, H, SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(g_hwndRight, nullptr, g_splitX + g_splitterW, 0,
        W - g_splitX - g_splitterW, H, SWP_NOZORDER | SWP_NOACTIVATE);

    if (g_hwndEmbedded && IsWindow(g_hwndEmbedded)) {
        RECT lrc;
        GetClientRect(g_hwndLeft, &lrc);
        SetWindowPos(g_hwndEmbedded, nullptr, 0, 0, lrc.right, lrc.bottom,
            SWP_NOZORDER | SWP_NOACTIVATE);
    }
}

//------------------------------------------------------------------------------
// プレビューウィンドウ探索 & 埋め込み
//------------------------------------------------------------------------------
struct WndEntry { HWND hwnd; wchar_t cls[256]; wchar_t title[256]; int w, h; };
struct EnumCtx  { DWORD pid; std::vector<WndEntry> entries; };

static BOOL CALLBACK EnumTopCb(HWND hwnd, LPARAM lp) {
    auto* ctx = reinterpret_cast<EnumCtx*>(lp);
    DWORD pid = 0; GetWindowThreadProcessId(hwnd, &pid);
    if (pid != ctx->pid) return TRUE;
    WndEntry e{}; e.hwnd = hwnd;
    GetClassNameW(hwnd, e.cls, 256); GetWindowTextW(hwnd, e.title, 256);
    RECT rc{}; GetWindowRect(hwnd, &rc);
    e.w = rc.right - rc.left; e.h = rc.bottom - rc.top;
    ctx->entries.push_back(e);
    return TRUE;
}

static void TryEmbedPreview() {
    HWND hostMain = g_editHandle->get_host_app_window();
    std::wstring log;
    if (!hostMain) {
        log = L"get_host_app_window() = null\nAviUtl2 本体ウィンドウを取得できませんでした。";
        if (g_logger) g_logger->warn(g_logger, log.c_str());
        if (g_hwndInfo) SetWindowTextW(g_hwndInfo, log.c_str());
        return;
    }
    DWORD hostPid = 0; GetWindowThreadProcessId(hostMain, &hostPid);
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
    if (g_hwndInfo) SetWindowTextW(g_hwndInfo, log.c_str());

    if (preview) {
        RECT lrc{}; GetClientRect(g_hwndLeft, &lrc);
        LONG style = GetWindowLongW(preview, GWL_STYLE);
        SetWindowLongW(preview, GWL_STYLE,
            (style & ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME)) | WS_CHILD | WS_VISIBLE);
        SetParent(preview, g_hwndLeft);
        SetWindowPos(preview, HWND_TOP, 0, 0, lrc.right, lrc.bottom, SWP_SHOWWINDOW | SWP_FRAMECHANGED);
        g_hwndEmbedded = preview;
        if (g_hwndInfo) ShowWindow(g_hwndInfo, SW_HIDE);
        if (g_logger) g_logger->log(g_logger, L"SetParent 完了。");
    } else {
        log += L"\n[!] プレビューウィンドウが見つかりませんでした。\n";
        if (g_hwndInfo) SetWindowTextW(g_hwndInfo, log.c_str());
    }
}

//------------------------------------------------------------------------------
// メインウィンドウ プロシージャ
//------------------------------------------------------------------------------
static LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE:
        LoadState();

        g_hwndLeft = CreateWindowExW(0, WC_STATIC, nullptr,
            WS_VISIBLE | WS_CHILD | SS_BLACKRECT,
            0, 0, 400, 400, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        g_hwndInfo = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT",
            L"プレビューウィンドウを検索中...",
            WS_VISIBLE | WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            0, 0, 400, 400, g_hwndLeft, nullptr, GetModuleHandleW(nullptr), nullptr);

        g_hwndRight = CreateWindowExW(0, WC_RIGHT, nullptr,
            WS_VISIBLE | WS_CHILD,
            400 + g_splitterW, 0, 400, 400, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        SetTimer(hwnd, TIMER_FIND_PREVIEW, 2000, nullptr);
        return 0;

    case WM_TIMER:
        if (wp == TIMER_FIND_PREVIEW) {
            KillTimer(hwnd, TIMER_FIND_PREVIEW);
            TryEmbedPreview();
        }
        return 0;

    case WM_SIZE: {
        RECT rc; GetClientRect(hwnd, &rc);
        int W = rc.right;
        if (W > MIN_LEFT_W + MIN_RIGHT_W + g_splitterW)
            g_splitX = (int)(W * g_splitRatio);
        DoLayout(hwnd);
        return 0;
    }

    case WM_SETCURSOR: {
        POINT pt; GetCursorPos(&pt); ScreenToClient(hwnd, &pt);
        if (pt.x >= g_splitX && pt.x < g_splitX + g_splitterW) {
            SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
            return TRUE;
        }
        break;
    }

    case WM_LBUTTONDOWN: {
        int x = GET_X_LPARAM(lp);
        if (x >= g_splitX && x < g_splitX + g_splitterW) {
            g_dragging = true; SetCapture(hwnd);
        }
        return 0;
    }

    case WM_MOUSEMOVE:
        if (g_dragging) {
            RECT rc; GetClientRect(hwnd, &rc); int W = rc.right;
            g_splitX = GET_X_LPARAM(lp);
            g_splitX = max(MIN_LEFT_W, min(g_splitX, W - MIN_RIGHT_W - g_splitterW));
            if (W > 0) g_splitRatio = (double)g_splitX / W;
            DoLayout(hwnd);
        }
        return 0;

    case WM_LBUTTONUP:
        if (g_dragging) { g_dragging = false; ReleaseCapture(); SaveState(); }
        return 0;

    case WM_ERASEBKGND: {
        HDC hdc = reinterpret_cast<HDC>(wp); RECT rc; GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_WINDOW + 1));
        return 1;
    }

    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT: {
        HDC hdc = reinterpret_cast<HDC>(wp);
        SetTextColor(hdc, g_clrText); SetBkColor(hdc, g_clrBg);
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
EXTERN_C __declspec(dllexport) DWORD RequiredVersion()                { return 2003300; }
EXTERN_C __declspec(dllexport) void  InitializeLogger(LOG_HANDLE* h)   { g_logger = h; }
EXTERN_C __declspec(dllexport) void  InitializeConfig(CONFIG_HANDLE* h) { g_config = h; }
EXTERN_C __declspec(dllexport) bool  InitializePlugin(DWORD)            { return true; }
EXTERN_C __declspec(dllexport) void  UninitializePlugin()               {}

EXTERN_C __declspec(dllexport) void RegisterPlugin(HOST_APP_TABLE* host) {
    host->set_plugin_information(L"Preview edit + Explorer version 0.5");

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    // ダークモードをプロセス全体に適用
    InitDarkMode();

    InitTheme();

    INITCOMMONCONTROLSEX icex{ sizeof(icex), ICC_WIN95_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&icex);

    HINSTANCE hi = GetModuleHandleW(nullptr);

    // Main ウィンドウクラス
    {
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpszClassName = WC_MAIN;
        wc.lpfnWndProc   = MainWndProc;
        wc.hInstance     = hi;
        wc.hbrBackground = g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_WINDOW + 1);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wc);
    }
    // Right ウィンドウクラス
    {
        WNDCLASSEXW wc{};
        wc.cbSize        = sizeof(wc);
        wc.lpszClassName = WC_RIGHT;
        wc.lpfnWndProc   = RightWndProc;
        wc.hInstance     = hi;
        wc.hbrBackground = g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_BTNFACE + 1);
        wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
        RegisterClassExW(&wc);
    }
    g_hwndMain = CreateWindowExW(
        0, WC_MAIN, L"Preview edit + Explorer",
        WS_POPUP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        nullptr, nullptr, hi, nullptr);
    if (!g_hwndMain) return;

    // ダークモードを自ウィンドウにも適用
    BOOL dark = TRUE;
    DwmSetWindowAttribute(g_hwndMain, 20, &dark, sizeof(dark));

    host->register_window_client(L"Preview edit + Explorer", g_hwndMain);
    g_editHandle = host->create_edit_handle();
    // 設定メニューは廃止
}
