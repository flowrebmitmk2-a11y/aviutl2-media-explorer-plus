//----------------------------------------------------------------------------------
//  MediaExplorerPlus - right pane only window client
//----------------------------------------------------------------------------------
#define NOMINMAX
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")

#include "MediaExplorer.h"
#include <cwctype>
#include <map>
#include <optional>
#include <string_view>

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
double g_splitRatio = 1.0;
int    g_splitX     = 0;
bool   g_dragging   = false;

int g_bigTabH   = 36;
int g_smallTabH = 24;
int g_splitterW = 0;

COLORREF g_clrBg   = RGB(0x20, 0x20, 0x20);
COLORREF g_clrText = RGB(0xFF, 0xFF, 0xFF);
HBRUSH   g_hBrushBg    = nullptr;
HFONT    g_hFontBig    = nullptr;
HFONT    g_hFontNormal = nullptr;

void Log(const wchar_t* msg) {
    if (g_logger) g_logger->log(g_logger, msg);
}

static void InitDarkMode() {
    HMODULE hUx = GetModuleHandleW(L"uxtheme.dll");
    if (!hUx) hUx = LoadLibraryW(L"uxtheme.dll");
    if (!hUx) return;
    using fnSetMode = int(WINAPI*)(int);
    auto pSetMode = reinterpret_cast<fnSetMode>(GetProcAddress(hUx, MAKEINTRESOURCEA(135)));
    if (pSetMode) pSetMode(2);
    using fnRefresh = void(WINAPI*)();
    auto pRefresh = reinterpret_cast<fnRefresh>(GetProcAddress(hUx, MAKEINTRESOURCEA(104)));
    if (pRefresh) pRefresh();
}

static void InitTheme() {
    LPCWSTR fontName = L"Yu Gothic UI";
    float fontSize = 13.0f;
    if (g_config) {
        FONT_INFO* fi = g_config->get_font_info(g_config, "Control");
        if (fi && fi->name && fi->name[0]) {
            fontName = fi->name;
            if (fi->size > 0.0f) fontSize = fi->size;
        }
    }
    if (g_config) {
        int raw = g_config->get_color_code(g_config, "Background");
        if (raw) g_clrBg = (COLORREF)raw;
        raw = g_config->get_color_code(g_config, "Text");
        if (raw) g_clrText = (COLORREF)raw;
    }

    if (g_hBrushBg) DeleteObject(g_hBrushBg);
    g_hBrushBg = CreateSolidBrush(g_clrBg);

    if (g_hFontBig) DeleteObject(g_hFontBig);
    if (g_hFontNormal) DeleteObject(g_hFontNormal);

    int bigH = -MulDiv((int)fontSize, 96, 72);
    int normalH = -MulDiv(std::max(9, (int)fontSize - 1), 96, 72);
    g_hFontBig = CreateFontW(bigH, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName);
    g_hFontNormal = CreateFontW(normalH, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName);
}

static std::wstring GetIniPath() {
    std::wstring base = g_config ? g_config->app_data_path : L".";
    return base + L"\\Plugin\\MediaExplorerPlus\\state.ini";
}

static void EnsureDir(const std::wstring& path) {
    auto pos = path.rfind(L'\\');
    if (pos == std::wstring::npos) return;
    EnsureDir(path.substr(0, pos));
    CreateDirectoryW(path.substr(0, pos).c_str(), nullptr);
}

using IniSection = std::map<std::wstring, std::wstring>;
using IniData = std::map<std::wstring, IniSection>;

static std::wstring TrimCopy(std::wstring_view value) {
    size_t begin = 0;
    while (begin < value.size() && iswspace(value[begin])) ++begin;
    size_t end = value.size();
    while (end > begin && iswspace(value[end - 1])) --end;
    return std::wstring(value.substr(begin, end - begin));
}

static bool IsValidUtf8(const std::vector<unsigned char>& bytes) {
    int result = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
        reinterpret_cast<LPCCH>(bytes.data()), (int)bytes.size(), nullptr, 0);
    return result > 0;
}

static std::wstring DecodeIniBytes(const std::vector<unsigned char>& bytes) {
    if (bytes.empty()) return L"";

    if (bytes.size() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) {
        return std::wstring(reinterpret_cast<const wchar_t*>(bytes.data() + 2),
            (bytes.size() - 2) / sizeof(wchar_t));
    }
    if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
        int len = MultiByteToWideChar(CP_UTF8, 0,
            reinterpret_cast<LPCCH>(bytes.data() + 3), (int)bytes.size() - 3, nullptr, 0);
        std::wstring text(len, L'\0');
        if (len > 0) {
            MultiByteToWideChar(CP_UTF8, 0,
                reinterpret_cast<LPCCH>(bytes.data() + 3), (int)bytes.size() - 3,
                text.data(), len);
        }
        return text;
    }

    UINT codePage = IsValidUtf8(bytes) ? CP_UTF8 : 932;
    int len = MultiByteToWideChar(codePage, 0,
        reinterpret_cast<LPCCH>(bytes.data()), (int)bytes.size(), nullptr, 0);
    std::wstring text(len, L'\0');
    if (len > 0) {
        MultiByteToWideChar(codePage, 0,
            reinterpret_cast<LPCCH>(bytes.data()), (int)bytes.size(),
            text.data(), len);
    }
    return text;
}

static IniData LoadIniData(const std::wstring& path) {
    IniData data;
    HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE,
        nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) return data;

    LARGE_INTEGER size{};
    if (!GetFileSizeEx(file, &size) || size.QuadPart < 0 || size.QuadPart > INT_MAX) {
        CloseHandle(file);
        return data;
    }

    std::vector<unsigned char> bytes((size_t)size.QuadPart);
    DWORD read = 0;
    if (!bytes.empty() && !ReadFile(file, bytes.data(), (DWORD)bytes.size(), &read, nullptr)) {
        CloseHandle(file);
        return data;
    }
    CloseHandle(file);
    bytes.resize(read);

    std::wstring text = DecodeIniBytes(bytes);
    std::wstring currentSection;
    size_t pos = 0;
    while (pos <= text.size()) {
        size_t end = text.find_first_of(L"\r\n", pos);
        std::wstring line = TrimCopy(text.substr(pos, end == std::wstring::npos ? std::wstring::npos : end - pos));
        if (!line.empty() && line[0] != L';' && line[0] != L'#') {
            if (line.front() == L'[' && line.back() == L']' && line.size() >= 2) {
                currentSection = line.substr(1, line.size() - 2);
            } else {
                size_t eq = line.find(L'=');
                if (eq != std::wstring::npos) {
                    data[currentSection][TrimCopy(line.substr(0, eq))] = TrimCopy(line.substr(eq + 1));
                }
            }
        }
        if (end == std::wstring::npos) break;
        pos = end + 1;
        if (end + 1 < text.size() && text[end] == L'\r' && text[end + 1] == L'\n') ++pos;
    }
    return data;
}

static bool SaveIniData(const std::wstring& path, const IniData& data) {
    std::wstring text;
    for (const auto& [sectionName, section] : data) {
        if (!sectionName.empty()) {
            text += L"[";
            text += sectionName;
            text += L"]\r\n";
        }
        for (const auto& [key, value] : section) {
            text += key;
            text += L"=";
            text += value;
            text += L"\r\n";
        }
    }

    int bytesLen = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), nullptr, 0, nullptr, nullptr);
    std::string utf8(bytesLen, '\0');
    if (bytesLen > 0) {
        WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), utf8.data(), bytesLen, nullptr, nullptr);
    }

    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) return false;
    DWORD written = 0;
    BOOL ok = WriteFile(file, utf8.data(), (DWORD)utf8.size(), &written, nullptr);
    CloseHandle(file);
    return ok == TRUE;
}

static std::wstring IniGetString(const IniData& data, const std::wstring& section, const std::wstring& key, const std::wstring& def = L"") {
    auto secIt = data.find(section);
    if (secIt == data.end()) return def;
    auto valIt = secIt->second.find(key);
    if (valIt == secIt->second.end()) return def;
    return valIt->second;
}

static int IniGetInt(const IniData& data, const std::wstring& section, const std::wstring& key, int def = 0) {
    std::wstring value = IniGetString(data, section, key);
    if (value.empty()) return def;
    return _wtoi(value.c_str());
}

static int InferMaxIndexedSection(const IniData& data, const std::wstring& prefix) {
    int maxIndex = -1;
    for (const auto& [section, values] : data) {
        (void)values;
        if (section.rfind(prefix, 0) != 0) continue;
        std::wstring tail = section.substr(prefix.size());
        if (tail.empty()) continue;
        bool allDigits = std::all_of(tail.begin(), tail.end(), iswdigit);
        if (!allDigits) continue;
        maxIndex = std::max(maxIndex, _wtoi(tail.c_str()));
    }
    return maxIndex;
}

static int InferMaxIndexedKey(const IniSection& section, const std::wstring& prefix, const std::wstring& suffix) {
    int maxIndex = -1;
    for (const auto& [key, value] : section) {
        (void)value;
        if (key.rfind(prefix, 0) != 0) continue;
        if (key.size() < prefix.size() + suffix.size()) continue;
        if (key.substr(key.size() - suffix.size()) != suffix) continue;
        std::wstring middle = key.substr(prefix.size(), key.size() - prefix.size() - suffix.size());
        if (middle.empty()) continue;
        bool allDigits = std::all_of(middle.begin(), middle.end(), iswdigit);
        if (!allDigits) continue;
        maxIndex = std::max(maxIndex, _wtoi(middle.c_str()));
    }
    return maxIndex;
}

void SaveState() {
    auto ini = GetIniPath();
    EnsureDir(ini);
    IniData data;
    data[L"Tabs"][L"BigTabCount"] = std::to_wstring(g_bigTabs.size());
    data[L"Tabs"][L"BigTabIndex"] = std::to_wstring(g_bigTabIdx);
    for (int i = 0; i < (int)g_bigTabs.size(); i++) {
        data[L"Tabs"][L"SmallTabIndex_" + std::to_wstring(i)] = std::to_wstring(g_bigTabs[i].smallTabIdx);
    }
    for (int i = 0; i < (int)g_bigTabs.size(); i++) {
        std::wstring sec = L"BigTab_" + std::to_wstring(i);
        data[sec][L"Name"] = g_bigTabs[i].name;
        data[sec][L"SmallTabCount"] = std::to_wstring(g_bigTabs[i].smallTabs.size());
        for (int j = 0; j < (int)g_bigTabs[i].smallTabs.size(); j++) {
            auto& st = g_bigTabs[i].smallTabs[j];
            data[sec][L"SmallTab_" + std::to_wstring(j) + L"_Name"] = st.name;
            data[sec][L"SmallTab_" + std::to_wstring(j) + L"_Path"] = st.path;
            data[sec][L"SmallTab_" + std::to_wstring(j) + L"_CurrentPath"] = st.currentPath;
        }
    }
    SaveIniData(ini, data);
}

void LoadState() {
    auto ini = GetIniPath();
    IniData data = LoadIniData(ini);
    int declaredBigCount = IniGetInt(data, L"Tabs", L"BigTabCount", 0);
    int inferredBigCount = InferMaxIndexedSection(data, L"BigTab_") + 1;
    int bigCount = std::max(declaredBigCount, inferredBigCount);
    g_bigTabIdx = IniGetInt(data, L"Tabs", L"BigTabIndex", 0);
    g_bigTabs.clear();

    if (bigCount == 0) {
        BigTab bt;
        bt.name = L"作業";
        bt.smallTabs.push_back({L"クイックアクセス", L""});
        bt.smallTabs.push_back({L"動画フォルダ", L""});
        g_bigTabs.push_back(bt);
        g_bigTabIdx = 0;
    } else {
        for (int i = 0; i < bigCount; i++) {
            std::wstring sec = L"BigTab_" + std::to_wstring(i);
            std::wstring name = IniGetString(data, sec, L"Name");
            if (name.empty()) continue;
            BigTab bt;
            bt.name = name;
            bt.smallTabIdx = IniGetInt(data, L"Tabs", L"SmallTabIndex_" + std::to_wstring(i), 0);
            auto secIt = data.find(sec);
            int inferredSmallCount = -1;
            if (secIt != data.end()) {
                inferredSmallCount = std::max({
                    InferMaxIndexedKey(secIt->second, L"SmallTab_", L"_Name"),
                    InferMaxIndexedKey(secIt->second, L"SmallTab_", L"_Path"),
                    InferMaxIndexedKey(secIt->second, L"SmallTab_", L"_CurrentPath")
                });
            }
            int stCount = std::max(IniGetInt(data, sec, L"SmallTabCount", 0), inferredSmallCount + 1);
            for (int j = 0; j < stCount; j++) {
                std::wstring prefix = L"SmallTab_" + std::to_wstring(j);
                std::wstring tabName = IniGetString(data, sec, prefix + L"_Name");
                std::wstring tabPath = IniGetString(data, sec, prefix + L"_Path");
                std::wstring currentPath = IniGetString(data, sec, prefix + L"_CurrentPath");
                if (tabName.empty() && tabPath.empty() && currentPath.empty()) continue;
                bt.smallTabs.push_back({tabName, tabPath, currentPath});
            }
            if (bt.smallTabs.empty()) bt.smallTabs.push_back({L"クイックアクセス", L"", L""});
            g_bigTabs.push_back(bt);
        }
    }
    if (g_bigTabs.empty()) {
        BigTab bt;
        bt.name = L"作業";
        bt.smallTabs.push_back({L"クイックアクセス", L""});
        bt.smallTabs.push_back({L"動画フォルダ", L""});
        g_bigTabs.push_back(bt);
    }
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) g_bigTabIdx = 0;
}

void DoLayout(HWND hwnd) {
    RECT rc{};
    GetClientRect(hwnd, &rc);
    if (g_hwndRight) {
        SetWindowPos(g_hwndRight, nullptr, 0, 0, rc.right, rc.bottom, SWP_NOZORDER | SWP_NOACTIVATE);
    }
}

static LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE:
        LoadState();
        g_hwndRight = CreateWindowExW(0, WC_RIGHT, nullptr,
            WS_VISIBLE | WS_CHILD,
            0, 0, 400, 400, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        return 0;
    case WM_SIZE:
        DoLayout(hwnd);
        return 0;
    case WM_ERASEBKGND: {
        HDC hdc = reinterpret_cast<HDC>(wp);
        RECT rc{};
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_WINDOW + 1));
        return 1;
    }
    case WM_DESTROY:
        SaveState();
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

EXTERN_C __declspec(dllexport) DWORD RequiredVersion() { return 2003300; }
EXTERN_C __declspec(dllexport) void InitializeLogger(LOG_HANDLE* h) { g_logger = h; }
EXTERN_C __declspec(dllexport) void InitializeConfig(CONFIG_HANDLE* h) { g_config = h; }
EXTERN_C __declspec(dllexport) bool InitializePlugin(DWORD) { return true; }
EXTERN_C __declspec(dllexport) void UninitializePlugin() {}

EXTERN_C __declspec(dllexport) void RegisterPlugin(HOST_APP_TABLE* host) {
    host->set_plugin_information(L"MediaExplorerPlus version 0.1");

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    InitDarkMode();
    InitTheme();

    INITCOMMONCONTROLSEX icex{ sizeof(icex), ICC_WIN95_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&icex);

    HINSTANCE hi = GetModuleHandleW(nullptr);

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpszClassName = WC_MAIN;
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hi;
    wc.hbrBackground = g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    RegisterClassExW(&wc);

    wc = {};
    wc.cbSize = sizeof(wc);
    wc.lpszClassName = WC_RIGHT;
    wc.lpfnWndProc = RightWndProc;
    wc.hInstance = hi;
    wc.hbrBackground = g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_BTNFACE + 1);
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    RegisterClassExW(&wc);

    g_hwndMain = CreateWindowExW(
        0, WC_MAIN, L"Media Explorer Plus",
        WS_POPUP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        nullptr, nullptr, hi, nullptr);
    if (!g_hwndMain) return;

    BOOL dark = TRUE;
    DwmSetWindowAttribute(g_hwndMain, 20, &dark, sizeof(dark));

    host->register_window_client(L"Media Explorer Plus", g_hwndMain);
    g_editHandle = host->create_edit_handle();
}
