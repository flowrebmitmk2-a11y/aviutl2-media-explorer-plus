//----------------------------------------------------------------------------------
//  PreviewExplorer - 右ペイン (IExplorerBrowser + ナビゲーション)
//----------------------------------------------------------------------------------
#include "MediaExplorer.h"
#include <commdlg.h>

// TabsSubclass.cpp で定義
LRESULT CALLBACK BigTabSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
LRESULT CALLBACK SmallTabSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);

// コントロール ID
constexpr int ID_BTN_BACK  = 201;
constexpr int ID_BTN_FWD   = 202;
constexpr int ID_BTN_UP    = 203;
constexpr int ID_ADDR_EDIT = 204;
constexpr int ID_BTN_COPY  = 205;
constexpr int ID_BTN_CUT   = 206;
constexpr int ID_BTN_PASTE = 207;
constexpr int ID_BTN_ALIAS = 208;  // エイリアス(ショートカット)作成

// ナビゲーションバー右側ボタン (RightPane.cpp 内でのみ使用)
static HWND s_hwndBtnCopy  = nullptr;
static HWND s_hwndBtnCut   = nullptr;
static HWND s_hwndBtnPaste = nullptr;
static HWND s_hwndBtnAlias = nullptr;

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
        if (!pidl) return S_OK;
        // タイマーフォールバックが既に予約されていれば取り消す
        if (g_hwndRight) KillTimer(g_hwndRight, 2 /*TIMER_ADDR_UPDATE*/);

        // パース可能な形式でパスを取得 (FS = 普通のパス、仮想フォルダ = "::{GUID}...")
        wchar_t parsePath[MAX_PATH]{};
        wchar_t dispPath[MAX_PATH]{};
        IShellItem* psi = nullptr;
        if (SUCCEEDED(SHCreateItemFromIDList(pidl, IID_PPV_ARGS(&psi)))) {
            LPWSTR pp = nullptr;
            // ファイルシステムパスを優先取得
            if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pp)) && pp) {
                wcsncpy_s(parsePath, pp, MAX_PATH - 1);
                wcsncpy_s(dispPath,  pp, MAX_PATH - 1);
                CoTaskMemFree(pp);
            } else {
                // 仮想フォルダ: 表示名をアドレスバーに、パース名を保存用に
                if (SUCCEEDED(psi->GetDisplayName(SIGDN_NORMALDISPLAY, &pp)) && pp) {
                    wcsncpy_s(dispPath, pp, MAX_PATH - 1);
                    CoTaskMemFree(pp);
                }
                if (SUCCEEDED(psi->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &pp)) && pp) {
                    wcsncpy_s(parsePath, pp, MAX_PATH - 1);
                    CoTaskMemFree(pp);
                }
            }
            psi->Release();
        }

        // アドレスバー更新
        if (g_hwndAddrEdit) SetWindowTextW(g_hwndAddrEdit, dispPath);

        // 現在のタブの位置状態を更新 (タブごとの状態保存)
        if (parsePath[0] &&
            g_bigTabIdx >= 0 && g_bigTabIdx < (int)g_bigTabs.size()) {
            auto& bt = g_bigTabs[g_bigTabIdx];
            if (bt.smallTabIdx >= 0 && bt.smallTabIdx < (int)bt.smallTabs.size())
                bt.smallTabs[bt.smallTabIdx].currentPath = parsePath;
        }
        return S_OK;
    }

    STDMETHODIMP OnNavigationFailed(PCIDLIST_ABSOLUTE) override { return S_OK; }
};

static CExplorerEvents g_ebEvents;

// アドレスバー遅延更新タイマー ID
constexpr UINT_PTR TIMER_ADDR_UPDATE = 2;

//------------------------------------------------------------------------------
// アドレスバーを現在のフォルダに更新
//  OnNavigationComplete が届かない場合のフォールバック用
//------------------------------------------------------------------------------
static void UpdateAddrBar() {
    if (!g_hwndAddrEdit || !g_peb) return;
    wchar_t path[MAX_PATH]{};

    IShellView* psv = nullptr;
    if (FAILED(g_peb->GetCurrentView(IID_PPV_ARGS(&psv)))) return;

    IFolderView2* pfv2 = nullptr;
    if (SUCCEEDED(psv->QueryInterface(IID_PPV_ARGS(&pfv2)))) {
        IPersistFolder2* ppf2 = nullptr;
        if (SUCCEEDED(pfv2->GetFolder(IID_PPV_ARGS(&ppf2)))) {
            PIDLIST_ABSOLUTE pidl = nullptr;
            if (SUCCEEDED(ppf2->GetCurFolder(&pidl))) {
                if (!SHGetPathFromIDListW(pidl, path)) {
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
                CoTaskMemFree(pidl);
            }
            ppf2->Release();
        }
        pfv2->Release();
    }
    psv->Release();

    SetWindowTextW(g_hwndAddrEdit, path);
}

//------------------------------------------------------------------------------
// タイムライン上のフォーカスオブジェクトのエイリアスを保存
//  - 保存ダイアログの初期ディレクトリ = 現在表示中のフォルダ
//  - 保存名デフォルト = プロジェクト名 (拡張子除く) + ".object"
//  - エイリアスデータは EDIT_SECTION::get_object_alias() で取得 (UTF-8 テキスト)
//------------------------------------------------------------------------------

// call_edit_section_param のコールバックで使うキャプチャ構造体
struct AliasCaptureData {
    std::string aliasBytes;   // get_object_alias の戻り値 (UTF-8)
    std::wstring projectName; // プロジェクトファイル名 (拡張子なし, ダイアログデフォルト名用)
};

static void CaptureAliasCallback(void* param, EDIT_SECTION* edit) {
    if (!edit || !param) return;
    auto* cap = static_cast<AliasCaptureData*>(param);

    OBJECT_HANDLE obj = edit->get_focus_object();
    if (!obj) return;

    LPCSTR data = edit->get_object_alias(obj);
    if (data && data[0]) cap->aliasBytes = data;

    // プロジェクト名を取得 (ファイル名のデフォルト用)
    if (g_editHandle) {
        PROJECT_FILE* pf = edit->get_project_file(g_editHandle);
        if (pf) {
            LPCWSTR projPath = pf->get_project_file_path();
            if (projPath && projPath[0]) {
                wchar_t tmp[MAX_PATH];
                wcscpy_s(tmp, PathFindFileNameW(projPath));
                PathRemoveExtensionW(tmp);
                cap->projectName = tmp;
            }
        }
    }
}

static void CreateAliasForSelected() {
    if (!g_editHandle) return;

    AliasCaptureData cap;
    g_editHandle->call_edit_section_param(&cap, CaptureAliasCallback);

    if (cap.aliasBytes.empty()) {
        MessageBoxW(g_hwndRight,
            L"タイムラインでオブジェクトを選択してください。",
            L"エイリアス保存", MB_OK | MB_ICONINFORMATION);
        return;
    }

    // デフォルト保存名: <プロジェクト名>.object (なければ "alias.object")
    std::wstring defaultName = cap.projectName.empty() ? L"alias" : cap.projectName;
    defaultName += L".object";

    // 保存先ダイアログ (初期ディレクトリ = 現在のエクスプローラーフォルダ)
    wchar_t savePath[MAX_PATH]{};
    wcscpy_s(savePath, defaultName.c_str());

    OPENFILENAMEW ofn{};
    ofn.lStructSize  = sizeof(ofn);
    ofn.hwndOwner    = g_hwndRight;
    ofn.lpstrFilter  = L"エイリアスファイル (*.object)\0*.object\0すべてのファイル (*.*)\0*.*\0";
    ofn.lpstrFile    = savePath;
    ofn.nMaxFile     = MAX_PATH;
    ofn.lpstrTitle   = L"エイリアスを名前を付けて保存";
    ofn.lpstrDefExt  = L"object";
    ofn.Flags        = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    std::wstring initDir = GetCurrentExplorerPath();
    if (!initDir.empty()) ofn.lpstrInitialDir = initDir.c_str();

    if (GetSaveFileNameW(&ofn)) {
        HANDLE hFile = CreateFileW(savePath, GENERIC_WRITE, 0, nullptr,
                                   CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hFile != INVALID_HANDLE_VALUE) {
            DWORD written;
            WriteFile(hFile, cap.aliasBytes.c_str(),
                      static_cast<DWORD>(cap.aliasBytes.size()), &written, nullptr);
            CloseHandle(hFile);
        } else {
            MessageBoxW(g_hwndRight, L"ファイルの作成に失敗しました。",
                        L"エイリアス保存エラー", MB_OK | MB_ICONERROR);
        }
    }
}

//------------------------------------------------------------------------------
// ファイルクリップボード操作
//  op: 0=コピー, 1=切り取り, 2=貼り付け
//------------------------------------------------------------------------------
static void ExecFileOp(int op) {
    if (!g_peb) return;

    if (op == 2) {
        // 貼り付け: IFileOperation を使用
        //  同フォルダ → FOF_RENAMEONCOLLISION で自動リネーム ("～のコピー")
        //  別フォルダ → デフォルト (IFileOperation が「上書き / スキップ / 両方保存」ダイアログを表示)
        std::wstring dest = GetCurrentExplorerPath();
        if (dest.empty()) return;

        IDataObject* pdo = nullptr;
        if (FAILED(OleGetClipboard(&pdo))) return;

        FORMATETC fe = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM sm{};
        if (SUCCEEDED(pdo->GetData(&fe, &sm))) {
            HDROP hDrop = (HDROP)GlobalLock(sm.hGlobal);
            if (hDrop) {
                UINT n = DragQueryFileW(hDrop, 0xFFFFFFFF, nullptr, 0);
                if (n > 0) {
                    // Move か Copy か判定
                    bool isMove = false;
                    UINT cfPDE = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
                    FORMATETC fePDE = { (CLIPFORMAT)cfPDE, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
                    STGMEDIUM smPDE{};
                    if (SUCCEEDED(pdo->GetData(&fePDE, &smPDE))) {
                        DWORD* pe = (DWORD*)GlobalLock(smPDE.hGlobal);
                        if (pe) { isMove = (*pe & DROPEFFECT_MOVE) != 0; GlobalUnlock(smPDE.hGlobal); }
                        ReleaseStgMedium(&smPDE);
                    }

                    // 全ソースファイルが貼り付け先と同じフォルダか確認
                    bool allSameFolder = !isMove;  // Move は同フォルダでも通常動作
                    for (UINT i = 0; i < n && allSameFolder; i++) {
                        wchar_t path[MAX_PATH]{};
                        DragQueryFileW(hDrop, i, path, MAX_PATH);
                        wchar_t parent[MAX_PATH]{};
                        wcscpy_s(parent, path);
                        PathRemoveFileSpecW(parent);  // 末尾の "\" は除去される
                        if (_wcsicmp(parent, dest.c_str()) != 0) allSameFolder = false;
                    }

                    IFileOperation* pfo = nullptr;
                    if (SUCCEEDED(CoCreateInstance(CLSID_FileOperation, nullptr,
                                                   CLSCTX_ALL, IID_PPV_ARGS(&pfo)))) {
                        DWORD flags = FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR;
                        if (allSameFolder)
                            flags |= FOF_RENAMEONCOLLISION;  // 同フォルダ: "～のコピー" 形式で自動リネーム
                        // 別フォルダ: ダイアログで「上書き / スキップ / 両方保存」を選択
                        pfo->SetOperationFlags(flags);
                        pfo->SetOwnerWindow(g_hwndRight);

                        IShellItem* psiDest = nullptr;
                        if (SUCCEEDED(SHCreateItemFromParsingName(dest.c_str(), nullptr,
                                                                   IID_PPV_ARGS(&psiDest)))) {
                            for (UINT i = 0; i < n; i++) {
                                wchar_t path[MAX_PATH]{};
                                DragQueryFileW(hDrop, i, path, MAX_PATH);
                                IShellItem* psiSrc = nullptr;
                                if (SUCCEEDED(SHCreateItemFromParsingName(path, nullptr,
                                                                           IID_PPV_ARGS(&psiSrc)))) {
                                    if (isMove)
                                        pfo->MoveItem(psiSrc, psiDest, nullptr, nullptr);
                                    else
                                        pfo->CopyItem(psiSrc, psiDest, nullptr, nullptr);
                                    psiSrc->Release();
                                }
                            }
                            psiDest->Release();
                        }
                        pfo->PerformOperations();
                        pfo->Release();
                    }
                }
                GlobalUnlock(sm.hGlobal);
            }
            ReleaseStgMedium(&sm);
        }
        pdo->Release();
        return;
    }

    // コピー / 切り取り: 選択ファイルを IDataObject としてクリップボードに設定
    IShellView* psv = nullptr;
    if (FAILED(g_peb->GetCurrentView(IID_PPV_ARGS(&psv)))) return;

    IFolderView* pfv = nullptr;
    if (SUCCEEDED(psv->QueryInterface(IID_PPV_ARGS(&pfv)))) {
        IShellItemArray* psia = nullptr;
        if (SUCCEEDED(pfv->Items(SVGIO_SELECTION, IID_PPV_ARGS(&psia)))) {
            IDataObject* pdo = nullptr;
            if (SUCCEEDED(psia->BindToHandler(nullptr, BHID_DataObject, IID_PPV_ARGS(&pdo)))) {
                UINT cfPDE = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
                FORMATETC fe = { (CLIPFORMAT)cfPDE, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
                STGMEDIUM sm{};
                sm.tymed   = TYMED_HGLOBAL;
                sm.hGlobal = GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD));
                if (sm.hGlobal) {
                    DWORD* p = (DWORD*)GlobalLock(sm.hGlobal);
                    *p = (op == 1) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
                    GlobalUnlock(sm.hGlobal);
                    pdo->SetData(&fe, &sm, TRUE);
                }
                OleSetClipboard(pdo);
                pdo->Release();
            }
            psia->Release();
        }
        pfv->Release();
    }
    psv->Release();
}

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

static LRESULT CALLBACK LowLevelKeyHook(int nCode, WPARAM wp, LPARAM lp) {
    if (nCode == HC_ACTION && (wp == WM_KEYDOWN || wp == WM_SYSKEYDOWN)) {
        auto* khs = reinterpret_cast<KBDLLHOOKSTRUCT*>(lp);
        bool ctrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
        if (ctrl && IsFocusInRightPane()) {
            HWND hFocus = GetFocus();
            UINT vk = khs->vkCode;
            if (hFocus == g_hwndAddrEdit) {
                // アドレスバー: テキスト編集ショートカット
                if      (vk == 'C') { SendMessageW(hFocus, WM_COPY,   0, 0); return 1; }
                else if (vk == 'X') { SendMessageW(hFocus, WM_CUT,    0, 0); return 1; }
                else if (vk == 'V') { SendMessageW(hFocus, WM_PASTE,  0, 0); return 1; }
                else if (vk == 'Z') { SendMessageW(hFocus, WM_UNDO,   0, 0); return 1; }
                else if (vk == 'A') { SendMessageW(hFocus, EM_SETSEL, 0, -1); return 1; }
            } else if (g_hwndRight &&
                       (vk == 'C' || vk == 'X' || vk == 'V' || vk == 'A')) {
                // シェルビュー: PostMessage でフック外で実行 (WPARAM = 仮想キーコード)
                PostMessageW(g_hwndRight, WM_EXEC_SHELL_CMD, (WPARAM)vk, 0);
                return 1;
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
    auto& st = bt.smallTabs[bt.smallTabIdx];

    // currentPath があればそちらを優先 (タブごとの位置状態)
    // なければ path (初期位置) へ
    const std::wstring& navPath = st.currentPath.empty() ? st.path : st.currentPath;

    if (navPath.empty() || navPath == L"QuickAccess") {
        PIDLIST_ABSOLUTE pidl = nullptr;
        const GUID FOLDERID_QA = {0x679f85cb, 0x0220, 0x4080,
            {0xb2, 0x9b, 0x55, 0x40, 0xcc, 0x05, 0xaa, 0xb6}};
        if (SUCCEEDED(SHGetKnownFolderIDList(FOLDERID_QA, 0, nullptr, &pidl))) {
            g_peb->BrowseToIDList(pidl, SBSP_ABSOLUTE);
            CoTaskMemFree(pidl);
        }
    } else {
        IShellItem* psi = nullptr;
        if (SUCCEEDED(SHCreateItemFromParsingName(navPath.c_str(), nullptr, IID_PPV_ARGS(&psi)))) {
            g_peb->BrowseToObject(psi, SBSP_ABSOLUTE);
            psi->Release();
        }
    }
    // OnNavigationComplete が届かない場合のフォールバック
    if (g_hwndRight) SetTimer(g_hwndRight, TIMER_ADDR_UPDATE, 300, nullptr);
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
    if (g_hwndRight) SetTimer(g_hwndRight, TIMER_ADDR_UPDATE, 300, nullptr);
}

static void NavForward() {
    if (!g_peb) return;
    IShellBrowser* psb = nullptr;
    if (SUCCEEDED(g_peb->QueryInterface(IID_PPV_ARGS(&psb)))) {
        psb->BrowseObject(nullptr, SBSP_NAVIGATEFORWARD);
        psb->Release();
    }
    if (g_hwndRight) SetTimer(g_hwndRight, TIMER_ADDR_UPDATE, 300, nullptr);
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
    if (g_hwndRight) SetTimer(g_hwndRight, TIMER_ADDR_UPDATE, 300, nullptr);
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

// ComputeTabLayout を使って多段時の実高さを計算する
static int CalcTabHeight(HWND hwndTab, HFONT hFont, int W, int baseH) {
    if (!hwndTab || W <= 0) return baseH;
    auto layout = ComputeTabLayout(hwndTab, hFont, W);
    if (layout.empty()) return baseH;
    int maxBottom = 0;
    for (auto& p : layout) maxBottom = max(maxBottom, p.rc.bottom);
    return maxBottom + 6;
}

void DoRightLayout(HWND hwnd) {
    RECT rc; GetClientRect(hwnd, &rc);
    int W = rc.right, H = rc.bottom;
    int y = 0;

    if (g_hwndBigTab) {
        int h = max(CalcTabHeight(g_hwndBigTab, g_hFontBig, W, g_bigTabH), g_bigTabH);
        SetWindowPos(g_hwndBigTab, nullptr, 0, y, W, h, SWP_NOZORDER | SWP_NOACTIVATE);
        y += h;
    }
    if (g_hwndSmallTab) {
        int h = max(CalcTabHeight(g_hwndSmallTab, g_hFontNormal, W, g_smallTabH), g_smallTabH);
        SetWindowPos(g_hwndSmallTab, nullptr, 0, y, W, h, SWP_NOZORDER | SWP_NOACTIVATE);
        y += h;
    }

    // ナビゲーションバー: [←][→][↑][アドレス入力][コピー][切取][貼付][リンク]
    if (g_hwndBtnBack) {
        constexpr int GAP    = 2;
        constexpr int CLIP_W = 44;
        int x = GAP;
        SetWindowPos(g_hwndBtnBack,  nullptr, x, y + 2, NAV_BTN_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); x += NAV_BTN_W + GAP;
        SetWindowPos(g_hwndBtnFwd,   nullptr, x, y + 2, NAV_BTN_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); x += NAV_BTN_W + GAP;
        SetWindowPos(g_hwndBtnUp,    nullptr, x, y + 2, NAV_BTN_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); x += NAV_BTN_W + GAP;
        // 右端: コピー・切取・貼付・リンク の 4 ボタン
        int clipStart = W - (CLIP_W + GAP) * 4;
        int addrW = max(40, clipStart - x - GAP);
        SetWindowPos(g_hwndAddrEdit,  nullptr, x,         y + 3, addrW,  NAV_BAR_H - 6, SWP_NOZORDER | SWP_NOACTIVATE);
        int cx = clipStart;
        SetWindowPos(s_hwndBtnCopy,   nullptr, cx, y + 2, CLIP_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); cx += CLIP_W + GAP;
        SetWindowPos(s_hwndBtnCut,    nullptr, cx, y + 2, CLIP_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); cx += CLIP_W + GAP;
        SetWindowPos(s_hwndBtnPaste,  nullptr, cx, y + 2, CLIP_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE); cx += CLIP_W + GAP;
        SetWindowPos(s_hwndBtnAlias,  nullptr, cx, y + 2, CLIP_W, NAV_BAR_H - 4, SWP_NOZORDER | SWP_NOACTIVATE);
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
        // TCS_MULTILINE: スクロール矢印を抑制するために必要。
        // 行入れ替えは WM_PAINT を完全上書きしているので視覚的に無関係。
        g_hwndBigTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE | TCS_HOTTRACK | TCS_MULTILINE,
            0, 0, 100, 36, hwnd, reinterpret_cast<HMENU>(101), hi, nullptr);
        if (g_hFontBig) SendMessageW(g_hwndBigTab, WM_SETFONT, (WPARAM)g_hFontBig, TRUE);

        // ── スモールタブ ──
        g_hwndSmallTab = CreateWindowExW(0, WC_TABCONTROLW, nullptr,
            WS_CHILD | WS_VISIBLE | TCS_MULTILINE,
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

        // タブ高さを動的計測 (フォント適用後に実際の行高さを取得)
        // TCS_MULTILINE では選択行が底部に配置されるため tr.bottom はコントロール高さと一致する場合がある。
        // 行の実高さは (tr.bottom - tr.top) で取得する。
        {
            RECT tr{};
            if (g_hwndBigTab && TabCtrl_GetItemRect(g_hwndBigTab, 0, &tr) && tr.bottom > tr.top)
                g_bigTabH = (tr.bottom - tr.top) + 8;
            if (g_hwndSmallTab && TabCtrl_GetItemRect(g_hwndSmallTab, 0, &tr) && tr.bottom > tr.top)
                g_smallTabH = (tr.bottom - tr.top) + 6;
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

        // ── クリップボード / エイリアスボタン ──
        s_hwndBtnCopy  = CreateWindowExW(0, L"BUTTON", L"コピー",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 44, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_BTN_COPY), hi, nullptr);
        s_hwndBtnCut   = CreateWindowExW(0, L"BUTTON", L"切取",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 44, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_BTN_CUT),  hi, nullptr);
        s_hwndBtnPaste = CreateWindowExW(0, L"BUTTON", L"貼付",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 44, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_BTN_PASTE), hi, nullptr);
        // 選択中のオブジェクトを現在のディレクトリに保存するダイアログを開く
        s_hwndBtnAlias = CreateWindowExW(0, L"BUTTON", L"A保存",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 44, NAV_BAR_H, hwnd, reinterpret_cast<HMENU>(ID_BTN_ALIAS), hi, nullptr);

        // ボタン・アドレスバーにダークテーマを適用
        for (HWND h : {g_hwndBtnBack, g_hwndBtnFwd, g_hwndBtnUp,
                       s_hwndBtnCopy, s_hwndBtnCut, s_hwndBtnPaste, s_hwndBtnAlias}) {
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

            // イベント登録 (IExplorerBrowser::Advise を直接使用)
            g_peb->Advise(&g_ebEvents, &g_ebEvents.m_cookie);

            NavigateExplorerToCurrentTab();
        } else {
            Log(L"IExplorerBrowser の作成に失敗しました。");
        }
        return 0;
    }

    case WM_SIZE:
        DoRightLayout(hwnd);
        return 0;

    case WM_TIMER:
        if (wp == TIMER_ADDR_UPDATE) {
            KillTimer(hwnd, TIMER_ADDR_UPDATE);
            UpdateAddrBar();
        }
        return 0;

    case WM_COMMAND: {
        int id = LOWORD(wp);
        if (id == ID_BTN_BACK)  { NavBack();                return 0; }
        if (id == ID_BTN_FWD)   { NavForward();             return 0; }
        if (id == ID_BTN_UP)    { NavUp();                  return 0; }
        if (id == ID_BTN_COPY)  { ExecFileOp(0);            return 0; }
        if (id == ID_BTN_CUT)   { ExecFileOp(1);            return 0; }
        if (id == ID_BTN_PASTE) { ExecFileOp(2);            return 0; }
        if (id == ID_BTN_ALIAS) { CreateAliasForSelected(); return 0; }
        break;
    }

    case WM_NOTIFY: {
        auto* nmhdr = reinterpret_cast<NMHDR*>(lp);

        // ── ボタン NM_CUSTOMDRAW : ダーク描画 ──
        if ((nmhdr->hwndFrom == g_hwndBtnBack  || nmhdr->hwndFrom == g_hwndBtnFwd  ||
             nmhdr->hwndFrom == g_hwndBtnUp    || nmhdr->hwndFrom == s_hwndBtnCopy ||
             nmhdr->hwndFrom == s_hwndBtnCut   || nmhdr->hwndFrom == s_hwndBtnPaste ||
             nmhdr->hwndFrom == s_hwndBtnAlias)
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
            if (sel >= 0) {
                g_bigTabIdx = sel;
                RebuildSmallTabs();
                DoRightLayout(hwnd);  // スモールタブ行数が変わる可能性があるため再計算
                NavigateExplorerToCurrentTab();  // タブの保存済み位置 (currentPath) へ移動
            }
            return 0;
        }
        if (nmhdr->hwndFrom == g_hwndSmallTab && nmhdr->code == TCN_SELCHANGE) {
            int sel = TabCtrl_GetCurSel(g_hwndSmallTab);
            if (sel >= 0 && g_bigTabIdx < (int)g_bigTabs.size()) {
                g_bigTabs[g_bigTabIdx].smallTabIdx = sel;
                NavigateExplorerToCurrentTab();  // タブの保存済み位置 (currentPath) へ移動
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

    case WM_EXEC_SHELL_CMD: {
        UINT vk = static_cast<UINT>(wp);
        if      (vk == 'C') ExecFileOp(0);
        else if (vk == 'X') ExecFileOp(1);
        else if (vk == 'V') ExecFileOp(2);
        return 0;
    }

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
            g_peb->Unadvise(g_ebEvents.m_cookie);
            g_ebEvents.m_cookie = 0;
        }
        if (g_peb) { g_peb->Destroy(); g_peb->Release(); g_peb = nullptr; }
        if (g_hFontBig)    { DeleteObject(g_hFontBig);    g_hFontBig    = nullptr; }
        if (g_hFontNormal) { DeleteObject(g_hFontNormal); g_hFontNormal = nullptr; }
        if (g_hBrushBg)    { DeleteObject(g_hBrushBg);    g_hBrushBg    = nullptr; }
        g_hwndBigTab = g_hwndSmallTab = nullptr;
        g_hwndBtnBack = g_hwndBtnFwd = g_hwndBtnUp = g_hwndAddrEdit = nullptr;
        s_hwndBtnCopy = s_hwndBtnCut = s_hwndBtnPaste = s_hwndBtnAlias = nullptr;
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}
