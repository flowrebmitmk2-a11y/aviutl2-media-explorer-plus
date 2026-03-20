//----------------------------------------------------------------------------------
//  PreviewExplorer - 入力ダイアログ (タブ名前変更用)
//  DialogBoxIndirectParamW を使用した標準 Win32 モーダルダイアログ実装
//----------------------------------------------------------------------------------
#include "MediaExplorer.h"
#include <vector>

//------------------------------------------------------------------------------
// メモリ内ダイアログテンプレートビルダー
//------------------------------------------------------------------------------
struct DlgBuf {
    std::vector<BYTE> data;
    void wrd(WORD w)  { data.push_back(BYTE(w)); data.push_back(BYTE(w >> 8)); }
    void dwd(DWORD d) { wrd(WORD(d)); wrd(WORD(d >> 16)); }
    void str(const wchar_t* s) { for (; *s; ++s) wrd(WORD(*s)); wrd(0); }
    void align4() { while (data.size() % 4) data.push_back(0); }

    void item(DWORD style, DWORD exStyle,
              short x, short y, short cx, short cy,
              WORD id, WORD cls, const wchar_t* text) {
        align4();
        dwd(style); dwd(exStyle);
        wrd(x); wrd(y); wrd(cx); wrd(cy);
        wrd(id);
        wrd(0xFFFF); wrd(cls);  // 定義済みクラス atom
        str(text);
        wrd(0);  // 追加データなし
    }
};

// ダイアログテンプレートをメモリに構築
static std::vector<BYTE> MakeInputDlgTemplate(const wchar_t* title, const wchar_t* prompt) {
    DlgBuf b;

    // DLGTEMPLATE ヘッダー
    b.dwd(DS_SETFONT | DS_MODALFRAME | DS_CENTER | DS_3DLOOK |
          WS_POPUP | WS_CAPTION | WS_SYSMENU);
    b.dwd(WS_EX_TOPMOST);
    b.wrd(4);               // コントロール数
    b.wrd(0); b.wrd(0);     // x, y (DS_CENTER で自動センタリング)
    b.wrd(158); b.wrd(62);  // cx, cy (ダイアログ単位)
    b.wrd(0);               // メニュー: なし
    b.wrd(0);               // クラス: デフォルト
    b.str(title);
    // DS_SETFONT: ポイントサイズ + フォント名
    b.wrd(9);
    b.str(L"MS Shell Dlg");

    // 0: STATIC (プロンプト)
    b.item(WS_CHILD | WS_VISIBLE | SS_LEFT,
           0, 7, 7, 144, 10, 0xFFFF, 0x0082, prompt);
    // 1: EDIT (入力欄)
    b.item(WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL,
           WS_EX_CLIENTEDGE, 7, 21, 144, 12, 1001, 0x0081, L"");
    // 2: OK ボタン
    b.item(WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON | WS_TABSTOP,
           0, 33, 41, 37, 14, IDOK, 0x0080, L"OK");
    // 3: キャンセルボタン
    b.item(WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
           0, 74, 41, 52, 14, IDCANCEL, 0x0080, L"キャンセル");

    return b.data;
}

//------------------------------------------------------------------------------
// ダイアログプロシージャ
//------------------------------------------------------------------------------
struct InputParam {
    const wchar_t* initial = nullptr;
    std::wstring   result;
    bool           ok = false;
};

static INT_PTR CALLBACK InputDlgProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_INITDIALOG: {
        auto* p = reinterpret_cast<InputParam*>(lp);
        SetWindowLongPtrW(hwnd, DWLP_USER, lp);
        if (p->initial)
            SetDlgItemTextW(hwnd, 1001, p->initial);
        SendDlgItemMessageW(hwnd, 1001, EM_SETSEL, 0, -1);
        return TRUE;  // フォーカスをデフォルト(最初の TABSTOP)に設定
    }
    case WM_COMMAND:
        if (LOWORD(wp) == IDOK) {
            auto* p = reinterpret_cast<InputParam*>(GetWindowLongPtrW(hwnd, DWLP_USER));
            if (p) {
                wchar_t buf[512]{};
                GetDlgItemTextW(hwnd, 1001, buf, 512);
                p->result = buf; p->ok = true;
            }
            EndDialog(hwnd, IDOK);
            return TRUE;
        }
        if (LOWORD(wp) == IDCANCEL) {
            EndDialog(hwnd, IDCANCEL);
            return TRUE;
        }
        break;
    }
    return FALSE;
}

//------------------------------------------------------------------------------
// 公開 API
//------------------------------------------------------------------------------
bool ShowInputDialog(HWND parent, const wchar_t* title,
                     const wchar_t* prompt, std::wstring& result) {
    InputParam p;
    p.initial = result.c_str();

    auto tmpl = MakeInputDlgTemplate(title, prompt);

    INT_PTR ret = DialogBoxIndirectParamW(
        GetModuleHandleW(nullptr),
        reinterpret_cast<DLGTEMPLATE*>(tmpl.data()),
        parent,
        InputDlgProc,
        reinterpret_cast<LPARAM>(&p));

    if (ret == IDOK) { result = p.result; return true; }
    return false;
}
