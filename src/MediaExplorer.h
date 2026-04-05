#pragma once
//----------------------------------------------------------------------------------
//  PreviewExplorer - 共有ヘッダー
//----------------------------------------------------------------------------------
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <string>
#include <vector>
#include <algorithm>

#include <aviutl2_sdk/plugin2.h>
#include <aviutl2_sdk/logger2.h>
#include <aviutl2_sdk/config2.h>

//------------------------------------------------------------------------------
// データ構造
//------------------------------------------------------------------------------
struct SmallTab {
    std::wstring name;
    std::wstring path;         // 初期位置 (右クリック「パス変更」等で設定)
    std::wstring currentPath;  // 最後に表示していた場所 (タブごとの位置状態)
};
struct BigTab   { std::wstring name; std::vector<SmallTab> smallTabs; int smallTabIdx = 0; };

// カスタム多段タブレイアウト (TabsSubclass.cpp で定義)
struct TabPos { RECT rc; };
std::vector<TabPos> ComputeTabLayout(HWND hwnd, HFONT hFont, int W);

//------------------------------------------------------------------------------
// ウィンドウクラス名
//------------------------------------------------------------------------------
constexpr wchar_t WC_MAIN[]  = L"MediaExplorerPlus_Main";
constexpr wchar_t WC_RIGHT[] = L"MediaExplorerPlus_Right";

//------------------------------------------------------------------------------
// レイアウト定数
//------------------------------------------------------------------------------
constexpr int MIN_LEFT_W  = 100;
constexpr int MIN_RIGHT_W = 200;
constexpr int NAV_BAR_H   = 28;   // ナビゲーションバー高さ (固定)
constexpr int NAV_BTN_W   = 28;   // 戻る/進む/上へ ボタン幅
constexpr int STATUS_H    = 20;   // 選択ファイル表示ステータスバー高さ

//------------------------------------------------------------------------------
// グローバル変数宣言 (定義は MediaExplorer.cpp)
//------------------------------------------------------------------------------
extern EDIT_HANDLE*   g_editHandle;
extern LOG_HANDLE*    g_logger;
extern CONFIG_HANDLE* g_config;

extern HWND g_hwndMain;
extern HWND g_hwndLeft;
extern HWND g_hwndRight;
extern HWND g_hwndInfo;
extern HWND g_hwndEmbedded;

// 右ペイン内コントロール
extern HWND g_hwndBigTab;
extern HWND g_hwndSmallTab;
extern HWND g_hwndBtnBack;
extern HWND g_hwndBtnFwd;
extern HWND g_hwndBtnUp;
extern HWND g_hwndAddrEdit;

extern IExplorerBrowser* g_peb;

extern std::vector<BigTab> g_bigTabs;
extern int    g_bigTabIdx;
extern double g_splitRatio;
extern int    g_splitX;
extern bool   g_dragging;

// 表示パラメータ (LoadState で確定)
extern int g_bigTabH;    // 大タブ高さ (動的計測後に上書き)
extern int g_smallTabH;  // 小タブ高さ (動的計測後に上書き)
extern int g_splitterW;

// テーマカラー
extern COLORREF g_clrBg;
extern COLORREF g_clrText;
extern HBRUSH   g_hBrushBg;
extern HFONT    g_hFontBig;
extern HFONT    g_hFontNormal;

//------------------------------------------------------------------------------
// 関数宣言
//------------------------------------------------------------------------------
void         Log(const wchar_t* msg);
void         SaveState();
void         LoadState();
void         DoLayout(HWND hwnd);
void         DoRightLayout(HWND hwnd);
void         RebuildSmallTabs();
void         NavigateExplorerToCurrentTab();
std::wstring GetCurrentExplorerPath();
bool         ShowInputDialog(HWND parent, const wchar_t* title,
                             const wchar_t* prompt, std::wstring& result);

// ウィンドウプロシージャ (RegisterPlugin から参照)
LRESULT CALLBACK RightWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
