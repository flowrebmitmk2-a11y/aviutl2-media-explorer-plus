//----------------------------------------------------------------------------------
//  PreviewExplorer - タブ右クリック + ドラッグ並び替え サブクラスプロシージャ
//----------------------------------------------------------------------------------
#include "MediaExplorer.h"

//------------------------------------------------------------------------------
// ドラッグ状態
//------------------------------------------------------------------------------
static struct {
    int   from    = -1;   // ドラッグ開始タブ
    int   over    = -1;   // 現在マウスが乗っているタブ
    bool  active  = false;
    POINT start   = {};
} s_bigDrag, s_smallDrag;

//------------------------------------------------------------------------------
// タブデータを fromIdx → toIdx へ移動し、選択インデックスを補正
//------------------------------------------------------------------------------
static void MoveBigTab(HWND hwnd, int from, int to) {
    if (from == to) return;
    if (from < 0 || from >= (int)g_bigTabs.size()) return;
    if (to   < 0 || to   >= (int)g_bigTabs.size()) return;

    BigTab moved = g_bigTabs[from];
    g_bigTabs.erase(g_bigTabs.begin() + from);

    int insertAt = (to > from) ? to - 1 : to;

    // g_bigTabIdx を補正
    if (g_bigTabIdx == from) {
        g_bigTabIdx = insertAt;
    } else {
        if (from < g_bigTabIdx) g_bigTabIdx--;
        if (insertAt <= g_bigTabIdx) g_bigTabIdx++;
    }

    g_bigTabs.insert(g_bigTabs.begin() + insertAt, std::move(moved));

    // タブコントロール再構築
    TabCtrl_DeleteAllItems(hwnd);
    for (int i = 0; i < (int)g_bigTabs.size(); i++) {
        TCITEMW tci{}; tci.mask = TCIF_TEXT;
        tci.pszText = const_cast<wchar_t*>(g_bigTabs[i].name.c_str());
        TabCtrl_InsertItem(hwnd, i, &tci);
    }
    TabCtrl_SetCurSel(hwnd, g_bigTabIdx);
    RebuildSmallTabs();
    NavigateExplorerToCurrentTab();
    SaveState();
}

static void MoveSmallTab(HWND hwnd, int from, int to) {
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    auto& bt = g_bigTabs[g_bigTabIdx];
    if (from == to) return;
    if (from < 0 || from >= (int)bt.smallTabs.size()) return;
    if (to   < 0 || to   >= (int)bt.smallTabs.size()) return;

    SmallTab moved = bt.smallTabs[from];
    bt.smallTabs.erase(bt.smallTabs.begin() + from);

    int insertAt = (to > from) ? to - 1 : to;

    if (bt.smallTabIdx == from) {
        bt.smallTabIdx = insertAt;
    } else {
        if (from < bt.smallTabIdx) bt.smallTabIdx--;
        if (insertAt <= bt.smallTabIdx) bt.smallTabIdx++;
    }

    bt.smallTabs.insert(bt.smallTabs.begin() + insertAt, std::move(moved));

    // タブコントロール再構築
    TabCtrl_DeleteAllItems(hwnd);
    for (int i = 0; i < (int)bt.smallTabs.size(); i++) {
        TCITEMW tci{}; tci.mask = TCIF_TEXT;
        tci.pszText = const_cast<wchar_t*>(bt.smallTabs[i].name.c_str());
        TabCtrl_InsertItem(hwnd, i, &tci);
    }
    TabCtrl_SetCurSel(hwnd, bt.smallTabIdx);
    NavigateExplorerToCurrentTab();
    SaveState();
}

//==============================================================================
// ビッグタブ サブクラスプロシージャ
//==============================================================================
// タブを自前描画する共通ヘルパー
static void DrawTabControl(HWND hwnd, HDC hdc, HFONT hFont, int dragOver) {
    RECT rc; GetClientRect(hwnd, &rc);
    FillRect(hdc, &rc, g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_3DFACE+1));

    int count = TabCtrl_GetItemCount(hwnd);
    int sel   = TabCtrl_GetCurSel(hwnd);
    auto clamp = [](int v) { return max(0, min(255, v)); };

    for (int i = 0; i < count; i++) {
        RECT tabRc{};
        if (!TabCtrl_GetItemRect(hwnd, i, &tabRc)) continue;

        bool selected = (i == sel);
        bool over     = (i == dragOver && dragOver >= 0);
        COLORREF bg = selected
            ? RGB(clamp(GetRValue(g_clrBg)+0x28), clamp(GetGValue(g_clrBg)+0x28), clamp(GetBValue(g_clrBg)+0x30))
            : over
            ? RGB(clamp(GetRValue(g_clrBg)+0x18), clamp(GetGValue(g_clrBg)+0x18), clamp(GetBValue(g_clrBg)+0x20))
            : RGB(clamp(GetRValue(g_clrBg)+0x0a), clamp(GetGValue(g_clrBg)+0x0a), clamp(GetBValue(g_clrBg)+0x0a));
        SetDCBrushColor(hdc, bg);
        FillRect(hdc, &tabRc, (HBRUSH)GetStockObject(DC_BRUSH));

        // 選択タブに下線
        if (selected) {
            RECT line = { tabRc.left, tabRc.bottom - 2, tabRc.right, tabRc.bottom };
            COLORREF accent = RGB(clamp(GetRValue(g_clrBg)+0x60),
                                  clamp(GetGValue(g_clrBg)+0x80),
                                  clamp(GetBValue(g_clrBg)+0xa0));
            SetDCBrushColor(hdc, accent);
            FillRect(hdc, &line, (HBRUSH)GetStockObject(DC_BRUSH));
        }

        wchar_t text[128]{};
        TCITEMW tci{}; tci.mask = TCIF_TEXT; tci.pszText = text; tci.cchTextMax = 127;
        TabCtrl_GetItem(hwnd, i, &tci);
        SetTextColor(hdc, g_clrText);
        SetBkMode(hdc, TRANSPARENT);
        HFONT hold = hFont ? (HFONT)SelectObject(hdc, hFont) : nullptr;
        DrawTextW(hdc, text, -1, &tabRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
        if (hold) SelectObject(hdc, hold);
    }
}

LRESULT CALLBACK BigTabSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                    UINT_PTR uid, DWORD_PTR data) {
    if (msg == WM_ERASEBKGND) return 1;  // WM_PAINT で全て描画するため不要

    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        DrawTabControl(hwnd, hdc, g_hFontBig, s_bigDrag.active ? s_bigDrag.over : -1);
        EndPaint(hwnd, &ps);
        return 0;
    }

    switch (msg) {
    // ── ドラッグ開始判定 ──
    case WM_LBUTTONDOWN: {
        TCHITTESTINFO hti{}; hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        int idx = TabCtrl_HitTest(hwnd, &hti);
        if (idx >= 0) {
            s_bigDrag.from   = idx;
            s_bigDrag.over   = idx;
            s_bigDrag.active = false;
            s_bigDrag.start  = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            SetCapture(hwnd);
        }
        break;  // 通常の選択処理も通す
    }

    case WM_MOUSEMOVE: {
        if (s_bigDrag.from < 0 || !(wp & MK_LBUTTON)) break;
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        if (!s_bigDrag.active) {
            // 閾値 (4px) を超えたらドラッグ開始
            if (abs(pt.x - s_bigDrag.start.x) > 4 || abs(pt.y - s_bigDrag.start.y) > 4)
                s_bigDrag.active = true;
        }
        if (s_bigDrag.active) {
            TCHITTESTINFO hti{}; hti.pt = pt;
            int over = TabCtrl_HitTest(hwnd, &hti);
            if (over >= 0 && over != s_bigDrag.over) {
                s_bigDrag.over = over;
                InvalidateRect(hwnd, nullptr, FALSE);  // ハイライト更新
            }
            SetCursor(LoadCursorW(nullptr, IDC_SIZEALL));
            return 0;
        }
        break;
    }

    case WM_SETCURSOR:
        if (s_bigDrag.active) {
            SetCursor(LoadCursorW(nullptr, IDC_SIZEALL));
            return TRUE;
        }
        break;

    case WM_LBUTTONUP: {
        bool wasDragging = s_bigDrag.active;
        int  from        = s_bigDrag.from;
        s_bigDrag.from   = -1;
        s_bigDrag.active = false;
        s_bigDrag.over   = -1;
        if (GetCapture() == hwnd) ReleaseCapture();

        if (wasDragging && from >= 0) {
            TCHITTESTINFO hti{}; hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            int to = TabCtrl_HitTest(hwnd, &hti);
            if (to >= 0 && to != from)
                MoveBigTab(hwnd, from, to);
            return 0;  // デフォルト処理スキップ（選択変更を防ぐ）
        }
        break;
    }

    case WM_CAPTURECHANGED:
        s_bigDrag.from   = -1;
        s_bigDrag.active = false;
        s_bigDrag.over   = -1;
        break;

    // ── 右クリックメニュー ──
    case WM_RBUTTONDOWN: {
        TCHITTESTINFO hti{}; hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        int hitIdx = TabCtrl_HitTest(hwnd, &hti);

        HMENU hm = CreatePopupMenu();
        AppendMenuW(hm, MF_STRING, 3001, L"追加");
        if (hitIdx >= 0) {
            AppendMenuW(hm, MF_STRING, 3002, L"名前変更");
            AppendMenuW(hm, MF_STRING, 3003, L"削除");
        }
        POINT pt; GetCursorPos(&pt);
        int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(hm);

        if (cmd == 3001) {
            BigTab bt; bt.name = L"新規";
            bt.smallTabs.push_back({ L"クイックアクセス", L"" });
            g_bigTabs.push_back(bt);
            g_bigTabIdx = (int)g_bigTabs.size() - 1;
            TCITEMW tci{}; tci.mask = TCIF_TEXT;
            tci.pszText = const_cast<wchar_t*>(g_bigTabs[g_bigTabIdx].name.c_str());
            TabCtrl_InsertItem(hwnd, g_bigTabIdx, &tci);
            TabCtrl_SetCurSel(hwnd, g_bigTabIdx);
            RebuildSmallTabs();
            NavigateExplorerToCurrentTab();
            SaveState();

        } else if (cmd == 3002 && hitIdx >= 0) {
            std::wstring name = g_bigTabs[hitIdx].name;
            if (ShowInputDialog(GetAncestor(hwnd, GA_ROOT), L"名前変更", L"新しい名前:", name)) {
                g_bigTabs[hitIdx].name = name;
                TCITEMW tci{}; tci.mask = TCIF_TEXT;
                tci.pszText = const_cast<wchar_t*>(name.c_str());
                TabCtrl_SetItem(hwnd, hitIdx, &tci);
                SaveState();
            }

        } else if (cmd == 3003 && hitIdx >= 0) {
            if (g_bigTabs.size() <= 1) {
                MessageBoxW(hwnd, L"最後のタブは削除できません。", L"削除", MB_OK | MB_ICONWARNING);
                return 0;
            }
            g_bigTabs.erase(g_bigTabs.begin() + hitIdx);
            TabCtrl_DeleteItem(hwnd, hitIdx);
            if (g_bigTabIdx >= (int)g_bigTabs.size()) g_bigTabIdx = (int)g_bigTabs.size() - 1;
            TabCtrl_SetCurSel(hwnd, g_bigTabIdx);
            RebuildSmallTabs();
            NavigateExplorerToCurrentTab();
            SaveState();
        }
        return 0;
    }
    } // switch

    return DefSubclassProc(hwnd, msg, wp, lp);
}

//==============================================================================
// スモールタブ サブクラスプロシージャ
//==============================================================================
LRESULT CALLBACK SmallTabSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                      UINT_PTR uid, DWORD_PTR data) {
    if (msg == WM_ERASEBKGND) return 1;

    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        DrawTabControl(hwnd, hdc, g_hFontNormal, s_smallDrag.active ? s_smallDrag.over : -1);
        EndPaint(hwnd, &ps);
        return 0;
    }

    switch (msg) {
    case WM_LBUTTONDOWN: {
        TCHITTESTINFO hti{}; hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        int idx = TabCtrl_HitTest(hwnd, &hti);
        if (idx >= 0) {
            s_smallDrag.from   = idx;
            s_smallDrag.over   = idx;
            s_smallDrag.active = false;
            s_smallDrag.start  = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            SetCapture(hwnd);
        }
        break;
    }

    case WM_MOUSEMOVE: {
        if (s_smallDrag.from < 0 || !(wp & MK_LBUTTON)) break;
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        if (!s_smallDrag.active) {
            if (abs(pt.x - s_smallDrag.start.x) > 4 || abs(pt.y - s_smallDrag.start.y) > 4)
                s_smallDrag.active = true;
        }
        if (s_smallDrag.active) {
            TCHITTESTINFO hti{}; hti.pt = pt;
            int over = TabCtrl_HitTest(hwnd, &hti);
            if (over >= 0 && over != s_smallDrag.over) {
                s_smallDrag.over = over;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            SetCursor(LoadCursorW(nullptr, IDC_SIZEALL));
            return 0;
        }
        break;
    }

    case WM_SETCURSOR:
        if (s_smallDrag.active) {
            SetCursor(LoadCursorW(nullptr, IDC_SIZEALL));
            return TRUE;
        }
        break;

    case WM_LBUTTONUP: {
        bool wasDragging = s_smallDrag.active;
        int  from        = s_smallDrag.from;
        s_smallDrag.from   = -1;
        s_smallDrag.active = false;
        s_smallDrag.over   = -1;
        if (GetCapture() == hwnd) ReleaseCapture();

        if (wasDragging && from >= 0) {
            TCHITTESTINFO hti{}; hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            int to = TabCtrl_HitTest(hwnd, &hti);
            if (to >= 0 && to != from)
                MoveSmallTab(hwnd, from, to);
            return 0;
        }
        break;
    }

    case WM_CAPTURECHANGED:
        s_smallDrag.from   = -1;
        s_smallDrag.active = false;
        s_smallDrag.over   = -1;
        break;

    case WM_RBUTTONDOWN: {
        TCHITTESTINFO hti{}; hti.pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
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
        int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(hm);

        if (cmd == 4001) {
            SmallTab st; st.name = L"新規"; st.path = L"";
            bt.smallTabs.push_back(st);
            int newIdx = (int)bt.smallTabs.size() - 1;
            TCITEMW tci{}; tci.mask = TCIF_TEXT;
            tci.pszText = const_cast<wchar_t*>(st.name.c_str());
            TabCtrl_InsertItem(hwnd, newIdx, &tci);
            bt.smallTabIdx = newIdx;
            TabCtrl_SetCurSel(hwnd, newIdx);
            NavigateExplorerToCurrentTab();
            SaveState();

        } else if (cmd == 4002 && hitIdx >= 0) {
            std::wstring name = bt.smallTabs[hitIdx].name;
            if (ShowInputDialog(GetAncestor(hwnd, GA_ROOT), L"名前変更", L"新しい名前:", name)) {
                bt.smallTabs[hitIdx].name = name;
                TCITEMW tci{}; tci.mask = TCIF_TEXT;
                tci.pszText = const_cast<wchar_t*>(name.c_str());
                TabCtrl_SetItem(hwnd, hitIdx, &tci);
                SaveState();
            }

        } else if (cmd == 4003 && hitIdx >= 0) {
            wchar_t disp[MAX_PATH]{};
            BROWSEINFOW bi{};
            bi.hwndOwner = GetAncestor(hwnd, GA_ROOT);
            bi.pszDisplayName = disp;
            bi.lpszTitle = L"フォルダを選択";
            bi.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
            if (pidl) {
                wchar_t path[MAX_PATH]{};
                SHGetPathFromIDListW(pidl, path);
                CoTaskMemFree(pidl);
                bt.smallTabs[hitIdx].path = path;
                if (hitIdx == bt.smallTabIdx) NavigateExplorerToCurrentTab();
                SaveState();
            }

        } else if (cmd == 4004 && hitIdx >= 0) {
            std::wstring path = GetCurrentExplorerPath();
            if (!path.empty()) { bt.smallTabs[hitIdx].path = path; SaveState(); }

        } else if (cmd == 4005 && hitIdx >= 0) {
            if (bt.smallTabs.size() <= 1) {
                MessageBoxW(hwnd, L"最後のタブは削除できません。", L"削除", MB_OK | MB_ICONWARNING);
                return 0;
            }
            bt.smallTabs.erase(bt.smallTabs.begin() + hitIdx);
            TabCtrl_DeleteItem(hwnd, hitIdx);
            if (bt.smallTabIdx >= (int)bt.smallTabs.size()) bt.smallTabIdx = (int)bt.smallTabs.size() - 1;
            TabCtrl_SetCurSel(hwnd, bt.smallTabIdx);
            NavigateExplorerToCurrentTab();
            SaveState();
        }
        return 0;
    }
    } // switch

    return DefSubclassProc(hwnd, msg, wp, lp);
}
