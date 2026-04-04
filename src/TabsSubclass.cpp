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
// タブコピーペースト用クリップボード [改善2]
//------------------------------------------------------------------------------
static BigTab   s_copiedBigTab;    // コピーした大タブ
static SmallTab s_copiedSmallTab;  // コピーした小タブ
static bool     s_hasCopiedBigTab   = false;  // 大タブがコピーされているか
static bool     s_hasCopiedSmallTab = false;  // 小タブがコピーされているか

static bool HasBigTabName(const std::wstring& name) {
    for (const auto& tab : g_bigTabs) {
        if (tab.name == name) return true;
    }
    return false;
}

static bool HasSmallTabName(const BigTab& bigTab, const std::wstring& name) {
    for (const auto& tab : bigTab.smallTabs) {
        if (tab.name == name) return true;
    }
    return false;
}

static std::wstring MakeUniqueCopiedBigTabName(const std::wstring& baseName) {
    if (!HasBigTabName(baseName)) return baseName;

    const std::wstring copiedName = baseName + L"のコピー";
    if (!HasBigTabName(copiedName)) return copiedName;

    for (int suffix = 2;; ++suffix) {
        std::wstring candidate = copiedName + L" " + std::to_wstring(suffix);
        if (!HasBigTabName(candidate)) return candidate;
    }
}

static std::wstring MakeUniqueCopiedSmallTabName(const BigTab& bigTab, const std::wstring& baseName) {
    if (!HasSmallTabName(bigTab, baseName)) return baseName;

    const std::wstring copiedName = baseName + L"のコピー";
    if (!HasSmallTabName(bigTab, copiedName)) return copiedName;

    for (int suffix = 2;; ++suffix) {
        std::wstring candidate = copiedName + L" " + std::to_wstring(suffix);
        if (!HasSmallTabName(bigTab, candidate)) return candidate;
    }
}

//------------------------------------------------------------------------------
// カスタム多段レイアウト計算
//  - ネイティブコントロールを単行のまま使い、描画/ヒットテスト/高さは自前で管理
//  - TCS_MULTILINE の「選択行が下に移動」問題を回避する
//------------------------------------------------------------------------------
std::vector<TabPos> ComputeTabLayout(HWND hwnd, HFONT hFont, int W) {
    int count = TabCtrl_GetItemCount(hwnd);
    std::vector<TabPos> result(count);
    if (count == 0 || W <= 0) return result;

    // 1行の高さ: ネイティブのアイテム 0 の rect から取得
    RECT baseRc{};
    int rowH = 22;
    if (TabCtrl_GetItemRect(hwnd, 0, &baseRc) && baseRc.bottom > baseRc.top)
        rowH = baseRc.bottom - baseRc.top;

    HDC hdc = GetDC(hwnd);
    HFONT oldFont = hFont ? (HFONT)SelectObject(hdc, hFont) : nullptr;

    int x = 0, y = 0;
    for (int i = 0; i < count; i++) {
        wchar_t text[128]{};
        TCITEMW tci{}; tci.mask = TCIF_TEXT; tci.pszText = text; tci.cchTextMax = 127;
        TabCtrl_GetItem(hwnd, i, &tci);

        SIZE sz{};
        GetTextExtentPoint32W(hdc, text, (int)wcslen(text), &sz);
        int w = max(sz.cx + 18, 40);

        if (x > 0 && x + w > W) { x = 0; y += rowH; }
        result[i].rc = { x, y, x + w, y + rowH };
        x += w;
    }

    if (oldFont) SelectObject(hdc, oldFont);
    ReleaseDC(hwnd, hdc);
    return result;
}

// layout から座標でタブを検索
static int HitTestLayout(const std::vector<TabPos>& layout, POINT pt) {
    for (int i = 0; i < (int)layout.size(); i++)
        if (PtInRect(&layout[i].rc, pt)) return i;
    return -1;
}

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

    if (g_bigTabIdx == from) {
        g_bigTabIdx = insertAt;
    } else {
        if (from < g_bigTabIdx) g_bigTabIdx--;
        if (insertAt <= g_bigTabIdx) g_bigTabIdx++;
    }

    g_bigTabs.insert(g_bigTabs.begin() + insertAt, std::move(moved));

    TabCtrl_DeleteAllItems(hwnd);
    for (int i = 0; i < (int)g_bigTabs.size(); i++) {
        TCITEMW tci{}; tci.mask = TCIF_TEXT;
        tci.pszText = const_cast<wchar_t*>(g_bigTabs[i].name.c_str());
        TabCtrl_InsertItem(hwnd, i, &tci);
    }
    TabCtrl_SetCurSel(hwnd, g_bigTabIdx);
    RebuildSmallTabs();
    if (g_hwndRight) DoRightLayout(g_hwndRight);
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

//------------------------------------------------------------------------------
// タブコピーペースト用ヘルパー関数 [改善2]
//------------------------------------------------------------------------------
static void CopyBigTab(int idx) {
    if (idx < 0 || idx >= (int)g_bigTabs.size()) return;
    s_copiedBigTab = g_bigTabs[idx];
    s_hasCopiedBigTab = true;
    s_hasCopiedSmallTab = false;  // 大タブコピー時は小タブ情報をクリア
}

static void PasteBigTab(HWND hwnd) {
    if (!s_hasCopiedBigTab) return;
    BigTab newTab = s_copiedBigTab;
    newTab.name = MakeUniqueCopiedBigTabName(newTab.name);
    g_bigTabs.push_back(newTab);
    g_bigTabIdx = (int)g_bigTabs.size() - 1;
    
    TCITEMW tci{}; tci.mask = TCIF_TEXT;
    tci.pszText = const_cast<wchar_t*>(newTab.name.c_str());
    TabCtrl_InsertItem(hwnd, g_bigTabIdx, &tci);
    TabCtrl_SetCurSel(hwnd, g_bigTabIdx);
    RebuildSmallTabs();
    if (g_hwndRight) DoRightLayout(g_hwndRight);
    NavigateExplorerToCurrentTab();
    SaveState();
}

static void CopySmallTab(int idx) {
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    auto& bt = g_bigTabs[g_bigTabIdx];
    if (idx < 0 || idx >= (int)bt.smallTabs.size()) return;
    s_copiedSmallTab = bt.smallTabs[idx];
    s_hasCopiedSmallTab = true;
    s_hasCopiedBigTab = false;  // 小タブコピー時は大タブ情報をクリア
}

static void PasteSmallTab(HWND hwnd) {
    if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return;
    if (!s_hasCopiedSmallTab) return;
    
    auto& bt = g_bigTabs[g_bigTabIdx];
    SmallTab newTab = s_copiedSmallTab;
    newTab.name = MakeUniqueCopiedSmallTabName(bt, newTab.name);
    bt.smallTabs.push_back(newTab);
    int newIdx = (int)bt.smallTabs.size() - 1;
    
    TCITEMW tci{}; tci.mask = TCIF_TEXT;
    tci.pszText = const_cast<wchar_t*>(newTab.name.c_str());
    TabCtrl_InsertItem(hwnd, newIdx, &tci);
    bt.smallTabIdx = newIdx;
    TabCtrl_SetCurSel(hwnd, newIdx);
    if (g_hwndRight) DoRightLayout(g_hwndRight);
    NavigateExplorerToCurrentTab();
    SaveState();
}

//==============================================================================
// 共通タブ描画ヘルパー (ComputeTabLayout を使用)
//==============================================================================
static void DrawTabControl(HWND hwnd, HDC hdc, HFONT hFont, int dragOver) {
    RECT rc; GetClientRect(hwnd, &rc);
    FillRect(hdc, &rc, g_hBrushBg ? g_hBrushBg : (HBRUSH)(COLOR_3DFACE+1));

    int count = TabCtrl_GetItemCount(hwnd);
    int sel   = TabCtrl_GetCurSel(hwnd);
    auto clamp = [](int v) { return max(0, min(255, v)); };

    auto layout = ComputeTabLayout(hwnd, hFont, rc.right);

    for (int i = 0; i < count && i < (int)layout.size(); i++) {
        RECT tabRc = layout[i].rc;

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

//==============================================================================
// ビッグタブ サブクラスプロシージャ
//==============================================================================
LRESULT CALLBACK BigTabSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp,
                                    UINT_PTR uid, DWORD_PTR data) {
    if (msg == WM_ERASEBKGND) return 1;

    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        DrawTabControl(hwnd, hdc, g_hFontBig, s_bigDrag.active ? s_bigDrag.over : -1);
        EndPaint(hwnd, &ps);
        return 0;
    }

    switch (msg) {
    case WM_LBUTTONDOWN: {
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        RECT crc{}; GetClientRect(hwnd, &crc);
        auto layout = ComputeTabLayout(hwnd, g_hFontBig, crc.right);
        int idx = HitTestLayout(layout, pt);
        if (idx >= 0) {
            int prev = TabCtrl_GetCurSel(hwnd);
            TabCtrl_SetCurSel(hwnd, idx);
            if (idx != prev) {
                // TCN_SELCHANGE を親へ通知
                NMHDR nm{}; nm.hwndFrom = hwnd; nm.idFrom = GetDlgCtrlID(hwnd); nm.code = TCN_SELCHANGE;
                SendMessageW(GetParent(hwnd), WM_NOTIFY, nm.idFrom, (LPARAM)&nm);
            }
            InvalidateRect(hwnd, nullptr, FALSE);
            // ドラッグ準備
            s_bigDrag.from   = idx;
            s_bigDrag.over   = idx;
            s_bigDrag.active = false;
            s_bigDrag.start  = pt;
            SetCapture(hwnd);
        }
        return 0;
    }

    case WM_MOUSEMOVE: {
        if (s_bigDrag.from < 0 || !(wp & MK_LBUTTON)) break;
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        if (!s_bigDrag.active) {
            if (abs(pt.x - s_bigDrag.start.x) > 4 || abs(pt.y - s_bigDrag.start.y) > 4)
                s_bigDrag.active = true;
        }
        if (s_bigDrag.active) {
            RECT crc{}; GetClientRect(hwnd, &crc);
            auto layout = ComputeTabLayout(hwnd, g_hFontBig, crc.right);
            int over = HitTestLayout(layout, pt);
            if (over >= 0 && over != s_bigDrag.over) {
                s_bigDrag.over = over;
                InvalidateRect(hwnd, nullptr, FALSE);
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
            POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            RECT crc{}; GetClientRect(hwnd, &crc);
            auto layout = ComputeTabLayout(hwnd, g_hFontBig, crc.right);
            int to = HitTestLayout(layout, pt);
            if (to >= 0 && to != from)
                MoveBigTab(hwnd, from, to);
            return 0;
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
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        RECT crc{}; GetClientRect(hwnd, &crc);
        auto layout = ComputeTabLayout(hwnd, g_hFontBig, crc.right);
        int hitIdx = HitTestLayout(layout, pt);

        HMENU hm = CreatePopupMenu();
        AppendMenuW(hm, MF_STRING, 3001, L"追加");
        if (hitIdx >= 0) {
            AppendMenuW(hm, MF_STRING, 3002, L"名前変更");
            AppendMenuW(hm, MF_STRING, 3004, L"初期位置に移動");
            AppendMenuW(hm, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(hm, MF_STRING, 3005, L"コピー");          // [改善2]
            // ペースト（コピーがあれば有効、なければ無効）
            UINT pasteFlags = s_hasCopiedBigTab ? MF_STRING : (MF_STRING | MF_GRAYED);
            AppendMenuW(hm, pasteFlags, 3006, L"ペースト");
            AppendMenuW(hm, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(hm, MF_STRING, 3003, L"削除");
        }
        POINT spt; GetCursorPos(&spt);
        int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON, spt.x, spt.y, 0, hwnd, nullptr);
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
            if (g_hwndRight) DoRightLayout(g_hwndRight);
            NavigateExplorerToCurrentTab();
            SaveState();

        } else if (cmd == 3002 && hitIdx >= 0) {
            std::wstring name = g_bigTabs[hitIdx].name;
            if (ShowInputDialog(GetAncestor(hwnd, GA_ROOT), L"名前変更", L"新しい名前:", name)) {
                g_bigTabs[hitIdx].name = name;
                TCITEMW tci{}; tci.mask = TCIF_TEXT;
                tci.pszText = const_cast<wchar_t*>(name.c_str());
                TabCtrl_SetItem(hwnd, hitIdx, &tci);
                InvalidateRect(hwnd, nullptr, FALSE);
                SaveState();
            }

        } else if (cmd == 3005 && hitIdx >= 0) {
            CopyBigTab(hitIdx);  // [改善2]

        } else if (cmd == 3006) {
            PasteBigTab(hwnd);   // [改善2]

        } else if (cmd == 3004 && hitIdx >= 0) {
            // 選択タブに切り替えてからそのタブの初期位置に移動
            if (g_bigTabIdx != hitIdx) {
                g_bigTabIdx = hitIdx;
                TabCtrl_SetCurSel(hwnd, hitIdx);
                RebuildSmallTabs();
                if (g_hwndRight) DoRightLayout(g_hwndRight);
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            NavigateExplorerToCurrentTab();

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
            if (g_hwndRight) DoRightLayout(g_hwndRight);
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
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        RECT crc{}; GetClientRect(hwnd, &crc);
        auto layout = ComputeTabLayout(hwnd, g_hFontNormal, crc.right);
        int idx = HitTestLayout(layout, pt);
        if (idx >= 0) {
            int prev = TabCtrl_GetCurSel(hwnd);
            TabCtrl_SetCurSel(hwnd, idx);
            if (idx != prev) {
                NMHDR nm{}; nm.hwndFrom = hwnd; nm.idFrom = GetDlgCtrlID(hwnd); nm.code = TCN_SELCHANGE;
                SendMessageW(GetParent(hwnd), WM_NOTIFY, nm.idFrom, (LPARAM)&nm);
            }
            InvalidateRect(hwnd, nullptr, FALSE);
            s_smallDrag.from   = idx;
            s_smallDrag.over   = idx;
            s_smallDrag.active = false;
            s_smallDrag.start  = pt;
            SetCapture(hwnd);
        }
        return 0;
    }

    case WM_MOUSEMOVE: {
        if (s_smallDrag.from < 0 || !(wp & MK_LBUTTON)) break;
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        if (!s_smallDrag.active) {
            if (abs(pt.x - s_smallDrag.start.x) > 4 || abs(pt.y - s_smallDrag.start.y) > 4)
                s_smallDrag.active = true;
        }
        if (s_smallDrag.active) {
            RECT crc{}; GetClientRect(hwnd, &crc);
            auto layout = ComputeTabLayout(hwnd, g_hFontNormal, crc.right);
            int over = HitTestLayout(layout, pt);
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
            POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            RECT crc{}; GetClientRect(hwnd, &crc);
            auto layout = ComputeTabLayout(hwnd, g_hFontNormal, crc.right);
            int to = HitTestLayout(layout, pt);
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
        POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
        RECT crc{}; GetClientRect(hwnd, &crc);
        auto layout = ComputeTabLayout(hwnd, g_hFontNormal, crc.right);
        int hitIdx = HitTestLayout(layout, pt);

        if (g_bigTabIdx < 0 || g_bigTabIdx >= (int)g_bigTabs.size()) return 0;
        auto& bt = g_bigTabs[g_bigTabIdx];

        HMENU hm = CreatePopupMenu();
        AppendMenuW(hm, MF_STRING, 4001, L"追加");
        if (hitIdx >= 0) {
            AppendMenuW(hm, MF_STRING, 4002, L"名前変更");
            AppendMenuW(hm, MF_STRING, 4003, L"パス変更...");
            AppendMenuW(hm, MF_STRING, 4004, L"現在の場所をパスに設定");
            AppendMenuW(hm, MF_STRING, 4006, L"初期位置に移動");
            AppendMenuW(hm, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(hm, MF_STRING, 4007, L"コピー");          // [改善2]
            // ペースト（コピーがあれば有効、なければ無効）
            UINT pasteFlags2 = s_hasCopiedSmallTab ? MF_STRING : (MF_STRING | MF_GRAYED);
            AppendMenuW(hm, pasteFlags2, 4008, L"ペースト");
            AppendMenuW(hm, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(hm, MF_STRING, 4005, L"削除");
        }
        POINT spt; GetCursorPos(&spt);
        int cmd = TrackPopupMenu(hm, TPM_RETURNCMD | TPM_RIGHTBUTTON, spt.x, spt.y, 0, hwnd, nullptr);
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
            if (g_hwndRight) DoRightLayout(g_hwndRight);
            NavigateExplorerToCurrentTab();
            SaveState();

        } else if (cmd == 4002 && hitIdx >= 0) {
            std::wstring name = bt.smallTabs[hitIdx].name;
            if (ShowInputDialog(GetAncestor(hwnd, GA_ROOT), L"名前変更", L"新しい名前:", name)) {
                bt.smallTabs[hitIdx].name = name;
                TCITEMW tci{}; tci.mask = TCIF_TEXT;
                tci.pszText = const_cast<wchar_t*>(name.c_str());
                TabCtrl_SetItem(hwnd, hitIdx, &tci);
                InvalidateRect(hwnd, nullptr, FALSE);
                SaveState();
            }

        } else if (cmd == 4007 && hitIdx >= 0) {
            CopySmallTab(hitIdx);  // [改善2]

        } else if (cmd == 4008) {
            PasteSmallTab(hwnd);   // [改善2]

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

        } else if (cmd == 4006 && hitIdx >= 0) {
            // 選択タブの初期位置に移動 (タブ切替時は自動ナビゲートしないため手動用)
            if (bt.smallTabIdx != hitIdx) {
                bt.smallTabIdx = hitIdx;
                TabCtrl_SetCurSel(hwnd, hitIdx);
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            NavigateExplorerToCurrentTab();

        } else if (cmd == 4005 && hitIdx >= 0) {
            if (bt.smallTabs.size() <= 1) {
                MessageBoxW(hwnd, L"最後のタブは削除できません。", L"削除", MB_OK | MB_ICONWARNING);
                return 0;
            }
            bt.smallTabs.erase(bt.smallTabs.begin() + hitIdx);
            TabCtrl_DeleteItem(hwnd, hitIdx);
            if (bt.smallTabIdx >= (int)bt.smallTabs.size()) bt.smallTabIdx = (int)bt.smallTabs.size() - 1;
            TabCtrl_SetCurSel(hwnd, bt.smallTabIdx);
            if (g_hwndRight) DoRightLayout(g_hwndRight);
            NavigateExplorerToCurrentTab();
            SaveState();
        }
        return 0;
    }
    } // switch

    return DefSubclassProc(hwnd, msg, wp, lp);
}
