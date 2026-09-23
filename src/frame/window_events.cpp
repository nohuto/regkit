// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "win32/text_transform.h"

namespace regkit
{
using namespace window_detail;

MainWindow::Impl::~Impl() = default;

bool MainWindow::Impl::Create(HINSTANCE instance)
{
    instance_ = instance;
    last_search_.criteria.search_keys = false;

    WNDCLASSEXW wc = {};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = MainWindow::Impl::WndProc;
    wc.hInstance = instance;
    wc.lpszClassName = kMainWindowClassName;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
    wc.hIconSm =
        static_cast<HICON>(LoadImageW(instance, MAKEINTRESOURCEW(IDI_APPICON), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
    wc.hbrBackground = nullptr;

    RegisterClassExW(&wc);

    std::wstring title = L"RegKit V";
    title.append(REGKIT_VERSION_STR_W);
    if (util::IsProcessTrustedInstaller())
    {
        title.append(L" - [TrustedInstaller]");
    }
    else if (util::IsProcessSystem())
    {
        title.append(L" - [SYSTEM]");
    }
    else if (util::IsProcessElevated())
    {
        title.append(L" - [Administrator]");
    }
    hwnd_ = CreateWindowExW(0, wc.lpszClassName, title.c_str(), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, CW_USEDEFAULT, CW_USEDEFAULT, 1200, 800, nullptr, nullptr, instance, this);
    if (hwnd_)
    {
        SetPropW(hwnd_, kRegKitWindowProperty, reinterpret_cast<HANDLE>(static_cast<INT_PTR>(1)));
        DragAcceptFiles(hwnd_, TRUE);
    }
    return hwnd_ != nullptr;
}

void MainWindow::Impl::Show(int cmd_show)
{
    int show_cmd = cmd_show;
    if (window_placement_loaded_ && window_width_ > 0 && window_height_ > 0)
    {
        show_cmd = window_maximized_ ? SW_MAXIMIZE : SW_SHOWNORMAL;
    }
    else if (window_placement_loaded_ && window_maximized_)
    {
        show_cmd = SW_MAXIMIZE;
    }
    if (show_cmd != SW_MAXIMIZE)
    {
        RECT rect = {};
        if (GetWindowRect(hwnd_, &rect))
        {
            RECT fitted = rect;
            win32::ClampToWorkArea(&fitted);
            if (!EqualRect(&fitted, &rect))
            {
                SetWindowPos(hwnd_, nullptr, fitted.left, fitted.top, fitted.right - fitted.left, fitted.bottom - fitted.top, SWP_NOZORDER | SWP_NOACTIVATE);
            }
        }
    }
    ShowWindow(hwnd_, show_cmd);
    UpdateWindow(hwnd_);
    PostMessageW(hwnd_, frame::message_id::kDeferredStartup, 0, 0);
    PostMessageW(hwnd_, frame::message_id::kLoadTraces, 0, 0);
    PostMessageW(hwnd_, frame::message_id::kLoadDefaults, 0, 0);
}

void MainWindow::Impl::QueueExternalJump(const std::wstring& target)
{
    queued_external_jump_target_ = target;
}

void MainWindow::Impl::BeginJumpUiBatch()
{
    if (jump_ui_batch_active_)
    {
        return;
    }
    jump_ui_batch_active_ = true;
    if (browse_.tree().hwnd())
    {
        SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, FALSE, 0);
    }
    if (browse_.values().hwnd())
    {
        SendMessageW(browse_.values().hwnd(), WM_SETREDRAW, FALSE, 0);
    }
}

void MainWindow::Impl::EndJumpUiBatch()
{
    if (!jump_ui_batch_active_)
    {
        return;
    }
    jump_ui_batch_active_ = false;
    if (browse_.tree().hwnd())
    {
        SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, TRUE, 0);
        InvalidateRect(browse_.tree().hwnd(), nullptr, TRUE);
    }
    if (browse_.values().hwnd())
    {
        SendMessageW(browse_.values().hwnd(), WM_SETREDRAW, TRUE, 0);
        InvalidateRect(browse_.values().hwnd(), nullptr, TRUE);
    }
}

void MainWindow::Impl::ApplyTreeSelectionEffects(RegistryNode* node)
{
    ResetValueFilter();
    UpdateAddressBar(node);
    UpdateValueListForNode(node);
    MarkTreeStateDirty();
}

void MainWindow::Impl::FocusAddressBarForExternalJump(bool defer_if_needed)
{
    if (!hwnd_)
    {
        return;
    }
    HWND target = browse_.address();
    if (!target)
    {
        return;
    }
    if (IsWindowVisible(hwnd_) && !IsIconic(hwnd_))
    {
        SetFocus(target);
        SendMessageW(target, EM_SETSEL, 0, -1);
        return;
    }
    if (defer_if_needed)
    {
        PostMessageW(hwnd_, frame::message_id::kFocusAddressBar, 0, 0);
    }
}

void MainWindow::Impl::CyclePaneFocus(bool forward)
{
    HWND panes[] = {toolbar_.hwnd(), browse_.address(), tab_, browse_.filter(), browse_.tree().hwnd(), browse_.values().hwnd(), search_results_list_, history_list_};
    HWND visible[_countof(panes)] = {};
    int count = 0;
    for (HWND pane : panes)
    {
        if (pane && IsWindowVisible(pane) && IsWindowEnabled(pane))
        {
            visible[count++] = pane;
        }
    }
    if (count == 0)
    {
        return;
    }
    HWND focus = GetFocus();
    int current = -1;
    for (int i = 0; i < count; ++i)
    {
        if (visible[i] == focus || IsChild(visible[i], focus))
        {
            current = i;
            break;
        }
    }
    const int next = current < 0 ? (forward ? 0 : count - 1) : (current + (forward ? 1 : count - 1)) % count;
    FocusPane(visible[next]);
}

void MainWindow::Impl::FocusPane(HWND pane)
{
    if (!pane)
    {
        return;
    }
    SetFocus(pane);
    if (pane != browse_.values().hwnd() && pane != search_results_list_ && pane != history_list_)
    {
        return;
    }
    if (ListView_GetItemCount(pane) <= 0 || ListView_GetNextItem(pane, -1, LVNI_FOCUSED) >= 0 ||
        ListView_GetSelectedCount(pane) > 0)
    {
        return;
    }
    SelectListRowAtIndex(pane, 0);
}

bool MainWindow::Impl::TranslateAccelerator(const MSG& msg)
{
    if (key_handles_window_ && IsDialogMessageW(key_handles_window_, const_cast<MSG*>(&msg)))
    {
        return true;
    }
    if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)
    {
        const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        const bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
        HWND focus = GetFocus();
        auto is_text_input = [](HWND hwnd) -> bool {
            if (!hwnd)
            {
                return false;
            }
            wchar_t cls[64] = {};
            GetClassNameW(hwnd, cls, static_cast<int>(_countof(cls)));
            if (util::EqualsInsensitive(cls, L"Edit"))
            {
                return true;
            }
            if (util::EqualsInsensitive(cls, L"RichEdit20W") || util::EqualsInsensitive(cls, L"RichEdit20A"))
            {
                return true;
            }
            if (util::EqualsInsensitive(cls, L"ComboBox") || util::EqualsInsensitive(cls, L"ComboBoxEx32"))
            {
                return true;
            }
            HWND parent = GetParent(hwnd);
            if (parent)
            {
                GetClassNameW(parent, cls, static_cast<int>(_countof(cls)));
                if (util::EqualsInsensitive(cls, L"ComboBox") || util::EqualsInsensitive(cls, L"ComboBoxEx32"))
                {
                    return true;
                }
            }
            return false;
        };
        const bool focus_edit = is_text_input(focus);
        if (!ctrl && !shift && !alt && msg.wParam == VK_F7 && (HIWORD(msg.lParam) & KF_REPEAT) != 0)
        {
            return true;
        }

        if ((alt && msg.wParam == 'D') || (ctrl && !alt && msg.wParam == 'L'))
        {
            if (browse_.address())
            {
                SetFocus(browse_.address());
                SendMessageW(browse_.address(), EM_SETSEL, 0, -1);
                return true;
            }
        }

        if (ctrl && !alt && !shift && msg.wParam == 'A')
        {
            if (SelectAllInFocusedList())
            {
                return true;
            }
            if (focus_edit && focus)
            {
                SendMessageW(focus, EM_SETSEL, 0, -1);
                return true;
            }
        }

        if (ctrl && !alt && shift && msg.wParam == 'G')
        {
            if (OpenSelectedSearchResult(true))
            {
                return true;
            }
        }

        if (!ctrl && !alt)
        {
            if (msg.wParam == VK_TAB && !focus_edit)
            {
                HWND tree = browse_.tree().hwnd();
                HWND values = browse_.values().hwnd();
                const bool tree_ready = tree && IsWindowVisible(tree);
                const bool values_ready = values && IsWindowVisible(values);
                HWND next = (focus == values && tree_ready) ? tree
                            : values_ready                  ? values
                            : tree_ready                    ? tree
                                                            : nullptr;
                if (next)
                {
                    FocusPane(next);
                    return true;
                }
            }
            if (msg.wParam == VK_F6)
            {
                CyclePaneFocus(!shift);
                return true;
            }
            if (msg.wParam == VK_BACK && !focus_edit)
            {
                HandleMenuCommand(cmd::kNavBack);
                return true;
            }
        }

        if (focus_edit)
        {
            if (msg.wParam == VK_DELETE || msg.wParam == VK_BACK || msg.wParam == VK_F2)
            {
                return false;
            }
            if (ctrl && !alt)
            {
                switch (msg.wParam)
                {
                case 'C':
                case 'V':
                case 'X':
                case 'Z':
                case 'Y':
                    return false;
                default:
                    break;
                }
            }
        }
    }
    if (accelerators_)
    {
        return ::TranslateAcceleratorW(hwnd_, accelerators_, const_cast<MSG*>(&msg)) != 0;
    }
    return false;
}

LRESULT CALLBACK MainWindow::Impl::WndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam)
{
    if (message == WM_NCCREATE)
    {
        auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
        auto* self = static_cast<MainWindow::Impl*>(create->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        self->hwnd_ = hwnd;
    }

    auto* self = reinterpret_cast<MainWindow::Impl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!self)
    {
        return DefWindowProcW(hwnd, message, wparam, lparam);
    }
    return self->HandleMessage(message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::AddressEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data)
{
    auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
    if (message == WM_CONTEXTMENU)
    {
        POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        if (pt.x == -1 && pt.y == -1)
        {
            RECT rect = {};
            GetWindowRect(hwnd, &rect);
            pt.x = rect.left + 8;
            pt.y = rect.top + (rect.bottom - rect.top) / 2;
        }
        if (self)
        {
            self->ShowAddressContextMenu(hwnd, pt);
        }
        return 0;
    }
    if (message == WM_CHAR && wparam == 0x16 && (GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        return 0;
    }
    if (message == WM_KEYDOWN && wparam == 'V' && (GetKeyState(VK_CONTROL) & 0x8000) != 0 &&
        (GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        if ((GetWindowLongPtrW(hwnd, GWL_STYLE) & ES_READONLY) == 0 && IsClipboardFormatAvailable(CF_UNICODETEXT))
        {
            SendMessageW(hwnd, EM_SETSEL, 0, -1);
            SendMessageW(hwnd, WM_PASTE, 0, 0);
            SendMessageW(hwnd, WM_KEYDOWN, VK_RETURN, 0);
        }
        return 0;
    }
    if (message == WM_KEYDOWN && wparam == VK_RETURN)
    {
        if (self && self->hwnd_)
        {
            SendMessageW(self->hwnd_, frame::message_id::kAddressEnter, 0, 0);
        }
        return 0;
    }
    if (message == WM_CHAR && wparam == VK_RETURN)
    {
        return 0;
    }
    if (message == WM_SETFOCUS)
    {
        LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
        SendMessageW(hwnd, EM_SETSEL, 0, -1);
        SendMessageW(hwnd, EM_SCROLLCARET, 0, 0);
        return result;
    }
    if (message == WM_KEYUP)
    {
        LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
        auto* window = reinterpret_cast<MainWindow::Impl*>(ref_data);
        if (window)
        {
            window->ApplyAutoCompleteTheme();
        }
        return result;
    }
    if (message == WM_LBUTTONDOWN)
    {
        if (GetFocus() != hwnd)
        {
            LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
            SendMessageW(hwnd, EM_SETSEL, 0, -1);
            SendMessageW(hwnd, EM_SCROLLCARET, 0, 0);
            return result;
        }
    }
    return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::FilterEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data)
{
    auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
    if (message == WM_KEYDOWN && (wparam == VK_RETURN || wparam == VK_DOWN))
    {
        if (self)
        {
            self->FocusFirstValue();
        }
        return 0;
    }
    if (message == WM_KEYDOWN && wparam == VK_ESCAPE)
    {
        if (self)
        {
            self->ClearValueFilter(GetWindowTextLengthW(hwnd) == 0);
        }
        return 0;
    }
    if (message == WM_CHAR && wparam == VK_ESCAPE)
    {
        return 0;
    }
    if (message == WM_CHAR && wparam == VK_RETURN)
    {
        return 0;
    }
    return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::TabProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data)
{
    auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
    if (!self)
    {
        return DefSubclassProc(hwnd, message, wparam, lparam);
    }

    switch (message)
    {
    case WM_ERASEBKGND:
        return 1;
    case WM_MOUSEMOVE:
        {
            POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
            self->UpdateTabHotState(hwnd, pt);
            if (!self->tab_mouse_tracking_)
            {
                TRACKMOUSEEVENT tme = {};
                tme.cbSize = sizeof(tme);
                tme.dwFlags = TME_LEAVE;
                tme.hwndTrack = hwnd;
                TrackMouseEvent(&tme);
                self->tab_mouse_tracking_ = true;
            }
            return 0;
        }
    case WM_MOUSELEAVE:
        self->tab_mouse_tracking_ = false;
        if (self->tab_hot_index_ != -1 || self->tab_close_hot_index_ != -1)
        {
            self->tab_hot_index_ = -1;
            self->tab_close_hot_index_ = -1;
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    case WM_LBUTTONDOWN:
        {
            POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
            TCHITTESTINFO hit = {};
            hit.pt = pt;
            int index = TabCtrl_HitTest(hwnd, &hit);
            RECT close_rect = {};
            if (self->GetTabCloseRect(index, &close_rect) && PtInRect(&close_rect, pt))
            {
                self->tab_close_down_index_ = index;
                SetCapture(hwnd);
                InvalidateRect(hwnd, nullptr, FALSE);
                return 0;
            }
            if (self->tab_close_down_index_ != -1)
            {
                self->tab_close_down_index_ = -1;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            break;
        }
    case WM_LBUTTONUP:
        {
            if (self->tab_close_down_index_ >= 0)
            {
                POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
                int close_index = self->tab_close_down_index_;
                self->tab_close_down_index_ = -1;
                ReleaseCapture();
                RECT close_rect = {};
                if (self->GetTabCloseRect(close_index, &close_rect) && PtInRect(&close_rect, pt))
                {
                    self->CloseTab(close_index);
                    self->tab_hot_index_ = -1;
                    self->tab_close_hot_index_ = -1;
                }
                InvalidateRect(hwnd, nullptr, FALSE);
                return 0;
            }
            break;
        }
    case WM_CAPTURECHANGED:
        if (self->tab_close_down_index_ >= 0)
        {
            self->tab_close_down_index_ = -1;
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps = {};
            HDC hdc = BeginPaint(hwnd, &ps);
            self->PaintTabControl(hwnd, hdc);
            EndPaint(hwnd, &ps);
            return 0;
        }
    default:
        break;
    }

    return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::ListViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data)
{
    auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
    if (self && InSendMessage() &&
        (message == WM_CHAR || message == WM_KEYDOWN || message == WM_SETFOCUS || message == LVM_GETITEMCOUNT ||
         message == LVM_FINDITEMW || message == LVM_GETITEMTEXTW || message == LVM_GETITEMW))
    {
        self->FlushExternalNavigation();
        if (message == WM_SETFOCUS && hwnd != self->browse_.values().hwnd())
        {
            SetFocus(self->browse_.values().hwnd());
            return 0;
        }
    }
    if (message == WM_PAINT && ListViewScrolledHorizontally(hwnd))
    {
        InvalidateListViewTail(hwnd);
    }

    if (message == WM_KEYDOWN && wparam == VK_RETURN && self && hwnd == self->browse_.values().hwnd())
    {
        self->value_activate_from_key_ = true;
    }

    if (message == WM_NCDESTROY)
    {
        RemovePropW(hwnd, kListScrollProp);
    }

    if (message == WM_MOUSEMOVE && self && self->value_tooltip_ &&
        (hwnd == self->browse_.values().hwnd() || hwnd == self->search_results_list_ || hwnd == self->history_list_))
    {
        LVHITTESTINFO hit = {};
        hit.pt.x = GET_X_LPARAM(lparam);
        hit.pt.y = GET_Y_LPARAM(lparam);
        const int item = ListView_SubItemHitTest(hwnd, &hit);
        if (hwnd != self->value_tip_list_ || item != self->value_tip_item_ || hit.iSubItem != self->value_tip_subitem_)
        {
            self->value_tip_list_ = hwnd;
            self->value_tip_item_ = item;
            self->value_tip_subitem_ = hit.iSubItem;
            SendMessageW(self->value_tooltip_, TTM_POP, 0, 0);
        }
    }
    if (message == WM_CHAR && self && (hwnd == self->browse_.values().hwnd() || InSendMessage()))
    {
        wchar_t ch = static_cast<wchar_t>(wparam);
        if (ch == L'\b' || (iswprint(ch) && ch != L'\r' && ch != L'\n' && ch != L'\t'))
        {
            self->HandleTypeToSelectList(ch);
            return 0;
        }
    }
    if (message == WM_SETFOCUS || message == WM_KILLFOCUS)
    {
        if (self && message == WM_SETFOCUS)
        {
            self->last_focus_ = hwnd;
        }
        SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
    }
    if (message == WM_UPDATEUISTATE)
    {
        LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
        SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
        return result;
    }
    if (message == WM_CTLCOLOREDIT)
    {
        HDC hdc = reinterpret_cast<HDC>(wparam);
        SetTextColor(hdc, Theme::Current().TextColor());
        SetBkColor(hdc, Theme::Current().PanelColor());
        return reinterpret_cast<LRESULT>(Theme::Current().PanelBrush());
    }
    if (message == WM_PRINTCLIENT)
    {
        HDC hdc = reinterpret_cast<HDC>(wparam);
        RECT rect = {};
        GetClientRect(hwnd, &rect);
        FillRect(hdc, &rect, Theme::Current().PanelBrush());
    }
    if (message == WM_THEMECHANGED)
    {
        if (self)
        {
            InvalidateRect(hwnd, nullptr, TRUE);
        }
    }
    return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::TreeViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data)
{
    auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
    if (message == TVM_SELECTITEM && wparam == TVGN_CARET && self && hwnd == self->regedit_compat_tree_.hwnd())
    {
        const HTREEITEM item = reinterpret_cast<HTREEITEM>(lparam);
        const bool unchanged = TreeView_GetSelection(hwnd) == item;
        const LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
        if (unchanged)
        {
            RegistryNode* node = self->regedit_compat_tree_.NodeFromItem(item);
            if (node)
            {
                self->QueueCompatJump(*node);
            }
        }
        return result;
    }
    if (message == WM_SETFOCUS && self)
    {
        self->last_focus_ = hwnd;
    }
    if (message == WM_KEYDOWN && self && hwnd == self->browse_.tree().hwnd())
    {
        switch (wparam)
        {
        case VK_RIGHT:
            self->browse_.set_tree_type_select_descend(true);
            break;
        case VK_LEFT:
        case VK_UP:
        case VK_DOWN:
        case VK_HOME:
        case VK_END:
        case VK_PRIOR:
        case VK_NEXT:
            self->browse_.set_tree_type_select_descend(false);
            break;
        default:
            break;
        }
    }
    if (message == WM_CHAR && self && hwnd == self->browse_.tree().hwnd())
    {
        wchar_t ch = static_cast<wchar_t>(wparam);
        if (ch == L'\b' || (iswprint(ch) && ch != L'\r' && ch != L'\n' && ch != L'\t'))
        {
            self->HandleTypeToSelectTree(ch);
            return 0;
        }
    }
    return DefSubclassProc(hwnd, message, wparam, lparam);
}

} // namespace regkit
