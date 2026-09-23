// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

namespace regkit
{

using namespace window_detail;

void MainWindow::Impl::PushUndo(changes::UndoOperation operation)
{
    if (is_replaying_)
    {
        return;
    }
    undo_stack_.Push(std::move(operation));
    if (toolbar_.hwnd())
    {
        SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
        SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, 0);
    }
}

void MainWindow::Impl::ClearRedo()
{
    undo_stack_.ClearRedo();
    if (toolbar_.hwnd())
    {
        SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
        SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, 0);
    }
}

MainWindow::Impl::ReplayResult MainWindow::Impl::ApplyUndoOperation(const changes::UndoOperation& operation, bool redo)
{
    if (!browse_.current_node())
    {
        return ReplayResult::kUnchanged;
    }
    bool ok = false;
    bool rename_left_both_names = false;
    std::optional<std::wstring> restored_value;
    is_replaying_ = true;
    switch (operation.type)
    {
    case changes::UndoOperation::Type::kCreateKey:
        {
            if (redo)
            {
                if (!operation.key_snapshot.name.empty())
                {
                    ok = changes::RestoreKey(operation.node, operation.key_snapshot);
                }
                else
                {
                    ok = RegistryStore::CreateKey(operation.node, operation.name);
                }
                if (ok)
                {
                    RefreshTreeSelection();
                    SelectChildKey(operation.node, operation.name);
                }
            }
            else
            {
                RegistryNode child = ChildNode(operation.node, operation.name);
                ok = RegistryStore::DeleteKey(child);
                if (ok)
                {
                    RefreshTreeSelection();
                }
            }
            break;
        }
    case changes::UndoOperation::Type::kDeleteKey:
        {
            if (redo)
            {
                RegistryNode child = ChildNode(operation.node, operation.name);
                ok = RegistryStore::DeleteKey(child);
                if (ok)
                {
                    RefreshTreeSelection();
                }
            }
            else
            {
                ok = changes::RestoreKey(operation.node, operation.key_snapshot);
                if (ok)
                {
                    RefreshTreeSelection();
                    SelectChildKey(operation.node, operation.key_snapshot.name);
                }
            }
            break;
        }
    case changes::UndoOperation::Type::kRenameKey:
        {
            std::wstring from = redo ? operation.name : operation.new_name;
            std::wstring to = redo ? operation.new_name : operation.name;
            RegistryNode child = ChildNode(operation.node, from);
            ok = RegistryStore::RenameKey(child, to);
            if (ok)
            {
                RefreshTreeSelection();
                std::wstring path = registry_path::Build(operation.node);
                if (!path.empty())
                {
                    path.append(L"\\");
                    path.append(to);
                    SelectTreePath(path);
                }
            }
            break;
        }
    case changes::UndoOperation::Type::kCreateValue:
        {
            if (redo)
            {
                ok = RegistryStore::SetValue(operation.node, operation.new_value.name, operation.new_value.type, operation.new_value.data);
                restored_value = operation.new_value.name;
            }
            else
            {
                ok = RegistryStore::DeleteValue(operation.node, operation.name);
            }
            break;
        }
    case changes::UndoOperation::Type::kDeleteValue:
        {
            if (redo)
            {
                ok = RegistryStore::DeleteValue(operation.node, operation.old_value.name);
            }
            else
            {
                ok = RegistryStore::SetValue(operation.node, operation.old_value.name, operation.old_value.type, operation.old_value.data);
                restored_value = operation.old_value.name;
            }
            break;
        }
    case changes::UndoOperation::Type::kModifyValue:
        {
            const ValueEntry& value = redo ? operation.new_value : operation.old_value;
            ok = RegistryStore::SetValue(operation.node, value.name, value.type, value.data);
            break;
        }
    case changes::UndoOperation::Type::kRenameValue:
        {
            std::wstring from = redo ? operation.name : operation.new_name;
            std::wstring to = redo ? operation.new_name : operation.name;
            ok = RegistryStore::RenameValue(operation.node, from, to, &rename_left_both_names);
            break;
        }
    default:
        break;
    }
    is_replaying_ = false;

    if (ok || rename_left_both_names)
    {
        MarkOfflineDirty();
    }
    if ((ok || rename_left_both_names) && browse_.current_node())
    {
        UpdateValueListForNode(browse_.current_node());
        if (ok && restored_value)
        {
            SelectValueAfterRefresh(*restored_value);
        }
    }
    if (rename_left_both_names)
    {
        ui::ShowError(hwnd_, L"The value was copied to the new name but the old name "
                             L"couldn't be removed. Both names now exist.");
    }
    if (toolbar_.hwnd())
    {
        SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
        SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, undo_stack_.CanRedo() ? TBSTATE_ENABLED : 0);
    }
    if (rename_left_both_names)
    {
        return ReplayResult::kPartial;
    }
    return ok ? ReplayResult::kSuccess : ReplayResult::kUnchanged;
}

bool MainWindow::Impl::SameNode(const RegistryNode& left, const RegistryNode& right) const
{
    if (left.root != right.root)
    {
        return false;
    }
    if (!EqualsInsensitive(left.subkey, right.subkey))
    {
        return false;
    }
    return EqualsInsensitive(left.root_name, right.root_name);
}

std::wstring MainWindow::Impl::MakeUniqueValueName(const RegistryNode& node, const std::wstring& base) const
{
    std::unordered_set<std::wstring> value_names;
    RegistryStore::KeyEnumResult enum_result;
    bool names_reserved = false;
    RegistryStore::EnumKeyStreaming(node, true, false, false, &enum_result, [&](const ValueInfo& value, const BYTE*, DWORD) {
        if (!names_reserved)
        {
            if (enum_result.info_valid)
            {
                value_names.reserve(enum_result.info.value_count);
            }
            names_reserved = true;
        }
        value_names.insert(ToLower(value.name));
        return true;
    },
                                    {});
    auto exists = [&](const std::wstring& candidate) -> bool { return value_names.contains(ToLower(candidate)); };

    std::wstring base_name = base;
    if (base_name.empty())
    {
        if (!exists(base_name))
        {
            return base_name;
        }
        base_name = L"Default";
    }
    if (!exists(base_name))
    {
        return base_name;
    }
    for (int i = 2; i < 10000; ++i)
    {
        std::wstring next = base_name + L" (" + std::to_wstring(i) + L")";
        if (!exists(next))
        {
            return next;
        }
    }
    return base_name;
}

std::wstring MainWindow::Impl::MakeUniqueKeyName(const RegistryNode& node, const std::wstring& base) const
{
    auto keys = RegistryStore::EnumSubKeyNames(node, false);
    auto exists = [&](const std::wstring& candidate) -> bool {
        for (const auto& key : keys)
        {
            if (EqualsInsensitive(key, candidate))
            {
                return true;
            }
        }
        return false;
    };

    std::wstring base_name = base;
    if (base_name.empty())
    {
        base_name = L"New Key";
    }
    if (!exists(base_name))
    {
        return base_name;
    }
    for (int i = 2; i < 10000; ++i)
    {
        std::wstring next = base_name + L" (" + std::to_wstring(i) + L")";
        if (!exists(next))
        {
            return next;
        }
    }
    return base_name;
}

bool MainWindow::Impl::ResolvePathToNode(const std::wstring& path, RegistryNode* node) const
{
    if (!node || path.empty())
    {
        return false;
    }
    for (const auto& root_entry : browse_.roots())
    {
        if (!StartsWithInsensitive(path, root_entry.path_name))
        {
            continue;
        }
        std::wstring rest = path.substr(root_entry.path_name.size());
        if (!rest.empty() && (rest.front() == L'\\' || rest.front() == L'/'))
        {
            rest.erase(rest.begin());
        }
        if (root_entry.subkey_prefix.empty())
        {
            node->root = root_entry.root;
            node->root_name = root_entry.path_name;
            node->subkey = rest;
            return true;
        }
        std::wstring prefix = root_entry.subkey_prefix;
        if (!rest.empty())
        {
            if (!StartsWithInsensitive(rest, prefix))
            {
                rest = prefix + L"\\" + rest;
            }
        }
        else
        {
            rest = prefix;
        }
        node->root = root_entry.root;
        node->root_name = root_entry.path_name;
        node->subkey = rest;
        return true;
    }
    return false;
}

void MainWindow::Impl::ShowHeaderMenu(HWND list, std::vector<ColumnInfo>& columns, std::vector<int>& widths, std::vector<bool>& visible, POINT screen_pt, int unavailable_column)
{
    HWND header_hwnd = ListView_GetHeader(list);
    if (!header_hwnd)
    {
        return;
    }
    POINT client_pt = screen_pt;
    ScreenToClient(header_hwnd, &client_pt);
    HDHITTESTINFO hit = {};
    hit.pt = client_pt;
    const int column_hit = static_cast<int>(SendMessageW(header_hwnd, HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hit)));

    HMENU menu = CreatePopupMenu();
    if (!menu)
    {
        return;
    }
    const UINT fit_flags = MF_STRING | ((column_hit >= 0) ? 0 : MF_GRAYED);
    AppendMenuW(menu, fit_flags, cmd::kHeaderSizeToFit, L"Size column to fit");
    AppendMenuW(menu, MF_STRING, cmd::kHeaderSizeAll, L"Size all columns to fit");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

    for (size_t i = 0; i < columns.size(); ++i)
    {
        if (static_cast<int>(i) == unavailable_column)
        {
            continue;
        }
        const UINT state = i < visible.size() && visible[i] ? MF_CHECKED : MF_UNCHECKED;
        AppendMenuW(menu, MF_STRING | state, cmd::kHeaderToggleBase + static_cast<int>(i), columns[i].title.c_str());
    }

    const int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_NONOTIFY, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
    DestroyMenu(menu);

    if (command == cmd::kHeaderSizeToFit && column_hit >= 0)
    {
        const int subitem = GetListViewColumnSubItem(list, column_hit);
        ListView_SetColumnWidth(list, column_hit, LVSCW_AUTOSIZE_USEHEADER);
        if (subitem >= 0 && static_cast<size_t>(subitem) < widths.size())
        {
            widths[static_cast<size_t>(subitem)] = ListView_GetColumnWidth(list, column_hit);
        }
    }
    else if (command == cmd::kHeaderSizeAll)
    {
        int last_visible = -1;
        for (size_t i = 0; i < columns.size(); ++i)
        {
            if (static_cast<int>(i) != unavailable_column && (i >= visible.size() || visible[i]))
            {
                last_visible = static_cast<int>(i);
            }
        }
        for (size_t i = 0; i < columns.size(); ++i)
        {
            if (static_cast<int>(i) == unavailable_column || (i < visible.size() && !visible[i]))
            {
                continue;
            }
            const int display = FindListViewColumnBySubItem(list, static_cast<int>(i));
            if (display < 0)
            {
                continue;
            }
            int width = 0;
            if (static_cast<int>(i) == last_visible)
            {
                width = CalcListViewColumnFitWidth(list, static_cast<int>(i), columns[i].width);
                ListView_SetColumnWidth(list, display, width);
            }
            else
            {
                ListView_SetColumnWidth(list, display, LVSCW_AUTOSIZE_USEHEADER);
                width = ListView_GetColumnWidth(list, display);
            }
            widths[i] = width;
        }
    }
    else if (command >= cmd::kHeaderToggleBase)
    {
        const int index = command - cmd::kHeaderToggleBase;
        if (index < 0 || static_cast<size_t>(index) >= columns.size() || index == unavailable_column)
        {
            return;
        }
        const bool show = !(static_cast<size_t>(index) < visible.size() && visible[static_cast<size_t>(index)]);
        if (list == browse_.values().hwnd())
        {
            ToggleValueColumn(index, show);
        }
        else if (list == history_list_)
        {
            ToggleHistoryColumn(index, show);
        }
        else if (list == search_results_list_)
        {
            ToggleSearchColumn(index, show);
        }
    }

    if (list == browse_.values().hwnd() && command != 0)
    {
        SaveSettings();
    }
}

void MainWindow::Impl::ShowValueHeaderMenu(POINT screen_pt)
{
    ShowHeaderMenu(browse_.values().hwnd(), browse_.columns().items, browse_.columns().widths, browse_.columns().visible, screen_pt);
}

void MainWindow::Impl::ShowHistoryHeaderMenu(POINT screen_pt)
{
    ShowHeaderMenu(history_list_, history_columns_, history_column_widths_, history_column_visible_, screen_pt);
}

void MainWindow::Impl::ShowSearchHeaderMenu(POINT screen_pt)
{
    const bool compare = IsCompareTabSelected();
    auto& columns = compare ? compare_columns_ : search_columns_;
    auto& widths = compare ? compare_column_widths_ : search_column_widths_;
    auto& visible = compare ? compare_column_visible_ : search_column_visible_;
    ShowHeaderMenu(search_results_list_, columns, widths, visible, screen_pt, compare && !IsCompareResultColumnAvailable() ? 4 : -1);
}

void MainWindow::Impl::ToggleValueColumn(int column, bool visible)
{
    if (column < 0 || static_cast<size_t>(column) >= browse_.columns().visible.size())
    {
        return;
    }
    if (visible == browse_.columns().visible[static_cast<size_t>(column)])
    {
        return;
    }

    if (visible)
    {
        int width = browse_.columns().widths[static_cast<size_t>(column)];
        if (width <= 0)
        {
            width = browse_.columns().items[static_cast<size_t>(column)].width;
        }
        browse_.columns().visible[static_cast<size_t>(column)] = true;
        browse_.columns().widths[static_cast<size_t>(column)] = width;
    }
    else
    {
        int display_index = FindListViewColumnBySubItem(browse_.values().hwnd(), column);
        int width = display_index >= 0 ? ListView_GetColumnWidth(browse_.values().hwnd(), display_index)
                                       : browse_.columns().widths[static_cast<size_t>(column)];
        if (width > 0)
        {
            browse_.columns().widths[static_cast<size_t>(column)] = width;
        }
        browse_.columns().visible[static_cast<size_t>(column)] = false;
    }
    ApplyValueColumns();
}

void MainWindow::Impl::ToggleHistoryColumn(int column, bool visible)
{
    if (column < 0 || static_cast<size_t>(column) >= history_column_visible_.size())
    {
        return;
    }
    if (visible == history_column_visible_[static_cast<size_t>(column)])
    {
        return;
    }

    if (visible)
    {
        int width = history_column_widths_[static_cast<size_t>(column)];
        if (width <= 0)
        {
            width = history_columns_[static_cast<size_t>(column)].width;
        }
        history_column_visible_[static_cast<size_t>(column)] = true;
        history_column_widths_[static_cast<size_t>(column)] = width;
    }
    else
    {
        int display_index = FindListViewColumnBySubItem(history_list_, column);
        int width = display_index >= 0 ? ListView_GetColumnWidth(history_list_, display_index)
                                       : history_column_widths_[static_cast<size_t>(column)];
        if (width > 0)
        {
            history_column_widths_[static_cast<size_t>(column)] = width;
        }
        history_column_visible_[static_cast<size_t>(column)] = false;
    }
    ApplyHistoryColumns();
}

void MainWindow::Impl::ToggleSearchColumn(int column, bool visible)
{
    bool compare = IsCompareTabSelected();
    auto& columns = compare ? compare_columns_ : search_columns_;
    auto& widths = compare ? compare_column_widths_ : search_column_widths_;
    auto& visibility = compare ? compare_column_visible_ : search_column_visible_;
    if (column < 0 || static_cast<size_t>(column) >= visibility.size() || static_cast<size_t>(column) >= columns.size())
    {
        return;
    }
    if (visible == visibility[static_cast<size_t>(column)])
    {
        return;
    }

    if (visible)
    {
        int width = widths[static_cast<size_t>(column)];
        if (width <= 0)
        {
            width = columns[static_cast<size_t>(column)].width;
        }
        visibility[static_cast<size_t>(column)] = true;
        widths[static_cast<size_t>(column)] = width;
    }
    else
    {
        int display_index = FindListViewColumnBySubItem(search_results_list_, column);
        int width = display_index >= 0 ? ListView_GetColumnWidth(search_results_list_, display_index)
                                       : widths[static_cast<size_t>(column)];
        if (width > 0)
        {
            widths[static_cast<size_t>(column)] = width;
        }
        visibility[static_cast<size_t>(column)] = false;
    }
    ApplySearchColumns(compare);
}

void MainWindow::Impl::DrawAddressButton(const DRAWITEMSTRUCT* info)
{
    if (!info)
    {
        return;
    }
    const Theme& theme = Theme::Current();
    HDC hdc = info->hDC;
    RECT rect = info->rcItem;
    bool pressed = (info->itemState & ODS_SELECTED) != 0;

    COLORREF bg_color = pressed ? theme.HoverColor() : theme.SurfaceColor();
    FillRect(hdc, &rect, appearance::CachedBrush(bg_color));

    HPEN pen = appearance::CachedPen(theme.BorderColor());
    HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
    MoveToEx(hdc, rect.left, rect.top + 3, nullptr);
    LineTo(hdc, rect.left, rect.bottom - 3);
    SelectObject(hdc, old_pen);

    if (info->CtlID == kAddressGoId)
    {
        if (address_go_icon_)
        {
            UINT dpi = win32::DpiForWindow(hwnd_);
            int icon_size = util::ScaleForDpi(kToolbarGlyphSize, dpi);
            int icon_x = rect.left + (rect.right - rect.left - icon_size) / 2;
            int icon_y = rect.top + (rect.bottom - rect.top - icon_size) / 2;
            if (info->itemState & ODS_DISABLED)
            {
                DrawState(hdc, nullptr, nullptr, reinterpret_cast<LPARAM>(address_go_icon_), 0, icon_x, icon_y, icon_size, icon_size, DST_ICON | DSS_DISABLED);
            }
            else
            {
                DrawIconEx(hdc, icon_x, icon_y, address_go_icon_, icon_size, icon_size, 0, nullptr, DI_NORMAL);
            }
        }
        else
        {
            POINT pts[3] = {
                {rect.left + 8, rect.top + 6},
                {rect.left + 8, rect.bottom - 6},
                {rect.right - 6, (rect.top + rect.bottom) / 2},
            };
            COLORREF arrow_color = theme.MutedTextColor();
            HBRUSH arrow_brush = appearance::CachedBrush(arrow_color);
            HBRUSH old_brush = reinterpret_cast<HBRUSH>(SelectObject(hdc, arrow_brush));
            HPEN arrow_pen = appearance::CachedPen(arrow_color);
            HPEN old_arrow = reinterpret_cast<HPEN>(SelectObject(hdc, arrow_pen));
            Polygon(hdc, pts, 3);
            SelectObject(hdc, old_arrow);
            SelectObject(hdc, old_brush);
        }
    }
}

void MainWindow::Impl::DrawHeaderCloseButton(const DRAWITEMSTRUCT* info)
{
    if (!info)
    {
        return;
    }
    const Theme& theme = Theme::Current();
    HDC hdc = info->hDC;
    RECT rect = info->rcItem;
    bool pressed = (info->itemState & ODS_SELECTED) != 0;

    COLORREF bg_color = pressed ? theme.HoverColor() : theme.HeaderColor();
    FillRect(hdc, &rect, appearance::CachedBrush(bg_color));

    DrawCloseGlyph(hdc, rect, theme.MutedTextColor(), win32::DpiForWindow(info->hwndItem));
}

void MainWindow::Impl::DrawFilterClearButton(const DRAWITEMSTRUCT* info)
{
    if (!info)
    {
        return;
    }
    const Theme& theme = Theme::Current();
    HDC hdc = info->hDC;
    RECT rect = info->rcItem;
    const bool pressed = (info->itemState & ODS_SELECTED) != 0;
    FillRect(hdc, &rect, appearance::CachedBrush(pressed ? theme.HoverColor() : theme.SurfaceColor()));
    HPEN pen = appearance::CachedPen(theme.BorderColor());
    HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
    MoveToEx(hdc, rect.left, rect.top + 3, nullptr);
    LineTo(hdc, rect.left, rect.bottom - 3);
    SelectObject(hdc, old_pen);
    DrawCloseGlyph(hdc, rect, theme.MutedTextColor(), win32::DpiForWindow(info->hwndItem));
}

void MainWindow::Impl::ClearValueFilter(bool focus_values)
{
    if (!browse_.filter())
    {
        return;
    }
    if (GetWindowTextLengthW(browse_.filter()) > 0)
    {
        SetWindowTextW(browse_.filter(), L"");
    }
    browse_.values().SetFilter(std::wstring());
    UpdateStatus();
    if (focus_values)
    {
        FocusPane(browse_.values().hwnd());
    }
}

bool MainWindow::Impl::SelectChildKey(const RegistryNode& parent, const std::wstring& name)
{
    if (name.empty())
    {
        return false;
    }
    std::wstring path = registry_path::Build(parent);
    if (path.empty())
    {
        return false;
    }
    path.append(L"\\");
    path.append(name);
    return SelectTreePath(path);
}

std::wstring MainWindow::Impl::TreeNeighbourPath(HTREEITEM item)
{
    HWND tree = browse_.tree().hwnd();
    if (!tree || !item)
    {
        return std::wstring();
    }
    HTREEITEM next = TreeView_GetNextSibling(tree, item);
    if (!next)
    {
        next = TreeView_GetPrevSibling(tree, item);
    }
    if (!next)
    {
        next = TreeView_GetParent(tree, item);
    }
    RegistryNode* node = next ? browse_.tree().NodeFromItem(next) : nullptr;
    return node ? registry_path::Build(*node) : std::wstring();
}

bool MainWindow::Impl::SelectTreePath(const std::wstring& path)
{
    if (!browse_.tree().hwnd())
    {
        return false;
    }
    std::vector<std::wstring> parts = BuildVisibleTreePathParts(path);
    if (parts.empty())
    {
        return false;
    }

    HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
    HTREEITEM current = root;
    for (const auto& part : parts)
    {
        TreeView_Expand(browse_.tree().hwnd(), current, TVE_EXPAND);
        HTREEITEM child = FindChildByText(browse_.tree().hwnd(), current, part);
        if (!child)
        {
            RefreshTreeItem(current);
            child = FindChildByText(browse_.tree().hwnd(), current, part);
        }
        if (!child)
        {
            return false;
        }
        current = child;
    }

    if (current)
    {
        TreeView_SelectItem(browse_.tree().hwnd(), current);
        TreeView_EnsureVisible(browse_.tree().hwnd(), current);
        return true;
    }
    return false;
}

bool MainWindow::Impl::SelectValueByName(const std::wstring& name)
{
    return browse_.SelectValue(name);
}

void MainWindow::Impl::SelectValueWhenReady(const std::wstring& name)
{
    pending_value_name_ = name;
    FocusPane(browse_.values().hwnd());
    if (!value_list_loading_ && SelectValueByName(name))
    {
        pending_value_name_.clear();
    }
}

void MainWindow::Impl::RestoreValueSelection()
{
    HWND list = browse_.values().hwnd();
    std::vector<std::wstring> names;
    names.swap(pending_value_selection_);
    const int top = pending_value_top_index_;
    pending_value_top_index_ = 0;
    pending_value_selection_key_.clear();
    if (!list)
    {
        return;
    }
    ListView_SetItemState(list, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    const std::unordered_set<std::wstring> wanted(names.begin(), names.end());
    bool first = true;
    const int rows = static_cast<int>(browse_.values().RowCount());
    for (int row_index = 0; row_index < rows; ++row_index)
    {
        const ListRow* row = browse_.values().RowAt(row_index);
        if (!row || row->kind != rowkind::kValue || !wanted.count(row->extra))
        {
            continue;
        }
        ListView_SetItemState(list, row_index, LVIS_SELECTED | (first ? LVIS_FOCUSED : 0), LVIS_SELECTED | LVIS_FOCUSED);
        first = false;
    }
    const int current_top = ListView_GetTopIndex(list);
    if (top > 0 && top != current_top)
    {
        RECT bounds = {};
        if (ListView_GetItemRect(list, 0, &bounds, LVIR_BOUNDS))
        {
            const int height = bounds.bottom - bounds.top;
            if (height > 0)
            {
                ListView_Scroll(list, 0, (top - current_top) * height);
            }
        }
    }
}

void MainWindow::Impl::SelectValueAfterRefresh(const std::wstring& name)
{
    if (!browse_.current_node())
    {
        return;
    }
    retained_value_name_ = name;
    retained_value_key_path_ = registry_path::Build(*browse_.current_node());
}

void MainWindow::Impl::SelectListRowAtIndex(HWND list, int index)
{
    if (!list || index < 0)
    {
        return;
    }
    const int count = ListView_GetItemCount(list);
    if (count <= 0)
    {
        return;
    }
    if (index >= count)
    {
        index = count - 1;
    }
    ListView_SetItemState(list, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetItemState(list, index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(list, index, FALSE);
}

void MainWindow::Impl::HandleTypeToSelectList(wchar_t ch)
{
    browse_.TypeSelectValues(ch, GetTickCount());
}

void MainWindow::Impl::HandleTypeToSelectTree(wchar_t ch)
{
    browse_.TypeSelectTree(ch, GetTickCount());
}

} // namespace regkit
