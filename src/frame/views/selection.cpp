// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

namespace regkit
{

using namespace window_detail;

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
    const std::unordered_set<std::wstring> wanted(names.begin(), names.end());
    win32ui::ListSelection selection;
    const int rows = static_cast<int>(browse_.values().RowCount());
    for (int row_index = 0; row_index < rows; ++row_index)
    {
        const ListRow* row = browse_.values().RowAt(row_index);
        if (row && row->kind == rowkind::kValue && wanted.contains(row->extra))
        {
            selection.rows.push_back(row_index);
        }
    }
    // first restored row takes focus and anchors Shift extensions
    selection.focused = selection.rows.empty() ? -1 : selection.rows.front();
    selection.mark = selection.focused;
    win32ui::ApplyListSelection(list, selection);
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
    win32ui::SelectOnlyListRow(list, std::min(index, count - 1));
}

}
