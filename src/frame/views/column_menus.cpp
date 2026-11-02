// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

namespace regkit
{

using namespace window_detail;

ui::ColumnSet* MainWindow::Impl::ColumnSetFor(HWND list)
{
    if (list == browse_.values().hwnd())
    {
        return &browse_.columns();
    }
    if (list == history_list_)
    {
        return &history_columns_;
    }
    if (list == search_results_list_)
    {
        return compare_columns_active_ ? &compare_columns_ : &search_columns_;
    }
    return nullptr;
}

void MainWindow::Impl::OnListColumnsChanged(HWND list)
{
    ui::ColumnSet* set = ColumnSetFor(list);
    if (!set)
    {
        return;
    }
    const std::vector<win32ui::ListColumn> columns = win32ui::ListColumns(list);
    for (size_t i = 0; i < columns.size() && i < set->widths.size() && i < set->visible.size(); ++i)
    {
        // unavailable columns retain user visibility choices for other views
        if (!columns[i].available)
        {
            continue;
        }
        set->visible[i] = columns[i].visible;
        // saved widths use DIPs, hidden columns report their last visible width
        set->widths[i] = columns[i].width_dip > 0 ? columns[i].width_dip : set->widths[i];
    }
    if (list == browse_.values().hwnd())
    {
        SaveSettings();
    }
}

void MainWindow::Impl::RescaleColumns()
{
    // widths use DIPs, lists scale them to the new DPI
    ApplyValueColumns();
    ApplyHistoryColumns();
    ApplySearchColumns(compare_columns_active_);
    // compare titles describe current sources, next search tab resets them
    search_results_view_tab_index_ = -1;
}

void MainWindow::Impl::AttachMainList(HWND list)
{
    if (!main_list_hooks_.cell_text)
    {
        main_list_hooks_ = win32ui::DialogListHooks();
        main_list_hooks_.context = this;
        main_list_hooks_.cell_text = [](void* context, HWND target, int row, int subitem, std::wstring* text) {
            *text = static_cast<MainWindow::Impl*>(context)->ListCellFieldText(target, row, subitem);
            return true;
        };
        main_list_hooks_.columns_changed = [](void* context, HWND target) { static_cast<MainWindow::Impl*>(context)->OnListColumnsChanged(target); };
    }
    win32ui::AttachListView(list, &main_list_hooks_);
    win32ui::ShowListGridButton(list);
}

}
