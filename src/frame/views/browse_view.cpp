// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include <win32ui/gdi/icon_loader.h>
#include <win32ui/window.h>
#include <win32ui/window_metrics.h>

#include <functional>

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "win32/translation.h"
#include "trace/trace_paths.h"

namespace regkit
{
using namespace window_detail;

void MainWindow::Impl::BuildImageLists()
{
    if (tree_images_)
    {
        ImageList_Destroy(tree_images_);
        tree_images_ = nullptr;
    }
    if (list_images_)
    {
        ImageList_Destroy(list_images_);
        list_images_ = nullptr;
    }

    UINT dpi = win32ui::DpiForWindow(hwnd_);

    const int icon_size = win32ui::ScaleForDpi(win32ui::kIconDip, dpi);
    tree_images_ = ImageList_Create(icon_size, icon_size, ILC_COLOR32, 12, 2);
    list_images_ = ImageList_Create(icon_size, icon_size, ILC_COLOR32, 10, 2);
    ImageList_SetBkColor(tree_images_, CLR_NONE);
    ImageList_SetBkColor(list_images_, CLR_NONE);

    auto add_icons = [&](std::initializer_list<HIMAGELIST> lists, std::initializer_list<int> resource_ids) {
        for (const int resource_id : resource_ids)
        {
            HICON icon = win32ui::LoadIconResource(resource_id, dpi);
            for (HIMAGELIST list : lists)
            {
                win32ui::ImageListAddOrBlank(list, icon, icon_size);
            }
            if (icon)
            {
                DestroyIcon(icon);
            }
        }
    };
    add_icons({tree_images_, list_images_}, {IDI_ICON_FOLDER, IDI_ICON_SYMLINK, IDI_ICON_DATABASE, IDI_ICON_FOLDER_SIM, IDI_ICON_FOLDER_DENIED, IDI_ICON_DATABASE_DENIED, IDI_ICON_FOLDER_VOLATILE});
    add_icons({tree_images_}, {IDI_ICON_ROOT_KEYS, IDI_APPICON, IDI_ICON_TREE_LOCAL_REGISTRY, IDI_ICON_TREE_REMOTE_REGISTRY, IDI_ICON_TREE_OFFLINE_REGISTRY});
    add_icons({list_images_}, {IDI_ICON_TEXT, IDI_ICON_BINARY, IDI_ICON_TRACE});
}

void MainWindow::Impl::CreateValueColumns()
{
    browse_.columns().Reset({
        {util::Tr(L"Name"), 260, LVCFMT_LEFT},
        {util::Tr(L"Type"), 120, LVCFMT_LEFT},
        {util::Tr(L"Data"), 160, LVCFMT_LEFT},
        {util::Tr(L"Default"), 200, LVCFMT_LEFT},
        {util::Tr(L"Read on boot"), 110, LVCFMT_LEFT},
        {util::Tr(L"Size"), 70, LVCFMT_RIGHT},
        {util::Tr(L"Date Modified"), 140, LVCFMT_LEFT},
        {util::Tr(L"Details"), 160, LVCFMT_LEFT},
        {util::Tr(L"Comment"), 220, LVCFMT_LEFT},
    });
    if (browse_.columns().saved)
    {
        auto patch_widths = [&](std::vector<int>& widths) {
            if (widths.size() == browse_.columns().items.size() - 1)
            {
                widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
            }
            else if (widths.size() == browse_.columns().items.size() - 2)
            {
                widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
                widths.push_back(browse_.columns().items[kValueColComment].width);
            }
            else if (widths.size() == browse_.columns().items.size() - 3)
            {
                widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
                widths.push_back(browse_.columns().items[kValueColDetails].width);
                widths.push_back(browse_.columns().items[kValueColComment].width);
            }
            else if (widths.size() == browse_.columns().items.size() - 4)
            {
                widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
                widths.insert(widths.begin() + kValueColReadOnBoot, browse_.columns().items[kValueColReadOnBoot].width);
                widths.push_back(browse_.columns().items[kValueColDetails].width);
                widths.push_back(browse_.columns().items[kValueColComment].width);
            }
        };
        auto patch_visible = [&](std::vector<bool>& visible) {
            if (visible.size() == browse_.columns().items.size() - 1)
            {
                visible.insert(visible.begin() + kValueColDefault, true);
            }
            else if (visible.size() == browse_.columns().items.size() - 2)
            {
                visible.insert(visible.begin() + kValueColDefault, true);
                visible.push_back(true);
            }
            else if (visible.size() == browse_.columns().items.size() - 3)
            {
                visible.insert(visible.begin() + kValueColDefault, true);
                visible.push_back(true);
                visible.push_back(true);
            }
            else if (visible.size() == browse_.columns().items.size() - 4)
            {
                visible.insert(visible.begin() + kValueColDefault, true);
                visible.insert(visible.begin() + kValueColReadOnBoot, true);
                visible.push_back(true);
                visible.push_back(true);
            }
        };
        patch_widths(browse_.columns().saved_widths);
        patch_visible(browse_.columns().saved_visible);
        for (size_t i = 0; i < browse_.columns().items.size(); ++i)
        {
            if (i < browse_.columns().saved_widths.size() && browse_.columns().saved_widths[i] > 0)
            {
                browse_.columns().widths[i] = browse_.columns().saved_widths[i];
            }
            if (i < browse_.columns().saved_visible.size())
            {
                browse_.columns().visible[i] = browse_.columns().saved_visible[i];
            }
        }
    }
    ApplyValueColumns();
}

void MainWindow::Impl::CreateHistoryColumns()
{
    history_columns_.Reset({
        {util::Tr(L"Time"), 140, LVCFMT_LEFT},
        {util::Tr(L"Action"), 280, LVCFMT_LEFT},
        {util::Tr(L"Old Data"), 220, LVCFMT_LEFT},
        {util::Tr(L"New Data"), 220, LVCFMT_LEFT},
    });
    ApplyHistoryColumns();
}

namespace
{

// hidden and unavailable columns keep subitems at width 0, unavailable columns also stay out of the header menu
void ApplyListColumns(HWND list, const ui::ColumnSet& set, const std::function<bool(size_t)>& unavailable = nullptr)
{
    std::vector<win32ui::ListColumn> columns(set.items.size());
    for (size_t i = 0; i < set.items.size(); ++i)
    {
        // saved widths use DIPs, 0 restores factory width used by Reset columns
        columns[i].title = set.items[i].title;
        columns[i].format = set.items[i].fmt;
        columns[i].width_dip = i < set.widths.size() && set.widths[i] > 0 ? set.widths[i] : set.items[i].width;
        columns[i].default_width_dip = set.items[i].width;
        columns[i].visible = i >= set.visible.size() || set.visible[i];
        columns[i].available = !unavailable || !unavailable(i);
    }
    const win32ui::RedrawScope list_redraw(list, RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
    const win32ui::RedrawScope header_redraw(ListView_GetHeader(list), 0);
    win32ui::SetListColumns(list, columns);
}

}

void MainWindow::Impl::ApplyValueColumns()
{
    HWND list = browse_.values().hwnd();
    if (!list)
    {
        return;
    }
    ApplyListColumns(list, browse_.columns(), [this](size_t index) { return settings_.simple_mode && index > kValueColData; });
    win32ui::UpdateListViewSort(list, browse_.columns().sort_column, browse_.columns().sort_ascending);
}

void MainWindow::Impl::ApplyHistoryColumns()
{
    if (!history_list_)
    {
        return;
    }
    ApplyListColumns(history_list_, history_columns_);
    win32ui::UpdateListViewSort(history_list_, history_sort_column_, history_sort_ascending_);
}

void MainWindow::Impl::CreateSearchColumns()
{
    if (!search_results_list_)
    {
        return;
    }
    search_columns_.Reset({
        {util::Tr(L"Path"), 320, LVCFMT_LEFT},
        {util::Tr(L"Value"), 180, LVCFMT_LEFT},
        {util::Tr(L"Type"), 110, LVCFMT_LEFT},
        {util::Tr(L"Data"), 360, LVCFMT_LEFT},
        {util::Tr(L"Size"), 80, LVCFMT_RIGHT},
        {util::Tr(L"Date Modified"), 150, LVCFMT_LEFT},
        {util::Tr(L"Source"), 190, LVCFMT_LEFT},
    });
    compare_columns_.Reset({
        {util::Tr(L"Path"), 320, LVCFMT_LEFT},
        {util::Tr(L"Value"), 180, LVCFMT_LEFT},
        {util::Tr(L"First Entry"), 320, LVCFMT_LEFT},
        {util::Tr(L"Second Entry"), 320, LVCFMT_LEFT},
        {util::Tr(L"Result"), 90, LVCFMT_LEFT},
    });
    compare_column_titles_ = {compare_columns_.items[2].title, compare_columns_.items[3].title};
    ApplySearchColumns(false);
}

void MainWindow::Impl::RefreshCompareColumnTitles()
{
    if (compare_columns_.items.size() < 4 || compare_column_titles_.size() < 2)
    {
        return;
    }
    const int index = tab_ ? SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) : -1;
    const SearchTab* tab = index >= 0 && static_cast<size_t>(index) < search_tabs_.size()
                               ? &search_tabs_[static_cast<size_t>(index)]
                               : nullptr;
    for (size_t side = 0; side < 2; ++side)
    {
        std::wstring title = compare_column_titles_[side];
        if (tab && tab->is_compare && side < tab->sources.size())
        {
            title += L" (" + search::SourceLabel(tab->sources[side]) + L")";
        }
        compare_columns_.items[side + 2].title = std::move(title);
    }
}

void MainWindow::Impl::ApplySearchColumns(bool compare)
{
    if (!search_results_list_)
    {
        return;
    }
    if (compare)
    {
        RefreshCompareColumnTitles();
    }
    const bool result_available = compare && IsCompareResultColumnAvailable();
    ApplyListColumns(search_results_list_, compare ? compare_columns_ : search_columns_, [&](size_t index) { return compare && index == 4 && !result_available; });
    compare_columns_active_ = compare;
    compare_result_column_active_ = result_available;
}

void MainWindow::Impl::WatchCurrentKey()
{
    const RegistryNode* node = browse_.current_node();
    const bool live = node && session_->mode == RegistryMode::kLocal && !node->simulated && !RegistryStore::IsVirtualRoot(node->root) &&
                      !RegistryStore::IsOfflineRoot(node->root) && node->root != HKEY_PERFORMANCE_DATA && node->root != HKEY_PERFORMANCE_TEXT &&
                      node->root != HKEY_PERFORMANCE_NLSTEXT;
    if (settings_.auto_refresh && live)
    {
        key_watcher_.Watch(hwnd_, frame::message_id::kRegistryChanged, node->root, node->subkey, ViewOf(*node));
    }
    else
    {
        key_watcher_.Stop();
    }
}

void MainWindow::Impl::ApplyAutoRefresh()
{
    KillTimer(hwnd_, kAutoRefreshTimerId);
    RegistryNode* node = browse_.current_node();
    if (!node)
    {
        return;
    }
    HWND tree = browse_.tree().hwnd();
    if (TreeView_GetEditControl(tree) || ListView_GetEditControl(browse_.values().hwnd()) || !pending_value_list_name_.empty())
    {
        SetTimer(hwnd_, kAutoRefreshTimerId, 250, nullptr);
        return;
    }
    key_watcher_.Rearm();
    auto_refresh_tick_ = GetTickCount64();
    // rebuilding tree items collapses children, rebuild only when subkey names change
    HTREEITEM item = TreeView_GetSelection(tree);
    if (item && (TreeView_GetItemState(tree, item, TVIS_EXPANDED) & TVIS_EXPANDED))
    {
        std::vector<std::wstring> shown;
        for (HTREEITEM child = TreeView_GetChild(tree, item); child; child = TreeView_GetNextSibling(tree, child))
        {
            const RegistryNode* entry = browse_.tree().NodeFromItem(child);
            if (entry && !entry->simulated)
            {
                shown.push_back(registry_path::Leaf(entry->subkey));
            }
        }
        std::vector<std::wstring> current = RegistryStore::EnumSubKeyNames(*node, false);
        const auto order = [](const std::wstring& left, const std::wstring& right) { return util::CompareInsensitive(left, right) < 0; };
        std::sort(shown.begin(), shown.end(), order);
        std::sort(current.begin(), current.end(), order);
        if (shown != current)
        {
            RefreshTreeSelection();
        }
    }
    merge_value_list_ = true;
    UpdateValueListForNode(browse_.current_node());
}

void MainWindow::Impl::UpdateValueListForNode(RegistryNode* node)
{
    if (updating_value_list_)
    {
        return;
    }
    const bool merge = merge_value_list_ && node;
    merge_value_list_ = false;
    WatchCurrentKey();
    updating_value_list_ = true;
    appended_value_name_.clear();
    retained_value_name_.clear();
    retained_value_key_path_.clear();
    retained_value_index_ = -1;
    if (browse_.current_node())
    {
        int selected = ListView_GetNextItem(browse_.values().hwnd(), -1, LVNI_SELECTED);
        if (selected >= 0 && ListView_GetNextItem(browse_.values().hwnd(), selected, LVNI_SELECTED) < 0)
        {
            if (const ListRow* row = browse_.values().RowAt(selected))
            {
                retained_value_index_ = selected;
                retained_value_key_path_ = registry_path::Build(*browse_.current_node());
                if (row->kind == rowkind::kValue)
                {
                    retained_value_name_ = row->extra;
                }
            }
        }
    }
    uint64_t generation = value_list_generation_.fetch_add(1) + 1;
    merge_value_list_generation_ = merge ? generation : 0;
    HWND list_hwnd = browse_.values().hwnd();
    if (!merge)
    {
        if (list_hwnd)
        {
            SendMessageW(list_hwnd, WM_SETREDRAW, FALSE, 0);
        }
        browse_.values().Clear();
        current_key_count_ = 0;
        current_value_count_ = 0;
    }
    if (!node)
    {
        if (list_hwnd && !jump_ui_batch_active_)
        {
            SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
            RedrawWindow(list_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
        }
        UpdateStatus();
        updating_value_list_ = false;
        value_list_loading_ = false;
        return;
    }

    RegistryNode snapshot = *node;
    RecordNavigation(registry_path::Build(snapshot));
    const std::wstring lookup_path = SourceLookupPath(snapshot);
    std::wstring trace_path = trace::NormalizeKeyPath(lookup_path);
    if (trace_path.empty())
    {
        trace_path = lookup_path;
    }
    std::wstring trace_path_lower = ToLower(trace_path);
    std::wstring default_path = trace::NormalizeKeyPathBasic(lookup_path);
    if (default_path.empty())
    {
        default_path = lookup_path;
    }
    std::wstring default_path_lower = ToLower(default_path);
    bool is_reg_file = IsRegFileTabSelected();
    auto trace_data_list = is_reg_file ? std::vector<ActiveTrace>() : active_traces_;
    auto default_data_list = is_reg_file ? std::vector<ActiveDefault>() : active_defaults_;
    bool show_simulated_keys = settings_.show_simulated_keys && !is_reg_file;
    constexpr size_t kDateColumn = static_cast<size_t>(kValueColDate);
    bool include_dates = browse_.columns().SortsBy(static_cast<int>(kDateColumn));
    if (kDateColumn < browse_.columns().visible.size() && browse_.columns().visible[kDateColumn])
    {
        include_dates = true;
    }
    constexpr size_t kDetailsColumn = static_cast<size_t>(kValueColDetails);
    bool include_details = browse_.columns().SortsBy(static_cast<int>(kDetailsColumn));
    if (kDetailsColumn < browse_.columns().visible.size() && browse_.columns().visible[kDetailsColumn])
    {
        include_details = true;
    }

    if (!merge)
    {
        if (list_hwnd && !jump_ui_batch_active_)
        {
            SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
            RedrawWindow(list_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
        }
        value_list_error_ = ERROR_SUCCESS;
        UpdateStatus();
    }
    updating_value_list_ = false;
    value_list_loading_ = true;

    int sort_column = browse_.columns().sort_column;
    bool sort_ascending = browse_.columns().sort_ascending;
    bool show_keys_in_list = settings_.show_keys_in_list;
    if (show_keys_in_list)
    {
        EnsureHiveListLoaded();
    }
    if (!value_loader_.running())
    {
        StartValueListWorker();
    }
    auto task = std::make_unique<ValueListTask>();
    task->generation = generation;
    task->snapshot = snapshot;
    task->trace_path_lower = std::move(trace_path_lower);
    task->default_path_lower = std::move(default_path_lower);
    task->include_dates = include_dates;
    task->sort_column = sort_column;
    task->sort_ascending = sort_ascending;
    task->sort_earlier = browse_.columns().sort_earlier;
    task->show_keys_in_list = show_keys_in_list;
    task->include_details = include_details;
    task->show_simulated_keys = show_simulated_keys;
    task->include_all_value_data = browse_.columns().SortsBy(kValueColData) || browse_.values().HasFilter();
    task->hwnd = hwnd_;
    task->trace_data_list = std::move(trace_data_list);
    task->default_data_list = std::move(default_data_list);
    task->hive_roots = show_keys_in_list ? hive_roots_ : nullptr;
    bool accepted = false;
    value_loader_.Submit(std::move(task), &accepted);
    // worker couldnt start, shown as unreadable instead of empty, the next selection or F5 tries again
    if (!accepted)
    {
        value_list_loading_ = false;
        value_list_error_ = ERROR_NOT_ENOUGH_MEMORY;
        UpdateStatus();
    }
}

void MainWindow::Impl::ScheduleValueListRename(LPARAM kind, const std::wstring& name)
{
    pending_value_list_kind_ = kind;
    pending_value_list_name_ = name;
}

void MainWindow::Impl::StartPendingValueListRename()
{
    if (pending_value_list_name_.empty() || !browse_.values().hwnd())
    {
        return;
    }
    if (pending_value_list_kind_ == rowkind::kKey && !settings_.show_keys_in_list)
    {
        pending_value_list_kind_ = 0;
        pending_value_list_name_.clear();
        return;
    }
    int index = -1;
    for (size_t i = 0; i < browse_.values().RowCount(); ++i)
    {
        const ListRow* row = browse_.values().RowAt(static_cast<int>(i));
        if (!row || row->kind != pending_value_list_kind_)
        {
            continue;
        }
        if (row->extra == pending_value_list_name_)
        {
            index = static_cast<int>(i);
            break;
        }
    }
    if (index >= 0 && IsWindowVisible(browse_.values().hwnd()))
    {
        SetFocus(browse_.values().hwnd());
        SelectListRowAtIndex(browse_.values().hwnd(), index);
        ListView_EditLabel(browse_.values().hwnd(), index);
    }
    pending_value_list_kind_ = 0;
    pending_value_list_name_.clear();
}

void MainWindow::Impl::ResetValueFilter()
{
    HWND filter = browse_.filter();
    if (!filter || GetWindowTextLengthW(filter) == 0)
    {
        return;
    }
    SetWindowTextW(filter, L"");
    browse_.values().SetFilter(std::wstring());
}

void MainWindow::Impl::FocusFirstValue()
{
    HWND list = browse_.values().hwnd();
    if (!list || ListView_GetItemCount(list) <= 0)
    {
        return;
    }
    SetFocus(list);
    SelectListRowAtIndex(list, 0);
}

void MainWindow::Impl::EnsureValueRowData(ListRow* row)
{
    if (!row || row->kind != rowkind::kValue || row->data_ready)
    {
        return;
    }
    if (row->value_data_size == 0)
    {
        row->data.clear();
        row->data_ready = true;
        browse_.values().InvalidateFilterCache(row);
        return;
    }
    if (!browse_.current_node())
    {
        return;
    }

    RegistryValue entry;
    if (!RegistryStore::QueryValue(*browse_.current_node(), row->extra, &entry))
    {
        row->data.clear();
        row->data_ready = true;
        browse_.values().InvalidateFilterCache(row);
        return;
    }

    row->value_type = entry.type;
    row->value_data_size = static_cast<DWORD>(entry.data.size());
    row->data = value_format::DisplayData(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
    row->data_ready = true;
    if (row->type.empty())
    {
        row->type = value_format::TypeName(entry.type);
    }
    row->size_value = row->value_data_size;
    row->has_size = true;
    if (row->size.empty() && row->value_data_size > 0)
    {
        row->size = std::to_wstring(row->value_data_size);
    }
    browse_.values().InvalidateFilterCache(row);
}

void MainWindow::Impl::UpdateAddressBar(RegistryNode* node)
{
    HWND address = browse_.address();
    if (!address)
    {
        return;
    }
    const std::wstring path = node ? registry_path::Build(*node) : std::wstring();
    if (!win32ui::SetTextIfChanged(address, path.c_str()))
    {
        return;
    }
    const int end = static_cast<int>(path.size());
    SendMessageW(address, EM_SETSEL, end, end);
    SendMessageW(address, EM_SCROLLCARET, 0, 0);
    UpdateGoButtonState();
}

void MainWindow::Impl::UpdateGoButtonState()
{
    HWND go = browse_.go_button();
    if (go && win32ui::SetEnabled(go, GetWindowTextLengthW(browse_.address()) > 0))
    {
        InvalidateRect(go, nullptr, TRUE);
    }
}

}
