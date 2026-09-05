// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

namespace {

} // namespace

void MainWindow::Impl::BuildImageLists() {
  if (tree_images_) {
    ImageList_Destroy(tree_images_);
    tree_images_ = nullptr;
  }
  if (list_images_) {
    ImageList_Destroy(list_images_);
    list_images_ = nullptr;
  }

  UINT dpi = win32::DpiForWindow(hwnd_);

  const int base_icon_size = kToolbarIconSize;
  const int icon_size = util::ScaleForDpi(base_icon_size, dpi);
  tree_images_ = ImageList_Create(icon_size, icon_size, ILC_COLOR32, 4, 2);
  list_images_ = ImageList_Create(icon_size, icon_size, ILC_COLOR32, 6, 2);
  ImageList_SetBkColor(tree_images_, CLR_NONE);
  ImageList_SetBkColor(list_images_, CLR_NONE);

  auto add_icon = [&](HIMAGELIST list, const wchar_t* name, int light_id, int dark_id) {
    HICON icon = LoadThemeIcon(name, light_id, dark_id, base_icon_size, dpi);
    util::ImageListAddOrBlank(list, icon, icon_size);
    if (icon) {
      DestroyIcon(icon);
    }
  };

  add_icon(tree_images_, L"folder.ico", IDI_ICON_LIGHT_FOLDER, IDI_ICON_DARK_FOLDER);
  add_icon(tree_images_, L"symlink.ico", IDI_ICON_LIGHT_SYMLINK, IDI_ICON_DARK_SYMLINK);
  add_icon(tree_images_, L"database.ico", IDI_ICON_LIGHT_DATABASE, IDI_ICON_DARK_DATABASE);
  add_icon(tree_images_, L"folder-sim.ico", IDI_ICON_LIGHT_FOLDER_SIM, IDI_ICON_DARK_FOLDER_SIM);
  add_icon(list_images_, L"folder.ico", IDI_ICON_LIGHT_FOLDER, IDI_ICON_DARK_FOLDER);
  add_icon(list_images_, L"symlink.ico", IDI_ICON_LIGHT_SYMLINK, IDI_ICON_DARK_SYMLINK);
  add_icon(list_images_, L"database.ico", IDI_ICON_LIGHT_DATABASE, IDI_ICON_DARK_DATABASE);
  add_icon(list_images_, L"folder-sim.ico", IDI_ICON_LIGHT_FOLDER_SIM, IDI_ICON_DARK_FOLDER_SIM);
  add_icon(list_images_, L"text.ico", IDI_ICON_LIGHT_TEXT, IDI_ICON_DARK_TEXT);
  add_icon(list_images_, L"binary.ico", IDI_ICON_LIGHT_BINARY, IDI_ICON_DARK_BINARY);
}

void MainWindow::Impl::CreateValueColumns() {
  browse_.columns().items = {
      {L"Name", 260, LVCFMT_LEFT},
      {L"Type", 120, LVCFMT_LEFT},
      {L"Data", 160, LVCFMT_LEFT},
      {L"Default", 200, LVCFMT_LEFT},
      {L"Read on boot", 110, LVCFMT_LEFT},
      {L"Size", 70, LVCFMT_RIGHT},
      {L"Date Modified", 140, LVCFMT_LEFT},
      {L"Details", 160, LVCFMT_LEFT},
      {L"Comment", 220, LVCFMT_LEFT},
  };
  browse_.columns().widths.clear();
  browse_.columns().visible.clear();
  browse_.columns().widths.reserve(browse_.columns().items.size());
  browse_.columns().visible.reserve(browse_.columns().items.size());
  for (const auto& column : browse_.columns().items) {
    browse_.columns().widths.push_back(column.width);
    browse_.columns().visible.push_back(true);
  }
  if (browse_.columns().saved) {
    auto patch_widths = [&](std::vector<int>& widths) {
      if (widths.size() == browse_.columns().items.size() - 1) {
        widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
      } else if (widths.size() == browse_.columns().items.size() - 2) {
        widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
        widths.push_back(browse_.columns().items[kValueColComment].width);
      } else if (widths.size() == browse_.columns().items.size() - 3) {
        widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
        widths.push_back(browse_.columns().items[kValueColDetails].width);
        widths.push_back(browse_.columns().items[kValueColComment].width);
      } else if (widths.size() == browse_.columns().items.size() - 4) {
        widths.insert(widths.begin() + kValueColDefault, browse_.columns().items[kValueColDefault].width);
        widths.insert(widths.begin() + kValueColReadOnBoot, browse_.columns().items[kValueColReadOnBoot].width);
        widths.push_back(browse_.columns().items[kValueColDetails].width);
        widths.push_back(browse_.columns().items[kValueColComment].width);
      }
    };
    auto patch_visible = [&](std::vector<bool>& visible) {
      if (visible.size() == browse_.columns().items.size() - 1) {
        visible.insert(visible.begin() + kValueColDefault, true);
      } else if (visible.size() == browse_.columns().items.size() - 2) {
        visible.insert(visible.begin() + kValueColDefault, true);
        visible.push_back(true);
      } else if (visible.size() == browse_.columns().items.size() - 3) {
        visible.insert(visible.begin() + kValueColDefault, true);
        visible.push_back(true);
        visible.push_back(true);
      } else if (visible.size() == browse_.columns().items.size() - 4) {
        visible.insert(visible.begin() + kValueColDefault, true);
        visible.insert(visible.begin() + kValueColReadOnBoot, true);
        visible.push_back(true);
        visible.push_back(true);
      }
    };
    patch_widths(browse_.columns().saved_widths);
    patch_visible(browse_.columns().saved_visible);
    for (size_t i = 0; i < browse_.columns().items.size(); ++i) {
      if (i < browse_.columns().saved_widths.size() && browse_.columns().saved_widths[i] > 0) {
        browse_.columns().widths[i] = browse_.columns().saved_widths[i];
        browse_.columns().items[i].width = browse_.columns().saved_widths[i];
      }
      if (i < browse_.columns().saved_visible.size()) {
        browse_.columns().visible[i] = browse_.columns().saved_visible[i];
      }
    }
  }
  ApplyValueColumns();
}

void MainWindow::Impl::CreateHistoryColumns() {
  history_columns_ = {
      {L"Time", 140, LVCFMT_LEFT},
      {L"Action", 280, LVCFMT_LEFT},
      {L"Old Data", 220, LVCFMT_LEFT},
      {L"New Data", 220, LVCFMT_LEFT},
  };
  history_column_widths_.clear();
  history_column_visible_.clear();
  history_column_widths_.reserve(history_columns_.size());
  history_column_visible_.reserve(history_columns_.size());
  for (const auto& column : history_columns_) {
    history_column_widths_.push_back(column.width);
    history_column_visible_.push_back(true);
  }
  ApplyHistoryColumns();
}

void MainWindow::Impl::ApplyValueColumns() {
  HWND list = browse_.values().hwnd();
  if (!list) {
    return;
  }
  HWND header = ListView_GetHeader(list);
  SendMessageW(list, WM_SETREDRAW, FALSE, 0);
  if (header) {
    SendMessageW(header, WM_SETREDRAW, FALSE, 0);
  }

  int count = header ? Header_GetItemCount(header) : 0;
  for (int i = count - 1; i >= 0; --i) {
    ListView_DeleteColumn(list, i);
  }

  value_column_subitems_.clear();
  int insert_index = 0;
  for (size_t i = 0; i < browse_.columns().items.size(); ++i) {
    if (i < browse_.columns().visible.size() && !browse_.columns().visible[i]) {
      continue;
    }
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT | LVCF_SUBITEM;
    col.pszText = const_cast<wchar_t*>(browse_.columns().items[i].title.c_str());
    int width = browse_.columns().widths[i];
    if (width <= 0) {
      width = browse_.columns().items[i].width;
    }
    col.cx = width;
    col.fmt = browse_.columns().items[i].fmt;
    col.iSubItem = static_cast<int>(i);
    ListView_InsertColumn(list, insert_index++, &col);
    value_column_subitems_.push_back(static_cast<int>(i));
  }

  header = ListView_GetHeader(list);
  if (header) {
    int size_display = FindListViewColumnBySubItem(list, kValueColSize);
    if (size_display >= 0) {
      HDITEMW item = {};
      item.mask = HDI_FORMAT;
      if (Header_GetItem(header, size_display, &item)) {
        item.fmt |= HDF_RIGHT;
        Header_SetItem(header, size_display, &item);
      }
    }
  }
  UpdateListViewSort(list, browse_.columns().sort_column, browse_.columns().sort_ascending);
  if (header) {
    AttachHeader(header);
    EnsureValueGridToolbar();
    SendMessageW(header, WM_SETREDRAW, TRUE, 0);
  }
  SendMessageW(list, WM_SETREDRAW, TRUE, 0);
  RedrawWindow(list, nullptr, nullptr,
               RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
}

void MainWindow::Impl::ApplyHistoryColumns() {
  if (!history_list_) {
    return;
  }
  HWND header = ListView_GetHeader(history_list_);
  SendMessageW(history_list_, WM_SETREDRAW, FALSE, 0);
  if (header) {
    SendMessageW(header, WM_SETREDRAW, FALSE, 0);
  }

  int count = header ? Header_GetItemCount(header) : 0;
  for (int i = count - 1; i >= 0; --i) {
    ListView_DeleteColumn(history_list_, i);
  }

  int insert_index = 0;
  for (size_t i = 0; i < history_columns_.size(); ++i) {
    if (i < history_column_visible_.size() && !history_column_visible_[i]) {
      continue;
    }
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT | LVCF_SUBITEM;
    col.pszText = const_cast<wchar_t*>(history_columns_[i].title.c_str());
    int width = history_column_widths_[i];
    if (width <= 0) {
      width = history_columns_[i].width;
    }
    col.cx = width;
    col.fmt = history_columns_[i].fmt;
    col.iSubItem = static_cast<int>(i);
    ListView_InsertColumn(history_list_, insert_index++, &col);
  }

  UpdateListViewSort(history_list_, history_sort_column_, history_sort_ascending_);
  header = ListView_GetHeader(history_list_);
  if (header) {
    AttachHeader(header);
    SendMessageW(header, WM_SETREDRAW, TRUE, 0);
  }
  SendMessageW(history_list_, WM_SETREDRAW, TRUE, 0);
  RedrawWindow(history_list_, nullptr, nullptr,
               RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
}

void MainWindow::Impl::CreateSearchColumns() {
  if (!search_results_list_) {
    return;
  }
  search_columns_ = {
      {L"Path", 320, LVCFMT_LEFT},
      {L"Value", 180, LVCFMT_LEFT},
      {L"Type", 110, LVCFMT_LEFT},
      {L"Data", 360, LVCFMT_LEFT},
      {L"Size", 80, LVCFMT_RIGHT},
      {L"Data Modified", 150, LVCFMT_LEFT},
  };
  search_column_widths_.clear();
  search_column_visible_.clear();
  search_column_widths_.reserve(search_columns_.size());
  search_column_visible_.reserve(search_columns_.size());
  for (const auto& column : search_columns_) {
    search_column_widths_.push_back(column.width);
    search_column_visible_.push_back(true);
  }
  compare_columns_ = {
      {L"Path", 320, LVCFMT_LEFT},
      {L"Value", 180, LVCFMT_LEFT},
      {L"First Entry", 320, LVCFMT_LEFT},
      {L"Second Entry", 320, LVCFMT_LEFT},
  };
  compare_column_widths_.clear();
  compare_column_visible_.clear();
  compare_column_widths_.reserve(compare_columns_.size());
  compare_column_visible_.reserve(compare_columns_.size());
  for (const auto& column : compare_columns_) {
    compare_column_widths_.push_back(column.width);
    compare_column_visible_.push_back(true);
  }
  ApplySearchColumns(false);
  HWND header = ListView_GetHeader(search_results_list_);
  AttachHeader(header);
}

void MainWindow::Impl::ApplySearchColumns(bool compare) {
  if (!search_results_list_) {
    return;
  }
  const auto& columns = compare ? compare_columns_ : search_columns_;
  auto& widths = compare ? compare_column_widths_ : search_column_widths_;
  auto& visible = compare ? compare_column_visible_ : search_column_visible_;
  HWND header = ListView_GetHeader(search_results_list_);
  SendMessageW(search_results_list_, WM_SETREDRAW, FALSE, 0);
  if (header) {
    SendMessageW(header, WM_SETREDRAW, FALSE, 0);
  }

  int count = header ? Header_GetItemCount(header) : 0;
  for (int i = count - 1; i >= 0; --i) {
    ListView_DeleteColumn(search_results_list_, i);
  }

  search_column_subitems_.clear();
  int insert_index = 0;
  for (size_t i = 0; i < columns.size(); ++i) {
    if (i < visible.size() && !visible[i]) {
      continue;
    }
    search_column_subitems_.push_back(static_cast<int>(i));
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT | LVCF_SUBITEM;
    col.pszText = const_cast<wchar_t*>(columns[i].title.c_str());
    int width = widths[i];
    if (width <= 0) {
      width = columns[i].width;
    }
    col.cx = width;
    col.fmt = columns[i].fmt;
    col.iSubItem = static_cast<int>(i);
    ListView_InsertColumn(search_results_list_, insert_index++, &col);
  }

  header = ListView_GetHeader(search_results_list_);
  if (header) {
    AttachHeader(header);
    SendMessageW(header, WM_SETREDRAW, TRUE, 0);
  }
  SendMessageW(search_results_list_, WM_SETREDRAW, TRUE, 0);
  RedrawWindow(search_results_list_, nullptr, nullptr,
               RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
  compare_columns_active_ = compare;
}

void MainWindow::Impl::UpdateValueListForNode(RegistryNode* node) {
  if (updating_value_list_) {
    return;
  }
  updating_value_list_ = true;
  appended_value_name_.clear();
  retained_value_name_.clear();
  retained_value_key_path_.clear();
  if (browse_.current_node()) {
    int selected = ListView_GetNextItem(browse_.values().hwnd(), -1, LVNI_SELECTED);
    if (selected >= 0 && ListView_GetNextItem(browse_.values().hwnd(), selected, LVNI_SELECTED) < 0) {
      if (const ListRow* row = browse_.values().RowAt(selected)) {
        if (row->kind == rowkind::kValue) {
          retained_value_name_ = row->extra;
          retained_value_key_path_ = registry_path::Build(*browse_.current_node());
        }
      }
    }
  }
  uint64_t generation = value_list_generation_.fetch_add(1) + 1;
  HWND list_hwnd = browse_.values().hwnd();
  if (list_hwnd) {
    SendMessageW(list_hwnd, WM_SETREDRAW, FALSE, 0);
  }

  browse_.values().Clear();
  current_key_count_ = 0;
  current_value_count_ = 0;
  if (!node) {
    if (list_hwnd && !jump_ui_batch_active_) {
      SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
      RedrawWindow(list_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
    }
    UpdateStatus();
    updating_value_list_ = false;
    value_list_loading_ = false;
    return;
  }

  RegistryNode snapshot = *node;
  std::wstring path = registry_path::Build(snapshot);
  RecordNavigation(path);
  std::wstring trace_path = NormalizeTraceKeyPath(path);
  if (trace_path.empty()) {
    trace_path = path;
  }
  std::wstring trace_path_lower = ToLower(trace_path);
  std::wstring default_path = NormalizeTraceKeyPathBasic(path);
  if (default_path.empty()) {
    default_path = path;
  }
  std::wstring default_path_lower = ToLower(default_path);
  bool is_reg_file = IsRegFileTabSelected();
  auto trace_data_list = is_reg_file ? std::vector<ActiveTrace>() : active_traces_;
  auto default_data_list = is_reg_file ? std::vector<ActiveDefault>() : active_defaults_;
  bool show_simulated_keys = show_simulated_keys_ && !is_reg_file;
  constexpr size_t kDateColumn = static_cast<size_t>(kValueColDate);
  bool include_dates = (browse_.columns().sort_column == static_cast<int>(kDateColumn));
  if (kDateColumn < browse_.columns().visible.size() && browse_.columns().visible[kDateColumn]) {
    include_dates = true;
  }
  constexpr size_t kDetailsColumn = static_cast<size_t>(kValueColDetails);
  bool include_details = (browse_.columns().sort_column == static_cast<int>(kDetailsColumn));
  if (kDetailsColumn < browse_.columns().visible.size() && browse_.columns().visible[kDetailsColumn]) {
    include_details = true;
  }

  if (list_hwnd && !jump_ui_batch_active_) {
    SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
    RedrawWindow(list_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
  }
  UpdateStatus();
  updating_value_list_ = false;
  value_list_loading_ = true;

  int sort_column = browse_.columns().sort_column;
  bool sort_ascending = browse_.columns().sort_ascending;
  bool show_keys_in_list = show_keys_in_list_;
  if (show_keys_in_list) {
    EnsureHiveListLoaded();
  }
  if (!value_loader_.running()) {
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
  task->show_keys_in_list = show_keys_in_list;
  task->include_details = include_details;
  task->show_simulated_keys = show_simulated_keys;
  task->include_all_value_data = sort_column == kValueColData || browse_.values().HasFilter();
  task->hwnd = hwnd_;
  task->trace_data_list = std::move(trace_data_list);
  task->default_data_list = std::move(default_data_list);
  task->hive_roots = show_keys_in_list ? hive_roots_ : nullptr;
  value_loader_.Submit(std::move(task));
}

void MainWindow::Impl::ScheduleValueListRename(LPARAM kind, const std::wstring& name) {
  pending_value_list_kind_ = kind;
  pending_value_list_name_ = name;
}

void MainWindow::Impl::StartPendingValueListRename() {
  if (pending_value_list_name_.empty() || !browse_.values().hwnd()) {
    return;
  }
  if (pending_value_list_kind_ == rowkind::kKey && !show_keys_in_list_) {
    pending_value_list_kind_ = 0;
    pending_value_list_name_.clear();
    return;
  }
  int index = -1;
  for (size_t i = 0; i < browse_.values().RowCount(); ++i) {
    const ListRow* row = browse_.values().RowAt(static_cast<int>(i));
    if (!row || row->kind != pending_value_list_kind_) {
      continue;
    }
    if (row->extra == pending_value_list_name_) {
      index = static_cast<int>(i);
      break;
    }
  }
  if (index >= 0 && IsWindowVisible(browse_.values().hwnd())) {
    SetFocus(browse_.values().hwnd());
    ListView_SetItemState(browse_.values().hwnd(), -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetItemState(browse_.values().hwnd(), index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(browse_.values().hwnd(), index, FALSE);
    ListView_EditLabel(browse_.values().hwnd(), index);
  }
  pending_value_list_kind_ = 0;
  pending_value_list_name_.clear();
}

void MainWindow::Impl::AttachBorder(HWND control) {
  if (!control || GetWindowSubclass(control, BorderProc, kBorderSubclassId, nullptr)) {
    return;
  }
  if (!SetWindowSubclass(control, BorderProc, kBorderSubclassId,
                         reinterpret_cast<DWORD_PTR>(this))) {
    return;
  }
  const LONG_PTR style = GetWindowLongPtrW(control, GWL_STYLE);
  SetWindowLongPtrW(control, GWL_STYLE, style | WS_BORDER);
  SetWindowPos(control, nullptr, 0, 0, 0, 0,
               SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

void MainWindow::Impl::AttachHeader(HWND header) {
  if (!header || GetWindowSubclass(header, HeaderProc, kHeaderSubclassId, nullptr)) {
    return;
  }
  SetWindowSubclass(header, HeaderProc, kHeaderSubclassId, reinterpret_cast<DWORD_PTR>(this));
}

void MainWindow::Impl::ResetValueFilter() {
  HWND filter = browse_.filter();
  if (!filter || GetWindowTextLengthW(filter) == 0) {
    return;
  }
  SetWindowTextW(filter, L"");
  browse_.values().SetFilter(std::wstring());
}

void MainWindow::Impl::FocusFirstValue() {
  HWND list = browse_.values().hwnd();
  if (!list || ListView_GetItemCount(list) <= 0) {
    return;
  }
  SetFocus(list);
  ListView_SetItemState(list, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
  ListView_SetItemState(list, 0, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
  ListView_EnsureVisible(list, 0, FALSE);
}

void MainWindow::Impl::EnsureValueRowData(ListRow* row) {
  if (!row || row->kind != rowkind::kValue || row->data_ready) {
    return;
  }
  if (row->value_data_size == 0) {
    row->data.clear();
    row->data_ready = true;
    browse_.values().InvalidateFilterCache(row);
    return;
  }
  if (!browse_.current_node()) {
    return;
  }

  ValueEntry entry;
  if (!RegistryStore::QueryValue(*browse_.current_node(), row->extra, &entry)) {
    row->data.clear();
    row->data_ready = true;
    browse_.values().InvalidateFilterCache(row);
    return;
  }

  row->value_type = entry.type;
  row->value_data_size = static_cast<DWORD>(entry.data.size());
  row->data = value_format::DisplayData(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
  row->data_ready = true;
  if (row->type.empty()) {
    row->type = value_format::TypeName(entry.type);
  }
  row->size_value = row->value_data_size;
  row->has_size = true;
  if (row->size.empty() && row->value_data_size > 0) {
    row->size = std::to_wstring(row->value_data_size);
  }
  browse_.values().InvalidateFilterCache(row);
}

void MainWindow::Impl::UpdateAddressBar(RegistryNode* node) {
  HWND address = browse_.address();
  if (!address) {
    return;
  }
  const std::wstring path = node ? registry_path::Build(*node) : std::wstring();
  if (static_cast<size_t>(GetWindowTextLengthW(address)) == path.size()) {
    std::wstring current(path.size() + 1, L'\0');
    GetWindowTextW(address, current.data(), static_cast<int>(current.size()));
    current.resize(path.size());
    if (current == path) {
      return;
    }
  }
  SetWindowTextW(address, path.c_str());
  UpdateGoButtonState();
}

void MainWindow::Impl::UpdateGoButtonState() {
  HWND go = browse_.go_button();
  if (!go) {
    return;
  }
  const bool enabled = GetWindowTextLengthW(browse_.address()) > 0;
  if ((IsWindowEnabled(go) != FALSE) == enabled) {
    return;
  }
  EnableWindow(go, enabled);
  InvalidateRect(go, nullptr, TRUE);
}

} // namespace regkit
