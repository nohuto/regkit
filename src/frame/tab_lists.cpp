// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

#include <unordered_set>

namespace regkit {
using namespace window_detail;

namespace {

struct StableListSelection {
  std::unordered_set<std::wstring> selected;
  std::wstring focused;
};

void AppendIdentityPart(std::wstring* key, const std::wstring& part) {
  key->push_back(L'|');
  key->append(std::to_wstring(part.size()));
  key->push_back(L':');
  key->append(part);
}

std::wstring ValueRowIdentity(const ListRow& row) {
  std::wstring key = std::to_wstring(static_cast<long long>(row.kind));
  AppendIdentityPart(&key, row.extra);
  return key;
}

std::wstring SearchResultIdentity(const search::Result& result) {
  return std::to_wstring(result.row_id);
}

std::wstring HistoryEntryIdentity(const HistoryEntry& entry) {
  std::wstring key = std::to_wstring(entry.timestamp);
  AppendIdentityPart(&key, entry.action);
  AppendIdentityPart(&key, entry.key_path);
  AppendIdentityPart(&key, entry.value_name);
  return key;
}

template <typename KeyAt>
StableListSelection CaptureListSelection(HWND list, KeyAt key_at) {
  StableListSelection state;
  if (!list) {
    return state;
  }
  state.selected.reserve(static_cast<size_t>(ListView_GetSelectedCount(list)));
  int index = -1;
  while ((index = ListView_GetNextItem(list, index, LVNI_SELECTED)) >= 0) {
    std::wstring key = key_at(index);
    if (!key.empty()) {
      state.selected.emplace(std::move(key));
    }
  }
  index = ListView_GetNextItem(list, -1, LVNI_FOCUSED);
  if (index >= 0) {
    state.focused = key_at(index);
  }
  return state;
}

template <typename KeyAt>
void RestoreListSelection(HWND list, const StableListSelection& state,
                          KeyAt key_at) {
  if (!list) {
    return;
  }
  if (state.selected.empty() && state.focused.empty()) {
    return;
  }
  SendMessageW(list, WM_SETREDRAW, FALSE, 0);
  ListView_SetItemState(list, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
  int focused_index = -1;
  const int count = ListView_GetItemCount(list);
  for (int index = 0; index < count; ++index) {
    std::wstring key = key_at(index);
    UINT state_mask = 0;
    if (state.selected.find(key) != state.selected.end()) {
      state_mask |= LVIS_SELECTED;
    }
    if (!state.focused.empty() && key == state.focused) {
      state_mask |= LVIS_FOCUSED;
      focused_index = index;
    }
    if (state_mask != 0) {
      ListView_SetItemState(list, index, state_mask, state_mask);
    }
  }
  SendMessageW(list, WM_SETREDRAW, TRUE, 0);
  if (focused_index >= 0) {
    ListView_EnsureVisible(list, focused_index, FALSE);
  }
  RedrawWindow(list, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
}

} // namespace

void MainWindow::Impl::SortValueList(int column, bool toggle) {
  if (column < 0 || static_cast<size_t>(column) >= browse_.columns().items.size()) {
    return;
  }
  if (toggle) {
    if (browse_.columns().sort_column == column) {
      browse_.columns().sort_ascending = !browse_.columns().sort_ascending;
    } else {
      browse_.columns().sort_column = column;
      browse_.columns().sort_ascending = true;
    }
  } else {
    browse_.columns().sort_column = column;
  }

  if (value_list_loading_ && browse_.current_node()) {
    UpdateValueListForNode(browse_.current_node());
    return;
  }

  StableListSelection selection = CaptureListSelection(
      browse_.values().hwnd(), [this](int index) {
        const ListRow* row = browse_.values().RowAt(index);
        return row ? ValueRowIdentity(*row) : std::wstring();
      });
  auto& rows = browse_.values().rows();
  if (browse_.columns().sort_column == kValueColData) {
    bool needs_data = false;
    for (const auto& row : rows) {
      if (row.kind == rowkind::kValue && !row.data_ready) {
        needs_data = true;
        break;
      }
    }
    if (needs_data && browse_.current_node()) {
      UpdateValueListForNode(browse_.current_node());
      return;
    }
    for (auto& row : rows) {
      EnsureValueRowData(&row);
    }
  }
  const bool was_updating = updating_value_list_;
  updating_value_list_ = true;
  SortValueRows(&rows, browse_.columns().sort_column, browse_.columns().sort_ascending);
  browse_.values().RebuildFilter();
  RestoreListSelection(browse_.values().hwnd(), selection,
                       [this](int index) {
                         const ListRow* row = browse_.values().RowAt(index);
                         return row ? ValueRowIdentity(*row) : std::wstring();
                       });
  updating_value_list_ = was_updating;
  if (!was_updating) {
    UpdateStatus();
  }

  UpdateListViewSort(browse_.values().hwnd(), browse_.columns().sort_column, browse_.columns().sort_ascending);
}

void MainWindow::Impl::SortHistoryList(int column, bool toggle) {
  if (!history_list_ || column < 0) {
    return;
  }
  if (toggle) {
    if (history_sort_column_ == column) {
      history_sort_ascending_ = !history_sort_ascending_;
    } else {
      history_sort_column_ = column;
      history_sort_ascending_ = true;
    }
  } else {
    history_sort_column_ = column;
  }

  auto history_key_at = [this](int index) {
    const auto& entries = change_history_.entries();
    if (index < 0 || static_cast<size_t>(index) >= entries.size()) {
      return std::wstring();
    }
    return HistoryEntryIdentity(entries[static_cast<size_t>(index)]);
  };
  StableListSelection selection =
      CaptureListSelection(history_list_, history_key_at);
  change_history_.Sort(history_sort_column_, history_sort_ascending_);
  RebuildHistoryList();
  RestoreListSelection(history_list_, selection, history_key_at);

  UpdateListViewSort(history_list_, history_sort_column_, history_sort_ascending_);
}

void MainWindow::Impl::SortSearchTabResults(SearchTab* tab) {
  if (!tab || tab->sort_column < 0) {
    if (tab) {
      tab->sort_dirty = false;
    }
    return;
  }
  if (tab->is_compare) {
    search::compare::SortRows(&tab->compare_rows, tab->sort_column,
                              tab->sort_ascending);
    tab->sort_dirty = false;
    return;
  }


  if (tab->sort_column == 3) {
    bool unresolved = false;
    for (const auto& result : tab->results) {
      if (result.data_state == search::DataState::kNotLoaded) {
        unresolved = true;
        break;
      }
    }
    if (unresolved) {
      QueueSearchSort(tab);
      tab->sort_dirty = false;
      return;
    }
  }
  search::SortResults(&tab->results, tab->sort_column, tab->sort_ascending);
  tab->sort_dirty = false;
}

void MainWindow::Impl::SortSearchResults(int column, bool toggle) {
  if (!search_results_list_ || column < 0) {
    return;
  }
  int sel = TabCtrl_GetCurSel(tab_);
  int index = SearchIndexFromTab(sel);
  if (index < 0 || static_cast<size_t>(index) >= search_tabs_.size()) {
    return;
  }
  EnsureSearchTabResultsLoaded(index);
  auto& tab = search_tabs_[static_cast<size_t>(index)];
  auto search_key_at = [&tab](int row) {
    return row >= 0 && static_cast<size_t>(row) < tab.results.size()
               ? SearchResultIdentity(tab.results[static_cast<size_t>(row)])
               : std::wstring();
  };
  StableListSelection selection =
      CaptureListSelection(search_results_list_, search_key_at);
  if (toggle) {
    if (tab.sort_column == column) {
      tab.sort_ascending = !tab.sort_ascending;
    } else {
      tab.sort_column = column;
      tab.sort_ascending = true;
    }
  } else {
    tab.sort_column = column;
  }
  SortSearchTabResults(&tab);
  RestoreListSelection(search_results_list_, selection, search_key_at);
  UpdateListViewSort(search_results_list_, tab.sort_column, tab.sort_ascending);
  RedrawWindow(search_results_list_, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
}

void MainWindow::Impl::ClearHistoryItems(bool delete_cache) {
  if (!history_list_) {
    return;
  }
  change_history_.entries().clear();
  ListView_DeleteAllItems(history_list_);

  if (delete_cache) {
    std::wstring path = HistoryCachePath();
    if (!path.empty()) {
      DeleteFileW(path.c_str());
    }
  }
}

void MainWindow::Impl::RemoveSelectedHistoryItems() {
  if (!history_list_) {
    return;
  }
  auto& entries = change_history_.entries();
  std::vector<int> selected;
  for (int index = ListView_GetNextItem(history_list_, -1, LVNI_SELECTED);
       index >= 0;
       index = ListView_GetNextItem(history_list_, index, LVNI_SELECTED)) {
    if (static_cast<size_t>(index) < entries.size()) {
      selected.push_back(index);
    }
  }
  if (selected.empty()) {
    return;
  }
  for (auto it = selected.rbegin(); it != selected.rend(); ++it) {
    entries.erase(entries.begin() + *it);
  }
  ListView_SetItemState(history_list_, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
  RebuildHistoryList();
  changes::WriteHistoryFile(HistoryCachePath(), entries);
}

void MainWindow::Impl::RebuildHistoryList() {
  if (!history_list_) {
    return;
  }
  const int count = static_cast<int>(change_history_.entries().size());
  ListView_SetItemCountEx(history_list_, count,
                          LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
  RedrawWindow(history_list_, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
  if (history_sort_column_ == 0 && history_sort_ascending_ && count > 0) {
    ListView_EnsureVisible(history_list_, count - 1, FALSE);
  }
}

void MainWindow::Impl::ResetNavigationState() {
  browse_.ResetNavigation();
  UpdateNavigationButtons();
}

void MainWindow::Impl::UpdateTabText(const std::wstring& text) {
  if (!tab_) {
    return;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (IsSearchTabIndex(index) || IsRegFileTabIndex(index)) {
    index = FindFirstRegistryTabIndex();
  }
  if (index < 0) {
    return;
  }
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(text.c_str());
  TabCtrl_SetItem(tab_, index, &item);
  UpdateTabWidth();
  InvalidateRect(tab_, nullptr, FALSE);
}

void MainWindow::Impl::MarkOfflineDirty() {
  if (IsRegFileTabSelected()) {
    int index = TabCtrl_GetCurSel(tab_);
    if (index >= 0 && static_cast<size_t>(index) < tabs_.size() && IsRegFileTabIndex(index)) {
      bool was_dirty = tabs_[static_cast<size_t>(index)].reg_file_dirty;
      tabs_[static_cast<size_t>(index)].reg_file_dirty = true;
      if (!was_dirty) {
        BuildMenus();
      }
    }
    return;
  }
  if (registry_mode_ != RegistryMode::kOffline) {
    return;
  }
  int index = CurrentRegistryTabIndex();
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kOffline) {
    return;
  }
  if (!entry.offline_dirty) {
    entry.offline_dirty = true;
    BuildMenus();
  }
}

void MainWindow::Impl::ClearOfflineDirty() {
  if (registry_mode_ != RegistryMode::kOffline) {
    return;
  }
  int index = CurrentRegistryTabIndex();
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kOffline) {
    return;
  }
  if (entry.offline_dirty) {
    entry.offline_dirty = false;
    BuildMenus();
  }
}

bool MainWindow::Impl::ConfirmCloseTab(int tab_index) {
  if (!tab_ || tab_index < 0 || static_cast<size_t>(tab_index) >= tabs_.size()) {
    return false;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
  if (entry.kind == TabEntry::Kind::kRegFile && entry.reg_file_dirty) {
    std::wstring message = L"The registry file has unsaved changes.\nSave "
                           L"before closing the tab?";
    int result = ui::PromptChoice(hwnd_, message, L"Unsaved changes", L"Save", L"Don't Save", L"Cancel");
    if (result == IDCANCEL) {
      return false;
    }
    if (result == IDNO) {
      return true;
    }
    if (SaveRegFileTab(tab_index)) {
      entry.reg_file_dirty = false;
      return true;
    }
    return false;
  }
  if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kOffline || !entry.offline_dirty) {
    return true;
  }
  if (tab_index != CurrentRegistryTabIndex()) {
    return true;
  }
  std::wstring message = L"The offline registry has unsaved changes.\nSave "
                         L"before closing the tab?";
  int result = ui::PromptChoice(hwnd_, message, L"Unsaved changes", L"Save", L"Don't Save", L"Cancel");
  if (result == IDCANCEL) {
    return false;
  }
  if (result == IDNO) {
    return true;
  }
  if (SaveOfflineRegistry()) {
    entry.offline_dirty = false;
    return true;
  }
  return false;
}

void MainWindow::Impl::CloseTab(int tab_index) {
  if (!tab_) {
    return;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (count <= 1 || tab_index < 0 || tab_index >= count) {
    return;
  }
  if (IsSearchTabIndex(tab_index)) {
    CloseSearchTab(tab_index);
    return;
  }
  const int registry_tab_count =
      static_cast<int>(std::count_if(tabs_.begin(), tabs_.end(), [](const TabEntry& entry) {
        return entry.kind == TabEntry::Kind::kRegistry;
      }));
  if (registry_tab_count <= 1 &&
      static_cast<size_t>(tab_index) < tabs_.size() &&
      tabs_[static_cast<size_t>(tab_index)].kind == TabEntry::Kind::kRegistry) {
    return;
  }
  if (!ConfirmCloseTab(tab_index)) {
    return;
  }

  if (IsRegFileTabIndex(tab_index)) {
    TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
    if (entry.reg_file_loading && !entry.reg_file_path.empty()) {
      std::wstring lower = ToLower(entry.reg_file_path);
      auto it = reg_file_parse_sessions_.find(lower);
      if (it != reg_file_parse_sessions_.end() && it->second) {
        it->second->work.CancelAndJoin();
        reg_file_parse_sessions_.erase(it);
      }
    }
    ReleaseRegFileRoots(&entry);
  }
  tabs_.erase(tabs_.begin() + tab_index);
  TabCtrl_DeleteItem(tab_, tab_index);

  if (active_search_tab_index_ == tab_index) {
    active_search_tab_index_ = -1;
  } else if (active_search_tab_index_ > tab_index) {
    --active_search_tab_index_;
  }

  int new_count = TabCtrl_GetItemCount(tab_);
  if (new_count > 0) {
    int new_index = std::min(tab_index, new_count - 1);
    TabCtrl_SetCurSel(tab_, new_index);
    ApplyTabSelection(new_index);
  }
  RefreshRegistryTabLabels();
  ApplyViewVisibility();
  UpdateSearchResultsView();
  UpdateStatus();
}

void MainWindow::Impl::SelectTabIndex(int index) {
  if (!tab_) {
    return;
  }
  const int current = TabCtrl_GetCurSel(tab_);
  if (current != index) {
    CaptureRegistryTabState(current);
  }
  TabCtrl_SetCurSel(tab_, index);
}

void MainWindow::Impl::OpenLocalRegistryTab() {
  if (!tab_) {
    return;
  }
  CaptureRegistryTabState(TabCtrl_GetCurSel(tab_));
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(L"Local Registry");
  int index = TabCtrl_GetItemCount(tab_);
  TabCtrl_InsertItem(tab_, index, &item);
  TabEntry entry;
  entry.kind = TabEntry::Kind::kRegistry;
  entry.registry_mode = RegistryMode::kLocal;
  tabs_.push_back(std::move(entry));
  RefreshRegistryTabLabels();
  TabCtrl_SetCurSel(tab_, index);
  SwitchToLocalRegistry();
  RestoreRegistryTabState(index);
  ApplyViewVisibility();
  UpdateSearchResultsView();
  UpdateStatus();
}

int MainWindow::Impl::CurrentRegistryTabIndex() const {
  if (!tab_) {
    return -1;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (index < 0) {
    return -1;
  }
  if (!IsSearchTabIndex(index) && !IsRegFileTabIndex(index)) {
    return index;
  }
  return FindFirstRegistryTabIndex();
}

void MainWindow::Impl::UpdateRegistryTabEntry(RegistryMode mode, const std::wstring& offline_path, const std::wstring& remote_machine) {
  int index = CurrentRegistryTabIndex();
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry) {
    return;
  }
  entry.registry_mode = mode;
  entry.offline_path = offline_path;
  entry.remote_machine = remote_machine;
}

void MainWindow::Impl::UpdateTabWidth() {
  if (!tab_) {
    return;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (count <= 0) {
    return;
  }
  bool has_close = count > 1;
  int pad_x = kTabTextPaddingX + (has_close ? (kTabCloseSize + kTabCloseGap) : 0);
  int pad_y = kTabInsetY + 2;
  TabCtrl_SetPadding(tab_, pad_x, pad_y);
  int text_height = 0;
  HDC hdc = GetDC(tab_);
  HFONT font = reinterpret_cast<HFONT>(SendMessageW(tab_, WM_GETFONT, 0, 0));
  HFONT old_font = nullptr;
  if (hdc && font) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, font));
  }
  if (hdc) {
    TEXTMETRICW tm = {};
    if (GetTextMetricsW(hdc, &tm)) {
      text_height = tm.tmHeight;
    }
  }

  if (hdc) {
    if (old_font) {
      SelectObject(hdc, old_font);
    }
    ReleaseDC(tab_, hdc);
  }

  int min_height = std::max<int>(24, text_height + pad_y * 2 + 2);
  SendMessageW(tab_, TCM_SETMINTABWIDTH, 0, static_cast<LPARAM>(kTabMinWidth));
  RECT item_rect = {};
  if (TabCtrl_GetItemRect(tab_, 0, &item_rect)) {
    int item_height = static_cast<int>(item_rect.bottom - item_rect.top);
    tab_height_ = std::max<int>(min_height, item_height);
  } else {
    tab_height_ = min_height;
  }
  InvalidateRect(tab_, nullptr, FALSE);
  if (hwnd_) {
    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    if (rect.right > 0 && rect.bottom > 0) {
      LayoutControls(rect.right, rect.bottom);
    }
  }
}

void MainWindow::Impl::BuildAccelerators() {
  if (accelerators_) {
    DestroyAcceleratorTable(accelerators_);
    accelerators_ = nullptr;
  }
  ACCEL accels[] = {
      {FVIRTKEY | FCONTROL, 'C', cmd::kEditCopy},
      {FVIRTKEY | FCONTROL, 'V', cmd::kEditPaste},
      {FVIRTKEY | FCONTROL, 'A', cmd::kViewSelectAll},
      {FVIRTKEY | FCONTROL, 'Z', cmd::kEditUndo},
      {FVIRTKEY | FCONTROL, 'Y', cmd::kEditRedo},
      {FVIRTKEY | FCONTROL, 'F', cmd::kEditFind},
      {FVIRTKEY | FCONTROL, 'G', cmd::kEditGoTo},
      {FVIRTKEY | FCONTROL, 'H', cmd::kEditReplace},
      {FVIRTKEY | FCONTROL, 'S', cmd::kFileSave},
      {FVIRTKEY | FCONTROL, 'E', cmd::kFileExport},
      {FVIRTKEY | FCONTROL | FSHIFT, 'C', cmd::kEditCopyKey},
      {FVIRTKEY, VK_DELETE, cmd::kEditDelete},
      {FVIRTKEY, VK_F2, cmd::kEditRename},
      {FVIRTKEY, VK_F5, cmd::kViewRefresh},
      {FVIRTKEY | FALT, VK_LEFT, cmd::kNavBack},
      {FVIRTKEY | FALT, VK_RIGHT, cmd::kNavForward},
      {FVIRTKEY | FALT, VK_UP, cmd::kNavUp},
  };
  accelerators_ = CreateAcceleratorTableW(accels, static_cast<int>(sizeof(accels) / sizeof(accels[0])));
}

bool MainWindow::Impl::SelectAllInFocusedList() {
  HWND focus = GetFocus();
  if (!focus) {
    return false;
  }
  if (focus != browse_.values().hwnd() && focus != history_list_ && focus != search_results_list_) {
    return false;
  }
  int count = ListView_GetItemCount(focus);
  if (count <= 0) {
    return true;
  }
  const bool value_list = focus == browse_.values().hwnd();
  const bool was_updating = updating_value_list_;
  if (value_list) {
    updating_value_list_ = true;
  }
  ListView_SetItemState(focus, -1, LVIS_SELECTED, LVIS_SELECTED);
  ListView_SetItemState(focus, 0, LVIS_FOCUSED, LVIS_FOCUSED);
  ListView_EnsureVisible(focus, 0, FALSE);
  if (value_list) {
    updating_value_list_ = was_updating;
    if (!was_updating) {
      UpdateStatus();
    }
  }
  return true;
}

bool MainWindow::Impl::InvertSelectionInFocusedList() {
  HWND focus = GetFocus();
  if (!focus) {
    return false;
  }
  if (focus != browse_.values().hwnd() && focus != history_list_ && focus != search_results_list_) {
    return false;
  }
  int count = ListView_GetItemCount(focus);
  if (count <= 0) {
    return true;
  }
  const bool value_list = focus == browse_.values().hwnd();
  const bool was_updating = updating_value_list_;
  if (value_list) {
    updating_value_list_ = true;
  }
  SendMessageW(focus, WM_SETREDRAW, FALSE, 0);
  int first_selected = -1;
  for (int i = 0; i < count; ++i) {
    UINT state = ListView_GetItemState(focus, i, LVIS_SELECTED);
    if (state & LVIS_SELECTED) {
      ListView_SetItemState(focus, i, 0, LVIS_SELECTED);
    } else {
      ListView_SetItemState(focus, i, LVIS_SELECTED, LVIS_SELECTED);
      if (first_selected < 0) {
        first_selected = i;
      }
    }
  }
  if (first_selected < 0) {
    first_selected = 0;
  }
  ListView_SetItemState(focus, first_selected, LVIS_FOCUSED, LVIS_FOCUSED);
  ListView_EnsureVisible(focus, first_selected, FALSE);
  SendMessageW(focus, WM_SETREDRAW, TRUE, 0);
  InvalidateRect(focus, nullptr, TRUE);
  if (value_list) {
    updating_value_list_ = was_updating;
    if (!was_updating) {
      UpdateStatus();
    }
  }
  return true;
}

void MainWindow::Impl::UpdateTabHotState(HWND hwnd, POINT pt) {
  int new_hot = -1;
  int new_close_hot = -1;

  TCHITTESTINFO hit = {};
  hit.pt = pt;
  int index = TabCtrl_HitTest(hwnd, &hit);
  if (index >= 0) {
    new_hot = index;
    RECT close_rect = {};
    if (GetTabCloseRect(index, &close_rect) && PtInRect(&close_rect, pt)) {
      new_close_hot = index;
    }
  }

  if (new_hot != tab_hot_index_ || new_close_hot != tab_close_hot_index_) {
    tab_hot_index_ = new_hot;
    tab_close_hot_index_ = new_close_hot;
    InvalidateRect(hwnd, nullptr, FALSE);
  }
}

bool MainWindow::Impl::GetTabCloseRect(int index, RECT* rect) const {
  if (!tab_ || !rect || index < 0) {
    return false;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (count <= 1) {
    return false;
  }
  RECT item_rect = {};
  if (!TabCtrl_GetItemRect(tab_, index, &item_rect)) {
    return false;
  }
  int header_bottom = item_rect.bottom + 1;
  RECT draw_rect = AdjustTabDrawRect(item_rect, header_bottom, false);
  RECT close_area = draw_rect;
  close_area.left = item_rect.left;
  close_area.right = item_rect.right;
  return CalcTabCloseRect(close_area, rect);
}

void MainWindow::Impl::DrawTabItem(HDC hdc, int index, const RECT& item_rect, int header_bottom, bool selected) {
  const Theme& theme = Theme::Current();
  RECT draw_rect = AdjustTabDrawRect(item_rect, header_bottom, selected);

  bool is_hot = (index == tab_hot_index_);
  bool close_hot = (index == tab_close_hot_index_);
  bool close_down = (index == tab_close_down_index_);

  COLORREF fill = selected ? theme.SurfaceColor() : theme.PanelColor();
  if (is_hot) {
    fill = theme.HoverColor();
  }
  HBRUSH fill_brush = appearance::CachedBrush(fill);
  FillRect(hdc, &draw_rect, fill_brush);

  HPEN border_pen = appearance::CachedPen(theme.BorderColor(), 1);
  HGDIOBJ old_pen = SelectObject(hdc, border_pen);
  MoveToEx(hdc, draw_rect.left, draw_rect.bottom, nullptr);
  LineTo(hdc, draw_rect.left, draw_rect.top);
  LineTo(hdc, draw_rect.right, draw_rect.top);
  LineTo(hdc, draw_rect.right, draw_rect.bottom);
  if (!selected) {
    LineTo(hdc, draw_rect.left, draw_rect.bottom);
  }
  SelectObject(hdc, old_pen);

  RECT close_rect = {};
  RECT close_area = draw_rect;
  close_area.left = item_rect.left;
  close_area.right = item_rect.right;
  bool has_close = TabCtrl_GetItemCount(tab_) > 1 && CalcTabCloseRect(close_area, &close_rect);

  RECT text_rect = draw_rect;
  text_rect.left = item_rect.left + kTabTextPaddingX;
  text_rect.right = item_rect.right - kTabTextPaddingX;
  if (has_close) {
    text_rect.right = std::max(text_rect.left, close_rect.left - kTabCloseGap);
  }

  COLORREF text_color = selected || is_hot ? theme.TextColor() : theme.MutedTextColor();
  SetTextColor(hdc, text_color);
  SetBkMode(hdc, TRANSPARENT);

  wchar_t text[256] = {};
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = text;
  item.cchTextMax = static_cast<int>(_countof(text));
  if (TabCtrl_GetItem(tab_, index, &item)) {
    DrawTextW(hdc, text, -1, &text_rect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
  }

  if (has_close) {
    if (close_down) {
      HBRUSH down_brush = appearance::CachedBrush(theme.SelectionColor());
      FillRect(hdc, &close_rect, down_brush);
    } else if (close_hot) {
      HBRUSH hot_brush = appearance::CachedBrush(theme.HoverColor());
      FillRect(hdc, &close_rect, hot_brush);
    }

    COLORREF close_color = close_down ? theme.SelectionTextColor() : theme.TextColor();
    if (icon_font_) {
      HFONT old_font = reinterpret_cast<HFONT>(SelectObject(hdc, icon_font_));
      SetTextColor(hdc, close_color);
      SetBkMode(hdc, TRANSPARENT);
      DrawTextW(hdc, L"\xE711", -1, &close_rect, DT_SINGLELINE | DT_VCENTER | DT_CENTER);
      SelectObject(hdc, old_font);
    } else {
      HPEN close_pen = appearance::CachedPen(close_color, 2);
      HGDIOBJ old_close_pen = SelectObject(hdc, close_pen);
      int pad = std::max<int>(2, static_cast<int>((close_rect.right - close_rect.left) / 4));
      MoveToEx(hdc, close_rect.left + pad, close_rect.top + pad, nullptr);
      LineTo(hdc, close_rect.right - pad, close_rect.bottom - pad);
      MoveToEx(hdc, close_rect.right - pad, close_rect.top + pad, nullptr);
      LineTo(hdc, close_rect.left + pad, close_rect.bottom - pad);
      SelectObject(hdc, old_close_pen);
    }
  }
}

void MainWindow::Impl::PaintTabControl(HWND hwnd, HDC hdc) {
  RECT client = {};
  GetClientRect(hwnd, &client);
  const Theme& theme = Theme::Current();
  FillRect(hdc, &client, theme.BackgroundBrush());

  HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
  HGDIOBJ old_font = nullptr;
  if (font) {
    old_font = SelectObject(hdc, font);
  }

  int count = TabCtrl_GetItemCount(hwnd);
  int current = TabCtrl_GetCurSel(hwnd);

  int header_bottom = client.top;
  RECT first_rect = {};
  if (count > 0 && TabCtrl_GetItemRect(hwnd, 0, &first_rect)) {
    int row_height = first_rect.bottom - first_rect.top;
    int rows = std::max(1, TabCtrl_GetRowCount(hwnd));
    header_bottom = first_rect.top + row_height * rows + 1;
  }

  if (header_bottom > client.top) {
    HPEN line_pen = appearance::CachedPen(theme.BorderColor(), 1);
    HGDIOBJ old_pen = SelectObject(hdc, line_pen);
    MoveToEx(hdc, client.left, header_bottom, nullptr);
    LineTo(hdc, client.right, header_bottom);
    SelectObject(hdc, old_pen);
  }

  for (int i = 0; i < count; ++i) {
    if (i == current) {
      continue;
    }
    RECT item_rect = {};
    if (TabCtrl_GetItemRect(hwnd, i, &item_rect)) {
      DrawTabItem(hdc, i, item_rect, header_bottom, false);
    }
  }
  if (current >= 0) {
    RECT item_rect = {};
    if (TabCtrl_GetItemRect(hwnd, current, &item_rect)) {
      DrawTabItem(hdc, current, item_rect, header_bottom, true);
    }
  }

  if (old_font) {
    SelectObject(hdc, old_font);
  }
}

} // namespace regkit
