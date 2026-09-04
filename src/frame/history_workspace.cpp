// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::AppendHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data) {
  HistoryEntry entry;
  entry.action = action;
  entry.old_data = old_data;
  entry.new_data = new_data;
  if (browse_.current_node()) {
    entry.key_path = registry_path::Build(*browse_.current_node());
  }
  AppendHistoryEntry(std::move(entry));
}

void MainWindow::Impl::AppendValueHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data, const RegistryNode& node, const std::wstring& value_name, HistoryEntry::RevertKind revert_kind, const ValueEntry* revert_value) {
  HistoryEntry entry;
  entry.action = action;
  entry.old_data = old_data;
  entry.new_data = new_data;
  entry.key_path = registry_path::Build(node);
  entry.value_name = value_name;
  entry.revert_kind = revert_kind;
  if (revert_value) {
    entry.revert_value = *revert_value;
  }
  AppendHistoryEntry(std::move(entry));
}

void MainWindow::Impl::AppendHistoryEntry(HistoryEntry entry) {
  if (!history_list_) {
    return;
  }

  const HistoryEntry appended =
      change_history_.Append(std::move(entry),
                             static_cast<size_t>(history_max_rows_));
  if (history_loaded_) {
    AppendHistoryCache(appended);
  }
  change_history_.Sort(history_sort_column_, history_sort_ascending_);
  RebuildHistoryList();
}

bool MainWindow::Impl::PrepareHistoryRevert(const HistoryEntry& entry, HistoryEntry* prepared) const {
  auto query_value = [this](const std::wstring& path,
                            const std::wstring& name, ValueEntry* value) {
    RegistryNode node;
    return ResolvePathToNode(path, &node) &&
           RegistryStore::QueryValue(node, name, value);
  };
  if (!changes::PrepareRevert(entry, query_value, prepared)) {
    return false;
  }
  if (prepared->revert_kind != HistoryEntry::RevertKind::kDeleteValue) {
    return true;
  }

  std::vector<const HistoryEntry*> later_entries;
  later_entries.reserve(change_history_.entries().size());
  for (const HistoryEntry& candidate : change_history_.entries()) {
    if (candidate.timestamp > entry.timestamp &&
        EqualsInsensitive(candidate.key_path, entry.key_path) &&
        StartsWithInsensitive(candidate.action, L"Rename value ")) {
      later_entries.push_back(&candidate);
    }
  }
  std::stable_sort(
      later_entries.begin(), later_entries.end(),
      [](const HistoryEntry* left, const HistoryEntry* right) {
        return left->timestamp < right->timestamp;
      });
  for (const HistoryEntry* candidate : later_entries) {
    if (!EqualsInsensitive(candidate->old_data, prepared->value_name)) {
      continue;
    }
    prepared->value_name = candidate->value_name.empty()
                               ? candidate->new_data
                               : candidate->value_name;
  }

  ValueEntry current;
  return query_value(prepared->key_path, prepared->value_name, &current);
}

bool MainWindow::Impl::OpenHistoryTarget(const HistoryEntry& entry) {
  if (entry.key_path.empty()) {
    return false;
  }
  RegistryNode node;
  std::wstring target = entry.key_path;
  KeyInfo info = {};
  if (!ResolvePathToNode(target, &node) || !RegistryStore::QueryKeyInfo(node, &info)) {
    if (!FindNearestExistingPath(target, &target) || target.empty()) {
      return false;
    }
    return NavigateToResolvedExternalJump(target, L"");
  }
  return NavigateToResolvedExternalJump(target, entry.value_name);
}

bool MainWindow::Impl::RevertHistoryEntry(const HistoryEntry& entry) {
  HistoryEntry prepared;
  if (!EnsureWritable() || !PrepareHistoryRevert(entry, &prepared)) {
    return false;
  }

  bool ok = false;
  is_replaying_ = true;
  switch (prepared.revert_kind) {
  case HistoryEntry::RevertKind::kSetValue: {
    RegistryNode node;
    if (ResolvePathToNode(prepared.key_path, &node)) {
      ok = RegistryStore::SetValue(node, prepared.revert_value.name, prepared.revert_value.type, prepared.revert_value.data);
    }
    break;
  }
  case HistoryEntry::RevertKind::kDeleteValue: {
    RegistryNode node;
    if (ResolvePathToNode(prepared.key_path, &node)) {
      ok = RegistryStore::DeleteValue(node, prepared.value_name);
    }
    break;
  }
  case HistoryEntry::RevertKind::kDeleteKey: {
    RegistryNode node;
    if (ResolvePathToNode(prepared.key_path, &node)) {
      std::wstring name = LeafName(node);
      if (!name.empty() && ui::ConfirmDelete(hwnd_, L"Revert Key Creation", name)) {
        ok = RegistryStore::DeleteKey(node);
      }
    }
    break;
  }
  default:
    break;
  }
  is_replaying_ = false;

  if (!ok) {
    ui::ShowError(hwnd_, L"Failed to revert history entry.");
    return false;
  }

  MarkOfflineDirty();
  HistoryEntry revert_entry;
  revert_entry.action = L"Revert: " + entry.action;
  revert_entry.key_path = prepared.key_path;
  revert_entry.value_name = prepared.value_name;
  AppendHistoryEntry(std::move(revert_entry));
  RefreshTreeSelection();
  if (browse_.current_node()) {
    UpdateValueListForNode(browse_.current_node());
  }
  return true;
}

void MainWindow::Impl::AppendHistoryCache(const HistoryEntry& entry) {
  changes::AppendHistoryFile(HistoryCachePath(), entry);
}

std::wstring MainWindow::Impl::CacheFolderPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  std::wstring cache = util::JoinPath(folder, L"cache");
  if (!cache.empty()) {
    SHCreateDirectoryExW(nullptr, cache.c_str(), nullptr);
  }

  return cache;
}

std::wstring MainWindow::Impl::HistoryCachePath() const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"history.tsv");
}

std::wstring MainWindow::Impl::TabsCachePath() const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"tabs.ini");
}

std::wstring MainWindow::Impl::SearchTabCachePath(const std::wstring& file) const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  if (file.empty()) {
    return L"";
  }
  return util::JoinPath(folder, file);
}

bool MainWindow::Impl::EnsureSearchTabResultsLoaded(int search_index) {
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return false;
  }
  SearchTab& tab = search_tabs_[static_cast<size_t>(search_index)];
  if (tab.results_loaded) {
    return true;
  }
  tab.last_ui_count = 0;
  if (tab.cache_file.empty()) {
    tab.results_loaded = true;
    return true;
  }

  // Reading, parsing and sorting a saved tab happens on a worker.
  if (tab.load_pending) {
    return false;
  }
  tab.load_pending = true;
  StartSearchTabLoadWorker();
  auto task = std::make_unique<SearchTabLoadTask>();
  task->generation = ++search_tab_load_generation_;
  task->tab_index = search_index;
  task->path = SearchTabCachePath(tab.cache_file);
  task->sort_column = tab.sort_column;
  task->sort_ascending = tab.sort_ascending;
  task->hwnd = hwnd_;
  tab.load_generation = task->generation;
  // A displaced request never runs, so its tab must be able to ask again.
  std::unique_ptr<SearchTabLoadTask> displaced =
      search_tab_loader_.Submit(std::move(task));
  if (displaced && displaced->tab_index >= 0 &&
      static_cast<size_t>(displaced->tab_index) < search_tabs_.size()) {
    SearchTab& stale = search_tabs_[static_cast<size_t>(displaced->tab_index)];
    if (stale.load_generation == displaced->generation) {
      stale.load_pending = false;
    }
  }
  return false;
}

void MainWindow::Impl::ApplySearchTabLoad(SearchTabLoadPayload* payload) {
  if (!payload) {
    return;
  }
  std::unique_ptr<SearchTabLoadPayload> owned(payload);
  if (owned->tab_index < 0 ||
      static_cast<size_t>(owned->tab_index) >= search_tabs_.size()) {
    return;
  }
  SearchTab& tab = search_tabs_[static_cast<size_t>(owned->tab_index)];
  if (tab.load_generation != owned->generation) {
    return;
  }
  tab.load_pending = false;
  tab.results = std::move(owned->rows);
  for (auto& row : tab.results) {
    row.row_id = tab.next_row_id++;
  }
  tab.sort_dirty = false;
  tab.results_loaded = true;
  if (SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) == owned->tab_index) {
    UpdateSearchResultsView();
    UpdateStatus();
  }
}

void MainWindow::Impl::ClearTabsCache() {
  std::wstring tabs_path = TabsCachePath();
  if (!tabs_path.empty()) {
    DeleteFileW(tabs_path.c_str());
  }
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return;
  }
  std::wstring pattern = util::JoinPath(folder, L"search_*.tsv");
  WIN32_FIND_DATAW data = {};
  HANDLE find = FindFirstFileW(pattern.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
      continue;
    }
    std::wstring path = util::JoinPath(folder, data.cFileName);
    DeleteFileW(path.c_str());
  } while (FindNextFileW(find, &data) != 0);
  FindClose(find);
}

void MainWindow::Impl::LoadTabs() {
  if (!tab_) {
    return;
  }
  tabs_.clear();
  search_tabs_.clear();
  active_search_tab_index_ = -1;
  TabCtrl_DeleteAllItems(tab_);

  int active_index = 0;
  bool loaded = false;
  if (save_tabs_) {
    workspace::TabState state;
    loaded = workspace::LoadTabs(TabsCachePath(), &state);
    if (loaded) {
      active_index = state.active_index;
      for (workspace::PersistedTab& saved : state.tabs) {
        std::wstring label = std::move(saved.label);
        if (saved.kind == workspace::PersistedTab::Kind::kRegistry) {
          if (label.empty()) {
            label = L"Local Registry";
          }
          TCITEMW item = {};
          item.mask = TCIF_TEXT;
          item.pszText = label.data();
          TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
          TabEntry entry;
          entry.kind = TabEntry::Kind::kRegistry;
          entry.registry_mode = RegistryMode::kLocal;
          entry.selected_path = std::move(saved.selected_path);
          entry.expanded_paths = std::move(saved.expanded_paths);
          tabs_.push_back(std::move(entry));
        } else {
          SearchTab search_tab;
          search_tab.label = label.empty() ? L"Find" : std::move(label);
          search_tab.cache_file = std::move(saved.search_cache_file);
          search_tab.results_loaded = search_tab.cache_file.empty();
          search_tab.is_compare =
              StartsWithInsensitive(search_tab.label, L"Compare:");
          search_tabs_.push_back(std::move(search_tab));
          const int search_index =
              static_cast<int>(search_tabs_.size() - 1);
          TCITEMW item = {};
          item.mask = TCIF_TEXT;
          item.pszText = search_tabs_.back().label.data();
          TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
          tabs_.push_back({TabEntry::Kind::kSearch, search_index});
        }
      }
    }
  }

  if (!loaded || tabs_.empty()) {
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(L"Local Registry");
    TabCtrl_InsertItem(tab_, 0, &item);
    TabEntry entry;
    entry.kind = TabEntry::Kind::kRegistry;
    entry.registry_mode = RegistryMode::kLocal;
    tabs_.push_back(std::move(entry));
    active_index = 0;
  }

  int count = TabCtrl_GetItemCount(tab_);
  if (count > 0) {
    int sel = ClampValue(active_index, 0, count - 1);
    TabCtrl_SetCurSel(tab_, sel);
    if (IsSearchTabIndex(sel)) {
      active_search_tab_index_ = sel;
    }
  }
  RefreshRegistryTabLabels();
}

bool MainWindow::Impl::SaveTabs() {
  if (!tab_) {
    return true;
  }
  int current_registry_tab = CurrentRegistryTabIndex();
  if (current_registry_tab >= 0 && !IsSearchTabIndex(current_registry_tab) && !IsRegFileTabIndex(current_registry_tab)) {
    CaptureRegistryTabState(current_registry_tab);
  }
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return false;
  }

  std::unordered_set<std::wstring> referenced_files;
  std::unordered_set<std::wstring> reserved_files;
  for (const auto& search_tab : search_tabs_) {
    if (!search_tab.cache_file.empty()) {
      reserved_files.insert(search_tab.cache_file);
    }
  }
  workspace::TabState state;
  int active_index = TabCtrl_GetCurSel(tab_);
  int saved_active_index = -1;

  int search_file_index = 0;
  int tab_count = TabCtrl_GetItemCount(tab_);
  int saved_index = 0;
  for (int i = 0; i < tab_count; ++i) {
    if (static_cast<size_t>(i) >= tabs_.size()) {
      break;
    }
    const auto& entry = tabs_[static_cast<size_t>(i)];
    if (entry.kind == TabEntry::Kind::kRegFile) {
      continue;
    }
    if (i == active_index) {
      saved_active_index = saved_index;
    }
    wchar_t text[256] = {};
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = text;
    item.cchTextMax = static_cast<int>(_countof(text));
    std::wstring label;
    if (TabCtrl_GetItem(tab_, i, &item)) {
      label = text;
    }
    if (entry.kind == TabEntry::Kind::kSearch) {
      int search_index = entry.search_index;
      if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
        continue;
      }
      SearchTab& search_tab = search_tabs_[static_cast<size_t>(search_index)];
      std::wstring file_name = search_tab.cache_file;
      if (file_name.empty()) {
        do {
          file_name = L"search_" + std::to_wstring(search_file_index++) + L".tsv";
        } while (referenced_files.find(file_name) != referenced_files.end() || reserved_files.find(file_name) != reserved_files.end());
        search_tab.cache_file = file_name;
        reserved_files.insert(file_name);
      }
      if (search_tab.results_loaded) {
        std::wstring result_path = SearchTabCachePath(file_name);
        search::SaveResults(result_path, search_tab.results);
      }
      referenced_files.insert(file_name);
      if (label.empty()) {
        label = search_tab.label;
      }
      workspace::PersistedTab saved;
      saved.kind = workspace::PersistedTab::Kind::kSearch;
      saved.label = std::move(label);
      saved.search_cache_file = std::move(file_name);
      state.tabs.push_back(std::move(saved));
    } else {
      const TabEntry& registry_entry = tabs_[static_cast<size_t>(i)];
      if (label.empty()) {
        if (registry_entry.registry_mode == RegistryMode::kLocal) {
          label = LocalRegistryTabLabel(i);
        } else {
          label = L"Local Registry";
        }
      }
      workspace::PersistedTab saved;
      saved.kind = workspace::PersistedTab::Kind::kRegistry;
      saved.label = std::move(label);
      saved.selected_path = registry_entry.selected_path;
      saved.expanded_paths = registry_entry.expanded_paths;
      state.tabs.push_back(std::move(saved));
    }
    ++saved_index;
  }
  if (saved_active_index < 0) {
    saved_active_index = 0;
  }
  state.active_index = saved_active_index;
  const bool saved = workspace::SaveTabs(TabsCachePath(), state);

  std::wstring pattern = util::JoinPath(folder, L"search_*.tsv");
  WIN32_FIND_DATAW data = {};
  HANDLE find = FindFirstFileW(pattern.c_str(), &data);
  if (find != INVALID_HANDLE_VALUE) {
    do {
      if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
        continue;
      }
      if (referenced_files.find(data.cFileName) != referenced_files.end()) {
        continue;
      }
      std::wstring stale = util::JoinPath(folder, data.cFileName);
      DeleteFileW(stale.c_str());
    } while (FindNextFileW(find, &data) != 0);
    FindClose(find);
  }
  return saved;
}

std::wstring MainWindow::Impl::CommentsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"comments.tsv");
}

void MainWindow::Impl::LoadComments() {
  value_comments_.Load(CommentsPath());
}

void MainWindow::Impl::SaveComments() const {
  value_comments_.Save(CommentsPath());
}

bool MainWindow::Impl::ImportCommentsFromFile(const std::wstring& path) {
  if (!value_comments_.Import(path)) {
    return false;
  }
  value_comments_.Save(CommentsPath());
  RefreshValueListComments();
  return true;
}

bool MainWindow::Impl::ExportCommentsToFile(const std::wstring& path) const {
  return value_comments_.Export(path);
}

void MainWindow::Impl::RefreshValueListComments() {
  if (!browse_.current_node()) {
    return;
  }
  std::wstring path = registry_path::Build(*browse_.current_node());
  bool changed = false;
  for (auto& row : browse_.values().rows()) {
    if (row.kind != rowkind::kValue) {
      if (!row.comment.empty()) {
        row.comment.clear();
        changed = true;
      }
      continue;
    }
    std::wstring value_key = changes::ValueComments::ValueKey(path, row.extra, row.value_type);
    std::wstring text;
    auto it = value_comments_.value_entries().find(value_key);
    if (it != value_comments_.value_entries().end()) {
      text = it->second.text;
    } else {
      std::wstring name_key = changes::ValueComments::NameKey(row.extra, row.value_type);
      auto it2 = value_comments_.name_entries().find(name_key);
      if (it2 != value_comments_.name_entries().end()) {
        text = it2->second.text;
      }
    }
    std::wstring display = FormatCommentDisplay(text);
    if (row.comment != display) {
      row.comment = display;
      changed = true;
    }
  }
  if (browse_.columns().sort_column == kValueColComment) {
    SortValueRows(&browse_.values().rows(), browse_.columns().sort_column, browse_.columns().sort_ascending);
    changed = true;
  }
  if (changed) {
    browse_.values().InvalidateFilterCache();
  }
  if (browse_.values().HasFilter()) {
    browse_.values().RebuildFilter();
  } else if (changed && browse_.values().hwnd()) {
    RedrawWindow(browse_.values().hwnd(), nullptr, nullptr,
                 RDW_INVALIDATE | RDW_NOERASE);
  }
}

bool MainWindow::Impl::EditValueComments(const std::vector<ListRow>& rows) {
  if (!browse_.current_node() || rows.empty()) {
    return false;
  }
  std::wstring path = registry_path::Build(*browse_.current_node());

  auto resolve_comment = [&](const ListRow& row, bool* from_name) {
    std::wstring value_key = changes::ValueComments::ValueKey(path, row.extra, row.value_type);
    auto value_it = value_comments_.value_entries().find(value_key);
    if (value_it != value_comments_.value_entries().end()) {
      if (from_name) {
        *from_name = false;
      }
      return value_it->second.text;
    }
    std::wstring name_key = changes::ValueComments::NameKey(row.extra, row.value_type);
    auto name_it = value_comments_.name_entries().find(name_key);
    if (name_it != value_comments_.name_entries().end()) {
      if (from_name) {
        *from_name = true;
      }
      return name_it->second.text;
    }
    if (from_name) {
      *from_name = false;
    }
    return std::wstring();
  };

  std::wstring initial;
  bool apply_all = true;
  bool have_initial = false;
  bool comments_match = true;
  for (const auto& row : rows) {
    if (row.kind != rowkind::kValue || row.simulated) {
      return false;
    }
    bool from_name = false;
    std::wstring comment = resolve_comment(row, &from_name);
    if (!have_initial) {
      initial = std::move(comment);
      have_initial = true;
    } else if (comment != initial) {
      comments_match = false;
    }
    apply_all = apply_all && from_name;
  }
  if (!comments_match) {
    initial.clear();
    apply_all = false;
  }

  editors::CommentRequest request;
  request.text = initial;
  request.apply_to_same_name = apply_all;
  editors::CommentResult result;
  if (!editors::EditComment(hwnd_, request, &result)) {
    return false;
  }
  std::wstring updated = std::move(result.text);
  const bool apply_all_out = result.apply_to_same_name;
  if (IsWhitespaceOnly(updated)) {
    updated.clear();
  }

  for (const auto& row : rows) {
    std::wstring value_key = changes::ValueComments::ValueKey(path, row.extra, row.value_type);
    std::wstring name_key = changes::ValueComments::NameKey(row.extra, row.value_type);
    if (updated.empty()) {
      value_comments_.value_entries().erase(value_key);
      value_comments_.name_entries().erase(name_key);
    } else if (apply_all_out) {
      changes::CommentEntry entry;
      entry.name = row.extra;
      entry.type = row.value_type;
      entry.text = updated;
      value_comments_.name_entries()[name_key] = std::move(entry);
      value_comments_.value_entries().erase(value_key);
    } else {
      changes::CommentEntry entry;
      entry.path = path;
      entry.name = row.extra;
      entry.type = row.value_type;
      entry.text = updated;
      value_comments_.value_entries()[value_key] = std::move(entry);
    }
  }
  SaveComments();
  RefreshValueListComments();
  return true;
}

void MainWindow::Impl::LoadSettings() {
  workspace::Settings settings;
  settings.clear_history_on_exit = clear_history_on_exit_;
  settings.clear_tabs_on_exit = clear_tabs_on_exit_;
  settings.show_toolbar = show_toolbar_;
  settings.show_address_bar = show_address_bar_;
  settings.show_filter_bar = show_filter_bar_;
  settings.show_tab_control = show_tab_control_;
  settings.show_tree = show_tree_;
  settings.show_history = show_history_;
  settings.show_status_bar = show_status_bar_;
  settings.show_keys_in_list = show_keys_in_list_;
  settings.show_simulated_keys = show_simulated_keys_;
  settings.show_extra_hives = show_extra_hives_;
  settings.show_value_grid = show_value_grid_;
  settings.save_tree_state = save_tree_state_;
  settings.save_tabs = save_tabs_;
  settings.always_run_as_admin = always_run_as_admin_;
  settings.always_run_as_system = always_run_as_system_;
  settings.always_run_as_trustedinstaller = always_run_as_trustedinstaller_;
  settings.always_on_top = always_on_top_;
  settings.single_instance = single_instance_;
  settings.read_only = read_only_;
  settings.window_x = window_x_;
  settings.window_y = window_y_;
  settings.window_width = window_width_;
  settings.window_height = window_height_;
  settings.window_maximized = window_maximized_;
  settings.tree_width = tree_width_;
  settings.history_height = history_height_;
  settings.theme_preset = active_theme_preset_;
  settings.icon_set = icon_set_;
  settings.use_custom_font = use_custom_font_;
  settings.font_face = custom_font_.lfFaceName;
  settings.font_size = appearance::FontPointSize(custom_font_, 9);
  settings.font_weight = custom_font_.lfWeight;
  settings.font_italic = custom_font_.lfItalic != FALSE;
  settings.value_column_widths = browse_.columns().saved_widths;
  settings.value_column_visible = browse_.columns().saved_visible;

  if (!workspace::LoadSettings(SettingsPath(), &settings)) {
    return;
  }

  clear_history_on_exit_ = settings.clear_history_on_exit;
  clear_tabs_on_exit_ = settings.clear_tabs_on_exit;
  show_toolbar_ = settings.show_toolbar;
  show_address_bar_ = settings.show_address_bar;
  show_filter_bar_ = settings.show_filter_bar;
  show_tab_control_ = settings.show_tab_control;
  show_tree_ = settings.show_tree;
  show_history_ = settings.show_history;
  show_status_bar_ = settings.show_status_bar;
  show_keys_in_list_ = settings.show_keys_in_list;
  show_simulated_keys_ = settings.show_simulated_keys;
  show_extra_hives_ = settings.show_extra_hives;
  show_value_grid_ = settings.show_value_grid;
  save_tree_state_ = settings.save_tree_state;
  save_tabs_ = settings.save_tabs;
  always_run_as_admin_ = settings.always_run_as_admin;
  always_run_as_system_ = settings.always_run_as_system;
  always_run_as_trustedinstaller_ = settings.always_run_as_trustedinstaller;
  always_on_top_ = settings.always_on_top;
  single_instance_ = settings.single_instance;
  read_only_ = settings.read_only;
  window_placement_loaded_ = settings.window_placement_present;
  window_x_ = settings.window_x;
  window_y_ = settings.window_y;
  window_width_ = settings.window_width;
  window_height_ = settings.window_height;
  window_maximized_ = settings.window_maximized;
  tree_width_ = settings.tree_width;
  history_height_ = settings.history_height;
  if (_wcsicmp(settings.theme_mode.c_str(), L"dark") == 0) {
    theme_mode_ = ThemeMode::kDark;
  } else if (_wcsicmp(settings.theme_mode.c_str(), L"light") == 0) {
    theme_mode_ = ThemeMode::kLight;
  } else if (_wcsicmp(settings.theme_mode.c_str(), L"custom") == 0) {
    theme_mode_ = ThemeMode::kCustom;
  } else {
    theme_mode_ = ThemeMode::kSystem;
  }
  active_theme_preset_ = std::move(settings.theme_preset);
  icon_set_ = IsIconSetName(settings.icon_set, kIconSetLegacyDefault)
                  ? kIconSetDefault
                  : (IsKnownIconSetName(settings.icon_set)
                         ? std::move(settings.icon_set)
                         : kIconSetDefault);
  use_custom_font_ = settings.use_custom_font;
  if (!settings.font_face.empty()) {
    wcsncpy_s(custom_font_.lfFaceName, settings.font_face.c_str(), _TRUNCATE);
  }
  if (settings.font_size > 0) {
    custom_font_.lfHeight = appearance::FontHeight(settings.font_size);
  }
  custom_font_.lfWeight = settings.font_weight;
  custom_font_.lfItalic = settings.font_italic ? TRUE : FALSE;
  recent_trace_paths_.Replace(std::move(settings.recent_traces));
  recent_default_paths_.Replace(std::move(settings.recent_defaults));
  browse_.columns().saved_widths = std::move(settings.value_column_widths);
  browse_.columns().saved_visible = std::move(settings.value_column_visible);
  browse_.columns().saved = !browse_.columns().saved_widths.empty() ||
                            !browse_.columns().saved_visible.empty();
  if (!save_tree_state_) {
    saved_tree_state_.Clear();
  }
}
void MainWindow::Impl::SaveSettings() const {
  workspace::Settings settings;
  settings.clear_history_on_exit = clear_history_on_exit_;
  settings.clear_tabs_on_exit = clear_tabs_on_exit_;
  settings.show_toolbar = show_toolbar_;
  settings.show_address_bar = show_address_bar_;
  settings.show_filter_bar = show_filter_bar_;
  settings.show_tab_control = show_tab_control_;
  settings.show_tree = show_tree_;
  settings.show_history = show_history_;
  settings.show_status_bar = show_status_bar_;
  settings.show_keys_in_list = show_keys_in_list_;
  settings.show_simulated_keys = show_simulated_keys_;
  settings.show_extra_hives = show_extra_hives_;
  settings.show_value_grid = show_value_grid_;
  settings.save_tree_state = save_tree_state_;
  settings.save_tabs = save_tabs_;
  settings.always_run_as_admin = always_run_as_admin_;
  settings.always_run_as_system = always_run_as_system_;
  settings.always_run_as_trustedinstaller = always_run_as_trustedinstaller_;
  settings.always_on_top = always_on_top_;
  settings.single_instance = single_instance_;
  settings.read_only = read_only_;
  settings.window_x = window_x_;
  settings.window_y = window_y_;
  settings.window_width = window_width_;
  settings.window_height = window_height_;
  settings.window_maximized = window_maximized_;
  if (hwnd_ && IsWindow(hwnd_)) {
    WINDOWPLACEMENT placement = {};
    placement.length = sizeof(placement);
    if (GetWindowPlacement(hwnd_, &placement)) {
      const RECT& normal = placement.rcNormalPosition;
      const int width = normal.right - normal.left;
      const int height = normal.bottom - normal.top;
      if (width > 0 && height > 0) {
        settings.window_x = normal.left;
        settings.window_y = normal.top;
        settings.window_width = width;
        settings.window_height = height;
      }
      settings.window_maximized = placement.showCmd == SW_SHOWMAXIMIZED;
    }
  }
  settings.tree_width = tree_width_;
  settings.history_height = history_height_;
  switch (theme_mode_) {
  case ThemeMode::kDark:
    settings.theme_mode = L"dark";
    break;
  case ThemeMode::kLight:
    settings.theme_mode = L"light";
    break;
  case ThemeMode::kCustom:
    settings.theme_mode = L"custom";
    break;
  default:
    settings.theme_mode = L"system";
    break;
  }
  settings.theme_preset = active_theme_preset_;
  settings.icon_set = IsKnownIconSetName(icon_set_) ? icon_set_ : kIconSetDefault;
  settings.use_custom_font = use_custom_font_;
  settings.font_face = custom_font_.lfFaceName;
  settings.font_size = appearance::FontPointSize(custom_font_, 9);
  settings.font_weight = custom_font_.lfWeight;
  settings.font_italic = custom_font_.lfItalic != FALSE;
  settings.recent_traces = recent_trace_paths_.items();
  settings.recent_defaults = recent_default_paths_.items();
  settings.value_column_widths = browse_.columns().widths;
  settings.value_column_visible = browse_.columns().visible;
  settings.value_column_widths.resize(browse_.columns().items.size(), 0);
  settings.value_column_visible.resize(browse_.columns().items.size(), true);
  workspace::SaveSettings(SettingsPath(), settings);
}
std::wstring MainWindow::Impl::SettingsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"settings.ini");
}

std::wstring MainWindow::Impl::TreeStatePath() const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"tree_state.ini");
}

void MainWindow::Impl::LoadTreeState() {
  saved_tree_state_.Clear();
  if (!save_tree_state_) {
    return;
  }
  workspace::LoadTreeState(TreeStatePath(), &saved_tree_state_);
}

void MainWindow::Impl::StartTreeStateWorker() {
  if (!save_tree_state_ || tree_state_saver_.running()) {
    return;
  }
  tree_state_saver_.Start(
      std::chrono::seconds(2),
      [this](workspace::TreeState state) {
        SaveTreeStateFile(state.selected_path, state.expanded_paths);
      });
}

void MainWindow::Impl::StopTreeStateWorker() {
  tree_state_saver_.Stop();
  if (save_tree_state_ && browse_.tree().hwnd() && IsWindow(browse_.tree().hwnd())) {
    std::wstring selected;
    std::vector<std::wstring> expanded;
    CaptureTreeState(&selected, &expanded);
    SaveTreeStateFile(selected, expanded);
  }
}

} // namespace regkit
