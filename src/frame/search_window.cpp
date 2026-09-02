// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

std::wstring MainWindow::Impl::NormalizeRegistryPath(const std::wstring& input) const {
  const std::wstring sid = util::GetCurrentUserSidString();
  std::wstring path = registry_path::Normalize(input, sid);
  auto strip_context = [&](const std::wstring& label) {
    if (label.empty()) {
      return;
    }
    const std::wstring prefix = label + L"\\";
    if (registry_path::StartsWith(path, prefix)) {
      path.erase(0, prefix.size());
    }
  };
  strip_context(TreeRootLabel());
  if (registry_mode_ == RegistryMode::kRemote) {
    strip_context(StripMachinePrefix(remote_machine_));
  }
  return registry_path::Normalize(path, sid);
}

std::wstring MainWindow::Impl::FormatRegistryPath(const std::wstring& path, RegistryPathFormat format) const {
  const std::wstring normalized = NormalizeRegistryPath(path);
  if (normalized.empty()) {
    return {};
  }
  std::wstring tree_root =
      registry_mode_ == RegistryMode::kLocal ? L"Computer" : TreeRootLabel();
  registry_path::Style style = registry_path::Style::kFull;
  switch (format) {
  case RegistryPathFormat::kAbbrev:
    style = registry_path::Style::kAbbreviated;
    break;
  case RegistryPathFormat::kRegedit:
    style = registry_path::Style::kRegeditAddress;
    break;
  case RegistryPathFormat::kRegFile:
    style = registry_path::Style::kRegFileHeader;
    break;
  case RegistryPathFormat::kPowerShellDrive:
    style = registry_path::Style::kPowerShellDrive;
    break;
  case RegistryPathFormat::kPowerShellProvider:
    style = registry_path::Style::kPowerShellProvider;
    break;
  case RegistryPathFormat::kEscaped:
    style = registry_path::Style::kEscaped;
    break;
  case RegistryPathFormat::kFull:
    break;
  }
  return registry_path::Format(normalized, style, tree_root);
}
bool MainWindow::Impl::FindNearestExistingPath(const std::wstring& path, std::wstring* nearest_path) const {
  return changes::FindNearestExistingPath(
      path,
      [this](const std::wstring& candidate) {
        RegistryNode node;
        KeyInfo info = {};
        return ResolvePathToNode(candidate, &node) &&
               RegistryStore::QueryKeyInfo(node, &info);
      },
      nearest_path);
}

bool MainWindow::Impl::CreateRegistryPath(const std::wstring& path) {
  RegistryNode node;
  if (!ResolvePathToNode(path, &node)) {
    return false;
  }
  if (node.subkey.empty()) {
    return true;
  }
  std::vector<std::wstring> parts = registry_path::Split(node.subkey);
  RegistryNode current = node;
  current.subkey.clear();
  bool created = false;
  for (const auto& part : parts) {
    if (!RegistryStore::CreateKey(current, part)) {
      return false;
    }
    created = true;
    if (current.subkey.empty()) {
      current.subkey = part;
    } else {
      current.subkey += L"\\" + part;
    }
  }
  if (created) {
    MarkOfflineDirty();
  }
  return true;
}

void MainWindow::Impl::UpdateStatus() {
  if (!status_bar_) {
    return;
  }
  RECT rc = {};
  GetClientRect(status_bar_, &rc);
  int total_width = rc.right - rc.left;
  if (total_width < 0) {
    total_width = 0;
  }
  LONG_PTR sb_style = GetWindowLongPtrW(status_bar_, GWL_STYLE);
  if (sb_style & SBARS_SIZEGRIP) {
    int grip = GetSystemMetrics(SM_CXVSCROLL);
    total_width = std::max(total_width - grip, 0);
  }
  auto measure_text = [&](HDC hdc, const std::wstring& text) -> int {
    if (!hdc || text.empty()) {
      return 0;
    }
    SIZE size = {};
    GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &size);
    return size.cx + 20;
  };
  if (IsSearchTabSelected()) {
    bool compare_selected = IsCompareTabSelected();
    int sel = TabCtrl_GetCurSel(tab_);
    int tab_index = SearchIndexFromTab(sel);
    size_t count = 0;
    if (tab_index >= 0 && static_cast<size_t>(tab_index) < search_tabs_.size()) {
      count = search_tabs_[static_cast<size_t>(tab_index)].results.size();
    }
    unsigned long long count_value = static_cast<unsigned long long>(count);
    wchar_t buffer[256] = {};
    if (compare_selected) {
      swprintf_s(buffer, L"Differences: %llu", count_value);
    } else if (search_running_) {
      uint64_t searched = search_progress_searched_.load();
      if (searched > 0) {
        swprintf_s(buffer, L"Searching... Results: ~%llu | Scanned: %llu", count_value, searched);
      } else {
        swprintf_s(buffer, L"Searching... Results: ~%llu", count_value);
      }
    } else if (search_duration_valid_ && search_duration_ms_ > 0) {
      double seconds = static_cast<double>(search_duration_ms_) / 1000.0;
      swprintf_s(buffer, L"Results: %llu (%.2fs)", count_value, seconds);
    } else {
      swprintf_s(buffer, L"Results: %llu", count_value);
    }
    int part = total_width;
    SendMessageW(status_bar_, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(&part));
    SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(buffer));
    return;
  }
  if (IsRegFileTabSelected()) {
    int sel = TabCtrl_GetCurSel(tab_);
    if (IsRegFileTabIndex(sel) && static_cast<size_t>(sel) < tabs_.size()) {
      const TabEntry& entry = tabs_[static_cast<size_t>(sel)];
      if (entry.reg_file_loading) {
        std::wstring label = entry.reg_file_label.empty() ? L"registry file" : entry.reg_file_label;
        std::wstring text = L"Loading " + label + L"...";
        int part = total_width;
        SendMessageW(status_bar_, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(&part));
        SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(text.c_str()));
        return;
      }
    }
  }

  int selected = ListView_GetSelectedCount(browse_.values().hwnd());
  wchar_t buffer[256] = {};
  std::wstring keys_text;
  std::wstring values_text;
  std::wstring selected_text;
  std::wstring path_text;
  if (browse_.current_node()) {
    path_text = registry_path::Build(*browse_.current_node());
  }
  swprintf_s(buffer, L"Keys: %d", current_key_count_);
  keys_text = buffer;
  swprintf_s(buffer, L"Values: %d", current_value_count_);
  values_text = buffer;
  swprintf_s(buffer, L"Selected: %d", selected);
  selected_text = buffer;

  HDC hdc = GetDC(status_bar_);
  HFONT old_font = nullptr;
  if (hdc && ui_font_) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }
  int values_width = measure_text(hdc, values_text);
  int selected_width = measure_text(hdc, selected_text);
  int keys_width = measure_text(hdc, keys_text);
  if (old_font) {
    SelectObject(hdc, old_font);
  }
  if (hdc) {
    ReleaseDC(status_bar_, hdc);
  }

  int part3 = total_width;
  int part2 = std::max(part3 - keys_width, 0);
  int part1 = std::max(part2 - selected_width, 0);
  int part0 = std::max(part1 - values_width, 0);
  int parts[4] = {part0, part1, part2, part3};
  SendMessageW(status_bar_, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));
  SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(path_text.c_str()));
  SendMessageW(status_bar_, SB_SETTEXTW, 1, reinterpret_cast<LPARAM>(values_text.c_str()));
  SendMessageW(status_bar_, SB_SETTEXTW, 2, reinterpret_cast<LPARAM>(selected_text.c_str()));
  SendMessageW(status_bar_, SB_SETTEXTW, 3, reinterpret_cast<LPARAM>(keys_text.c_str()));
}

bool MainWindow::Impl::IsSearchTabSelected() const {
  if (!tab_) {
    return false;
  }
  int index = TabCtrl_GetCurSel(tab_);
  return IsSearchTabIndex(index);
}

bool MainWindow::Impl::IsRegFileTabSelected() const {
  if (!tab_) {
    return false;
  }
  int index = TabCtrl_GetCurSel(tab_);
  return IsRegFileTabIndex(index);
}

bool MainWindow::Impl::IsCompareTabSelected() const {
  if (!tab_) {
    return false;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (!IsSearchTabIndex(index)) {
    return false;
  }
  int search_index = SearchIndexFromTab(index);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return false;
  }
  return search_tabs_[static_cast<size_t>(search_index)].is_compare;
}

bool MainWindow::Impl::IsSearchTabIndex(int index) const {
  if (index < 0) {
    return false;
  }
  if (static_cast<size_t>(index) >= tabs_.size()) {
    return false;
  }
  return tabs_[static_cast<size_t>(index)].kind == TabEntry::Kind::kSearch;
}

bool MainWindow::Impl::IsRegFileTabIndex(int index) const {
  if (index < 0) {
    return false;
  }
  if (static_cast<size_t>(index) >= tabs_.size()) {
    return false;
  }
  return tabs_[static_cast<size_t>(index)].kind == TabEntry::Kind::kRegFile;
}

int MainWindow::Impl::SearchIndexFromTab(int index) const {
  if (!IsSearchTabIndex(index)) {
    return -1;
  }
  return tabs_[static_cast<size_t>(index)].search_index;
}

int MainWindow::Impl::FindFirstRegistryTabIndex() const {
  for (size_t i = 0; i < tabs_.size(); ++i) {
    if (tabs_[i].kind == TabEntry::Kind::kRegistry) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

void MainWindow::Impl::SyncRegFileTabSelection() {
  if (!tab_) {
    return;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (!IsRegFileTabIndex(index)) {
    return;
  }
  if (static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  const TabEntry& entry = tabs_[static_cast<size_t>(index)];
  registry_mode_ = RegistryMode::kLocal;
  browse_.set_source_kind(browse::SourceKind::kVirtual);
  std::vector<RegistryRootEntry> roots;
  roots.reserve(entry.reg_file_roots.size());
  for (const auto& root : entry.reg_file_roots) {
    if (!root.root) {
      continue;
    }
    RegistryRootEntry reg_root;
    reg_root.root = root.root;
    reg_root.display_name = root.name;
    reg_root.path_name = root.name;
    reg_root.subkey_prefix = L"";
    reg_root.group = RegistryRootGroup::kStandard;
    roots.push_back(std::move(reg_root));
  }
  ApplyRegistryRoots(roots);
  RestoreRegistryTabState(index);
}

void MainWindow::Impl::UpdateSearchResultsView() {
  if (!search_results_list_) {
    return;
  }
  int sel = TabCtrl_GetCurSel(tab_);
  if (!IsSearchTabIndex(sel)) {
    ListView_SetItemCountEx(search_results_list_, 0, LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
    search_results_view_tab_index_ = -1;
    return;
  }
  int search_index = SearchIndexFromTab(sel);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return;
  }
  EnsureSearchTabResultsLoaded(search_index);
  bool force_redraw = (search_results_view_tab_index_ != sel);
  search_results_view_tab_index_ = sel;
  auto& tab = search_tabs_[static_cast<size_t>(search_index)];
  bool compare = tab.is_compare;
  if (compare != compare_columns_active_) {
    ApplySearchColumns(compare);
    force_redraw = true;
  }
  int max_sort_col = compare ? 3 : 5;
  if (tab.sort_column > max_sort_col) {
    tab.sort_column = -1;
  }
  UpdateListViewSort(search_results_list_, tab.sort_column, tab.sort_ascending);
  HWND header = ListView_GetHeader(search_results_list_);
  if (header) {
    InvalidateRect(header, nullptr, TRUE);
  }
  size_t count = tab.results.size();
  size_t old_count = tab.last_ui_count;
  if (force_redraw || count != old_count) {
    ListView_SetItemCountEx(search_results_list_, static_cast<int>(count), LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
    if (force_redraw || count < old_count) {
      InvalidateRect(search_results_list_, nullptr, TRUE);
    } else if (count > old_count) {
      int first = static_cast<int>(old_count);
      int last = static_cast<int>(count - 1);
      ListView_RedrawItems(search_results_list_, first, last);
    }
    tab.last_ui_count = count;
  }
}

void MainWindow::Impl::StartSearch(const SearchDialogResult& options) {
  if (options.criteria.query.empty()) {
    ui::ShowWarning(hwnd_, L"Enter text to find.");
    return;
  }

  bool matcher_ok = true;
  TextMatcher matcher(options.criteria.query, options.criteria.use_regex, options.criteria.match_case, options.criteria.match_whole, &matcher_ok);
  if (!matcher_ok) {
    ui::ShowError(hwnd_, L"Invalid regex.");
    return;
  }

  bool want_registry = options.search_standard_hives || options.search_registry_root;
  bool want_trace = options.search_trace_values && !active_traces_.empty();
  std::wstring registry_scope_path;
  std::wstring scope_path;
  if (options.scope == SearchScope::kCurrentKey) {
    if (!options.start_key.empty()) {
      registry_scope_path = options.start_key;
      scope_path = NormalizeRegistryPath(options.start_key);
    } else if (browse_.current_node()) {
      registry_scope_path = registry_path::Build(*browse_.current_node());
      scope_path = NormalizeRegistryPath(registry_scope_path);
    } else {
      ui::ShowError(hwnd_, L"Select a starting key first.");
      return;
    }
  }

  std::vector<RegistryNode> start_nodes;
  if (want_registry) {
    if (options.scope == SearchScope::kCurrentKey) {
      if (!registry_scope_path.empty()) {
        RegistryNode node;
        if (ResolvePathToNode(registry_scope_path, &node)) {
          start_nodes.push_back(node);
        } else {
          std::wstring normalized = NormalizeRegistryPath(registry_scope_path);
          if (!normalized.empty() && ResolvePathToNode(normalized, &node)) {
            start_nodes.push_back(node);
          } else {
            ui::ShowError(hwnd_, L"Starting key path was not found.");
            return;
          }
        }
      } else if (browse_.current_node()) {
        start_nodes.push_back(*browse_.current_node());
      } else {
        ui::ShowError(hwnd_, L"Select a starting key first.");
        return;
      }
    } else {
      std::unordered_set<std::wstring> seen;
      auto add_root = [&](const RegistryRootEntry& entry) {
        std::wstring key = ToLower(entry.path_name.empty() ? entry.display_name : entry.path_name);
        if (key.empty()) {
          return;
        }
        if (!seen.insert(key).second) {
          return;
        }
        RegistryNode node;
        node.root = entry.root;
        node.root_name = entry.path_name;
        node.subkey = entry.subkey_prefix;
        start_nodes.push_back(std::move(node));
      };

      if (options.search_standard_hives) {
        for (const auto& path : options.root_paths) {
          for (const auto& root : browse_.roots()) {
            if (_wcsicmp(root.path_name.c_str(), path.c_str()) == 0 || _wcsicmp(root.display_name.c_str(), path.c_str()) == 0) {
              add_root(root);
              break;
            }
          }
        }
        if (start_nodes.empty()) {
          for (const auto& root : browse_.roots()) {
            if (root.group == RegistryRootGroup::kStandard) {
              add_root(root);
            }
          }
        }
      }
      if (options.search_registry_root) {
        for (const auto& root : browse_.roots()) {
          if (_wcsicmp(root.path_name.c_str(), L"REGISTRY") == 0 || _wcsicmp(root.display_name.c_str(), L"REGISTRY") == 0) {
            add_root(root);
            break;
          }
        }
      }
    }
  }

  if (want_registry && start_nodes.empty()) {
    ui::ShowError(hwnd_, L"Select at least one top-level key.");
    return;
  }
  if (!want_registry && !want_trace) {
    return;
  }
  if (!tab_) {
    return;
  }

  CancelSearch();

  search::Criteria criteria = options.criteria;
  criteria.start_nodes = start_nodes;
  criteria.exclude_paths = options.exclude_paths;

  std::wstring label = L"Find";
  if (!criteria.query.empty()) {
    label = L"Find: " + criteria.query;
    constexpr size_t kMaxLabel = 48;
    if (label.size() > kMaxLabel) {
      label.resize(kMaxLabel - 3);
      label.append(L"...");
    }
  }

  int tab_index = -1;
  int search_index = -1;
  bool reuse_tab = options.result_mode == SearchResultMode::kReuseTab;
  if (reuse_tab) {
    int sel = TabCtrl_GetCurSel(tab_);
    int candidate = IsSearchTabIndex(sel) ? sel : active_search_tab_index_;
    if (IsSearchTabIndex(candidate)) {
      int index = SearchIndexFromTab(candidate);
      if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size() && !search_tabs_[static_cast<size_t>(index)].is_compare) {
        tab_index = candidate;
        search_index = index;
      }
    }
  }

  if (search_index >= 0) {
    SearchTab& tab = search_tabs_[static_cast<size_t>(search_index)];
    tab.label = label;
    tab.results.clear();
    tab.last_ui_count = 0;
    tab.is_compare = false;
    tab.sort_dirty = false;
    tab.open_in_new_tab = options.open_in_new_tab;
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(tab.label.c_str());
    TabCtrl_SetItem(tab_, tab_index, &item);
  } else {
    SearchTab tab;
    tab.label = label;
    tab.is_compare = false;
    tab.open_in_new_tab = options.open_in_new_tab;
    search_tabs_.push_back(std::move(tab));
    search_index = static_cast<int>(search_tabs_.size() - 1);
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(search_tabs_.back().label.c_str());
    tab_index = TabCtrl_GetItemCount(tab_);
    TabCtrl_InsertItem(tab_, tab_index, &item);
    tabs_.push_back({TabEntry::Kind::kSearch, search_index});
  }

  UpdateTabWidth();
  SelectTabIndex(tab_index);
  active_search_tab_index_ = tab_index;
  search_results_view_tab_index_ = -1;
  search_progress_searched_.store(0);
  search_progress_total_.store(0);
  search_progress_percent_ = 0;
  search_progress_posted_.store(false);
  search_posted_.store(false);
  {
    std::lock_guard<std::mutex> lock(search_mutex_);
    std::vector<PendingSearchResult>().swap(search_pending_);
  }
  search_last_refresh_tick_ = 0;
  search_start_tick_ = GetTickCount64();
  search_duration_ms_ = 0;
  search_duration_valid_ = false;
  search_running_ = true;

  if (search_progress_) {
    SendMessageW(search_progress_, PBM_SETMARQUEE, TRUE, 30);
  }

  ApplyViewVisibility();
  UpdateSearchResultsView();
  UpdateStatus();

  std::vector<ActiveTrace> traces = active_traces_;
  std::vector<std::wstring> exclude_paths = options.exclude_paths;
  std::wstring scope_lower = ToLower(scope_path);
  bool scope_recursive = criteria.recursive;
  bool trace_enabled = want_trace;
  bool registry_enabled = want_registry && !criteria.start_nodes.empty();

  const uint64_t generation = search_session_.Start(
      [this, criteria, traces, exclude_paths, scope_lower, scope_recursive,
       trace_enabled, registry_enabled, matcher](
          uint64_t generation, std::atomic_bool& cancel) mutable {
        auto should_stop = [&]() { return cancel.load(); };

        std::vector<PendingSearchResult> batch;
        batch.reserve(kSearchQueueBatch);
        std::mutex batch_mutex;

        auto flush = [&]() {
          std::vector<PendingSearchResult> pending;
          {
            std::lock_guard<std::mutex> lock(batch_mutex);
            if (batch.empty()) {
              return;
            }
            pending.swap(batch);
          }
          {
            std::lock_guard<std::mutex> lock(search_mutex_);
            search_pending_.reserve(search_pending_.size() + pending.size());
            for (auto& item : pending) {
              search_pending_.push_back(std::move(item));
            }
          }
          if (!search_posted_.exchange(true)) {
            PostMessageW(hwnd_, frame::message_id::kSearchResults, static_cast<WPARAM>(generation), 0);
          }
        };

        auto queue_result = [&](search::Result&& result) {
          bool should_flush = false;
          {
            std::lock_guard<std::mutex> lock(batch_mutex);
            PendingSearchResult pending;
            pending.generation = generation;
            pending.result = std::move(result);
            batch.push_back(std::move(pending));
            should_flush = batch.size() >= kSearchQueueBatch;
          }
          if (should_flush) {
            flush();
          }
        };

        auto is_excluded = [&](const std::wstring& path) {
          if (exclude_paths.empty()) {
            return false;
          }
          for (const auto& exclude : exclude_paths) {
            if (exclude.empty()) {
              continue;
            }
            if (FindStringOrdinal(FIND_FROMSTART, path.c_str(), static_cast<int>(path.size()), exclude.c_str(), static_cast<int>(exclude.size()), TRUE) >= 0) {
              return true;
            }
          }
          return false;
        };

        auto key_in_scope = [&](const std::wstring& key_lower) {
          if (scope_lower.empty()) {
            return true;
          }
          if (key_lower == scope_lower) {
            return true;
          }
          if (!scope_recursive) {
            return false;
          }
          if (key_lower.size() <= scope_lower.size()) {
            return false;
          }
          if (key_lower.compare(0, scope_lower.size(), scope_lower) != 0) {
            return false;
          }
          return key_lower[scope_lower.size()] == L'\\';
        };

        if (trace_enabled) {
          for (const auto& trace : traces) {
            if (should_stop()) {
              break;
            }
            if (!trace.data) {
              continue;
            }
            std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
            for (const auto& key_path : trace.data->key_paths) {
              if (should_stop()) {
                break;
              }
              if (key_path.empty()) {
                continue;
              }
              if (is_excluded(key_path)) {
                continue;
              }
              std::wstring key_lower = ToLower(key_path);
              if (!trace.selection ||
                  !trace::IncludesKey(*trace.selection, key_lower)) {
                continue;
              }
              if (!key_in_scope(key_lower)) {
                continue;
              }
              std::wstring key_name = registry_path::Leaf(key_path);

              if (criteria.search_keys) {
                TextMatch match = matcher.Match(key_name);
                if (match.matched) {
                  search::Result result;
                  result.key_path = key_path;
                  result.key_name = key_name;
                  result.type_text = L"Trace Key";
                  result.is_key = true;
                  size_t path_start = key_path.size() >= key_name.size() ? key_path.size() - key_name.size() : 0;
                  result.match_field = search::MatchField::kPath;
                  result.match_start = static_cast<int>(path_start + match.start);
                  result.match_length = static_cast<int>(match.length);
                  queue_result(std::move(result));
                }
              }

              if (criteria.search_values) {
                auto it = trace.data->values_by_key.find(key_lower);
                if (it != trace.data->values_by_key.end()) {
                  for (const auto& value_name : it->second.values_display) {
                    if (should_stop()) {
                      break;
                    }
                    std::wstring value_lower = ToLower(value_name);
                    if (!trace.selection ||
                        !trace::IncludesValue(*trace.selection, key_lower,
                                              value_lower)) {
                      continue;
                    }
                    std::wstring display = value_name.empty() ? L"(Default)" : value_name;
                    TextMatch match = matcher.Match(display);
                    if (!match.matched) {
                      continue;
                    }
                    search::Result result;
                    result.key_path = key_path;
                    result.key_name = key_name;
                    result.value_name = value_name;
                    result.display_name = display;
                    result.type_text = L"Trace Value";
                    result.is_key = false;
                    result.match_field = search::MatchField::kName;
                    result.match_start = static_cast<int>(match.start);
                    result.match_length = static_cast<int>(match.length);
                    queue_result(std::move(result));
                  }
                }
              }
            }
          }
          flush();
        }

        if (!should_stop() && registry_enabled) {
          std::atomic<uint64_t> last_progress_tick{0};
          auto progress_cb = [&](uint64_t searched, uint64_t total) {
            search_progress_searched_.store(searched);
            search_progress_total_.store(total);
            uint64_t now = GetTickCount64();
            uint64_t last = last_progress_tick.load();
            if (now - last < kSearchProgressUiMs && searched < total) {
              return;
            }
            if (last_progress_tick.compare_exchange_strong(last, now)) {
              if (!search_progress_posted_.exchange(true)) {
                PostMessageW(hwnd_, frame::message_id::kSearchProgress, static_cast<WPARAM>(generation), 0);
              }
            }
          };
          bool ok = search::Run(
              criteria, &cancel,
              [&](search::Result&& result) -> bool {
                if (should_stop()) {
                  return false;
                }
                queue_result(std::move(result));
                return !should_stop();
              },
              progress_cb, false);
          flush();
          if (!ok) {
            PostMessageW(hwnd_, frame::message_id::kSearchFailed, static_cast<WPARAM>(generation), 0);
            return;
          }
        }

        flush();
        PostMessageW(hwnd_, frame::message_id::kSearchFinished, static_cast<WPARAM>(generation), 0);
      });
  search_tabs_[static_cast<size_t>(search_index)].generation = generation;
}

void MainWindow::Impl::StartReplace(const ReplaceDialogResult& options) {
  if (read_only_) {
    ui::ShowWarning(hwnd_, L"Read only mode is enabled.");
    return;
  }
  if (options.find_text.empty()) {
    return;
  }

  RegistryNode start;
  if (!options.start_key.empty()) {
    if (!ResolvePathToNode(options.start_key, &start)) {
      ui::ShowError(hwnd_, L"Starting key path was not found.");
      return;
    }
  } else if (browse_.current_node()) {
    start = *browse_.current_node();
  } else {
    ui::ShowError(hwnd_, L"Select a starting key first.");
    return;
  }

  search::Replacer matcher(options);
  if (!matcher.valid()) {
    ui::ShowError(hwnd_, L"Invalid replace pattern.");
    return;
  }

  if (replace_result_pending_) {
    ui::ShowWarning(hwnd_, L"Replace is already running.");
    return;
  }

  const HWND hwnd = hwnd_;
  replace_result_pending_ = true;
  replace_session_.Start(
      [this, start, options, matcher, hwnd](
          uint64_t generation, std::atomic_bool& cancel) mutable {
        auto payload = std::make_unique<ReplacePayload>();
        payload->generation = generation;
        std::vector<RegistryNode> stack;
        stack.push_back(start);

        while (!stack.empty() && !cancel.load()) {
          RegistryNode node = std::move(stack.back());
          stack.pop_back();

          std::vector<ValueEntry> values;
          RegistryStore::KeyEnumResult enum_result;
          bool values_reserved = false;
          RegistryStore::EnumKeyStreaming(
              node, true, true, false, &enum_result,
              [&](const ValueInfo& info, const BYTE* data, DWORD data_size) {
                if (!values_reserved) {
                  if (enum_result.info_valid) {
                    values.reserve(enum_result.info.value_count);
                  }
                  values_reserved = true;
                }
                ValueEntry value;
                value.name = info.name;
                value.type = info.type;
                if (data_size > 0 && data) {
                  value.data.assign(data, data + data_size);
                }
                values.push_back(std::move(value));
                return !cancel.load();
              },
              {});

          for (const auto& value : values) {
            if (cancel.load()) {
              break;
            }

            std::wstring current_name = value.name;
            std::wstring replaced_name;
            if (!current_name.empty() &&
                matcher.Replace(current_name, &replaced_name) &&
                replaced_name != current_name) {
              if (replaced_name.empty()) {
                continue;
              }
              std::wstring unique =
                  MakeUniqueValueName(node, replaced_name);
              if (!RegistryStore::RenameValue(node, current_name, unique)) {
                ++payload->failures;
              } else {
                ReplacePayload::Change change;
                change.undo.type =
                    changes::UndoOperation::Type::kRenameValue;
                change.undo.node = node;
                change.undo.name = current_name;
                change.undo.new_name = unique;
                change.history.action = L"Rename value " + current_name;
                change.history.old_data = current_name;
                change.history.new_data = unique;
                change.history.key_path = registry_path::Build(node);
                change.history.value_name = unique;
                payload->changes.push_back(std::move(change));
                current_name = unique;
              }
            }

            if (value.type != REG_SZ &&
                value.type != REG_EXPAND_SZ &&
                value.type != REG_MULTI_SZ) {
              continue;
            }

            std::vector<BYTE> new_data = value.data;
            bool changed = false;
            if (value.type == REG_MULTI_SZ) {
              auto parts = value_format::MultiStringItems(value.data);
              for (auto& part : parts) {
                std::wstring updated;
                if (matcher.Replace(part, &updated) && updated != part) {
                  part = std::move(updated);
                  changed = true;
                }
              }
              if (changed) {
                new_data = value_format::MultiStringData(parts);
              }
            } else {
              const std::wstring text = value_format::Data(
                  value.type, value.data.data(),
                  static_cast<DWORD>(value.data.size()));
              std::wstring updated;
              if (matcher.Replace(text, &updated) && updated != text) {
                new_data = value_format::StringData(updated);
                changed = true;
              }
            }

            if (!changed) {
              continue;
            }
            if (!RegistryStore::SetValue(node, current_name, value.type,
                                         new_data)) {
              ++payload->failures;
              continue;
            }

            ValueEntry old_value = value;
            old_value.name = current_name;
            ValueEntry new_value = value;
            new_value.name = current_name;
            new_value.data = new_data;
            ReplacePayload::Change change;
            change.undo.type =
                changes::UndoOperation::Type::kModifyValue;
            change.undo.node = node;
            change.undo.old_value = old_value;
            change.undo.new_value = new_value;
            change.history.action = L"Modify value " + current_name;
            change.history.old_data = value_format::Data(
                value.type, value.data.data(),
                static_cast<DWORD>(value.data.size()));
            change.history.new_data = value_format::Data(
                value.type, new_data.data(),
                static_cast<DWORD>(new_data.size()));
            change.history.key_path = registry_path::Build(node);
            change.history.value_name = current_name;
            change.history.revert_kind =
                HistoryEntry::RevertKind::kSetValue;
            change.history.revert_value = std::move(old_value);
            payload->changes.push_back(std::move(change));
          }

          if (options.recursive && !cancel.load()) {
            auto subkeys =
                RegistryStore::EnumSubKeyNames(node, false);
            for (const auto& name : subkeys) {
              stack.push_back(MakeChildNode(node, name));
            }
          }
        }

        payload->cancelled = cancel.load();
        if (hwnd && IsWindow(hwnd) &&
            PostMessageW(hwnd, frame::message_id::kReplaceReady,
                         static_cast<WPARAM>(generation),
                         reinterpret_cast<LPARAM>(payload.get()))) {
          ReleasePostedPayload(payload);
        }
      });
}

void MainWindow::Impl::ApplyReplacePayload(ReplacePayload* payload) {
  if (!payload) {
    return;
  }
  std::unique_ptr<ReplacePayload> owned(payload);
  if (!replace_session_.IsCurrent(owned->generation)) {
    return;
  }
  replace_session_.Join();
  replace_result_pending_ = false;
  CommitReplacePayload(std::move(owned), true);
}

void MainWindow::Impl::CommitReplacePayload(
    std::unique_ptr<ReplacePayload> payload, bool show_failures) {
  if (!payload) {
    return;
  }
  for (auto& change : payload->changes) {
    PushUndo(std::move(change.undo));
    AppendHistoryEntry(std::move(change.history));
  }
  if (!payload->changes.empty()) {
    MarkOfflineDirty();
  }
  if (browse_.current_node()) {
    UpdateValueListForNode(browse_.current_node());
  }
  if (show_failures && payload->failures > 0) {
    const std::wstring message =
        L"Replace finished with some failures.\nReplaced: " +
        std::to_wstring(payload->changes.size()) + L"\nFailed: " +
        std::to_wstring(payload->failures);
    ui::ShowError(hwnd_, message);
  }
}

void MainWindow::Impl::StopReplace() {
  replace_session_.CancelAndJoin();
  MSG message = {};
  while (PeekMessageW(&message, hwnd_, frame::message_id::kReplaceReady,
                      frame::message_id::kReplaceReady, PM_REMOVE)) {
    CommitReplacePayload(
        std::unique_ptr<ReplacePayload>(
            reinterpret_cast<ReplacePayload*>(message.lParam)),
        false);
  }
  replace_result_pending_ = false;
}

void MainWindow::Impl::CancelSearch() {
  search_session_.CancelAndJoin();
  search_running_ = false;
  search_start_tick_ = 0;
  search_duration_ms_ = 0;
  search_duration_valid_ = false;
  search_progress_percent_ = 0;
  search_progress_searched_.store(0);
  search_progress_total_.store(0);
  search_progress_posted_.store(false);
  search_posted_.store(false);
  {
    std::lock_guard<std::mutex> lock(search_mutex_);
    std::vector<PendingSearchResult>().swap(search_pending_);
  }
  if (search_progress_) {
    SendMessageW(search_progress_, PBM_SETMARQUEE, FALSE, 0);
  }
  ApplyViewVisibility();
  UpdateStatus();
}

void MainWindow::Impl::CloseSearchTab(int tab_index) {
  if (!tab_ || !IsSearchTabIndex(tab_index)) {
    return;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (tab_index >= count) {
    return;
  }
  if (search_running_ && active_search_tab_index_ == tab_index) {
    CancelSearch();
  }
  int search_index = SearchIndexFromTab(tab_index);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return;
  }

  bool was_active = TabCtrl_GetCurSel(tab_) == tab_index;

  search_tabs_.erase(search_tabs_.begin() + search_index);
  tabs_.erase(tabs_.begin() + tab_index);
  for (auto& entry : tabs_) {
    if (entry.kind == TabEntry::Kind::kSearch && entry.search_index > search_index) {
      --entry.search_index;
    }
  }
  TabCtrl_DeleteItem(tab_, tab_index);
  if (active_search_tab_index_ == tab_index) {
    active_search_tab_index_ = -1;
  } else if (active_search_tab_index_ > tab_index) {
    --active_search_tab_index_;
  }

  int new_count = TabCtrl_GetItemCount(tab_);
  if (was_active && new_count > 0) {
    int next = std::min(tab_index, new_count - 1);
    SelectTabIndex(next);
  }
  UpdateTabWidth();
  UpdateSearchResultsView();
  ApplyViewVisibility();
  UpdateStatus();
}

} // namespace regkit
