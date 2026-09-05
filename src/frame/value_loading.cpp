// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::StartValueListWorker() {
  if (value_loader_.running()) {
    return;
  }
  value_loader_.Start(
      [this](std::unique_ptr<ValueListTask> task,
             const std::atomic_bool& stopping) {
        if (!task || stopping.load() ||
            task->generation != value_list_generation_.load()) {
          return;
        }

        auto payload = std::make_unique<ValueListPayload>();
        payload->generation = task->generation;
        struct KeyMetadata {
          int image_index = kFolderIconIndex;
          bool is_link = false;
        };
        std::unordered_map<std::wstring, KeyMetadata> key_metadata_cache;
        key_metadata_cache.reserve(256);

        auto resolve_key_icon = [&](const RegistryNode& node, bool* is_link) -> int {
          if (is_link) {
            *is_link = false;
          }
          if (node.simulated) {
            return kFolderSimIconIndex;
          }
          std::wstring cache_key = registry_path::Build(node);
          auto cached = key_metadata_cache.find(cache_key);
          if (cached != key_metadata_cache.end()) {
            if (is_link) {
              *is_link = cached->second.is_link;
            }
            return cached->second.image_index;
          }
          KeyMetadata metadata;
          std::wstring nt_path = registry_path::BuildNative(node);
          if (!nt_path.empty() && task->hive_roots && task->hive_roots->find(ToLower(nt_path)) != task->hive_roots->end()) {
            metadata.image_index = kDatabaseIconIndex;
            key_metadata_cache.emplace(std::move(cache_key), metadata);
            return metadata.image_index;
          }
          std::wstring link_target;
          if (RegistryStore::QuerySymbolicLinkTarget(node, &link_target)) {
            if (is_link) {
              *is_link = true;
            }
            metadata.image_index = kSymlinkIconIndex;
            metadata.is_link = true;
            key_metadata_cache.emplace(std::move(cache_key), metadata);
            return metadata.image_index;
          }
          key_metadata_cache.emplace(std::move(cache_key), metadata);
          return metadata.image_index;
        };

        auto subkeys = RegistryStore::EnumSubKeyNames(task->snapshot, false);
        std::unordered_set<std::wstring> existing_keys;
        existing_keys.reserve(subkeys.size());
        for (const auto& name : subkeys) {
          existing_keys.insert(ToLower(name));
        }
        std::vector<std::wstring> simulated_subkeys;
        auto append_trace_children = [&](const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) {
          if (!out) {
            return;
          }
          out->clear();
          if (!task->show_simulated_keys) {
            return;
          }
          if (task->trace_data_list.empty()) {
            return;
          }
          if (!node.root_name.empty() && EqualsInsensitive(node.root_name, L"REGISTRY")) {
            return;
          }
          std::wstring path = registry_path::Build(node);
          std::wstring trace_path = NormalizeTraceKeyPath(path);
          if (trace_path.empty()) {
            trace_path = path;
          }
          std::wstring key_lower = ToLower(trace_path);
          std::unordered_set<std::wstring> seen;
          for (const auto& trace : task->trace_data_list) {
            if (!trace.data) {
              continue;
            }
            std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
            if (!trace.selection ||
                !trace::IncludesKey(*trace.selection, key_lower)) {
              continue;
            }
            auto it = trace.data->children_by_key.find(key_lower);
            if (it == trace.data->children_by_key.end()) {
              continue;
            }
            for (const auto& name : it->second) {
              if (name.empty()) {
                continue;
              }
              std::wstring name_lower = ToLower(name);
              if (existing_lower.find(name_lower) != existing_lower.end()) {
                continue;
              }
              if (!seen.insert(name_lower).second) {
                continue;
              }
              out->push_back(name);
            }
          }
          std::sort(out->begin(), out->end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
        };
        append_trace_children(task->snapshot, existing_keys, &simulated_subkeys);
        payload->key_count = static_cast<int>(subkeys.size() + simulated_subkeys.size());
        payload->rows.reserve((task->show_keys_in_list ? subkeys.size() + simulated_subkeys.size() : 0) + 16);

        struct TraceMatch {
          std::wstring label;
          trace::KeyValues values;
          const trace::Selection* selection = nullptr;
        };
        std::vector<TraceMatch> trace_matches;
        if (!task->trace_data_list.empty()) {
          for (const auto& trace : task->trace_data_list) {
            if (!trace.data) {
              continue;
            }
            std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
            if (!trace.selection ||
                !trace::IncludesKey(*trace.selection,
                                    task->trace_path_lower)) {
              continue;
            }
            auto it = trace.data->values_by_key.find(task->trace_path_lower);
            if (it == trace.data->values_by_key.end()) {
              continue;
            }
            TraceMatch match;
            match.label = trace.label.empty() ? L"Trace" : trace.label;
            match.values = it->second;
            match.selection = trace.selection.get();
            trace_matches.push_back(std::move(match));
          }
        }

        struct DefaultMatch {
          std::wstring label;
          defaults::Key values;
          const trace::Selection* selection = nullptr;
        };
        std::vector<DefaultMatch> default_keys;
        if (!task->default_data_list.empty() && !task->default_path_lower.empty()) {
          default_keys.reserve(task->default_data_list.size());
          for (const auto& defaults : task->default_data_list) {
            if (!defaults.data) {
              continue;
            }
            std::shared_lock<std::shared_mutex> defaults_lock(*defaults.data->mutex);
            if (!defaults.selection ||
                !trace::IncludesKey(*defaults.selection,
                                    task->default_path_lower)) {
              continue;
            }
            auto it = defaults.data->values_by_key.find(task->default_path_lower);
            if (it == defaults.data->values_by_key.end()) {
              continue;
            }
            default_keys.push_back(
                {ShortDefaultLabel(defaults.label, defaults.source_path),
                 it->second, defaults.selection.get()});
          }
        }
        auto resolve_default_data = [&](const std::wstring& value_name) -> std::wstring {
          if (default_keys.empty()) {
            return {};
          }
          std::wstring value_lower = ToLower(value_name);
          struct DefaultGroup {
            std::wstring text;
            std::wstring labels;
          };
          std::vector<DefaultGroup> groups;
          for (const auto& match : default_keys) {
            if (match.selection &&
                !trace::IncludesValue(*match.selection,
                                      task->default_path_lower,
                                      value_lower)) {
              continue;
            }
            auto it = match.values.values.find(value_lower);
            std::wstring text = it != match.values.values.end()
                                    ? it->second.data
                                    : std::wstring(L"(Missing)");
            auto group = std::find_if(
                groups.begin(), groups.end(),
                [&](const DefaultGroup& entry) { return entry.text == text; });
            if (group == groups.end()) {
              groups.push_back({std::move(text), match.label});
              continue;
            }
            if (!match.label.empty()) {
              if (!group->labels.empty()) {
                group->labels.append(L", ");
              }
              group->labels.append(match.label);
            }
          }
          if (groups.empty()) {
            return {};
          }
          if (groups.size() == 1) {
            return std::move(groups.front().text);
          }
          std::wstring text;
          for (size_t i = 0; i < groups.size(); ++i) {
            if (i > 0) {
              text.append(L" - ");
            }
            text.append(groups[i].text);
            if (!groups[i].labels.empty()) {
              text.append(L" (").append(groups[i].labels).append(L")");
            }
          }
          return text;
        };

        if (task->show_keys_in_list) {
          for (const auto& name : subkeys) {
            ListRow row;
            row.name = name;
            bool is_link = false;
            RegistryNode child = task->snapshot;
            child.subkey = task->snapshot.subkey.empty() ? name : task->snapshot.subkey + L"\\" + name;
            row.image_index = resolve_key_icon(child, &is_link);
            row.type = is_link ? L"Link" : L"Key";
            row.extra = name;
            row.kind = rowkind::kKey;
            if (task->include_dates || task->include_details) {
              KeyInfo info = {};
              if (RegistryStore::QueryKeyInfo(child, &info)) {
                if (task->include_dates) {
                  row.date = FormatFileTime(info.last_write);
                  row.date_value = FileTimeToUint64(info.last_write);
                  row.has_date = (row.date_value != 0);
                }
                if (task->include_details) {
                  row.detail_key_count = info.subkey_count;
                  row.detail_value_count = info.value_count;
                  row.has_details = true;
                  row.details = L"Keys: " + std::to_wstring(info.subkey_count) + L", Values: " + std::to_wstring(info.value_count);
                }
              }
            }
            payload->rows.emplace_back(std::move(row));
          }
          for (const auto& name : simulated_subkeys) {
            if (name.empty()) {
              continue;
            }
            ListRow row;
            row.name = name;
            RegistryNode child = task->snapshot;
            child.subkey = task->snapshot.subkey.empty() ? name : task->snapshot.subkey + L"\\" + name;
            child.simulated = true;
            row.image_index = resolve_key_icon(child, nullptr);
            row.simulated = true;
            row.type = L"Key";
            row.extra = name;
            row.kind = rowkind::kKey;
            payload->rows.emplace_back(std::move(row));
          }
        }

        std::wstring link_target;
        bool has_link = RegistryStore::QuerySymbolicLinkTarget(task->snapshot, &link_target);
        bool track_existing = (!trace_matches.empty()) || has_link;
        std::unordered_set<std::wstring> existing_values;
        if (track_existing) {
          existing_values.reserve(64);
        }

        auto gather_labels = [&](const std::wstring& value_lower) -> std::vector<std::wstring> {
          std::vector<std::wstring> labels;
          for (const auto& match : trace_matches) {
            if (match.values.values_lower.find(value_lower) != match.values.values_lower.end()) {
              if (match.selection &&
                  !trace::IncludesValue(*match.selection,
                                        task->trace_path_lower,
                                        value_lower)) {
                continue;
              }
              labels.push_back(match.label);
            }
          }
          if (labels.size() < 2) {
            return labels;
          }
          std::vector<std::wstring> unique;
          unique.reserve(labels.size());
          std::unordered_set<std::wstring> seen;
          for (const auto& label : labels) {
            std::wstring key = ToLower(label);
            if (seen.insert(key).second) {
              unique.push_back(label);
            }
          }
          return unique;
        };
        auto format_read_on_boot = [&](const std::vector<std::wstring>& labels) -> std::wstring {
          if (labels.empty()) {
            return L"No";
          }
          std::wstring out = L"Yes (";
          for (size_t i = 0; i < labels.size(); ++i) {
            if (i > 0) {
              out.append(L", ");
            }
            out.append(labels[i]);
          }
          out.push_back(L')');
          return out;
        };
        bool have_traces = !trace_matches.empty();

        bool has_default = false;
        bool has_symbolic_value = false;
        constexpr DWORD kEagerValueDataLimit = 64u * 1024u;
        const DWORD max_data_size = task->include_all_value_data ? MAXDWORD : kEagerValueDataLimit;
        auto append_value = [&](const ValueInfo& value, const BYTE* data, DWORD data_size) -> bool {
          if (value.name.empty()) {
            has_default = true;
          }
          if (EqualsInsensitive(value.name, L"SymbolicLinkValue")) {
            has_symbolic_value = true;
          }
          ListRow row = MakeValueListRow(value.name, value.type, data,
                                         data_size);
          row.default_data = resolve_default_data(value.name);
          if (!have_traces) {
            row.read_on_boot.clear();
          } else {
            std::wstring lower = ToLower(value.name);
            row.read_on_boot = format_read_on_boot(gather_labels(lower));
            if (track_existing) {
              existing_values.insert(lower);
            }
          }
          payload->rows.emplace_back(std::move(row));
          ++payload->value_count;
          return !stopping.load() &&
                 task->generation == value_list_generation_.load();
        };
        RegistryStore::EnumKeyStreaming(task->snapshot, true, true, false, nullptr, append_value, RegistryStore::SubkeyStreamCallback(), max_data_size);

        if (!has_symbolic_value && has_link && !link_target.empty()) {
          ListRow row;
          row.name = L"SymbolicLinkValue";
          row.type = L"REG_LINK";
          row.data = link_target;
          row.data_ready = true;
          row.image_index = UseBinaryValueIcon(REG_LINK) ? kBinaryIconIndex : kValueIconIndex;
          row.kind = rowkind::kValue;
          row.extra = L"SymbolicLinkValue";
          row.default_data = resolve_default_data(row.extra);
          DWORD link_bytes = static_cast<DWORD>((link_target.size() + 1) * sizeof(wchar_t));
          row.size = std::to_wstring(link_bytes);
          row.size_value = link_bytes;
          row.value_data_size = link_bytes;
          row.has_size = true;
          row.value_type = REG_LINK;
          row.read_on_boot = have_traces ? L"No" : L"";
          row.simulated = true;
          payload->rows.emplace_back(std::move(row));
        }

        if (!has_default) {
          ListRow row;
          row.name = L"(Default)";
          row.type = L"REG_SZ";
          row.data = L"(value not set)";
          row.data_ready = true;
          row.image_index = kValueIconIndex;
          row.kind = rowkind::kValue;
          row.extra = L"";
          row.default_data = resolve_default_data(row.extra);
          row.size = L"0";
          row.size_value = 0;
          row.has_size = true;
          row.value_type = REG_SZ;
          if (!have_traces) {
            row.read_on_boot.clear();
          } else {
            row.read_on_boot = format_read_on_boot(gather_labels(L""));
            if (track_existing) {
              existing_values.insert(L"");
            }
          }
          payload->rows.emplace_back(std::move(row));
          payload->value_count += 1;
        }

        size_t trace_added = 0;
        if (!trace_matches.empty()) {
          for (const auto& match : trace_matches) {
            payload->rows.reserve(payload->rows.size() + match.values.values_display.size());
            for (const auto& value_name : match.values.values_display) {
              std::wstring value_lower = ToLower(value_name);
              if (match.selection &&
                  !trace::IncludesValue(*match.selection,
                                        task->trace_path_lower,
                                        value_lower)) {
                continue;
              }
              if (existing_values.find(value_lower) != existing_values.end()) {
                continue;
              }
              ListRow row;
              row.name = value_name.empty() ? L"(Default)" : value_name;
              row.type = L"TRACE";
              row.data = L"(value not set)";
              row.read_on_boot = format_read_on_boot(gather_labels(value_lower));
              row.image_index = kValueIconIndex;
              row.kind = rowkind::kValue;
              row.extra = value_name;
              row.data_ready = true;
              row.default_data = resolve_default_data(value_name);
              payload->rows.emplace_back(std::move(row));
              ++trace_added;
              existing_values.insert(value_lower);
            }
          }
          payload->value_count += static_cast<int>(trace_added);
        }

        if (task->sort_column != kValueColComment) {
          SortValueRows(&payload->rows, task->sort_column, task->sort_ascending);
        }
        if (stopping.load() ||
            task->generation != value_list_generation_.load()) {
          return;
        }
        if (PostMessageW(task->hwnd, frame::message_id::kValueListReady, static_cast<WPARAM>(task->generation), reinterpret_cast<LPARAM>(payload.get())) != 0) {
          ReleasePostedPayload(payload);
        }
      });
}

void MainWindow::Impl::StartValuePreviewWorker() {
  if (value_preview_loader_.running()) {
    return;
  }
  value_preview_loader_.Start(
      [this](std::unique_ptr<ValuePreviewTask> task, const std::atomic_bool& stopping) {
        if (!task || task->generation != value_list_generation_.load()) {
          return;
        }
        auto payload = std::make_unique<ValuePreviewPayload>();
        payload->generation = task->generation;
        payload->items.reserve(task->indices.size());
        for (size_t i = 0; i < task->indices.size(); ++i) {
          if (stopping.load() || task->generation != value_list_generation_.load()) {
            return;
          }
          ValueEntry entry;
          ValuePreviewItem item;
          item.index = task->indices[i];
          item.name = task->names[i];
          const bool ok =
              RegistryStore::QueryValue(task->snapshot, item.name, &entry);
          if (ok) {
            item.type = entry.type;
            item.size = static_cast<DWORD>(entry.data.size());
            item.preview = value_format::DisplayData(
                entry.type, entry.data.data(),
                std::min<DWORD>(item.size, kValuePreviewBytes));
            if (item.preview.size() > kValuePreviewLimit) {
              item.preview.resize(kValuePreviewLimit);
            }
          }
          payload->items.push_back(std::move(item));
        }
        if (stopping.load() || task->generation != value_list_generation_.load()) {
          return;
        }
        if (PostMessageW(task->hwnd, frame::message_id::kValuePreviewReady,
                         static_cast<WPARAM>(task->generation),
                         reinterpret_cast<LPARAM>(payload.get())) != 0) {
          payload.release();
        }
      });
}

void MainWindow::Impl::QueueValuePreviews(int first, int last) {
  if (!browse_.current_node() || first < 0 || last < first) {
    return;
  }
  auto task = std::make_unique<ValuePreviewTask>();
  for (int i = first; i <= last; ++i) {
    ListRow* row = browse_.values().MutableRowAt(i);
    if (!row || row->kind != rowkind::kValue || row->data_ready) {
      continue;
    }
    if (row->value_data_size == 0) {
      row->data.clear();
      row->data_ready = true;
      browse_.values().InvalidateFilterCache(row);
      continue;
    }
    task->indices.push_back(i);
    task->names.push_back(row->extra);
  }
  if (task->indices.empty()) {
    return;
  }
  StartValuePreviewWorker();
  task->generation = value_list_generation_.load();
  task->snapshot = *browse_.current_node();
  task->hwnd = hwnd_;
  value_preview_loader_.Submit(std::move(task));
}

void MainWindow::Impl::StartSearchPreviewWorker() {
  if (search_preview_loader_.running()) {
    return;
  }
  search_preview_loader_.Start(
      [this](std::unique_ptr<SearchPreviewTask> task,
             const std::atomic_bool& stopping) {
        if (!task) {
          return;
        }
        auto payload = std::make_unique<SearchPreviewPayload>();
        payload->generation = task->generation;
        payload->tab_index = task->tab_index;
        payload->items.reserve(task->requests.size());
        for (const auto& request : task->requests) {
          if (stopping.load()) {
            return;
          }
          ValueEntry entry;
          SearchPreviewItem item;
          item.index = request.index;
          item.row_id = request.row_id;
          if (RegistryStore::QueryValue(request.node, request.value_name, &entry)) {
            item.type = entry.type;
            item.data_size = static_cast<DWORD>(entry.data.size());
            item.preview = value_format::DisplayData(
                entry.type, entry.data.data(),
                std::min<DWORD>(item.data_size, kValuePreviewBytes));
            if (item.preview.size() > kValuePreviewLimit) {
              item.preview.resize(kValuePreviewLimit);
            }
          }
          payload->items.push_back(std::move(item));
        }
        if (stopping.load()) {
          return;
        }
        if (PostMessageW(task->hwnd, frame::message_id::kSearchPreviewReady,
                         static_cast<WPARAM>(task->generation),
                         reinterpret_cast<LPARAM>(payload.get())) != 0) {
          payload.release();
        }
      });
}

void MainWindow::Impl::QueueSearchPreviews(int first, int last) {
  const int index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
  if (index < 0 || static_cast<size_t>(index) >= search_tabs_.size()) {
    return;
  }
  SearchTab& tab = search_tabs_[static_cast<size_t>(index)];
  if (tab.is_compare || first < 0 || last < first) {
    return;
  }
  auto task = std::make_unique<SearchPreviewTask>();
  const size_t count = tab.results.size();
  for (int i = first; i <= last && static_cast<size_t>(i) < count; ++i) {
    search::Result& row = tab.results[static_cast<size_t>(i)];
    if (row.data_state != search::DataState::kNotLoaded ||
        search::IsKeyRow(row)) {
      continue;
    }
    SearchPreviewRequest request;
    if (!ResolvePathToNode(row.key_path, &request.node)) {
      row.data_text.clear();
      row.data_state = search::DataState::kLoaded;
      continue;
    }
    request.index = i;
    request.row_id = row.row_id;
    request.value_name = row.value_name;
    task->requests.push_back(std::move(request));
  }
  if (task->requests.empty()) {
    return;
  }
  StartSearchPreviewWorker();
  task->generation = tab.generation;
  task->tab_index = index;
  task->hwnd = hwnd_;
  search_preview_loader_.Submit(std::move(task));
}

void MainWindow::Impl::StartSearchSortWorker() {
  if (search_sort_loader_.running()) {
    return;
  }
  search_sort_loader_.Start(
      [this](std::unique_ptr<SearchSortTask> task,
             const std::atomic_bool& stopping) {
        if (!task) {
          return;
        }
        for (size_t i = 0; i < task->rows.size(); ++i) {
          if (stopping.load()) {
            return;
          }
          search::Result& row = task->rows[i];
          if (row.data_state != search::DataState::kNotLoaded) {
            continue;
          }
          row.data_state = search::DataState::kLoaded;
          ValueEntry entry;
          if (!RegistryStore::QueryValue(task->nodes[i], row.value_name,
                                         &entry)) {
            row.data_text.clear();
            continue;
          }
          row.type = entry.type;
          row.data_size = static_cast<DWORD>(entry.data.size());
          std::wstring preview = value_format::DisplayData(
              entry.type, entry.data.data(),
              std::min<DWORD>(row.data_size, kValuePreviewBytes));
          if (preview.size() > kValuePreviewLimit) {
            preview.resize(kValuePreviewLimit);
          }
          row.data_text = std::move(preview);
        }
        if (stopping.load()) {
          return;
        }
        search::SortResults(&task->rows, task->column, task->ascending);
        auto payload = std::make_unique<SearchSortPayload>();
        payload->generation = task->generation;
        payload->tab_index = task->tab_index;
        payload->rows = std::move(task->rows);
        if (PostMessageW(task->hwnd, frame::message_id::kSearchSortReady,
                         static_cast<WPARAM>(task->generation),
                         reinterpret_cast<LPARAM>(payload.get())) != 0) {
          payload.release();
        }
      });
}

void MainWindow::Impl::QueueSearchSort(SearchTab* tab) {
  if (!tab) {
    return;
  }
  int tab_index = -1;
  for (size_t i = 0; i < search_tabs_.size(); ++i) {
    if (&search_tabs_[i] == tab) {
      tab_index = static_cast<int>(i);
      break;
    }
  }
  if (tab_index < 0) {
    return;
  }
  auto task = std::make_unique<SearchSortTask>();
  task->rows = tab->results;
  task->nodes.resize(task->rows.size());
  for (size_t i = 0; i < task->rows.size(); ++i) {
    if (task->rows[i].data_state == search::DataState::kNotLoaded) {
      RegistryNode node;
      if (ResolvePathToNode(task->rows[i].key_path, &node)) {
        task->nodes[i] = std::move(node);
      } else {
        task->rows[i].data_state = search::DataState::kLoaded;
      }
    }
  }
  StartSearchSortWorker();
  task->generation = tab->generation;
  task->tab_index = tab_index;
  task->column = tab->sort_column;
  task->ascending = tab->sort_ascending;
  task->hwnd = hwnd_;
  search_sort_loader_.Submit(std::move(task));
}

void MainWindow::Impl::StartSearchTabLoadWorker() {
  if (search_tab_loader_.running()) {
    return;
  }
  search_tab_loader_.Start(
      [](std::unique_ptr<SearchTabLoadTask> task,
         const std::atomic_bool& stopping) {
        if (!task) {
          return;
        }
        auto payload = std::make_unique<SearchTabLoadPayload>();
        payload->generation = task->generation;
        payload->tab_index = task->tab_index;

        payload->ok = search::LoadResults(task->path, &payload->rows);
        if (payload->ok && task->sort_column >= 0) {
          search::SortResults(&payload->rows, task->sort_column,
                              task->sort_ascending);
        }
        if (stopping.load()) {
          return;
        }
        if (PostMessageW(task->hwnd, frame::message_id::kSearchTabLoadReady,
                         static_cast<WPARAM>(task->generation),
                         reinterpret_cast<LPARAM>(payload.get())) != 0) {
          payload.release();
        }
      });
}

void MainWindow::Impl::StopValueListWorker() {
  value_loader_.Stop();
  value_preview_loader_.Stop();
  search_preview_loader_.Stop();
  search_sort_loader_.Stop();
  search_tab_loader_.Stop();
}

} // namespace regkit
