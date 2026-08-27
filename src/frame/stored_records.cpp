// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::StartDefaultLoadWorker() {
  if (default_load_session_.running()) {
    return;
  }
  std::wstring active_path = ActiveDefaultsPath();
  const HWND hwnd = hwnd_;
  default_load_session_.StartIfIdle(
      [this, active_path, hwnd](uint64_t generation,
                                const std::atomic_bool& cancel) {
        auto payload = std::make_unique<DefaultLoadPayload>();
        payload->generation = generation;
        std::wstring content;
        if (!util::ReadTextFile(
                active_path, &content, nullptr,
                static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
          return;
        }

        std::vector<std::wstring> entries;
        size_t start = 0;
        while (start < content.size()) {
          size_t end = content.find(L'\n', start);
          if (end == std::wstring::npos) {
            end = content.size();
          }
          std::wstring line = content.substr(start, end - start);
          if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
          }
          start = end + 1;
          line = TrimWhitespace(line);
          if (line.empty() || line.front() == L'#') {
            continue;
          }
          if (StartsWithInsensitive(line, L"default=")) {
            line.erase(0, wcslen(L"default="));
          }
          line = TrimWhitespace(line);
          if (!line.empty()) {
            entries.push_back(std::move(line));
          }
        }

        std::unordered_set<std::wstring> loaded;
        for (const auto& entry : entries) {
          if (cancel.load()) {
            return;
          }
          std::wstring source = entry;
          std::wstring use_label;
          if (!FileExists(source)) {
            std::wstring bundled = ResolveBundledDefaultPath(source);
            if (!bundled.empty() && FileExists(bundled)) {
              source = bundled;
              use_label = entry;
            } else {
              continue;
            }
          }
          if (use_label.empty()) {
            use_label = FileBaseName(source);
          }
          if (use_label.empty()) {
            use_label = L"Default";
          }
          std::wstring source_lower = ToLower(source);
          if (!loaded.insert(source_lower).second) {
            continue;
          }
          defaults::Data data;
          if (!defaults::Load(
                  source,
                  [](const std::wstring& path) {
                    return NormalizeTraceKeyPathBasic(path);
                  },
                  &data, nullptr, nullptr,
                  &cancel)) {
            continue;
          }
          std::shared_ptr<const defaults::Data> default_data =
              std::make_shared<defaults::Data>(std::move(data));
          trace::Selection selection = {};
          selection.select_all = true;
          selection.recursive = true;
          payload->defaults.push_back(
              {use_label, source, default_data,
               std::make_shared<trace::Selection>(selection)});
        }

        if (cancel.load()) {
          return;
        }
        if (hwnd && IsWindow(hwnd) &&
            PostMessageW(hwnd, frame::message_id::kDefaultLoadReady, 0,
                         reinterpret_cast<LPARAM>(payload.get()))) {
          ReleasePostedPayload(payload);
        }
      });
}

void MainWindow::Impl::StopDefaultLoadWorker() {
  default_load_session_.CancelAndJoin();
}

void MainWindow::Impl::StopDefaultParseSessions() {
  for (auto& entry : default_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->work.CancelAndJoin();
  }
  default_parse_sessions_.clear();
}

void MainWindow::Impl::StopRegFileParseSessions() {
  for (auto& entry : reg_file_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->work.CancelAndJoin();
  }
  reg_file_parse_sessions_.clear();
}

void MainWindow::Impl::MarkTreeStateDirty() {
  if (!save_tree_state_ || !browse_.tree().hwnd() || !IsWindow(browse_.tree().hwnd())) {
    return;
  }
  std::wstring selected;
  std::vector<std::wstring> expanded;
  CaptureTreeState(&selected, &expanded);
  workspace::TreeState state;
  state.selected_path = std::move(selected);
  state.expanded_paths = std::move(expanded);
  tree_state_saver_.Submit(std::move(state));
}

void MainWindow::Impl::SaveTreeStateFile(const std::wstring& selected, const std::vector<std::wstring>& expanded) const {
  workspace::TreeState state;
  state.selected_path = selected;
  state.expanded_paths = expanded;
  workspace::SaveTreeState(TreeStatePath(), state);
}

std::wstring MainWindow::Impl::ActiveTracesPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"active_traces.ini");
}

std::wstring MainWindow::Impl::ActiveDefaultsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"active_defaults.ini");
}

std::wstring MainWindow::Impl::TraceSettingsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"trace_settings.ini");
}

void MainWindow::Impl::LoadTraceSettings() {
  trace_selection_cache_.clear();
  std::wstring path = TraceSettingsPath();
  if (path.empty()) {
    return;
  }
  auto parse_bool = [](const std::wstring& value) -> bool { return (_wcsicmp(value.c_str(), L"1") == 0 || _wcsicmp(value.c_str(), L"true") == 0 || _wcsicmp(value.c_str(), L"yes") == 0); };
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, MutableData(buffer), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(buffer);
  if (content.empty()) {
    return;
  }

  trace::Selection selection = {};
  selection.select_all = true;
  selection.recursive = true;
  std::wstring current_path;
  std::wstring current_label;
  bool has_entry = false;

  auto normalize_selection = [&]() {
    std::vector<std::wstring> cleaned;
    cleaned.reserve(selection.key_paths.size());
    std::unordered_set<std::wstring> seen;
    for (auto& path : selection.key_paths) {
      std::wstring trimmed = TrimWhitespace(path);
      if (trimmed.empty()) {
        continue;
      }
      std::wstring lower = ToLower(trimmed);
      if (seen.insert(lower).second) {
        cleaned.push_back(std::move(trimmed));
      }
    }
    for (const auto& entry : selection.values_by_key) {
      if (entry.first.empty()) {
        continue;
      }
      if (seen.insert(entry.first).second) {
        cleaned.push_back(entry.first);
      }
    }
    selection.key_paths.swap(cleaned);
    if (selection.key_paths.empty() && selection.values_by_key.empty()) {
      selection.select_all = true;
    }
  };

  auto flush_entry = [&]() {
    if (!has_entry) {
      return;
    }
    normalize_selection();
    std::wstring key = current_path;
    if (key.empty() && !current_label.empty()) {
      std::wstring resolved = ResolveBundledTracePath(current_label);
      key = resolved.empty() ? current_label : resolved;
    }
    key = TrimWhitespace(key);
    if (!key.empty()) {
      trace_selection_cache_[ToLower(key)] = selection;
    }
    selection = {};
    selection.select_all = true;
    selection.recursive = true;
    current_path.clear();
    current_label.clear();
    has_entry = false;
  };

  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    line = TrimWhitespace(line);
    if (line.empty()) {
      flush_entry();
      continue;
    }
    if (line.front() == L'#') {
      continue;
    }
    if (line.front() == L'[') {
      flush_entry();
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = TrimWhitespace(line.substr(0, sep));
    std::wstring value = line.substr(sep + 1);
    if (key.empty()) {
      continue;
    }
    has_entry = true;
    if (EqualsInsensitive(key, L"path")) {
      current_path = value;
    } else if (EqualsInsensitive(key, L"label")) {
      current_label = value;
    } else if (EqualsInsensitive(key, L"select_all")) {
      selection.select_all = parse_bool(value);
    } else if (EqualsInsensitive(key, L"recursive")) {
      selection.recursive = parse_bool(value);
    } else if (EqualsInsensitive(key, L"key_path") || EqualsInsensitive(key, L"key")) {
      selection.key_paths.push_back(value);
    } else if (EqualsInsensitive(key, L"value")) {
      size_t bar = value.find(L'|');
      if (bar == std::wstring::npos) {
        continue;
      }
      std::wstring key_part = TrimWhitespace(value.substr(0, bar));
      std::wstring value_part = TrimWhitespace(value.substr(bar + 1));
      if (key_part.empty()) {
        continue;
      }
      if (value_part == L"@") {
        value_part.clear();
      }
      std::wstring key_lower = ToLower(key_part);
      std::wstring value_lower = ToLower(value_part);
      selection.values_by_key[key_lower].insert(value_lower);
    }
  }
  flush_entry();
}

void MainWindow::Impl::SaveTraceSettings() const {
  std::wstring path = TraceSettingsPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& trace : active_traces_) {
    if (!trace.data || !trace.selection) {
      continue;
    }
    content.append(L"[trace]\n");
    if (!trace.label.empty()) {
      content.append(L"label=");
      content.append(trace.label);
      content.push_back(L'\n');
    }
    if (!trace.source_path.empty()) {
      content.append(L"path=");
      content.append(trace.source_path);
      content.push_back(L'\n');
    }
    content.append(L"select_all=");
    content.append(trace.selection->select_all ? L"1\n" : L"0\n");
    content.append(L"recursive=");
    content.append(trace.selection->recursive ? L"1\n" : L"0\n");
    for (const auto& key_path : trace.selection->key_paths) {
      if (key_path.empty()) {
        continue;
      }
      content.append(L"key=");
      content.append(key_path);
      content.push_back(L'\n');
    }
    for (const auto& entry : trace.selection->values_by_key) {
      if (entry.first.empty()) {
        continue;
      }
      for (const auto& value_name : entry.second) {
        content.append(L"value=");
        content.append(entry.first);
        content.push_back(L'|');
        if (value_name.empty()) {
          content.append(L"@");
        } else {
          content.append(value_name);
        }
        content.push_back(L'\n');
      }
    }
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  if (!utf8.empty()) {
    DWORD written = 0;
    WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  }
  CloseHandle(file);
}

void MainWindow::Impl::SaveActiveTraces() const {
  std::wstring path = ActiveTracesPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& trace : active_traces_) {
    if (trace.source_path.empty()) {
      continue;
    }
    std::wstring entry = trace.source_path;
    if (!trace.label.empty()) {
      std::wstring bundled = ResolveBundledTracePath(trace.label);
      if (!bundled.empty() && EqualsInsensitive(bundled, trace.source_path)) {
        entry = trace.label;
      }
    }
    content.append(entry);
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  if (utf8.empty() && !content.empty()) {
    return;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  if (!utf8.empty()) {
    DWORD written = 0;
    WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  }
  CloseHandle(file);
}

void MainWindow::Impl::SaveActiveDefaults() const {
  std::wstring path = ActiveDefaultsPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& defaults : active_defaults_) {
    if (defaults.source_path.empty()) {
      continue;
    }
    std::wstring entry = defaults.source_path;
    if (!defaults.label.empty()) {
      std::wstring bundled = ResolveBundledDefaultPath(defaults.label);
      if (!bundled.empty() && EqualsInsensitive(bundled, defaults.source_path)) {
        entry = defaults.label;
      }
    }
    content.append(entry);
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  if (utf8.empty() && !content.empty()) {
    return;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  if (!utf8.empty()) {
    DWORD written = 0;
    WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  }
  CloseHandle(file);
}

bool MainWindow::Impl::HasActiveTraces() const {
  return !active_traces_.empty();
}

bool MainWindow::Impl::RemoveTraceByPath(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::wstring target = TrimWhitespace(path);
  if (target.empty()) {
    return false;
  }
  std::wstring target_lower = ToLower(target);
  auto session_it = trace_parse_sessions_.find(target_lower);
  if (session_it != trace_parse_sessions_.end()) {
    if (session_it->second) {
      session_it->second->work.CancelAndJoin();
    }
    trace_parse_sessions_.erase(session_it);
  }
  size_t removed = 0;
  active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(),
                                      [&](const ActiveTrace& trace) {
                                        if (!EqualsInsensitive(trace.source_path, target)) {
                                          return false;
                                        }
                                        ++removed;
                                        return true;
                                      }),
                       active_traces_.end());
  if (removed == 0) {
    return false;
  }
  trace_selection_cache_.erase(target_lower);
  SaveActiveTraces();
  SaveTraceSettings();
  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(browse_.current_node());
  SaveSettings();
  return true;
}

bool MainWindow::Impl::RemoveTraceByLabel(const std::wstring& label) {
  if (label.empty()) {
    return false;
  }
  for (auto it = trace_parse_sessions_.begin(); it != trace_parse_sessions_.end();) {
    if (it->second && _wcsicmp(it->second->label.c_str(), label.c_str()) == 0) {
      it->second->work.CancelAndJoin();
      it = trace_parse_sessions_.erase(it);
      continue;
    }
    ++it;
  }
  size_t removed = 0;
  active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(),
                                      [&](const ActiveTrace& trace) {
                                        if (_wcsicmp(trace.label.c_str(), label.c_str()) != 0) {
                                          return false;
                                        }
                                        ++removed;
                                        return true;
                                      }),
                       active_traces_.end());
  if (removed == 0) {
    return false;
  }
  trace_selection_cache_.clear();
  for (const auto& trace : active_traces_) {
    if (!trace.source_path.empty()) {
      if (trace.selection) {
        trace_selection_cache_[ToLower(trace.source_path)] = *trace.selection;
      }
    }
  }
  SaveActiveTraces();
  SaveTraceSettings();
  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(browse_.current_node());
  SaveSettings();
  return true;
}

bool MainWindow::Impl::RemoveDefaultByPath(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::wstring target = TrimWhitespace(path);
  if (target.empty()) {
    return false;
  }
  std::wstring target_lower = ToLower(target);
  auto session_it = default_parse_sessions_.find(target_lower);
  if (session_it != default_parse_sessions_.end()) {
    if (session_it->second) {
      session_it->second->work.CancelAndJoin();
    }
    default_parse_sessions_.erase(session_it);
  }
  size_t removed = 0;
  active_defaults_.erase(std::remove_if(active_defaults_.begin(), active_defaults_.end(),
                                        [&](const ActiveDefault& defaults) {
                                          if (!EqualsInsensitive(defaults.source_path, target)) {
                                            return false;
                                          }
                                          ++removed;
                                          return true;
                                        }),
                         active_defaults_.end());
  if (removed == 0) {
    return false;
  }
  SaveActiveDefaults();
  BuildMenus();
  UpdateValueListForNode(browse_.current_node());
  SaveSettings();
  return true;
}

} // namespace regkit
