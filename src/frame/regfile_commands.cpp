// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

bool MainWindow::Impl::SaveRegFileTab(int tab_index) {
  if (!IsRegFileTabIndex(tab_index) || static_cast<size_t>(tab_index) >= tabs_.size()) {
    return false;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
  if (entry.reg_file_path.empty()) {
    return false;
  }
  std::wstring content;
  if (!BuildRegFileContent(entry, &content)) {
    return false;
  }
  if (!util::WriteTextFile(entry.reg_file_path, content, true)) {
    ui::ShowError(hwnd_, L"Failed to save registry file.");
    return false;
  }
  if (entry.reg_file_dirty) {
    entry.reg_file_dirty = false;
    BuildMenus();
  }
  AppendHistoryEntry(L"Save .reg file " + FileNameOnly(entry.reg_file_path), L"", entry.reg_file_path);
  return true;
}

bool MainWindow::Impl::ExportRegFileTab(int tab_index, const std::wstring& path) {
  if (!IsRegFileTabIndex(tab_index) || static_cast<size_t>(tab_index) >= tabs_.size()) {
    return false;
  }
  if (path.empty()) {
    return false;
  }
  std::wstring content;
  if (!BuildRegFileContent(tabs_[static_cast<size_t>(tab_index)], &content)) {
    return false;
  }
  std::wstring target = EnsureRegExtension(path);
  if (!util::WriteTextFile(target, content, true)) {
    ui::ShowError(hwnd_, L"Failed to export registry file.");
    return false;
  }
  return true;
}

bool MainWindow::Impl::BuildRegFileContent(const TabEntry& entry, std::wstring* out) const {
  if (!out) {
    return false;
  }
  out->clear();
  if (entry.kind != TabEntry::Kind::kRegFile) {
    return false;
  }

  regfile::Writer writer;
  std::function<void(const VirtualRegistryKey&,
                     const std::wstring&)>
      append_key;
  append_key = [&](const VirtualRegistryKey& key,
                   const std::wstring& full_path) {
    if (!key.values.empty()) {
      std::vector<const VirtualRegistryValue*> values;
      values.reserve(key.values.size());
      for (const auto& source : key.values) {
        values.push_back(&source.second);
      }
      writer.AppendKey(full_path, std::move(values));
    }
    std::vector<const VirtualRegistryKey*> children;
    children.reserve(key.children.size());
    for (const auto& child : key.children) {
      if (child.second) {
        children.push_back(child.second.get());
      }
    }
    std::sort(children.begin(), children.end(), [](const VirtualRegistryKey* left, const VirtualRegistryKey* right) {
      if (!left || !right) {
        return left != nullptr;
      }
      return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
    });
    for (const auto* child : children) {
      if (!child) {
        continue;
      }
      append_key(*child, full_path + L"\\" + child->name);
    }
  };

  for (const auto& root : entry.reg_file_roots) {
    if (!root.data || !root.data->root) {
      continue;
    }
    std::wstring root_name = root.name;
    if (root_name.empty()) {
      root_name = root.data->root_name;
    }
    if (root_name.empty()) {
      continue;
    }
    append_key(*root.data->root, root_name);
  }
  *out = std::move(writer).Finish();
  return true;
}

void MainWindow::Impl::ReleaseRegFileRoots(TabEntry* entry) {
  if (!entry) {
    return;
  }
  for (auto& root : entry->reg_file_roots) {
    if (root.root) {
      RegistryStore::UnregisterVirtualRoot(root.root);
      root.root = nullptr;
    }
    root.data.reset();
  }
  entry->reg_file_roots.clear();
}

bool MainWindow::Impl::OpenRegFileTab(const std::wstring& path) {
  if (!tab_ || path.empty()) {
    return false;
  }
  if (!FileExists(path)) {
    ui::ShowError(hwnd_, L"Registry file not found.");
    return false;
  }
  std::wstring label = FileNameOnly(path);
  if (label.empty()) {
    label = L"Registry File";
  }
  std::wstring path_lower = ToLower(path);
  auto start_parse = [&]() {
    if (reg_file_parse_sessions_.find(path_lower) != reg_file_parse_sessions_.end()) {
      return;
    }
    auto session = std::make_unique<RegFileParseSession>();
    session->source_path = path;
    session->source_lower = path_lower;
    HWND hwnd = hwnd_;
    RegFileParseSession* session_ptr = session.get();
    session->work.Start([this, session_ptr, hwnd](
                            uint64_t generation,
                            std::atomic_bool& cancel) {
      auto payload = std::make_unique<RegFileParsePayload>();
      payload->generation = generation;
      payload->source_path = session_ptr->source_path;
      payload->source_lower = session_ptr->source_lower;
      std::wstring parse_error;
      std::vector<ParsedRegFileRoot> parsed_roots;
      bool cancelled = false;
      if (!ParseRegFileToVirtualRoots(payload->source_path, &parsed_roots,
                                      &parse_error, &cancel,
                                      &cancelled)) {
        if (!cancelled && parse_error.empty()) {
          parse_error = L"Failed to read registry file.";
        }
      }
      payload->roots = std::move(parsed_roots);
      payload->error = std::move(parse_error);
      payload->cancelled = cancelled;
      if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, frame::message_id::kRegFileLoadReady, 0, reinterpret_cast<LPARAM>(payload.get()))) {
        return;
      }
      ReleasePostedPayload(payload);
    });
    reg_file_parse_sessions_.emplace(path_lower, std::move(session));
  };

  for (size_t i = 0; i < tabs_.size(); ++i) {
    TabEntry& entry = tabs_[i];
    if (entry.kind != TabEntry::Kind::kRegFile) {
      continue;
    }
    if (EqualsInsensitive(entry.reg_file_path, path)) {
      entry.reg_file_path = path;
      entry.reg_file_label = label;
      entry.reg_file_loading = true;
      TCITEMW item = {};
      item.mask = TCIF_TEXT;
      item.pszText = const_cast<wchar_t*>(label.c_str());
      TabCtrl_SetItem(tab_, static_cast<int>(i), &item);
      TabCtrl_SetCurSel(tab_, static_cast<int>(i));
      SyncRegFileTabSelection();
      ApplyViewVisibility();
      UpdateStatus();
      start_parse();
      AppendHistoryEntry(L"Open .reg file " + FileNameOnly(path), L"", path);
      return true;
    }
  }

  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(label.c_str());
  int index = TabCtrl_GetItemCount(tab_);
  TabCtrl_InsertItem(tab_, index, &item);
  TabEntry entry;
  entry.kind = TabEntry::Kind::kRegFile;
  entry.reg_file_path = path;
  entry.reg_file_label = label;
  entry.reg_file_dirty = false;
  entry.reg_file_loading = true;
  tabs_.push_back(std::move(entry));
  UpdateTabWidth();
  TabCtrl_SetCurSel(tab_, index);
  SyncRegFileTabSelection();
  ApplyViewVisibility();
  UpdateStatus();
  start_parse();
  AppendHistoryEntry(L"Open .reg file " + FileNameOnly(path), L"", path);
  return true;
}

} // namespace regkit
