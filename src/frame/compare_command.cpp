// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"

namespace regkit {
using namespace command_detail;

void MainWindow::Impl::StartCompareRegistries() {
  CompareDialogDefaults defaults;
  defaults.registry_roots.reserve(browse_.roots().size());
  std::unordered_set<std::wstring> seen_roots;
  for (const auto& root : browse_.roots()) {
    if (root.path_name.empty()) {
      continue;
    }
    std::wstring key = ToLower(root.path_name);
    if (seen_roots.insert(key).second) {
      defaults.registry_roots.push_back(root.path_name);
    }
  }
  if (defaults.registry_roots.empty()) {
    defaults.registry_roots = {L"HKEY_LOCAL_MACHINE", L"HKEY_CURRENT_USER", L"HKEY_CLASSES_ROOT", L"HKEY_USERS", L"HKEY_CURRENT_CONFIG"};
  }

  CompareDialogSelection left;
  CompareDialogSelection right;
  left.type = CompareSourceType::kRegistry;
  right.type = CompareSourceType::kRegistry;
  left.recursive = true;
  right.recursive = true;
  if (browse_.current_node()) {
    std::wstring root_name = browse_.current_node()->root_name.empty() ? registry_path::RootName(browse_.current_node()->root) : browse_.current_node()->root_name;
    left.root = root_name;
    right.root = root_name;
    left.path = browse_.current_node()->subkey;
    right.path = browse_.current_node()->subkey;
  } else if (!defaults.registry_roots.empty()) {
    left.root = defaults.registry_roots.front();
    right.root = defaults.registry_roots.front();
  }
  defaults.left = left;
  defaults.right = right;

  CompareDialogResult selection;
  if (!ShowCompareDialog(hwnd_, defaults, &selection)) {
    return;
  }

  auto normalize_base = [&](const CompareDialogSelection& sel, std::wstring* out_base) -> bool {
    if (!out_base) {
      return false;
    }
    if (sel.type == CompareSourceType::kRegistry) {
      std::wstring base;
      if (!sel.path.empty()) {
        std::wstring normalized_path = NormalizeRegistryPath(sel.path);
        if (StartsWithInsensitive(normalized_path, L"HKEY_") || StartsWithInsensitive(normalized_path, L"REGISTRY")) {
          base = normalized_path;
        }
      }
      if (base.empty()) {
        base = sel.root;
        if (!sel.path.empty()) {
          base += L"\\" + sel.path;
        }
        base = NormalizeRegistryPath(base);
      }
      if (base.empty()) {
        return false;
      }
      *out_base = base;
      return true;
    }
    std::wstring base = NormalizeRegistryPath(sel.key_path);
    if (base.empty()) {
      return false;
    }
    *out_base = base;
    return true;
  };

  auto build_snapshot =
      [&](const CompareDialogSelection& source,
          search::compare::Snapshot* snapshot,
          std::wstring* error) -> bool {
    std::wstring base;
    if (!normalize_base(source, &base)) {
      if (error) {
        *error = L"Invalid registry path.";
      }
      return false;
    }
    if (source.type == CompareSourceType::kRegFile) {
      return search::compare::LoadRegFile(
          source.file_path, base, source.recursive,
          [this](const std::wstring& path) {
            return NormalizeRegistryPath(path);
          },
          snapshot, error);
    }

    RegistryNode node;
    KeyInfo info = {};
    if (!ResolvePathToNode(base, &node) ||
        !RegistryStore::QueryKeyInfo(node, &info)) {
      if (error) {
        *error = L"Registry path not found: " + base;
      }
      return false;
    }
    return search::compare::CaptureRegistry(
        base, node, source.recursive, snapshot);
  };

  search::compare::Snapshot left_snapshot;
  search::compare::Snapshot right_snapshot;
  std::wstring error;
  if (!build_snapshot(selection.left, &left_snapshot, &error)) {
    if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return;
  }
  error.clear();
  if (!build_snapshot(selection.right, &right_snapshot, &error)) {
    if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return;
  }

  std::vector<search::Result> results =
      search::compare::Diff(left_snapshot, right_snapshot);
  std::wstring tab_label = L"Registry Comparision";

  SearchTab tab;
  tab.label = std::move(tab_label);
  tab.results = std::move(results);
  tab.is_compare = true;
  search_tabs_.push_back(std::move(tab));
  int search_index = static_cast<int>(search_tabs_.size() - 1);
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(search_tabs_.back().label.c_str());
  int tab_index = TabCtrl_GetItemCount(tab_);
  TabCtrl_InsertItem(tab_, tab_index, &item);
  tabs_.push_back({TabEntry::Kind::kSearch, search_index});

  UpdateTabWidth();
  SelectTabIndex(tab_index);
  active_search_tab_index_ = tab_index;
  UpdateSearchResultsView();
  ApplyViewVisibility();
  UpdateStatus();
}

} // namespace regkit
