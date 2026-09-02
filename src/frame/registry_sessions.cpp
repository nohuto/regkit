// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::ReleaseRemoteRegistry() {
  if (remote_hklm_) {
    RegCloseKey(remote_hklm_);
    remote_hklm_ = nullptr;
  }
  if (remote_hku_) {
    RegCloseKey(remote_hku_);
    remote_hku_ = nullptr;
  }
  remote_machine_.clear();
}

bool MainWindow::Impl::UnloadOfflineRegistry(std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (offline_roots_.empty()) {
    return true;
  }
  ClearOfflineDirty();
  for (HKEY root : offline_roots_) {
    if (!RegistryStore::CloseOfflineHive(root, error)) {
      return false;
    }
  }
  RegistryStore::SetOfflineRoots({});
  offline_roots_.clear();
  offline_root_labels_.clear();
  offline_root_paths_.clear();
  offline_root_ = nullptr;
  offline_mount_.clear();
  offline_root_name_.clear();
  return true;
}

void MainWindow::Impl::ApplyRegistryRoots(const std::vector<RegistryRootEntry>& roots) {
  browse_.roots() = roots;
  ResetHiveListCache();
  browse_.set_current_node(nullptr);
  browse_.values().Clear();
  current_key_count_ = 0;
  current_value_count_ = 0;
  browse_.tree().SetRegeditLayout(false);
  browse_.tree().SetRootLabel(TreeRootLabel());
  browse_.tree().PopulateRoots(browse_.roots());
  ResetNavigationState();
  UpdateStatus();

  SelectDefaultTreeItem();
}

std::vector<std::wstring> MainWindow::Impl::BuildVisibleTreePathParts(const std::wstring& path) const {
  std::vector<std::wstring> parts = registry_path::Split(path);
  if (parts.empty()) {
    return parts;
  }

  std::wstring root_label = TreeRootLabel();
  if (!root_label.empty() && !parts.empty() && EqualsInsensitive(parts.front(), root_label)) {
    parts.erase(parts.begin());
  }
  if (!parts.empty() && EqualsInsensitive(parts.front(), L"Computer")) {
    parts.erase(parts.begin());
  }

  auto is_standard_root = [](const std::wstring& name) -> bool {
    if (StartsWithInsensitive(name, L"HKEY_")) {
      return true;
    }
    return EqualsInsensitive(name, L"HKLM") || EqualsInsensitive(name, L"HKCU") || EqualsInsensitive(name, L"HKCR") || EqualsInsensitive(name, L"HKU") || EqualsInsensitive(name, L"HKCC");
  };
  if (!parts.empty() && EqualsInsensitive(parts.front(), L"Registry")) {
    if (parts.size() > 1 && is_standard_root(parts[1])) {
      if (false) {
        parts.erase(parts.begin());
      } else {
        parts.front() = kStandardGroupLabel;
      }
    } else if (!false) {
      parts.front() = kRealGroupLabel;
    }
  } else if (!parts.empty() && EqualsInsensitive(parts.front(), L"Real Registry")) {
    if (false) {
      parts.clear();
      return parts;
    }
    parts.front() = kRealGroupLabel;
    if (parts.size() > 1 && EqualsInsensitive(parts[1], kRealGroupLabel)) {
      parts.erase(parts.begin() + 1);
    }
  }

  if (registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty()) {
    std::wstring machine = StripMachinePrefix(remote_machine_);
    if (!machine.empty() && !parts.empty() && EqualsInsensitive(parts.front(), machine)) {
      parts.erase(parts.begin());
    }
  }
  if (registry_mode_ == RegistryMode::kOffline && !offline_root_labels_.empty() && parts.size() >= 2) {
    std::wstring root_name = offline_root_name_;
    auto is_offline_label = [&](const std::wstring& name) {
      for (const auto& label : offline_root_labels_) {
        if (EqualsInsensitive(label, name)) {
          return true;
        }
      }
      return false;
    };
    if (!root_name.empty() && EqualsInsensitive(parts[0], root_name) && is_offline_label(parts[1])) {
      parts.erase(parts.begin());
    }
  }

  if (false) {
    if (!parts.empty() && EqualsInsensitive(parts.front(), kStandardGroupLabel)) {
      parts.erase(parts.begin());
    }
    if (!parts.empty() && EqualsInsensitive(parts.front(), kRealGroupLabel)) {
      parts.clear();
      return parts;
    }
    return parts;
  }

  if (!parts.empty()) {
    if (!EqualsInsensitive(parts.front(), kStandardGroupLabel) && !EqualsInsensitive(parts.front(), kRealGroupLabel)) {
      if (EqualsInsensitive(parts.front(), L"REGISTRY")) {
        parts.insert(parts.begin(), kRealGroupLabel);
      } else {
        parts.insert(parts.begin(), kStandardGroupLabel);
      }
    }
  }
  return parts;
}

void MainWindow::Impl::RefreshVisibleRegistryTreeLayout(bool preserve_selection) {
  if (!browse_.tree().hwnd() || browse_.roots().empty()) {
    return;
  }

  std::wstring selected_path;
  std::vector<std::wstring> expanded_paths;
  if (preserve_selection) {
    CaptureTreeState(&selected_path, &expanded_paths);
  }

  browse_.set_current_node(nullptr);
  browse_.values().Clear();
  browse_.tree().SetRegeditLayout(false);
  browse_.tree().SetRootLabel(TreeRootLabel());
  browse_.tree().PopulateRoots(browse_.roots());

  if (!preserve_selection) {
    SelectDefaultTreeItem();
    return;
  }

  std::sort(expanded_paths.begin(), expanded_paths.end(), [](const std::wstring& left, const std::wstring& right) {
    if (left.size() != right.size()) {
      return left.size() < right.size();
    }
    return _wcsicmp(left.c_str(), right.c_str()) < 0;
  });
  for (const auto& path : expanded_paths) {
    ExpandTreePath(path);
  }
  if (!selected_path.empty() && SelectTreePath(selected_path)) {
    return;
  }
  SelectDefaultTreeItem();
}

std::wstring MainWindow::Impl::TreeRootLabel() const {
  if (false) {
    return L"Computer";
  }
  if (registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty()) {
    return StripMachinePrefix(remote_machine_);
  }
  wchar_t buffer[MAX_COMPUTERNAME_LENGTH + 1] = {};
  DWORD size = static_cast<DWORD>(_countof(buffer));
  if (GetComputerNameW(buffer, &size) && size > 0) {
    return std::wstring(buffer, size);
  }
  return L"Computer";
}

void MainWindow::Impl::SelectDefaultTreeItem() {
  if (!browse_.tree().hwnd()) {
    return;
  }
  HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
  if (!root) {
    return;
  }
  HTREEITEM group = TreeView_GetChild(browse_.tree().hwnd(), root);
  HTREEITEM standard_group = nullptr;
  while (group) {
    wchar_t text[128] = {};
    TVITEMW tvi = {};
    tvi.mask = TVIF_TEXT;
    tvi.hItem = group;
    tvi.pszText = text;
    tvi.cchTextMax = static_cast<int>(_countof(text));
    if (TreeView_GetItem(browse_.tree().hwnd(), &tvi)) {
      if (_wcsicmp(text, kStandardGroupLabel) == 0) {
        standard_group = group;
        break;
      }
    }
    group = TreeView_GetNextSibling(browse_.tree().hwnd(), group);
  }
  if (standard_group) {
    TreeView_SelectItem(browse_.tree().hwnd(), standard_group);
    return;
  }
  group = TreeView_GetChild(browse_.tree().hwnd(), root);
  while (group) {
    RegistryNode* node = browse_.tree().NodeFromItem(group);
    if (node) {
      TreeView_SelectItem(browse_.tree().hwnd(), group);
      return;
    }
    HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), group);
    if (child) {
      TreeView_SelectItem(browse_.tree().hwnd(), child);
      return;
    }
    group = TreeView_GetNextSibling(browse_.tree().hwnd(), group);
  }
}

void MainWindow::Impl::CaptureRegistryTabState(int index) {
  if (!browse_.tree().hwnd() || index < 0 ||
      static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind == TabEntry::Kind::kSearch) {
    return;
  }
  CaptureTreeState(&entry.selected_path, &entry.expanded_paths);
}

void MainWindow::Impl::ResetRegistryTreeState() {
  if (!browse_.tree().hwnd()) {
    return;
  }
  HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
  if (!root) {
    return;
  }

  SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, FALSE, 0);
  std::function<void(HTREEITEM)> collapse = [&](HTREEITEM item) {
    while (item) {
      HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), item);
      if (child) {
        collapse(child);
      }
      TreeView_Expand(browse_.tree().hwnd(), item, TVE_COLLAPSE);
      item = TreeView_GetNextSibling(browse_.tree().hwnd(), item);
    }
  };
  HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), root);
  if (child) {
    collapse(child);
  }
  TreeView_SelectItem(browse_.tree().hwnd(), root);
  SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, TRUE, 0);
  RedrawWindow(browse_.tree().hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}

void MainWindow::Impl::RestoreRegistryTabState(int index) {
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size() || !browse_.tree().hwnd()) {
    return;
  }
  const TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind == TabEntry::Kind::kSearch) {
    return;
  }
  ResetRegistryTreeState();
  if (entry.expanded_paths.empty() && entry.selected_path.empty()) {
    SelectDefaultTreeItem();
    return;
  }
  std::vector<std::wstring> expanded;
  expanded.reserve(entry.expanded_paths.size());
  std::unordered_set<std::wstring> seen;
  seen.reserve(entry.expanded_paths.size());
  for (const auto& path : entry.expanded_paths) {
    if (path.empty()) {
      continue;
    }
    std::wstring key = ToLower(path);
    if (seen.insert(key).second) {
      expanded.push_back(path);
    }
  }
  std::sort(expanded.begin(), expanded.end(), [](const std::wstring& left, const std::wstring& right) {
    if (left.size() != right.size()) {
      return left.size() < right.size();
    }
    return _wcsicmp(left.c_str(), right.c_str()) < 0;
  });
  for (const auto& path : expanded) {
    ExpandTreePath(path);
  }
  if (!entry.selected_path.empty() && SelectTreePath(entry.selected_path)) {
    return;
  }
  SelectDefaultTreeItem();
}

std::wstring MainWindow::Impl::LocalRegistryTabLabel(int index) const {
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return L"Local Registry";
  }
  int local_count = 0;
  int local_index = 0;
  for (size_t i = 0; i < tabs_.size(); ++i) {
    const TabEntry& entry = tabs_[i];
    if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kLocal) {
      continue;
    }
    ++local_count;
    if (static_cast<int>(i) == index) {
      local_index = local_count;
    }
  }
  if (local_count <= 1 || local_index <= 1) {
    return L"Local Registry";
  }
  return L"Local Registry (" + std::to_wstring(local_index) + L")";
}

void MainWindow::Impl::RefreshRegistryTabLabels() {
  if (!tab_) {
    return;
  }
  for (size_t i = 0; i < tabs_.size(); ++i) {
    const TabEntry& entry = tabs_[i];
    if (entry.kind != TabEntry::Kind::kRegistry) {
      continue;
    }
    std::wstring label;
    if (entry.registry_mode == RegistryMode::kLocal) {
      label = LocalRegistryTabLabel(static_cast<int>(i));
    } else {
      continue;
    }
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(label.c_str());
    TabCtrl_SetItem(tab_, static_cast<int>(i), &item);
  }
  UpdateTabWidth();
  InvalidateRect(tab_, nullptr, FALSE);
}

void MainWindow::Impl::AppendRealRegistryRoot(std::vector<RegistryRootEntry>* roots) {
  if (!roots || registry_mode_ != RegistryMode::kLocal) {
    return;
  }
  if (!registry_root_.get()) {
    registry_root_ = util::OpenNativeRegistryRoot();
  }
  if (!registry_root_.get()) {
    return;
  }
  RegistryRootEntry entry;
  entry.root = registry_root_.get();
  entry.display_name = L"REGISTRY";
  entry.path_name = L"REGISTRY";
  entry.subkey_prefix = L"";
  entry.group = RegistryRootGroup::kReal;
  roots->push_back(std::move(entry));
}

bool MainWindow::Impl::SwitchToLocalRegistry() {
  bool needs_reload = registry_mode_ != RegistryMode::kLocal;
  if (!needs_reload) {
    if (browse_.roots().empty()) {
      needs_reload = true;
    } else if (RegistryStore::IsVirtualRoot(browse_.roots().front().root)) {
      needs_reload = true;
    } else {
      auto has_root = [&](HKEY root) -> bool {
        for (const auto& entry_root : browse_.roots()) {
          if (entry_root.root == root) {
            return true;
          }
        }
        return false;
      };
      if (!has_root(HKEY_CLASSES_ROOT) || !has_root(HKEY_CURRENT_USER) || !has_root(HKEY_LOCAL_MACHINE) || !has_root(HKEY_USERS) || !has_root(HKEY_CURRENT_CONFIG)) {
        needs_reload = true;
      }
    }
  }
  if (!needs_reload) {
    return true;
  }
  if (registry_mode_ == RegistryMode::kOffline) {
    std::wstring error;
    if (!UnloadOfflineRegistry(&error)) {
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      return false;
    }
  }
  ReleaseRemoteRegistry();
  registry_mode_ = RegistryMode::kLocal;
  browse_.set_source_kind(browse::SourceKind::kLocal);
  UpdateRegistryTabEntry(RegistryMode::kLocal, L"", L"");
  std::vector<RegistryRootEntry> roots = RegistryStore::DefaultRoots(show_extra_hives_);
  AppendRealRegistryRoot(&roots);
  ApplyRegistryRoots(roots);
  RefreshRegistryTabLabels();
  return true;
}

bool MainWindow::Impl::SwitchToRemoteRegistry() {
  std::wstring machine = remote_machine_;
  editors::TextRequest request;
  request.title = L"Connect to Remote Registry";
  request.label = L"Computer name (e.g. \\\\MACHINE):";
  request.text = machine;
  editors::TextResult text_result;
  if (!editors::EditText(hwnd_, request, &text_result)) {
    return false;
  }
  machine = std::move(text_result.text);
  machine = NormalizeMachineName(machine);
  if (machine.empty()) {
    ui::ShowError(hwnd_, L"Computer name is required.");
    return false;
  }

  HKEY hklm = nullptr;
  LONG result = RegConnectRegistryW(machine.c_str(), HKEY_LOCAL_MACHINE, &hklm);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }

  HKEY hku = nullptr;
  LONG hku_result = RegConnectRegistryW(machine.c_str(), HKEY_USERS, &hku);

  if (registry_mode_ == RegistryMode::kOffline) {
    std::wstring error;
    if (!UnloadOfflineRegistry(&error)) {
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      if (hku) {
        RegCloseKey(hku);
      }
      RegCloseKey(hklm);
      return false;
    }
  }

  ReleaseRemoteRegistry();
  registry_mode_ = RegistryMode::kRemote;
  browse_.set_source_kind(browse::SourceKind::kRemote);
  remote_machine_ = machine;
  remote_hklm_ = hklm;
  remote_hku_ = hku;
  UpdateRegistryTabEntry(RegistryMode::kRemote, L"", remote_machine_);

  std::wstring prefix = machine + L"\\";
  std::vector<RegistryRootEntry> roots;
  roots.push_back({remote_hklm_, L"HKEY_LOCAL_MACHINE", prefix + L"HKEY_LOCAL_MACHINE", L""});
  if (remote_hku_) {
    roots.push_back({remote_hku_, L"HKEY_USERS", prefix + L"HKEY_USERS", L""});
  }

  UpdateTabText(L"Remote Registry (" + StripMachinePrefix(machine) + L")");
  ApplyRegistryRoots(roots);
  RefreshRegistryTabLabels();

  if (hku_result != ERROR_SUCCESS) {
    std::wstring message = L"Connected to HKEY_LOCAL_MACHINE, but HKEY_USERS was unavailable.\n";
    message += FormatWin32Error(hku_result);
    ui::ShowError(hwnd_, message);
  }
  return true;
}

bool MainWindow::Impl::SwitchToOfflineRegistry() {
  const int choice = ui::PromptChoice(
      hwnd_, L"Load the offline registry from a single hive file, or from a folder of hives?",
      L"Offline Registry", L"Hive File", L"Folder", L"Cancel", 60, 470);
  std::wstring hive_path;
  HRESULT hr = S_OK;
  if (choice == IDYES) {
    hr = win32::ChooseFileToOpen(hwnd_, kOfflineHiveFilter, &hive_path);
  } else if (choice == IDNO) {
    hr = win32::ChooseFolder(hwnd_, &hive_path);
  } else {
    return false;
  }
  if (!ui::ReportFileDialogResult(hwnd_, hr)) {
    return false;
  }
  return LoadOfflineRegistryFromPath(hive_path, true);
}

bool MainWindow::Impl::LoadOfflineRegistryFromPath(const std::wstring& path, bool open_new_tab) {
  if (registry_mode_ == RegistryMode::kOffline && !offline_roots_.empty()) {
    std::wstring error;
    if (!UnloadOfflineRegistry(&error)) {
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      return false;
    }
  }

  std::wstring selection_path = TrimTrailingSeparators(path);
  if (selection_path.empty()) {
    return false;
  }

  bool is_dir = IsDirectoryPath(selection_path);
  std::vector<OfflineHiveCandidate> candidates;
  if (is_dir) {
    CollectOfflineHivesInFolder(selection_path, &candidates);
    if (candidates.empty()) {
      ui::ShowError(hwnd_, L"The selected folder does not contain a registry hive file.");
      return false;
    }
  } else {
    std::wstring mount_name = TrimWhitespace(FileBaseName(selection_path));
    if (mount_name.empty()) {
      mount_name = L"OfflineHive";
    }
    candidates.push_back({selection_path, mount_name});
  }

  offline_root_name_ = ResolveOfflineRootName(selection_path, is_dir, browse_.current_node());
  if (offline_root_name_.empty()) {
    offline_root_name_ = L"HKEY_LOCAL_MACHINE";
  }

  std::wstring error;
  std::vector<HKEY> handles;
  std::vector<std::wstring> labels;
  std::vector<std::wstring> paths;
  std::vector<RegistryRootEntry> roots;
  handles.reserve(candidates.size());
  labels.reserve(candidates.size());
  paths.reserve(candidates.size());
  roots.reserve(candidates.size());
  auto close_handles = [&](std::vector<HKEY>* to_close) {
    if (!to_close) {
      return;
    }
    for (HKEY root : *to_close) {
      RegistryStore::CloseOfflineHive(root, nullptr);
    }
  };
  for (const auto& candidate : candidates) {
    HKEY hive_handle = nullptr;
    if (!RegistryStore::OpenOfflineHive(candidate.path, &hive_handle, &error)) {
      close_handles(&handles);
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      return false;
    }
    std::wstring label = TrimWhitespace(candidate.label);
    if (label.empty()) {
      label = TrimWhitespace(FileBaseName(candidate.path));
      if (label.empty()) {
        label = L"OfflineHive";
      }
    }
    std::wstring path_name = offline_root_name_ + L"\\" + label;
    roots.push_back({hive_handle, label, path_name, L""});
    handles.push_back(hive_handle);
    labels.push_back(label);
    paths.push_back(candidate.path);
  }

  if (tab_ && open_new_tab) {
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(L"Offline Registry");
    int index = TabCtrl_GetItemCount(tab_);
    TabCtrl_InsertItem(tab_, index, &item);
    TabEntry entry;
    entry.kind = TabEntry::Kind::kRegistry;
    entry.registry_mode = RegistryMode::kOffline;
    entry.offline_path = selection_path;
    tabs_.push_back(std::move(entry));
    UpdateTabWidth();
    suppress_tab_change_ = true;
    SelectTabIndex(index);
    suppress_tab_change_ = false;
  }

  ReleaseRemoteRegistry();
  registry_mode_ = RegistryMode::kOffline;
  browse_.set_source_kind(browse::SourceKind::kOffline);
  offline_roots_ = std::move(handles);
  offline_root_labels_ = std::move(labels);
  offline_root_paths_ = std::move(paths);
  if (offline_roots_.size() == 1) {
    offline_root_ = offline_roots_.front();
    offline_mount_ = offline_root_labels_.front();
  } else {
    offline_root_ = nullptr;
    offline_mount_.clear();
  }
  RegistryStore::SetOfflineRoots(offline_roots_);

  std::wstring tab_text = L"Offline Registry";
  if (offline_roots_.size() == 1 && !offline_root_name_.empty() && !offline_mount_.empty()) {
    tab_text = L"Offline Registry (" + offline_root_name_ + L"\\" + offline_mount_ + L")";
  } else if (!offline_root_name_.empty()) {
    tab_text = L"Offline Registry (" + offline_root_name_ + L")";
  }
  UpdateTabText(tab_text);
  UpdateRegistryTabEntry(RegistryMode::kOffline, selection_path, L"");
  ApplyRegistryRoots(roots);
  RefreshRegistryTabLabels();
  HistoryEntry history;
  history.action = L"Load offline registry";
  history.new_data = selection_path;
  AppendHistoryEntry(std::move(history));
  return true;
}

bool MainWindow::Impl::SaveOfflineRegistry() {
  if (registry_mode_ != RegistryMode::kOffline || offline_roots_.empty()) {
    ui::ShowError(hwnd_, L"No offline registry is loaded.");
    return false;
  }
  if (offline_roots_.size() > 1) {
    if (offline_root_paths_.size() != offline_roots_.size()) {
      ui::ShowError(hwnd_, L"Failed to resolve offline hive paths for saving.");
      return false;
    }
    for (size_t i = 0; i < offline_roots_.size(); ++i) {
      const std::wstring& path = offline_root_paths_[i];
      if (path.empty()) {
        ui::ShowError(hwnd_, L"Failed to resolve offline hive path for saving.");
        return false;
      }
      DWORD attrs = GetFileAttributesW(path.c_str());
      if (attrs != INVALID_FILE_ATTRIBUTES) {
        if (!DeleteFileW(path.c_str())) {
          ui::ShowError(hwnd_, FormatWin32Error(GetLastError()));
          return false;
        }
      }
      std::wstring error;
      if (!RegistryStore::SaveOfflineHive(offline_roots_[i], path, &error)) {
        ui::ShowError(hwnd_, error.empty() ? L"Failed to save offline hive." : error);
        return false;
      }
    }
    ClearOfflineDirty();
    HistoryEntry history;
    history.action = L"Save offline registry";
    history.new_data = std::to_wstring(offline_roots_.size()) + L" hives";
    AppendHistoryEntry(std::move(history));
    return true;
  }
  if (!offline_root_) {
    ui::ShowError(hwnd_, L"No offline registry is loaded.");
    return false;
  }

  std::wstring path;
  if (!PromptSaveFile(hwnd_, L"Hive Files (*.*)\0*.*\0", &path)) {
    return false;
  }

  DWORD attrs = GetFileAttributesW(path.c_str());
  if (attrs != INVALID_FILE_ATTRIBUTES) {
    if (!DeleteFileW(path.c_str())) {
      ui::ShowError(hwnd_, FormatWin32Error(GetLastError()));
      return false;
    }
  }

  std::wstring error;
  if (!RegistryStore::SaveOfflineHive(offline_root_, path, &error)) {
    ui::ShowError(hwnd_, error.empty() ? L"Failed to save offline hive." : error);
    return false;
  }
  ClearOfflineDirty();
  HistoryEntry history;
  history.action = L"Save offline registry";
  history.new_data = path;
  AppendHistoryEntry(std::move(history));
  return true;
}

void MainWindow::Impl::NavigateToAddress() {
  wchar_t buffer[512] = {};
  GetWindowTextW(browse_.address(), buffer, static_cast<int>(_countof(buffer)));
  std::wstring path = NormalizeRegistryPath(buffer);
  if (path.empty()) {
    return;
  }
  std::wstring jump_key_path;
  std::wstring jump_value_name;
  if (ResolveExternalJumpTarget(path, &jump_key_path, &jump_value_name) && !jump_value_name.empty()) {
    if (NavigateToResolvedExternalJump(jump_key_path, jump_value_name)) {
      AddAddressHistory(jump_key_path);
    }
    return;
  }
  if (SelectTreePath(path)) {
    AddAddressHistory(path);
  } else {
    std::wstring nearest;
    if (!FindNearestExistingPath(path, &nearest) || nearest.empty()) {
      ui::ShowWarning(hwnd_, L"Registry path not found.");
      return;
    }
    std::wstring message = L"The registry key doesn't exist:";
    if (read_only_) {
      message += L"\nRead only mode is enabled.";
      int result = ui::PromptKeyChoice(hwnd_, message, path, L"Registry path not found", L"Go to nearest key", L"", L"Cancel", 70);
      if (result == IDYES) {
        if (SelectTreePath(nearest)) {
          AddAddressHistory(nearest);
        }
      }
      return;
    }
    int result = ui::PromptKeyChoice(hwnd_, message, path, L"Registry path not found", L"Go to nearest key", L"Create key", L"Cancel", 70);
    if (result == IDYES) {
      if (SelectTreePath(nearest)) {
        AddAddressHistory(nearest);
      }
      return;
    }
    if (result == IDNO) {
      if (!CreateRegistryPath(path)) {
        ui::ShowError(hwnd_, L"Failed to create registry key.");
        return;
      }
      RefreshTreePath(nearest);
      if (SelectTreePath(path)) {
        AddAddressHistory(path);
      }
    }
  }
}

void MainWindow::Impl::ApplyQueuedExternalJump() {
  if (queued_external_jump_target_.empty()) {
    return;
  }
  std::wstring target = std::move(queued_external_jump_target_);
  queued_external_jump_target_.clear();
  NavigateToExternalJump(target);
}

bool MainWindow::Impl::ResolveExternalJumpTarget(const std::wstring& target, std::wstring* key_path, std::wstring* value_name) const {
  if (!key_path || !value_name) {
    return false;
  }
  key_path->clear();
  value_name->clear();

  std::wstring normalized = NormalizeRegistryPath(target);
  if (normalized.empty()) {
    return false;
  }

  RegistryNode node;
  auto has_value = [&](const RegistryNode& candidate, const std::wstring& candidate_value) -> bool {
    if (candidate_value.empty()) {
      return false;
    }
    bool found = false;
    RegistryStore::EnumKeyStreaming(
        candidate, true, false, false, nullptr,
        [&](const ValueInfo& value, const BYTE*, DWORD) {
          found = EqualsInsensitive(value.name, candidate_value);
          return !found;
        },
        {});
    return found;
  };
  if (ResolvePathToNode(normalized, &node)) {
    KeyInfo info = {};
    if (RegistryStore::QueryKeyInfo(node, &info)) {
      *key_path = std::move(normalized);
      return true;
    }
  }

  size_t search_pos = normalized.size();
  while (search_pos > 0) {
    size_t slash = normalized.rfind(L'\\', search_pos - 1);
    if (slash == std::wstring::npos || slash + 1 >= normalized.size()) {
      break;
    }
    std::wstring candidate_key = normalized.substr(0, slash);
    std::wstring candidate_value = normalized.substr(slash + 1);
    if (!candidate_value.empty() && ResolvePathToNode(candidate_key, &node)) {
      KeyInfo info = {};
      if (RegistryStore::QueryKeyInfo(node, &info)) {
        if (!has_value(node, candidate_value)) {
          if (slash == 0) {
            break;
          }
          search_pos = slash;
          continue;
        }
        *key_path = std::move(candidate_key);
        *value_name = std::move(candidate_value);
        return true;
      }
    }
    if (slash == 0) {
      break;
    }
    search_pos = slash;
  }

  std::wstring nearest;
  if (FindNearestExistingPath(normalized, &nearest) && !nearest.empty()) {
    *key_path = std::move(nearest);
    return true;
  }
  return false;
}

bool MainWindow::Impl::NavigateToExternalJump(const std::wstring& target) {
  if (registry_mode_ != RegistryMode::kLocal && !SwitchToLocalRegistry()) {
    return false;
  }
  std::wstring key_path;
  std::wstring value_name;
  if (!ResolveExternalJumpTarget(target, &key_path, &value_name)) {
    return false;
  }
  return NavigateToResolvedExternalJump(key_path, value_name);
}

bool MainWindow::Impl::SearchResultOpensInNewTab() const {
  if (!tab_) {
    return false;
  }
  const int index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
  return index >= 0 && static_cast<size_t>(index) < search_tabs_.size() &&
         search_tabs_[static_cast<size_t>(index)].open_in_new_tab;
}

void MainWindow::Impl::ActivateRegistryTab() {
  const int registry_tab = FindFirstRegistryTabIndex();
  if (registry_tab < 0) {
    OpenLocalRegistryTab();
    return;
  }
  suppress_tab_change_ = true;
  SelectTabIndex(registry_tab);
  suppress_tab_change_ = false;
  ApplyTabSelection(registry_tab);
}

bool MainWindow::Impl::NavigateToResolvedExternalJump(const std::wstring& key_path, const std::wstring& value_name) {
  if (key_path.empty()) {
    return false;
  }

  int sel = TabCtrl_GetCurSel(tab_);
  if (sel < 0 || static_cast<size_t>(sel) >= tabs_.size() || tabs_[static_cast<size_t>(sel)].kind != TabEntry::Kind::kRegistry) {
    ActivateRegistryTab();
  }

  if (registry_mode_ != RegistryMode::kLocal) {
    if (!SwitchToLocalRegistry()) {
      return false;
    }
  }

  ApplyViewVisibility();
  UpdateStatus();

  BeginJumpUiBatch();
  if (!SelectTreePath(key_path)) {
    EndJumpUiBatch();
    return false;
  }
  ApplyTreeSelectionEffects(browse_.current_node());
  EndJumpUiBatch();

  AddAddressHistory(key_path);

  pending_external_value_key_path_.clear();
  pending_external_value_name_.clear();
  if (!value_name.empty()) {
    pending_external_value_key_path_ = key_path;
    pending_external_value_name_ = value_name;
    if (!SelectValueByName(value_name) && browse_.current_node() && !value_list_loading_) {
      UpdateValueListForNode(browse_.current_node());
    }
  }
  return true;
}

} // namespace regkit
