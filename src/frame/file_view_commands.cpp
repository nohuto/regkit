// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"

namespace regkit {
using namespace command_detail;

bool MainWindow::Impl::HandleDynamicCommand(int command_id) {
  if (command_id >= cmd::kFavoritesItemBase && command_id <= cmd::kFavoritesItemMax) {
    if (!favorites_loaded_) {
      RefreshFavoritesCache();
    }
    size_t index = static_cast<size_t>(command_id - cmd::kFavoritesItemBase);
    if (index < favorites_cache_.size()) {
      SelectTreePath(favorites_cache_[index]);
      return true;
    }
  }
  if (command_id >= cmd::kDefaultBundledBase && command_id <= cmd::kDefaultBundledMax) {
    size_t index = static_cast<size_t>(command_id - cmd::kDefaultBundledBase);
    if (index < bundled_defaults_.size()) {
      const auto& entry = bundled_defaults_[index];
      if (RemoveDefaultByPath(entry.path)) {
        return true;
      }
      LoadDefaultFromFile(entry.label, entry.path);
      return true;
    }
  }
  if (command_id >= cmd::kDefaultRecentBase && command_id <= cmd::kDefaultRecentMax) {
    size_t index = static_cast<size_t>(command_id - cmd::kDefaultRecentBase);
    if (index < recent_default_paths_.items().size()) {
      std::wstring path = recent_default_paths_.items()[index];
      std::wstring label = FileBaseName(path);
      if (label.empty()) {
        label = L"Default";
      }
      if (RemoveDefaultByPath(path)) {
        return true;
      }
      if (LoadDefaultFromFile(label, path)) {
        AddRecentDefaultPath(path);
        BuildMenus();
        SaveSettings();
      }
      return true;
    }
  }
  if (command_id >= cmd::kTraceRecentBase && command_id <= cmd::kTraceRecentMax) {
    size_t index = static_cast<size_t>(command_id - cmd::kTraceRecentBase);
    if (index < recent_trace_paths_.items().size()) {
      std::wstring path = recent_trace_paths_.items()[index];
      std::wstring label = FileBaseName(path);
      if (label.empty()) {
        label = L"Trace";
      }
      if (RemoveTraceByPath(path)) {
        return true;
      }
      if (LoadTraceFromFile(label, path)) {
        AddRecentTracePath(path);
        BuildMenus();
        SaveSettings();
        SaveActiveTraces();
      }
      return true;
    }
  }

  return false;
}

bool MainWindow::Impl::HandleFileCommand(int command_id) {
  switch (command_id) {
  case cmd::kFileExit:
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  case cmd::kFileImport: {
    if (!EnsureWritable()) {
      return true;
    }
    std::wstring path;
    if (!PromptOpenFilePath(hwnd_, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    std::wstring error;
    if (ImportRegFileFromPath(path, &error)) {
      AppendHistoryEntry(L"Import .reg file " + FileNameOnly(path), L"", path);
    } else if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return true;
  }
  case cmd::kFileOpenRegFile: {
    std::wstring path;
    if (!PromptOpenFilePath(hwnd_, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    OpenRegFileTab(path);
    return true;
  }
  case cmd::kFileSave: {
    if (!EnsureWritable()) {
      return true;
    }
    if (IsRegFileTabSelected()) {
      int tab_index = TabCtrl_GetCurSel(tab_);
      if (tab_index >= 0 && static_cast<size_t>(tab_index) < tabs_.size()) {
        if (tabs_[static_cast<size_t>(tab_index)].reg_file_dirty) {
          SaveRegFileTab(tab_index);
        }
      }
      return true;
    }
    if (registry_mode_ == RegistryMode::kOffline) {
      int index = CurrentRegistryTabIndex();
      if (index >= 0 && static_cast<size_t>(index) < tabs_.size()) {
        if (tabs_[static_cast<size_t>(index)].offline_dirty) {
          SaveOfflineRegistry();
        }
      }
      return true;
    }
    return true;
  }
  case cmd::kFileExport: {
    if (IsRegFileTabSelected()) {
      int tab_index = TabCtrl_GetCurSel(tab_);
      std::wstring path;
      if (!PromptSaveFilePath(hwnd_, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
        return true;
      }
      if (ExportRegFileTab(tab_index, path)) {
        AppendHistoryEntry(L"Export .reg tab " + FileNameOnly(path), L"", path);
      }
      return true;
    }
    if (!browse_.current_node()) {
      return true;
    }
    if (browse_.values().hwnd() && GetFocus() == browse_.values().hwnd()) {
      std::vector<std::wstring> selected_values;
      std::vector<std::wstring> selected_keys;
      int index = -1;
      while ((index = ListView_GetNextItem(browse_.values().hwnd(), index, LVNI_SELECTED)) >= 0) {
        const ListRow* row = browse_.values().RowAt(index);
        if (!row) {
          continue;
        }
        if (row->kind == rowkind::kValue) {
          selected_values.push_back(row->extra);
        } else if (row->kind == rowkind::kKey) {
          selected_keys.push_back(row->extra);
        }
      }
      if (!selected_values.empty() || !selected_keys.empty()) {
        auto dedupe = [](std::vector<std::wstring>* items) {
          if (!items) {
            return;
          }
          std::unordered_set<std::wstring> seen;
          std::vector<std::wstring> unique;
          unique.reserve(items->size());
          for (const auto& item : *items) {
            std::wstring key = ToLower(item);
            if (seen.insert(key).second) {
              unique.push_back(item);
            }
          }
          *items = std::move(unique);
        };
        dedupe(&selected_values);
        dedupe(&selected_keys);
        std::wstring error;
        std::wstring path = registry_path::Build(*browse_.current_node());
        if (ExportRegFileSelection(hwnd_, path, selected_values, selected_keys, &error)) {
          HistoryEntry entry;
          entry.action = L"Export registry selection";
          entry.old_data = std::to_wstring(selected_keys.size()) + L" keys, " + std::to_wstring(selected_values.size()) + L" values";
          entry.key_path = path;
          AppendHistoryEntry(std::move(entry));
        } else if (!error.empty()) {
          ui::ShowError(hwnd_, error);
        }
        return true;
      }
    }
    std::wstring error;
    std::wstring path = registry_path::Build(*browse_.current_node());
    if (ExportRegFile(hwnd_, path, &error)) {
      HistoryEntry entry;
      entry.action = L"Export registry key";
      entry.key_path = path;
      entry.new_data = path;
      AppendHistoryEntry(std::move(entry));
    } else if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return true;
  }
  case cmd::kFileImportComments: {
    std::wstring path;
    if (!PromptOpenFilePath(hwnd_, L"RegKit Comment Files (*.rkc)\0*.rkc\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    if (ImportCommentsFromFile(path)) {
      AppendHistoryEntry(L"Import comments " + FileNameOnly(path), L"", path);
    } else {
      ui::ShowError(hwnd_, L"Failed to import comments.");
    }
    return true;
  }
  case cmd::kFileExportComments: {
    std::wstring path;
    if (!PromptSaveFilePath(hwnd_, L"RegKit Comment Files (*.rkc)\0*.rkc\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    if (ExportCommentsToFile(path)) {
      AppendHistoryEntry(L"Export comments " + FileNameOnly(path), L"", path);
    } else {
      ui::ShowError(hwnd_, L"Failed to export comments.");
    }
    return true;
  }
  case cmd::kFileLoadHive: {
    if (!EnsureWritable()) {
      return true;
    }
    if (registry_mode_ == RegistryMode::kRemote) {
      ui::ShowError(hwnd_, L"Loading hives is not supported for remote registries.");
      return true;
    }
    std::wstring error;
    HKEY root = HKEY_LOCAL_MACHINE;
    if (LoadHive(hwnd_, &root, &error)) {
      const std::wstring root_name = registry_path::RootName(root);
      AppendHistoryEntry(L"Load hive", L"", root_name);
      SelectTreePath(root_name);
      RefreshTreeSelection();
      UpdateValueListForNode(browse_.current_node());
    } else if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return true;
  }
  case cmd::kFileUnloadHive: {
    if (!EnsureWritable()) {
      return true;
    }
    if (registry_mode_ == RegistryMode::kRemote) {
      ui::ShowError(hwnd_, L"Unloading hives is not supported for remote registries.");
      return true;
    }
    HKEY root = HKEY_LOCAL_MACHINE;
    std::wstring subkey;
    if (browse_.current_node() && (browse_.current_node()->root == HKEY_LOCAL_MACHINE || browse_.current_node()->root == HKEY_USERS)) {
      root = browse_.current_node()->root;
      subkey = browse_.current_node()->subkey;
    }
    if (subkey.empty() || subkey.find(L'\\') != std::wstring::npos) {
      ui::ShowError(hwnd_, L"Select a hive you loaded under HKEY_LOCAL_MACHINE or HKEY_USERS first.");
      return true;
    }
    if (ui::PromptKeyChoice(hwnd_, L"Unload this key and all of its subkeys?", subkey,
                            L"Unload Hive", L"Unload", L"", L"Cancel") != IDYES) {
      return true;
    }
    std::wstring error;
    if (!UnloadHive(hwnd_, root, subkey, &error)) {
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      return true;
    }
    AppendHistoryEntry(L"Unload hive", subkey, L"");
    SelectTreePath(registry_path::RootName(root));
    RefreshTreeSelection();
    UpdateValueListForNode(browse_.current_node());
    return true;
  }
  case cmd::kFileSaveOfflineHive:
    SaveOfflineRegistry();
    return true;
  case cmd::kFileClearHistoryOnExit:
    clear_history_on_exit_ = !clear_history_on_exit_;
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kFileClearTabsOnExit:
    clear_tabs_on_exit_ = !clear_tabs_on_exit_;
    SaveSettings();
    BuildMenus();
    return true;
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleViewCommand(int command_id) {
  switch (command_id) {
  case cmd::kViewRefresh:
    RefreshTreeSelection();
    RefreshMatchingTreeNodes();
    UpdateValueListForNode(browse_.current_node());
    return true;
  case cmd::kViewAddressBar:
    show_address_bar_ = !show_address_bar_;
    SaveSettings();
    ApplyViewVisibility();
    BuildMenus();
    return true;
  case cmd::kViewFilterBar:
    show_filter_bar_ = !show_filter_bar_;
    SaveSettings();
    ApplyViewVisibility();
    BuildMenus();
    return true;
  case cmd::kViewTabControl:
    show_tab_control_ = !show_tab_control_;
    SaveSettings();
    ApplyViewVisibility();
    BuildMenus();
    return true;
  case cmd::kTreeToggleExpand: {
    if (!browse_.tree().hwnd()) {
      return true;
    }
    HTREEITEM item = TreeView_GetSelection(browse_.tree().hwnd());
    if (!item) {
      return true;
    }
    TVITEMW tvi = {};
    tvi.hItem = item;
    tvi.mask = TVIF_STATE | TVIF_CHILDREN;
    tvi.stateMask = TVIS_EXPANDED;
    if (!TreeView_GetItem(browse_.tree().hwnd(), &tvi)) {
      return true;
    }
    bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
    bool has_child = TreeView_GetChild(browse_.tree().hwnd(), item) != nullptr || tvi.cChildren != 0;
    if (!expanded && !has_child) {
      return true;
    }
    TreeView_Expand(browse_.tree().hwnd(), item, expanded ? TVE_COLLAPSE : TVE_EXPAND);
    return true;
  }
  case cmd::kTreeExpandAll: {
    HWND tree = browse_.tree().hwnd();
    HTREEITEM root = tree ? TreeView_GetSelection(tree) : nullptr;
    if (!root) {
      return true;
    }
    HCURSOR previous = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    SendMessageW(tree, WM_SETREDRAW, FALSE, 0);
    constexpr int kExpandAllKeyLimit = 5000;
    std::vector<HTREEITEM> pending{root};
    int expanded_keys = 0;
    while (!pending.empty() && expanded_keys < kExpandAllKeyLimit) {
      HTREEITEM item = pending.back();
      pending.pop_back();
      TreeView_Expand(tree, item, TVE_EXPAND);
      ++expanded_keys;
      for (HTREEITEM child = TreeView_GetChild(tree, item); child;
           child = TreeView_GetNextSibling(tree, child)) {
        pending.push_back(child);
      }
    }
    const bool truncated = !pending.empty();
    SendMessageW(tree, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(tree, nullptr, TRUE);
    TreeView_EnsureVisible(tree, root);
    SetCursor(previous);
    MarkTreeStateDirty();
    if (truncated) {
      ui::ShowWarning(hwnd_, L"This key has too many subkeys to expand at once. " +
                                 std::to_wstring(expanded_keys) +
                                 L" keys were expanded.");
    }
    return true;
  }
  case cmd::kViewSelectAll:
    if (!SelectAllInFocusedList()) {
      HWND focus = GetFocus();
      if (focus) {
        SendMessageW(focus, EM_SETSEL, 0, -1);
      }
    }
    return true;
  case cmd::kEditInvertSelection:
    InvertSelectionInFocusedList();
    return true;
  case cmd::kViewToolbar:
    show_toolbar_ = !show_toolbar_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewKeyTree:
    show_tree_ = !show_tree_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewKeysInList:
    show_keys_in_list_ = !show_keys_in_list_;
    BuildMenus();
    UpdateValueListForNode(browse_.current_node());
    SaveSettings();
    return true;
  case cmd::kViewSimulatedKeys:
    show_simulated_keys_ = !show_simulated_keys_;
    BuildMenus();
    RefreshTreeSelection();
    UpdateValueListForNode(browse_.current_node());
    SaveSettings();
    return true;
  case cmd::kViewHistory:
    show_history_ = !show_history_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewStatusBar:
    show_status_bar_ = !show_status_bar_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewExtraHives:
    show_extra_hives_ = !show_extra_hives_;
    SaveSettings();
    BuildMenus();
    if (registry_mode_ == RegistryMode::kLocal) {
      std::vector<RegistryRootEntry> roots = RegistryStore::DefaultRoots(show_extra_hives_);
      AppendRealRegistryRoot(&roots);
      ApplyRegistryRoots(roots);
    }
    return true;
  case cmd::kViewSaveTreeState:
    if (save_tree_state_) {
      StopTreeStateWorker();
      save_tree_state_ = false;
      saved_tree_state_.selected_path.clear();
      saved_tree_state_.expanded_paths.clear();
    } else {
      save_tree_state_ = true;
      LoadTreeState();
      tree_state_restored_ = false;
      RestoreTreeState();
      StartTreeStateWorker();
      MarkTreeStateDirty();
    }
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsSaveTabs:
    save_tabs_ = !save_tabs_;
    if (!save_tabs_) {
      ClearTabsCache();
    } else {
      SaveTabs();
    }
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsReadOnly:
    read_only_ = !read_only_;
    SaveSettings();
    BuildMenus();
    if (toolbar_.hwnd()) {
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditPaste, read_only_ ? 0 : TBSTATE_ENABLED);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditDelete, read_only_ ? 0 : TBSTATE_ENABLED);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                   !read_only_ && undo_stack_.CanUndo() ? TBSTATE_ENABLED
                                                        : 0);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo,
                   !read_only_ && undo_stack_.CanRedo() ? TBSTATE_ENABLED
                                                        : 0);
    }
    return true;
  case cmd::kOptionsCompareRegistries:
    StartCompareRegistries();
    return true;
  case cmd::kViewFont: {
    FontDialogResult result = {};
    if (ShowFontDialog(hwnd_, DefaultLogFont(), !use_custom_font_, custom_font_, &result)) {
      use_custom_font_ = !result.use_default;
      custom_font_ = result.font;
      UpdateUIFont();
      SaveSettings();
    }
    return true;
  }
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleTraceDefaultCommand(int command_id) {
  switch (command_id) {
  case cmd::kTraceLoad23H2:
    if (RemoveTraceByLabel(L"23H2")) {
      return true;
    }
    LoadBundledTrace(L"23H2");
    return true;
  case cmd::kTraceLoad24H2:
    if (RemoveTraceByLabel(L"24H2")) {
      return true;
    }
    LoadBundledTrace(L"24H2");
    return true;
  case cmd::kTraceLoad25H2:
    if (RemoveTraceByLabel(L"25H2")) {
      return true;
    }
    LoadBundledTrace(L"25H2");
    return true;
  case cmd::kTraceLoadCustom:
    LoadTraceFromPrompt();
    return true;
  case cmd::kTraceClear:
    ClearTrace();
    return true;
  case cmd::kDefaultLoadCustom:
    LoadDefaultFromPrompt();
    return true;
  case cmd::kDefaultClear:
    ClearDefaults();
    return true;
  case cmd::kDefaultEditActive: {
    std::vector<std::wstring> active;
    active.reserve(active_defaults_.size());
    for (const auto& defaults : active_defaults_) {
      active.push_back(defaults.source_path);
    }
    std::wstring content = JoinLines(active);
    editors::TextRequest request;
    request.title = L"Edit Active Defaults";
    request.label = L"One default path per line.";
    request.text = content;
    request.multiline = true;
    editors::TextResult result;
    if (editors::EditText(hwnd_, request, &result)) {
      content = std::move(result.text);
      std::vector<std::wstring> lines = SplitLines(content);
      active_defaults_.clear();
      for (const auto& line : lines) {
        AddDefaultFromFile(L"", line, false, false, false);
      }
      SaveActiveDefaults();
      BuildMenus();
      UpdateValueListForNode(browse_.current_node());
      SaveSettings();
    }
    return true;
  }
  case cmd::kTraceEditActive: {
    std::vector<std::wstring> active;
    active.reserve(active_traces_.size());
    for (const auto& trace : active_traces_) {
      active.push_back(trace.source_path);
    }
    std::wstring content = JoinLines(active);
    editors::TextRequest request;
    request.title = L"Edit Active Traces";
    request.label = L"One trace path per line.";
    request.text = content;
    request.multiline = true;
    editors::TextResult result;
    if (editors::EditText(hwnd_, request, &result)) {
      content = std::move(result.text);
      std::vector<std::wstring> lines = SplitLines(content);
      LoadTraceSettings();
      active_traces_.clear();
      for (const auto& line : lines) {
        AddTraceFromFile(L"", line, nullptr, false, false);
      }
      SaveActiveTraces();
      SaveTraceSettings();
      BuildMenus();
      RefreshTreeSelection();
      UpdateValueListForNode(browse_.current_node());
      SaveSettings();
    }
    return true;
  }
  case cmd::kTraceGuide:
    win32::ShellOpen(hwnd_, L"https://github.com/nohuto/regkit/blob/main/guides/wpr-wpa.md");
    return true;
  case cmd::kDefaultEditRecent: {
    std::wstring content = JoinLines(recent_default_paths_.items());
    editors::TextRequest request;
    request.title = L"Edit Recent Defaults";
    request.label = L"One default path per line.";
    request.text = content;
    request.multiline = true;
    editors::TextResult result;
    if (editors::EditText(hwnd_, request, &result)) {
      content = std::move(result.text);
      std::vector<std::wstring> updated = SplitLines(content);
      recent_default_paths_.Replace(std::move(updated));
      NormalizeRecentDefaultList();
      SaveSettings();
      BuildMenus();
    }
    return true;
  }
  case cmd::kTraceEditRecent: {
    std::wstring content = JoinLines(recent_trace_paths_.items());
    editors::TextRequest request;
    request.title = L"Edit Recent Traces";
    request.label = L"One trace path per line.";
    request.text = content;
    request.multiline = true;
    editors::TextResult result;
    if (editors::EditText(hwnd_, request, &result)) {
      content = std::move(result.text);
      std::vector<std::wstring> updated = SplitLines(content);
      recent_trace_paths_.Replace(std::move(updated));
      NormalizeRecentTraceList();
      SaveSettings();
      BuildMenus();
    }
    return true;
  }
  default:
    return false;
  }
}

} // namespace regkit
