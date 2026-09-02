// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"
#include "frame/research_links.h"

namespace regkit {
using namespace command_detail;

std::wstring MainWindow::Impl::CommandShortcutText(int command_id) const {
  switch (command_id) {
  case cmd::kEditCopy:
    return L"Ctrl+C";
  case cmd::kEditPaste:
    return L"Ctrl+V";
  case cmd::kEditUndo:
    return L"Ctrl+Z";
  case cmd::kEditRedo:
    return L"Ctrl+Y";
  case cmd::kEditFind:
    return L"Ctrl+F";
  case cmd::kEditReplace:
    return L"Ctrl+H";
  case cmd::kEditGoTo:
    return L"Ctrl+G";
  case cmd::kEditRename:
    return L"F2";
  case cmd::kEditDelete:
    return L"Del";
  case cmd::kEditCopyKey:
    return L"Ctrl+Shift+C";
  case cmd::kViewSelectAll:
    return L"Ctrl+A";
  case cmd::kFileSave:
    return L"Ctrl+S";
  case cmd::kFileExport:
    return L"Ctrl+E";
  case cmd::kViewRefresh:
    return L"F5";
  case cmd::kNavBack:
    return L"Alt+Left";
  case cmd::kNavForward:
    return L"Alt+Right";
  case cmd::kNavUp:
    return L"Alt+Up";
  default:
    return L"";
  }
}

std::wstring MainWindow::Impl::CommandTooltipText(int command_id) const {
  switch (command_id) {
  case cmd::kRegistryLocal:
    return L"Local Registry";
  case cmd::kRegistryNetwork:
    return L"Remote Registry";
  case cmd::kRegistryOffline:
    return L"Offline Registry";
  case cmd::kEditFind:
    return L"Find";
  case cmd::kEditReplace:
    return L"Replace";
  case cmd::kFileSave:
    return L"Save";
  case cmd::kFileExport:
    return L"Export";
  case cmd::kEditUndo:
    return L"Undo";
  case cmd::kEditRedo:
    return L"Redo";
  case cmd::kEditCopy:
    return L"Copy";
  case cmd::kEditPaste:
    return L"Paste";
  case cmd::kEditDelete:
    return L"Delete";
  case cmd::kViewRefresh:
    return L"Refresh";
  case cmd::kNavBack:
    return L"Back";
  case cmd::kNavForward:
    return L"Forward";
  case cmd::kNavUp:
    return L"Up";
  default:
    return L"";
  }
}

bool MainWindow::Impl::EnsureWritable() {
  if (!read_only_) {
    return true;
  }
  ui::ShowWarning(hwnd_, L"Read only mode is enabled.");
  return false;
}

void MainWindow::Impl::BuildMenus() {
  if (deferred_startup_complete_) {
    SyncReplaceRegeditState();
  }
  if (deferred_startup_complete_ && !favorites_loaded_) {
    RefreshFavoritesCache();
  }
  if (deferred_startup_complete_ && !bundled_defaults_loaded_) {
    RefreshBundledDefaultsCache();
  }
  menu_items_.clear();
  bool can_modify = !read_only_;
  HMENU menu = CreateMenu();
  HMENU file_menu = CreatePopupMenu();
  auto append_menu = [&](HMENU target, UINT flags, int command, const wchar_t* text) {
    std::wstring shortcut = CommandShortcutText(command);
    if (!shortcut.empty()) {
      std::wstring combined = std::wstring(text) + L"\t" + shortcut;
      AppendMenuW(target, flags, command, combined.c_str());
      return;
    }
    AppendMenuW(target, flags, command, text);
  };
  bool can_save = false;
  if (can_modify && tab_) {
    int sel = TabCtrl_GetCurSel(tab_);
    if (sel >= 0 && static_cast<size_t>(sel) < tabs_.size()) {
      const auto& entry = tabs_[static_cast<size_t>(sel)];
      if (entry.kind == TabEntry::Kind::kRegFile) {
        can_save = entry.reg_file_dirty;
      } else if (entry.kind == TabEntry::Kind::kRegistry && entry.registry_mode == RegistryMode::kOffline) {
        can_save = entry.offline_dirty;
      }
    }
  }
  UINT save_flags = MF_STRING | (can_save ? 0 : MF_GRAYED);
  append_menu(file_menu, save_flags, cmd::kFileSave, L"Save");
  append_menu(file_menu, MF_STRING, cmd::kFileOpenRegFile, L"Open .reg File...");
  UINT import_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
  append_menu(file_menu, import_flags, cmd::kFileImport, L"Import...");
  append_menu(file_menu, MF_STRING, cmd::kFileExport, L"Export...");
  append_menu(file_menu, MF_STRING, cmd::kFileImportComments, L"Import Comments...");
  append_menu(file_menu, MF_STRING, cmd::kFileExportComments, L"Export Comments...");
  append_menu(file_menu, MF_STRING, cmd::kOptionsCompareRegistries, L"Compare Registries...");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  UINT hive_modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
  append_menu(file_menu, hive_modify_flags, cmd::kFileLoadHive, L"Load Hive...");
  append_menu(file_menu, hive_modify_flags, cmd::kFileUnloadHive, L"Unload Hive...");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  UINT local_flags = MF_STRING | (registry_mode_ == RegistryMode::kLocal ? MF_CHECKED : MF_UNCHECKED);
  UINT remote_flags = MF_STRING | (registry_mode_ == RegistryMode::kRemote ? MF_CHECKED : MF_UNCHECKED);
  UINT offline_flags = MF_STRING | (registry_mode_ == RegistryMode::kOffline ? MF_CHECKED : MF_UNCHECKED);
  append_menu(file_menu, local_flags, cmd::kRegistryLocal, L"Local Registry");
  append_menu(file_menu, remote_flags, cmd::kRegistryNetwork, L"Remote Registry...");
  append_menu(file_menu, offline_flags, cmd::kRegistryOffline, L"Offline Registry...");
  UINT save_offline_flags = MF_STRING | ((registry_mode_ == RegistryMode::kOffline && !offline_mount_.empty()) ? 0 : MF_GRAYED);
  append_menu(file_menu, save_offline_flags, cmd::kFileSaveOfflineHive, L"Save Offline Hive...");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  UINT clear_flags = MF_STRING | (clear_history_on_exit_ ? MF_CHECKED : MF_UNCHECKED);
  append_menu(file_menu, clear_flags, cmd::kFileClearHistoryOnExit, L"Clear History on Exit");
  UINT clear_tabs_flags = MF_STRING | (clear_tabs_on_exit_ ? MF_CHECKED : MF_UNCHECKED);
  append_menu(file_menu, clear_tabs_flags, cmd::kFileClearTabsOnExit, L"Clear Tabs on Exit");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(file_menu, MF_STRING, cmd::kFileExit, L"Exit");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(file_menu), L"File");

  HMENU edit_menu = CreatePopupMenu();
  UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
  append_menu(edit_menu, modify_flags, cmd::kEditModify, L"Modify...");
  append_menu(edit_menu, modify_flags, cmd::kEditModifyBinary, L"Modify Binary Data...");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, modify_flags, cmd::kEditUndo, L"Undo");
  append_menu(edit_menu, modify_flags, cmd::kEditRedo, L"Redo");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  HMENU edit_new = CreatePopupMenu();
  AppendMenuW(edit_new, MF_STRING, cmd::kNewKey, L"Key");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewString, L"String Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewBinary, L"Binary Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");
  AppendMenuW(edit_menu, MF_POPUP | (can_modify ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(edit_new), L"New");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, MF_STRING, cmd::kEditCopy, L"Copy");
  append_menu(edit_menu, modify_flags, cmd::kEditPaste, L"Paste");
  append_menu(edit_menu, modify_flags, cmd::kEditRename, L"Rename");
  append_menu(edit_menu, modify_flags, cmd::kEditDelete, L"Delete");
  append_menu(edit_menu, MF_STRING, cmd::kViewSelectAll, L"Select All");
  append_menu(edit_menu, MF_STRING, cmd::kEditInvertSelection, L"Invert Selection");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, MF_STRING, cmd::kEditCopyKey, L"Copy Key Name");
  append_menu(edit_menu, MF_STRING, cmd::kEditCopyKeyPath, L"Copy Key Path");
  AppendMenuW(edit_menu, MF_POPUP, reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
  UINT permissions_flags = MF_STRING | ((browse_.current_node() && can_modify) ? 0 : MF_GRAYED);
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, MF_STRING, cmd::kEditGoTo, L"Go to...");
  append_menu(edit_menu, MF_STRING, cmd::kEditFind, L"Find...");
  append_menu(edit_menu, modify_flags, cmd::kEditReplace, L"Replace...");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, permissions_flags, cmd::kEditPermissions, L"Permissions...");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(edit_menu), L"Edit");

  HMENU view_menu = CreatePopupMenu();
  append_menu(view_menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
  append_menu(view_menu, MF_STRING, cmd::kViewSelectAll, L"Select All");
  AppendMenuW(view_menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(view_menu, MF_STRING | (show_toolbar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewToolbar, L"Toolbar");
  AppendMenuW(view_menu, MF_STRING | (show_address_bar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewAddressBar, L"Address Bar");
  AppendMenuW(view_menu, MF_STRING | (show_filter_bar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewFilterBar, L"Filter Bar");
  AppendMenuW(view_menu, MF_STRING | (show_tab_control_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewTabControl, L"Tab Control");
  AppendMenuW(view_menu, MF_STRING | (show_tree_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewKeyTree, L"Key Tree");
  AppendMenuW(view_menu, MF_STRING | (show_keys_in_list_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewKeysInList, L"Keys in List");
  AppendMenuW(view_menu, MF_STRING | (show_value_grid_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewGridLines, L"Grid Lines");
  UINT simulated_flags = MF_STRING | (show_simulated_keys_ ? MF_CHECKED : MF_UNCHECKED);
  if (!HasActiveTraces()) {
    simulated_flags |= MF_GRAYED;
  }
  AppendMenuW(view_menu, simulated_flags, cmd::kViewSimulatedKeys, L"Simulated Keys");
  AppendMenuW(view_menu, MF_STRING | (show_history_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewHistory, L"History");
  AppendMenuW(view_menu, MF_STRING | (show_status_bar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewStatusBar, L"Status Bar");
  UINT extra_flags = MF_STRING | (show_extra_hives_ ? MF_CHECKED : MF_UNCHECKED);
  if (registry_mode_ != RegistryMode::kLocal) {
    extra_flags |= MF_GRAYED;
  }
  AppendMenuW(view_menu, extra_flags, cmd::kViewExtraHives, L"Show Extra Hives");
  AppendMenuW(view_menu, MF_SEPARATOR, 0, nullptr);
  UINT hive_flags = MF_STRING |
                    (ResolveSelectedHiveFilePath().empty() ? MF_GRAYED : 0);
  append_menu(view_menu, hive_flags, cmd::kOptionsHiveFileDir, L"Open Hive File");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(view_menu), L"View");

  HMENU options_menu = CreatePopupMenu();
  HMENU theme_menu = CreatePopupMenu();
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kSystem ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeSystem, L"System");
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kLight ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeLight, L"Light");
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kDark ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeDark, L"Dark");
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kCustom ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeCustom, L"Custom");
  AppendMenuW(theme_menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(theme_menu, MF_STRING, cmd::kOptionsThemePresets, L"Theme Presets...");
  AppendMenuW(options_menu, MF_POPUP, reinterpret_cast<UINT_PTR>(theme_menu), L"Theme");
  HMENU icon_menu = CreatePopupMenu();
  auto icon_flags = [&](const wchar_t* name) -> UINT { return MF_STRING | (_wcsicmp(icon_set_.c_str(), name) == 0 ? MF_CHECKED : MF_UNCHECKED); };
  AppendMenuW(icon_menu, icon_flags(kIconSetDefault), cmd::kOptionsIconSetDefault, L"Phosphor + RegEdit");
  AppendMenuW(icon_menu, icon_flags(kIconSetPhosphor), cmd::kOptionsIconSetPhosphor, L"Phosphor");
  AppendMenuW(icon_menu, icon_flags(kIconSetLucide), cmd::kOptionsIconSetLucide, L"Lucide");
  AppendMenuW(icon_menu, icon_flags(kIconSetMaterialSymbols), cmd::kOptionsIconSetMaterialSymbols, L"Material Symbols");
  AppendMenuW(icon_menu, icon_flags(kIconSetCustom), cmd::kOptionsIconSetCustom, L"Custom");
  AppendMenuW(options_menu, MF_POPUP, reinterpret_cast<UINT_PTR>(icon_menu), L"Icons");
  AppendMenuW(options_menu, MF_STRING, cmd::kViewFont, L"Font...");
  AppendMenuW(options_menu, MF_SEPARATOR, 0, nullptr);
  bool is_elevated = util::IsProcessElevated();
  bool is_system = util::IsProcessSystem();
  bool is_ti = util::IsProcessTrustedInstaller();
  UINT admin_flags = MF_STRING | (is_elevated ? MF_GRAYED : 0);
  AppendMenuW(options_menu, admin_flags, cmd::kOptionsRestartAdmin, L"Restart as Admin");
  AppendMenuW(options_menu, MF_STRING | (always_run_as_admin_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsAlwaysRunAdmin, L"Always run as Admin");
  UINT system_flags = MF_STRING | (is_system ? MF_GRAYED : 0);
  AppendMenuW(options_menu, system_flags, cmd::kOptionsRestartSystem, L"Restart as SYSTEM");
  AppendMenuW(options_menu, MF_STRING | (always_run_as_system_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsAlwaysRunSystem, L"Always run as SYSTEM");
  UINT ti_flags = MF_STRING | (is_ti ? MF_GRAYED : 0);
  AppendMenuW(options_menu, ti_flags, cmd::kOptionsRestartTrustedInstaller, L"Restart as TI");
  AppendMenuW(options_menu, MF_STRING | (always_run_as_trustedinstaller_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsAlwaysRunTrustedInstaller, L"Always run as TI");
  AppendMenuW(options_menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(options_menu, MF_STRING, cmd::kOptionsOpenDefaultRegedit, L"Open Default Regedit");
  UINT replace_flags = MF_STRING | ((is_elevated || is_system || is_ti) ? 0 : MF_GRAYED);
  AppendMenuW(options_menu, replace_flags | (replace_regedit_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsReplaceRegedit, L"Replace Regedit");
  AppendMenuW(options_menu, MF_STRING | (single_instance_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsSingleInstance, L"Single Instance");
  AppendMenuW(options_menu, MF_STRING | (save_tabs_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsSaveTabs, L"Save Tabs");
  AppendMenuW(options_menu, MF_STRING | (read_only_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsReadOnly, L"Read Only Mode");
  AppendMenuW(options_menu, MF_STRING | (save_tree_state_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewSaveTreeState, L"Save Previous Tree State");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(options_menu), L"Options");

  HMENU favorites_menu = CreatePopupMenu();
  AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesAdd, L"Add to Favorites...");
  AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesRemove, L"Remove Favorite");
  AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesEdit, L"Edit Favorites...");
  AppendMenuW(favorites_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(favorites_menu, MF_STRING, cmd::kFavoritesImport, L"Import Favorites...");
  append_menu(favorites_menu, MF_STRING, cmd::kFavoritesImportRegedit, L"Import Regedit Favorites");
  append_menu(favorites_menu, MF_STRING, cmd::kFavoritesExport, L"Export Favorites...");
  if (!favorites_cache_.empty()) {
    AppendMenuW(favorites_menu, MF_SEPARATOR, 0, nullptr);
    int limit = std::min(static_cast<int>(favorites_cache_.size()), cmd::kFavoritesItemMax - cmd::kFavoritesItemBase + 1);
    for (int i = 0; i < limit; ++i) {
      AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesItemBase + i, favorites_cache_[static_cast<size_t>(i)].c_str());
    }
  }
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(favorites_menu), L"Favorites");

  HMENU window_menu = CreatePopupMenu();
  AppendMenuW(window_menu, MF_STRING | (single_instance_ ? MF_GRAYED : 0), cmd::kWindowNew, L"New Window");
  AppendMenuW(window_menu, MF_STRING, cmd::kWindowClose, L"Close Window");
  AppendMenuW(window_menu, MF_STRING | (always_on_top_ ? MF_CHECKED : MF_UNCHECKED), cmd::kWindowAlwaysOnTop, L"Always on Top");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(window_menu), L"Window");

  HMENU trace_menu = CreatePopupMenu();
  auto has_label = [&](const wchar_t* label) -> bool {
    for (const auto& trace : active_traces_) {
      if (_wcsicmp(trace.label.c_str(), label) == 0) {
        return true;
      }
    }
    return false;
  };
  auto has_path = [&](const std::wstring& path) -> bool {
    for (const auto& trace : active_traces_) {
      if (EqualsInsensitive(trace.source_path, path)) {
        return true;
      }
    }
    return false;
  };
  bool trace_23h2 = has_label(L"23H2");
  bool trace_24h2 = has_label(L"24H2");
  bool trace_25h2 = has_label(L"25H2");
  bool has_recent_trace = false;
  append_menu(trace_menu, MF_STRING | (trace_23h2 ? MF_CHECKED : MF_UNCHECKED), cmd::kTraceLoad23H2, L"23H2");
  append_menu(trace_menu, MF_STRING | (trace_24h2 ? MF_CHECKED : MF_UNCHECKED), cmd::kTraceLoad24H2, L"24H2");
  append_menu(trace_menu, MF_STRING | (trace_25h2 ? MF_CHECKED : MF_UNCHECKED), cmd::kTraceLoad25H2, L"25H2");
  int recent_limit = std::min(static_cast<int>(recent_trace_paths_.items().size()), cmd::kTraceRecentMax - cmd::kTraceRecentBase + 1);
  for (int i = 0; i < recent_limit; ++i) {
    const std::wstring& path = recent_trace_paths_.items()[static_cast<size_t>(i)];
    if (path.empty()) {
      continue;
    }
    has_recent_trace = true;
    std::wstring name = FileNameOnly(path);
    if (name.empty()) {
      name = L"Trace";
    }
    UINT flags = MF_STRING;
    if (has_path(path)) {
      flags |= MF_CHECKED;
    }
    append_menu(trace_menu, flags, cmd::kTraceRecentBase + i, name.c_str());
  }
  AppendMenuW(trace_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(trace_menu, MF_STRING, cmd::kTraceGuide, L"Guide");
  append_menu(trace_menu, MF_STRING, cmd::kTraceLoadCustom, L"Open Trace File...");
  UINT edit_recent_flags = MF_STRING | (has_recent_trace ? 0 : MF_GRAYED);
  append_menu(trace_menu, edit_recent_flags, cmd::kTraceEditRecent, L"Edit Recent Traces...");
  append_menu(trace_menu, MF_STRING, cmd::kTraceEditActive, L"Edit Active Traces...");
  UINT clear_trace_flags = MF_STRING | (!active_traces_.empty() ? 0 : MF_GRAYED);
  append_menu(trace_menu, clear_trace_flags, cmd::kTraceClear, L"Clear Trace");

  HMENU default_menu = CreatePopupMenu();
  auto has_default_path = [&](const std::wstring& path) -> bool {
    for (const auto& defaults : active_defaults_) {
      if (EqualsInsensitive(defaults.source_path, path)) {
        return true;
      }
    }
    return false;
  };
  for (size_t i = 0; i < bundled_defaults_.size(); ++i) {
    const auto& entry = bundled_defaults_[i];
    UINT flags = MF_STRING;
    if (has_default_path(entry.path)) {
      flags |= MF_CHECKED;
    }
    append_menu(default_menu, flags, cmd::kDefaultBundledBase + static_cast<int>(i), entry.label.c_str());
  }
  bool has_recent_default = false;
  int default_recent_limit = std::min(static_cast<int>(recent_default_paths_.items().size()), cmd::kDefaultRecentMax - cmd::kDefaultRecentBase + 1);
  for (int i = 0; i < default_recent_limit; ++i) {
    const std::wstring& path = recent_default_paths_.items()[static_cast<size_t>(i)];
    if (path.empty()) {
      continue;
    }
    has_recent_default = true;
    std::wstring name = FileNameOnly(path);
    if (name.empty()) {
      name = L"Default";
    }
    UINT flags = MF_STRING;
    if (has_default_path(path)) {
      flags |= MF_CHECKED;
    }
    append_menu(default_menu, flags, cmd::kDefaultRecentBase + i, name.c_str());
  }
  AppendMenuW(default_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(default_menu, MF_STRING, cmd::kDefaultLoadCustom, L"Open Default File...");
  UINT edit_default_recent_flags = MF_STRING | (has_recent_default ? 0 : MF_GRAYED);
  append_menu(default_menu, edit_default_recent_flags, cmd::kDefaultEditRecent, L"Edit Recent Defaults...");
  append_menu(default_menu, MF_STRING, cmd::kDefaultEditActive, L"Edit Active Defaults...");
  UINT clear_default_flags = MF_STRING | (!active_defaults_.empty() ? 0 : MF_GRAYED);
  append_menu(default_menu, clear_default_flags, cmd::kDefaultClear, L"Clear Defaults");

  HMENU help_menu = CreatePopupMenu();
  AppendMenuW(help_menu, MF_STRING, cmd::kHelpContents, L"Help");
  AppendMenuW(help_menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(help_menu, MF_STRING, cmd::kHelpAbout, L"About RegKit");

  HMENU research_menu = CreatePopupMenu();
  int research_command = cmd::kResearchItemBase;
  for (const auto& link : frame::ResearchLinks()) {
    AppendMenuW(research_menu, MF_STRING, research_command++, link.name);
    if (link.separator_after) {
      AppendMenuW(research_menu, MF_SEPARATOR, 0, nullptr);
    }
  }

  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(trace_menu), L"Trace");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(default_menu), L"Default");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(research_menu), L"Research");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(help_menu), L"Help");

  PrepareMenusForOwnerDraw(menu, true);

  HMENU old_menu = GetMenu(hwnd_);
  SetMenu(hwnd_, menu);
  DrawMenuBar(hwnd_);
  if (old_menu) {
    DestroyMenu(old_menu);
  }
}

void MainWindow::Impl::RefreshFavoritesCache() {
  favorites_cache_.clear();
  FavoritesStore::Load(&favorites_cache_);
  favorites_loaded_ = true;
}

void MainWindow::Impl::RefreshBundledDefaultsCache() {
  bundled_defaults_.clear();
  bundled_defaults_loaded_ = true;

  std::wstring module_dir = util::GetModuleDirectory();
  if (module_dir.empty()) {
    return;
  }
  std::wstring assets = util::JoinPath(module_dir, L"assets");
  std::wstring defaults_dir = util::JoinPath(assets, L"defaults");
  std::wstring pattern = util::JoinPath(defaults_dir, L"*.reg");
  WIN32_FIND_DATAW data = {};
  HANDLE find = FindFirstFileW(pattern.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
      continue;
    }
    std::wstring file_name = data.cFileName;
    std::wstring label = FileBaseName(file_name);
    if (label.empty()) {
      continue;
    }
    BundledDefault entry;
    entry.label = std::move(label);
    entry.path = util::JoinPath(defaults_dir, file_name);
    bundled_defaults_.push_back(std::move(entry));
  } while (FindNextFileW(find, &data));
  FindClose(find);

  std::sort(bundled_defaults_.begin(), bundled_defaults_.end(), [](const BundledDefault& left, const BundledDefault& right) { return _wcsicmp(left.label.c_str(), right.label.c_str()) < 0; });
  size_t bundled_limit = std::min(bundled_defaults_.size(), static_cast<size_t>(cmd::kDefaultBundledMax - cmd::kDefaultBundledBase + 1));
  if (bundled_defaults_.size() > bundled_limit) {
    bundled_defaults_.resize(bundled_limit);
  }
}

bool MainWindow::Impl::HandleMenuCommand(int command_id) {
  frame::CommandContext context;
  context.context = this;
  context.dynamic = [](void* value, int id) {
    return static_cast<MainWindow::Impl*>(value)->HandleDynamicCommand(id);
  };
  context.file = [](void* value, int id) {
    return static_cast<MainWindow::Impl*>(value)->HandleFileCommand(id);
  };
  context.view = [](void* value, int id) {
    return static_cast<MainWindow::Impl*>(value)->HandleViewCommand(id);
  };
  context.trace_defaults = [](void* value, int id) {
    return static_cast<MainWindow::Impl*>(value)->HandleTraceDefaultCommand(id);
  };
  context.workspace_appearance = [](void* value, int id) {
    return static_cast<MainWindow::Impl*>(value)->HandleWorkspaceAppearanceCommand(id);
  };
  context.navigate_clipboard = [](void* value, int id) {
    return static_cast<MainWindow::Impl*>(value)->HandleNavigateClipboardCommand(id);
  };
  context.mutation = [](void* value, int id) {
    return static_cast<MainWindow::Impl*>(value)->HandleMutationCommand(id);
  };
  return frame::DispatchCommand(command_id, context);
}

} // namespace regkit
