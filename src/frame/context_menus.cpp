// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"

namespace regkit {
using namespace command_detail;

void MainWindow::Impl::RecordNavigation(const std::wstring& path) {
  if (browse_.RecordNavigation(path)) {
    UpdateNavigationButtons();
  }
}

void MainWindow::Impl::NavigateBack() {
  if (const auto target = browse_.Back()) {
    SelectTreePath(*target);
    UpdateNavigationButtons();
  }
}

void MainWindow::Impl::NavigateForward() {
  if (const auto target = browse_.Forward()) {
    SelectTreePath(*target);
    UpdateNavigationButtons();
  }
}

void MainWindow::Impl::NavigateUp() {
  if (const auto target = browse_.Up()) {
    SelectTreePath(*target);
    UpdateNavigationButtons();
  }
}

void MainWindow::Impl::UpdateNavigationButtons() {
  if (!toolbar_.hwnd()) {
    return;
  }
  const browse::NavigationAvailability available = browse_.navigation();
  SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kNavBack,
               available.back ? TBSTATE_ENABLED : 0);
  SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kNavForward,
               available.forward ? TBSTATE_ENABLED : 0);
  SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kNavUp,
               available.up ? TBSTATE_ENABLED : 0);
}

void MainWindow::Impl::ShowAddressContextMenu(HWND edit, POINT screen_pt) {
  if (!edit) {
    return;
  }
  enum AddressCommand {
    kAddressUndo = 1,
    kAddressCut,
    kAddressCopy,
    kAddressPaste,
    kAddressPasteGo,
    kAddressDelete,
    kAddressSelectAll,
  };

  DWORD selection_start = 0;
  DWORD selection_end = 0;
  SendMessageW(edit, EM_GETSEL, reinterpret_cast<WPARAM>(&selection_start), reinterpret_cast<LPARAM>(&selection_end));
  const bool has_selection = selection_end > selection_start;
  const bool writable = (GetWindowLongPtrW(edit, GWL_STYLE) & ES_READONLY) == 0;
  const bool can_undo = SendMessageW(edit, EM_CANUNDO, 0, 0) != 0;
  const bool has_text = GetWindowTextLengthW(edit) > 0;
  const bool has_clipboard_text = IsClipboardFormatAvailable(CF_UNICODETEXT) != FALSE;

  HMENU menu = CreatePopupMenu();
  auto append = [&](bool enabled, int command, const wchar_t* text, const wchar_t* shortcut) {
    std::wstring label = text;
    label.push_back(L'\t');
    label.append(shortcut);
    AppendMenuW(menu, MF_STRING | (enabled ? MF_ENABLED : MF_GRAYED), static_cast<UINT_PTR>(command), label.c_str());
  };

  append(can_undo && writable, kAddressUndo, L"Undo", L"Ctrl+Z");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  append(has_selection && writable, kAddressCut, L"Cut", L"Ctrl+X");
  append(has_selection, kAddressCopy, L"Copy", L"Ctrl+C");
  append(has_clipboard_text && writable, kAddressPaste, L"Paste", L"Ctrl+V");
  append(has_clipboard_text && writable, kAddressPasteGo, L"Paste and Go", L"Ctrl+Shift+V");
  append(has_selection && writable, kAddressDelete, L"Delete", L"Del");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  append(has_text, kAddressSelectAll, L"Select All", L"Ctrl+A");

  const int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  switch (command) {
  case kAddressUndo:
    SendMessageW(edit, WM_UNDO, 0, 0);
    break;
  case kAddressCut:
    SendMessageW(edit, WM_CUT, 0, 0);
    break;
  case kAddressCopy:
    SendMessageW(edit, WM_COPY, 0, 0);
    break;
  case kAddressPaste:
    SendMessageW(edit, WM_PASTE, 0, 0);
    break;
  case kAddressPasteGo:
    SendMessageW(edit, EM_SETSEL, 0, -1);
    SendMessageW(edit, WM_PASTE, 0, 0);
    SendMessageW(edit, WM_KEYDOWN, VK_RETURN, 0);
    break;
  case kAddressDelete:
    SendMessageW(edit, WM_CLEAR, 0, 0);
    break;
  case kAddressSelectAll:
    SendMessageW(edit, EM_SETSEL, 0, -1);
    break;
  default:
    break;
  }
}

void MainWindow::Impl::ShowTreeContextMenu(POINT screen_pt) {
  if (!browse_.tree().hwnd()) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(browse_.tree().hwnd(), &client_pt);
  TVHITTESTINFO hit = {};
  hit.pt = client_pt;
  HTREEITEM item = TreeView_HitTest(browse_.tree().hwnd(), &hit);
  if (item) {
    TreeView_SelectItem(browse_.tree().hwnd(), item);
  }
  SetFocus(browse_.tree().hwnd());
  HTREEITEM target = item ? item : TreeView_GetSelection(browse_.tree().hwnd());
  RegistryNode* node = browse_.tree().NodeFromItem(target);

  HMENU menu = CreatePopupMenu();
  bool has_node = node != nullptr;
  bool can_rename = has_node && !node->subkey.empty();
  bool is_simulated = has_node && node->simulated;
  bool can_modify = !read_only_;
  UINT edit_flags = MF_STRING | (has_node ? 0 : MF_GRAYED);
  UINT modify_flags = MF_STRING | ((has_node && can_modify) ? 0 : MF_GRAYED);
  UINT rename_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);
  UINT delete_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);

  bool expanded = false;
  bool can_toggle = false;
  bool has_children = false;
  if (target) {
    TVITEMW tvi = {};
    tvi.hItem = target;
    tvi.mask = TVIF_STATE | TVIF_CHILDREN;
    tvi.stateMask = TVIS_EXPANDED;
    if (TreeView_GetItem(browse_.tree().hwnd(), &tvi)) {
      expanded = (tvi.state & TVIS_EXPANDED) != 0;
      has_children = TreeView_GetChild(browse_.tree().hwnd(), target) != nullptr || tvi.cChildren != 0;
      can_toggle = expanded || has_children;
    }
  }
  std::wstring expand_label = expanded ? L"Collapse Key" : L"Expand Key";
  UINT expand_flags = MF_STRING | (can_toggle ? 0 : MF_GRAYED);
  UINT expand_all_flags = MF_STRING | (has_children ? 0 : MF_GRAYED);
  auto equals_insensitive = [](const std::wstring& left, const wchar_t* right) -> bool { return _wcsicmp(left.c_str(), right) == 0; };
  bool can_open_hive = false;
  if (has_node) {
    bool is_root = false;
    std::wstring hive_path = LookupHivePath(*node, &is_root);
    if (!hive_path.empty() && is_root) {
      if (node->subkey.empty() && (node->root == HKEY_CURRENT_USER || equals_insensitive(node->root_name, L"HKEY_CURRENT_USER"))) {
        can_open_hive = false;
      } else {
        can_open_hive = true;
      }
    }
  }

  HMENU new_value = CreatePopupMenu();
  AppendMenuW(new_value, MF_STRING, cmd::kNewString, L"String Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewBinary, L"Binary Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");

  AppendMenuW(menu, edit_flags, cmd::kEditCopyKey, L"Copy Key Name");
  AppendMenuW(menu, edit_flags, cmd::kEditCopyKeyPath, L"Copy Key Path");
  AppendMenuW(menu, MF_POPUP | (has_node ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
  if (!is_simulated) {
    AppendMenuW(menu, modify_flags, cmd::kEditPermissions, L"Permissions...");
    if (can_open_hive) {
      AppendMenuW(menu, MF_STRING, cmd::kOptionsHiveFileDir, L"On-Disk Hive File");
    }
  }
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, expand_flags, cmd::kTreeToggleExpand, expand_label.c_str());
  AppendMenuW(menu, expand_all_flags, cmd::kTreeExpandAll, L"Expand All Subkeys");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  if (is_simulated) {
    AppendMenuW(menu, modify_flags, cmd::kCreateSimulatedKey, L"Create Key");
  } else {
    AppendMenuW(menu, modify_flags, cmd::kNewKey, L"New Key");
    AppendMenuW(menu, MF_POPUP | ((has_node && can_modify) ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(new_value), L"New Value");
  }
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  if (!is_simulated) {
    AppendMenuW(menu, edit_flags, cmd::kFileExport, L"Export...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  }
  AppendMenuW(menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
  if (!is_simulated) {
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, rename_flags, cmd::kEditRename, L"Rename");
    AppendMenuW(menu, delete_flags, cmd::kEditDelete, L"Delete");
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command != 0) {
    HandleMenuCommand(command);
  }
}

void MainWindow::Impl::ShowValueContextMenu(POINT screen_pt) {
  if (!browse_.values().hwnd()) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(browse_.values().hwnd(), &client_pt);
  LVHITTESTINFO hit = {};
  hit.pt = client_pt;
  int index = ListView_HitTest(browse_.values().hwnd(), &hit);
  const ListRow* row = nullptr;
  if (index >= 0) {
    const UINT state = ListView_GetItemState(browse_.values().hwnd(), index, LVIS_SELECTED);
    if ((state & LVIS_SELECTED) == 0) {
      ListView_SetItemState(browse_.values().hwnd(), -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    }
    ListView_SetItemState(browse_.values().hwnd(), index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    row = browse_.values().RowAt(index);
  }
  SetFocus(browse_.values().hwnd());

  HMENU menu = CreatePopupMenu();
  if (row && row->kind == rowkind::kKey) {
    auto equals_insensitive = [](const std::wstring& left, const wchar_t* right) -> bool { return _wcsicmp(left.c_str(), right) == 0; };
    bool is_simulated = row->simulated;
    bool can_rename = !row->extra.empty();
    bool can_modify = !read_only_;
    UINT edit_flags = MF_STRING;
    UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
    UINT rename_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);
    UINT delete_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);
    UINT expand_flags = MF_STRING | MF_GRAYED;
    std::wstring expand_label = L"Expand Key";
    bool can_open_hive = false;
    if (browse_.current_node()) {
      RegistryNode target = *browse_.current_node();
      if (!row->extra.empty()) {
        target = MakeChildNode(*browse_.current_node(), row->extra);
      }
      bool is_root = false;
      std::wstring hive_path = LookupHivePath(target, &is_root);
      if (!hive_path.empty() && is_root) {
        if (target.subkey.empty() && (target.root == HKEY_CURRENT_USER || equals_insensitive(target.root_name, L"HKEY_CURRENT_USER"))) {
          can_open_hive = false;
        } else {
          can_open_hive = true;
        }
      }
    }

    HMENU new_value = CreatePopupMenu();
    AppendMenuW(new_value, MF_STRING, cmd::kNewString, L"String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewBinary, L"Binary Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");

    AppendMenuW(menu, edit_flags, cmd::kEditCopyKey, L"Copy Key Name");
    AppendMenuW(menu, edit_flags, cmd::kEditCopyKeyPath, L"Copy Key Path");
    AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
    if (!is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kEditPermissions, L"Permissions...");
      if (can_open_hive) {
        AppendMenuW(menu, MF_STRING, cmd::kOptionsHiveFileDir, L"On-Disk Hive File");
      }
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, expand_flags, cmd::kTreeToggleExpand, expand_label.c_str());
    AppendMenuW(menu, expand_flags, cmd::kTreeExpandAll, L"Expand All Subkeys");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kCreateSimulatedKey, L"Create Key");
    } else {
      AppendMenuW(menu, modify_flags, cmd::kNewKey, L"New Key");
      AppendMenuW(menu, MF_POPUP | (can_modify ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(new_value), L"New Value");
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (!is_simulated) {
      AppendMenuW(menu, edit_flags, cmd::kFileExport, L"Export...");
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
    if (!is_simulated) {
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
      AppendMenuW(menu, rename_flags, cmd::kEditRename, L"Rename");
      AppendMenuW(menu, delete_flags, cmd::kEditDelete, L"Delete");
    }
  } else if (row && row->kind == rowkind::kValue) {
    std::vector<ListRow> selected_rows = SelectedListRows(browse_.values());
    const bool all_values = !selected_rows.empty() &&
                            std::all_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& selected) {
                              return selected.kind == rowkind::kValue && !selected.simulated;
                            });
    const bool single_value = all_values && selected_rows.size() == 1;
    bool can_modify = !read_only_ && single_value;
    bool can_delete = !read_only_ && all_values;
    bool can_comment = all_values;
    bool can_export = !row->simulated && browse_.current_node() && !browse_.current_node()->simulated;
    UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
    UINT delete_flags = MF_STRING | (can_delete ? 0 : MF_GRAYED);
    UINT single_flags = MF_STRING | (single_value ? 0 : MF_GRAYED);
    UINT export_flags = MF_STRING | (can_export ? 0 : MF_GRAYED);
    UINT comment_flags = MF_STRING | (can_comment ? 0 : MF_GRAYED);
    AppendMenuW(menu, modify_flags, cmd::kEditModify, L"Modify...");
    AppendMenuW(menu, modify_flags, cmd::kEditModifyBinary, L"Modify Binary Data...");
    AppendMenuW(menu, comment_flags, cmd::kEditModifyComment, L"Modify Comment...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, single_flags, cmd::kEditCopyValueName, L"Copy Value Name");
    AppendMenuW(menu, single_flags, cmd::kEditCopyValueData, L"Copy Value Data");
    AppendMenuW(menu, export_flags, cmd::kFileExport, L"Export...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, modify_flags, cmd::kEditRename, L"Rename");
    AppendMenuW(menu, delete_flags, cmd::kEditDelete, L"Delete");
  } else {
    bool is_simulated = browse_.current_node() && browse_.current_node()->simulated;
    bool can_modify = !read_only_;
    UINT edit_flags = MF_STRING | (browse_.current_node() ? 0 : MF_GRAYED);
    UINT modify_flags = MF_STRING | ((browse_.current_node() && can_modify) ? 0 : MF_GRAYED);
    HMENU new_value = CreatePopupMenu();
    AppendMenuW(new_value, MF_STRING, cmd::kNewString, L"String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewBinary, L"Binary Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");

    AppendMenuW(menu, edit_flags, cmd::kEditCopyKey, L"Copy Key Name");
    AppendMenuW(menu, edit_flags, cmd::kEditCopyKeyPath, L"Copy Key Path");
    AppendMenuW(menu, MF_POPUP | (browse_.current_node() ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kCreateSimulatedKey, L"Create Key");
    } else {
      AppendMenuW(menu, modify_flags, cmd::kNewKey, L"New Key");
      AppendMenuW(menu, MF_POPUP | ((browse_.current_node() && can_modify) ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(new_value), L"New Value");
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (!is_simulated) {
      AppendMenuW(menu, edit_flags, cmd::kFileExport, L"Export...");
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (!is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kEditPermissions, L"Permissions...");
    }
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command != 0) {
    HandleMenuCommand(command);
  }
}

void MainWindow::Impl::ShowHistoryContextMenu(POINT screen_pt) {
  if (!history_list_) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(history_list_, &client_pt);
  LVHITTESTINFO hit = {};
  hit.pt = client_pt;
  int index = ListView_HitTest(history_list_, &hit);
  if (index >= 0) {
    if ((ListView_GetItemState(history_list_, index, LVIS_SELECTED) &
         LVIS_SELECTED) == 0) {
      ListView_SetItemState(history_list_, -1, 0,
                            LVIS_SELECTED | LVIS_FOCUSED);
    }
    ListView_SetItemState(history_list_, index,
                          LVIS_SELECTED | LVIS_FOCUSED,
                          LVIS_SELECTED | LVIS_FOCUSED);
  }
  const HistoryEntry* entry = nullptr;
  if (index >= 0) {
    LVITEMW item = {};
    item.mask = LVIF_PARAM;
    item.iItem = index;
    if (ListView_GetItem(history_list_, &item)) {
      size_t history_index = static_cast<size_t>(item.lParam);
      if (history_index < change_history_.entries().size()) {
        entry = &change_history_.entries()[history_index];
      }
    }
  }

  HMENU menu = CreatePopupMenu();
  HistoryEntry prepared_revert;
  const bool can_revert = entry && PrepareHistoryRevert(*entry, &prepared_revert);
  UINT open_flags = MF_STRING | ((entry && !entry->key_path.empty()) ? 0 : MF_GRAYED);
  UINT revert_flags = MF_STRING | (can_revert ? 0 : MF_GRAYED);
  AppendMenuW(menu, open_flags, cmd::kHistoryOpenTarget, L"Open Entry");
  AppendMenuW(menu, revert_flags, cmd::kHistoryRevert, L"Revert");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKey, L"Copy");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING, cmd::kEditDelete, L"Clear History");

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command == cmd::kHistoryOpenTarget && entry) {
    if (!OpenHistoryTarget(*entry)) {
      ui::ShowError(hwnd_, L"Failed to open history target.");
    }
  } else if (command == cmd::kHistoryRevert && entry) {
    RevertHistoryEntry(*entry);
  } else if (command == cmd::kEditCopyKey && entry) {
    const std::wstring combined = entry->time_text + L" | " + entry->action +
                                  L" | " + entry->old_data + L" | " +
                                  entry->new_data;
    ui::CopyTextToClipboard(hwnd_, combined);
  } else if (command == cmd::kEditDelete) {
    ClearHistoryItems(true);
  }
}

void MainWindow::Impl::ShowSearchResultContextMenu(POINT screen_pt) {
  if (!search_results_list_) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(search_results_list_, &client_pt);
  LVHITTESTINFO hit = {};
  hit.pt = client_pt;
  int index = ListView_HitTest(search_results_list_, &hit);
  if (index < 0) {
    return;
  }

  int sel_tab = TabCtrl_GetCurSel(tab_);
  int search_index = SearchIndexFromTab(sel_tab);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return;
  }
  if (static_cast<size_t>(index) >= SearchRowCount(search_index)) {
    return;
  }

  if ((ListView_GetItemState(search_results_list_, index, LVIS_SELECTED) &
       LVIS_SELECTED) == 0) {
    ListView_SetItemState(search_results_list_, -1, 0,
                          LVIS_SELECTED | LVIS_FOCUSED);
  }
  ListView_SetItemState(search_results_list_, index,
                        LVIS_SELECTED | LVIS_FOCUSED,
                        LVIS_SELECTED | LVIS_FOCUSED);
  std::wstring key_path = SearchRowKeyPath(search_index, index);
  // Null on compare tabs, where value operations do not apply.
  const search::Result* result = SearchResultAt(index);
  const bool is_key_row = !result || search::IsKeyRow(*result);
  if (key_path.empty()) {
    return;
  }

  RegistryNode node;
  bool node_ok = ResolvePathToNode(key_path, &node);
  KeyInfo info = {};
  bool key_exists = node_ok && RegistryStore::QueryKeyInfo(node, &info);
  bool can_modify = !read_only_;
  bool can_rename = key_exists && can_modify && (!is_key_row || !node.subkey.empty());
  bool can_delete = key_exists && can_modify && (!is_key_row || !node.subkey.empty());
  bool can_export = key_exists;
  bool can_permissions = key_exists && can_modify;
  bool can_open_hive = false;
  if (key_exists) {
    bool is_root = false;
    std::wstring hive_path = LookupHivePath(node, &is_root);
    if (!hive_path.empty() && is_root) {
      if (node.subkey.empty() && (node.root == HKEY_CURRENT_USER || EqualsInsensitive(node.root_name, L"HKEY_CURRENT_USER"))) {
        can_open_hive = false;
      } else {
        can_open_hive = true;
      }
    }
  }

  enum {
    kSearchOpenKey = 51000,
    kSearchOpenKeyNewTab = 51001,
    kSearchModify = 51002,
    kSearchModifyBinary = 51003,
    kSearchModifyComment = 51004,
    kSearchCopyKeyName = 51005,
    kSearchCopyKeyPath = 51006,
    kSearchCopyKeyPathAbbrev = 51013,
    kSearchCopyKeyPathRegedit = 51014,
    kSearchCopyKeyPathRegFile = 51015,
    kSearchCopyKeyPathPowerShell = 51016,
    kSearchCopyKeyPathPowerShellProvider = 51017,
    kSearchCopyKeyPathEscaped = 51018,
    kSearchPermissions = 51007,
    kSearchOpenHive = 51008,
    kSearchExport = 51009,
    kSearchRename = 51010,
    kSearchDelete = 51011,
  };

  auto build_copy_path_menu = [&]() -> HMENU {
    HMENU submenu = CreatePopupMenu();
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathAbbrev, L"Abbreviated (HKLM)");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathRegedit, L"Regedit Address Bar");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathRegFile, L".reg File Header");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathPowerShell, L"PowerShell Drive");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathPowerShellProvider, L"PowerShell Provider");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathEscaped, L"Escaped Backslashes");
    return submenu;
  };

  HMENU menu = CreatePopupMenu();
  AppendMenuW(menu, MF_STRING, kSearchOpenKey, L"Open Key");
  AppendMenuW(menu, MF_STRING, kSearchOpenKeyNewTab, L"Open Key in New Tab");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  if (!is_key_row) {
    UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
    AppendMenuW(menu, modify_flags, kSearchModify, L"Modify...");
    AppendMenuW(menu, modify_flags, kSearchModifyBinary, L"Modify Binary Data...");
    AppendMenuW(menu, MF_STRING, kSearchModifyComment, L"Modify Comment...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  }
  AppendMenuW(menu, MF_STRING, kSearchCopyKeyName, L"Copy Key Name");
  AppendMenuW(menu, MF_STRING, kSearchCopyKeyPath, L"Copy Key Path");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(build_copy_path_menu()), L"Copy Key Path As");
  UINT permissions_flags = MF_STRING | (can_permissions ? 0 : MF_GRAYED);
  AppendMenuW(menu, permissions_flags, kSearchPermissions, L"Permissions...");
  UINT open_hive_flags = MF_STRING | (can_open_hive ? 0 : MF_GRAYED);
  AppendMenuW(menu, open_hive_flags, kSearchOpenHive, L"On-Disk Hive File");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  UINT export_flags = MF_STRING | (can_export ? 0 : MF_GRAYED);
  AppendMenuW(menu, export_flags, kSearchExport, L"Export...");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  UINT rename_flags = MF_STRING | (can_rename ? 0 : MF_GRAYED);
  UINT delete_flags = MF_STRING | (can_delete ? 0 : MF_GRAYED);
  AppendMenuW(menu, rename_flags, kSearchRename, L"Rename");
  AppendMenuW(menu, delete_flags, kSearchDelete, L"Delete");

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);
  if (command == 0) {
    return;
  }

  auto open_key = [&](bool new_tab) {
    if (!tab_) {
      return;
    }
    if (new_tab) {
      OpenLocalRegistryTab();
    } else {
      ActivateRegistryTab();
    }
    ApplyViewVisibility();
    UpdateStatus();
    SelectTreePath(key_path);
  };
  auto focus_key = [&]() {
    open_key(false);
    if (browse_.tree().hwnd()) {
      SetFocus(browse_.tree().hwnd());
    }
  };
  auto run_on_value = [&](int command) {
    if (is_key_row) {
      return;
    }
    open_key(false);
    if (SelectValueByName(result->value_name)) {
      if (browse_.values().hwnd()) {
        SetFocus(browse_.values().hwnd());
      }
      HandleMenuCommand(command);
      return;
    }
    pending_external_value_key_path_ = key_path;
    pending_external_value_name_ = result->value_name;
    pending_value_command_ = command;
  };
  auto run_on_key = [&](int command) {
    focus_key();
    HandleMenuCommand(command);
  };
  auto run_on_target = [&](int command) {
    if (is_key_row) {
      run_on_key(command);
    } else {
      run_on_value(command);
    }
  };

  switch (command) {
  case kSearchOpenKey:
    open_key(SearchResultOpensInNewTab());
    return;
  case kSearchOpenKeyNewTab:
    open_key(true);
    return;
  case kSearchModify:
    run_on_value(cmd::kEditModify);
    return;
  case kSearchModifyBinary:
    run_on_value(cmd::kEditModifyBinary);
    return;
  case kSearchModifyComment:
    run_on_value(cmd::kEditModifyComment);
    return;
  case kSearchCopyKeyName: {
    std::wstring name;
    if (node_ok) {
      name = LeafName(node);
    } else {
      size_t pos = key_path.find_last_of(L"\\/");
      name = (pos == std::wstring::npos) ? key_path : key_path.substr(pos + 1);
    }
    if (!name.empty()) {
      ui::CopyTextToClipboard(hwnd_, name);
    }
    return;
  }
  case kSearchCopyKeyPath:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, key_path);
    }
    return;
  case kSearchCopyKeyPathAbbrev:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kAbbrev));
    }
    return;
  case kSearchCopyKeyPathRegedit:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kRegedit));
    }
    return;
  case kSearchCopyKeyPathRegFile:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kRegFile));
    }
    return;
  case kSearchCopyKeyPathPowerShell:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kPowerShellDrive));
    }
    return;
  case kSearchCopyKeyPathPowerShellProvider:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kPowerShellProvider));
    }
    return;
  case kSearchCopyKeyPathEscaped:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kEscaped));
    }
    return;
  case kSearchPermissions:
    if (node_ok && key_exists) {
      ShowPermissionsDialog(node);
    }
    return;
  case kSearchOpenHive:
    if (can_open_hive) {
      focus_key();
      HandleMenuCommand(cmd::kOptionsHiveFileDir);
    }
    return;
  case kSearchExport:
    if (can_export) {
      run_on_target(cmd::kFileExport);
    }
    return;
  case kSearchRename:
    if (can_rename) {
      run_on_target(cmd::kEditRename);
    }
    return;
  case kSearchDelete:
    if (can_delete) {
      run_on_target(cmd::kEditDelete);
    }
    return;
  default:
    return;
  }
}

} // namespace regkit
