// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"

namespace regkit {
using namespace command_detail;

namespace {

std::wstring NativeTargetPath(const std::wstring& path) {
  const std::wstring normalized =
      registry_path::Normalize(path, util::GetCurrentUserSidString());
  if (normalized.empty()) {
    return {};
  }
  const size_t split = normalized.find(L'\\');
  const std::wstring root =
      split == std::wstring::npos ? normalized : normalized.substr(0, split);
  const std::wstring rest =
      split == std::wstring::npos ? std::wstring() : normalized.substr(split);
  std::wstring native;
  if (_wcsicmp(root.c_str(), L"HKEY_LOCAL_MACHINE") == 0) {
    native = L"\\REGISTRY\\MACHINE";
  } else if (_wcsicmp(root.c_str(), L"HKEY_USERS") == 0) {
    native = L"\\REGISTRY\\USER";
  } else if (_wcsicmp(root.c_str(), L"HKEY_CURRENT_USER") == 0) {
    const std::wstring sid = util::GetCurrentUserSidString();
    if (sid.empty()) {
      return {};
    }
    native = L"\\REGISTRY\\USER\\" + sid;
  } else if (_wcsicmp(root.c_str(), L"HKEY_CLASSES_ROOT") == 0) {
    native = L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes";
  } else if (_wcsicmp(root.c_str(), L"HKEY_CURRENT_CONFIG") == 0) {
    native =
        L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Hardware Profiles\\Current";
  } else {
    return {};
  }
  return native + rest;
}

} // namespace

bool MainWindow::Impl::HandleMutationCommand(int command_id) {
  if (command_id == cmd::kEditResetDefault ||
      (command_id >= cmd::kResetDefaultBase &&
       command_id <= cmd::kResetDefaultMax)) {
    return HandleResetDefaultCommand(command_id);
  }
  switch (command_id) {
  case cmd::kNewSymbolicLink: {
    if (!EnsureWritable() || !browse_.current_node()) {
      return true;
    }
    if (registry_mode_ != RegistryMode::kLocal) {
      ui::ShowWarning(hwnd_, L"Symbolic links can only be created in the local registry.");
      return true;
    }
    editors::SymbolicLinkResult link;
    if (!editors::PromptSymbolicLink(
            hwnd_, MakeUniqueKeyName(*browse_.current_node(), L"New Link"),
            [](HWND owner, std::wstring* selected) {
              return ShowBrowseKeyDialog(owner, selected);
            },
            &link)) {
      return true;
    }
    const std::wstring native = NativeTargetPath(link.target);
    if (native.empty()) {
      ui::ShowError(hwnd_, L"The target must be a full registry path under a standard root.");
      return true;
    }
    DWORD error = ERROR_SUCCESS;
    if (!registry_backend::live::CreateRegistryLink(*browse_.current_node(),
                                                    link.name, native, &error)) {
      std::wstring message = L"Failed to create the symbolic link.";
      const std::wstring detail = util::FormatWin32Error(error);
      if (!detail.empty()) {
        message += L"\n";
        message += detail;
      }
      ui::ShowError(hwnd_, message);
      return true;
    }
    AppendHistoryEntry(L"Create symbolic link " + link.name, L"", native);
    MarkOfflineDirty();
    RefreshTreeSelection();
    RefreshMatchingTreeNodes();
    UpdateValueListForNode(browse_.current_node());
    return true;
  }
  case cmd::kCreateSimulatedKey:
  case cmd::kNewKey:
  case cmd::kNewString:
  case cmd::kNewExpandString:
  case cmd::kNewBinary:
  case cmd::kNewDword:
  case cmd::kNewQword:
  case cmd::kNewMultiString:
    return HandleCreateCommand(command_id);
  case cmd::kEditModify:
  case cmd::kEditModifyBinary:
  case cmd::kEditChangeType:
  case cmd::kEditModifyComment:
    return HandleModifyCommand(command_id);
  case cmd::kEditRename:
    return HandleRenameCommand(command_id);
  case cmd::kEditDelete:
    return HandleDeleteCommand(command_id);
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleCreateCommand(int command_id) {
  switch (command_id) {
  case cmd::kCreateSimulatedKey: {
    if (!EnsureWritable()) {
      return true;
    }
    RegistryNode target;
    bool has_target = false;
    const ListRow* row = SelectedValueRow(browse_.values(), nullptr);
    if (row && row->kind == rowkind::kKey && row->simulated && browse_.current_node()) {
      target = MakeChildNode(*browse_.current_node(), row->extra);
      has_target = true;
    } else if (browse_.current_node() && browse_.current_node()->simulated) {
      target = *browse_.current_node();
      has_target = true;
    }
    if (!has_target) {
      return true;
    }
    std::wstring path = registry_path::Build(target);
    if (!CreateRegistryPath(path)) {
      ui::ShowError(hwnd_, L"Failed to create the key.");
      return true;
    }
    UpdateSimulatedChain(TreeView_GetSelection(browse_.tree().hwnd()));
    RefreshTreeSelection();
    RefreshMatchingTreeNodes();
    UpdateValueListForNode(browse_.current_node());
    return true;
  }
  case cmd::kNewKey: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!browse_.current_node()) {
      return true;
    }
    std::wstring name = MakeUniqueKeyName(*browse_.current_node(), L"New Key");
    if (name.empty()) {
      return true;
    }
    if (!RegistryStore::CreateKey(*browse_.current_node(), name)) {
      ui::ShowError(hwnd_, L"Failed to create key.");
    } else {
      AppendHistoryEntry(L"Create key " + name, L"", L"");
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kCreateKey;
      op.node = *browse_.current_node();
      op.name = name;
      op.key_snapshot.name = name;
      PushUndo(std::move(op));
      std::wstring path = registry_path::Build(*browse_.current_node());
      if (!path.empty()) {
        path.append(L"\\");
        path.append(name);
      }
      HWND focus = GetFocus();
      bool edit_in_list = (focus == browse_.values().hwnd()) && show_keys_in_list_ && browse_.values().hwnd();
      if (edit_in_list) {
        ScheduleValueListRename(rowkind::kKey, name);
        UpdateValueListForNode(browse_.current_node());
      } else {
        std::wstring parent_path = registry_path::Build(*browse_.current_node());
        HTREEITEM parent_item = TreeView_GetSelection(browse_.tree().hwnd());
        if (!parent_item && !parent_path.empty()) {
          SelectTreePath(parent_path);
          parent_item = TreeView_GetSelection(browse_.tree().hwnd());
        }

        if (parent_item) {
          TreeView_SelectItem(browse_.tree().hwnd(), parent_item);
        }
        RefreshTreeSelection();
        RefreshMatchingTreeNodes();

        HTREEITEM target = nullptr;
        if (parent_item) {
          target = FindChildByText(browse_.tree().hwnd(), parent_item, name);
          if (target) {
            TreeView_SelectItem(browse_.tree().hwnd(), target);
            TreeView_EnsureVisible(browse_.tree().hwnd(), target);
          }
        }
        if (!target && !path.empty()) {
          if (SelectTreePath(path)) {
            target = TreeView_GetSelection(browse_.tree().hwnd());
          }
        }
        if (target) {
          SetFocus(browse_.tree().hwnd());
          TreeView_EditLabel(browse_.tree().hwnd(), target);
        }
        UpdateValueListForNode(browse_.current_node());
      }
    }
    return true;
  }
  case cmd::kNewString:
  case cmd::kNewExpandString:
  case cmd::kNewBinary:
  case cmd::kNewDword:
  case cmd::kNewQword:
  case cmd::kNewMultiString: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!browse_.current_node()) {
      return true;
    }
    DWORD type = REG_SZ;
    std::wstring base_name = L"New Value";
    switch (command_id) {
    case cmd::kNewExpandString:
      type = REG_EXPAND_SZ;
      base_name = L"New Expandable String Value";
      break;
    case cmd::kNewBinary:
      type = REG_BINARY;
      base_name = L"New Binary Value";
      break;
    case cmd::kNewDword:
      type = REG_DWORD;
      base_name = L"New DWORD Value";
      break;
    case cmd::kNewQword:
      type = REG_QWORD;
      base_name = L"New QWORD Value";
      break;
    case cmd::kNewMultiString:
      type = REG_MULTI_SZ;
      base_name = L"New Multi-String Value";
      break;
    default:
      type = REG_SZ;
      break;
    }
    std::wstring value_name = MakeUniqueValueName(*browse_.current_node(), base_name);
    if (value_name.empty()) {
      return true;
    }
    std::vector<BYTE> data;
    if (type == REG_SZ || type == REG_EXPAND_SZ) {
      data.resize(sizeof(wchar_t), 0);
    } else if (type == REG_MULTI_SZ) {
      data.resize(sizeof(wchar_t) * 2, 0);
    } else if (type == REG_DWORD) {
      data.resize(sizeof(DWORD), 0);
    } else if (type == REG_QWORD) {
      data.resize(sizeof(unsigned long long), 0);
    }
    if (!RegistryStore::SetValue(*browse_.current_node(), value_name, type, data)) {
      ui::ShowError(hwnd_, L"Failed to set value.");
    } else {
      std::wstring data_text = value_format::Data(type, data.data(), static_cast<DWORD>(data.size()));
      AppendValueHistoryEntry(L"Create value " + value_name, L"", data_text, *browse_.current_node(), value_name, HistoryEntry::RevertKind::kDeleteValue);
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kCreateValue;
      op.node = *browse_.current_node();
      op.name = value_name;
      op.new_value.name = value_name;
      op.new_value.type = type;
      op.new_value.data = data;
      PushUndo(std::move(op));

      value_list_generation_.fetch_add(1);
      value_list_loading_ = false;

      const DWORD data_size = static_cast<DWORD>(data.size());
      ListRow row = MakeValueListRow(value_name, type, data.data(), data_size);
      browse_.values().AppendRow(std::move(row), true);
      ++current_value_count_;
      UpdateStatus();

      appended_value_name_ = value_name;
      ScheduleValueListRename(rowkind::kValue, value_name);
      StartPendingValueListRename();
    }
    return true;
  }
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleModifyCommand(int command_id) {
  switch (command_id) {
  case cmd::kEditChangeType: {
    if (!EnsureWritable() || !browse_.current_node()) {
      return true;
    }
    std::vector<ListRow> rows = SelectedListRows(browse_.values());
    if (rows.size() != 1 || rows.front().kind != rowkind::kValue) {
      return true;
    }
    ValueEntry entry;
    if (!GetValueEntry(*browse_.current_node(), rows.front().extra, &entry)) {
      ui::ShowError(hwnd_, L"Failed to read value.");
      return true;
    }
    editors::CustomValueRequest request;
    request.value_name = entry.name;
    request.type = entry.type;
    request.data = entry.data;
    editors::CustomValueResult result;
    if (!editors::EditCustomValue(hwnd_, request, &result)) {
      return true;
    }
    if (result.type == entry.type && result.data == entry.data) {
      return true;
    }
    if (!RegistryStore::SetValue(*browse_.current_node(), entry.name, result.type,
                                 result.data)) {
      ui::ShowError(hwnd_, L"Failed to update value.");
      return true;
    }
    const std::wstring old_text = value_format::Data(
        entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
    const std::wstring new_text = value_format::Data(
        result.type, result.data.data(), static_cast<DWORD>(result.data.size()));
    AppendValueHistoryEntry(L"Change type " + entry.name, old_text, new_text,
                            *browse_.current_node(), entry.name,
                            HistoryEntry::RevertKind::kSetValue, &entry);
    MarkOfflineDirty();
    changes::UndoOperation op;
    op.type = changes::UndoOperation::Type::kModifyValue;
    op.node = *browse_.current_node();
    op.old_value = entry;
    op.new_value = entry;
    op.new_value.type = result.type;
    op.new_value.data = result.data;
    PushUndo(std::move(op));
    UpdateValueListForNode(browse_.current_node());
    return true;
  }
  case cmd::kEditModify:
  case cmd::kEditModifyBinary: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!browse_.current_node()) {
      return true;
    }
    std::vector<ListRow> selected_rows = SelectedListRows(browse_.values());
    if (selected_rows.size() != 1 || selected_rows.front().kind != rowkind::kValue) {
      return true;
    }
    const ListRow* row = &selected_rows.front();
    ValueEntry entry;
    if (!GetValueEntry(*browse_.current_node(), row->extra, &entry)) {
      if (HasActiveTraces() && (row->type.empty() || EqualsInsensitive(row->type, L"TRACE"))) {
        bool needs_create = browse_.current_node()->simulated;
        DWORD type = REG_SZ;
        std::vector<BYTE> data;
        editors::CustomValueRequest request;
        request.value_name = row->extra;
        request.type = type;
        editors::CustomValueResult result;
        if (!editors::EditCustomValue(hwnd_, request, &result)) {
          return true;
        }
        type = result.type;
        data = std::move(result.data);
        if (needs_create) {
          std::wstring path = registry_path::Build(*browse_.current_node());
          if (!CreateRegistryPath(path)) {
            ui::ShowError(hwnd_, L"Failed to create the key.");
            return true;
          }
          UpdateSimulatedChain(TreeView_GetSelection(browse_.tree().hwnd()));
        }
        if (!RegistryStore::SetValue(*browse_.current_node(), row->extra, type, data)) {
          ui::ShowError(hwnd_, L"Failed to set value.");
          return true;
        }
        std::wstring display_name = row->extra.empty() ? L"(Default)" : row->extra;
        std::wstring data_text = value_format::Data(type, data.data(), static_cast<DWORD>(data.size()));
        AppendValueHistoryEntry(L"Create value " + display_name, L"", data_text, *browse_.current_node(), row->extra, HistoryEntry::RevertKind::kDeleteValue);
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kCreateValue;
        op.node = *browse_.current_node();
        op.name = row->extra;
        op.new_value.name = row->extra;
        op.new_value.type = type;
        op.new_value.data = data;
        PushUndo(std::move(op));
        RefreshTreeSelection();
        RefreshMatchingTreeNodes();
        UpdateValueListForNode(browse_.current_node());
        return true;
      }
      ui::ShowError(hwnd_, L"Failed to read value.");
      return true;
    }
    std::wstring old_text = value_format::Data(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
    DWORD base_type = value_format::NormalizeType(entry.type);
    bool supports_extended_dialog = base_type == REG_SZ || base_type == REG_EXPAND_SZ || base_type == REG_MULTI_SZ || base_type == REG_DWORD || base_type == REG_DWORD_BIG_ENDIAN || base_type == REG_QWORD || base_type == REG_LINK;
    std::vector<BYTE> new_data;
    if (command_id == cmd::kEditModifyBinary || base_type == REG_BINARY || base_type == REG_NONE || base_type == REG_RESOURCE_LIST || base_type == REG_FULL_RESOURCE_DESCRIPTOR || base_type == REG_RESOURCE_REQUIREMENTS_LIST) {
      std::wstring type_label = value_format::TypeName(entry.type);
      editors::BinaryRequest request;
      request.value_name = entry.name;
      request.value_type = std::move(type_label);
      request.data = entry.data;
      editors::BinaryResult result;
      if (!editors::EditBinary(hwnd_, request, &result)) {
        return true;
      }
      new_data = std::move(result.data);
    } else if (command_id == cmd::kEditModify && supports_extended_dialog) {
      std::wstring type_label = value_format::TypeName(entry.type);
      editors::FlaggedValueRequest request;
      request.value_name = entry.name;
      request.value_type = std::move(type_label);
      request.base_type = base_type;
      request.data = entry.data;
      editors::FlaggedValueResult result;
      if (!editors::EditFlaggedValue(hwnd_, request, &result)) {
        return true;
      }
      new_data = std::move(result.data);
    } else {
      std::wstring type_label = value_format::TypeName(entry.type);
      editors::BinaryRequest request;
      request.value_name = entry.name;
      request.value_type = std::move(type_label);
      request.data = entry.data;
      editors::BinaryResult result;
      if (!editors::EditBinary(hwnd_, request, &result)) {
        return true;
      }
      new_data = std::move(result.data);
    }
    if (new_data == entry.data) {
      return true;
    }
    if (!RegistryStore::SetValue(*browse_.current_node(), entry.name, entry.type, new_data)) {
      ui::ShowError(hwnd_, L"Failed to update value.");
    } else {
      std::wstring new_text = value_format::Data(entry.type, new_data.data(), static_cast<DWORD>(new_data.size()));
      AppendValueHistoryEntry(L"Modify value " + entry.name, old_text, new_text, *browse_.current_node(), entry.name, HistoryEntry::RevertKind::kSetValue, &entry);
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kModifyValue;
      op.node = *browse_.current_node();
      op.old_value = entry;
      op.new_value = entry;
      op.new_value.data = new_data;
      PushUndo(std::move(op));
      UpdateValueListForNode(browse_.current_node());
    }
    return true;
  }
  case cmd::kEditModifyComment: {
    if (!browse_.current_node()) {
      return true;
    }
    std::vector<ListRow> selected_rows = SelectedListRows(browse_.values());
    if (selected_rows.empty()) {
      return true;
    }
    if (std::any_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& row) {
          return row.kind != rowkind::kValue || row.simulated;
        })) {
      return true;
    }
    EditValueComments(selected_rows);
    return true;
  }
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleRenameCommand(int command_id) {
  switch (command_id) {
  case cmd::kEditRename: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!browse_.current_node()) {
      return true;
    }
    HWND focus = GetFocus();
    std::vector<ListRow> selected_rows = SelectedListRows(browse_.values());
    if (focus == browse_.values().hwnd() && selected_rows.size() > 1) {
      return true;
    }

    const ListRow* row = selected_rows.empty() ? nullptr : &selected_rows.front();
    if (focus == browse_.tree().hwnd() || (!row && browse_.current_node())) {
      if (browse_.current_node()->subkey.empty()) {
        return true;
      }
      HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
      if (selected) {
        SetFocus(browse_.tree().hwnd());
        TreeView_EditLabel(browse_.tree().hwnd(), selected);
      }
      return true;
    }
    if (row && row->kind == rowkind::kKey) {
      if (row->extra.empty()) {
        return true;
      }
      int index = -1;
      SelectedValueRow(browse_.values(), &index);
      if (focus == browse_.values().hwnd() && index >= 0) {
        SetFocus(browse_.values().hwnd());
        ListView_EditLabel(browse_.values().hwnd(), index);
        return true;
      }
      std::wstring path = registry_path::Build(*browse_.current_node());
      if (!row->extra.empty()) {
        path.append(L"\\");
        path.append(row->extra);
      }
      if (SelectTreePath(path)) {
        HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
        if (selected) {
          SetFocus(browse_.tree().hwnd());
          TreeView_EditLabel(browse_.tree().hwnd(), selected);
        }
      }
      return true;
    }
    if (row && row->kind == rowkind::kValue) {
      if (row->simulated) {
        return true;
      }
      if (row->extra.empty()) {
        return true;
      }
      int index = -1;
      SelectedValueRow(browse_.values(), &index);
      if (index >= 0) {
        SetFocus(browse_.values().hwnd());
        ListView_EditLabel(browse_.values().hwnd(), index);
      }
      return true;
    }
    return true;
  }
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleResetDefaultCommand(int command_id) {
  if (!default_reset_enabled_ || !EnsureWritable() || !browse_.current_node()) {
    return true;
  }
  std::vector<ListRow> rows = SelectedListRows(browse_.values());
  if (rows.size() != 1 || rows.front().kind != rowkind::kValue) {
    return true;
  }
  const std::wstring name = rows.front().extra;
  const std::vector<DefaultValueChoice> choices = CollectDefaultChoices(name);
  size_t index = 0;
  if (command_id != cmd::kEditResetDefault) {
    index = static_cast<size_t>(command_id - cmd::kResetDefaultBase);
  }
  if (index >= choices.size()) {
    return true;
  }
  const DefaultValueChoice& choice = choices[index];
  const std::wstring display_name = name.empty() ? L"(Default)" : name;
  ValueEntry entry;
  const bool exists = GetValueEntry(*browse_.current_node(), name, &entry);

  if (!choice.present) {
    if (!exists) {
      return true;
    }
    if (!ui::ConfirmDelete(hwnd_, L"Delete Value", display_name)) {
      return true;
    }
    if (!RegistryStore::DeleteValue(*browse_.current_node(), name)) {
      ui::ShowError(hwnd_, L"Failed to delete value.");
      return true;
    }
    AppendValueHistoryEntry(L"Reset value " + display_name, display_name, L"",
                            *browse_.current_node(), name,
                            HistoryEntry::RevertKind::kSetValue, &entry);
    changes::UndoOperation op;
    op.type = changes::UndoOperation::Type::kDeleteValue;
    op.node = *browse_.current_node();
    op.old_value = std::move(entry);
    PushUndo(std::move(op));
    MarkOfflineDirty();
    UpdateValueListForNode(browse_.current_node());
    return true;
  }

  if (exists && entry.type == choice.type && entry.data == choice.data) {
    return true;
  }
  if (!RegistryStore::SetValue(*browse_.current_node(), name, choice.type,
                               choice.data)) {
    ui::ShowError(hwnd_, L"Failed to update value.");
    return true;
  }
  const std::wstring old_text =
      exists ? value_format::Data(entry.type, entry.data.data(),
                                  static_cast<DWORD>(entry.data.size()))
             : std::wstring();
  const std::wstring new_text = value_format::Data(
      choice.type, choice.data.data(), static_cast<DWORD>(choice.data.size()));
  AppendValueHistoryEntry(
      L"Reset value " + display_name, old_text, new_text,
      *browse_.current_node(), name,
      exists ? HistoryEntry::RevertKind::kSetValue
             : HistoryEntry::RevertKind::kDeleteValue,
      exists ? &entry : nullptr);
  changes::UndoOperation op;
  op.node = *browse_.current_node();
  if (exists) {
    op.type = changes::UndoOperation::Type::kModifyValue;
    op.old_value = entry;
    op.new_value = entry;
    op.new_value.type = choice.type;
    op.new_value.data = choice.data;
  } else {
    op.type = changes::UndoOperation::Type::kCreateValue;
    op.name = name;
    op.new_value.name = name;
    op.new_value.type = choice.type;
    op.new_value.data = choice.data;
  }
  PushUndo(std::move(op));
  MarkOfflineDirty();
  UpdateValueListForNode(browse_.current_node());
  return true;
}

bool MainWindow::Impl::HandleDeleteCommand(int command_id) {
  switch (command_id) {
  case cmd::kEditDelete: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!browse_.current_node()) {
      return true;
    }
    HWND focus = GetFocus();
    bool tree_focus = (focus == browse_.tree().hwnd());
    if (tree_focus && browse_.current_node() && !browse_.current_node()->subkey.empty()) {
      std::wstring name = LeafName(*browse_.current_node());
      if (!ui::ConfirmDelete(hwnd_, L"Delete Key", name)) {
        return true;
      }
      RegistryNode target = *browse_.current_node();
      RegistryNode parent = target;
      size_t pos = parent.subkey.rfind(L'\\');
      parent.subkey = (pos == std::wstring::npos) ? L"" : parent.subkey.substr(0, pos);
      changes::KeySnapshot snapshot = changes::CaptureKey(target);
      if (!RegistryStore::DeleteKey(target)) {
        ui::ShowError(hwnd_, L"Failed to delete key.");
      } else {
        AppendHistoryEntry(L"Delete key " + name, name, L"");
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteKey;
        op.node = parent;
        op.name = name;
        op.key_snapshot = std::move(snapshot);
        PushUndo(std::move(op));
        std::wstring parent_path = registry_path::Build(parent);
        bool selected_parent = false;
        if (!parent_path.empty()) {
          selected_parent = SelectTreePath(parent_path);
        }
        RefreshTreeSelection();
        RefreshMatchingTreeNodes();
        if (!selected_parent) {
          UpdateValueListForNode(browse_.current_node());
        }
      }
      return true;
    }

    std::vector<ListRow> selected_rows = SelectedListRows(browse_.values());
    if (focus == browse_.values().hwnd() && selected_rows.size() > 1) {
      const bool all_values = std::all_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& selected) {
        return selected.kind == rowkind::kValue && !selected.simulated;
      });
      if (!all_values) {
        ui::ShowWarning(hwnd_, L"Bulk deletion only supports registry values.");
        return true;
      }

      std::vector<ValueEntry> entries;
      entries.reserve(selected_rows.size());
      for (const auto& selected : selected_rows) {
        ValueEntry entry;
        if (!GetValueEntry(*browse_.current_node(), selected.extra, &entry)) {
          ui::ShowError(hwnd_, L"Failed to read all selected values.");
          return true;
        }
        entries.push_back(std::move(entry));
      }

      std::vector<std::wstring> names;
      names.reserve(entries.size());
      for (const auto& entry : entries) {
        names.push_back(entry.name);
      }
      if (!ui::ConfirmDelete(hwnd_, L"Delete Values", names)) {
        return true;
      }

      size_t deleted = 0;
      for (auto& entry : entries) {
        if (!RegistryStore::DeleteValue(*browse_.current_node(), entry.name)) {
          continue;
        }
        std::wstring display_name = entry.name.empty() ? L"(Default)" : entry.name;
        AppendValueHistoryEntry(L"Delete value " + display_name, display_name, L"", *browse_.current_node(), entry.name, HistoryEntry::RevertKind::kSetValue, &entry);
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteValue;
        op.node = *browse_.current_node();
        op.old_value = std::move(entry);
        PushUndo(std::move(op));
        ++deleted;
      }

      if (deleted > 0) {
        MarkOfflineDirty();
        UpdateValueListForNode(browse_.current_node());
      }
      if (deleted != selected_rows.size()) {
        std::wstring message = L"Deleted " + std::to_wstring(deleted) + L" of " +
                               std::to_wstring(selected_rows.size()) + L" selected values.";
        ui::ShowError(hwnd_, message);
      }
      return true;
    }

    const ListRow* row = selected_rows.empty() ? nullptr : &selected_rows.front();
    if (row && row->kind == rowkind::kKey) {
      if (!ui::ConfirmDelete(hwnd_, L"Delete Key", row->extra)) {
        return true;
      }
      RegistryNode child = MakeChildNode(*browse_.current_node(), row->extra);
      changes::KeySnapshot snapshot = changes::CaptureKey(child);
      if (!RegistryStore::DeleteKey(child)) {
        ui::ShowError(hwnd_, L"Failed to delete key.");
      } else {
        AppendHistoryEntry(L"Delete key " + row->extra, row->extra, L"");
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteKey;
        op.node = *browse_.current_node();
        op.name = row->extra;
        op.key_snapshot = std::move(snapshot);
        PushUndo(std::move(op));
        RefreshTreeSelection();
        RefreshMatchingTreeNodes();
        UpdateValueListForNode(browse_.current_node());
      }
      return true;
    }
    if (row && row->kind == rowkind::kValue) {
      if (row->simulated) {
        return true;
      }
      std::wstring display_name = row->extra.empty() ? L"(Default)" : row->extra;
      if (!ui::ConfirmDelete(hwnd_, L"Delete Value", display_name)) {
        return true;
      }
      ValueEntry entry;
      if (!GetValueEntry(*browse_.current_node(), row->extra, &entry)) {
        ui::ShowError(hwnd_, L"Failed to read value.");
        return true;
      }
      if (!RegistryStore::DeleteValue(*browse_.current_node(), row->extra)) {
        ui::ShowError(hwnd_, L"Failed to delete value.");
      } else {
        AppendValueHistoryEntry(L"Delete value " + display_name, display_name, L"", *browse_.current_node(), row->extra, HistoryEntry::RevertKind::kSetValue, &entry);
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteValue;
        op.node = *browse_.current_node();
        op.old_value = std::move(entry);
        PushUndo(std::move(op));
        UpdateValueListForNode(browse_.current_node());
      }
      return true;
    }
    return true;
  }
  default:
    return false;
  }
}

} // namespace regkit
