// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"
#include "registry/resource_list.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/system/error_text.h>

namespace regkit
{
using namespace command_detail;

namespace
{

std::wstring NativeTargetPath(const std::wstring& path)
{
    const std::wstring normalized = registry_path::Normalize(path, util::GetCurrentUserSidString());
    if (normalized.empty())
    {
        return {};
    }
    const size_t split = normalized.find(L'\\');
    const std::wstring root = split == std::wstring::npos ? normalized : normalized.substr(0, split);
    const std::wstring rest = split == std::wstring::npos ? std::wstring() : normalized.substr(split);
    std::wstring native;
    if (util::EqualsInsensitive(root, L"HKEY_LOCAL_MACHINE"))
    {
        native = L"\\REGISTRY\\MACHINE";
    }
    else if (util::EqualsInsensitive(root, L"HKEY_USERS"))
    {
        native = L"\\REGISTRY\\USER";
    }
    else if (util::EqualsInsensitive(root, L"HKEY_CURRENT_USER"))
    {
        const std::wstring sid = util::GetCurrentUserSidString();
        if (sid.empty())
        {
            return {};
        }
        native = L"\\REGISTRY\\USER\\" + sid;
    }
    else if (util::EqualsInsensitive(root, L"HKEY_CLASSES_ROOT"))
    {
        native = L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes";
    }
    else if (util::EqualsInsensitive(root, L"HKEY_CURRENT_CONFIG"))
    {
        native = L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Hardware Profiles\\Current";
    }
    else
    {
        return {};
    }
    return native + rest;
}

}

bool MainWindow::Impl::HandleMutationCommand(int command_id)
{
    if (command_id == cmd::kEditResetDefault ||
        (command_id >= cmd::kResetDefaultBase && command_id <= cmd::kResetDefaultMax))
    {
        return HandleResetDefaultCommand(command_id);
    }
    switch (command_id)
    {
    case cmd::kNewSymbolicLink:
        {
            if (!EnsureWritable() || !browse_.current_node())
            {
                return true;
            }
            if (session_->mode != RegistryMode::kLocal)
            {
                ui::ShowWarning(hwnd_, util::Tr(L"Symbolic links can only be created in the local registry."));
                return true;
            }
            const bool from_list = GetFocus() == browse_.values().hwnd();
            editors::SymbolicLinkResult link;
            if (!editors::PromptSymbolicLink(
                    hwnd_,
                    MakeUniqueKeyName(*browse_.current_node(), L"New Link"),
                    [](HWND owner, std::wstring* selected) { return ShowBrowseKeyDialog(owner, selected); },
                    &link
                ))
            {
                return true;
            }
            const std::wstring native = NativeTargetPath(link.target);
            if (native.empty())
            {
                ui::ShowError(hwnd_, util::Tr(L"The target must be a full registry path under a standard root."));
                return true;
            }
            KeyDetails parent;
            // native names are available even when key queries are denied
            RegistryStore::QueryKeyDetails(*browse_.current_node(), &parent);
            if (registry_path::LinkLeavesUserHives(parent.native.native_name, native) &&
                ui::PromptChoice(hwnd_, util::Tr(L"Windows doesn't follow a link from a user hive into another user's hive or HKEY_LOCAL_MACHINE, so this link may not open. Create it anyway?"), util::Tr(L"New Symbolic Link"), util::Tr(L"Create"), util::Tr(L"Cancel"), L"") != IDYES)
            {
                return true;
            }
            DWORD error = ERROR_SUCCESS;
            if (!registry_backend::live::CreateRegistryLink(*browse_.current_node(), link.name, native, &error))
            {
                std::wstring message = util::Tr(L"Failed to create the symbolic link.");
                const std::wstring detail = win32ui::FormatWin32Error(error);
                if (!detail.empty())
                {
                    message += L"\n";
                    message += detail;
                }
                ui::ShowError(hwnd_, message);
                return true;
            }
            AppendHistoryEntry(L"Create symbolic link " + link.name, L"", native);
            MarkOfflineDirty();
            const std::wstring link_path = registry_path::Build(*browse_.current_node()) + L"\\" + link.name;
            RefreshTreeSelection();
            // like a new key, the tree selects the link unless it was made from the value list
            if (!from_list && SelectTreePath(link_path))
            {
                SetFocus(browse_.tree().hwnd());
            }
            else
            {
                UpdateValueListForNode(browse_.current_node());
            }
            return true;
        }
    case cmd::kCreateSimulatedKey:
    case cmd::kNewKey:
    case cmd::kNewVolatileKey:
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
    case cmd::kEditBits:
        return HandleModifyCommand(command_id);
    case cmd::kEditRename:
        return HandleRenameCommand(command_id);
    case cmd::kEditDelete:
        return HandleDeleteCommand(command_id);
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleCreateCommand(int command_id)
{
    if (TreeView_GetEditControl(browse_.tree().hwnd()) || ListView_GetEditControl(browse_.values().hwnd()) ||
        !pending_value_list_name_.empty())
    {
        return true;
    }
    switch (command_id)
    {
    case cmd::kCreateSimulatedKey:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            RegistryNode target;
            bool has_target = false;
            const ListRow* row = SelectedValueRow(browse_.values(), nullptr);
            if (row && row->kind == rowkind::kKey && row->simulated && browse_.current_node())
            {
                target = ChildNode(*browse_.current_node(), row->extra);
                has_target = true;
            }
            else if (browse_.current_node() && browse_.current_node()->simulated)
            {
                target = *browse_.current_node();
                has_target = true;
            }
            if (!has_target)
            {
                return true;
            }
            std::wstring path = registry_path::Build(target);
            if (!CreateRegistryPath(path))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to create the key."));
                return true;
            }
            UpdateSimulatedChain(TreeView_GetSelection(browse_.tree().hwnd()));
            RefreshTreeSelection();
            UpdateValueListForNode(browse_.current_node());
            return true;
        }
    case cmd::kNewKey:
    case cmd::kNewVolatileKey:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (!browse_.current_node() || (command_id == cmd::kNewVolatileKey && (session_->mode == RegistryMode::kOffline || ShowsRegFile())))
            {
                return true;
            }
            std::wstring name = MakeUniqueKeyName(*browse_.current_node(), L"New Key");
            if (name.empty())
            {
                return true;
            }
            bool made_volatile = false;
            if (!RegistryStore::CreateKey(*browse_.current_node(), name, {.is_volatile = command_id == cmd::kNewVolatileKey}, &made_volatile))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to create the key."));
            }
            else
            {
                AppendHistoryEntry((made_volatile ? L"Create volatile key " : L"Create key ") + name, L"", L"");
                if (made_volatile && command_id == cmd::kNewKey)
                {
                    SetStatusMessage(util::Tr(L"Created a volatile key, because its parent key is volatile."));
                }
                MarkOfflineDirty();
                changes::UndoOperation op;
                op.type = changes::UndoOperation::Type::kCreateKey;
                op.node = *browse_.current_node();
                op.name = name;
                op.key_snapshot.name = name;
                op.key_snapshot.is_volatile = made_volatile;
                PushUndo(std::move(op));
                std::wstring path = registry_path::Build(*browse_.current_node());
                if (!path.empty())
                {
                    path.append(L"\\");
                    path.append(name);
                }
                HWND focus = GetFocus();
                bool edit_in_list = (focus == browse_.values().hwnd()) && settings_.show_keys_in_list && browse_.values().hwnd();
                if (edit_in_list)
                {
                    ScheduleValueListRename(rowkind::kKey, name);
                    UpdateValueListForNode(browse_.current_node());
                }
                else
                {
                    std::wstring parent_path = registry_path::Build(*browse_.current_node());
                    HTREEITEM parent_item = TreeView_GetSelection(browse_.tree().hwnd());
                    if (!parent_item && !parent_path.empty())
                    {
                        SelectTreePath(parent_path);
                        parent_item = TreeView_GetSelection(browse_.tree().hwnd());
                    }

                    if (parent_item)
                    {
                        TreeView_SelectItem(browse_.tree().hwnd(), parent_item);
                    }
                    HTREEITEM target = browse_.tree().InsertChild(parent_item, name);
                    if (!target)
                    {
                        RefreshTreeSelection();
                    }
                    else
                    {
                        RefreshMatchingTreeNodes();
                    }

                    if (!target && parent_item)
                    {
                        target = FindChildByText(browse_.tree().hwnd(), parent_item, name);
                    }
                    if (target)
                    {
                        TreeView_SelectItem(browse_.tree().hwnd(), target);
                        TreeView_EnsureVisible(browse_.tree().hwnd(), target);
                    }
                    if (!target && !path.empty())
                    {
                        if (SelectTreePath(path))
                        {
                            target = TreeView_GetSelection(browse_.tree().hwnd());
                        }
                    }
                    if (target)
                    {
                        SetFocus(browse_.tree().hwnd());
                        TreeView_EditLabel(browse_.tree().hwnd(), target);
                    }
                }
            }
            return true;
        }
    case cmd::kNewString:
    case cmd::kNewExpandString:
    case cmd::kNewBinary:
    case cmd::kNewDword:
    case cmd::kNewQword:
    case cmd::kNewMultiString:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (!browse_.current_node())
            {
                return true;
            }
            DWORD type = REG_SZ;
            std::wstring base_name = L"New Value";
            switch (command_id)
            {
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
            const std::optional<std::wstring> free_name = MakeUniqueValueName(*browse_.current_node(), base_name);
            if (!free_name)
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to set value."));
                return true;
            }
            const std::wstring& value_name = *free_name;
            std::vector<BYTE> data;
            if (type == REG_SZ || type == REG_EXPAND_SZ)
            {
                data.resize(sizeof(wchar_t), 0);
            }
            else if (type == REG_MULTI_SZ)
            {
                data.resize(sizeof(wchar_t) * 2, 0);
            }
            else if (type == REG_DWORD)
            {
                data.resize(sizeof(DWORD), 0);
            }
            else if (type == REG_QWORD)
            {
                data.resize(sizeof(unsigned long long), 0);
            }
            if (!RegistryStore::SetValue(*browse_.current_node(), value_name, type, data))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to set value."));
            }
            else
            {
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

                // a running load misses the new value and cancelling it hides the rest, a new load has both
                ScheduleValueListRename(rowkind::kValue, value_name);
                if (value_list_loading_)
                {
                    UpdateValueListForNode(browse_.current_node());
                    appended_value_name_ = value_name;
                }
                else
                {
                    value_list_generation_.fetch_add(1);
                    const DWORD data_size = static_cast<DWORD>(data.size());
                    ListRow row = MakeValueListRow(value_name, type, data.data(), data_size);
                    browse_.values().AppendRow(std::move(row), true);
                    ++current_value_count_;
                    UpdateStatus();
                    appended_value_name_ = value_name;
                    StartPendingValueListRename();
                }
            }
            return true;
        }
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleModifyCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kEditChangeType:
        {
            if (!browse_.current_node())
            {
                return true;
            }
            const std::vector<const ListRow*> rows = SelectedListRows(browse_.values());
            if (rows.size() != 1 || rows.front()->kind != rowkind::kValue)
            {
                return true;
            }
            RegistryValue entry;
            if (!GetValueEntry(*browse_.current_node(), rows.front()->extra, &entry))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to read value."));
                return true;
            }
            editors::CustomValueRequest request;
            request.value_name = entry.name;
            request.key_path = registry_path::Build(*browse_.current_node());
            request.type = entry.type;
            request.data = entry.data;
            request.read_only = settings_.read_only;
            editors::CustomValueResult result;
            if (!editors::EditCustomValue(hwnd_, request, &result))
            {
                return true;
            }
            if (settings_.read_only)
            {
                return true;
            }
            if (result.type == entry.type && result.data == entry.data)
            {
                return true;
            }
            if (!RegistryStore::SetValue(*browse_.current_node(), entry.name, result.type, result.data))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to update value."));
                return true;
            }
            const std::wstring old_text =
                value_format::Data(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
            const std::wstring new_text =
                value_format::Data(result.type, result.data.data(), static_cast<DWORD>(result.data.size()));
            AppendValueHistoryEntry(L"Change type " + entry.name, old_text, new_text, *browse_.current_node(), entry.name, HistoryEntry::RevertKind::kSetValue, &entry);
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
    case cmd::kEditModifyBinary:
    case cmd::kEditBits:
        {
            if (!browse_.current_node())
            {
                return true;
            }
            const std::vector<const ListRow*> selected_rows = SelectedListRows(browse_.values());
            if (selected_rows.size() != 1 || selected_rows.front()->kind != rowkind::kValue)
            {
                return true;
            }
            // the editor runs a modal loop
            const ListRow selected = *selected_rows.front();
            const ListRow* row = &selected;
            RegistryValue entry;
            if (!GetValueEntry(*browse_.current_node(), row->extra, &entry))
            {
                if (settings_.read_only)
                {
                    return true;
                }
                if (HasActiveTraces() && (row->type.empty() || row->type == L"TRACE"))
                {
                    bool needs_create = browse_.current_node()->simulated;
                    editors::CustomValueRequest request;
                    request.value_name = row->extra;
                    request.key_path = registry_path::Build(*browse_.current_node());
                    request.type = command_id == cmd::kEditModifyBinary ? REG_BINARY : REG_SZ;
                    editors::CustomValueResult result;
                    if (!editors::EditCustomValue(hwnd_, request, &result))
                    {
                        return true;
                    }
                    const DWORD type = result.type;
                    const std::vector<BYTE> data = std::move(result.data);
                    if (needs_create)
                    {
                        std::wstring path = registry_path::Build(*browse_.current_node());
                        if (!CreateRegistryPath(path))
                        {
                            ui::ShowError(hwnd_, util::Tr(L"Failed to create the key."));
                            return true;
                        }
                        UpdateSimulatedChain(TreeView_GetSelection(browse_.tree().hwnd()));
                    }
                    if (!RegistryStore::SetValue(*browse_.current_node(), row->extra, type, data))
                    {
                        ui::ShowError(hwnd_, util::Tr(L"Failed to set value."));
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
                    UpdateValueListForNode(browse_.current_node());
                    return true;
                }
                ui::ShowError(hwnd_, util::Tr(L"Failed to read value."));
                return true;
            }
            std::wstring old_text =
                value_format::Data(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
            DWORD base_type = value_format::NormalizeType(entry.type);
            bool supports_extended_dialog =
                base_type == REG_SZ || base_type == REG_EXPAND_SZ || base_type == REG_MULTI_SZ || base_type == REG_DWORD ||
                base_type == REG_DWORD_BIG_ENDIAN || base_type == REG_QWORD || base_type == REG_LINK;
            std::vector<BYTE> new_data;
            // resource lists contain kernel hardware assignments, Modify shows decoded data with hex fallback
            const bool resource = command_id == cmd::kEditModify && resource_list::IsResourceType(base_type);
            const bool edit_binary = command_id == cmd::kEditModifyBinary || base_type == REG_BINARY || base_type == REG_NONE || (resource && ShowResourceList(entry));
            if (resource && !edit_binary)
            {
                return true;
            }
            const std::wstring key_path = registry_path::Build(*browse_.current_node());
            if (command_id == cmd::kEditBits)
            {
                editors::BitsRequest request;
                request.value_name = entry.name;
                request.key_path = key_path;
                request.base_type = base_type;
                request.data = entry.data;
                request.read_only = settings_.read_only;
                if (!editors::EditValueBits(hwnd_, request, &new_data))
                {
                    return true;
                }
            }
            else if (edit_binary)
            {
                editors::BinaryRequest request;
                request.value_name = entry.name;
                request.key_path = key_path;
                request.data = entry.data;
                request.read_only = settings_.read_only;
                editors::BinaryResult result;
                if (!editors::EditBinary(hwnd_, request, &result))
                {
                    return true;
                }
                new_data = std::move(result.data);
            }
            else if (command_id == cmd::kEditModify && supports_extended_dialog)
            {
                editors::FlaggedValueRequest request;
                request.value_name = entry.name;
                request.key_path = key_path;
                request.base_type = base_type;
                request.data = entry.data;
                request.read_only = settings_.read_only;
                editors::FlaggedValueResult result;
                if (!editors::EditFlaggedValue(hwnd_, request, &result))
                {
                    return true;
                }
                new_data = std::move(result.data);
            }
            else
            {
                editors::BinaryRequest request;
                request.value_name = entry.name;
                request.key_path = key_path;
                request.data = entry.data;
                request.read_only = settings_.read_only;
                editors::BinaryResult result;
                if (!editors::EditBinary(hwnd_, request, &result))
                {
                    return true;
                }
                new_data = std::move(result.data);
            }
            if (settings_.read_only)
            {
                return true;
            }
            if (new_data == entry.data)
            {
                return true;
            }
            if (!RegistryStore::SetValue(*browse_.current_node(), entry.name, entry.type, new_data))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to update value."));
            }
            else
            {
                std::wstring new_text =
                    value_format::Data(entry.type, new_data.data(), static_cast<DWORD>(new_data.size()));
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
    case cmd::kEditModifyComment:
        {
            const std::vector<changes::CommentTarget> targets = CommentTargets();
            if (!targets.empty())
            {
                EditComments(targets);
            }
            return true;
        }
    default:
        return false;
    }
}

std::vector<changes::CommentTarget> MainWindow::Impl::CommentTargets()
{
    std::vector<changes::CommentTarget> targets;
    if (!browse_.current_node())
    {
        return targets;
    }
    const std::vector<const ListRow*> selected_rows = SelectedListRows(browse_.values());
    const std::wstring path = CommentKeyPath(*browse_.current_node());
    if (GetFocus() == browse_.tree().hwnd() || selected_rows.empty())
    {
        targets.push_back({path, {}, 0, 0, true});
    }
    else if (selected_rows.size() == 1 && selected_rows.front()->kind == rowkind::kKey &&
             !selected_rows.front()->extra.empty())
    {
        targets.push_back({registry_path::JoinSubkey(path, selected_rows.front()->extra), {}, 0, 0, true});
    }
    else
    {
        // simulated SymbolicLinkValue rows support comments like real values
        for (const ListRow* row : selected_rows)
        {
            if (row->kind != rowkind::kValue)
            {
                return {};
            }
            targets.push_back({path, row->extra, row->value_type, row->value_data_size});
        }
    }
    return targets;
}

bool MainWindow::Impl::HandleRenameCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kEditRename:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (!browse_.current_node())
            {
                return true;
            }
            HWND focus = GetFocus();
            const std::vector<const ListRow*> selected_rows = SelectedListRows(browse_.values());
            if (focus == browse_.values().hwnd() && selected_rows.size() > 1)
            {
                return true;
            }

            const ListRow* row = selected_rows.empty() ? nullptr : selected_rows.front();
            if (focus == browse_.tree().hwnd() || (!row && browse_.current_node()))
            {
                if (browse_.current_node()->subkey.empty())
                {
                    return true;
                }
                HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
                if (selected)
                {
                    SetFocus(browse_.tree().hwnd());
                    TreeView_EditLabel(browse_.tree().hwnd(), selected);
                }
                return true;
            }
            if (row && row->kind == rowkind::kKey)
            {
                if (row->extra.empty())
                {
                    return true;
                }
                int index = -1;
                SelectedValueRow(browse_.values(), &index);
                if (focus == browse_.values().hwnd() && index >= 0)
                {
                    SetFocus(browse_.values().hwnd());
                    ListView_EditLabel(browse_.values().hwnd(), index);
                    return true;
                }
                std::wstring path = registry_path::Build(*browse_.current_node());
                if (!row->extra.empty())
                {
                    path.append(L"\\");
                    path.append(row->extra);
                }
                if (SelectTreePath(path))
                {
                    HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
                    if (selected)
                    {
                        SetFocus(browse_.tree().hwnd());
                        TreeView_EditLabel(browse_.tree().hwnd(), selected);
                    }
                }
                return true;
            }
            if (row && row->kind == rowkind::kValue)
            {
                if (row->simulated)
                {
                    return true;
                }
                if (row->extra.empty())
                {
                    return true;
                }
                int index = -1;
                SelectedValueRow(browse_.values(), &index);
                if (index >= 0)
                {
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

bool MainWindow::Impl::HandleResetDefaultCommand(int command_id)
{
    if (!settings_.default_reset_enabled || !EnsureWritable() || !browse_.current_node())
    {
        return true;
    }
    const std::vector<const ListRow*> rows = SelectedListRows(browse_.values());
    if (rows.size() != 1 || rows.front()->kind != rowkind::kValue)
    {
        return true;
    }
    const std::wstring name = rows.front()->extra;
    const std::vector<DefaultValueChoice> choices = CollectDefaultChoices(name);
    size_t index = 0;
    if (command_id != cmd::kEditResetDefault)
    {
        index = static_cast<size_t>(command_id - cmd::kResetDefaultBase);
    }
    if (index >= choices.size())
    {
        return true;
    }
    const DefaultValueChoice& choice = choices[index];
    const std::wstring display_name = name.empty() ? L"(Default)" : name;
    RegistryValue entry;
    const bool exists = GetValueEntry(*browse_.current_node(), name, &entry);

    if (!choice.present)
    {
        if (!exists)
        {
            return true;
        }
        if (!ui::ConfirmDelete(hwnd_, util::Tr(L"Delete Value"), registry_path::DisplayName(name)))
        {
            return true;
        }
        if (!RegistryStore::DeleteValue(*browse_.current_node(), name))
        {
            ui::ShowError(hwnd_, util::Tr(L"Failed to delete value."));
            return true;
        }
        AppendValueHistoryEntry(L"Delete value " + display_name, display_name, L"", *browse_.current_node(), name, HistoryEntry::RevertKind::kSetValue, &entry);
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteValue;
        op.node = *browse_.current_node();
        op.old_value = std::move(entry);
        PushUndo(std::move(op));
        MarkOfflineDirty();
        UpdateValueListForNode(browse_.current_node());
        return true;
    }

    if (exists && entry.type == choice.type && entry.data == choice.data)
    {
        return true;
    }
    if (!RegistryStore::SetValue(*browse_.current_node(), name, choice.type, choice.data))
    {
        ui::ShowError(hwnd_, util::Tr(L"Failed to update value."));
        return true;
    }
    const std::wstring old_text =
        exists ? value_format::Data(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()))
               : std::wstring();
    const std::wstring new_text =
        value_format::Data(choice.type, choice.data.data(), static_cast<DWORD>(choice.data.size()));
    AppendValueHistoryEntry(L"Reset value " + display_name, old_text, new_text, *browse_.current_node(), name, exists ? HistoryEntry::RevertKind::kSetValue : HistoryEntry::RevertKind::kDeleteValue, exists ? &entry : nullptr);
    changes::UndoOperation op;
    op.node = *browse_.current_node();
    if (exists)
    {
        op.type = changes::UndoOperation::Type::kModifyValue;
        op.old_value = entry;
        op.new_value = entry;
        op.new_value.type = choice.type;
        op.new_value.data = choice.data;
    }
    else
    {
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

bool MainWindow::Impl::HandleDeleteCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kEditDelete:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (!browse_.current_node())
            {
                return true;
            }
            HWND focus = GetFocus();
            bool tree_focus = (focus == browse_.tree().hwnd());
            if (tree_focus && browse_.current_node() && !browse_.current_node()->subkey.empty())
            {
                std::wstring name = LeafName(*browse_.current_node());
                if (!ui::ConfirmDelete(hwnd_, util::Tr(L"Delete Key"), registry_path::DisplayName(name)))
                {
                    return true;
                }
                RegistryNode target = *browse_.current_node();
                RegistryNode parent = target;
                parent.subkey = registry_path::Parent(target.subkey);
                HTREEITEM deleted_item = TreeView_GetSelection(browse_.tree().hwnd());
                HTREEITEM parent_item = deleted_item ? TreeView_GetParent(browse_.tree().hwnd(), deleted_item) : nullptr;
                const std::wstring neighbour_path = TreeNeighbourPath(deleted_item);
                changes::KeySnapshot snapshot = changes::CaptureKey(target);
                const bool restorable = snapshot.complete;
                if (!restorable && !ui::ConfirmDelete(hwnd_, util::Tr(L"Delete Key"), registry_path::DisplayName(name), util::Tr(L"Part of this key couldn't be read, so this delete can't be "
                                                                                                                                 L"undone. Delete anyway?")))
                {
                    return true;
                }
                if (!RegistryStore::DeleteKey(target))
                {
                    ui::ShowError(hwnd_, util::Tr(L"Failed to delete key."));
                }
                else
                {
                    AppendHistoryEntry(L"Delete key " + name, name, L"");
                    MarkOfflineDirty();
                    if (restorable)
                    {
                        changes::UndoOperation op;
                        op.type = changes::UndoOperation::Type::kDeleteKey;
                        op.node = parent;
                        op.name = name;
                        op.key_snapshot = std::move(snapshot);
                        PushUndo(std::move(op));
                    }
                    if (deleted_item)
                    {
                        TreeView_DeleteItem(browse_.tree().hwnd(), deleted_item);
                    }
                    std::wstring next_path = neighbour_path.empty() ? registry_path::Build(parent) : neighbour_path;
                    const bool moved = !next_path.empty() && SelectTreePath(next_path);
                    RefreshMatchingTreeNodes(parent_item);
                    if (!moved)
                    {
                        UpdateValueListForNode(browse_.current_node());
                    }
                }
                return true;
            }

            std::vector<const ListRow*> selected_rows = SelectedListRows(browse_.values());
            if (focus == browse_.values().hwnd() && selected_rows.size() > 1)
            {
                const bool all_values =
                    std::all_of(selected_rows.begin(), selected_rows.end(), [](const ListRow* selected) {
                        return selected->kind == rowkind::kValue && !selected->simulated;
                    });
                if (!all_values)
                {
                    ui::ShowWarning(hwnd_, util::Tr(L"Bulk deletion only supports registry values."));
                    return true;
                }
                // unset defaults and trace only rows have nothing to delete
                std::erase_if(selected_rows, [](const ListRow* selected) { return selected->missing; });
                if (selected_rows.empty())
                {
                    return true;
                }

                std::vector<RegistryValue> entries;
                entries.reserve(selected_rows.size());
                for (const ListRow* selected : selected_rows)
                {
                    RegistryValue entry;
                    if (!GetValueEntry(*browse_.current_node(), selected->extra, &entry))
                    {
                        ui::ShowError(hwnd_, util::Tr(L"Failed to read all selected values."));
                        return true;
                    }
                    entries.push_back(std::move(entry));
                }

                std::vector<std::wstring> names;
                names.reserve(entries.size());
                for (const auto& entry : entries)
                {
                    names.push_back(entry.name);
                }
                if (!ui::ConfirmDelete(hwnd_, util::Tr(L"Delete Values"), names))
                {
                    return true;
                }

                std::vector<changes::UndoOperation> steps;
                std::vector<HistoryEntry> history;
                steps.reserve(entries.size());
                history.reserve(entries.size());
                for (auto& entry : entries)
                {
                    if (!RegistryStore::DeleteValue(*browse_.current_node(), entry.name))
                    {
                        continue;
                    }
                    std::wstring display_name = entry.name.empty() ? L"(Default)" : entry.name;
                    history.push_back(ValueHistoryEntry(L"Delete value " + display_name, display_name, L"", *browse_.current_node(), entry.name, HistoryEntry::RevertKind::kSetValue, &entry));
                    changes::UndoOperation op;
                    op.type = changes::UndoOperation::Type::kDeleteValue;
                    op.node = *browse_.current_node();
                    op.old_value = std::move(entry);
                    steps.push_back(std::move(op));
                }
                const size_t deleted = steps.size();
                AppendHistoryEntries(std::move(history));
                PushUndo(std::move(steps));

                if (deleted > 0)
                {
                    MarkOfflineDirty();
                    UpdateValueListForNode(browse_.current_node());
                }
                if (deleted != selected_rows.size())
                {
                    ui::ShowError(hwnd_, util::Tr(L"Some of the selected values couldn't be deleted."));
                }
                return true;
            }

            // the confirmation runs a modal loop
            const ListRow selected = selected_rows.empty() ? ListRow{} : *selected_rows.front();
            const ListRow* row = selected_rows.empty() ? nullptr : &selected;
            if (row && row->kind == rowkind::kKey)
            {
                if (!ui::ConfirmDelete(hwnd_, util::Tr(L"Delete Key"), row->name))
                {
                    return true;
                }
                RegistryNode child = ChildNode(*browse_.current_node(), row->extra);
                changes::KeySnapshot snapshot = changes::CaptureKey(child);
                const bool restorable = snapshot.complete;
                if (!restorable && !ui::ConfirmDelete(hwnd_, util::Tr(L"Delete Key"), row->name, util::Tr(L"Part of this key couldn't be read, so this delete can't be "
                                                                                                          L"undone. Delete anyway?")))
                {
                    return true;
                }
                if (!RegistryStore::DeleteKey(child))
                {
                    ui::ShowError(hwnd_, util::Tr(L"Failed to delete key."));
                }
                else
                {
                    AppendHistoryEntry(L"Delete key " + row->extra, row->extra, L"");
                    MarkOfflineDirty();
                    if (restorable)
                    {
                        changes::UndoOperation op;
                        op.type = changes::UndoOperation::Type::kDeleteKey;
                        op.node = *browse_.current_node();
                        op.name = row->extra;
                        op.key_snapshot = std::move(snapshot);
                        PushUndo(std::move(op));
                    }
                    RefreshTreeSelection();
                    UpdateValueListForNode(browse_.current_node());
                }
                return true;
            }
            if (row && row->kind == rowkind::kValue)
            {
                if (row->simulated || row->missing)
                {
                    return true;
                }
                std::wstring display_name = row->extra.empty() ? L"(Default)" : row->extra;
                if (!ui::ConfirmDelete(hwnd_, util::Tr(L"Delete Value"), registry_path::DisplayName(row->extra)))
                {
                    return true;
                }
                RegistryValue entry;
                if (!GetValueEntry(*browse_.current_node(), row->extra, &entry))
                {
                    ui::ShowError(hwnd_, util::Tr(L"Failed to read value."));
                    return true;
                }
                if (!RegistryStore::DeleteValue(*browse_.current_node(), row->extra))
                {
                    ui::ShowError(hwnd_, util::Tr(L"Failed to delete value."));
                }
                else
                {
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

}
