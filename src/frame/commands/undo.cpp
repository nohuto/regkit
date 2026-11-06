// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "regfile/registry_transfer.h"
#include "win32/translation.h"

namespace regkit
{

using namespace window_detail;

void MainWindow::Impl::PushUndo(changes::UndoOperation operation)
{
    std::vector<changes::UndoOperation> steps;
    steps.push_back(std::move(operation));
    PushUndo(std::move(steps));
}

void MainWindow::Impl::PushUndo(std::vector<changes::UndoOperation> steps)
{
    if (is_replaying_)
    {
        return;
    }
    session_->undo.Push(std::move(steps));
    UpdateToolbarStates();
}

void MainWindow::Impl::ClearRedo()
{
    session_->undo.ClearRedo();
    UpdateToolbarStates();
}

namespace
{

using Type = changes::UndoOperation::Type;

bool TouchesKeys(const changes::UndoOperation& operation)
{
    return operation.type == Type::kCreateKey || operation.type == Type::kDeleteKey || operation.type == Type::kRenameKey || operation.type == Type::kReplaceKey ||
           std::any_of(operation.steps.begin(), operation.steps.end(), TouchesKeys);
}

// registry part of one step, caller refreshes the view once
bool ReplayStep(const changes::UndoOperation& operation, bool redo, bool* rename_left_both_names, size_t* replayed = nullptr)
{
    switch (operation.type)
    {
    case Type::kCreateKey:
        if (!redo)
        {
            return RegistryStore::DeleteKey(ChildNode(operation.node, operation.name));
        }
        return operation.key_snapshot.name.empty() ? RegistryStore::CreateKey(operation.node, operation.name) : changes::RestoreKey(operation.node, operation.key_snapshot);
    case Type::kDeleteKey:
        return redo ? RegistryStore::DeleteKey(ChildNode(operation.node, operation.name)) : changes::RestoreKey(operation.node, operation.key_snapshot);
    case Type::kRenameKey:
        return RegistryStore::RenameKey(ChildNode(operation.node, redo ? operation.name : operation.new_name), redo ? operation.new_name : operation.name);
    case Type::kCreateValue:
        return redo ? RegistryStore::SetValue(operation.node, operation.new_value.name, operation.new_value.type, operation.new_value.data)
                    : RegistryStore::DeleteValue(operation.node, operation.name);
    case Type::kDeleteValue:
        return redo ? RegistryStore::DeleteValue(operation.node, operation.old_value.name)
                    : RegistryStore::SetValue(operation.node, operation.old_value.name, operation.old_value.type, operation.old_value.data);
    case Type::kModifyValue:
        {
            const RegistryValue& value = redo ? operation.new_value : operation.old_value;
            return RegistryStore::SetValue(operation.node, value.name, value.type, value.data);
        }
    case Type::kReplaceKey:
        {
            const auto& copy = redo ? operation.hive_after : operation.hive_before;
            if (!copy || RestoreKeyCopy(operation.node.root, operation.node.subkey, ViewOf(operation.node), copy->get()) != ERROR_SUCCESS)
            {
                return false;
            }
            // the restored content had no volatile keys, the earlier content gets them back
            return redo || std::all_of(operation.volatile_keys.begin(), operation.volatile_keys.end(), [](const auto& entry) {
                       return changes::RestoreKey(entry.first, entry.second);
                   });
        }
    case Type::kRenameValue:
        return RegistryStore::RenameValue(operation.node, redo ? operation.name : operation.new_name, redo ? operation.new_name : operation.name, rename_left_both_names);
    case Type::kGroup:
        {
            // undo stops at the first failed step, earlier steps may depend on it
            const size_t count = operation.steps.size();
            size_t index = 0;
            while (index < count && ReplayStep(operation.steps[redo ? index : count - 1 - index], redo, rename_left_both_names))
            {
                ++index;
            }
            if (replayed)
            {
                *replayed = index;
            }
            return index == count;
        }
    }
    return false;
}

}

MainWindow::Impl::ReplayResult MainWindow::Impl::ApplyUndoOperation(const changes::UndoOperation& operation, bool redo, size_t* replayed, bool* left_both_names)
{
    if (!browse_.current_node())
    {
        return ReplayResult::kUnchanged;
    }
    bool rename_left_both_names = false;
    std::optional<std::wstring> restored_value;
    size_t steps = 0;
    is_replaying_ = true;
    const bool ok = ReplayStep(operation, redo, &rename_left_both_names, &steps);
    is_replaying_ = false;
    if (replayed)
    {
        *replayed = steps;
    }

    const bool changed_some = operation.type == Type::kGroup && steps > 0;
    if (TouchesKeys(operation) && (ok || changed_some))
    {
        // group keys can lie anywhere in the tree, not just under selection
        if (operation.type == Type::kGroup)
        {
            RefreshWholeTree();
        }
        else
        {
            RefreshTreeSelection();
        }
    }
    if (left_both_names)
    {
        *left_both_names = rename_left_both_names;
    }
    if (ok)
    {
        switch (operation.type)
        {
        case Type::kCreateKey:
        case Type::kDeleteKey:
            if (redo == (operation.type == Type::kCreateKey))
            {
                SelectChildKey(operation.node, operation.type == Type::kCreateKey ? operation.name : operation.key_snapshot.name);
            }
            break;
        case Type::kRenameKey:
            if (std::wstring path = registry_path::Build(operation.node); !path.empty())
            {
                SelectTreePath(path + L"\\" + (redo ? operation.new_name : operation.name));
            }
            break;
        case Type::kCreateValue:
        case Type::kDeleteValue:
            if (redo == (operation.type == Type::kCreateValue))
            {
                restored_value = redo ? operation.new_value.name : operation.old_value.name;
            }
            break;
        default:
            break;
        }
    }

    if (ok || rename_left_both_names || changed_some)
    {
        MarkOfflineDirty();
    }
    if ((ok || rename_left_both_names || changed_some) && browse_.current_node())
    {
        UpdateValueListForNode(browse_.current_node());
        if (ok && restored_value)
        {
            SelectValueAfterRefresh(*restored_value);
        }
    }
    if (rename_left_both_names)
    {
        ui::ShowError(hwnd_, util::Tr(L"The value was copied to the new name but the old name "
                                      L"couldn't be removed. Both names now exist."));
    }
    UpdateToolbarStates();
    // caller splits partly replayed groups, unchanged groups stay intact
    if (rename_left_both_names || (changed_some && !ok))
    {
        if (!rename_left_both_names)
        {
            ui::ShowError(hwnd_, redo ? util::Tr(L"The change couldn't be redone.") : util::Tr(L"The change couldn't be undone."));
        }
        return ReplayResult::kPartial;
    }
    return ok ? ReplayResult::kSuccess : ReplayResult::kUnchanged;
}

bool MainWindow::Impl::SameNode(const RegistryNode& left, const RegistryNode& right) const
{
    if (left.root != right.root || ViewOf(left) != ViewOf(right))
    {
        return false;
    }
    if (!EqualsInsensitive(left.subkey, right.subkey))
    {
        return false;
    }
    return EqualsInsensitive(left.root_name, right.root_name);
}

std::optional<std::wstring> MainWindow::Impl::MakeUniqueValueName(const RegistryNode& node, const std::wstring& base) const
{
    std::unordered_set<std::wstring> value_names;
    RegistryStore::KeyEnumResult enum_result;
    bool names_reserved = false;
    const bool listed = RegistryStore::EnumKeyStreaming(node, true, false, false, &enum_result, [&](const ValueInfo& value, const BYTE*, DWORD) {
        if (!names_reserved)
        {
            if (enum_result.info_valid)
            {
                value_names.reserve(enum_result.info.value_count);
            }
            names_reserved = true;
        }
        value_names.insert(ToLower(value.name));
        return true;
    },
                                                        {});
    if (!listed || enum_result.error != ERROR_SUCCESS)
    {
        return std::nullopt;
    }
    auto exists = [&](const std::wstring& candidate) -> bool { return value_names.contains(ToLower(candidate)); };

    std::wstring base_name = base;
    if (base_name.empty())
    {
        if (!exists(base_name))
        {
            return base_name;
        }
        base_name = L"Default";
    }
    if (!exists(base_name))
    {
        return base_name;
    }
    for (int i = 2; i < 10000; ++i)
    {
        std::wstring next = base_name + L" (" + std::to_wstring(i) + L")";
        if (!exists(next))
        {
            return next;
        }
    }
    return std::nullopt;
}

std::wstring MainWindow::Impl::MakeUniqueKeyName(const RegistryNode& node, const std::wstring& base) const
{
    auto keys = RegistryStore::EnumSubKeyNames(node, false);
    auto exists = [&](const std::wstring& candidate) -> bool {
        for (const auto& key : keys)
        {
            if (EqualsInsensitive(key, candidate))
            {
                return true;
            }
        }
        return false;
    };

    std::wstring base_name = base;
    if (base_name.empty())
    {
        base_name = L"New Key";
    }
    if (!exists(base_name))
    {
        return base_name;
    }
    for (int i = 2; i < 10000; ++i)
    {
        std::wstring next = base_name + L" (" + std::to_wstring(i) + L")";
        if (!exists(next))
        {
            return next;
        }
    }
    return base_name;
}

bool MainWindow::Impl::ResolvePathToNode(const std::wstring& path, RegistryNode* node) const
{
    return ResolvePathInRoots(browse_.roots(), path, node);
}

bool MainWindow::Impl::ResolvePathInRoots(const std::vector<RegistryRootEntry>& roots, const std::wstring& path, RegistryNode* node)
{
    if (!node || path.empty())
    {
        return false;
    }
    for (const auto& root_entry : roots)
    {
        const size_t name_size = root_entry.path_name.size();
        if (!StartsWithInsensitive(path, root_entry.path_name) ||
            (path.size() > name_size && path[name_size] != L'\\' && path[name_size] != L'/'))
        {
            continue;
        }
        const std::wstring_view tail = std::wstring_view(path).substr(name_size);
        const auto subkey = [&root_entry](std::wstring rest) {
            if (!rest.empty() && (rest.front() == L'\\' || rest.front() == L'/'))
            {
                rest.erase(rest.begin());
            }
            const std::wstring& prefix = root_entry.subkey_prefix;
            if (rest.empty())
            {
                return prefix;
            }
            return prefix.empty() || registry_path::HasComponentPrefix(rest, prefix) ? rest : prefix + L"\\" + rest;
        };
        node->root = root_entry.root;
        node->view = root_entry.view;
        node->root_name = root_entry.path_name;
        node->subkey = subkey(registry_path::RawName(tail));
        // U+2400 shows an embedded NUL but is valid in names, an existing literal name wins
        if (tail.find(L'\x2400') != std::wstring_view::npos && !RegistryStore::KeyExists(*node))
        {
            RegistryNode literal = *node;
            literal.subkey = subkey(std::wstring(tail));
            if (RegistryStore::KeyExists(literal))
            {
                node->subkey = std::move(literal.subkey);
            }
        }
        return true;
    }
    return false;
}

}
