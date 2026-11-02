// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_store.h"

#include "registry/registry_backends.h"
#include "registry/registry_path.h"
#include "win32/text_transform.h"

#include <win32ui/handle_owner.h>

#include <algorithm>
#include <atomic>

namespace regkit
{
namespace
{

using registry_backend::virtual_store::FindRoot;

template <typename VirtualCall, typename OfflineCall, typename LiveCall>
decltype(auto) Dispatch(const RegistryNode& node, VirtualCall&& virtual_call, OfflineCall&& offline_call, LiveCall&& live_call)
{
    if (std::shared_ptr<VirtualRegistryData> data = FindRoot(node.root))
    {
        return virtual_call(*data);
    }
    if (registry_backend::offline::Owns(node.root))
    {
        return offline_call();
    }
    return live_call();
}

std::atomic<uint64_t> g_key_revision{0};

std::wstring CaseRenameStep(std::wstring_view old_name, std::wstring_view new_name)
{
    return old_name != new_name && util::EqualsInsensitive(old_name, new_name) ? std::wstring(new_name) + L"~" + std::to_wstring(GetTickCount64()) : std::wstring();
}

}

std::vector<RegistryRootEntry> RegistryStore::DefaultRoots(bool include_extra)
{
    std::vector<RegistryRootEntry> roots = {
        {HKEY_CLASSES_ROOT, L"HKEY_CLASSES_ROOT", L"HKEY_CLASSES_ROOT", L""},
        {HKEY_CURRENT_USER, L"HKEY_CURRENT_USER", L"HKEY_CURRENT_USER", L""},
        {HKEY_LOCAL_MACHINE, L"HKEY_LOCAL_MACHINE", L"HKEY_LOCAL_MACHINE", L""},
        {HKEY_USERS, L"HKEY_USERS", L"HKEY_USERS", L""},
        {HKEY_CURRENT_CONFIG, L"HKEY_CURRENT_CONFIG", L"HKEY_CURRENT_CONFIG", L""},
    };
    if (include_extra)
    {
        roots.push_back({HKEY_PERFORMANCE_DATA, L"HKEY_PERFORMANCE_DATA", L"HKEY_PERFORMANCE_DATA", L""});
        roots.push_back({HKEY_PERFORMANCE_TEXT, L"HKEY_PERFORMANCE_TEXT", L"HKEY_PERFORMANCE_TEXT", L""});
        roots.push_back({HKEY_PERFORMANCE_NLSTEXT, L"HKEY_PERFORMANCE_NLSTEXT", L"HKEY_PERFORMANCE_NLSTEXT", L""});
    }
    return roots;
}

bool RegistryStore::OpenOfflineHive(const std::wstring& path, HKEY* root, std::wstring* error)
{
    return registry_backend::offline::OpenHive(path, root, error);
}

void RegistryStore::AddOfflineRoot(HKEY root)
{
    registry_backend::offline::AddRoot(root);
}

void RegistryStore::RemoveOfflineRoot(HKEY root)
{
    registry_backend::offline::RemoveRoot(root);
}

bool RegistryStore::SaveOfflineHive(HKEY root, const std::wstring& path, std::wstring* error)
{
    return registry_backend::offline::SaveHive(root, path, error);
}

bool RegistryStore::CloseOfflineHive(HKEY root, std::wstring* error)
{
    return registry_backend::offline::CloseHive(root, error);
}

HKEY RegistryStore::RegisterVirtualRoot(const std::wstring& root_name, const std::shared_ptr<VirtualRegistryData>& data)
{
    return registry_backend::virtual_store::RegisterRoot(root_name, data);
}

void RegistryStore::UnregisterVirtualRoot(HKEY root)
{
    registry_backend::virtual_store::UnregisterRoot(root);
}

HKEY RegistryStore::DuplicateVirtualRoot(HKEY root)
{
    std::wstring name;
    const auto data = FindRoot(root, &name);
    return data ? registry_backend::virtual_store::RegisterRoot(name, data) : nullptr;
}

bool RegistryStore::IsVirtualRoot(HKEY root)
{
    return static_cast<bool>(FindRoot(root));
}

bool RegistryStore::GetVirtualRootName(HKEY root, std::wstring* root_name)
{
    return static_cast<bool>(FindRoot(root, root_name));
}

uint64_t RegistryStore::KeyRevision() noexcept
{
    return g_key_revision.load();
}

void RegistryStore::NoteKeyChange() noexcept
{
    ++g_key_revision;
}

bool RegistryStore::HasSubKeys(const RegistryNode& node)
{
    return Dispatch(
        node,
        [&](const VirtualRegistryData& data) { return registry_backend::virtual_store::HasSubKeys(data, node); },
        [&] { return registry_backend::offline::HasSubKeys(node); },
        [&] { return registry_backend::live::HasSubKeys(node); }
    );
}

bool RegistryStore::QueryKeyInfo(const RegistryNode& node, KeyInfo* info)
{
    if (!info)
    {
        return false;
    }
    return Dispatch(
        node,
        [&](const VirtualRegistryData& data) {
            return registry_backend::virtual_store::QueryKeyInfo(data, node, info);
        },
        [&] { return registry_backend::offline::QueryKeyInfo(node, info); },
        [&] { return registry_backend::live::QueryKeyInfo(node, info); }
    );
}

bool RegistryStore::KeyExists(const RegistryNode& node)
{
    KeyInfo info = {};
    if (QueryKeyInfo(node, &info))
    {
        return true;
    }
    if (node.subkey.empty())
    {
        return false;
    }
    RegistryNode parent = node;
    parent.subkey = registry_path::Parent(node.subkey);
    const std::wstring leaf = registry_path::Leaf(node.subkey);
    const std::vector<std::wstring> names = EnumSubKeyNames(parent, false);
    return std::any_of(names.begin(), names.end(), [&](const std::wstring& name) { return util::EqualsInsensitive(name, leaf); });
}

bool RegistryStore::QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target, bool* denied)
{
    if (!target)
    {
        return false;
    }
    target->clear();
    if (FindRoot(node.root))
    {
        return false;
    }
    return registry_backend::offline::Owns(node.root)
               ? registry_backend::offline::QuerySymbolicLinkTarget(node, target)
               : registry_backend::live::QuerySymbolicLinkTarget(node, target, denied);
}

bool RegistryStore::IsBrokenLink(const RegistryNode& node, std::wstring* target)
{
    std::wstring ignored;
    return Dispatch(
        node,
        [&](VirtualRegistryData&) { return false; },
        [&] { return false; },
        [&] { return registry_backend::live::IsBrokenLink(node, target ? target : &ignored); }
    );
}

KeyInspection RegistryStore::InspectKey(const RegistryNode& node, bool want_info, bool want_source)
{
    KeyInspection result;
    if (FindRoot(node.root) || registry_backend::offline::Owns(node.root))
    {
        std::wstring target;
        result.link = QuerySymbolicLinkTarget(node, &target, &result.denied);
    }
    else
    {
        result = registry_backend::live::InspectKey(node, want_info, want_source);
    }
    if (want_info && !result.info_valid)
    {
        result.info_valid = QueryKeyInfo(node, &result.info);
    }
    return result;
}

std::vector<std::wstring> RegistryStore::EnumSubKeyNames(const RegistryNode& node, bool sorted)
{
    return Dispatch(
        node,
        [&](const VirtualRegistryData& data) {
            return registry_backend::virtual_store::EnumSubKeyNames(data, node, sorted);
        },
        [&] { return registry_backend::offline::EnumSubKeyNames(node, sorted); },
        [&] { return registry_backend::live::EnumSubKeyNames(node, sorted); }
    );
}

bool RegistryStore::EnumKeyStreaming(const RegistryNode& node, bool include_values, bool include_data, bool include_subkeys, KeyEnumResult* out_info, const ValueStreamCallback& value_callback, const SubkeyStreamCallback& subkey_callback, DWORD max_data_size, EnumerationScratch* scratch, bool ordered, bool open_link)
{
    if (out_info)
    {
        out_info->info = {};
        out_info->info_valid = false;
    }
    return Dispatch(
        node,
        [&](const VirtualRegistryData& data) {
            return registry_backend::virtual_store::EnumKeyStreaming(data, node, include_values, include_data, include_subkeys, out_info, value_callback, subkey_callback, max_data_size, scratch, ordered);
        },
        [&] {
            return registry_backend::offline::EnumKeyStreaming(node, include_values, include_data, include_subkeys, out_info, value_callback, subkey_callback, max_data_size, scratch, ordered);
        },
        [&] {
            return registry_backend::live::EnumKeyStreaming(node, include_values, include_data, include_subkeys, out_info, value_callback, subkey_callback, max_data_size, scratch, ordered, open_link);
        }
    );
}

bool RegistryStore::IsOfflineRoot(HKEY root)
{
    return registry_backend::offline::Owns(root);
}

bool RegistryStore::QueryValue(const RegistryNode& node, const std::wstring& value_name, RegistryValue* out)
{
    if (!out)
    {
        return false;
    }
    return Dispatch(
        node,
        [&](const VirtualRegistryData& data) {
            return registry_backend::virtual_store::QueryValue(data, node, value_name, out);
        },
        [&] { return registry_backend::offline::QueryValue(node, value_name, out); },
        [&] { return registry_backend::live::QueryValue(node, value_name, out); }
    );
}

LONG RegistryStore::SetKeyControlFlags(const RegistryNode& node, ULONG flags)
{
    NoteKeyChange();
    win32ui::UniqueHKey key;
    const LONG status = util::OpenRegistryPath(node.root, node.subkey, KEY_SET_VALUE | ViewOf(node), false, &key);
    return status == ERROR_SUCCESS ? util::SetKeyControlFlags(key.get(), flags) : status;
}

bool RegistryStore::QueryKeyDetails(const RegistryNode& node, KeyDetails* details, bool open_link)
{
    *details = {};
    return Dispatch(
        node,
        [&](const VirtualRegistryData& data) {
            return registry_backend::virtual_store::QueryKeyInfo(data, node, &details->info);
        },
        [&] { return registry_backend::offline::QueryKeyDetails(node, details); },
        [&] { return registry_backend::live::QueryKeyDetails(node, details, open_link); }
    );
}

std::wstring RegistryStore::OfflineControlSet(const RegistryNode& node)
{
    return Dispatch(
        node,
        [&](VirtualRegistryData&) { return std::wstring(); },
        [&] { return registry_backend::offline::SelectedControlSet(node); },
        [&] { return std::wstring(); }
    );
}

bool RegistryStore::CreateKey(const RegistryNode& node, const std::wstring& name, const KeyCreateOptions& options, bool* created_volatile)
{
    NoteKeyChange();
    if (created_volatile)
    {
        *created_volatile = false;
    }
    if (name.empty())
    {
        return false;
    }
    return Dispatch(
        node,
        [&](VirtualRegistryData& data) { return registry_backend::virtual_store::CreateKey(data, node, name); },
        [&] { return registry_backend::offline::CreateKey(node, name, options.class_name); },
        [&] { return registry_backend::live::CreateKey(node, name, options, created_volatile); }
    );
}

bool RegistryStore::CreateKeyLink(const RegistryNode& node, const std::wstring& name, const std::wstring& nt_target, const KeyCreateOptions& options)
{
    NoteKeyChange();
    return Dispatch(
        node,
        [&](VirtualRegistryData&) { return false; },
        [&] { return false; },
        [&] { return registry_backend::live::CreateRegistryLink(node, name, nt_target, nullptr, options); }
    );
}

bool RegistryStore::ReadKeyLink(const RegistryNode& node, std::wstring* target)
{
    if (node.subkey.empty())
    {
        return false;
    }
    return Dispatch(
        node,
        [&](VirtualRegistryData&) { return false; },
        [&] { return false; },
        [&] { return registry_backend::live::ReadKeyLink(node, target); }
    );
}

bool RegistryStore::ReadKeySecurity(const RegistryNode& node, SECURITY_INFORMATION* parts, std::vector<BYTE>* descriptor, LONG* error)
{
    return Dispatch(
        node,
        [&](VirtualRegistryData&) { return false; },
        [&] { return registry_backend::offline::ReadKeySecurity(node, parts, descriptor, error); },
        [&] { return registry_backend::live::ReadKeySecurity(node, parts, descriptor, error); }
    );
}

bool RegistryStore::WriteKeySecurity(const RegistryNode& node, SECURITY_INFORMATION parts, const std::vector<BYTE>& descriptor, const FILETIME* last_write)
{
    NoteKeyChange();
    return Dispatch(
        node,
        [&](VirtualRegistryData&) { return descriptor.empty(); },
        [&] { return registry_backend::offline::WriteKeySecurity(node, parts, descriptor); },
        [&] { return registry_backend::live::WriteKeySecurity(node, parts, descriptor, last_write); }
    );
}

bool RegistryStore::DeleteKey(const RegistryNode& node)
{
    NoteKeyChange();
    if (node.subkey.empty())
    {
        return false;
    }
    return Dispatch(
        node,
        [&](VirtualRegistryData& data) { return registry_backend::virtual_store::DeleteKey(data, node); },
        [&] { return registry_backend::offline::DeleteKey(node); },
        [&] { return registry_backend::live::DeleteKey(node); }
    );
}

bool RegistryStore::RenameKey(const RegistryNode& node, const std::wstring& new_name)
{
    NoteKeyChange();
    if (node.subkey.empty() || new_name.empty())
    {
        return false;
    }
    const std::wstring old_name = registry_path::Leaf(node.subkey);
    if (const std::wstring step = CaseRenameStep(old_name, new_name); !step.empty())
    {
        RegistryNode moved = node;
        moved.subkey = registry_path::Parent(node.subkey);
        moved = registry_path::ChildNode(moved, step);
        return RenameKey(node, step) && (RenameKey(moved, new_name) || (RenameKey(moved, old_name), false));
    }
    return Dispatch(
        node,
        [&](VirtualRegistryData& data) { return registry_backend::virtual_store::RenameKey(data, node, new_name); },
        [&] { return registry_backend::offline::RenameKey(node, new_name); },
        [&] { return registry_backend::live::RenameKey(node, new_name); }
    );
}

bool RegistryStore::DeleteValue(const RegistryNode& node, const std::wstring& value_name)
{
    return Dispatch(
        node,
        [&](VirtualRegistryData& data) { return registry_backend::virtual_store::DeleteValue(data, node, value_name); },
        [&] { return registry_backend::offline::DeleteValue(node, value_name); },
        [&] { return registry_backend::live::DeleteValue(node, value_name); }
    );
}

bool RegistryStore::SetValue(const RegistryNode& node, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& data)
{
    return Dispatch(
        node,
        [&](VirtualRegistryData& virtual_data) {
            return registry_backend::virtual_store::SetValue(virtual_data, node, value_name, type, data);
        },
        [&] { return registry_backend::offline::SetValue(node, value_name, type, data); },
        [&] { return registry_backend::live::SetValue(node, value_name, type, data); }
    );
}

bool RegistryStore::RenameValue(const RegistryNode& node, const std::wstring& old_name, const std::wstring& new_name, bool* both_names_left)
{
    if (both_names_left)
    {
        *both_names_left = false;
    }
    if (new_name.empty())
    {
        return false;
    }
    if (const std::wstring step = CaseRenameStep(old_name, new_name); !step.empty())
    {
        return RenameValue(node, old_name, step, both_names_left) && (RenameValue(node, step, new_name, both_names_left) || (RenameValue(node, step, old_name, nullptr), false));
    }
    return Dispatch(
        node,
        [&](VirtualRegistryData& data) {
            return registry_backend::virtual_store::RenameValue(data, node, old_name, new_name, both_names_left);
        },
        [&] { return registry_backend::offline::RenameValue(node, old_name, new_name, both_names_left); },
        [&] { return registry_backend::live::RenameValue(node, old_name, new_name, both_names_left); }
    );
}

}
