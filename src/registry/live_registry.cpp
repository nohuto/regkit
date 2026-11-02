// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_backends.h"

#include "registry/key_algorithms.h"
#include "win32/registry_native.h"
#include "win32/registry_view.h"

#include <win32ui/handle_owner.h>

namespace regkit::registry_backend::live
{
namespace
{

class LiveKey : public RegistryKeyHandle
{
  public:
    LiveKey(const RegistryNode& node, REGSAM access, bool open_link = false)
    {
        if (node.root)
        {
            error_ = util::OpenRegistryPath(node.root, node.subkey, access | ViewOf(node), open_link, &key_);
        }
    }
    LiveKey() = default;
    using RegistryKeyHandle::RegistryKeyHandle;
    LONG error() const noexcept
    {
        return error_;
    }

  private:
    LONG error_ = ERROR_INVALID_HANDLE;
};
LiveKey OpenChild(const RegistryNode& node, REGSAM parent_access, REGSAM child_access, bool open_link, std::wstring* name = nullptr)
{
    RegistryNode parent_node;
    std::wstring leaf;
    win32ui::UniqueHKey child;
    if (SplitNode(node, &parent_node, &leaf))
    {
        LiveKey parent(parent_node, parent_access);
        if (parent)
        {
            util::OpenRegistryPath(parent.get(), leaf, child_access | ViewOf(node), open_link, &child);
        }
    }
    if (name)
    {
        *name = std::move(leaf);
    }
    return LiveKey(std::move(child));
}

bool KernelHandle(const RegistryNode& node)
{
    return util::IsLocalRoot(node.root) || util::EqualsInsensitive(node.root_name, L"REGISTRY");
}

// opens the key itself, not a link target
LiveKey OpenInspected(const RegistryNode& node, const std::wstring& native, LONG* error)
{
    if (node.view || (native.empty() && node.root == HKEY_CLASSES_ROOT && !node.subkey.empty()))
    {
        // HKCR has no native path and 32 bit views redirect, use Win32 opens
        win32ui::UniqueHKey handle;
        *error = util::OpenRegistryPath(node.root, node.subkey, KEY_QUERY_VALUE | ViewOf(node), true, &handle);
        return LiveKey(std::move(handle));
    }
    return LiveKey(util::OpenNativeRegistryKey(native, KEY_QUERY_VALUE, true, error));
}

LiveKey OpenSecurity(const RegistryNode& node, REGSAM access, SECURITY_INFORMATION* parts)
{
    if (*parts & SACL_SECURITY_INFORMATION)
    {
        LiveKey key(node, access | ACCESS_SYSTEM_SECURITY, true);
        if (key)
        {
            return key;
        }
        *parts &= ~SACL_SECURITY_INFORMATION;
    }
    return LiveKey(node, access, true);
}

}

bool HasSubKeys(const RegistryNode& node)
{
    LiveKey key(node, KEY_QUERY_VALUE);
    return key && registry_backend::HasSubKeys(key);
}

bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info)
{
    LiveKey key(node, kKeyReadAccess);
    return key && registry_backend::QueryKeyInfo(key, info);
}

bool QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target, bool* denied)
{
    target->clear();
    LONG error = ERROR_SUCCESS;
    const LiveKey key = OpenInspected(node, registry_path::BuildNative(node), &error);
    if (denied)
    {
        *denied = error == ERROR_ACCESS_DENIED;
    }
    return key && ReadLinkTarget(key, target) && !target->empty();
}

bool IsBrokenLink(const RegistryNode& node, std::wstring* target)
{
    if (!QuerySymbolicLinkTarget(node, target, nullptr))
    {
        return false;
    }
    LONG error = ERROR_SUCCESS;
    util::OpenNativeRegistryKey(*target, KEY_QUERY_VALUE, false, &error);
    return error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND;
}

KeyInspection InspectKey(const RegistryNode& node, bool want_info, bool want_source)
{
    KeyInspection result;
    LONG error = ERROR_SUCCESS;
    // HARDWARE never redirects, native paths also identify volatile hives for 32 bit nodes
    const std::wstring native = registry_path::BuildNative(node);
    const LiveKey key = OpenInspected(node, native, &error);
    // merged keys report their backing hive
    if (want_source && key && node.root == HKEY_CLASSES_ROOT)
    {
        result.class_source = registry_path::ClassesSource(util::QueryKeyName(key.get()));
    }
    result.denied = error == ERROR_ACCESS_DENIED;
    std::wstring target;
    // flags avoid reading link values for ordinary keys
    const std::optional<ULONG> flags = key ? util::QueryKeyFlags(key.get()) : std::nullopt;
    result.link = key && (!flags || (*flags & util::kKeyFlagLink)) && ReadLinkTarget(key, &target) && !target.empty();
    result.is_volatile = (flags.value_or(0) & util::kKeyFlagVolatile) || registry_path::InVolatileHive(native);
    if (key && want_info && !result.link)
    {
        result.info_valid = registry_backend::QueryKeyInfo(key, &result.info);
    }
    return result;
}

std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node, bool sorted)
{
    LiveKey key(node, kKeyReadAccess);
    return key ? SubKeyNames(key, sorted) : std::vector<std::wstring>();
}

bool EnumKeyStreaming(const RegistryNode& node, bool include_values, bool include_data, bool include_subkeys, RegistryStore::KeyEnumResult* out_info, const RegistryStore::ValueStreamCallback& value_callback, const RegistryStore::SubkeyStreamCallback& subkey_callback, DWORD max_data_size, EnumerationScratch* scratch, bool, bool open_link)
{
    LiveKey key(node, kKeyReadAccess, open_link);
    if (key && out_info && out_info->want_options && KernelHandle(node))
    {
        out_info->options.is_volatile = util::QueryKeyFlags(key.get()).value_or(0) & util::kKeyFlagVolatile;
    }
    if (!key && out_info)
    {
        out_info->error = key.error();
    }
    return key && EnumerateKey(key, include_values, include_data, include_subkeys, out_info, value_callback, subkey_callback, max_data_size, scratch);
}

bool QueryValue(const RegistryNode& node, const std::wstring& value_name, RegistryValue* out)
{
    LiveKey key(node, KEY_QUERY_VALUE);
    return key && registry_backend::QueryValue(key, value_name, out);
}

bool QueryKeyDetails(const RegistryNode& node, KeyDetails* details, bool open_link)
{
    LiveKey key(node, kKeyReadAccess, open_link);
    const bool read = key && registry_backend::QueryKeyDetails(key, details) == ERROR_SUCCESS;
    // native names require no access, available even for keys denying queries
    const LiveKey named = key ? LiveKey() : LiveKey(node, MAXIMUM_ALLOWED, open_link);
    if (KernelHandle(node) && (key || named))
    {
        details->native = util::QueryNativeKeyInfo((key ? key : named).get());
    }
    return read;
}

bool CreateKey(const RegistryNode& node, const std::wstring& name, const KeyCreateOptions& options, bool* created_volatile)
{
    LiveKey parent(node, KEY_WRITE);
    if (!parent)
    {
        return false;
    }
    win32ui::UniqueHKey created;
    DWORD disposition = 0;
    DWORD flags = options.is_volatile ? REG_OPTION_VOLATILE : REG_OPTION_NON_VOLATILE;
    LONG result = util::CreateRegistryKey(parent.get(), name, KEY_READ | KEY_WRITE, flags, &created, &disposition, options.class_name);
    if (result == ERROR_CHILD_MUST_BE_VOLATILE)
    {
        flags = REG_OPTION_VOLATILE;
        result = util::CreateRegistryKey(parent.get(), name, KEY_READ | KEY_WRITE, flags, &created, &disposition, options.class_name);
    }
    if (created_volatile)
    {
        *created_volatile = flags == REG_OPTION_VOLATILE;
    }
    return result == ERROR_SUCCESS && disposition == REG_CREATED_NEW_KEY;
}

bool CreateRegistryLink(const RegistryNode& node, const std::wstring& name, const std::wstring& nt_target, DWORD* error, const KeyCreateOptions& options)
{
    LONG result = ERROR_ACCESS_DENIED;
    LiveKey parent(node, KEY_WRITE);
    win32ui::UniqueHKey created;
    DWORD disposition = 0;
    if (parent)
    {
        const DWORD flags = (options.is_volatile ? REG_OPTION_VOLATILE : REG_OPTION_NON_VOLATILE) | REG_OPTION_CREATE_LINK;
        result = util::CreateRegistryKey(parent.get(), name, KEY_SET_VALUE | KEY_CREATE_LINK | DELETE, flags, &created, &disposition, options.class_name);
        if (result == ERROR_CHILD_MUST_BE_VOLATILE)
        {
            result = util::CreateRegistryKey(parent.get(), name, KEY_SET_VALUE | KEY_CREATE_LINK | DELETE, flags | REG_OPTION_VOLATILE, &created, &disposition, options.class_name);
        }
        if (result == ERROR_SUCCESS && disposition != REG_CREATED_NEW_KEY)
        {
            result = ERROR_ALREADY_EXISTS;
        }
    }
    if (result == ERROR_SUCCESS)
    {
        result = RegSetValueExW(created.get(), L"SymbolicLinkValue", 0, REG_LINK, reinterpret_cast<const BYTE*>(nt_target.c_str()), static_cast<DWORD>(nt_target.size() * sizeof(wchar_t)));
        if (result != ERROR_SUCCESS)
        {
            util::DeleteNativeRegistryKey(created.get());
        }
    }
    if (error)
    {
        *error = static_cast<DWORD>(result);
    }
    return result == ERROR_SUCCESS;
}

bool ReadKeyLink(const RegistryNode& node, std::wstring* target)
{
    std::wstring value;
    const LiveKey link = OpenChild(node, kKeyReadAccess, KEY_QUERY_VALUE, true);
    if (!link || !ReadLinkTarget(link, &value))
    {
        return false;
    }
    if (target)
    {
        *target = std::move(value);
    }
    return true;
}

bool ReadKeySecurity(const RegistryNode& node, SECURITY_INFORMATION* parts, std::vector<BYTE>* descriptor, LONG* error)
{
    descriptor->clear();
    const LiveKey key = OpenSecurity(node, READ_CONTROL, parts);
    if (!key && error)
    {
        *error = key.error();
    }
    return key && ReadSecurity(key, parts, descriptor, error);
}

bool WriteKeySecurity(const RegistryNode& node, SECURITY_INFORMATION parts, const std::vector<BYTE>& descriptor, const FILETIME* last_write)
{
    // security writes change last write time, restore it through the same handle
    const SECURITY_INFORMATION requested = parts;
    LiveKey key = OpenSecurity(node, WRITE_DAC | WRITE_OWNER | (last_write ? KEY_SET_VALUE : 0), &parts);
    if (!key && last_write)
    {
        parts = requested;
        key = OpenSecurity(node, WRITE_DAC | WRITE_OWNER, &parts);
    }
    const bool written = key && (descriptor.empty() || WriteSecurity(key, parts, descriptor));
    if (written && last_write && KernelHandle(node))
    {
        util::SetKeyLastWriteTime(key.get(), *last_write);
    }
    return written;
}

bool DeleteKey(const RegistryNode& node)
{
    const LiveKey target =
        OpenChild(node, KEY_ENUMERATE_SUB_KEYS, DELETE | KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE, true);
    return target && util::DeleteRegistryTree(target.get()) == ERROR_SUCCESS;
}

bool RenameKey(const RegistryNode& node, const std::wstring& new_name)
{
    RegistryNode parent_node;
    std::wstring old_name;
    if (!SplitNode(node, &parent_node, &old_name))
    {
        return false;
    }
    LiveKey parent(parent_node, KEY_WRITE);
    return parent && util::RenameRegistryKey(parent.get(), old_name, new_name) == ERROR_SUCCESS;
}

bool DeleteValue(const RegistryNode& node, const std::wstring& value_name)
{
    LiveKey key(node, KEY_SET_VALUE);
    return key && key.DeleteValue(value_name) == ERROR_SUCCESS;
}

bool SetValue(const RegistryNode& node, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& data)
{
    LiveKey key(node, KEY_SET_VALUE);
    return key && key.SetValue(value_name, type, data.empty() ? nullptr : data.data(), static_cast<DWORD>(data.size())) == ERROR_SUCCESS;
}

bool RenameValue(const RegistryNode& node, const std::wstring& old_name, const std::wstring& new_name, bool* both_names_left)
{
    LiveKey key(node, KEY_QUERY_VALUE | KEY_SET_VALUE);
    return key && registry_backend::RenameValue(key, old_name, new_name, both_names_left);
}

}
