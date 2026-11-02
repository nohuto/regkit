// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "win32/registry_native.h"
#include "win32/text_transform.h"

#include <win32ui/handle_owner.h>

#include <algorithm>
#include <string>
#include <utility>
#include <vector>

namespace regkit::registry_backend
{

inline constexpr REGSAM kKeyReadAccess = KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS;
inline constexpr size_t kMaxKeyNameLength = 255;
inline constexpr size_t kMaxValueNameLength = 16383;
// RegEnumValueW cost grows with buffer size, start small and retry only larger values
inline constexpr DWORD kValueDataProbe = 64 * 1024;
inline constexpr SECURITY_INFORMATION kKeySecurityInformation =
    OWNER_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | DACL_SECURITY_INFORMATION;

inline const wchar_t* ValueNameArg(const std::wstring& name)
{
    return name.empty() ? nullptr : name.c_str();
}

inline bool HasNull(const std::wstring& name)
{
    return name.find(L'\0') != std::wstring::npos;
}

class RegistryKeyHandle
{
  public:
    RegistryKeyHandle() = default;
    explicit RegistryKeyHandle(win32ui::UniqueHKey key)
        : key_(std::move(key))
    {
    }

    explicit operator bool() const noexcept
    {
        return static_cast<bool>(key_);
    }
    HKEY get() const noexcept
    {
        return key_.get();
    }

    LONG QueryInfo(DWORD* subkeys, DWORD* max_subkey_length, DWORD* values, DWORD* max_value_name_length, DWORD* max_value_data_length, FILETIME* last_write, wchar_t* class_name = nullptr, DWORD* class_length = nullptr, DWORD* max_class_length = nullptr) const
    {
        return RegQueryInfoKeyW(key_.get(), class_name, class_length, nullptr, subkeys, max_subkey_length, max_class_length, values, max_value_name_length, max_value_data_length, nullptr, last_write);
    }
    LONG EnumKey(DWORD index, wchar_t* name, DWORD* length) const
    {
        return RegEnumKeyExW(key_.get(), index, name, length, nullptr, nullptr, nullptr, nullptr);
    }
    LONG EnumValue(DWORD index, wchar_t* name, DWORD* name_length, DWORD* type, BYTE* data, DWORD* data_length) const
    {
        return RegEnumValueW(key_.get(), index, name, name_length, nullptr, type, data, data_length);
    }
    LONG GetValue(const std::wstring& name, DWORD* type, BYTE* data, DWORD* size) const
    {
        return HasNull(name) ? util::QueryValueCounted(key_.get(), name, type, data, size)
                             : RegQueryValueExW(key_.get(), ValueNameArg(name), nullptr, type, data, size);
    }
    LONG SetValue(const std::wstring& name, DWORD type, const BYTE* data, DWORD size) const
    {
        return HasNull(name) ? util::SetValueCounted(key_.get(), name, type, data, size)
                             : RegSetValueExW(key_.get(), ValueNameArg(name), 0, type, data, size);
    }
    LONG DeleteValue(const std::wstring& name) const
    {
        return HasNull(name) ? util::DeleteValueCounted(key_.get(), name) : RegDeleteValueW(key_.get(), ValueNameArg(name));
    }
    LONG GetSecurity(SECURITY_INFORMATION information, PSECURITY_DESCRIPTOR descriptor, DWORD* size) const
    {
        return RegGetKeySecurity(key_.get(), information, descriptor, size);
    }
    LONG SetSecurity(SECURITY_INFORMATION information, PSECURITY_DESCRIPTOR descriptor) const
    {
        return RegSetKeySecurity(key_.get(), information, descriptor);
    }

  protected:
    win32ui::UniqueHKey key_;
};

inline bool SplitNode(const RegistryNode& node, RegistryNode* parent, std::wstring* name)
{
    *name = registry_path::Leaf(node.subkey);
    *parent = node;
    parent->subkey = registry_path::Parent(node.subkey);
    return !name->empty();
}

inline void SortNames(std::vector<std::wstring>* names)
{
    std::sort(names->begin(), names->end(), [](const std::wstring& left, const std::wstring& right) {
        return util::CompareInsensitive(left, right) < 0;
    });
}

template <typename Key>
bool HasSubKeys(const Key& key)
{
    DWORD count = 0;
    return key.QueryInfo(&count, nullptr, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS && count > 0;
}

template <typename Key>
bool QueryKeyInfo(const Key& key, KeyInfo* info)
{
    return key.QueryInfo(&info->subkey_count, nullptr, &info->value_count, nullptr, nullptr, &info->last_write) ==
           ERROR_SUCCESS;
}

template <typename Key>
LONG QueryKeyDetails(const Key& key, KeyDetails* details)
{
    std::wstring& name = details->class_name;
    LONG result = ERROR_MORE_DATA;
    for (DWORD capacity : {DWORD(MAX_PATH), DWORD(32767)})
    {
        if (result != ERROR_MORE_DATA)
        {
            break;
        }
        name.resize(capacity);
        DWORD length = capacity;
        result = key.QueryInfo(&details->info.subkey_count, &details->max_subkey_name, &details->info.value_count, &details->max_value_name, &details->max_value_data, &details->info.last_write, name.data(), &length, &details->max_class);
        name.resize(result == ERROR_SUCCESS ? length : 0);
    }
    return result;
}

template <typename Key>
LONG EnumKeyName(const Key& key, DWORD index, std::wstring* name, DWORD* length)
{
    *length = static_cast<DWORD>(name->size());
    LONG result = key.EnumKey(index, name->data(), length);
    if (result == ERROR_MORE_DATA)
    {
        name->resize(kMaxKeyNameLength + 1);
        *length = static_cast<DWORD>(name->size());
        result = key.EnumKey(index, name->data(), length);
    }
    return result;
}

template <typename Key>
std::vector<std::wstring> SubKeyNames(const Key& key, bool sorted)
{
    std::vector<std::wstring> names;
    DWORD count = 0;
    DWORD max_length = 0;
    if (key.QueryInfo(&count, &max_length, nullptr, nullptr, nullptr, nullptr) != ERROR_SUCCESS)
    {
        return names;
    }
    names.reserve(count);
    std::wstring buffer(max_length + 1, L'\0');
    for (DWORD index = 0; index < count; ++index)
    {
        DWORD length = 0;
        const LONG result = EnumKeyName(key, index, &buffer, &length);
        if (result == ERROR_NO_MORE_ITEMS)
        {
            break;
        }
        if (result == ERROR_SUCCESS)
        {
            names.emplace_back(buffer.data(), length);
        }
    }
    if (sorted)
    {
        SortNames(&names);
    }
    return names;
}

template <typename Key>
bool EnumerateKey(const Key& key, bool include_values, bool include_data, bool include_subkeys, RegistryStore::KeyEnumResult* out_info, const RegistryStore::ValueStreamCallback& value_callback, const RegistryStore::SubkeyStreamCallback& subkey_callback, DWORD max_data_size, EnumerationScratch* scratch)
{
    EnumerationScratch local;
    EnumerationScratch& buffers = scratch ? *scratch : local;
    KeyDetails details;
    const bool want_options = out_info && out_info->want_options;
    const LONG status = want_options ? QueryKeyDetails(key, &details)
                                     : key.QueryInfo(&details.info.subkey_count, &details.max_subkey_name, &details.info.value_count, &details.max_value_name, &details.max_value_data, &details.info.last_write);
    if (status != ERROR_SUCCESS)
    {
        if (out_info)
        {
            out_info->error = status;
        }
        return false;
    }
    const auto fail = [&](LONG result) {
        if (out_info && out_info->error == ERROR_SUCCESS && result != ERROR_NO_MORE_ITEMS)
        {
            out_info->error = result;
        }
    };
    const KeyInfo& info = details.info;
    const DWORD max_subkey_length = details.max_subkey_name;
    const DWORD max_value_name_length = details.max_value_name;
    const DWORD max_value_data_length = details.max_value_data;
    if (out_info)
    {
        out_info->info = info;
        out_info->info_valid = true;
        if (want_options)
        {
            out_info->options.class_name = std::move(details.class_name);
        }
    }

    if (include_values && value_callback)
    {
        std::wstring& name = buffers.value_name;
        std::vector<BYTE>& data = buffers.value_data;
        name.resize(static_cast<size_t>(max_value_name_length) + 1);
        if (include_data)
        {
            data.resize(std::min({max_value_data_length, max_data_size, kValueDataProbe}));
        }
        ValueInfo value;
        for (DWORD index = 0; index < info.value_count; ++index)
        {
            DWORD name_length = static_cast<DWORD>(name.size());
            DWORD data_length = include_data ? static_cast<DWORD>(std::min<size_t>(data.size(), kValueDataProbe)) : 0;
            DWORD type = 0;
            BYTE* buffer = include_data && !data.empty() ? data.data() : nullptr;
            LONG result = key.EnumValue(index, name.data(), &name_length, &type, buffer, &data_length);
            if (result == ERROR_MORE_DATA || (result == ERROR_SUCCESS && include_data && !buffer && data_length > 0 && data_length <= max_data_size))
            {
                name.resize(std::max(name.size(), kMaxValueNameLength + 1));
                buffer = nullptr;
                if (include_data && data_length <= max_data_size)
                {
                    data.resize(std::max<size_t>(data.size(), data_length));
                    buffer = data.empty() ? nullptr : data.data();
                }
                name_length = static_cast<DWORD>(name.size());
                data_length = buffer ? static_cast<DWORD>(data.size()) : 0;
                result = key.EnumValue(index, name.data(), &name_length, &type, buffer, &data_length);
            }
            // oversized buffers let API terminators overwrite the last byte of odd sized strings
            if (result == ERROR_SUCCESS && buffer && (data_length & 1) && data_length < data.size() && (type == REG_SZ || type == REG_EXPAND_SZ))
            {
                name_length = static_cast<DWORD>(name.size());
                result = key.EnumValue(index, name.data(), &name_length, &type, buffer, &data_length);
            }
            if (result != ERROR_SUCCESS)
            {
                fail(result);
                if (result == ERROR_NO_MORE_ITEMS)
                {
                    break;
                }
                continue;
            }
            value.name.assign(name.data(), name_length);
            value.type = type;
            value.data_size = data_length;
            if (!value_callback(value, buffer && data_length > 0 ? buffer : nullptr, data_length))
            {
                return false;
            }
        }
    }

    if (include_subkeys && subkey_callback)
    {
        std::wstring& name = buffers.subkey_name;
        name.resize(static_cast<size_t>(max_subkey_length) + 1);
        for (DWORD index = 0; index < info.subkey_count; ++index)
        {
            DWORD name_length = 0;
            const LONG result = EnumKeyName(key, index, &name, &name_length);
            if (result != ERROR_SUCCESS)
            {
                fail(result);
                if (result == ERROR_NO_MORE_ITEMS)
                {
                    break;
                }
                continue;
            }
            if (!subkey_callback(std::wstring(name.data(), name_length)))
            {
                return false;
            }
        }
    }
    return true;
}

struct KeyContents
{
    std::vector<RegistryValue> values;
    std::vector<std::wstring> subkeys;
};

inline LONG ReadKeyContents(HKEY root, const std::wstring& subkey, REGSAM view, bool include_data, KeyContents* contents, bool* is_volatile = nullptr)
{
    win32ui::UniqueHKey handle;
    const LONG status = util::OpenRegistryPath(root, subkey, kKeyReadAccess | view, false, &handle);
    if (status != ERROR_SUCCESS)
    {
        return status;
    }
    if (is_volatile)
    {
        *is_volatile = util::IsLocalRoot(root) && ((util::QueryKeyFlags(handle.get()).value_or(0) & util::kKeyFlagVolatile) ||
                                                   (root == HKEY_LOCAL_MACHINE && registry_path::InVolatileHive(L"\\REGISTRY\\MACHINE\\" + subkey)));
    }
    RegistryStore::KeyEnumResult result;
    EnumerateKey(
        RegistryKeyHandle(std::move(handle)),
        true,
        include_data,
        true,
        &result,
        [&](const ValueInfo& info, const BYTE* data, DWORD size) {
            contents->values.push_back({info.name, info.type, data ? std::vector<BYTE>(data, data + size) : std::vector<BYTE>()});
            return true;
        },
        [&](const std::wstring& name) {
            contents->subkeys.push_back(name);
            return true;
        },
        MAXDWORD,
        nullptr
    );
    return result.error;
}

template <typename Key>
LONG ReadValue(const Key& key, const std::wstring& name, DWORD* type, std::vector<BYTE>* data)
{
    data->clear();
    DWORD size = 0;
    LONG result = key.GetValue(name, type, nullptr, &size);
    for (int attempt = 0; attempt < 4 && (result == ERROR_SUCCESS || result == ERROR_MORE_DATA); ++attempt)
    {
        if (result == ERROR_SUCCESS && size <= data->size())
        {
            data->resize(size);
            return result;
        }
        data->resize(size);
        result = key.GetValue(name, type, data->empty() ? nullptr : data->data(), &size);
    }
    data->clear();
    return result == ERROR_SUCCESS ? ERROR_MORE_DATA : result;
}

template <typename Key>
bool QueryValue(const Key& key, const std::wstring& value_name, RegistryValue* out)
{
    DWORD type = 0;
    std::vector<BYTE> data;
    if (ReadValue(key, value_name, &type, &data) != ERROR_SUCCESS)
    {
        return false;
    }
    out->name = value_name;
    out->type = type;
    out->data = std::move(data);
    return true;
}

template <typename Key>
bool ReadLinkTarget(const Key& key, std::wstring* target)
{
    DWORD type = 0;
    std::vector<BYTE> data;
    static const std::wstring kLinkValue = L"SymbolicLinkValue";
    if (ReadValue(key, kLinkValue, &type, &data) != ERROR_SUCCESS || type != REG_LINK)
    {
        return false;
    }
    target->assign(reinterpret_cast<const wchar_t*>(data.data()), data.size() / sizeof(wchar_t));
    while (!target->empty() && target->back() == L'\0')
    {
        target->pop_back();
    }
    return true;
}

template <typename Key>
bool RenameValue(const Key& key, const std::wstring& old_name, const std::wstring& new_name, bool* both_names_left)
{
    DWORD type = 0;
    DWORD existing_size = 0;
    std::vector<BYTE> data;
    if (ReadValue(key, old_name, &type, &data) != ERROR_SUCCESS ||
        key.GetValue(new_name, nullptr, nullptr, &existing_size) != ERROR_FILE_NOT_FOUND ||
        key.SetValue(new_name, type, data.empty() ? nullptr : data.data(), static_cast<DWORD>(data.size())) != ERROR_SUCCESS)
    {
        return false;
    }
    if (key.DeleteValue(old_name) != ERROR_SUCCESS)
    {
        if (key.DeleteValue(new_name) != ERROR_SUCCESS && both_names_left)
        {
            *both_names_left = true;
        }
        return false;
    }
    return true;
}

template <typename Key>
bool ReadSecurity(const Key& key, SECURITY_INFORMATION* parts, std::vector<BYTE>* descriptor, LONG* error = nullptr)
{
    // fall back to owner, group and DACL if label or SACL cannot be read
    LONG result = ERROR_SUCCESS;
    for (const SECURITY_INFORMATION request : {*parts, kKeySecurityInformation})
    {
        DWORD size = 0;
        result = key.GetSecurity(request, nullptr, &size);
        if ((result == ERROR_INSUFFICIENT_BUFFER || result == ERROR_MORE_DATA) && size > 0)
        {
            descriptor->resize(size);
            result = key.GetSecurity(request, descriptor->data(), &size);
        }
        if (result == ERROR_SUCCESS && size > 0)
        {
            descriptor->resize(size);
            *parts = request;
            return true;
        }
    }
    descriptor->clear();
    if (error)
    {
        *error = result;
    }
    return false;
}

template <typename Key>
bool WriteSecurity(const Key& key, SECURITY_INFORMATION parts, const std::vector<BYTE>& descriptor)
{
    // keep DACL if the token cannot set owner, label or SACL
    if (descriptor.empty())
    {
        return false;
    }
    for (const SECURITY_INFORMATION request : {parts, parts & kKeySecurityInformation, parts & DACL_SECURITY_INFORMATION})
    {
        if (request && key.SetSecurity(request, const_cast<BYTE*>(descriptor.data())) == ERROR_SUCCESS)
        {
            return true;
        }
    }
    return false;
}

}
