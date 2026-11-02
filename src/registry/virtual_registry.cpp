// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_backends.h"

#include "registry/key_algorithms.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <memory>
#include <mutex>
#include <unordered_map>

namespace regkit::registry_backend::virtual_store
{
namespace
{

struct RootEntry
{
    std::wstring root_name;
    std::shared_ptr<VirtualRegistryData> data;
    std::unique_ptr<int> handle_tag;
};

std::mutex g_roots_mutex;
std::unordered_map<HKEY, RootEntry> g_roots;

template <typename Key>
Key* FindKey(Key* current, const std::wstring& subkey)
{
    std::wstring part;
    for (size_t start = 0; current && start < subkey.size(); ++start)
    {
        const size_t end = std::min(subkey.find(L'\\', start), subkey.size());
        if (end > start)
        {
            part.assign(subkey, start, end - start);
            CharLowerBuffW(part.data(), static_cast<DWORD>(part.size()));
            const auto child = current->children.find(part);
            current = child == current->children.end() ? nullptr : child->second.get();
        }
        start = end;
    }
    return current;
}

std::vector<std::wstring> ChildNames(const VirtualRegistryKey& key, bool sorted)
{
    std::vector<std::wstring> names;
    names.reserve(key.children.size());
    for (const auto& child : key.children)
    {
        if (child.second)
        {
            names.push_back(child.second->name);
        }
    }
    if (sorted)
    {
        SortNames(&names);
    }
    return names;
}

}

HKEY RegisterRoot(const std::wstring& root_name, const std::shared_ptr<VirtualRegistryData>& data)
{
    if (!data)
    {
        return nullptr;
    }
    auto handle_tag = std::make_unique<int>(0);
    const HKEY handle = reinterpret_cast<HKEY>(handle_tag.get());
    RootEntry entry;
    entry.root_name = root_name;
    entry.data = data;
    entry.handle_tag = std::move(handle_tag);
    std::lock_guard<std::mutex> lock(g_roots_mutex);
    g_roots.emplace(handle, std::move(entry));
    return handle;
}

void UnregisterRoot(HKEY root)
{
    std::lock_guard<std::mutex> lock(g_roots_mutex);
    g_roots.erase(root);
}

std::shared_ptr<VirtualRegistryData> FindRoot(HKEY root, std::wstring* root_name)
{
    if (root_name)
    {
        root_name->clear();
    }
    if (!root)
    {
        return {};
    }
    std::lock_guard<std::mutex> lock(g_roots_mutex);
    const auto entry = g_roots.find(root);
    if (entry == g_roots.end())
    {
        return {};
    }
    if (root_name)
    {
        *root_name = entry->second.root_name;
    }
    return entry->second.data;
}

bool HasSubKeys(const VirtualRegistryData& data, const RegistryNode& node)
{
    std::shared_lock<std::shared_mutex> lock(*data.mutex);
    const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    return key && !key->children.empty();
}

bool QueryKeyInfo(const VirtualRegistryData& data, const RegistryNode& node, KeyInfo* info)
{
    std::shared_lock<std::shared_mutex> lock(*data.mutex);
    if (!info)
    {
        return false;
    }
    const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    if (!key)
    {
        return false;
    }
    info->subkey_count = static_cast<DWORD>(key->children.size());
    info->value_count = static_cast<DWORD>(key->values.size());
    info->last_write = {};
    return true;
}

std::vector<std::wstring> EnumSubKeyNames(const VirtualRegistryData& data, const RegistryNode& node, bool sorted)
{
    std::shared_lock<std::shared_mutex> lock(*data.mutex);
    const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    if (!key)
    {
        return {};
    }
    return ChildNames(*key, sorted);
}

bool EnumKeyStreaming(const VirtualRegistryData& data, const RegistryNode& node, bool include_values, bool include_data, bool include_subkeys, RegistryStore::KeyEnumResult* out_info, const RegistryStore::ValueStreamCallback& value_callback, const RegistryStore::SubkeyStreamCallback& subkey_callback, DWORD max_data_size, EnumerationScratch* scratch, bool ordered)
{
    (void)scratch;
    std::shared_lock<std::shared_mutex> lock(*data.mutex);
    const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    if (!key)
    {
        return false;
    }
    if (out_info)
    {
        out_info->info.subkey_count = static_cast<DWORD>(key->children.size());
        out_info->info.value_count = static_cast<DWORD>(key->values.size());
        out_info->info.last_write = {};
        out_info->info_valid = true;
    }
    if (include_values && value_callback)
    {
        std::vector<const RegistryValue*> values;
        values.reserve(key->values.size());
        for (const auto& value : key->values)
        {
            values.push_back(&value.second);
        }
        if (ordered)
        {
            std::sort(values.begin(), values.end(), [](const RegistryValue* left, const RegistryValue* right) {
                return util::CompareInsensitive(left->name, right->name) < 0;
            });
        }
        for (const RegistryValue* value : values)
        {
            ValueInfo info;
            info.name = value->name;
            info.type = value->type;
            info.data_size = static_cast<DWORD>(value->data.size());
            const bool data_available = include_data && value->data.size() <= max_data_size;
            const BYTE* buffer = data_available && !value->data.empty() ? value->data.data() : nullptr;
            if (!value_callback(info, buffer, static_cast<DWORD>(value->data.size())))
            {
                return false;
            }
        }
    }
    if (include_subkeys && subkey_callback)
    {
        for (const auto& name : ChildNames(*key, ordered))
        {
            if (!subkey_callback(name))
            {
                return false;
            }
        }
    }
    return true;
}

bool QueryValue(const VirtualRegistryData& data, const RegistryNode& node, const std::wstring& value_name, RegistryValue* out)
{
    std::shared_lock<std::shared_mutex> lock(*data.mutex);
    if (!out)
    {
        return false;
    }
    const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    if (!key)
    {
        return false;
    }
    const auto value = key->values.find(util::ToLower(value_name));
    if (value == key->values.end())
    {
        return false;
    }
    *out = value->second;
    return true;
}

bool CreateKey(VirtualRegistryData& data, const RegistryNode& node, const std::wstring& name)
{
    std::unique_lock<std::shared_mutex> lock(*data.mutex);
    VirtualRegistryKey* parent = FindKey(data.root.get(), node.subkey);
    if (!parent)
    {
        return false;
    }
    const std::wstring lower_name = util::ToLower(name);
    if (parent->children.find(lower_name) != parent->children.end())
    {
        return false;
    }
    auto child = std::make_unique<VirtualRegistryKey>();
    child->name = name;
    parent->children.emplace(lower_name, std::move(child));
    return true;
}

bool DeleteKey(VirtualRegistryData& data, const RegistryNode& node)
{
    std::unique_lock<std::shared_mutex> lock(*data.mutex);
    RegistryNode parent_node;
    std::wstring name;
    if (!SplitNode(node, &parent_node, &name))
    {
        return false;
    }
    VirtualRegistryKey* parent = FindKey(data.root.get(), parent_node.subkey);
    return parent && parent->children.erase(util::ToLower(name)) != 0;
}

bool RenameKey(VirtualRegistryData& data, const RegistryNode& node, const std::wstring& new_name)
{
    std::unique_lock<std::shared_mutex> lock(*data.mutex);
    RegistryNode parent_node;
    std::wstring old_name;
    if (!SplitNode(node, &parent_node, &old_name))
    {
        return false;
    }
    VirtualRegistryKey* parent = FindKey(data.root.get(), parent_node.subkey);
    if (!parent)
    {
        return false;
    }
    std::wstring new_lower = util::ToLower(new_name);
    parent->children.reserve(parent->children.size() + 1);
    auto source = parent->children.find(util::ToLower(old_name));
    if (source == parent->children.end() || !source->second || parent->children.find(new_lower) != parent->children.end())
    {
        return false;
    }
    std::wstring name = new_name;
    auto node_handle = parent->children.extract(source);
    node_handle.key() = std::move(new_lower);
    node_handle.mapped()->name.swap(name);
    parent->children.insert(std::move(node_handle));
    return true;
}

bool DeleteValue(VirtualRegistryData& data, const RegistryNode& node, const std::wstring& value_name)
{
    std::unique_lock<std::shared_mutex> lock(*data.mutex);
    VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    return key && key->values.erase(util::ToLower(value_name)) != 0;
}

bool SetValue(VirtualRegistryData& data, const RegistryNode& node, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& value_data)
{
    std::unique_lock<std::shared_mutex> lock(*data.mutex);
    VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    if (!key)
    {
        return false;
    }
    RegistryValue& value = key->values[util::ToLower(value_name)];
    value.name = value_name;
    value.type = type;
    value.data = value_data;
    return true;
}

bool RenameValue(VirtualRegistryData& data, const RegistryNode& node, const std::wstring& old_name, const std::wstring& new_name, bool* both_names_left)
{
    if (both_names_left)
    {
        *both_names_left = false;
    }
    std::unique_lock<std::shared_mutex> lock(*data.mutex);
    VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
    if (!key)
    {
        return false;
    }
    std::wstring new_lower = util::ToLower(new_name);
    key->values.reserve(key->values.size() + 1);
    auto source = key->values.find(util::ToLower(old_name));
    if (source == key->values.end() || key->values.find(new_lower) != key->values.end())
    {
        return false;
    }
    std::wstring name = new_name;
    auto node_handle = key->values.extract(source);
    node_handle.key() = std::move(new_lower);
    node_handle.mapped().name.swap(name);
    key->values.insert(std::move(node_handle));
    return true;
}

}
