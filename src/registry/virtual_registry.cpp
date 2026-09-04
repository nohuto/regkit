// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_backends.h"

#include "registry/registry_path.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <memory>
#include <mutex>
#include <unordered_map>

namespace regkit::registry_backend::virtual_store {
namespace {

struct RootEntry {
  std::wstring root_name;
  std::shared_ptr<VirtualRegistryData> data;
  std::unique_ptr<int> handle_tag;
};

std::mutex g_roots_mutex;
std::unordered_map<HKEY, RootEntry> g_roots;

VirtualRegistryKey* FindKey(VirtualRegistryKey* root,
                            const std::wstring& subkey) {
  if (!root) {
    return nullptr;
  }
  if (subkey.empty()) {
    return root;
  }
  VirtualRegistryKey* current = root;
  for (const auto& part : registry_path::Split(subkey)) {
    auto child = current->children.find(util::ToLower(part));
    if (child == current->children.end()) {
      return nullptr;
    }
    current = child->second.get();
  }
  return current;
}

const VirtualRegistryKey* FindKey(const VirtualRegistryKey* root,
                                  const std::wstring& subkey) {
  if (!root) {
    return nullptr;
  }
  if (subkey.empty()) {
    return root;
  }
  const VirtualRegistryKey* current = root;
  for (const auto& part : registry_path::Split(subkey)) {
    auto child = current->children.find(util::ToLower(part));
    if (child == current->children.end()) {
      return nullptr;
    }
    current = child->second.get();
  }
  return current;
}

bool SplitNode(const RegistryNode& node, std::wstring* parent,
               std::wstring* name) {
  if (!parent || !name) {
    return false;
  }
  *parent = registry_path::Parent(node.subkey);
  *name = registry_path::Leaf(node.subkey);
  return !name->empty();
}

std::vector<std::wstring> ChildNames(const VirtualRegistryKey& key,
                                     bool sorted) {
  std::vector<std::wstring> names;
  names.reserve(key.children.size());
  for (const auto& child : key.children) {
    if (child.second) {
      names.push_back(child.second->name);
    }
  }
  if (sorted) {
    std::sort(names.begin(), names.end(),
              [](const std::wstring& left, const std::wstring& right) {
                return _wcsicmp(left.c_str(), right.c_str()) < 0;
              });
  }
  return names;
}

} // namespace

HKEY RegisterRoot(const std::wstring& root_name,
                  const std::shared_ptr<VirtualRegistryData>& data) {
  if (!data) {
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

void UnregisterRoot(HKEY root) {
  std::lock_guard<std::mutex> lock(g_roots_mutex);
  g_roots.erase(root);
}

std::shared_ptr<VirtualRegistryData>
FindRoot(HKEY root, std::wstring* root_name) {
  if (root_name) {
    root_name->clear();
  }
  if (!root) {
    return {};
  }
  std::lock_guard<std::mutex> lock(g_roots_mutex);
  const auto entry = g_roots.find(root);
  if (entry == g_roots.end()) {
    return {};
  }
  if (root_name) {
    *root_name = entry->second.root_name;
  }
  return entry->second.data;
}

bool HasSubKeys(const VirtualRegistryData& data, const RegistryNode& node) {
  const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  return key && !key->children.empty();
}

bool QueryKeyInfo(const VirtualRegistryData& data, const RegistryNode& node,
                  KeyInfo* info) {
  if (!info) {
    return false;
  }
  const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  if (!key) {
    return false;
  }
  info->subkey_count = static_cast<DWORD>(key->children.size());
  info->value_count = static_cast<DWORD>(key->values.size());
  info->last_write = {};
  return true;
}

std::vector<std::wstring> EnumSubKeyNames(const VirtualRegistryData& data,
                                          const RegistryNode& node,
                                          bool sorted) {
  const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  if (!key) {
    return {};
  }
  return ChildNames(*key, sorted);
}

bool EnumKeyStreaming(
    const VirtualRegistryData& data, const RegistryNode& node,
    bool include_values, bool include_data, bool include_subkeys,
    RegistryStore::KeyEnumResult* out_info,
    const RegistryStore::ValueStreamCallback& value_callback,
    const RegistryStore::SubkeyStreamCallback& subkey_callback,
    DWORD max_data_size, EnumerationScratch* scratch, bool ordered) {
  (void)scratch;
  const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  if (!key) {
    return false;
  }
  if (out_info) {
    out_info->info.subkey_count = static_cast<DWORD>(key->children.size());
    out_info->info.value_count = static_cast<DWORD>(key->values.size());
    out_info->info.last_write = {};
    out_info->info_valid = true;
  }
  if (include_values && value_callback) {
    std::vector<const RegistryValue*> values;
    values.reserve(key->values.size());
    for (const auto& value : key->values) {
      values.push_back(&value.second);
    }
    if (ordered) {
      std::sort(values.begin(), values.end(),
                [](const RegistryValue* left, const RegistryValue* right) {
                  return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
                });
    }
    for (const RegistryValue* value : values) {
      ValueInfo info;
      info.name = value->name;
      info.type = value->type;
      info.data_size = static_cast<DWORD>(value->data.size());
      const bool data_available =
          include_data && value->data.size() <= max_data_size;
      const BYTE* buffer =
          data_available && !value->data.empty() ? value->data.data() : nullptr;
      if (!value_callback(info, buffer,
                          static_cast<DWORD>(value->data.size()))) {
        return false;
      }
    }
  }
  if (include_subkeys && subkey_callback) {
    for (const auto& name : ChildNames(*key, ordered)) {
      if (!subkey_callback(name)) {
        return false;
      }
    }
  }
  return true;
}

bool QueryValue(const VirtualRegistryData& data, const RegistryNode& node,
                const std::wstring& value_name, ValueEntry* out) {
  if (!out) {
    return false;
  }
  const VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  if (!key) {
    return false;
  }
  const auto value = key->values.find(util::ToLower(value_name));
  if (value == key->values.end()) {
    return false;
  }
  *out = value->second;
  return true;
}

bool CreateKey(VirtualRegistryData& data, const RegistryNode& node,
               const std::wstring& name) {
  VirtualRegistryKey* parent = FindKey(data.root.get(), node.subkey);
  if (!parent) {
    return false;
  }
  const std::wstring lower_name = util::ToLower(name);
  if (parent->children.find(lower_name) != parent->children.end()) {
    return true;
  }
  auto child = std::make_unique<VirtualRegistryKey>();
  child->name = name;
  parent->children.emplace(lower_name, std::move(child));
  return true;
}

bool DeleteKey(VirtualRegistryData& data, const RegistryNode& node) {
  std::wstring parent_path;
  std::wstring name;
  if (!SplitNode(node, &parent_path, &name)) {
    return false;
  }
  VirtualRegistryKey* parent = FindKey(data.root.get(), parent_path);
  return parent &&
         parent->children.erase(util::ToLower(name)) != 0;
}

bool RenameKey(VirtualRegistryData& data, const RegistryNode& node,
               const std::wstring& new_name) {
  std::wstring parent_path;
  std::wstring old_name;
  if (!SplitNode(node, &parent_path, &old_name)) {
    return false;
  }
  VirtualRegistryKey* parent = FindKey(data.root.get(), parent_path);
  if (!parent) {
    return false;
  }
  const std::wstring old_lower = util::ToLower(old_name);
  const std::wstring new_lower = util::ToLower(new_name);
  auto source = parent->children.find(old_lower);
  if (source == parent->children.end() ||
      parent->children.find(new_lower) != parent->children.end()) {
    return false;
  }
  std::unique_ptr<VirtualRegistryKey> moved = std::move(source->second);
  parent->children.erase(source);
  if (moved) {
    moved->name = new_name;
  }
  parent->children.emplace(new_lower, std::move(moved));
  return true;
}

bool DeleteValue(VirtualRegistryData& data, const RegistryNode& node,
                 const std::wstring& value_name) {
  VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  return key && key->values.erase(util::ToLower(value_name)) != 0;
}

bool SetValue(VirtualRegistryData& data, const RegistryNode& node,
              const std::wstring& value_name, DWORD type,
              const std::vector<BYTE>& value_data) {
  VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  if (!key) {
    return false;
  }
  RegistryValue& value = key->values[util::ToLower(value_name)];
  value.name = value_name;
  value.type = type;
  value.data = value_data;
  return true;
}

bool RenameValue(VirtualRegistryData& data, const RegistryNode& node,
                 const std::wstring& old_name,
                 const std::wstring& new_name) {
  VirtualRegistryKey* key = FindKey(data.root.get(), node.subkey);
  if (!key) {
    return false;
  }
  const std::wstring old_lower = util::ToLower(old_name);
  const std::wstring new_lower = util::ToLower(new_name);
  auto source = key->values.find(old_lower);
  if (source == key->values.end() ||
      key->values.find(new_lower) != key->values.end()) {
    return false;
  }
  RegistryValue value = std::move(source->second);
  key->values.erase(source);
  value.name = new_name;
  key->values.emplace(new_lower, std::move(value));
  return true;
}

} // namespace regkit::registry_backend::virtual_store
