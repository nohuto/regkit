// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_backends.h"

#include "registry/registry_path.h"
#include "win32/registry_native.h"
#include "win32/registry_view.h"

#include <algorithm>
#include <utility>
#include <vector>

namespace regkit::registry_backend::live {
namespace {

util::UniqueHKey OpenKey(const RegistryNode& node, REGSAM access) {
  util::UniqueHKey key;
  if (!node.root) {
    return key;
  }
  const wchar_t* subkey =
      node.subkey.empty() ? nullptr : node.subkey.c_str();
  if (RegOpenKeyExW(node.root, subkey, 0, access | win32::kDefaultRegistryView, key.put()) !=
      ERROR_SUCCESS) {
    key.reset();
  }
  return key;
}

bool SplitNode(const RegistryNode& node, RegistryNode* parent,
               std::wstring* name) {
  if (!parent || !name) {
    return false;
  }
  *name = registry_path::Leaf(node.subkey);
  if (name->empty()) {
    return false;
  }
  *parent = node;
  parent->subkey = registry_path::Parent(node.subkey);
  return true;
}

} // namespace

bool HasSubKeys(const RegistryNode& node) {
  util::UniqueHKey key = OpenKey(node, KEY_ENUMERATE_SUB_KEYS);
  if (!key.get()) {
    return false;
  }
  wchar_t name[1] = {};
  DWORD name_length = static_cast<DWORD>(_countof(name));
  const LONG result =
      RegEnumKeyExW(key.get(), 0, name, &name_length, nullptr, nullptr,
                    nullptr, nullptr);
  return result == ERROR_SUCCESS || result == ERROR_MORE_DATA;
}

bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info) {
  if (!info) {
    return false;
  }
  util::UniqueHKey key = OpenKey(node, KEY_READ);
  if (!key.get()) {
    return false;
  }
  DWORD subkey_count = 0;
  DWORD value_count = 0;
  FILETIME last_write = {};
  if (RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, &subkey_count,
                       nullptr, nullptr, &value_count, nullptr, nullptr,
                       nullptr, &last_write) != ERROR_SUCCESS) {
    return false;
  }
  info->subkey_count = subkey_count;
  info->value_count = value_count;
  info->last_write = last_write;
  return true;
}

bool QuerySymbolicLinkTarget(const RegistryNode& node,
                             std::wstring* target) {
  if (!target) {
    return false;
  }
  target->clear();
  const std::wstring native_path = registry_path::BuildNative(node);
  util::UniqueHKey key =
      util::OpenNativeRegistryKey(native_path, KEY_QUERY_VALUE, true);
  if (!key.get()) {
    return false;
  }
  DWORD type = 0;
  DWORD size = 0;
  LONG result =
      RegQueryValueExW(key.get(), L"SymbolicLinkValue", nullptr, &type,
                       nullptr, &size);
  if (result != ERROR_SUCCESS ||
      (type != REG_LINK && type != REG_SZ && type != REG_EXPAND_SZ) ||
      size == 0) {
    return false;
  }
  std::vector<wchar_t> buffer(size / sizeof(wchar_t) + 1, L'\0');
  result = RegQueryValueExW(
      key.get(), L"SymbolicLinkValue", nullptr, &type,
      reinterpret_cast<LPBYTE>(buffer.data()), &size);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  std::wstring value(buffer.data(), size / sizeof(wchar_t));
  while (!value.empty() && value.back() == L'\0') {
    value.pop_back();
  }
  if (value.empty()) {
    return false;
  }
  *target = std::move(value);
  return true;
}

std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node,
                                          bool sorted) {
  std::vector<std::wstring> names;
  util::UniqueHKey key = OpenKey(node, KEY_READ);
  if (!key.get()) {
    return names;
  }
  DWORD subkey_count = 0;
  DWORD max_name_length = 0;
  if (RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, &subkey_count,
                       &max_name_length, nullptr, nullptr, nullptr, nullptr,
                       nullptr, nullptr) != ERROR_SUCCESS) {
    return names;
  }
  names.reserve(subkey_count);
  std::wstring buffer(max_name_length + 1, L'\0');
  for (DWORD index = 0; index < subkey_count; ++index) {
    DWORD name_length = static_cast<DWORD>(buffer.size());
    FILETIME last_write = {};
    if (RegEnumKeyExW(key.get(), index, buffer.data(), &name_length, nullptr,
                      nullptr, nullptr, &last_write) == ERROR_SUCCESS) {
      buffer[name_length] = L'\0';
      names.emplace_back(buffer.c_str());
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

bool EnumKeyStreaming(
    const RegistryNode& node, bool include_values, bool include_data,
    bool include_subkeys, RegistryStore::KeyEnumResult* out_info,
    const RegistryStore::ValueStreamCallback& value_callback,
    const RegistryStore::SubkeyStreamCallback& subkey_callback,
    DWORD max_data_size, EnumerationScratch* scratch, bool ordered) {
  (void)ordered;
  EnumerationScratch local;
  EnumerationScratch& buffers = scratch ? *scratch : local;
  util::UniqueHKey key = OpenKey(node, KEY_READ);
  if (!key.get()) {
    return false;
  }

  DWORD subkey_count = 0;
  DWORD max_subkey_length = 0;
  DWORD value_count = 0;
  DWORD max_value_name_length = 0;
  DWORD max_value_data_length = 0;
  FILETIME last_write = {};
  if (RegQueryInfoKeyW(
          key.get(), nullptr, nullptr, nullptr, &subkey_count,
          &max_subkey_length, nullptr, &value_count, &max_value_name_length,
          &max_value_data_length, nullptr, &last_write) != ERROR_SUCCESS) {
    return false;
  }
  if (out_info) {
    out_info->info.subkey_count = subkey_count;
    out_info->info.value_count = value_count;
    out_info->info.last_write = last_write;
    out_info->info_valid = true;
  }

  if (include_values && value_callback) {
    std::wstring& name = buffers.value_name;
    std::vector<BYTE>& data = buffers.value_data;
    if (name.size() < static_cast<size_t>(max_value_name_length) + 1) {
      name.resize(static_cast<size_t>(max_value_name_length) + 1);
    }
    if (include_data) {
      const size_t needed = std::min(max_value_data_length, max_data_size);
      if (data.size() < needed) {
        data.resize(needed);
      }
    }
    for (DWORD index = 0; index < value_count; ++index) {
      DWORD name_length = static_cast<DWORD>(name.size());
      DWORD data_length =
          include_data ? static_cast<DWORD>(data.size()) : 0;
      DWORD type = 0;
      LONG result = RegEnumValueW(
          key.get(), index, name.data(), &name_length, nullptr, &type,
          include_data && !data.empty() ? data.data() : nullptr,
          &data_length);
      if (result == ERROR_MORE_DATA && include_data &&
          data_length <= max_data_size) {
        if (data.size() < data_length) {
          data.resize(data_length);
        }
        name_length = static_cast<DWORD>(name.size());
        data_length = static_cast<DWORD>(data.size());
        result = RegEnumValueW(
            key.get(), index, name.data(), &name_length, nullptr, &type,
            data.empty() ? nullptr : data.data(), &data_length);
      }
      const bool data_available =
          include_data && result == ERROR_SUCCESS;
      if (result != ERROR_SUCCESS &&
          !(result == ERROR_MORE_DATA &&
            (!include_data || data_length > max_data_size))) {
        continue;
      }
      name[name_length] = L'\0';
      ValueInfo info;
      info.name.assign(name.c_str(), name_length);
      info.type = type;
      info.data_size = data_length;
      const BYTE* buffer =
          data_available && data_length > 0 ? data.data() : nullptr;
      if (!value_callback(info, buffer, data_length)) {
        return false;
      }
    }
  }

  if (include_subkeys && subkey_callback) {
    std::wstring& name = buffers.subkey_name;
    if (name.size() < static_cast<size_t>(max_subkey_length) + 1) {
      name.resize(static_cast<size_t>(max_subkey_length) + 1);
    }
    for (DWORD index = 0; index < subkey_count; ++index) {
      DWORD name_length = static_cast<DWORD>(name.size());
      FILETIME child_write = {};
      if (RegEnumKeyExW(key.get(), index, name.data(), &name_length, nullptr,
                        nullptr, nullptr, &child_write) != ERROR_SUCCESS) {
        continue;
      }
      name[name_length] = L'\0';
      if (!subkey_callback(name.c_str())) {
        return false;
      }
    }
  }
  return true;
}

bool QueryValue(const RegistryNode& node, const std::wstring& value_name,
                ValueEntry* out) {
  if (!out) {
    return false;
  }
  util::UniqueHKey key = OpenKey(node, KEY_QUERY_VALUE);
  if (!key.get()) {
    return false;
  }
  const wchar_t* name =
      value_name.empty() ? nullptr : value_name.c_str();
  DWORD type = 0;
  DWORD size = 0;
  if (RegQueryValueExW(key.get(), name, nullptr, &type, nullptr, &size) !=
      ERROR_SUCCESS) {
    return false;
  }
  std::vector<BYTE> data(size);
  if (RegQueryValueExW(key.get(), name, nullptr, &type,
                       data.empty() ? nullptr : data.data(),
                       &size) != ERROR_SUCCESS) {
    return false;
  }
  data.resize(size);
  out->name = value_name;
  out->type = type;
  out->data = std::move(data);
  return true;
}

bool CreateKey(const RegistryNode& node, const std::wstring& name) {
  util::UniqueHKey parent = OpenKey(node, KEY_WRITE);
  if (!parent.get()) {
    return false;
  }
  util::UniqueHKey created;
  DWORD disposition = 0;
  return RegCreateKeyExW(parent.get(), name.c_str(), 0, nullptr,
                         REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE,
                         nullptr, created.put(), &disposition) ==
         ERROR_SUCCESS;
}

bool DeleteKey(const RegistryNode& node) {
  RegistryNode parent;
  std::wstring name;
  if (!SplitNode(node, &parent, &name)) {
    return false;
  }
  util::UniqueHKey key = OpenKey(parent, KEY_WRITE);
  return key.get() &&
         RegDeleteTreeW(key.get(), name.c_str()) == ERROR_SUCCESS;
}

bool RenameKey(const RegistryNode& node, const std::wstring& new_name) {
  RegistryNode parent;
  std::wstring old_name;
  if (!SplitNode(node, &parent, &old_name)) {
    return false;
  }
  util::UniqueHKey key = OpenKey(parent, KEY_WRITE);
  return key.get() &&
         RegRenameKey(key.get(), old_name.c_str(), new_name.c_str()) ==
             ERROR_SUCCESS;
}

bool DeleteValue(const RegistryNode& node,
                 const std::wstring& value_name) {
  util::UniqueHKey key = OpenKey(node, KEY_SET_VALUE);
  return key.get() &&
         RegDeleteValueW(key.get(), value_name.c_str()) == ERROR_SUCCESS;
}

bool SetValue(const RegistryNode& node, const std::wstring& value_name,
              DWORD type, const std::vector<BYTE>& data) {
  util::UniqueHKey key = OpenKey(node, KEY_SET_VALUE);
  if (!key.get()) {
    return false;
  }
  return RegSetValueExW(
             key.get(), value_name.c_str(), 0, type,
             data.empty() ? nullptr : data.data(),
             static_cast<DWORD>(data.size())) == ERROR_SUCCESS;
}

bool RenameValue(const RegistryNode& node, const std::wstring& old_name,
                 const std::wstring& new_name) {
  util::UniqueHKey key =
      OpenKey(node, KEY_QUERY_VALUE | KEY_SET_VALUE);
  if (!key.get()) {
    return false;
  }
  DWORD type = 0;
  DWORD size = 0;
  if (RegQueryValueExW(key.get(), old_name.c_str(), nullptr, &type, nullptr,
                       &size) != ERROR_SUCCESS) {
    return false;
  }
  std::vector<BYTE> data(size);
  if (RegQueryValueExW(key.get(), old_name.c_str(), nullptr, &type,
                       data.empty() ? nullptr : data.data(),
                       &size) != ERROR_SUCCESS ||
      RegSetValueExW(key.get(), new_name.c_str(), 0, type,
                     data.empty() ? nullptr : data.data(),
                     size) != ERROR_SUCCESS) {
    return false;
  }
  return RegDeleteValueW(key.get(), old_name.c_str()) == ERROR_SUCCESS;
}

} // namespace regkit::registry_backend::live
