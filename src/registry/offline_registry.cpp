// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_backends.h"

#include "registry/registry_path.h"
#include "win32/system_error.h"

#include <algorithm>
#include <utility>
#include <vector>

#include <winternl.h>

namespace regkit::registry_backend::offline {
namespace {

using ORHKEY = void*;
using OROpenHiveFn = DWORD(WINAPI*)(PCWSTR, ORHKEY*);
using ORCloseHiveFn = DWORD(WINAPI*)(ORHKEY);
using ORSaveHiveFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, DWORD, DWORD);
using OROpenKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, ORHKEY*);
using ORCloseKeyFn = DWORD(WINAPI*)(ORHKEY);
using ORCreateKeyFn = DWORD(WINAPI*)(
    ORHKEY, PCWSTR, PWSTR, DWORD, PSECURITY_DESCRIPTOR, ORHKEY*, DWORD*);
using ORDeleteKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);
using ORQueryInfoKeyFn = DWORD(WINAPI*)(
    ORHKEY, PWSTR, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*,
    DWORD*, FILETIME*);
using OREnumKeyFn =
    DWORD(WINAPI*)(ORHKEY, DWORD, PWSTR, DWORD*, PWSTR, DWORD*, FILETIME*);
using ORGetValueFn =
    DWORD(WINAPI*)(ORHKEY, PCWSTR, PCWSTR, DWORD*, void*, DWORD*);
using ORSetValueFn =
    DWORD(WINAPI*)(ORHKEY, PCWSTR, DWORD, const BYTE*, DWORD);
using ORDeleteValueFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);
using OREnumValueFn =
    DWORD(WINAPI*)(ORHKEY, DWORD, PWSTR, DWORD*, DWORD*, BYTE*, DWORD*);
using ORRenameKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);

template <typename Function>
Function LoadFunction(HMODULE module, const char* name) {
  return reinterpret_cast<Function>(GetProcAddress(module, name));
}

class OffregApi {
public:
  OffregApi() {
    module_ = LoadLibraryW(L"offreg.dll");
    if (!module_) {
      return;
    }
    open_hive = LoadFunction<OROpenHiveFn>(module_, "OROpenHive");
    close_hive = LoadFunction<ORCloseHiveFn>(module_, "ORCloseHive");
    save_hive = LoadFunction<ORSaveHiveFn>(module_, "ORSaveHive");
    open_key = LoadFunction<OROpenKeyFn>(module_, "OROpenKey");
    close_key = LoadFunction<ORCloseKeyFn>(module_, "ORCloseKey");
    create_key = LoadFunction<ORCreateKeyFn>(module_, "ORCreateKey");
    delete_key = LoadFunction<ORDeleteKeyFn>(module_, "ORDeleteKey");
    query_info = LoadFunction<ORQueryInfoKeyFn>(module_, "ORQueryInfoKey");
    enum_key = LoadFunction<OREnumKeyFn>(module_, "OREnumKey");
    get_value = LoadFunction<ORGetValueFn>(module_, "ORGetValue");
    set_value = LoadFunction<ORSetValueFn>(module_, "ORSetValue");
    delete_value = LoadFunction<ORDeleteValueFn>(module_, "ORDeleteValue");
    enum_value = LoadFunction<OREnumValueFn>(module_, "OREnumValue");
    rename_key = LoadFunction<ORRenameKeyFn>(module_, "ORRenameKey");
    if (!valid()) {
      FreeLibrary(module_);
      module_ = nullptr;
    }
  }

  ~OffregApi() {
    if (module_) {
      FreeLibrary(module_);
    }
  }

  OffregApi(const OffregApi&) = delete;
  OffregApi& operator=(const OffregApi&) = delete;

  bool valid() const noexcept {
    return module_ && open_hive && close_hive && save_hive && open_key &&
           close_key && create_key && delete_key && query_info && enum_key &&
           get_value && set_value && delete_value && enum_value && rename_key;
  }

  OROpenHiveFn open_hive = nullptr;
  ORCloseHiveFn close_hive = nullptr;
  ORSaveHiveFn save_hive = nullptr;
  OROpenKeyFn open_key = nullptr;
  ORCloseKeyFn close_key = nullptr;
  ORCreateKeyFn create_key = nullptr;
  ORDeleteKeyFn delete_key = nullptr;
  ORQueryInfoKeyFn query_info = nullptr;
  OREnumKeyFn enum_key = nullptr;
  ORGetValueFn get_value = nullptr;
  ORSetValueFn set_value = nullptr;
  ORDeleteValueFn delete_value = nullptr;
  OREnumValueFn enum_value = nullptr;
  ORRenameKeyFn rename_key = nullptr;

private:
  HMODULE module_ = nullptr;
};

OffregApi* Api() {
  static OffregApi api;
  return api.valid() ? &api : nullptr;
}

class OfflineKey {
public:
  OfflineKey() noexcept = default;
  OfflineKey(ORHKEY handle, ORCloseKeyFn close) noexcept
      : handle_(handle), close_(close) {}
  ~OfflineKey() { reset(); }
  OfflineKey(const OfflineKey&) = delete;
  OfflineKey& operator=(const OfflineKey&) = delete;
  OfflineKey(OfflineKey&& other) noexcept
      : handle_(std::exchange(other.handle_, nullptr)),
        close_(std::exchange(other.close_, nullptr)) {}
  OfflineKey& operator=(OfflineKey&& other) noexcept {
    if (this != &other) {
      reset();
      handle_ = std::exchange(other.handle_, nullptr);
      close_ = std::exchange(other.close_, nullptr);
    }
    return *this;
  }

  ORHKEY get() const noexcept { return handle_; }

private:
  void reset() noexcept {
    if (handle_ && close_) {
      close_(handle_);
    }
    handle_ = nullptr;
    close_ = nullptr;
  }

  ORHKEY handle_ = nullptr;
  ORCloseKeyFn close_ = nullptr;
};

std::vector<HKEY> g_roots;

OfflineKey OpenKey(const RegistryNode& node, OffregApi* api) {
  ORHKEY root = reinterpret_cast<ORHKEY>(node.root);
  if (!api || !root) {
    return {};
  }
  if (node.subkey.empty()) {
    return {root, nullptr};
  }
  ORHKEY key = nullptr;
  if (api->open_key(root, node.subkey.c_str(), &key) != ERROR_SUCCESS ||
      !key) {
    return {};
  }
  return {key, api->close_key};
}

bool DeleteSubtree(OffregApi& api, ORHKEY parent,
                   const std::wstring& name) {
  if (!parent || name.empty()) {
    return false;
  }
  ORHKEY child_handle = nullptr;
  if (api.open_key(parent, name.c_str(), &child_handle) != ERROR_SUCCESS ||
      !child_handle) {
    return false;
  }
  OfflineKey child(child_handle, api.close_key);

  DWORD child_count = 0;
  DWORD max_name_length = 0;
  if (api.query_info(child.get(), nullptr, nullptr, &child_count,
                     &max_name_length, nullptr, nullptr, nullptr, nullptr,
                     nullptr, nullptr) == ERROR_SUCCESS &&
      child_count > 0) {
    std::wstring buffer(max_name_length + 1, L'\0');
    std::vector<std::wstring> children;
    children.reserve(child_count);
    for (DWORD index = 0; index < child_count; ++index) {
      DWORD length = static_cast<DWORD>(buffer.size());
      if (api.enum_key(child.get(), index, buffer.data(), &length, nullptr,
                       nullptr, nullptr) == ERROR_SUCCESS) {
        children.emplace_back(buffer.data(), length);
      }
    }
    for (const std::wstring& child_name : children) {
      if (!DeleteSubtree(api, child.get(), child_name)) {
        return false;
      }
    }
  }
  child = {};
  return api.delete_key(parent, name.c_str()) == ERROR_SUCCESS;
}

bool GetOsVersion(DWORD* major, DWORD* minor) {
  const HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
  if (!major || !minor || !ntdll) {
    return false;
  }
  using RtlGetVersionFn = NTSTATUS(WINAPI*)(PRTL_OSVERSIONINFOW);
  const auto get_version = reinterpret_cast<RtlGetVersionFn>(
      GetProcAddress(ntdll, "RtlGetVersion"));
  if (!get_version) {
    return false;
  }
  RTL_OSVERSIONINFOW version = {};
  version.dwOSVersionInfoSize = sizeof(version);
  if (get_version(&version) != 0) {
    return false;
  }
  *major = version.dwMajorVersion;
  *minor = version.dwMinorVersion;
  return true;
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

bool OpenHive(const std::wstring& path, HKEY* root, std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (!root) {
    return false;
  }
  *root = nullptr;
  OffregApi* api = Api();
  if (!api) {
    if (error) {
      *error = L"offreg.dll is not available.";
    }
    return false;
  }
  ORHKEY hive = nullptr;
  const DWORD result = api->open_hive(path.c_str(), &hive);
  if (result != ERROR_SUCCESS || !hive) {
    if (error) {
      *error = util::FormatWin32Error(result);
    }
    return false;
  }
  *root = reinterpret_cast<HKEY>(hive);
  return true;
}

bool SaveHive(HKEY root, const std::wstring& path, std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (!root) {
    return false;
  }
  OffregApi* api = Api();
  if (!api) {
    if (error) {
      *error = L"offreg.dll is not available.";
    }
    return false;
  }
  DWORD major = 10;
  DWORD minor = 0;
  GetOsVersion(&major, &minor);
  const DWORD result =
      api->save_hive(reinterpret_cast<ORHKEY>(root), path.c_str(), major,
                     minor);
  if (result != ERROR_SUCCESS) {
    if (error) {
      *error = util::FormatWin32Error(result);
    }
    return false;
  }
  return true;
}

bool CloseHive(HKEY root, std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (!root) {
    return true;
  }
  OffregApi* api = Api();
  if (!api) {
    if (error) {
      *error = L"offreg.dll is not available.";
    }
    return false;
  }
  const DWORD result = api->close_hive(reinterpret_cast<ORHKEY>(root));
  if (result != ERROR_SUCCESS) {
    if (error) {
      *error = util::FormatWin32Error(result);
    }
    return false;
  }
  return true;
}

void SetRoots(const std::vector<HKEY>& roots) {
  g_roots.clear();
  g_roots.reserve(roots.size());
  for (HKEY root : roots) {
    if (root) {
      g_roots.push_back(root);
    }
  }
}

bool Owns(HKEY root) {
  return root &&
         std::find(g_roots.begin(), g_roots.end(), root) != g_roots.end();
}

bool HasSubKeys(const RegistryNode& node) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  if (!api || !key.get()) {
    return false;
  }
  DWORD count = 0;
  return api->query_info(key.get(), nullptr, nullptr, &count, nullptr,
                         nullptr, nullptr, nullptr, nullptr, nullptr,
                         nullptr) == ERROR_SUCCESS &&
         count > 0;
}

bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  if (!info || !api || !key.get()) {
    return false;
  }
  DWORD subkey_count = 0;
  DWORD value_count = 0;
  FILETIME last_write = {};
  if (api->query_info(key.get(), nullptr, nullptr, &subkey_count, nullptr,
                      nullptr, &value_count, nullptr, nullptr, nullptr,
                      &last_write) != ERROR_SUCCESS) {
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
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  if (!api || !key.get()) {
    return false;
  }
  DWORD type = 0;
  DWORD size = 0;
  DWORD result = api->get_value(key.get(), nullptr, L"SymbolicLinkValue",
                                &type, nullptr, &size);
  if ((result != ERROR_SUCCESS && result != ERROR_MORE_DATA) ||
      (type != REG_LINK && type != REG_SZ && type != REG_EXPAND_SZ) ||
      size == 0) {
    return false;
  }
  std::vector<wchar_t> buffer(size / sizeof(wchar_t) + 1, L'\0');
  result = api->get_value(key.get(), nullptr, L"SymbolicLinkValue", &type,
                          buffer.data(), &size);
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
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  if (!api || !key.get()) {
    return names;
  }
  DWORD count = 0;
  DWORD max_name_length = 0;
  if (api->query_info(key.get(), nullptr, nullptr, &count, &max_name_length,
                      nullptr, nullptr, nullptr, nullptr, nullptr,
                      nullptr) != ERROR_SUCCESS) {
    return names;
  }
  names.reserve(count);
  std::wstring buffer(max_name_length + 1, L'\0');
  for (DWORD index = 0; index < count; ++index) {
    DWORD length = static_cast<DWORD>(buffer.size());
    if (api->enum_key(key.get(), index, buffer.data(), &length, nullptr,
                      nullptr, nullptr) == ERROR_SUCCESS) {
      names.emplace_back(buffer.data(), length);
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
    DWORD max_data_size) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  if (!api || !key.get()) {
    return false;
  }
  DWORD subkey_count = 0;
  DWORD max_subkey_length = 0;
  DWORD value_count = 0;
  DWORD max_value_name_length = 0;
  DWORD max_value_data_length = 0;
  FILETIME last_write = {};
  if (api->query_info(
          key.get(), nullptr, nullptr, &subkey_count, &max_subkey_length,
          nullptr, &value_count, &max_value_name_length,
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
    std::wstring name(max_value_name_length + 1, L'\0');
    std::vector<BYTE> data;
    if (include_data) {
      data.resize(std::min(max_value_data_length, max_data_size));
    }
    for (DWORD index = 0; index < value_count; ++index) {
      DWORD name_length = static_cast<DWORD>(name.size());
      DWORD data_length =
          include_data ? static_cast<DWORD>(data.size()) : 0;
      DWORD type = 0;
      DWORD result = api->enum_value(
          key.get(), index, name.data(), &name_length, &type,
          include_data && !data.empty() ? data.data() : nullptr,
          &data_length);
      if (result == ERROR_MORE_DATA && include_data &&
          data_length <= max_data_size) {
        data.resize(data_length);
        name_length = static_cast<DWORD>(name.size());
        data_length = static_cast<DWORD>(data.size());
        result = api->enum_value(
            key.get(), index, name.data(), &name_length, &type,
            data.empty() ? nullptr : data.data(), &data_length);
      }
      const bool data_available =
          include_data && result == ERROR_SUCCESS;
      if (result != ERROR_SUCCESS &&
          !(result == ERROR_MORE_DATA &&
            (!include_data || data_length > max_data_size))) {
        continue;
      }
      ValueInfo info;
      info.name.assign(name.data(), name_length);
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
    std::wstring name(max_subkey_length + 1, L'\0');
    for (DWORD index = 0; index < subkey_count; ++index) {
      DWORD name_length = static_cast<DWORD>(name.size());
      if (api->enum_key(key.get(), index, name.data(), &name_length, nullptr,
                        nullptr, nullptr) != ERROR_SUCCESS) {
        continue;
      }
      if (!subkey_callback(std::wstring(name.data(), name_length))) {
        return false;
      }
    }
  }
  return true;
}

bool QueryValue(const RegistryNode& node, const std::wstring& value_name,
                ValueEntry* out) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  if (!out || !api || !key.get()) {
    return false;
  }
  const wchar_t* name =
      value_name.empty() ? nullptr : value_name.c_str();
  DWORD type = 0;
  DWORD size = 0;
  DWORD result =
      api->get_value(key.get(), nullptr, name, &type, nullptr, &size);
  if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
    return false;
  }
  std::vector<BYTE> data(size);
  result = api->get_value(key.get(), nullptr, name, &type,
                          data.empty() ? nullptr : data.data(), &size);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  data.resize(size);
  out->name = value_name;
  out->type = type;
  out->data = std::move(data);
  return true;
}

bool CreateKey(const RegistryNode& node, const std::wstring& name) {
  OffregApi* api = Api();
  OfflineKey parent = OpenKey(node, api);
  if (!api || !parent.get()) {
    return false;
  }
  ORHKEY created_handle = nullptr;
  DWORD disposition = 0;
  const DWORD result =
      api->create_key(parent.get(), name.c_str(), nullptr, 0, nullptr,
                      &created_handle, &disposition);
  OfflineKey created(created_handle, api->close_key);
  return result == ERROR_SUCCESS;
}

bool DeleteKey(const RegistryNode& node) {
  RegistryNode parent;
  std::wstring name;
  if (!SplitNode(node, &parent, &name)) {
    return false;
  }
  OffregApi* api = Api();
  OfflineKey key = OpenKey(parent, api);
  return api && key.get() && DeleteSubtree(*api, key.get(), name);
}

bool RenameKey(const RegistryNode& node, const std::wstring& new_name) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  return api && key.get() &&
         api->rename_key(key.get(), new_name.c_str()) == ERROR_SUCCESS;
}

bool DeleteValue(const RegistryNode& node,
                 const std::wstring& value_name) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  return api && key.get() &&
         api->delete_value(key.get(), value_name.empty()
                                          ? nullptr
                                          : value_name.c_str()) ==
             ERROR_SUCCESS;
}

bool SetValue(const RegistryNode& node, const std::wstring& value_name,
              DWORD type, const std::vector<BYTE>& data) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  return api && key.get() &&
         api->set_value(key.get(),
                        value_name.empty() ? nullptr : value_name.c_str(),
                        type, data.empty() ? nullptr : data.data(),
                        static_cast<DWORD>(data.size())) == ERROR_SUCCESS;
}

bool RenameValue(const RegistryNode& node, const std::wstring& old_name,
                 const std::wstring& new_name) {
  OffregApi* api = Api();
  OfflineKey key = OpenKey(node, api);
  if (!api || !key.get()) {
    return false;
  }
  DWORD type = 0;
  DWORD size = 0;
  DWORD result = api->get_value(key.get(), nullptr, old_name.c_str(), &type,
                                nullptr, &size);
  if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
    return false;
  }
  std::vector<BYTE> data(size);
  result = api->get_value(key.get(), nullptr, old_name.c_str(), &type,
                          data.empty() ? nullptr : data.data(), &size);
  if (result != ERROR_SUCCESS ||
      api->set_value(key.get(), new_name.c_str(), type,
                     data.empty() ? nullptr : data.data(),
                     size) != ERROR_SUCCESS) {
    return false;
  }
  return api->delete_value(key.get(), old_name.c_str()) == ERROR_SUCCESS;
}

} // namespace regkit::registry_backend::offline
