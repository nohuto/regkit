// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#include "registry/registry_provider.h"

#include <algorithm>
#include <cwctype>
#include <memory>
#include <mutex>
#include <shlwapi.h>
#include <unordered_map>
#include <vector>

#include <winternl.h>

#include "win32/win32_helpers.h"

namespace regkit {

namespace {

#ifndef OBJ_OPENLINK
#define OBJ_OPENLINK 0x00000008L
#endif

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif

using NtOpenKeyFn = NTSTATUS(NTAPI*)(PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES);

using ORHKEY = void*;
using OROpenHiveFn = DWORD(WINAPI*)(PCWSTR, ORHKEY*);
using ORCloseHiveFn = DWORD(WINAPI*)(ORHKEY);
using ORSaveHiveFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, DWORD, DWORD);
using OROpenKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, ORHKEY*);
using ORCloseKeyFn = DWORD(WINAPI*)(ORHKEY);
using ORCreateKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, PWSTR, DWORD, PSECURITY_DESCRIPTOR, ORHKEY*, DWORD*);
using ORDeleteKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);
using ORQueryInfoKeyFn = DWORD(WINAPI*)(ORHKEY, PWSTR, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, FILETIME*);
using OREnumKeyFn = DWORD(WINAPI*)(ORHKEY, DWORD, PWSTR, DWORD*, PWSTR, DWORD*, FILETIME*);
using ORGetValueFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, PCWSTR, DWORD*, void*, DWORD*);
using ORSetValueFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, DWORD, const BYTE*, DWORD);
using ORDeleteValueFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);
using OREnumValueFn = DWORD(WINAPI*)(ORHKEY, DWORD, PWSTR, DWORD*, DWORD*, BYTE*, DWORD*);
using ORRenameKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);

struct OffregApi {
  HMODULE module = nullptr;
  bool loaded = false;
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

  bool EnsureLoaded() {
    if (loaded) {
      return module != nullptr;
    }
    loaded = true;
    module = LoadLibraryW(L"offreg.dll");
    if (!module) {
      return false;
    }
    open_hive = reinterpret_cast<OROpenHiveFn>(GetProcAddress(module, "OROpenHive"));
    close_hive = reinterpret_cast<ORCloseHiveFn>(GetProcAddress(module, "ORCloseHive"));
    save_hive = reinterpret_cast<ORSaveHiveFn>(GetProcAddress(module, "ORSaveHive"));
    open_key = reinterpret_cast<OROpenKeyFn>(GetProcAddress(module, "OROpenKey"));
    close_key = reinterpret_cast<ORCloseKeyFn>(GetProcAddress(module, "ORCloseKey"));
    create_key = reinterpret_cast<ORCreateKeyFn>(GetProcAddress(module, "ORCreateKey"));
    delete_key = reinterpret_cast<ORDeleteKeyFn>(GetProcAddress(module, "ORDeleteKey"));
    query_info = reinterpret_cast<ORQueryInfoKeyFn>(GetProcAddress(module, "ORQueryInfoKey"));
    enum_key = reinterpret_cast<OREnumKeyFn>(GetProcAddress(module, "OREnumKey"));
    get_value = reinterpret_cast<ORGetValueFn>(GetProcAddress(module, "ORGetValue"));
    set_value = reinterpret_cast<ORSetValueFn>(GetProcAddress(module, "ORSetValue"));
    delete_value = reinterpret_cast<ORDeleteValueFn>(GetProcAddress(module, "ORDeleteValue"));
    enum_value = reinterpret_cast<OREnumValueFn>(GetProcAddress(module, "OREnumValue"));
    rename_key = reinterpret_cast<ORRenameKeyFn>(GetProcAddress(module, "ORRenameKey"));
    if (!open_hive || !close_hive || !save_hive || !open_key || !close_key || !create_key || !delete_key || !query_info || !enum_key || !get_value || !set_value || !delete_value || !enum_value || !rename_key) {
      FreeLibrary(module);
      module = nullptr;
      return false;
    }
    return true;
  }
};

OffregApi* GetOffreg() {
  static OffregApi api;
  return api.EnsureLoaded() ? &api : nullptr;
}

std::vector<HKEY> g_offline_roots;

struct VirtualRootEntry {
  std::wstring root_name;
  std::shared_ptr<RegistryProvider::VirtualRegistryData> data;
  std::unique_ptr<int> handle_tag;
};

std::mutex g_virtual_mutex;
std::unordered_map<HKEY, VirtualRootEntry> g_virtual_roots;

NtOpenKeyFn LoadNtOpenKey() {
  HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
  if (!ntdll) {
    return nullptr;
  }
  return reinterpret_cast<NtOpenKeyFn>(GetProcAddress(ntdll, "NtOpenKey"));
}

NtOpenKeyFn GetNtOpenKey() {
  static NtOpenKeyFn open_fn = LoadNtOpenKey();
  return open_fn;
}

bool IsOfflineNode(const RegistryNode& node) {
  if (!node.root || g_offline_roots.empty()) {
    return false;
  }
  return std::find(g_offline_roots.begin(), g_offline_roots.end(), node.root) != g_offline_roots.end();
}

struct OfflineKey {
  ORHKEY handle = nullptr;
  bool close = false;
};

OfflineKey OpenOfflineKey(const RegistryNode& node) {
  OfflineKey result;
  OffregApi* api = GetOffreg();
  if (!api) {
    return result;
  }
  ORHKEY root = reinterpret_cast<ORHKEY>(node.root);
  if (!root) {
    return result;
  }
  if (node.subkey.empty()) {
    result.handle = root;
    result.close = false;
    return result;
  }
  ORHKEY key = nullptr;
  if (api->open_key(root, node.subkey.c_str(), &key) != ERROR_SUCCESS || !key) {
    return result;
  }
  result.handle = key;
  result.close = true;
  return result;
}

void CloseOfflineKey(const OfflineKey& key) {
  if (key.close && key.handle) {
    if (OffregApi* api = GetOffreg()) {
      api->close_key(key.handle);
    }
  }
}

bool DeleteOfflineSubtree(ORHKEY parent, const std::wstring& subkey) {
  OffregApi* api = GetOffreg();
  if (!api || !parent || subkey.empty()) {
    return false;
  }
  ORHKEY child = nullptr;
  DWORD result = api->open_key(parent, subkey.c_str(), &child);
  if (result != ERROR_SUCCESS || !child) {
    return false;
  }

  DWORD subkey_count = 0;
  DWORD max_subkey_len = 0;
  result = api->query_info(child, nullptr, nullptr, &subkey_count, &max_subkey_len, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr);
  if (result == ERROR_SUCCESS && subkey_count > 0) {
    std::wstring buffer;
    buffer.resize(max_subkey_len + 1);
    std::vector<std::wstring> children;
    children.reserve(subkey_count);
    for (DWORD index = 0; index < subkey_count; ++index) {
      DWORD name_len = static_cast<DWORD>(buffer.size());
      if (api->enum_key(child, index, buffer.data(), &name_len, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
        buffer[name_len] = L'\0';
        children.emplace_back(buffer.c_str());
      }
    }
    for (const auto& name : children) {
      if (!DeleteOfflineSubtree(child, name)) {
        api->close_key(child);
        return false;
      }
    }
  }

  api->close_key(child);
  return api->delete_key(parent, subkey.c_str()) == ERROR_SUCCESS;
}

bool GetOsVersion(DWORD* major, DWORD* minor) {
  if (!major || !minor) {
    return false;
  }
  HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
  if (!ntdll) {
    return false;
  }
  using RtlGetVersionFn = NTSTATUS(WINAPI*)(PRTL_OSVERSIONINFOW);
  auto fn = reinterpret_cast<RtlGetVersionFn>(GetProcAddress(ntdll, "RtlGetVersion"));
  if (!fn) {
    return false;
  }
  RTL_OSVERSIONINFOW info = {};
  info.dwOSVersionInfoSize = sizeof(info);
  if (fn(&info) != 0) {
    return false;
  }
  *major = info.dwMajorVersion;
  *minor = info.dwMinorVersion;
  return true;
}

std::wstring FormatWin32Error(DWORD code) {
  if (code == 0) {
    return L"";
  }
  wchar_t buffer[512] = {};
  DWORD len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, code, 0, buffer, static_cast<DWORD>(_countof(buffer)), nullptr);
  if (len == 0) {
    return L"Unknown error.";
  }
  return buffer;
}

bool SplitSubKey(const std::wstring& subkey, std::wstring* parent, std::wstring* name) {
  if (!parent || !name) {
    return false;
  }
  size_t pos = subkey.rfind(L'\\');
  if (pos == std::wstring::npos) {
    *parent = L"";
    *name = subkey;
    return !name->empty();
  }
  *parent = subkey.substr(0, pos);
  *name = subkey.substr(pos + 1);
  return !name->empty();
}

std::wstring ToLower(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (wchar_t ch : text) {
    out.push_back(static_cast<wchar_t>(towlower(ch)));
  }
  return out;
}

std::vector<std::wstring> SplitPath(const std::wstring& path) {
  std::vector<std::wstring> parts;
  std::wstring current;
  for (wchar_t ch : path) {
    if (ch == L'\\' || ch == L'/') {
      if (!current.empty()) {
        parts.push_back(current);
        current.clear();
      }
    } else {
      current.push_back(ch);
    }
  }
  if (!current.empty()) {
    parts.push_back(current);
  }
  return parts;
}

bool GetVirtualRootData(HKEY root, std::shared_ptr<RegistryProvider::VirtualRegistryData>* data, std::wstring* root_name) {
  if (data) {
    data->reset();
  }
  if (root_name) {
    root_name->clear();
  }
  if (!root) {
    return false;
  }
  std::lock_guard<std::mutex> lock(g_virtual_mutex);
  auto it = g_virtual_roots.find(root);
  if (it == g_virtual_roots.end()) {
    return false;
  }
  if (data) {
    *data = it->second.data;
  }
  if (root_name) {
    *root_name = it->second.root_name;
  }
  return true;
}

RegistryProvider::VirtualRegistryKey* FindVirtualKey(RegistryProvider::VirtualRegistryKey* root, const std::wstring& subkey) {
  if (!root) {
    return nullptr;
  }
  if (subkey.empty()) {
    return root;
  }
  auto parts = SplitPath(subkey);
  RegistryProvider::VirtualRegistryKey* current = root;
  for (const auto& part : parts) {
    std::wstring lower = ToLower(part);
    auto it = current->children.find(lower);
    if (it == current->children.end()) {
      return nullptr;
    }
    current = it->second.get();
  }
  return current;
}

const RegistryProvider::VirtualRegistryKey* FindVirtualKey(const RegistryProvider::VirtualRegistryKey* root, const std::wstring& subkey) {
  if (!root) {
    return nullptr;
  }
  if (subkey.empty()) {
    return root;
  }
  auto parts = SplitPath(subkey);
  const RegistryProvider::VirtualRegistryKey* current = root;
  for (const auto& part : parts) {
    std::wstring lower = ToLower(part);
    auto it = current->children.find(lower);
    if (it == current->children.end()) {
      return nullptr;
    }
    current = it->second.get();
  }
  return current;
}

} // namespace

std::vector<RegistryRootEntry> RegistryProvider::DefaultRoots(bool include_extra) {
  std::vector<RegistryRootEntry> roots = {
      {HKEY_CLASSES_ROOT, L"HKEY_CLASSES_ROOT", L"HKEY_CLASSES_ROOT", L""}, {HKEY_CURRENT_USER, L"HKEY_CURRENT_USER", L"HKEY_CURRENT_USER", L""}, {HKEY_LOCAL_MACHINE, L"HKEY_LOCAL_MACHINE", L"HKEY_LOCAL_MACHINE", L""}, {HKEY_USERS, L"HKEY_USERS", L"HKEY_USERS", L""}, {HKEY_CURRENT_CONFIG, L"HKEY_CURRENT_CONFIG", L"HKEY_CURRENT_CONFIG", L""},
  };
  if (include_extra) {
    roots.push_back({HKEY_PERFORMANCE_DATA, L"HKEY_PERFORMANCE_DATA", L"HKEY_PERFORMANCE_DATA", L""});
    roots.push_back({HKEY_PERFORMANCE_TEXT, L"HKEY_PERFORMANCE_TEXT", L"HKEY_PERFORMANCE_TEXT", L""});
    roots.push_back({HKEY_PERFORMANCE_NLSTEXT, L"HKEY_PERFORMANCE_NLSTEXT", L"HKEY_PERFORMANCE_NLSTEXT", L""});
  }
  return roots;
}

std::wstring RegistryProvider::RootName(HKEY root) {
  std::wstring virtual_name;
  if (GetVirtualRootData(root, nullptr, &virtual_name)) {
    return virtual_name;
  }
  if (root == HKEY_CLASSES_ROOT) {
    return L"HKEY_CLASSES_ROOT";
  }
  if (root == HKEY_CURRENT_USER) {
    return L"HKEY_CURRENT_USER";
  }
  if (root == HKEY_LOCAL_MACHINE) {
    return L"HKEY_LOCAL_MACHINE";
  }
  if (root == HKEY_USERS) {
    return L"HKEY_USERS";
  }
  if (root == HKEY_CURRENT_CONFIG) {
    return L"HKEY_CURRENT_CONFIG";
  }
  if (root == HKEY_PERFORMANCE_DATA) {
    return L"HKEY_PERFORMANCE_DATA";
  }
  if (root == HKEY_PERFORMANCE_TEXT) {
    return L"HKEY_PERFORMANCE_TEXT";
  }
  if (root == HKEY_PERFORMANCE_NLSTEXT) {
    return L"HKEY_PERFORMANCE_NLSTEXT";
  }
  return L"";
}

std::wstring RegistryProvider::BuildPath(const RegistryNode& node) {
  std::wstring root = node.root_name.empty() ? RootName(node.root) : node.root_name;
  if (node.subkey.empty()) {
    return root;
  }
  return root + L"\\" + node.subkey;
}

std::wstring RegistryProvider::BuildNtPath(const RegistryNode& node) {
  if (IsVirtualRoot(node.root)) {
    return L"";
  }
  if (!node.root_name.empty() && _wcsicmp(node.root_name.c_str(), L"REGISTRY") == 0) {
    std::wstring root_path = L"\\REGISTRY";
    if (!node.subkey.empty()) {
      root_path += L"\\";
      root_path += node.subkey;
    }
    return root_path;
  }

  std::wstring root_path;
  if (node.root == HKEY_LOCAL_MACHINE) {
    root_path = L"\\REGISTRY\\MACHINE";
  } else if (node.root == HKEY_USERS) {
    root_path = L"\\REGISTRY\\USER";
  } else if (node.root == HKEY_CURRENT_USER) {
    std::wstring sid = util::GetCurrentUserSidString();
    if (sid.empty()) {
      return L"";
    }
    root_path = L"\\REGISTRY\\USER\\" + sid;
  } else if (node.root == HKEY_CURRENT_CONFIG) {
    root_path = L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Hardware\\Profiles\\Current";
  } else if (node.root == HKEY_CLASSES_ROOT) {
    root_path = L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes";
  } else {
    return L"";
  }

  if (!node.subkey.empty()) {
    root_path += L"\\";
    root_path += node.subkey;
  }
  return root_path;
}

bool RegistryProvider::OpenOfflineHive(const std::wstring& path, HKEY* root, std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (!root) {
    return false;
  }
  *root = nullptr;
  OffregApi* api = GetOffreg();
  if (!api) {
    if (error) {
      *error = L"offreg.dll is not available.";
    }
    return false;
  }
  ORHKEY hive = nullptr;
  DWORD result = api->open_hive(path.c_str(), &hive);
  if (result != ERROR_SUCCESS || !hive) {
    if (error) {
      *error = FormatWin32Error(result);
    }
    return false;
  }
  *root = reinterpret_cast<HKEY>(hive);
  return true;
}

bool RegistryProvider::SaveOfflineHive(HKEY root, const std::wstring& path, std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (!root) {
    return false;
  }
  OffregApi* api = GetOffreg();
  if (!api) {
    if (error) {
      *error = L"offreg.dll is not available.";
    }
    return false;
  }
  DWORD major = 0;
  DWORD minor = 0;
  if (!GetOsVersion(&major, &minor)) {
    major = 10;
    minor = 0;
  }
  DWORD result = api->save_hive(reinterpret_cast<ORHKEY>(root), path.c_str(), major, minor);
  if (result != ERROR_SUCCESS) {
    if (error) {
      *error = FormatWin32Error(result);
    }
    return false;
  }
  return true;
}

bool RegistryProvider::CloseOfflineHive(HKEY root, std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (!root) {
    return true;
  }
  OffregApi* api = GetOffreg();
  if (!api) {
    if (error) {
      *error = L"offreg.dll is not available.";
    }
    return false;
  }
  DWORD result = api->close_hive(reinterpret_cast<ORHKEY>(root));
  if (result != ERROR_SUCCESS) {
    if (error) {
      *error = FormatWin32Error(result);
    }
    return false;
  }
  return true;
}

void RegistryProvider::SetOfflineRoot(HKEY root) {
  g_offline_roots.clear();
  if (root) {
    g_offline_roots.push_back(root);
  }
}

void RegistryProvider::SetOfflineRoots(const std::vector<HKEY>& roots) {
  g_offline_roots.clear();
  g_offline_roots.reserve(roots.size());
  for (HKEY root : roots) {
    if (root) {
      g_offline_roots.push_back(root);
    }
  }
}

HKEY RegistryProvider::RegisterVirtualRoot(const std::wstring& root_name, const std::shared_ptr<VirtualRegistryData>& data) {
  if (!data) {
    return nullptr;
  }
  auto handle_tag = std::make_unique<int>(0);
  HKEY handle = reinterpret_cast<HKEY>(handle_tag.get());
  VirtualRootEntry entry;
  entry.root_name = root_name;
  entry.data = data;
  entry.handle_tag = std::move(handle_tag);
  {
    std::lock_guard<std::mutex> lock(g_virtual_mutex);
    g_virtual_roots.emplace(handle, std::move(entry));
  }
  return handle;
}

void RegistryProvider::UnregisterVirtualRoot(HKEY root) {
  std::lock_guard<std::mutex> lock(g_virtual_mutex);
  g_virtual_roots.erase(root);
}

bool RegistryProvider::IsVirtualRoot(HKEY root) {
  std::lock_guard<std::mutex> lock(g_virtual_mutex);
  return g_virtual_roots.find(root) != g_virtual_roots.end();
}

bool RegistryProvider::OpenKey(const RegistryNode& node, REGSAM sam, HKEY* key) {
  if (!node.root) {
    return false;
  }
  const wchar_t* sub = node.subkey.empty() ? nullptr : node.subkey.c_str();
  return RegOpenKeyExW(node.root, sub, 0, sam, key) == ERROR_SUCCESS;
}

bool RegistryProvider::HasSubKeys(const RegistryNode& node) {
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    const VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    return key && !key->children.empty();
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    DWORD subkey_count = 0;
    DWORD result = api->query_info(key.handle, nullptr, nullptr, &subkey_count, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr);
    CloseOfflineKey(key);
    return result == ERROR_SUCCESS && subkey_count > 0;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_ENUMERATE_SUB_KEYS, key.put())) {
    return false;
  }
  wchar_t name[1] = {};
  DWORD name_len = static_cast<DWORD>(_countof(name));
  LONG result = RegEnumKeyExW(key.get(), 0, name, &name_len, nullptr, nullptr, nullptr, nullptr);
  return result == ERROR_SUCCESS || result == ERROR_MORE_DATA;
}

bool RegistryProvider::QueryKeyInfo(const RegistryNode& node, KeyInfo* info) {
  if (!info) {
    return false;
  }
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    const VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return false;
    }
    info->subkey_count = static_cast<DWORD>(key->children.size());
    info->value_count = static_cast<DWORD>(key->values.size());
    info->last_write = {};
    return true;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    DWORD subkey_count = 0;
    DWORD value_count = 0;
    FILETIME last_write = {};
    DWORD result = api->query_info(key.handle, nullptr, nullptr, &subkey_count, nullptr, nullptr, &value_count, nullptr, nullptr, nullptr, &last_write);
    CloseOfflineKey(key);
    if (result != ERROR_SUCCESS) {
      return false;
    }
    info->subkey_count = subkey_count;
    info->value_count = value_count;
    info->last_write = last_write;
    return true;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_READ, key.put())) {
    return false;
  }

  DWORD subkey_count = 0;
  DWORD value_count = 0;
  FILETIME last_write = {};
  LONG result = RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, &subkey_count, nullptr, nullptr, &value_count, nullptr, nullptr, nullptr, &last_write);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  info->subkey_count = subkey_count;
  info->value_count = value_count;
  info->last_write = last_write;
  return true;
}

bool RegistryProvider::QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target) {
  if (!target) {
    return false;
  }
  target->clear();
  if (IsVirtualRoot(node.root)) {
    return false;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    DWORD type = 0;
    DWORD size = 0;
    DWORD result = api->get_value(key.handle, nullptr, L"SymbolicLinkValue", &type, nullptr, &size);
    if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
      CloseOfflineKey(key);
      return false;
    }
    if ((type != REG_LINK && type != REG_SZ && type != REG_EXPAND_SZ) || size == 0) {
      CloseOfflineKey(key);
      return false;
    }
    std::vector<wchar_t> buffer(size / sizeof(wchar_t) + 1, L'\0');
    result = api->get_value(key.handle, nullptr, L"SymbolicLinkValue", &type, buffer.data(), &size);
    CloseOfflineKey(key);
    if (result != ERROR_SUCCESS) {
      return false;
    }
    size_t length = size / sizeof(wchar_t);
    std::wstring value(buffer.data(), length);
    while (!value.empty() && value.back() == L'\0') {
      value.pop_back();
    }
    if (value.empty()) {
      return false;
    }
    *target = std::move(value);
    return true;
  }
  std::wstring nt_path = BuildNtPath(node);
  if (nt_path.empty()) {
    return false;
  }
  NtOpenKeyFn open_fn = GetNtOpenKey();
  if (!open_fn) {
    return false;
  }
  UNICODE_STRING name = {};
  name.Buffer = const_cast<PWSTR>(nt_path.c_str());
  name.Length = static_cast<USHORT>(nt_path.size() * sizeof(wchar_t));
  name.MaximumLength = name.Length;
  OBJECT_ATTRIBUTES attrs = {};
  InitializeObjectAttributes(&attrs, &name, OBJ_CASE_INSENSITIVE | OBJ_OPENLINK, nullptr, nullptr);

  HANDLE handle = nullptr;
  NTSTATUS status = open_fn(&handle, KEY_QUERY_VALUE, &attrs);
  if (!NT_SUCCESS(status) || !handle) {
    return false;
  }
  HKEY key = reinterpret_cast<HKEY>(handle);
  DWORD type = 0;
  DWORD size = 0;
  LONG result = RegQueryValueExW(key, L"SymbolicLinkValue", nullptr, &type, nullptr, &size);
  if (result != ERROR_SUCCESS || (type != REG_LINK && type != REG_SZ && type != REG_EXPAND_SZ) || size == 0) {
    RegCloseKey(key);
    return false;
  }
  std::vector<wchar_t> buffer(size / sizeof(wchar_t) + 1, L'\0');
  result = RegQueryValueExW(key, L"SymbolicLinkValue", nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &size);
  RegCloseKey(key);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  size_t length = size / sizeof(wchar_t);
  std::wstring value(buffer.data(), length);
  while (!value.empty() && value.back() == L'\0') {
    value.pop_back();
  }
  if (value.empty()) {
    return false;
  }
  *target = std::move(value);
  return true;
}

std::vector<std::wstring> RegistryProvider::EnumSubKeyNames(const RegistryNode& node, bool sorted) {
  std::vector<std::wstring> names;
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return names;
    }
    const VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return names;
    }
    names.reserve(key->children.size());
    for (const auto& entry : key->children) {
      if (entry.second) {
        names.push_back(entry.second->name);
      }
    }
    if (sorted) {
      std::sort(names.begin(), names.end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
    }
    return names;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return names;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return names;
    }
    DWORD subkey_count = 0;
    DWORD max_subkey_len = 0;
    DWORD result = api->query_info(key.handle, nullptr, nullptr, &subkey_count, &max_subkey_len, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr);
    if (result != ERROR_SUCCESS) {
      CloseOfflineKey(key);
      return names;
    }
    names.reserve(subkey_count);
    std::wstring buffer;
    buffer.resize(max_subkey_len + 1);
    for (DWORD index = 0; index < subkey_count; ++index) {
      DWORD name_len = static_cast<DWORD>(buffer.size());
      result = api->enum_key(key.handle, index, buffer.data(), &name_len, nullptr, nullptr, nullptr);
      if (result == ERROR_SUCCESS) {
        buffer[name_len] = L'\0';
        names.emplace_back(buffer.c_str());
      }
    }
    CloseOfflineKey(key);
    if (sorted) {
      std::sort(names.begin(), names.end());
    }
    return names;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_READ, key.put())) {
    return names;
  }

  DWORD subkey_count = 0;
  DWORD max_subkey_len = 0;
  if (RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, &subkey_count, &max_subkey_len, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr) != ERROR_SUCCESS) {
    return names;
  }

  names.reserve(subkey_count);
  std::wstring buffer;
  buffer.resize(max_subkey_len + 1);

  for (DWORD index = 0; index < subkey_count; ++index) {
    DWORD name_len = static_cast<DWORD>(buffer.size());
    FILETIME ft = {};
    LONG result = RegEnumKeyExW(key.get(), index, buffer.data(), &name_len, nullptr, nullptr, nullptr, &ft);
    if (result == ERROR_SUCCESS) {
      buffer[name_len] = L'\0';
      names.emplace_back(buffer.c_str());
    }
  }

  if (sorted) {
    std::sort(names.begin(), names.end());
  }
  return names;
}

std::vector<ValueInfo> RegistryProvider::EnumValueInfo(const RegistryNode& node) {
  std::vector<ValueInfo> values;
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return values;
    }
    const VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return values;
    }
    std::vector<const VirtualRegistryValue*> ordered;
    ordered.reserve(key->values.size());
    for (const auto& entry : key->values) {
      ordered.push_back(&entry.second);
    }
    std::sort(ordered.begin(), ordered.end(), [](const VirtualRegistryValue* left, const VirtualRegistryValue* right) {
      if (!left || !right) {
        return left != nullptr;
      }
      return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
    });
    values.reserve(ordered.size());
    for (const auto* value : ordered) {
      if (!value) {
        continue;
      }
      ValueInfo info;
      info.name = value->name;
      info.type = value->type;
      info.data_size = static_cast<DWORD>(value->data.size());
      values.emplace_back(std::move(info));
    }
    return values;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return values;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return values;
    }
    DWORD value_count = 0;
    DWORD max_value_name_len = 0;
    DWORD max_value_data_len = 0;
    DWORD result = api->query_info(key.handle, nullptr, nullptr, nullptr, nullptr, nullptr, &value_count, &max_value_name_len, &max_value_data_len, nullptr, nullptr);
    if (result != ERROR_SUCCESS) {
      CloseOfflineKey(key);
      return values;
    }
    values.reserve(value_count);
    std::wstring name_buffer;
    name_buffer.resize(max_value_name_len + 1);
    std::vector<BYTE> data;
    data.resize(max_value_data_len == 0 ? 1 : max_value_data_len);
    for (DWORD index = 0; index < value_count; ++index) {
      DWORD name_len = static_cast<DWORD>(name_buffer.size());
      DWORD data_len = static_cast<DWORD>(data.size());
      DWORD type = 0;
      result = api->enum_value(key.handle, index, name_buffer.data(), &name_len, &type, data.data(), &data_len);
      if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
        continue;
      }
      name_buffer[name_len] = L'\0';
      ValueInfo info;
      info.name.assign(name_buffer.c_str(), name_len);
      info.type = type;
      info.data_size = data_len;
      values.emplace_back(std::move(info));
    }
    CloseOfflineKey(key);
    return values;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_READ, key.put())) {
    return values;
  }

  DWORD value_count = 0;
  DWORD max_value_name_len = 0;
  if (RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, &value_count, &max_value_name_len, nullptr, nullptr, nullptr) != ERROR_SUCCESS) {
    return values;
  }

  values.reserve(value_count);
  std::wstring name_buffer;
  name_buffer.resize(max_value_name_len + 1);

  for (DWORD index = 0; index < value_count; ++index) {
    DWORD name_len = static_cast<DWORD>(name_buffer.size());
    DWORD data_len = 0;
    DWORD type = 0;
    LONG result = RegEnumValueW(key.get(), index, name_buffer.data(), &name_len, nullptr, &type, nullptr, &data_len);
    if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
      continue;
    }
    name_buffer[name_len] = L'\0';
    ValueInfo info;
    info.name.assign(name_buffer.c_str(), name_len);
    info.type = type;
    info.data_size = data_len;
    values.emplace_back(std::move(info));
  }

  return values;
}

std::vector<ValueEntry> RegistryProvider::EnumValues(const RegistryNode& node) {
  std::vector<ValueEntry> values;
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return values;
    }
    const VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return values;
    }
    std::vector<const VirtualRegistryValue*> ordered;
    ordered.reserve(key->values.size());
    for (const auto& entry : key->values) {
      ordered.push_back(&entry.second);
    }
    std::sort(ordered.begin(), ordered.end(), [](const VirtualRegistryValue* left, const VirtualRegistryValue* right) {
      if (!left || !right) {
        return left != nullptr;
      }
      return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
    });
    values.reserve(ordered.size());
    for (const auto* value : ordered) {
      if (!value) {
        continue;
      }
      ValueEntry entry;
      entry.name = value->name;
      entry.type = value->type;
      entry.data = value->data;
      values.emplace_back(std::move(entry));
    }
    return values;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return values;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return values;
    }
    DWORD value_count = 0;
    DWORD max_value_name_len = 0;
    DWORD max_value_data_len = 0;
    DWORD result = api->query_info(key.handle, nullptr, nullptr, nullptr, nullptr, nullptr, &value_count, &max_value_name_len, &max_value_data_len, nullptr, nullptr);
    if (result != ERROR_SUCCESS) {
      CloseOfflineKey(key);
      return values;
    }
    values.reserve(value_count);
    std::wstring name_buffer;
    name_buffer.resize(max_value_name_len + 1);
    std::vector<BYTE> data;
    data.resize(max_value_data_len == 0 ? 1 : max_value_data_len);
    for (DWORD index = 0; index < value_count; ++index) {
      DWORD name_len = static_cast<DWORD>(name_buffer.size());
      DWORD data_len = static_cast<DWORD>(data.size());
      DWORD type = 0;
      result = api->enum_value(key.handle, index, name_buffer.data(), &name_len, &type, data.data(), &data_len);
      if (result != ERROR_SUCCESS) {
        continue;
      }
      ValueEntry entry;
      entry.name.assign(name_buffer.c_str(), name_len);
      entry.type = type;
      entry.data.assign(data.begin(), data.begin() + data_len);
      values.emplace_back(std::move(entry));
    }
    CloseOfflineKey(key);
    return values;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_READ, key.put())) {
    return values;
  }

  DWORD value_count = 0;
  DWORD max_value_name_len = 0;
  DWORD max_value_data_len = 0;
  if (RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, &value_count, &max_value_name_len, &max_value_data_len, nullptr, nullptr) != ERROR_SUCCESS) {
    return values;
  }

  values.reserve(value_count);
  std::wstring name_buffer;
  name_buffer.resize(max_value_name_len + 1);

  std::vector<BYTE> data;
  data.resize(max_value_data_len == 0 ? 1 : max_value_data_len);

  for (DWORD index = 0; index < value_count; ++index) {
    DWORD name_len = static_cast<DWORD>(name_buffer.size());
    DWORD data_len = static_cast<DWORD>(data.size());
    DWORD type = 0;
    LONG result = RegEnumValueW(key.get(), index, name_buffer.data(), &name_len, nullptr, &type, data.data(), &data_len);
    if (result != ERROR_SUCCESS) {
      continue;
    }

    ValueEntry entry;
    entry.name.assign(name_buffer.c_str(), name_len);
    entry.type = type;
    entry.data.assign(data.begin(), data.begin() + data_len);
    values.emplace_back(std::move(entry));
  }

  return values;
}

bool RegistryProvider::EnumKeyStreaming(const RegistryNode& node, bool include_values, bool include_data, bool include_subkeys, KeyEnumResult* out_info, const ValueStreamCallback& value_callback, const SubkeyStreamCallback& subkey_callback) {
  if (out_info) {
    out_info->info = {};
    out_info->info_valid = false;
  }
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    const VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
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
      std::vector<const VirtualRegistryValue*> ordered;
      ordered.reserve(key->values.size());
      for (const auto& entry : key->values) {
        ordered.push_back(&entry.second);
      }
      std::sort(ordered.begin(), ordered.end(), [](const VirtualRegistryValue* left, const VirtualRegistryValue* right) {
        if (!left || !right) {
          return left != nullptr;
        }
        return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
      });
      for (const auto* value : ordered) {
        if (!value) {
          continue;
        }
        ValueInfo info;
        info.name = value->name;
        info.type = value->type;
        info.data_size = static_cast<DWORD>(value->data.size());
        const BYTE* buffer = include_data && !value->data.empty() ? value->data.data() : nullptr;
        if (!value_callback(info, buffer, static_cast<DWORD>(value->data.size()))) {
          return false;
        }
      }
    }
    if (include_subkeys && subkey_callback) {
      std::vector<std::wstring> names;
      names.reserve(key->children.size());
      for (const auto& entry : key->children) {
        if (entry.second) {
          names.push_back(entry.second->name);
        }
      }
      std::sort(names.begin(), names.end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
      for (const auto& name : names) {
        if (!subkey_callback(name)) {
          return false;
        }
      }
    }
    return true;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    DWORD subkey_count = 0;
    DWORD max_subkey_len = 0;
    DWORD value_count = 0;
    DWORD max_value_name_len = 0;
    DWORD max_value_data_len = 0;
    FILETIME last_write = {};
    DWORD result = api->query_info(key.handle, nullptr, nullptr, &subkey_count, &max_subkey_len, nullptr, &value_count, &max_value_name_len, &max_value_data_len, nullptr, &last_write);
    if (result != ERROR_SUCCESS) {
      CloseOfflineKey(key);
      return false;
    }
    if (out_info) {
      out_info->info.subkey_count = subkey_count;
      out_info->info.value_count = value_count;
      out_info->info.last_write = last_write;
      out_info->info_valid = true;
    }

    if (include_values && value_callback) {
      std::wstring name_buffer;
      name_buffer.resize(max_value_name_len + 1);
      std::vector<BYTE> data;
      if (include_data) {
        data.resize(max_value_data_len == 0 ? 1 : max_value_data_len);
      }
      for (DWORD index = 0; index < value_count; ++index) {
        DWORD name_len = static_cast<DWORD>(name_buffer.size());
        DWORD data_len = include_data ? static_cast<DWORD>(data.size()) : 0;
        DWORD type = 0;
        result = api->enum_value(key.handle, index, name_buffer.data(), &name_len, &type, include_data ? data.data() : nullptr, &data_len);
        if (result == ERROR_MORE_DATA && include_data) {
          data.resize(data_len == 0 ? 1 : data_len);
          name_len = static_cast<DWORD>(name_buffer.size());
          data_len = static_cast<DWORD>(data.size());
          result = api->enum_value(key.handle, index, name_buffer.data(), &name_len, &type, data.data(), &data_len);
        }
        if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
          continue;
        }
        name_buffer[name_len] = L'\0';
        ValueInfo info;
        info.name.assign(name_buffer.c_str(), name_len);
        info.type = type;
        info.data_size = data_len;
        const BYTE* buffer = include_data && data_len > 0 ? data.data() : nullptr;
        if (!value_callback(info, buffer, data_len)) {
          CloseOfflineKey(key);
          return false;
        }
      }
    }

    if (include_subkeys && subkey_callback) {
      std::wstring subkey_buffer;
      subkey_buffer.resize(max_subkey_len + 1);
      for (DWORD index = 0; index < subkey_count; ++index) {
        DWORD name_len = static_cast<DWORD>(subkey_buffer.size());
        result = api->enum_key(key.handle, index, subkey_buffer.data(), &name_len, nullptr, nullptr, nullptr);
        if (result != ERROR_SUCCESS) {
          continue;
        }
        subkey_buffer[name_len] = L'\0';
        if (!subkey_callback(subkey_buffer.c_str())) {
          CloseOfflineKey(key);
          return false;
        }
      }
    }

    CloseOfflineKey(key);
    return true;
  }

  util::UniqueHKey key;
  if (!OpenKey(node, KEY_READ, key.put())) {
    return false;
  }

  DWORD subkey_count = 0;
  DWORD max_subkey_len = 0;
  DWORD value_count = 0;
  DWORD max_value_name_len = 0;
  DWORD max_value_data_len = 0;
  FILETIME last_write = {};
  if (RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, &subkey_count, &max_subkey_len, nullptr, &value_count, &max_value_name_len, &max_value_data_len, nullptr, &last_write) != ERROR_SUCCESS) {
    return false;
  }
  if (out_info) {
    out_info->info.subkey_count = subkey_count;
    out_info->info.value_count = value_count;
    out_info->info.last_write = last_write;
    out_info->info_valid = true;
  }

  if (include_values && value_callback) {
    std::wstring name_buffer;
    name_buffer.resize(max_value_name_len + 1);
    std::vector<BYTE> data;
    if (include_data) {
      data.resize(max_value_data_len == 0 ? 1 : max_value_data_len);
    }
    for (DWORD index = 0; index < value_count; ++index) {
      DWORD name_len = static_cast<DWORD>(name_buffer.size());
      DWORD data_len = include_data ? static_cast<DWORD>(data.size()) : 0;
      DWORD type = 0;
      LONG result = RegEnumValueW(key.get(), index, name_buffer.data(), &name_len, nullptr, &type, include_data ? data.data() : nullptr, &data_len);
      if (result == ERROR_MORE_DATA && include_data) {
        data.resize(data_len == 0 ? 1 : data_len);
        name_len = static_cast<DWORD>(name_buffer.size());
        data_len = static_cast<DWORD>(data.size());
        result = RegEnumValueW(key.get(), index, name_buffer.data(), &name_len, nullptr, &type, data.data(), &data_len);
      }
      if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
        continue;
      }
      name_buffer[name_len] = L'\0';
      ValueInfo info;
      info.name.assign(name_buffer.c_str(), name_len);
      info.type = type;
      info.data_size = data_len;
      const BYTE* buffer = include_data && data_len > 0 ? data.data() : nullptr;
      if (!value_callback(info, buffer, data_len)) {
        return false;
      }
    }
  }

  if (include_subkeys && subkey_callback) {
    std::wstring subkey_buffer;
    subkey_buffer.resize(max_subkey_len + 1);
    for (DWORD index = 0; index < subkey_count; ++index) {
      DWORD name_len = static_cast<DWORD>(subkey_buffer.size());
      FILETIME ft = {};
      LONG result = RegEnumKeyExW(key.get(), index, subkey_buffer.data(), &name_len, nullptr, nullptr, nullptr, &ft);
      if (result != ERROR_SUCCESS) {
        continue;
      }
      subkey_buffer[name_len] = L'\0';
      if (!subkey_callback(subkey_buffer.c_str())) {
        return false;
      }
    }
  }

  return true;
}

bool RegistryProvider::QueryValue(const RegistryNode& node, const std::wstring& value_name, ValueEntry* out) {
  if (!out) {
    return false;
  }
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    const VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return false;
    }
    std::wstring name_lower = ToLower(value_name);
    auto it = key->values.find(name_lower);
    if (it == key->values.end()) {
      return false;
    }
    out->name = it->second.name;
    out->type = it->second.type;
    out->data = it->second.data;
    return true;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    const wchar_t* name = value_name.empty() ? nullptr : value_name.c_str();
    DWORD type = 0;
    DWORD size = 0;
    DWORD result = api->get_value(key.handle, nullptr, name, &type, nullptr, &size);
    if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
      CloseOfflineKey(key);
      return false;
    }
    std::vector<BYTE> data;
    if (size > 0) {
      data.resize(size);
    }
    result = api->get_value(key.handle, nullptr, name, &type, data.empty() ? nullptr : data.data(), &size);
    CloseOfflineKey(key);
    if (result != ERROR_SUCCESS) {
      return false;
    }
    data.resize(size);
    out->name = value_name;
    out->type = type;
    out->data = std::move(data);
    return true;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_QUERY_VALUE, key.put())) {
    return false;
  }

  const wchar_t* name = value_name.empty() ? nullptr : value_name.c_str();
  DWORD type = 0;
  DWORD size = 0;
  LONG result = RegQueryValueExW(key.get(), name, nullptr, &type, nullptr, &size);
  if (result != ERROR_SUCCESS) {
    return false;
  }

  std::vector<BYTE> data;
  if (size > 0) {
    data.resize(size);
  }
  result = RegQueryValueExW(key.get(), name, nullptr, &type, data.empty() ? nullptr : data.data(), &size);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  data.resize(size);
  out->name = value_name;
  out->type = type;
  out->data = std::move(data);
  return true;
}

DWORD RegistryProvider::NormalizeValueType(DWORD type) {
  DWORD base = type & 0xFFFF;
  switch (base) {
  case REG_NONE:
  case REG_SZ:
  case REG_EXPAND_SZ:
  case REG_MULTI_SZ:
  case REG_DWORD:
  case REG_QWORD:
  case REG_BINARY:
  case REG_RESOURCE_LIST:
  case REG_FULL_RESOURCE_DESCRIPTOR:
  case REG_RESOURCE_REQUIREMENTS_LIST:
  case REG_LINK:
  case REG_DWORD_BIG_ENDIAN:
    return base;
  default:
    return type;
  }
}

std::wstring RegistryProvider::FormatValueType(DWORD type) {
  DWORD base_type = NormalizeValueType(type);
  bool has_flags = base_type != type;
  const wchar_t* label = nullptr;
  switch (base_type) {
  case REG_NONE:
    label = L"REG_NONE";
    break;
  case REG_SZ:
    label = L"REG_SZ";
    break;
  case REG_EXPAND_SZ:
    label = L"REG_EXPAND_SZ";
    break;
  case REG_MULTI_SZ:
    label = L"REG_MULTI_SZ";
    break;
  case REG_DWORD:
    label = L"REG_DWORD";
    break;
  case REG_QWORD:
    label = L"REG_QWORD";
    break;
  case REG_BINARY:
    label = L"REG_BINARY";
    break;
  case REG_RESOURCE_LIST:
    label = L"REG_RESOURCE_LIST";
    break;
  case REG_FULL_RESOURCE_DESCRIPTOR:
    label = L"REG_FULL_RESOURCE_DESCRIPTOR";
    break;
  case REG_RESOURCE_REQUIREMENTS_LIST:
    label = L"REG_RESOURCE_REQUIREMENTS_LIST";
    break;
  case REG_LINK:
    label = L"REG_LINK";
    break;
  case REG_DWORD_BIG_ENDIAN:
    label = L"REG_DWORD_BIG_ENDIAN";
    break;
  default:
    break;
  }
  if (!label) {
    wchar_t buffer[32] = {};
    swprintf_s(buffer, L"REG_UNKNOWN (0x%X)", type);
    return buffer;
  }
  if (has_flags) {
    wchar_t buffer[64] = {};
    swprintf_s(buffer, L"%s (0x%X)", label, type);
    return buffer;
  }
  return label;
}

std::wstring RegistryProvider::FormatValueData(DWORD type, const BYTE* data, DWORD size) {
  if (!data || size == 0) {
    return L"";
  }

  DWORD base_type = NormalizeValueType(type);
  switch (base_type) {
  case REG_SZ:
  case REG_EXPAND_SZ: {
    size_t wchar_count = size / sizeof(wchar_t);
    std::wstring text(reinterpret_cast<const wchar_t*>(data), wchar_count);
    while (!text.empty() && text.back() == L'\0') {
      text.pop_back();
    }
    return text;
  }
  case REG_MULTI_SZ: {
    size_t wchar_count = size / sizeof(wchar_t);
    const wchar_t* start = reinterpret_cast<const wchar_t*>(data);
    std::wstring joined;
    size_t offset = 0;
    while (offset < wchar_count) {
      const wchar_t* current = start + offset;
      size_t len = wcsnlen_s(current, wchar_count - offset);
      if (len == 0) {
        break;
      }
      if (!joined.empty()) {
        joined.append(L"; ");
      }
      joined.append(current, len);
      offset += len + 1;
    }
    return joined;
  }
  case REG_DWORD: {
    if (size >= sizeof(DWORD)) {
      DWORD value = *reinterpret_cast<const DWORD*>(data);
      wchar_t buffer[32] = {};
      swprintf_s(buffer, L"0x%08X (%u)", value, value);
      return buffer;
    }
    break;
  }
  case REG_QWORD: {
    if (size >= sizeof(unsigned long long)) {
      unsigned long long value = *reinterpret_cast<const unsigned long long*>(data);
      wchar_t buffer[48] = {};
      swprintf_s(buffer, L"0x%016llX (%llu)", value, value);
      return buffer;
    }
    break;
  }
  case REG_BINARY: {
    return util::ToHex(data, size, 32);
  }
  case REG_RESOURCE_LIST:
  case REG_FULL_RESOURCE_DESCRIPTOR:
  case REG_RESOURCE_REQUIREMENTS_LIST: {
    return util::ToHex(data, size, 32);
  }
  case REG_LINK: {
    size_t wchar_count = size / sizeof(wchar_t);
    std::wstring text(reinterpret_cast<const wchar_t*>(data), wchar_count);
    while (!text.empty() && text.back() == L'\0') {
      text.pop_back();
    }
    return text;
  }
  default:
    break;
  }

  return util::ToHex(data, size, 32);
}

std::wstring RegistryProvider::FormatValueDataForDisplay(DWORD type, const BYTE* data, DWORD size) {
  DWORD base_type = NormalizeValueType(type);
  std::wstring base = FormatValueData(type, data, size);
  if (base.empty()) {
    return base;
  }

  if ((base_type == REG_SZ || base_type == REG_EXPAND_SZ) && base.front() == L'@') {
    wchar_t resolved[512] = {};
    if (SUCCEEDED(SHLoadIndirectString(base.c_str(), resolved, static_cast<UINT>(_countof(resolved)), nullptr))) {
      if (wcslen(resolved) > 0) {
        return resolved;
      }
    }
  }

  if (base_type == REG_EXPAND_SZ) {
    wchar_t expanded[512] = {};
    DWORD length = ExpandEnvironmentStringsW(base.c_str(), expanded, static_cast<DWORD>(_countof(expanded)));
    if (length > 0 && length < _countof(expanded)) {
      std::wstring result(expanded);
      if (!result.empty() && result != base) {
        return result;
      }
    }
  }

  return base;
}

bool RegistryProvider::CreateKey(const RegistryNode& node, const std::wstring& name) {
  if (name.empty()) {
    return false;
  }
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    VirtualRegistryKey* parent = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!parent) {
      return false;
    }
    std::wstring name_lower = ToLower(name);
    auto it = parent->children.find(name_lower);
    if (it != parent->children.end()) {
      return true;
    }
    auto child = std::make_unique<VirtualRegistryKey>();
    child->name = name;
    parent->children.emplace(name_lower, std::move(child));
    return true;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    ORHKEY created = nullptr;
    DWORD disposition = 0;
    DWORD result = api->create_key(key.handle, name.c_str(), nullptr, 0, nullptr, &created, &disposition);
    if (created) {
      api->close_key(created);
    }
    CloseOfflineKey(key);
    return result == ERROR_SUCCESS;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_WRITE, key.put())) {
    return false;
  }
  HKEY created = nullptr;
  DWORD disposition = 0;
  LONG result = RegCreateKeyExW(key.get(), name.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, &created, &disposition);
  if (created) {
    RegCloseKey(created);
  }
  return result == ERROR_SUCCESS;
}

bool RegistryProvider::DeleteKey(const RegistryNode& node) {
  if (node.subkey.empty()) {
    return false;
  }
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    std::wstring parent_path;
    std::wstring name;
    if (!SplitSubKey(node.subkey, &parent_path, &name)) {
      return false;
    }
    VirtualRegistryKey* parent = FindVirtualKey(virtual_data->root.get(), parent_path);
    if (!parent) {
      return false;
    }
    std::wstring name_lower = ToLower(name);
    return parent->children.erase(name_lower) > 0;
  }
  std::wstring parent_path;
  std::wstring name;
  if (!SplitSubKey(node.subkey, &parent_path, &name)) {
    return false;
  }
  RegistryNode parent = node;
  parent.subkey = parent_path;
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(parent);
    if (!key.handle) {
      return false;
    }
    bool ok = DeleteOfflineSubtree(key.handle, name);
    CloseOfflineKey(key);
    return ok;
  }
  util::UniqueHKey key;
  if (!OpenKey(parent, KEY_WRITE, key.put())) {
    return false;
  }
  return RegDeleteTreeW(key.get(), name.c_str()) == ERROR_SUCCESS;
}

bool RegistryProvider::RenameKey(const RegistryNode& node, const std::wstring& new_name) {
  if (node.subkey.empty() || new_name.empty()) {
    return false;
  }
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    std::wstring parent_path;
    std::wstring name;
    if (!SplitSubKey(node.subkey, &parent_path, &name)) {
      return false;
    }
    VirtualRegistryKey* parent = FindVirtualKey(virtual_data->root.get(), parent_path);
    if (!parent) {
      return false;
    }
    std::wstring old_lower = ToLower(name);
    std::wstring new_lower = ToLower(new_name);
    auto it = parent->children.find(old_lower);
    if (it == parent->children.end()) {
      return false;
    }
    if (parent->children.find(new_lower) != parent->children.end()) {
      return false;
    }
    std::unique_ptr<VirtualRegistryKey> moved = std::move(it->second);
    parent->children.erase(it);
    if (moved) {
      moved->name = new_name;
    }
    parent->children.emplace(new_lower, std::move(moved));
    return true;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    DWORD result = api->rename_key(key.handle, new_name.c_str());
    CloseOfflineKey(key);
    return result == ERROR_SUCCESS;
  }
  std::wstring parent_path;
  std::wstring name;
  if (!SplitSubKey(node.subkey, &parent_path, &name)) {
    return false;
  }
  RegistryNode parent = node;
  parent.subkey = parent_path;
  util::UniqueHKey key;
  if (!OpenKey(parent, KEY_WRITE, key.put())) {
    return false;
  }
  return RegRenameKey(key.get(), name.c_str(), new_name.c_str()) == ERROR_SUCCESS;
}

bool RegistryProvider::DeleteValue(const RegistryNode& node, const std::wstring& value_name) {
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return false;
    }
    std::wstring name_lower = ToLower(value_name);
    return key->values.erase(name_lower) > 0;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    DWORD result = api->delete_value(key.handle, value_name.empty() ? nullptr : value_name.c_str());
    CloseOfflineKey(key);
    return result == ERROR_SUCCESS;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_SET_VALUE, key.put())) {
    return false;
  }
  return RegDeleteValueW(key.get(), value_name.c_str()) == ERROR_SUCCESS;
}

bool RegistryProvider::SetValue(const RegistryNode& node, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& data) {
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return false;
    }
    std::wstring name_lower = ToLower(value_name);
    VirtualRegistryValue& value = key->values[name_lower];
    value.name = value_name;
    value.type = type;
    value.data = data;
    return true;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    const BYTE* buffer = data.empty() ? nullptr : data.data();
    DWORD size = static_cast<DWORD>(data.size());
    DWORD result = api->set_value(key.handle, value_name.empty() ? nullptr : value_name.c_str(), type, buffer, size);
    CloseOfflineKey(key);
    return result == ERROR_SUCCESS;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_SET_VALUE, key.put())) {
    return false;
  }
  const BYTE* buffer = data.empty() ? nullptr : data.data();
  DWORD size = static_cast<DWORD>(data.size());
  return RegSetValueExW(key.get(), value_name.c_str(), 0, type, buffer, size) == ERROR_SUCCESS;
}

bool RegistryProvider::RenameValue(const RegistryNode& node, const std::wstring& old_name, const std::wstring& new_name) {
  if (new_name.empty()) {
    return false;
  }
  std::shared_ptr<VirtualRegistryData> virtual_data;
  if (GetVirtualRootData(node.root, &virtual_data, nullptr)) {
    if (!virtual_data || !virtual_data->root) {
      return false;
    }
    VirtualRegistryKey* key = FindVirtualKey(virtual_data->root.get(), node.subkey);
    if (!key) {
      return false;
    }
    std::wstring old_lower = ToLower(old_name);
    std::wstring new_lower = ToLower(new_name);
    auto it = key->values.find(old_lower);
    if (it == key->values.end()) {
      return false;
    }
    if (key->values.find(new_lower) != key->values.end()) {
      return false;
    }
    VirtualRegistryValue value = std::move(it->second);
    key->values.erase(it);
    value.name = new_name;
    key->values.emplace(new_lower, std::move(value));
    return true;
  }
  if (IsOfflineNode(node)) {
    OffregApi* api = GetOffreg();
    if (!api) {
      return false;
    }
    OfflineKey key = OpenOfflineKey(node);
    if (!key.handle) {
      return false;
    }
    DWORD type = 0;
    DWORD size = 0;
    DWORD result = api->get_value(key.handle, nullptr, old_name.c_str(), &type, nullptr, &size);
    if (result != ERROR_SUCCESS && result != ERROR_MORE_DATA) {
      CloseOfflineKey(key);
      return false;
    }
    std::vector<BYTE> data;
    if (size > 0) {
      data.resize(size);
    }
    result = api->get_value(key.handle, nullptr, old_name.c_str(), &type, data.empty() ? nullptr : data.data(), &size);
    if (result != ERROR_SUCCESS) {
      CloseOfflineKey(key);
      return false;
    }
    result = api->set_value(key.handle, new_name.c_str(), type, data.empty() ? nullptr : data.data(), size);
    if (result != ERROR_SUCCESS) {
      CloseOfflineKey(key);
      return false;
    }
    result = api->delete_value(key.handle, old_name.c_str());
    CloseOfflineKey(key);
    return result == ERROR_SUCCESS;
  }
  util::UniqueHKey key;
  if (!OpenKey(node, KEY_QUERY_VALUE | KEY_SET_VALUE, key.put())) {
    return false;
  }

  DWORD type = 0;
  DWORD size = 0;
  LONG result = RegQueryValueExW(key.get(), old_name.c_str(), nullptr, &type, nullptr, &size);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  std::vector<BYTE> data(size);
  result = RegQueryValueExW(key.get(), old_name.c_str(), nullptr, &type, data.data(), &size);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  result = RegSetValueExW(key.get(), new_name.c_str(), 0, type, data.data(), size);
  if (result != ERROR_SUCCESS) {
    return false;
  }
  return RegDeleteValueW(key.get(), old_name.c_str()) == ERROR_SUCCESS;
}

} // namespace regkit
