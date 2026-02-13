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

#pragma once

#include <windows.h>

#include <functional>
#include <memory>
#include <string>
#include <unordered_map>
#include <vector>

namespace regkit {

struct RegistryNode {
  HKEY root = nullptr;
  std::wstring subkey;
  std::wstring root_name;
  bool children_loaded = false;
  bool simulated = false;
};

enum class RegistryRootGroup {
  kStandard,
  kReal,
};

struct RegistryRootEntry {
  HKEY root = nullptr;
  std::wstring display_name;
  std::wstring path_name;
  std::wstring subkey_prefix;
  RegistryRootGroup group = RegistryRootGroup::kStandard;
};

struct ValueEntry {
  std::wstring name;
  DWORD type = 0;
  std::vector<BYTE> data;
};

struct ValueInfo {
  std::wstring name;
  DWORD type = 0;
  DWORD data_size = 0;
};

struct KeyInfo {
  DWORD subkey_count = 0;
  DWORD value_count = 0;
  FILETIME last_write = {};
};

class RegistryProvider {
public:
  struct VirtualRegistryValue {
    std::wstring name;
    DWORD type = REG_NONE;
    std::vector<BYTE> data;
  };

  struct VirtualRegistryKey {
    std::wstring name;
    std::unordered_map<std::wstring, VirtualRegistryValue> values;
    std::unordered_map<std::wstring, std::unique_ptr<VirtualRegistryKey>> children;
  };

  struct VirtualRegistryData {
    std::wstring root_name;
    std::unique_ptr<VirtualRegistryKey> root;
  };

  static std::vector<RegistryRootEntry> DefaultRoots(bool include_extra = false);
  static std::wstring RootName(HKEY root);
  static std::wstring BuildPath(const RegistryNode& node);
  static std::wstring BuildNtPath(const RegistryNode& node);
  static bool HasSubKeys(const RegistryNode& node);
  static std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node, bool sorted = true);
  static std::vector<ValueInfo> EnumValueInfo(const RegistryNode& node);
  static std::vector<ValueEntry> EnumValues(const RegistryNode& node);
  using ValueStreamCallback = std::function<bool(const ValueInfo& info, const BYTE* data, DWORD data_size)>;
  using SubkeyStreamCallback = std::function<bool(const std::wstring& name)>;
  struct KeyEnumResult {
    KeyInfo info;
    bool info_valid = false;
  };
  static bool EnumKeyStreaming(const RegistryNode& node, bool include_values, bool include_data, bool include_subkeys, KeyEnumResult* out_info, const ValueStreamCallback& value_callback, const SubkeyStreamCallback& subkey_callback);
  static bool QueryValue(const RegistryNode& node, const std::wstring& value_name, ValueEntry* out);
  static DWORD NormalizeValueType(DWORD type);
  static std::wstring FormatValueType(DWORD type);
  static std::wstring FormatValueData(DWORD type, const BYTE* data, DWORD size);
  static std::wstring FormatValueDataForDisplay(DWORD type, const BYTE* data, DWORD size);
  static bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info);
  static bool QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target);
  static bool OpenOfflineHive(const std::wstring& path, HKEY* root, std::wstring* error);
  static bool SaveOfflineHive(HKEY root, const std::wstring& path, std::wstring* error);
  static bool CloseOfflineHive(HKEY root, std::wstring* error);
  static void SetOfflineRoot(HKEY root);
  static void SetOfflineRoots(const std::vector<HKEY>& roots);
  static HKEY RegisterVirtualRoot(const std::wstring& root_name, const std::shared_ptr<VirtualRegistryData>& data);
  static void UnregisterVirtualRoot(HKEY root);
  static bool IsVirtualRoot(HKEY root);
  static bool CreateKey(const RegistryNode& node, const std::wstring& name);
  static bool DeleteKey(const RegistryNode& node);
  static bool RenameKey(const RegistryNode& node, const std::wstring& new_name);
  static bool DeleteValue(const RegistryNode& node, const std::wstring& value_name);
  static bool SetValue(const RegistryNode& node, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& data);
  static bool RenameValue(const RegistryNode& node, const std::wstring& old_name, const std::wstring& new_name);

private:
  static bool OpenKey(const RegistryNode& node, REGSAM sam, HKEY* key);
};

} // namespace regkit
