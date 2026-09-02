// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_store.h"
#include "registry/virtual_registry.h"

namespace regkit::registry_backend {

namespace live {

bool HasSubKeys(const RegistryNode& node);
bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info);
bool QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target);
std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node,
                                          bool sorted);
bool EnumKeyStreaming(
    const RegistryNode& node, bool include_values, bool include_data,
    bool include_subkeys, RegistryStore::KeyEnumResult* out_info,
    const RegistryStore::ValueStreamCallback& value_callback,
    const RegistryStore::SubkeyStreamCallback& subkey_callback,
    DWORD max_data_size);
bool QueryValue(const RegistryNode& node, const std::wstring& value_name,
                ValueEntry* out);
bool CreateKey(const RegistryNode& node, const std::wstring& name);
bool DeleteKey(const RegistryNode& node);
bool RenameKey(const RegistryNode& node, const std::wstring& new_name);
bool DeleteValue(const RegistryNode& node, const std::wstring& value_name);
bool SetValue(const RegistryNode& node, const std::wstring& value_name,
              DWORD type, const std::vector<BYTE>& data);
bool RenameValue(const RegistryNode& node, const std::wstring& old_name,
                 const std::wstring& new_name);

} // namespace live

namespace offline {

bool OpenHive(const std::wstring& path, HKEY* root, std::wstring* error);
bool SaveHive(HKEY root, const std::wstring& path, std::wstring* error);
bool CloseHive(HKEY root, std::wstring* error);
void SetRoots(const std::vector<HKEY>& roots);
bool Owns(HKEY root);
void AddRoot(HKEY root);
void RemoveRoot(HKEY root);

bool HasSubKeys(const RegistryNode& node);
bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info);
bool QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target);
std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node,
                                          bool sorted);
bool EnumKeyStreaming(
    const RegistryNode& node, bool include_values, bool include_data,
    bool include_subkeys, RegistryStore::KeyEnumResult* out_info,
    const RegistryStore::ValueStreamCallback& value_callback,
    const RegistryStore::SubkeyStreamCallback& subkey_callback,
    DWORD max_data_size);
bool QueryValue(const RegistryNode& node, const std::wstring& value_name,
                ValueEntry* out);
bool CreateKey(const RegistryNode& node, const std::wstring& name);
bool DeleteKey(const RegistryNode& node);
bool RenameKey(const RegistryNode& node, const std::wstring& new_name);
bool DeleteValue(const RegistryNode& node, const std::wstring& value_name);
bool SetValue(const RegistryNode& node, const std::wstring& value_name,
              DWORD type, const std::vector<BYTE>& data);
bool RenameValue(const RegistryNode& node, const std::wstring& old_name,
                 const std::wstring& new_name);

} // namespace offline

namespace virtual_store {

HKEY RegisterRoot(const std::wstring& root_name,
                  const std::shared_ptr<VirtualRegistryData>& data);
void UnregisterRoot(HKEY root);
std::shared_ptr<VirtualRegistryData> FindRoot(HKEY root,
                                              std::wstring* root_name = nullptr);

bool HasSubKeys(const VirtualRegistryData& data, const RegistryNode& node);
bool QueryKeyInfo(const VirtualRegistryData& data, const RegistryNode& node,
                  KeyInfo* info);
std::vector<std::wstring> EnumSubKeyNames(const VirtualRegistryData& data,
                                          const RegistryNode& node,
                                          bool sorted);
bool EnumKeyStreaming(
    const VirtualRegistryData& data, const RegistryNode& node,
    bool include_values, bool include_data, bool include_subkeys,
    RegistryStore::KeyEnumResult* out_info,
    const RegistryStore::ValueStreamCallback& value_callback,
    const RegistryStore::SubkeyStreamCallback& subkey_callback,
    DWORD max_data_size);
bool QueryValue(const VirtualRegistryData& data, const RegistryNode& node,
                const std::wstring& value_name, ValueEntry* out);
bool CreateKey(VirtualRegistryData& data, const RegistryNode& node,
               const std::wstring& name);
bool DeleteKey(VirtualRegistryData& data, const RegistryNode& node);
bool RenameKey(VirtualRegistryData& data, const RegistryNode& node,
               const std::wstring& new_name);
bool DeleteValue(VirtualRegistryData& data, const RegistryNode& node,
                 const std::wstring& value_name);
bool SetValue(VirtualRegistryData& data, const RegistryNode& node,
              const std::wstring& value_name, DWORD type,
              const std::vector<BYTE>& value_data);
bool RenameValue(VirtualRegistryData& data, const RegistryNode& node,
                 const std::wstring& old_name,
                 const std::wstring& new_name);

} // namespace virtual_store
} // namespace regkit::registry_backend
