// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "registry/registry_provider.h"
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
    bool include_subkeys, RegistryProvider::KeyEnumResult* out_info,
    const RegistryProvider::ValueStreamCallback& value_callback,
    const RegistryProvider::SubkeyStreamCallback& subkey_callback,
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

bool HasSubKeys(const RegistryNode& node);
bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info);
bool QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target);
std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node,
                                         bool sorted);
bool EnumKeyStreaming(
    const RegistryNode& node, bool include_values, bool include_data,
    bool include_subkeys, RegistryProvider::KeyEnumResult* out_info,
    const RegistryProvider::ValueStreamCallback& value_callback,
    const RegistryProvider::SubkeyStreamCallback& subkey_callback,
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
    RegistryProvider::KeyEnumResult* out_info,
    const RegistryProvider::ValueStreamCallback& value_callback,
    const RegistryProvider::SubkeyStreamCallback& subkey_callback,
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
