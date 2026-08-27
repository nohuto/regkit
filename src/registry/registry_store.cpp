// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_store.h"

#include "registry/registry_backends.h"

namespace regkit {
namespace {

using registry_backend::virtual_store::FindRoot;

template <typename VirtualCall, typename OfflineCall, typename LiveCall>
decltype(auto) Dispatch(const RegistryNode& node, VirtualCall&& virtual_call,
                        OfflineCall&& offline_call, LiveCall&& live_call) {
  if (std::shared_ptr<VirtualRegistryData> data = FindRoot(node.root)) {
    return virtual_call(*data);
  }
  if (registry_backend::offline::Owns(node.root)) {
    return offline_call();
  }
  return live_call();
}

} // namespace

std::vector<RegistryRootEntry> RegistryStore::DefaultRoots(
    bool include_extra) {
  std::vector<RegistryRootEntry> roots = {
      {HKEY_CLASSES_ROOT, L"HKEY_CLASSES_ROOT", L"HKEY_CLASSES_ROOT", L""},
      {HKEY_CURRENT_USER, L"HKEY_CURRENT_USER", L"HKEY_CURRENT_USER", L""},
      {HKEY_LOCAL_MACHINE, L"HKEY_LOCAL_MACHINE", L"HKEY_LOCAL_MACHINE",
       L""},
      {HKEY_USERS, L"HKEY_USERS", L"HKEY_USERS", L""},
      {HKEY_CURRENT_CONFIG, L"HKEY_CURRENT_CONFIG", L"HKEY_CURRENT_CONFIG",
       L""},
  };
  if (include_extra) {
    roots.push_back({HKEY_PERFORMANCE_DATA, L"HKEY_PERFORMANCE_DATA",
                     L"HKEY_PERFORMANCE_DATA", L""});
    roots.push_back({HKEY_PERFORMANCE_TEXT, L"HKEY_PERFORMANCE_TEXT",
                     L"HKEY_PERFORMANCE_TEXT", L""});
    roots.push_back({HKEY_PERFORMANCE_NLSTEXT, L"HKEY_PERFORMANCE_NLSTEXT",
                     L"HKEY_PERFORMANCE_NLSTEXT", L""});
  }
  return roots;
}

bool RegistryStore::OpenOfflineHive(const std::wstring& path, HKEY* root,
                                    std::wstring* error) {
  return registry_backend::offline::OpenHive(path, root, error);
}

bool RegistryStore::SaveOfflineHive(HKEY root, const std::wstring& path,
                                    std::wstring* error) {
  return registry_backend::offline::SaveHive(root, path, error);
}

bool RegistryStore::CloseOfflineHive(HKEY root, std::wstring* error) {
  return registry_backend::offline::CloseHive(root, error);
}

void RegistryStore::SetOfflineRoots(const std::vector<HKEY>& roots) {
  registry_backend::offline::SetRoots(roots);
}

HKEY RegistryStore::RegisterVirtualRoot(
    const std::wstring& root_name,
    const std::shared_ptr<VirtualRegistryData>& data) {
  return registry_backend::virtual_store::RegisterRoot(root_name, data);
}

void RegistryStore::UnregisterVirtualRoot(HKEY root) {
  registry_backend::virtual_store::UnregisterRoot(root);
}

bool RegistryStore::IsVirtualRoot(HKEY root) {
  return static_cast<bool>(FindRoot(root));
}

bool RegistryStore::GetVirtualRootName(HKEY root,
                                       std::wstring* root_name) {
  return static_cast<bool>(FindRoot(root, root_name));
}

bool RegistryStore::HasSubKeys(const RegistryNode& node) {
  return Dispatch(
      node,
      [&](const VirtualRegistryData& data) {
        return registry_backend::virtual_store::HasSubKeys(data, node);
      },
      [&] { return registry_backend::offline::HasSubKeys(node); },
      [&] { return registry_backend::live::HasSubKeys(node); });
}

bool RegistryStore::QueryKeyInfo(const RegistryNode& node, KeyInfo* info) {
  if (!info) {
    return false;
  }
  return Dispatch(
      node,
      [&](const VirtualRegistryData& data) {
        return registry_backend::virtual_store::QueryKeyInfo(data, node,
                                                             info);
      },
      [&] { return registry_backend::offline::QueryKeyInfo(node, info); },
      [&] { return registry_backend::live::QueryKeyInfo(node, info); });
}

bool RegistryStore::QuerySymbolicLinkTarget(const RegistryNode& node,
                                            std::wstring* target) {
  if (!target) {
    return false;
  }
  target->clear();
  if (FindRoot(node.root)) {
    return false;
  }
  return registry_backend::offline::Owns(node.root)
             ? registry_backend::offline::QuerySymbolicLinkTarget(node,
                                                                  target)
             : registry_backend::live::QuerySymbolicLinkTarget(node, target);
}

std::vector<std::wstring> RegistryStore::EnumSubKeyNames(
    const RegistryNode& node, bool sorted) {
  return Dispatch(
      node,
      [&](const VirtualRegistryData& data) {
        return registry_backend::virtual_store::EnumSubKeyNames(data, node,
                                                                sorted);
      },
      [&] {
        return registry_backend::offline::EnumSubKeyNames(node, sorted);
      },
      [&] { return registry_backend::live::EnumSubKeyNames(node, sorted); });
}

bool RegistryStore::EnumKeyStreaming(
    const RegistryNode& node, bool include_values, bool include_data,
    bool include_subkeys, KeyEnumResult* out_info,
    const ValueStreamCallback& value_callback,
    const SubkeyStreamCallback& subkey_callback, DWORD max_data_size) {
  if (out_info) {
    out_info->info = {};
    out_info->info_valid = false;
  }
  return Dispatch(
      node,
      [&](const VirtualRegistryData& data) {
        return registry_backend::virtual_store::EnumKeyStreaming(
            data, node, include_values, include_data, include_subkeys,
            out_info, value_callback, subkey_callback, max_data_size);
      },
      [&] {
        return registry_backend::offline::EnumKeyStreaming(
            node, include_values, include_data, include_subkeys, out_info,
            value_callback, subkey_callback, max_data_size);
      },
      [&] {
        return registry_backend::live::EnumKeyStreaming(
            node, include_values, include_data, include_subkeys, out_info,
            value_callback, subkey_callback, max_data_size);
      });
}

bool RegistryStore::QueryValue(const RegistryNode& node,
                               const std::wstring& value_name,
                               ValueEntry* out) {
  if (!out) {
    return false;
  }
  return Dispatch(
      node,
      [&](const VirtualRegistryData& data) {
        return registry_backend::virtual_store::QueryValue(data, node,
                                                           value_name, out);
      },
      [&] {
        return registry_backend::offline::QueryValue(node, value_name, out);
      },
      [&] {
        return registry_backend::live::QueryValue(node, value_name, out);
      });
}

bool RegistryStore::CreateKey(const RegistryNode& node,
                              const std::wstring& name) {
  if (name.empty()) {
    return false;
  }
  return Dispatch(
      node,
      [&](VirtualRegistryData& data) {
        return registry_backend::virtual_store::CreateKey(data, node, name);
      },
      [&] { return registry_backend::offline::CreateKey(node, name); },
      [&] { return registry_backend::live::CreateKey(node, name); });
}

bool RegistryStore::DeleteKey(const RegistryNode& node) {
  if (node.subkey.empty()) {
    return false;
  }
  return Dispatch(
      node,
      [&](VirtualRegistryData& data) {
        return registry_backend::virtual_store::DeleteKey(data, node);
      },
      [&] { return registry_backend::offline::DeleteKey(node); },
      [&] { return registry_backend::live::DeleteKey(node); });
}

bool RegistryStore::RenameKey(const RegistryNode& node,
                              const std::wstring& new_name) {
  if (node.subkey.empty() || new_name.empty()) {
    return false;
  }
  return Dispatch(
      node,
      [&](VirtualRegistryData& data) {
        return registry_backend::virtual_store::RenameKey(data, node,
                                                          new_name);
      },
      [&] { return registry_backend::offline::RenameKey(node, new_name); },
      [&] { return registry_backend::live::RenameKey(node, new_name); });
}

bool RegistryStore::DeleteValue(const RegistryNode& node,
                                const std::wstring& value_name) {
  return Dispatch(
      node,
      [&](VirtualRegistryData& data) {
        return registry_backend::virtual_store::DeleteValue(data, node,
                                                            value_name);
      },
      [&] {
        return registry_backend::offline::DeleteValue(node, value_name);
      },
      [&] { return registry_backend::live::DeleteValue(node, value_name); });
}

bool RegistryStore::SetValue(const RegistryNode& node,
                             const std::wstring& value_name, DWORD type,
                             const std::vector<BYTE>& data) {
  return Dispatch(
      node,
      [&](VirtualRegistryData& virtual_data) {
        return registry_backend::virtual_store::SetValue(
            virtual_data, node, value_name, type, data);
      },
      [&] {
        return registry_backend::offline::SetValue(node, value_name, type,
                                                   data);
      },
      [&] {
        return registry_backend::live::SetValue(node, value_name, type, data);
      });
}

bool RegistryStore::RenameValue(const RegistryNode& node,
                                const std::wstring& old_name,
                                const std::wstring& new_name) {
  if (new_name.empty()) {
    return false;
  }
  return Dispatch(
      node,
      [&](VirtualRegistryData& data) {
        return registry_backend::virtual_store::RenameValue(
            data, node, old_name, new_name);
      },
      [&] {
        return registry_backend::offline::RenameValue(node, old_name,
                                                      new_name);
      },
      [&] {
        return registry_backend::live::RenameValue(node, old_name, new_name);
      });
}

} // namespace regkit
