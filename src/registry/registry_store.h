// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_value.h"
#include "registry/virtual_registry.h"
#include "win32/registry_native.h"
#include "win32/registry_view.h"

#include <win32ui/windows_config.h>

#include <windows.h>

#include <cstdint>
#include <functional>
#include <memory>
#include <string>
#include <vector>

namespace regkit
{

struct RegistryNode
{
    HKEY root = nullptr;
    std::wstring subkey;
    std::wstring root_name;
    REGSAM view = 0;
    bool children_loaded = false;
    bool simulated = false;
    signed char has_children = -1;
    int icon = -1;
};

enum class RegistryRootGroup
{
    kStandard,
    kReal,
};

struct RegistryRootEntry
{
    HKEY root = nullptr;
    std::wstring display_name;
    std::wstring path_name;
    std::wstring subkey_prefix;
    RegistryRootGroup group = RegistryRootGroup::kStandard;
    REGSAM view = 0;
};

// 0 uses native view, local 32 bit roots set KEY_WOW64_32KEY inherited by all children
inline REGSAM ViewOf(const RegistryNode& node)
{
    return node.view ? node.view : win32::kDefaultRegistryView;
}

struct ValueInfo
{
    std::wstring name;
    DWORD type = 0;
    DWORD data_size = 0;
};

struct EnumerationScratch
{
    std::wstring value_name;
    std::wstring subkey_name;
    std::vector<BYTE> value_data;
};

struct KeyInfo
{
    DWORD subkey_count = 0;
    DWORD value_count = 0;
    FILETIME last_write = {};
};

enum class ClassSource : unsigned char
{
    kNone,
    kUser,
    kMachine,
};

struct KeyInspection
{
    bool link = false;
    bool denied = false;
    bool is_volatile = false;
    ClassSource class_source = ClassSource::kNone;
    bool info_valid = false;
    KeyInfo info;
};

struct KeyDetails
{
    KeyInfo info;
    DWORD max_subkey_name = 0;
    DWORD max_class = 0;
    DWORD max_value_name = 0;
    DWORD max_value_data = 0;
    std::wstring class_name;
    util::NativeKeyInfo native;
};

struct KeyCreateOptions
{
    std::wstring class_name;
    bool is_volatile = false;
};

class RegistryStore
{
  public:
    static std::vector<RegistryRootEntry> DefaultRoots(bool include_extra = false);
    static bool HasSubKeys(const RegistryNode& node);
    // counts key creates, deletes, renames and access changes so other views can detect stale data
    static uint64_t KeyRevision() noexcept;
    static void NoteKeyChange() noexcept;
    static std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node, bool sorted = true);
    using ValueStreamCallback = std::function<bool(const ValueInfo& info, const BYTE* data, DWORD data_size)>;
    using SubkeyStreamCallback = std::function<bool(const std::wstring& name)>;
    struct KeyEnumResult
    {
        KeyInfo info;
        bool info_valid = false;
        bool want_options = false;
        LONG error = ERROR_SUCCESS;
        KeyCreateOptions options;
    };
    static bool EnumKeyStreaming(const RegistryNode& node, bool include_values, bool include_data, bool include_subkeys, KeyEnumResult* out_info, const ValueStreamCallback& value_callback, const SubkeyStreamCallback& subkey_callback, DWORD max_data_size = MAXDWORD, EnumerationScratch* scratch = nullptr, bool ordered = true, bool open_link = false);
    static bool IsOfflineRoot(HKEY root);
    static bool QueryValue(const RegistryNode& node, const std::wstring& value_name, RegistryValue* out);
    static bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info);
    static bool KeyExists(const RegistryNode& node);
    static KeyInspection InspectKey(const RegistryNode& node, bool want_info, bool want_source = false);
    static bool QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target, bool* denied = nullptr);
    static bool IsBrokenLink(const RegistryNode& node, std::wstring* target = nullptr);
    static bool OpenOfflineHive(const std::wstring& path, HKEY* root, std::wstring* error);
    static bool SaveOfflineHive(HKEY root, const std::wstring& path, std::wstring* error);
    static bool CloseOfflineHive(HKEY root, std::wstring* error);
    static void AddOfflineRoot(HKEY root);
    static void RemoveOfflineRoot(HKEY root);
    static HKEY RegisterVirtualRoot(const std::wstring& root_name, const std::shared_ptr<VirtualRegistryData>& data);
    static void UnregisterVirtualRoot(HKEY root);
    // independent registration of the same data, remains valid after releasing the first
    static HKEY DuplicateVirtualRoot(HKEY root);
    static bool IsVirtualRoot(HKEY root);
    static bool GetVirtualRootName(HKEY root, std::wstring* root_name);
    static bool QueryKeyDetails(const RegistryNode& node, KeyDetails* details, bool open_link = false);
    static std::wstring OfflineControlSet(const RegistryNode& node);
    static LONG SetKeyControlFlags(const RegistryNode& node, ULONG flags);
    static bool CreateKey(const RegistryNode& node, const std::wstring& name, const KeyCreateOptions& options = {}, bool* created_volatile = nullptr);
    static bool CreateKeyLink(const RegistryNode& node, const std::wstring& name, const std::wstring& nt_target, const KeyCreateOptions& options = {});
    static bool ReadKeyLink(const RegistryNode& node, std::wstring* target);
    static bool ReadKeySecurity(const RegistryNode& node, SECURITY_INFORMATION* parts, std::vector<BYTE>* descriptor, LONG* error = nullptr);
    static bool WriteKeySecurity(const RegistryNode& node, SECURITY_INFORMATION parts, const std::vector<BYTE>& descriptor, const FILETIME* last_write = nullptr);
    static bool DeleteKey(const RegistryNode& node);
    static bool RenameKey(const RegistryNode& node, const std::wstring& new_name);
    static bool DeleteValue(const RegistryNode& node, const std::wstring& value_name);
    static bool SetValue(const RegistryNode& node, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& data);
    static bool RenameValue(const RegistryNode& node, const std::wstring& old_name, const std::wstring& new_name, bool* both_names_left = nullptr);
};

}
