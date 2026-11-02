// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/handle_owner.h>
#include <win32ui/windows_config.h>

#include <optional>
#include <string>

namespace regkit::util
{

inline constexpr ULONG kKeyFlagVolatile = 0x1;
inline constexpr ULONG kKeyFlagLink = 0x2;

struct NativeKeyInfo
{
    std::wstring native_name;
    std::optional<ULONG> key_flags;
    std::optional<ULONG> control_flags;
    std::optional<ULONG> virtualization;
    std::optional<ULONG> trust;
    std::optional<ULONG> layer;
};

// remote roots are RPC handles, never pass to native calls
inline bool IsLocalRoot(HKEY root) noexcept
{
    return root == HKEY_CLASSES_ROOT || root == HKEY_CURRENT_USER || root == HKEY_LOCAL_MACHINE || root == HKEY_USERS || root == HKEY_CURRENT_CONFIG;
}

win32ui::UniqueHKey OpenNativeRegistryKey(const std::wstring& path, REGSAM access, bool open_link = false, LONG* error = nullptr);
win32ui::UniqueHKey OpenNativeRegistryRoot();
LONG OpenRegistryPath(HKEY root, const std::wstring& subkey, REGSAM access, bool open_link, win32ui::UniqueHKey* key);
void SetBackupRestoreMode(bool enable);
LONG CreateRegistryKey(HKEY parent, const std::wstring& name, REGSAM access, DWORD options, win32ui::UniqueHKey* key, DWORD* disposition, const std::wstring& class_name = {});
LONG RenameRegistryKey(HKEY parent, const std::wstring& old_name, const std::wstring& new_name);
LONG DeleteRegistryTree(HKEY key);
bool DeleteNativeRegistryKey(HKEY key);
std::optional<ULONG> QueryKeyFlags(HKEY key);
std::wstring QueryKeyName(HKEY key);
NativeKeyInfo QueryNativeKeyInfo(HKEY key);
LONG SetKeyLastWriteTime(HKEY key, const FILETIME& time);
inline constexpr ULONG kKeyDontVirtualize = 0x2;
inline constexpr ULONG kKeyDontSilentFail = 0x4;
inline constexpr ULONG kKeyRecurseFlag = 0x8;
LONG SetKeyControlFlags(HKEY key, ULONG flags);
LONG QueryValueCounted(HKEY key, const std::wstring& name, DWORD* type, BYTE* data, DWORD* size);
LONG SetValueCounted(HKEY key, const std::wstring& name, DWORD type, const BYTE* data, DWORD size);
LONG DeleteValueCounted(HKEY key, const std::wstring& name);

}
