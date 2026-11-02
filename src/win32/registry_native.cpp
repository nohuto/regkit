// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/registry_native.h"

#include "win32/process_rights.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system_api.h>

#include <winternl.h>

#include <algorithm>
#include <atomic>
#include <cstring>
#include <limits>
#include <vector>

namespace regkit::util
{
namespace
{

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif

#ifndef OBJ_OPENLINK
#define OBJ_OPENLINK 0x00000100L
#endif

using NtOpenKeyFn = NTSTATUS(NTAPI*)(PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES);
using NtOpenKeyExFn = NTSTATUS(NTAPI*)(PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES, ULONG);
using NtCreateKeyFn = NTSTATUS(NTAPI*)(PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES, ULONG, PUNICODE_STRING, ULONG, PULONG);
using NtRenameKeyFn = NTSTATUS(NTAPI*)(HANDLE, PUNICODE_STRING);
using NtDeleteKeyFn = NTSTATUS(NTAPI*)(HANDLE);
using NtQueryKeyFn = NTSTATUS(NTAPI*)(HANDLE, ULONG, PVOID, ULONG, PULONG);
using NtSetInformationKeyFn = NTSTATUS(NTAPI*)(HANDLE, ULONG, PVOID, ULONG);
using NtQueryValueKeyFn = NTSTATUS(NTAPI*)(HANDLE, PUNICODE_STRING, ULONG, PVOID, ULONG, PULONG);
using NtSetValueKeyFn = NTSTATUS(NTAPI*)(HANDLE, PUNICODE_STRING, ULONG, ULONG, PVOID, ULONG);
using NtDeleteValueKeyFn = NTSTATUS(NTAPI*)(HANDLE, PUNICODE_STRING);

constexpr REGSAM kViewFlags = KEY_WOW64_32KEY | KEY_WOW64_64KEY;
constexpr NTSTATUS kBufferOverflow = static_cast<NTSTATUS>(0x80000005L);
constexpr NTSTATUS kBufferTooSmall = static_cast<NTSTATUS>(0xC0000023L);
constexpr ULONG kKeyNameInformation = 3;
constexpr ULONG kKeyFlagsInformation = 5;
constexpr ULONG kKeyVirtualizationInformation = 6;
constexpr ULONG kKeyTrustInformation = 8;
constexpr ULONG kKeyLayerInformation = 9;
constexpr ULONG kKeyWriteTimeInformation = 0;
constexpr ULONG kKeyControlFlagsInformation = 2;
constexpr ULONG kKeyValuePartialInformation = 2;
constexpr ULONG kPartialHeader = 3 * sizeof(ULONG);

// bit 0 tags remote RPC handles, never pass to native calls, bit 1 tags HKCR handles
HANDLE RootHandle(HKEY key)
{
    const ULONG_PTR value = reinterpret_cast<ULONG_PTR>(key);
    return (value & 1) ? nullptr : reinterpret_cast<HANDLE>(value & ~static_cast<ULONG_PTR>(3));
}

bool CountedName(const std::wstring& text, UNICODE_STRING* name)
{
    // counted names preserve embedded NULs truncated by Win32 APIs
    if (text.size() * sizeof(wchar_t) > (std::numeric_limits<USHORT>::max)())
    {
        return false;
    }
    name->Buffer = const_cast<PWSTR>(text.c_str());
    name->Length = static_cast<USHORT>(text.size() * sizeof(wchar_t));
    name->MaximumLength = name->Length;
    return true;
}

std::atomic_bool g_backup_restore{false};

LONG OpenNative(HKEY parent, const std::wstring& path, REGSAM access, bool open_link, win32ui::UniqueHKey* key, ULONG options = 0)
{
    static const auto open_key = win32ui::ImportProc<NtOpenKeyFn>(L"ntdll.dll", "NtOpenKey");
    static const auto open_key_ex = win32ui::ImportProc<NtOpenKeyExFn>(L"ntdll.dll", "NtOpenKeyEx");
    UNICODE_STRING name = {};
    if (path.empty() || !CountedName(path, &name))
    {
        return ERROR_INVALID_PARAMETER;
    }
    options |= open_link ? REG_OPTION_OPEN_LINK : 0;
    if (options ? !open_key_ex : !open_key)
    {
        return ERROR_CALL_NOT_IMPLEMENTED;
    }
    OBJECT_ATTRIBUTES attributes = {};
    // use OBJ_OPENLINK to open the link key instead of following its target
    InitializeObjectAttributes(&attributes, &name, OBJ_CASE_INSENSITIVE | (open_link ? OBJ_OPENLINK : 0ul), parent ? RootHandle(parent) : nullptr, nullptr);
    HANDLE handle = nullptr;
    // remove Win32 registry view flags before the native call
    const NTSTATUS status = options ? open_key_ex(&handle, access & ~kViewFlags, &attributes, options) : open_key(&handle, access & ~kViewFlags, &attributes);
    if (NT_SUCCESS(status))
    {
        key->reset(reinterpret_cast<HKEY>(handle));
    }
    return win32ui::Win32FromNtStatus(status);
}

template <size_t N>
bool QueryKeyWords(HKEY key, ULONG info_class, ULONG (&words)[N])
{
    static const auto query = win32ui::ImportProc<NtQueryKeyFn>(L"ntdll.dll", "NtQueryKey");
    ULONG length = 0;
    return query && NT_SUCCESS(query(RootHandle(key), info_class, words, sizeof(words), &length));
}

std::optional<ULONG> KeyWord(HKEY key, ULONG info_class)
{
    ULONG word[1] = {};
    return QueryKeyWords(key, info_class, word) ? std::optional<ULONG>(word[0]) : std::nullopt;
}

}

win32ui::UniqueHKey OpenNativeRegistryKey(const std::wstring& path, REGSAM access, bool open_link, LONG* error)
{
    win32ui::UniqueHKey key;
    LONG result = OpenNative(nullptr, path, access, open_link, &key);
    if (result == ERROR_ACCESS_DENIED && g_backup_restore)
    {
        result = OpenNative(nullptr, path, access, open_link, &key, REG_OPTION_BACKUP_RESTORE);
    }
    if (error)
    {
        *error = result;
    }
    return key;
}

win32ui::UniqueHKey OpenNativeRegistryRoot()
{
    return OpenNativeRegistryKey(L"\\REGISTRY", KEY_READ);
}

LONG OpenRegistryPath(HKEY root, const std::wstring& subkey, REGSAM access, bool open_link, win32ui::UniqueHKey* key)
{
    key->reset();
    constexpr size_t kMaxLevelsPerOpen = 32;
    size_t levels = 0;
    for (size_t pos = subkey.find(L'\\'); pos != std::wstring::npos; pos = subkey.find(L'\\', pos + 1))
    {
        if (++levels == kMaxLevelsPerOpen)
        {
            win32ui::UniqueHKey middle;
            const LONG result = OpenRegistryPath(root, subkey.substr(0, pos), MAXIMUM_ALLOWED | (access & kViewFlags), false, &middle);
            return result == ERROR_SUCCESS ? OpenRegistryPath(middle.get(), subkey.substr(pos + 1), access, open_link, key) : result;
        }
    }
    const bool merged = root == HKEY_CLASSES_ROOT;
    root = MapCurrentUserRoot(root);
    const size_t null_char = subkey.find(L'\0');
    LONG result = ERROR_SUCCESS;
    if (null_char == std::wstring::npos)
    {
        result = RegOpenKeyExW(root, subkey.empty() ? nullptr : subkey.c_str(), open_link ? REG_OPTION_OPEN_LINK : 0, access, key->put());
    }
    else
    {
        const size_t split = subkey.rfind(L'\\', null_char);
        const std::wstring prefix = split == std::wstring::npos ? std::wstring() : subkey.substr(0, split);
        win32ui::UniqueHKey parent;
        result = RegOpenKeyExW(root, prefix.c_str(), 0, MAXIMUM_ALLOWED | (access & kViewFlags), parent.put());
        if (result == ERROR_SUCCESS)
        {
            result = OpenNative(parent.get(), split == std::wstring::npos ? subkey : subkey.substr(split + 1), access, open_link, key);
        }
    }

    const ULONG_PTR value = reinterpret_cast<ULONG_PTR>(root);
    const bool predefined = (value & ~static_cast<ULONG_PTR>(0xFF)) == reinterpret_cast<ULONG_PTR>(HKEY_CLASSES_ROOT);
    if (result == ERROR_ACCESS_DENIED && g_backup_restore && !merged && !subkey.empty() && (predefined || !(value & 3)))
    {
        win32ui::UniqueHKey base;
        if (predefined)
        {
            result = RegOpenKeyExW(root, nullptr, 0, KEY_QUERY_VALUE | (access & kViewFlags), base.put());
        }
        if (result == ERROR_SUCCESS || !predefined)
        {
            result = OpenNative(predefined ? base.get() : root, subkey, access, open_link, key, REG_OPTION_BACKUP_RESTORE);
        }
    }
    return result;
}

void SetBackupRestoreMode(bool enable)
{
    g_backup_restore = enable;
}

LONG CreateRegistryKey(HKEY parent, const std::wstring& name, REGSAM access, DWORD options, win32ui::UniqueHKey* key, DWORD* disposition, const std::wstring& class_name)
{
    key->reset();
    if (name.find(L'\0') == std::wstring::npos)
    {
        const LPWSTR key_class = class_name.empty() ? nullptr : const_cast<LPWSTR>(class_name.c_str());
        LONG result = RegCreateKeyExW(parent, name.c_str(), 0, key_class, options, access, nullptr, key->put(), disposition);
        if (result == ERROR_ACCESS_DENIED && g_backup_restore)
        {
            result = RegCreateKeyExW(parent, name.c_str(), 0, key_class, options | REG_OPTION_BACKUP_RESTORE, access, nullptr, key->put(), disposition);
        }
        return result;
    }
    // use NtCreateKey to keep the full name when it contains embedded nulls
    static const auto create_key = win32ui::ImportProc<NtCreateKeyFn>(L"ntdll.dll", "NtCreateKey");
    UNICODE_STRING counted = {};
    UNICODE_STRING counted_class = {};
    if (!create_key || !CountedName(name, &counted) || !CountedName(class_name, &counted_class))
    {
        return ERROR_INVALID_PARAMETER;
    }
    OBJECT_ATTRIBUTES attributes = {};
    InitializeObjectAttributes(&attributes, &counted, OBJ_CASE_INSENSITIVE | ((options & REG_OPTION_CREATE_LINK) ? OBJ_OPENLINK : 0ul), RootHandle(parent), nullptr);
    HANDLE handle = nullptr;
    ULONG created = 0;
    const NTSTATUS status = create_key(&handle, access & ~kViewFlags, &attributes, 0, class_name.empty() ? nullptr : &counted_class, options, &created);
    if (NT_SUCCESS(status))
    {
        key->reset(reinterpret_cast<HKEY>(handle));
        if (disposition)
        {
            *disposition = created;
        }
    }
    return win32ui::Win32FromNtStatus(status);
}

LONG RenameRegistryKey(HKEY parent, const std::wstring& old_name, const std::wstring& new_name)
{
    // NtRenameKey renames the link itself, RegRenameKey follows targets and truncates names at embedded NULs
    static const auto rename_key = win32ui::ImportProc<NtRenameKeyFn>(L"ntdll.dll", "NtRenameKey");
    win32ui::UniqueHKey key;
    LONG result = OpenRegistryPath(parent, old_name, KEY_WRITE, true, &key);
    if (result == ERROR_SUCCESS && !RootHandle(key.get()))
    {
        return new_name.find(L'\0') == std::wstring::npos ? RegRenameKey(key.get(), nullptr, new_name.c_str()) : ERROR_INVALID_PARAMETER;
    }
    UNICODE_STRING counted = {};
    if (result == ERROR_SUCCESS && (!rename_key || !CountedName(new_name, &counted)))
    {
        result = ERROR_INVALID_PARAMETER;
    }
    return result == ERROR_SUCCESS ? win32ui::Win32FromNtStatus(rename_key(RootHandle(key.get()), &counted)) : result;
}

LONG DeleteRegistryTree(HKEY key)
{
    static const auto delete_key = win32ui::ImportProc<NtDeleteKeyFn>(L"ntdll.dll", "NtDeleteKey");
    if (!key || !delete_key)
    {
        return ERROR_INVALID_PARAMETER;
    }
    wchar_t name[256] = {};
    while (true)
    {
        DWORD length = static_cast<DWORD>(_countof(name));
        // each removal moves the next child to index zero
        const LONG result = RegEnumKeyExW(key, 0, name, &length, nullptr, nullptr, nullptr, nullptr);
        if (result == ERROR_NO_MORE_ITEMS)
        {
            break;
        }
        win32ui::UniqueHKey child;
        // open link keys directly so deletion never goes into their targets
        LONG removed = result == ERROR_SUCCESS
                           ? OpenRegistryPath(key, std::wstring(name, length), DELETE | KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE, true, &child)
                           : result;
        if (removed == ERROR_SUCCESS)
        {
            removed = DeleteRegistryTree(child.get());
        }
        if (removed != ERROR_SUCCESS)
        {
            return removed;
        }
    }
    // empty subkeys delete the key behind a remote handle
    return RootHandle(key) ? win32ui::Win32FromNtStatus(delete_key(RootHandle(key))) : RegDeleteKeyExW(key, L"", 0, 0);
}

bool DeleteNativeRegistryKey(HKEY key)
{
    static const auto delete_key = win32ui::ImportProc<NtDeleteKeyFn>(L"ntdll.dll", "NtDeleteKey");
    return key && delete_key && NT_SUCCESS(delete_key(RootHandle(key)));
}

std::optional<ULONG> QueryKeyFlags(HKEY key)
{
    ULONG flags[3] = {};
    return QueryKeyWords(key, kKeyFlagsInformation, flags) ? std::optional<ULONG>(flags[1]) : std::nullopt;
}

std::wstring QueryKeyName(HKEY key)
{
    static const auto query = win32ui::ImportProc<NtQueryKeyFn>(L"ntdll.dll", "NtQueryKey");
    std::vector<ULONG> buffer(130);
    for (int attempt = 0; query && attempt < 3; ++attempt)
    {
        ULONG needed = 0;
        const NTSTATUS status = query(RootHandle(key), kKeyNameInformation, buffer.data(), static_cast<ULONG>(buffer.size() * sizeof(ULONG)), &needed);
        if (NT_SUCCESS(status))
        {
            return std::wstring(reinterpret_cast<const wchar_t*>(buffer.data() + 1), buffer[0] / sizeof(wchar_t));
        }
        if (status != kBufferOverflow && status != kBufferTooSmall)
        {
            break;
        }
        buffer.resize(needed / sizeof(ULONG) + 1);
    }
    return {};
}

NativeKeyInfo QueryNativeKeyInfo(HKEY key)
{
    NativeKeyInfo info;
    info.native_name = QueryKeyName(key);
    ULONG flags[3] = {};
    if (QueryKeyWords(key, kKeyFlagsInformation, flags))
    {
        info.key_flags = flags[1];
        info.control_flags = flags[2];
    }
    info.virtualization = KeyWord(key, kKeyVirtualizationInformation);
    info.trust = KeyWord(key, kKeyTrustInformation);
    info.layer = KeyWord(key, kKeyLayerInformation);
    return info;
}

LONG SetKeyLastWriteTime(HKEY key, const FILETIME& time)
{
    static const auto set_information = win32ui::ImportProc<NtSetInformationKeyFn>(L"ntdll.dll", "NtSetInformationKey");
    LARGE_INTEGER value = {};
    value.LowPart = time.dwLowDateTime;
    value.HighPart = static_cast<LONG>(time.dwHighDateTime);
    return set_information ? win32ui::Win32FromNtStatus(set_information(RootHandle(key), kKeyWriteTimeInformation, &value, sizeof(value))) : ERROR_CALL_NOT_IMPLEMENTED;
}

LONG SetKeyControlFlags(HKEY key, ULONG flags)
{
    static const auto set_information = win32ui::ImportProc<NtSetInformationKeyFn>(L"ntdll.dll", "NtSetInformationKey");
    return set_information ? win32ui::Win32FromNtStatus(set_information(RootHandle(key), kKeyControlFlagsInformation, &flags, sizeof(flags))) : ERROR_CALL_NOT_IMPLEMENTED;
}

LONG QueryValueCounted(HKEY key, const std::wstring& name, DWORD* type, BYTE* data, DWORD* size)
{
    // RegQueryValueExW behavior for names requiring counted strings
    static const auto query_value = win32ui::ImportProc<NtQueryValueKeyFn>(L"ntdll.dll", "NtQueryValueKey");
    UNICODE_STRING counted = {};
    if (!query_value || !CountedName(name, &counted) || !size || (data ? *size : 0) > MAXDWORD - kPartialHeader)
    {
        return ERROR_INVALID_PARAMETER;
    }
    std::vector<BYTE> buffer(kPartialHeader + (data ? *size : 0));
    ULONG needed = 0;
    const NTSTATUS status = query_value(RootHandle(key), &counted, kKeyValuePartialInformation, buffer.data(), static_cast<ULONG>(buffer.size()), &needed);
    const bool header = NT_SUCCESS(status) || status == kBufferOverflow;
    if (!header && status != kBufferTooSmall)
    {
        return win32ui::Win32FromNtStatus(status);
    }
    const auto* words = reinterpret_cast<const ULONG*>(buffer.data());
    if (type && header)
    {
        *type = words[1];
    }
    *size = header ? words[2] : (std::max)(needed, kPartialHeader) - kPartialHeader;
    if (!NT_SUCCESS(status))
    {
        return data ? ERROR_MORE_DATA : ERROR_SUCCESS;
    }
    if (data)
    {
        std::memcpy(data, buffer.data() + kPartialHeader, *size);
    }
    return ERROR_SUCCESS;
}

LONG SetValueCounted(HKEY key, const std::wstring& name, DWORD type, const BYTE* data, DWORD size)
{
    static const auto set_value = win32ui::ImportProc<NtSetValueKeyFn>(L"ntdll.dll", "NtSetValueKey");
    UNICODE_STRING counted = {};
    if (!set_value || !CountedName(name, &counted))
    {
        return ERROR_INVALID_PARAMETER;
    }
    return win32ui::Win32FromNtStatus(set_value(RootHandle(key), &counted, 0, type, const_cast<BYTE*>(data), size));
}

LONG DeleteValueCounted(HKEY key, const std::wstring& name)
{
    static const auto delete_value = win32ui::ImportProc<NtDeleteValueKeyFn>(L"ntdll.dll", "NtDeleteValueKey");
    UNICODE_STRING counted = {};
    if (!delete_value || !CountedName(name, &counted))
    {
        return ERROR_INVALID_PARAMETER;
    }
    return win32ui::Win32FromNtStatus(delete_value(RootHandle(key), &counted));
}
}
