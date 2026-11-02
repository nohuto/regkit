// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/key_handle_snapshot.h"

#include "win32/process_rights.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system_api.h>

#include <tlhelp32.h>
#include <winternl.h>

#include <algorithm>
#include <cstddef>
#include <unordered_map>

namespace regkit::win32
{

namespace
{

constexpr ULONG kSystemExtendedHandleInformation = 64;
constexpr ULONG kObjectNameInformation = 1;
constexpr NTSTATUS kInfoLengthMismatch = static_cast<NTSTATUS>(0xC0000004L);
constexpr NTSTATUS kBufferOverflow = static_cast<NTSTATUS>(0x80000005L);
constexpr NTSTATUS kBufferTooSmall = static_cast<NTSTATUS>(0xC0000023L);
constexpr size_t kInitialTableBytes = 256 * 1024;
constexpr size_t kTableMarginBytes = 64 * 1024;
constexpr size_t kMaxTableBytes = 512 * 1024 * 1024;

struct HandleEntry
{
    PVOID object;
    ULONG_PTR process_id;
    ULONG_PTR handle;
    ULONG access;
    USHORT back_trace_index;
    USHORT type_index;
    ULONG attributes;
    ULONG reserved;
};

struct HandleTable
{
    ULONG_PTR count;
    ULONG_PTR reserved;
    HandleEntry entries[1];
};

using QuerySystemInformationFn = NTSTATUS(NTAPI*)(ULONG, PVOID, ULONG, PULONG);
using QueryObjectFn = NTSTATUS(NTAPI*)(HANDLE, ULONG, PVOID, ULONG, PULONG);

DWORD QueryHandleTable(std::vector<ULONG_PTR>* buffer)
{
    static const auto query = win32ui::ImportProc<QuerySystemInformationFn>(L"ntdll.dll", "NtQuerySystemInformation");
    if (!query)
    {
        return ERROR_PROC_NOT_FOUND;
    }
    size_t bytes = kInitialTableBytes;
    for (int attempt = 0; attempt < 8; ++attempt)
    {
        buffer->resize(bytes / sizeof(ULONG_PTR) + 1);
        ULONG needed = 0;
        const NTSTATUS status = query(kSystemExtendedHandleInformation, buffer->data(), static_cast<ULONG>(buffer->size() * sizeof(ULONG_PTR)), &needed);
        if (status >= 0)
        {
            return ERROR_SUCCESS;
        }
        if (status != kInfoLengthMismatch)
        {
            return win32ui::Win32FromNtStatus(status);
        }
        bytes = needed > bytes ? needed + kTableMarginBytes : bytes * 2;
        if (bytes > kMaxTableBytes)
        {
            return ERROR_NOT_ENOUGH_MEMORY;
        }
    }
    return ERROR_INSUFFICIENT_BUFFER;
}

void QueryName(QueryObjectFn query, HANDLE handle, std::vector<ULONG_PTR>* buffer, std::wstring* name)
{
    for (int attempt = 0; attempt < 4; ++attempt)
    {
        const ULONG bytes = static_cast<ULONG>(buffer->size() * sizeof(ULONG_PTR));
        ULONG needed = 0;
        const NTSTATUS status = query(handle, kObjectNameInformation, buffer->data(), bytes, &needed);
        if (status >= 0)
        {
            const auto* text = reinterpret_cast<const UNICODE_STRING*>(buffer->data());
            if (text->Buffer)
            {
                name->assign(text->Buffer, text->Length / sizeof(wchar_t));
            }
            return;
        }
        if ((status != kInfoLengthMismatch && status != kBufferOverflow && status != kBufferTooSmall) ||
            needed <= bytes)
        {
            return;
        }
        buffer->resize(needed / sizeof(ULONG_PTR) + 1);
    }
}

std::unordered_map<ULONG_PTR, std::wstring> ProcessNames()
{
    std::unordered_map<ULONG_PTR, std::wstring> names;
    win32ui::UniqueHandle snapshot(CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0));
    PROCESSENTRY32W entry = {};
    entry.dwSize = sizeof(entry);
    if (snapshot && Process32FirstW(snapshot.get(), &entry))
    {
        do
        {
            names.emplace(entry.th32ProcessID, entry.szExeFile);
        } while (Process32NextW(snapshot.get(), &entry));
    }
    return names;
}

DWORD LoadKeyEntries(std::vector<HandleEntry>* keys)
{
    win32ui::UniqueHKey probe;
    const LSTATUS opened = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE", 0, KEY_QUERY_VALUE, probe.put());
    if (opened != ERROR_SUCCESS)
    {
        return static_cast<DWORD>(opened);
    }
    std::vector<ULONG_PTR> buffer;
    const DWORD error = QueryHandleTable(&buffer);
    if (error != ERROR_SUCCESS)
    {
        return error;
    }
    const auto* table = reinterpret_cast<const HandleTable*>(buffer.data());
    const size_t capacity = (buffer.size() * sizeof(ULONG_PTR) - offsetof(HandleTable, entries)) / sizeof(HandleEntry);
    const HandleEntry* begin = table->entries;
    const HandleEntry* end = begin + std::min<size_t>(table->count, capacity);
    const ULONG_PTR self = GetCurrentProcessId();
    const ULONG_PTR probe_handle = reinterpret_cast<ULONG_PTR>(probe.get());
    const HandleEntry* probe_entry = std::find_if(
        begin,
        end,
        [&](const HandleEntry& entry) { return entry.process_id == self && entry.handle == probe_handle; }
    );
    if (probe_entry == end)
    {
        return ERROR_NOT_SUPPORTED;
    }
    for (const HandleEntry* entry = begin; entry != end; ++entry)
    {
        if (entry->type_index == probe_entry->type_index && entry != probe_entry)
        {
            keys->push_back(*entry);
        }
    }
    return ERROR_SUCCESS;
}

}

KeyHandleSnapshot SnapshotKeyHandles(const std::atomic_bool& cancel)
{
    KeyHandleSnapshot snapshot;
    static const auto query_object = win32ui::ImportProc<QueryObjectFn>(L"ntdll.dll", "NtQueryObject");
    if (!query_object)
    {
        snapshot.error = ERROR_PROC_NOT_FOUND;
        return snapshot;
    }
    const util::PrivilegeScope debug({SE_DEBUG_NAME});
    std::vector<HandleEntry> keys;
    snapshot.error = LoadKeyEntries(&keys);
    if (snapshot.error != ERROR_SUCCESS)
    {
        return snapshot;
    }
    const ULONG_PTR self = GetCurrentProcessId();
    std::sort(keys.begin(), keys.end(), [](const HandleEntry& left, const HandleEntry& right) {
        return left.process_id != right.process_id ? left.process_id < right.process_id : left.handle < right.handle;
    });

    const std::unordered_map<ULONG_PTR, std::wstring> names = ProcessNames();
    std::vector<ULONG_PTR> name_buffer(1024 / sizeof(ULONG_PTR));
    snapshot.handles.reserve(keys.size());
    for (size_t first = 0; first < keys.size() && !cancel.load();)
    {
        const ULONG_PTR pid = keys[first].process_id;
        size_t last = first;
        while (last < keys.size() && keys[last].process_id == pid)
        {
            ++last;
        }
        win32ui::UniqueHandle process(pid == self ? nullptr : OpenProcess(PROCESS_DUP_HANDLE, FALSE, static_cast<DWORD>(pid)));
        const HANDLE source = pid == self ? GetCurrentProcess() : process.get();
        if (!source)
        {
            ++snapshot.inaccessible_processes;
            first = last;
            continue;
        }
        const auto process_name = names.find(pid);
        for (; first < last && !cancel.load(); ++first)
        {
            const HandleEntry& entry = keys[first];
            KeyHandle row;
            row.process_id = pid;
            row.handle = entry.handle;
            row.object = reinterpret_cast<ULONG_PTR>(entry.object);
            row.access = entry.access;
            row.attributes = entry.attributes;
            if (process_name != names.end())
            {
                row.process = process_name->second;
            }
            win32ui::UniqueHandle duplicate;
            if (DuplicateHandle(source, reinterpret_cast<HANDLE>(entry.handle), GetCurrentProcess(), duplicate.put(), 0, FALSE, DUPLICATE_SAME_ACCESS))
            {
                QueryName(query_object, duplicate.get(), &name_buffer, &row.name);
            }
            snapshot.handles.push_back(std::move(row));
        }
        first = last;
    }
    return snapshot;
}

DWORD CloseKeyHandles(const std::vector<KeyHandle>& targets, size_t* closed)
{
    *closed = 0;
    const util::PrivilegeScope debug({SE_DEBUG_NAME});
    std::vector<HandleEntry> keys;
    DWORD result = LoadKeyEntries(&keys);
    if (result != ERROR_SUCCESS)
    {
        return result;
    }
    const ULONG_PTR self = GetCurrentProcessId();
    for (const KeyHandle& target : targets)
    {
        const bool unchanged =
            target.process_id != self && std::any_of(keys.begin(), keys.end(), [&](const HandleEntry& entry) {
                return entry.process_id == target.process_id && entry.handle == target.handle &&
                       reinterpret_cast<ULONG_PTR>(entry.object) == target.object && entry.access == target.access;
            });
        if (!unchanged)
        {
            result = ERROR_INVALID_HANDLE;
            continue;
        }
        win32ui::UniqueHandle process(OpenProcess(PROCESS_DUP_HANDLE, FALSE, static_cast<DWORD>(target.process_id)));
        if (process && DuplicateHandle(process.get(), reinterpret_cast<HANDLE>(target.handle), nullptr, nullptr, 0, FALSE, DUPLICATE_CLOSE_SOURCE))
        {
            ++*closed;
        }
        else
        {
            result = GetLastError();
        }
    }
    return *closed == targets.size() ? ERROR_SUCCESS : result;
}

}
