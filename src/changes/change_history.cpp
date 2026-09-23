// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/change_history.h"

#include "records/escaped_fields.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/file_text.h"
#include "win32/handle_owner.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <cerrno>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <limits>

namespace regkit::changes
{
namespace
{

int CompareEntry(const HistoryEntry& left, const HistoryEntry& right, int column)
{
    switch (column)
    {
    case 1:
        return util::CompareListText(left.action, right.action);
    case 2:
        return util::CompareListText(left.old_data, right.old_data);
    case 3:
        return util::CompareListText(left.new_data, right.new_data);
    default:
        return left.timestamp < right.timestamp ? -1 : (left.timestamp > right.timestamp ? 1 : 0);
    }
}

void Trim(std::vector<HistoryEntry>* entries, size_t maximum)
{
    if (!entries)
    {
        return;
    }
    if (maximum == 0)
    {
        entries->clear();
        return;
    }
    if (entries->size() <= maximum)
    {
        return;
    }
    auto cut = entries->end() - static_cast<std::ptrdiff_t>(maximum);
    std::nth_element(entries->begin(), cut, entries->end(), [](const HistoryEntry& left, const HistoryEntry& right) {
        return left.timestamp < right.timestamp;
    });
    entries->erase(entries->begin(), cut);
}

void Stamp(HistoryEntry* entry)
{
    if (!entry || (entry->timestamp != 0 && !entry->time_text.empty()))
    {
        return;
    }
    SYSTEMTIME local = {};
    GetLocalTime(&local);
    FILETIME now = {};
    GetSystemTimeAsFileTime(&now);
    ULARGE_INTEGER value = {};
    value.LowPart = now.dwLowDateTime;
    value.HighPart = now.dwHighDateTime;
    entry->timestamp = value.QuadPart;
    entry->time_text = util::FormatLocalTime(local, true);
}

void DecodeRevert(const std::vector<std::wstring>& fields, HistoryEntry* entry)
{
    uint64_t kind = 0;
    uint64_t type = 0;
    ValueEntry value;
    if (!entry || fields.size() < 11 ||
        !record_fields::ParseUnsigned(fields[7], static_cast<uint64_t>(HistoryEntry::RevertKind::kDeleteKey), &kind) ||
        !record_fields::ParseUnsigned(fields[9], MAXDWORD, &type) || !value_format::ParseHex(fields[10], &value.data))
    {
        return;
    }
    value.name = fields[8];
    value.type = static_cast<DWORD>(type);
    entry->revert_kind = static_cast<HistoryEntry::RevertKind>(kind);
    entry->revert_value = std::move(value);
}

} // namespace

HistoryEntry ChangeHistory::Append(HistoryEntry entry, size_t maximum)
{
    Stamp(&entry);
    HistoryEntry appended = entry;
    entries_.push_back(std::move(entry));
    Trim(&entries_, maximum);
    return appended;
}

void ChangeHistory::Replace(std::vector<HistoryEntry> entries, size_t maximum)
{
    entries_ = std::move(entries);
    Trim(&entries_, maximum);
}

void ChangeHistory::Clear()
{
    entries_.clear();
}

void ChangeHistory::Sort(int column, bool ascending)
{
    std::stable_sort(entries_.begin(), entries_.end(), [column, ascending](const HistoryEntry& left, const HistoryEntry& right) {
        const int comparison = CompareEntry(left, right, column);
        return comparison != 0 && (ascending ? comparison < 0 : comparison > 0);
    });
}

const std::vector<HistoryEntry>& ChangeHistory::entries() const noexcept
{
    return entries_;
}

std::vector<HistoryEntry>& ChangeHistory::entries() noexcept
{
    return entries_;
}

HistoryDocument ParseHistory(const std::wstring& content)
{
    HistoryDocument document;
    for (const std::wstring_view line : record_fields::Lines(content))
    {
        if (line.empty())
        {
            continue;
        }
        auto fields = record_fields::DecodeRecord(line);
        HistoryEntry entry;
        if (fields.size() < 5 || !record_fields::ParseUnsigned(fields[0], UINT64_MAX, &entry.timestamp))
        {
            continue;
        }
        entry.time_text = std::move(fields[1]);
        entry.action = std::move(fields[2]);
        entry.old_data = std::move(fields[3]);
        entry.new_data = std::move(fields[4]);
        if (fields.size() >= 7)
        {
            entry.key_path = std::move(fields[5]);
            entry.value_name = std::move(fields[6]);
        }
        if (fields.size() >= 11)
        {
            document.source_version = HistoryDocument::kCurrentVersion;
            DecodeRevert(fields, &entry);
        }
        document.entries.push_back(std::move(entry));
    }
    return document;
}

std::wstring SerializeHistoryEntry(const HistoryEntry& entry)
{
    std::wstring line;
    record_fields::AppendRecord(&line, {std::to_wstring(entry.timestamp), entry.time_text, entry.action, entry.old_data, entry.new_data, entry.key_path, entry.value_name, std::to_wstring(static_cast<int>(entry.revert_kind)), entry.revert_value.name, std::to_wstring(entry.revert_value.type), util::ToHex(entry.revert_value.data)});
    return line;
}

bool WriteHistoryFile(const std::wstring& path, const std::vector<HistoryEntry>& entries)
{
    std::wstring content;
    for (const auto& entry : entries)
    {
        content += SerializeHistoryEntry(entry);
    }
    if (path.empty())
    {
        return false;
    }
    if (content.empty())
    {
        return DeleteFileW(path.c_str()) != 0 || GetLastError() == ERROR_FILE_NOT_FOUND;
    }
    return util::WriteTextFile(path, content, false);
}

bool AppendHistoryFile(const std::wstring& path, const HistoryEntry& entry)
{
    const std::string bytes = util::WideToUtf8(SerializeHistoryEntry(entry));
    util::UniqueHandle file(path.empty() ? INVALID_HANDLE_VALUE : CreateFileW(path.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    DWORD written = 0;
    return file && !bytes.empty() &&
           WriteFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) &&
           written == bytes.size();
}
bool PrepareRevert(const HistoryEntry& entry, const QueryValue& query_value, HistoryEntry* prepared)
{
    if (!prepared || entry.key_path.empty())
    {
        return false;
    }
    *prepared = entry;
    if (prepared->revert_kind != HistoryEntry::RevertKind::kNone)
    {
        return true;
    }

    auto suffix = [&](std::wstring_view prefix, std::wstring* value_name) {
        if (!util::StartsWithInsensitive(entry.action, prefix))
        {
            return false;
        }
        *value_name = entry.action.substr(prefix.size());
        if (*value_name == L"(Default)")
        {
            value_name->clear();
        }
        return true;
    };

    std::wstring value_name;
    ValueEntry current;
    if (suffix(L"Create value ", &value_name))
    {
        prepared->value_name = std::move(value_name);
        prepared->revert_kind = HistoryEntry::RevertKind::kDeleteValue;
        return true;
    }
    if (!suffix(L"Modify value ", &value_name) || entry.old_data.empty() || !query_value ||
        !query_value(entry.key_path, value_name, &current))
    {
        return false;
    }

    const DWORD type = value_format::NormalizeType(current.type);
    if (type != REG_DWORD && type != REG_DWORD_BIG_ENDIAN && type != REG_QWORD)
    {
        return false;
    }
    const wchar_t* start = entry.old_data.c_str();
    while (*start && iswspace(*start))
    {
        ++start;
    }
    if (start[0] != L'0' || (start[1] != L'x' && start[1] != L'X'))
    {
        return false;
    }
    errno = 0;
    wchar_t* end = nullptr;
    unsigned long long value = wcstoull(start + 2, &end, 16);
    if (end == start + 2 || errno == ERANGE)
    {
        return false;
    }
    while (*end && iswspace(*end))
    {
        ++end;
    }
    if ((*end != L'\0' && *end != L'(') ||
        ((type == REG_DWORD || type == REG_DWORD_BIG_ENDIAN) && value > std::numeric_limits<DWORD>::max()))
    {
        return false;
    }

    current.name = value_name;
    if (type == REG_QWORD)
    {
        current.data.resize(sizeof(value));
        memcpy(current.data.data(), &value, sizeof(value));
    }
    else
    {
        const DWORD dword =
            type == REG_DWORD_BIG_ENDIAN ? _byteswap_ulong(static_cast<DWORD>(value)) : static_cast<DWORD>(value);
        current.data.resize(sizeof(dword));
        memcpy(current.data.data(), &dword, sizeof(dword));
    }
    prepared->value_name = value_name;
    prepared->revert_value = std::move(current);
    prepared->revert_kind = HistoryEntry::RevertKind::kSetValue;
    return true;
}

bool FindNearestExistingPath(const std::wstring& path, const PathExists& path_exists, std::wstring* nearest_path)
{
    if (!nearest_path || !path_exists)
    {
        return false;
    }
    nearest_path->clear();
    std::wstring candidate = registry_path::Clean(path);
    while (!candidate.empty())
    {
        if (path_exists(candidate))
        {
            *nearest_path = std::move(candidate);
            return true;
        }
        const std::wstring parent = registry_path::Parent(candidate);
        if (parent.empty() || parent == candidate)
        {
            return false;
        }
        candidate = parent;
    }
    return false;
}

} // namespace regkit::changes
