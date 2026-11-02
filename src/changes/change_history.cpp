// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/change_history.h"

#include "records/escaped_fields.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/date_time.h>

#include <algorithm>

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
    const size_t excess = entries->size() - maximum;
    std::vector<uint64_t> stamps;
    stamps.reserve(entries->size());
    for (const HistoryEntry& entry : *entries)
    {
        stamps.push_back(entry.timestamp);
    }
    std::nth_element(stamps.begin(), stamps.begin() + static_cast<std::ptrdiff_t>(excess - 1), stamps.end());
    const uint64_t cutoff = stamps[excess - 1];
    size_t at_cutoff = excess - static_cast<size_t>(std::count_if(stamps.begin(), stamps.end(), [cutoff](uint64_t stamp) { return stamp < cutoff; }));
    auto kept = entries->begin();
    for (auto entry = entries->begin(); entry != entries->end(); ++entry)
    {
        if (entry->timestamp < cutoff)
        {
            continue;
        }
        if (entry->timestamp == cutoff && at_cutoff > 0)
        {
            --at_cutoff;
            continue;
        }
        if (kept != entry)
        {
            *kept = std::move(*entry);
        }
        ++kept;
    }
    entries->erase(kept, entries->end());
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
    // strictly increasing timestamps preserve batch order within one clock tick
    static uint64_t last = 0;
    last = std::max<uint64_t>(win32ui::FileTimeTicks(now), last + 1);
    entry->timestamp = last;
    entry->time_text = win32ui::FormatDateTime(local, true);
}

void DecodeRevert(const std::vector<std::wstring>& fields, HistoryEntry* entry)
{
    uint64_t kind = 0;
    uint64_t type = 0;
    RegistryValue value;
    if (!entry ||
        !record_fields::ParseUnsigned(fields[7], static_cast<uint64_t>(HistoryEntry::RevertKind::kDeleteKey), &kind) ||
        !record_fields::ParseUnsigned(fields[9], MAXDWORD, &type) || !win32ui::ParseHexBytes(fields[10], &value.data))
    {
        return;
    }
    value.name = fields[8];
    value.type = static_cast<DWORD>(type);
    entry->revert_kind = static_cast<HistoryEntry::RevertKind>(kind);
    entry->revert_value = std::move(value);
}

}

std::vector<HistoryEntry> ChangeHistory::Append(std::vector<HistoryEntry> entries, size_t maximum)
{
    for (HistoryEntry& entry : entries)
    {
        Stamp(&entry);
    }
    entries_.insert(entries_.end(), entries.begin(), entries.end());
    Trim(&entries_, maximum);
    return entries;
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

void ChangeHistory::Sort(int column, bool ascending, const util::SortKeys& earlier)
{
    std::stable_sort(entries_.begin(), entries_.end(), [&](const HistoryEntry& left, const HistoryEntry& right) {
        return util::CompareSorted(column, ascending, earlier, [&](int key) { return CompareEntry(left, right, key); }) < 0;
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

HistoryDocument ParseHistory(const std::wstring& content, std::wstring* error)
{
    HistoryDocument document;
    size_t line_number = 0;
    for (const std::wstring_view line : win32ui::LineViews(content))
    {
        ++line_number;
        if (line.empty())
        {
            continue;
        }
        auto fields = record_fields::DecodeRecord(line);
        HistoryEntry entry;
        if (fields.size() != 15 || !record_fields::ParseUnsigned(fields[0], UINT64_MAX, &entry.timestamp))
        {
            if (error)
            {
                *error = util::BadLineReason(line_number);
            }
            return {};
        }
        entry.time_text = std::move(fields[1]);
        entry.action = std::move(fields[2]);
        entry.old_data = std::move(fields[3]);
        entry.new_data = std::move(fields[4]);
        entry.key_path = std::move(fields[5]);
        entry.value_name = std::move(fields[6]);
        DecodeRevert(fields, &entry);
        entry.backup_mode = fields[11] == L"1";
        uint64_t kind = 0;
        uint64_t view = 0;
        if (record_fields::ParseUnsigned(fields[12], 3, &kind) && record_fields::ParseUnsigned(fields[14], MAXDWORD, &view))
        {
            entry.source_kind = static_cast<int>(kind);
            entry.source_name = std::move(fields[13]);
            entry.source_view = static_cast<REGSAM>(view);
        }
        document.entries.push_back(std::move(entry));
    }
    return document;
}

std::wstring SerializeHistoryEntry(const HistoryEntry& entry)
{
    std::wstring line;
    record_fields::AppendRecord(&line, {std::to_wstring(entry.timestamp), entry.time_text, entry.action, entry.old_data, entry.new_data, entry.key_path, entry.value_name, std::to_wstring(static_cast<int>(entry.revert_kind)), entry.revert_value.name, std::to_wstring(entry.revert_value.type), win32ui::HexBytes(entry.revert_value.data), entry.backup_mode ? L"1" : L"0", std::to_wstring(entry.source_kind), entry.source_name, std::to_wstring(entry.source_view)});
    return line;
}

bool WriteHistoryFile(const std::wstring& path, const std::vector<HistoryEntry>& entries)
{
    if (path.empty())
    {
        return false;
    }
    std::wstring content;
    for (const auto& entry : entries)
    {
        content += SerializeHistoryEntry(entry);
    }
    if (content.empty())
    {
        return DeleteFileW(path.c_str()) != 0 || GetLastError() == ERROR_FILE_NOT_FOUND;
    }
    return win32ui::WriteTextFileAtomic(path.c_str(), content) == ERROR_SUCCESS;
}

bool AppendHistoryFile(const std::wstring& path, const std::vector<HistoryEntry>& entries)
{
    std::wstring lines;
    for (const HistoryEntry& entry : entries)
    {
        lines += SerializeHistoryEntry(entry);
    }
    const std::string bytes = win32ui::WideToMultiByte(lines, CP_UTF8);
    win32ui::UniqueHandle file(path.empty() ? INVALID_HANDLE_VALUE : CreateFileW(path.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
    DWORD written = 0;
    return file && !bytes.empty() &&
           WriteFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr) &&
           written == bytes.size();
}
bool PrepareRevert(const HistoryEntry& entry, HistoryEntry* prepared)
{
    if (!prepared || entry.key_path.empty() || entry.revert_kind == HistoryEntry::RevertKind::kNone)
    {
        return false;
    }
    *prepared = entry;
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

}
