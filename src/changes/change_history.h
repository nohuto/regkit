// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_store.h"
#include "win32/text_transform.h"

#include <functional>
#include <string>
#include <vector>

namespace regkit
{

struct HistoryEntry
{
    enum class RevertKind
    {
        kNone,
        kSetValue,
        kDeleteValue,
        kDeleteKey,
    };

    uint64_t timestamp = 0;
    std::wstring time_text;
    std::wstring action;
    std::wstring old_data;
    std::wstring new_data;
    std::wstring key_path;
    std::wstring value_name;
    RevertKind revert_kind = RevertKind::kNone;
    RegistryValue revert_value;
    bool backup_mode = false;
    int source_kind = -1;
    std::wstring source_name;
    REGSAM source_view = 0;
};

namespace changes
{

struct HistoryDocument
{
    std::vector<HistoryEntry> entries;
};

using PathExists = std::function<bool(const std::wstring&)>;

class ChangeHistory
{
  public:
    std::vector<HistoryEntry> Append(std::vector<HistoryEntry> entries, size_t maximum);
    void Replace(std::vector<HistoryEntry> entries, size_t maximum);
    void Clear();
    void Sort(int column, bool ascending, const util::SortKeys& earlier = {});

    const std::vector<HistoryEntry>& entries() const noexcept;
    std::vector<HistoryEntry>& entries() noexcept;

  private:
    std::vector<HistoryEntry> entries_;
};

HistoryDocument ParseHistory(const std::wstring& content, std::wstring* error = nullptr);
std::wstring SerializeHistoryEntry(const HistoryEntry& entry);
bool AppendHistoryFile(const std::wstring& path, const std::vector<HistoryEntry>& entries);
bool WriteHistoryFile(const std::wstring& path, const std::vector<HistoryEntry>& entries);
bool PrepareRevert(const HistoryEntry& entry, HistoryEntry* prepared);
bool FindNearestExistingPath(const std::wstring& path, const PathExists& path_exists, std::wstring* nearest_path);

}
}
