// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "registry/registry_provider.h"

#include <functional>
#include <string>
#include <vector>

namespace regkit {

struct HistoryEntry {
  enum class RevertKind {
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
  ValueEntry revert_value;
};

namespace changes {

struct HistoryDocument {
  static constexpr int kCurrentVersion = 2;
  int source_version = 1;
  std::vector<HistoryEntry> entries;
};

using QueryValue =
    std::function<bool(const std::wstring&, const std::wstring&, ValueEntry*)>;
using PathExists = std::function<bool(const std::wstring&)>;

class ChangeHistory {
public:
  HistoryEntry Append(HistoryEntry entry, size_t maximum);
  void Replace(std::vector<HistoryEntry> entries, size_t maximum);
  void Clear();
  void Sort(int column, bool ascending);

  const std::vector<HistoryEntry>& entries() const noexcept;
  std::vector<HistoryEntry>& entries() noexcept;

private:
  std::vector<HistoryEntry> entries_;
};

HistoryDocument ParseHistory(const std::wstring& content);
std::wstring SerializeHistoryEntry(const HistoryEntry& entry);
bool AppendHistoryFile(const std::wstring& path, const HistoryEntry& entry);
bool PrepareRevert(const HistoryEntry& entry, const QueryValue& query_value,
                   HistoryEntry* prepared);
bool FindNearestExistingPath(const std::wstring& path,
                             const PathExists& path_exists,
                             std::wstring* nearest_path);

} // namespace changes
} // namespace regkit
