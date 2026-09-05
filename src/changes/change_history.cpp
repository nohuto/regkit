// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/change_history.h"

#include "records/escaped_fields.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <cerrno>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <limits>

namespace regkit::changes {
namespace {

int CompareText(const std::wstring& left, const std::wstring& right) {
  const int result = _wcsicmp(left.c_str(), right.c_str());
  return result < 0 ? -1 : (result > 0 ? 1 : 0);
}

int CompareEntry(const HistoryEntry& left, const HistoryEntry& right,
                 int column) {
  switch (column) {
  case 1:
    return CompareText(left.action, right.action);
  case 2:
    return CompareText(left.old_data, right.old_data);
  case 3:
    return CompareText(left.new_data, right.new_data);
  default:
    return left.timestamp < right.timestamp
               ? -1
               : (left.timestamp > right.timestamp ? 1 : 0);
  }
}

void Trim(std::vector<HistoryEntry>* entries, size_t maximum) {
  if (!entries) {
    return;
  }
  if (maximum == 0) {
    entries->clear();
    return;
  }
  if (entries->size() <= maximum) {
    return;
  }
  auto cut = entries->end() - static_cast<std::ptrdiff_t>(maximum);
  std::nth_element(
      entries->begin(), cut, entries->end(),
      [](const HistoryEntry& left, const HistoryEntry& right) {
        return left.timestamp < right.timestamp;
      });
  entries->erase(entries->begin(), cut);
}

void Stamp(HistoryEntry* entry) {
  if (!entry || (entry->timestamp != 0 && !entry->time_text.empty())) {
    return;
  }
  SYSTEMTIME local = {};
  GetLocalTime(&local);
  wchar_t text[64] = {};
  swprintf_s(text, L"%d/%d/%d %d:%02d:%02d", local.wMonth, local.wDay,
             local.wYear, local.wHour, local.wMinute, local.wSecond);
  FILETIME now = {};
  GetSystemTimeAsFileTime(&now);
  ULARGE_INTEGER value = {};
  value.LowPart = now.dwLowDateTime;
  value.HighPart = now.dwHighDateTime;
  entry->timestamp = value.QuadPart;
  entry->time_text = text;
}

void DecodeRevert(const std::vector<std::wstring>& fields,
                  HistoryEntry* entry) {
  if (!entry || fields.size() < 11) {
    return;
  }
  try {
    const int kind = std::stoi(fields[7]);
    if (kind < static_cast<int>(HistoryEntry::RevertKind::kNone) ||
        kind > static_cast<int>(HistoryEntry::RevertKind::kDeleteKey)) {
      return;
    }
    entry->revert_kind = static_cast<HistoryEntry::RevertKind>(kind);
    entry->revert_value.name = record_fields::Unescape(fields[8]);
    const unsigned long type = std::stoul(fields[9]);
    if (type > std::numeric_limits<DWORD>::max()) {
      entry->revert_kind = HistoryEntry::RevertKind::kNone;
      return;
    }
    entry->revert_value.type = static_cast<DWORD>(type);
    if (!value_format::ParseHex(record_fields::Unescape(fields[10]),
                                &entry->revert_value.data)) {
      entry->revert_kind = HistoryEntry::RevertKind::kNone;
      entry->revert_value = {};
    }
  } catch (...) {
    entry->revert_kind = HistoryEntry::RevertKind::kNone;
    entry->revert_value = {};
  }
}

} // namespace

HistoryEntry ChangeHistory::Append(HistoryEntry entry, size_t maximum) {
  Stamp(&entry);
  HistoryEntry appended = entry;
  entries_.push_back(std::move(entry));
  Trim(&entries_, maximum);
  return appended;
}

void ChangeHistory::Replace(std::vector<HistoryEntry> entries,
                            size_t maximum) {
  entries_ = std::move(entries);
  Trim(&entries_, maximum);
}

void ChangeHistory::Clear() {
  entries_.clear();
}

void ChangeHistory::Sort(int column, bool ascending) {
  std::stable_sort(
      entries_.begin(), entries_.end(),
      [column, ascending](const HistoryEntry& left,
                          const HistoryEntry& right) {
        const int comparison = CompareEntry(left, right, column);
        return comparison != 0 &&
               (ascending ? comparison < 0 : comparison > 0);
      });
}

const std::vector<HistoryEntry>& ChangeHistory::entries() const noexcept {
  return entries_;
}

std::vector<HistoryEntry>& ChangeHistory::entries() noexcept {
  return entries_;
}

HistoryDocument ParseHistory(const std::wstring& content) {
  HistoryDocument document;
  for (const std::wstring& line : record_fields::Lines(content)) {
    if (line.empty()) {
      continue;
    }
    const auto fields = record_fields::Split(line);
    if (fields.size() < 5) {
      continue;
    }
    HistoryEntry entry;
    try {
      entry.timestamp = std::stoull(fields[0]);
    } catch (...) {
      continue;
    }
    entry.time_text = record_fields::Unescape(fields[1]);
    entry.action = record_fields::Unescape(fields[2]);
    entry.old_data = record_fields::Unescape(fields[3]);
    entry.new_data = record_fields::Unescape(fields[4]);
    if (fields.size() >= 7) {
      entry.key_path = record_fields::Unescape(fields[5]);
      entry.value_name = record_fields::Unescape(fields[6]);
    }
    if (fields.size() >= 11) {
      document.source_version = HistoryDocument::kCurrentVersion;
      DecodeRevert(fields, &entry);
    }
    document.entries.push_back(std::move(entry));
  }
  return document;
}

std::wstring SerializeHistoryEntry(const HistoryEntry& entry) {
  std::wstring line = std::to_wstring(entry.timestamp);
  const std::wstring fields[] = {
      entry.time_text, entry.action, entry.old_data, entry.new_data,
      entry.key_path, entry.value_name};
  for (const std::wstring& field : fields) {
    line.push_back(L'\t');
    line.append(record_fields::Escape(field));
  }
  line.push_back(L'\t');
  line.append(std::to_wstring(static_cast<int>(entry.revert_kind)));
  line.push_back(L'\t');
  line.append(record_fields::Escape(entry.revert_value.name));
  line.push_back(L'\t');
  line.append(std::to_wstring(entry.revert_value.type));
  line.push_back(L'\t');
  line.append(record_fields::Escape(util::ToHex(
      entry.revert_value.data.data(), entry.revert_value.data.size(),
      entry.revert_value.data.size())));
  line.push_back(L'\n');
  return line;
}

bool WriteHistoryFile(const std::wstring& path,
                      const std::vector<HistoryEntry>& entries) {
  if (path.empty()) {
    return false;
  }
  std::string bytes;
  for (const auto& entry : entries) {
    bytes += util::WideToUtf8(SerializeHistoryEntry(entry));
  }
  if (bytes.empty()) {
    return DeleteFileW(path.c_str()) != 0 ||
           GetLastError() == ERROR_FILE_NOT_FOUND;
  }
  HANDLE file =
      CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr,
                  CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  DWORD written = 0;
  const bool success =
      WriteFile(file, bytes.data(), static_cast<DWORD>(bytes.size()), &written,
                nullptr) != 0 &&
      written == static_cast<DWORD>(bytes.size());
  CloseHandle(file);
  return success;
}

bool AppendHistoryFile(const std::wstring& path, const HistoryEntry& entry) {
  if (path.empty()) {
    return false;
  }
  const std::string bytes = util::WideToUtf8(SerializeHistoryEntry(entry));
  if (bytes.empty()) {
    return false;
  }
  HANDLE file =
      CreateFileW(path.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, nullptr,
                  OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  DWORD written = 0;
  const bool success =
      WriteFile(file, bytes.data(), static_cast<DWORD>(bytes.size()),
                &written, nullptr) != 0 &&
      written == static_cast<DWORD>(bytes.size());
  CloseHandle(file);
  return success;
}

bool PrepareRevert(const HistoryEntry& entry, const QueryValue& query_value,
                   HistoryEntry* prepared) {
  if (!prepared || entry.key_path.empty()) {
    return false;
  }
  *prepared = entry;
  if (prepared->revert_kind != HistoryEntry::RevertKind::kNone) {
    return true;
  }

  auto suffix = [&](const wchar_t* prefix, std::wstring* value_name) {
    const size_t length = wcslen(prefix);
    if (entry.action.size() < length ||
        _wcsnicmp(entry.action.c_str(), prefix, length) != 0) {
      return false;
    }
    *value_name = entry.action.substr(length);
    if (*value_name == L"(Default)") {
      value_name->clear();
    }
    return true;
  };

  std::wstring value_name;
  ValueEntry current;
  if (suffix(L"Create value ", &value_name)) {
    prepared->value_name = std::move(value_name);
    prepared->revert_kind = HistoryEntry::RevertKind::kDeleteValue;
    return true;
  }
  if (!suffix(L"Modify value ", &value_name) || entry.old_data.empty() ||
      !query_value ||
      !query_value(entry.key_path, value_name, &current)) {
    return false;
  }

  const DWORD type = value_format::NormalizeType(current.type);
  if (type != REG_DWORD && type != REG_DWORD_BIG_ENDIAN &&
      type != REG_QWORD) {
    return false;
  }
  const wchar_t* start = entry.old_data.c_str();
  while (*start && iswspace(*start)) {
    ++start;
  }
  if (start[0] != L'0' || (start[1] != L'x' && start[1] != L'X')) {
    return false;
  }
  errno = 0;
  wchar_t* end = nullptr;
  unsigned long long value = wcstoull(start + 2, &end, 16);
  if (end == start + 2 || errno == ERANGE) {
    return false;
  }
  while (*end && iswspace(*end)) {
    ++end;
  }
  if ((*end != L'\0' && *end != L'(') ||
      ((type == REG_DWORD || type == REG_DWORD_BIG_ENDIAN) &&
       value > std::numeric_limits<DWORD>::max())) {
    return false;
  }

  current.name = value_name;
  if (type == REG_QWORD) {
    current.data.resize(sizeof(value));
    memcpy(current.data.data(), &value, sizeof(value));
  } else if (type == REG_DWORD_BIG_ENDIAN) {
    current.data.resize(sizeof(DWORD));
    for (size_t index = 0; index < sizeof(DWORD); ++index) {
      current.data[sizeof(DWORD) - 1 - index] =
          static_cast<BYTE>(value & 0xff);
      value >>= 8;
    }
  } else {
    const DWORD dword = static_cast<DWORD>(value);
    current.data.resize(sizeof(dword));
    memcpy(current.data.data(), &dword, sizeof(dword));
  }
  prepared->value_name = value_name;
  prepared->revert_value = std::move(current);
  prepared->revert_kind = HistoryEntry::RevertKind::kSetValue;
  return true;
}

bool FindNearestExistingPath(const std::wstring& path,
                             const PathExists& path_exists,
                             std::wstring* nearest_path) {
  if (!nearest_path || !path_exists) {
    return false;
  }
  nearest_path->clear();
  std::wstring candidate = registry_path::Clean(path);
  while (!candidate.empty()) {
    if (path_exists(candidate)) {
      *nearest_path = std::move(candidate);
      return true;
    }
    const std::wstring parent = registry_path::Parent(candidate);
    if (parent.empty() || parent == candidate) {
      return false;
    }
    candidate = parent;
  }
  return false;
}

} // namespace regkit::changes
