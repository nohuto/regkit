// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/result_file.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"

#include <limits>
#include <string_view>
#include <utility>

namespace regkit::search {

namespace {

constexpr wchar_t kRecordVersionTag[] = L"#regkit-search-2";

FILETIME FileTimeFromString(const std::wstring& text) {
  const unsigned long long value =
      text.empty() ? 0ull : _wcstoui64(text.c_str(), nullptr, 10);
  FILETIME time = {};
  time.dwLowDateTime = static_cast<DWORD>(value & 0xFFFFFFFFull);
  time.dwHighDateTime = static_cast<DWORD>(value >> 32);
  return time;
}

std::wstring FileTimeToString(const FILETIME& time) {
  const unsigned long long value =
      (static_cast<unsigned long long>(time.dwHighDateTime) << 32) |
      static_cast<unsigned long long>(time.dwLowDateTime);
  return std::to_wstring(value);
}

MatchField ToMatchField(int value) {
  return value < 0 || value > static_cast<int>(MatchField::kData)
             ? MatchField::kNone
             : static_cast<MatchField>(value);
}

Result ParseLegacyRecord(const std::vector<std::wstring>& fields) {
  Result result;
  result.key_path = record_fields::Unescape(fields[0]);
  result.value_name = record_fields::Unescape(fields[2]);
  result.type = static_cast<DWORD>(_wtoi(fields[5].c_str()));
  result.data_text = record_fields::Unescape(fields[6]);
  result.data_size =
      static_cast<DWORD>(_wtoi(record_fields::Unescape(fields[7]).c_str()));
  const size_t base = fields.size() >= 14 ? 10 : 9;
  result.kind = _wtoi(fields[base].c_str()) != 0 ? ResultKind::kKey
                                                 : ResultKind::kValue;
  result.match_field = ToMatchField(_wtoi(fields[base + 1].c_str()));
  const int start = _wtoi(fields[base + 2].c_str());
  result.match_start = start < 0 ? 0u : static_cast<uint32_t>(start);
  result.match_length = static_cast<uint32_t>(_wtoi(fields[base + 3].c_str()));
  bool loaded = true;
  if (fields.size() > base + 4) {
    loaded = _wtoi(fields[base + 4].c_str()) != 0;
  }
  result.data_state = result.kind == ResultKind::kKey ? DataState::kNotApplicable
                      : loaded                        ? DataState::kLoaded
                                                      : DataState::kNotLoaded;
  return result;
}

Result ParseVersionedRecord(const std::vector<std::wstring>& fields) {
  Result result;
  result.key_path = record_fields::Unescape(fields[0]);
  result.value_name = record_fields::Unescape(fields[1]);
  result.data_text = record_fields::Unescape(fields[2]);
  result.type = static_cast<DWORD>(_wcstoui64(fields[3].c_str(), nullptr, 10));
  result.data_size =
      static_cast<DWORD>(_wcstoui64(fields[4].c_str(), nullptr, 10));
  result.modified = FileTimeFromString(fields[5]);
  result.match_field = ToMatchField(_wtoi(fields[6].c_str()));
  result.match_start =
      static_cast<uint32_t>(_wcstoui64(fields[7].c_str(), nullptr, 10));
  result.match_length =
      static_cast<uint32_t>(_wcstoui64(fields[8].c_str(), nullptr, 10));
  const int kind = _wtoi(fields[9].c_str());
  result.kind = kind < 0 || kind > static_cast<int>(ResultKind::kTraceValue)
                    ? ResultKind::kValue
                    : static_cast<ResultKind>(kind);
  const int state = _wtoi(fields[10].c_str());
  result.data_state = state < 0 || state > static_cast<int>(DataState::kLoaded)
                          ? DataState::kNotApplicable
                          : static_cast<DataState>(state);
  return result;
}

} // namespace

std::vector<Result> ParseResults(const std::wstring& content) {
  std::vector<Result> results;
  bool versioned = false;
  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    start = end + 1;
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    if (line.empty()) {
      continue;
    }
    if (line == kRecordVersionTag) {
      versioned = true;
      continue;
    }
    const auto fields = record_fields::Split(line);
    if (versioned) {
      if (fields.size() >= 11) {
        results.push_back(ParseVersionedRecord(fields));
      }
      continue;
    }
    if (fields.size() < 13) {
      continue;
    }
    results.push_back(ParseLegacyRecord(fields));
  }
  return results;
}

std::wstring SerializeResults(const std::vector<Result>& results) {
  std::wstring content = kRecordVersionTag;
  content += L'\n';
  for (const auto& result : results) {
    content += record_fields::Escape(result.key_path) + L'\t';
    content += record_fields::Escape(result.value_name) + L'\t';
    content += record_fields::Escape(result.data_text) + L'\t';
    content += std::to_wstring(result.type) + L'\t';
    content += std::to_wstring(result.data_size) + L'\t';
    content += FileTimeToString(result.modified) + L'\t';
    content += std::to_wstring(static_cast<int>(result.match_field)) + L'\t';
    content += std::to_wstring(result.match_start) + L'\t';
    content += std::to_wstring(result.match_length) + L'\t';
    content += std::to_wstring(static_cast<int>(result.kind)) + L'\t';
    content += std::to_wstring(static_cast<int>(result.data_state)) + L'\n';
  }
  return content;
}

bool LoadResults(const std::wstring& path,
                 std::vector<Result>* results) {
  if (!results || path.empty()) {
    return false;
  }
  std::vector<BYTE> bytes;
  if (!util::ReadFileBytes(
          path, &bytes,
          static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
    return false;
  }
  size_t offset = 0;
  if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB &&
      bytes[2] == 0xBF) {
    offset = 3;
  }
  const std::wstring content = util::Utf8ToWide(std::string_view(
      reinterpret_cast<const char*>(bytes.data() + offset),
      bytes.size() - offset));
  *results = ParseResults(content);
  return true;
}

bool SaveResults(const std::wstring& path,
                 const std::vector<Result>& results) {
  return !path.empty() &&
         util::WriteTextFile(path, SerializeResults(results), false);
}

} // namespace regkit::search
