// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "search/result_file.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"

#include <limits>
#include <string_view>
#include <utility>

namespace regkit::search {

std::vector<Result> ParseResults(const std::wstring& content) {
  std::vector<Result> results;
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
    const auto fields = record_fields::Split(line);
    if (fields.size() < 13) {
      continue;
    }

    Result result;
    result.key_path = record_fields::Unescape(fields[0]);
    result.key_name = record_fields::Unescape(fields[1]);
    result.value_name = record_fields::Unescape(fields[2]);
    result.display_name = record_fields::Unescape(fields[3]);
    result.type_text = record_fields::Unescape(fields[4]);
    result.type = static_cast<DWORD>(_wtoi(fields[5].c_str()));
    result.data = record_fields::Unescape(fields[6]);
    result.size_text = record_fields::Unescape(fields[7]);
    result.date_text = record_fields::Unescape(fields[8]);
    size_t base = 9;
    if (fields.size() >= 14) {
      result.comment = record_fields::Unescape(fields[9]);
      base = 10;
    }
    result.is_key = _wtoi(fields[base].c_str()) != 0;
    const int match_field = _wtoi(fields[base + 1].c_str());
    result.match_field =
        match_field < 0 ||
                match_field > static_cast<int>(MatchField::kData)
            ? MatchField::kNone
            : static_cast<MatchField>(match_field);
    result.match_start = _wtoi(fields[base + 2].c_str());
    result.match_length = _wtoi(fields[base + 3].c_str());
    if (fields.size() > base + 4) {
      result.data_loaded = _wtoi(fields[base + 4].c_str()) != 0;
    }
    results.push_back(std::move(result));
  }
  return results;
}

std::wstring SerializeResults(const std::vector<Result>& results) {
  std::wstring content;
  for (const auto& result : results) {
    content += record_fields::Escape(result.key_path) + L'\t';
    content += record_fields::Escape(result.key_name) + L'\t';
    content += record_fields::Escape(result.value_name) + L'\t';
    content += record_fields::Escape(result.display_name) + L'\t';
    content += record_fields::Escape(result.type_text) + L'\t';
    content += std::to_wstring(result.type) + L'\t';
    content += record_fields::Escape(result.data) + L'\t';
    content += record_fields::Escape(result.size_text) + L'\t';
    content += record_fields::Escape(result.date_text) + L'\t';
    content += record_fields::Escape(result.comment) + L'\t';
    content += (result.is_key ? L"1\t" : L"0\t");
    content += std::to_wstring(static_cast<int>(result.match_field)) + L'\t';
    content += std::to_wstring(result.match_start) + L'\t';
    content += std::to_wstring(result.match_length) + L'\t';
    content += (result.data_loaded ? L"1\n" : L"0\n");
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
