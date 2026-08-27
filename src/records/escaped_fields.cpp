// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "records/escaped_fields.h"

namespace regkit::record_fields {

std::wstring Escape(const std::wstring& text) {
  std::wstring escaped;
  escaped.reserve(text.size());
  for (wchar_t character : text) {
    switch (character) {
    case L'\\':
      escaped.append(L"\\\\");
      break;
    case L'\t':
      escaped.append(L"\\t");
      break;
    case L'\r':
      escaped.append(L"\\r");
      break;
    case L'\n':
      escaped.append(L"\\n");
      break;
    default:
      escaped.push_back(character);
      break;
    }
  }
  return escaped;
}

std::wstring Unescape(const std::wstring& text) {
  std::wstring unescaped;
  unescaped.reserve(text.size());
  for (size_t index = 0; index < text.size(); ++index) {
    const wchar_t character = text[index];
    if (character == L'\\' && index + 1 < text.size()) {
      switch (text[index + 1]) {
      case L'\\':
        unescaped.push_back(L'\\');
        ++index;
        continue;
      case L't':
        unescaped.push_back(L'\t');
        ++index;
        continue;
      case L'r':
        unescaped.push_back(L'\r');
        ++index;
        continue;
      case L'n':
        unescaped.push_back(L'\n');
        ++index;
        continue;
      default:
        break;
      }
    }
    unescaped.push_back(character);
  }
  return unescaped;
}

std::vector<std::wstring> Split(const std::wstring& line) {
  std::vector<std::wstring> fields;
  size_t start = 0;
  for (;;) {
    const size_t separator = line.find(L'\t', start);
    if (separator == std::wstring::npos) {
      fields.emplace_back(line.substr(start));
      return fields;
    }
    fields.emplace_back(line.substr(start, separator - start));
    start = separator + 1;
  }
}

std::vector<std::wstring> Lines(const std::wstring& content) {
  std::vector<std::wstring> lines;
  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    lines.push_back(std::move(line));
    start = end + 1;
  }
  return lines;
}

} // namespace regkit::record_fields
