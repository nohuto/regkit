// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/text_transform.h"

#include <cwctype>

namespace util {

std::wstring WindowText(HWND window) {
  const int length = window ? GetWindowTextLengthW(window) : 0;
  if (length <= 0) {
    return {};
  }
  std::wstring text(static_cast<size_t>(length), L'\0');
  const int copied = GetWindowTextW(window, text.data(), length + 1);
  text.resize(copied > 0 ? static_cast<size_t>(copied) : 0);
  return text;
}

std::wstring ToLower(const std::wstring& text) {
  std::wstring result = text;
  for (wchar_t& character : result) {
    character = towlower(character);
  }
  return result;
}

std::wstring TrimWhitespace(const std::wstring& text) {
  size_t first = 0;
  while (first < text.size() && iswspace(text[first])) {
    ++first;
  }
  size_t last = text.size();
  while (last > first && iswspace(text[last - 1])) {
    --last;
  }
  return text.substr(first, last - first);
}

std::wstring ExpandEnvironmentStringsDynamic(const std::wstring& text) {
  if (text.empty()) {
    return {};
  }
  const DWORD needed = ExpandEnvironmentStringsW(text.c_str(), nullptr, 0);
  if (needed == 0) {
    return {};
  }
  std::wstring expanded(needed, L'\0');
  const DWORD written =
      ExpandEnvironmentStringsW(text.c_str(), expanded.data(), needed);
  if (written == 0 || written > needed) {
    return {};
  }
  while (!expanded.empty() && expanded.back() == L'\0') {
    expanded.pop_back();
  }
  return expanded;
}

std::wstring ToHex(const BYTE* data, size_t size, size_t max_bytes) {
  if (!data || size == 0) {
    return {};
  }
  const size_t count =
      max_bytes == 0 || max_bytes > size ? size : max_bytes;
  static constexpr wchar_t digits[] = L"0123456789abcdef";
  std::wstring output;
  output.reserve(count * 3 + (count < size ? 4 : 0));
  for (size_t index = 0; index < count; ++index) {
    output.push_back(digits[data[index] >> 4]);
    output.push_back(digits[data[index] & 0x0F]);
    if (index + 1 < count) {
      output.push_back(L' ');
    }
  }
  if (count < size) {
    output += L" ...";
  }
  return output;
}

} // namespace util
