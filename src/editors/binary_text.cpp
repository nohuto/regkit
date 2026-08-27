// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/binary_text.h"

#include <algorithm>
#include <cwctype>

namespace regkit::editors::binary_text {

std::wstring Hex(std::span<const BYTE> data) {
  if (data.empty()) {
    return {};
  }
  std::wstring text;
  text.reserve(data.size() * 3 - 1);
  constexpr wchar_t digits[] = L"0123456789ABCDEF";
  for (const BYTE byte : data) {
    if (!text.empty()) {
      text.push_back(L' ');
    }
    text.push_back(digits[byte >> 4]);
    text.push_back(digits[byte & 0x0F]);
  }
  return text;
}

std::wstring Preview(std::span<const BYTE> data, int group_bytes,
                     bool unicode) {
  constexpr size_t kBytesPerLine = 16;
  group_bytes = std::max(group_bytes, 1);
  const int offset_width = data.size() > 0xFFFF ? 8 : 4;
  std::wstring text;
  text.reserve(data.size() * 5);

  for (size_t offset = 0; offset < data.size(); offset += kBytesPerLine) {
    wchar_t offset_text[16] = {};
    swprintf_s(offset_text, L"%0*X", offset_width,
               static_cast<unsigned int>(offset));
    text.append(offset_text);
    text.append(L"  ");

    for (size_t index = 0; index < kBytesPerLine; ++index) {
      const size_t position = offset + index;
      if (position < data.size()) {
        wchar_t hex[4] = {};
        swprintf_s(hex, L"%02X", static_cast<unsigned int>(data[position]));
        text.append(hex);
      } else {
        text.append(L"  ");
      }
      text.push_back(L' ');
      if (index == 7) {
        text.push_back(L' ');
      }
      if (group_bytes > 1 &&
          (index + 1) % static_cast<size_t>(group_bytes) == 0) {
        text.push_back(L' ');
      }
    }

    text.push_back(L' ');
    if (unicode) {
      for (size_t index = 0; index < kBytesPerLine; index += 2) {
        const size_t position = offset + index;
        wchar_t character = L'.';
        if (position + 1 < data.size()) {
          const wchar_t value = static_cast<wchar_t>(
              data[position] | (data[position + 1] << 8));
          if (iswprint(value)) {
            character = value;
          }
        }
        text.push_back(character);
      }
    } else {
      for (size_t index = 0; index < kBytesPerLine; ++index) {
        const size_t position = offset + index;
        wchar_t character = L' ';
        if (position < data.size()) {
          const wchar_t value = static_cast<wchar_t>(data[position]);
          character = iswprint(value) ? value : L'.';
        }
        text.push_back(character);
      }
    }
    text.append(L"\r\n");
  }
  return text;
}

} // namespace regkit::editors::binary_text
