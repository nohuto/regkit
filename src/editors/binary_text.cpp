// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/binary_text.h"

#include <array>
#include <cwctype>

namespace regkit::editors::binary_text {

namespace {

constexpr wchar_t kHexDigits[] = L"0123456789ABCDEF";

const wchar_t* AnsiCharTable() {
  static const std::array<wchar_t, 256> table = [] {
    std::array<wchar_t, 256> entries = {};
    for (int index = 0; index < 256; ++index) {
      const char narrow = static_cast<char>(index);
      wchar_t wide = 0;
      const bool converted = MultiByteToWideChar(CP_ACP, MB_ERR_INVALID_CHARS,
                                                 &narrow, 1, &wide, 1) == 1;
      entries[static_cast<size_t>(index)] =
          (converted && iswprint(wide)) ? wide : L'.';
    }
    return entries;
  }();
  return table.data();
}

} // namespace

std::wstring Hex(std::span<const BYTE> data) {
  if (data.empty()) {
    return {};
  }
  std::wstring text;
  text.reserve(data.size() * 3 - 1);
  for (const BYTE byte : data) {
    if (!text.empty()) {
      text.push_back(L' ');
    }
    text.push_back(kHexDigits[byte >> 4]);
    text.push_back(kHexDigits[byte & 0x0F]);
  }
  return text;
}

std::wstring Preview(std::span<const BYTE> data, int group_bytes,
                     bool unicode) {
  constexpr size_t kBytesPerLine = 16;
  const size_t group =
      (group_bytes == 2 || group_bytes == 4 || group_bytes == 8)
          ? static_cast<size_t>(group_bytes)
          : 1;
  const int offset_width = data.size() > 0xFFFF ? 8 : 4;
  const wchar_t* ansi = AnsiCharTable();
  std::wstring text;
  text.reserve(data.size() * 5);

  for (size_t offset = 0; offset < data.size(); offset += kBytesPerLine) {
    wchar_t offset_text[16] = {};
    swprintf_s(offset_text, L"%0*X", offset_width,
               static_cast<unsigned int>(offset));
    text.append(offset_text);
    text.append(L"  ");

    for (size_t start = 0; start < kBytesPerLine; start += group) {
      for (size_t index = group; index-- > 0;) {
        const size_t position = offset + start + index;
        if (position < data.size()) {
          text.push_back(kHexDigits[data[position] >> 4]);
          text.push_back(kHexDigits[data[position] & 0x0F]);
        } else {
          text.append(L"  ");
        }
      }
      text.push_back(L' ');
    }
    text.append(L"  ");

    if (unicode) {
      for (size_t index = 0; index + 1 < kBytesPerLine; index += 2) {
        const size_t position = offset + index;
        if (position + 1 >= data.size()) {
          break;
        }
        const wchar_t value = static_cast<wchar_t>(
            data[position] | (data[position + 1] << 8));
        text.push_back(iswprint(value) ? value : L'.');
      }
    } else {
      for (size_t index = 0; index < kBytesPerLine; ++index) {
        const size_t position = offset + index;
        if (position >= data.size()) {
          break;
        }
        text.push_back(ansi[data[position]]);
      }
    }
    text.append(L"\r\n");
  }
  return text;
}

} // namespace regkit::editors::binary_text
