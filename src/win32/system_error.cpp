// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/system_error.h"

#include <iterator>

namespace util {

std::wstring FormatWin32Error(DWORD code) {
  if (code == ERROR_SUCCESS) {
    return {};
  }
  wchar_t buffer[512] = {};
  DWORD length = FormatMessageW(
      FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, code,
      0, buffer, static_cast<DWORD>(std::size(buffer)), nullptr);
  if (length == 0) {
    return L"Unknown error.";
  }
  while (length > 0 &&
         (buffer[length - 1] == L'\r' || buffer[length - 1] == L'\n')) {
    buffer[--length] = L'\0';
  }
  return buffer;
}

} // namespace util
