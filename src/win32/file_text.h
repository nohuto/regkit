// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace util {

std::string WideToUtf8(const std::wstring& text);
std::wstring Utf8ToWide(std::string_view text);
bool ReadFileBytes(
    const std::wstring& path, std::vector<BYTE>* output,
    uint64_t max_bytes = 64ull * 1024ull * 1024ull,
    DWORD share_mode = FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE);
bool ReadTextFile(
    const std::wstring& path, std::wstring* output, bool* utf16 = nullptr,
    uint64_t max_bytes = 64ull * 1024ull * 1024ull,
    DWORD share_mode = FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE);
bool WriteTextFile(const std::wstring& path, const std::wstring& text,
                   bool utf16);

} // namespace util
