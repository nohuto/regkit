// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <cstddef>
#include <string>

namespace util {

std::wstring WindowText(HWND window);
std::wstring ToLower(const std::wstring& text);
std::wstring TrimWhitespace(const std::wstring& text);
std::wstring ExpandEnvironmentStringsDynamic(const std::wstring& text);
std::wstring ToHex(const BYTE* data, size_t size, size_t max_bytes);

} // namespace util
