// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace regkit::value_format {

DWORD NormalizeType(DWORD type);
std::wstring TypeName(DWORD type);
std::wstring Data(DWORD type, const BYTE* data, DWORD size);
std::wstring DisplayData(DWORD type, const BYTE* data, DWORD size);

bool ParseHex(std::wstring_view text, std::vector<BYTE>* output);
std::vector<BYTE> StringData(std::wstring_view text);
bool DecodeString(std::span<const BYTE> data, std::wstring* output);
std::vector<std::wstring> MultiStringItems(const std::vector<BYTE>& data);
std::vector<BYTE> MultiStringData(const std::vector<std::wstring>& items);
std::wstring MultiStringText(const std::vector<BYTE>& data);
std::vector<BYTE> MultiStringData(std::wstring_view lines);

} // namespace regkit::value_format
