// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <cstddef>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace util
{

inline constexpr int HexDigitValue(wchar_t character)
{
    return character >= L'0' && character <= L'9'   ? character - L'0'
           : character >= L'a' && character <= L'f' ? character - L'a' + 10
           : character >= L'A' && character <= L'F' ? character - L'A' + 10
                                                    : -1;
}
std::wstring WindowText(HWND window);
std::wstring DialogText(HWND dialog, int control_id);
bool ParseUnsignedNumber(std::wstring_view text, int base, unsigned long long* value);
std::wstring ToLower(std::wstring_view text);
std::wstring TrimWhitespace(std::wstring_view text);
std::wstring FormatLocalTime(const SYSTEMTIME& time, bool with_seconds = false);
bool IsBlank(std::wstring_view text);
std::wstring ExpandEnvironmentStringsDynamic(const std::wstring& text);
std::wstring ToHex(std::span<const BYTE> data, wchar_t separator = L' ', bool uppercase = false, size_t max_bytes = 0);

int CompareInsensitive(std::wstring_view left, std::wstring_view right);
int CompareListText(std::wstring_view left, std::wstring_view right);
bool EqualsInsensitive(std::wstring_view left, std::wstring_view right);
bool StartsWithInsensitive(std::wstring_view text, std::wstring_view prefix);
bool EndsWithInsensitive(std::wstring_view text, std::wstring_view suffix);
bool ContainsInsensitive(std::wstring_view text, std::wstring_view needle);
int FindInsensitive(std::wstring_view text, std::wstring_view needle);
bool ParseBool(std::wstring_view text);

std::vector<std::wstring> SplitLines(std::wstring_view text);
std::wstring JoinLines(const std::vector<std::wstring>& lines);

} // namespace util
