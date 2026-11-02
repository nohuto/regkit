// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <cstdint>
#include <initializer_list>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace regkit::record_fields
{

std::wstring Escape(std::wstring_view text);
std::wstring Unescape(std::wstring_view text);
void AppendRecord(std::wstring* output, std::initializer_list<std::wstring_view> fields);
void AppendRecord(std::wstring* output, std::span<const std::wstring> fields);
std::vector<std::wstring> DecodeRecord(std::wstring_view line);
bool ParseUnsigned(std::wstring_view text, uint64_t maximum, uint64_t* value);

}
