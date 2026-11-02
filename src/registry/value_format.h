// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <cstdint>
#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace regkit::value_format
{

struct TypeLabel
{
    DWORD type;
    const wchar_t* name;
};

std::span<const TypeLabel> TypeLabels();
DWORD NormalizeType(DWORD type);
std::wstring TypeName(DWORD type);
std::wstring ByteCount(size_t size);
std::wstring Data(DWORD type, const BYTE* data, DWORD size);
std::wstring DisplayData(DWORD type, const BYTE* data, DWORD size, bool resolve_indirect = true);

uint64_t ReadUnsigned(std::span<const BYTE> data, size_t width, bool big_endian = false);
std::vector<BYTE> UnsignedBytes(uint64_t value, size_t width, bool big_endian = false);
std::vector<BYTE> StringData(std::wstring_view text);
bool DecodeString(std::span<const BYTE> data, std::wstring* output);
std::vector<std::wstring> MultiStringItems(std::span<const BYTE> data);
std::vector<BYTE> MultiStringData(const std::vector<std::wstring>& items);
std::wstring MultiStringText(const std::vector<BYTE>& data);
std::vector<BYTE> MultiStringData(std::wstring_view lines);

}
