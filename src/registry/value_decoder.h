// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <vector>

namespace regkit::value_decoder
{

enum class TransformId
{
    kNone,
    kBase64,
    kBase64Url,
    kHex,
    kPercent,
};

enum class DecoderId
{
    kRawBytes,
    kUtf8,
    kUtf16Le,
    kUtf16Be,
    kAscii,
    kFileTime,
    kSystemTime,
    kUnixSeconds,
    kUnixMilliseconds,
    kGuid,
    kSid,
    kSecurityDescriptor,
    kIpv4,
    kIpv6,
};

struct TransformEntry
{
    TransformId id = TransformId::kNone;
    const wchar_t* name = nullptr;
};

struct DecoderEntry
{
    DecoderId id = DecoderId::kRawBytes;
    const wchar_t* name = nullptr;
};

struct Field
{
    std::wstring name;
    std::wstring value;
};

struct Decoded
{
    bool ok = false;
    std::wstring error;
    std::vector<Field> fields;
};

std::vector<TransformEntry> AvailableTransforms(DWORD type, const std::vector<BYTE>& data);
std::vector<DecoderEntry> AvailableDecoders(const BYTE* data, size_t size);

bool Transform(TransformId id, DWORD type, const std::vector<BYTE>& source, std::vector<BYTE>* out, std::wstring* error);
Decoded Decode(DecoderId id, const BYTE* data, size_t size);
DecoderId Suggest(const std::wstring& key_path, const std::wstring& value_name, size_t size);

}
