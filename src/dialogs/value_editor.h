// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <span>
#include <string>
#include <vector>

namespace regkit::editors
{

struct CustomValueRequest
{
    std::wstring value_name;
    DWORD type = REG_SZ;
    std::vector<BYTE> data;
    bool read_only = false;
};

struct CustomValueResult
{
    DWORD type = REG_SZ;
    std::vector<BYTE> data;
};

struct FlaggedValueRequest
{
    std::wstring value_name;
    DWORD base_type = REG_SZ;
    std::span<const BYTE> data;
    bool read_only = false;
};

struct FlaggedValueResult
{
    std::vector<BYTE> data;
};

struct BitsRequest
{
    std::wstring value_name;
    std::wstring key_path;
    DWORD base_type = REG_DWORD;
    std::span<const BYTE> data;
    bool read_only = false;
};

bool EditValueBits(HWND owner, const BitsRequest& request, std::vector<BYTE>* data);
bool EditCustomValue(HWND owner, const CustomValueRequest& request, CustomValueResult* result);
bool EditFlaggedValue(HWND owner, const FlaggedValueRequest& request, FlaggedValueResult* result);

}
