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

struct BinaryRequest
{
    std::wstring value_name;
    std::wstring key_path;
    std::span<const BYTE> data;
    bool read_only = false;
};

struct BinaryResult
{
    std::vector<BYTE> data;
};

bool EditBinary(HWND owner, const BinaryRequest& request, BinaryResult* result);

}
