// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <vector>

namespace regkit::editors
{

struct DecodeRequest
{
    std::wstring value_name;
    std::wstring key_path;
    DWORD type = REG_NONE;
    std::vector<BYTE> data;
};

void ShowValueDecoder(HWND owner, const DecodeRequest& request);

}
