// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "dialogs/bitfield_definition.h"

#include <cstdint>
#include <span>
#include <string>
#include <vector>

namespace regkit::editors
{

struct BitfieldRequest
{
    std::wstring value_name;
    std::wstring key_path;
    uint64_t value = 0;
    unsigned bit_count = 32;
    std::span<const BYTE> data;
    bool read_only = false;
};

struct BitfieldResult
{
    uint64_t value = 0;
    std::vector<BYTE> data;
};

bool EditBitfield(HWND owner, const BitfieldRequest& request, BitfieldResult* result);

}
