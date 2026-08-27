// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <span>
#include <string>
#include <vector>

namespace regkit::editors {

struct BinaryRequest {
  std::wstring value_name;
  std::wstring value_type = L"REG_BINARY";
  std::span<const BYTE> data;
};

struct BinaryResult {
  std::vector<BYTE> data;
};

bool EditBinary(HWND owner, const BinaryRequest& request,
                BinaryResult* result);

} // namespace regkit::editors
