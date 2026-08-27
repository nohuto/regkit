// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>
#include <vector>

namespace regkit {

struct RegistryValue {
  std::wstring name;
  DWORD type = REG_NONE;
  std::vector<BYTE> data;
};

} // namespace regkit
