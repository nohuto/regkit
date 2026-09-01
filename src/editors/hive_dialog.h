// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include <string>

namespace regkit::editors {

struct LoadHiveResult {
  std::wstring file;
  std::wstring key_name;
  HKEY root = HKEY_LOCAL_MACHINE;
};

bool ChooseHiveToLoad(HWND owner, LoadHiveResult* result);

} // namespace regkit::editors
