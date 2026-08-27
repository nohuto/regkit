// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>
#include <vector>

namespace regkit::query_prompts {

bool ShowDataTypes(HWND owner, std::vector<DWORD>* types);
bool ShowRegistryKey(HWND owner, std::wstring* selected_path);

} // namespace regkit::query_prompts
