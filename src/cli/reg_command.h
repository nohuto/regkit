// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <vector>

namespace regkit::cli
{

bool Execute(const std::vector<std::wstring>& args, int* exit_code);
bool PrintErrorToTerminal(const std::wstring& text);

}
