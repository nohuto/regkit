// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>

namespace util {

std::wstring GetCurrentUserSidString();
bool IsProcessElevated();
bool IsProcessSystem();
bool IsProcessTrustedInstaller();
bool LaunchProcessAsSystem(const std::wstring& command_line,
                           const std::wstring& work_dir,
                           DWORD* error_code = nullptr);
bool LaunchProcessAsTrustedInstaller(const std::wstring& command_line,
                                     const std::wstring& work_dir,
                                     DWORD* error_code = nullptr);

} // namespace util
