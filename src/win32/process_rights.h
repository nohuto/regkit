// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

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

