// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>
#include <vector>

namespace regkit::win32 {

inline constexpr wchar_t kRestartSystemArg[] = L"--restart-system";
inline constexpr wchar_t kRestartTiArg[] = L"--restart-ti";
inline constexpr wchar_t kRestartParentArg[] = L"--restart-from-pid";

std::wstring RestartArguments(const wchar_t* target_arg, DWORD parent_pid);

HRESULT LaunchElevated(HWND owner, const std::wstring& exe, const std::wstring& arguments);

DWORD RestartParentPid(const std::vector<std::wstring>& args);

void WaitForParentExit(DWORD parent_pid);

} // namespace regkit::win32
