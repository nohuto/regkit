// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <vector>

namespace regkit::win32
{

inline constexpr wchar_t kRestartSystemArg[] = L"--restart-system";
inline constexpr wchar_t kRestartTiArg[] = L"--restart-ti";
inline constexpr wchar_t kRestartParentArg[] = L"--restart-from-pid";
inline constexpr wchar_t kRestartParentTimeArg[] = L"--restart-from-time";
inline constexpr wchar_t kRestartUserArg[] = L"--restart-user";
inline constexpr wchar_t kRestartAdminArg[] = L"--restart-admin";
inline constexpr wchar_t kRestartDataDirArg[] = L"--data-dir";
inline constexpr wchar_t kRestartSessionArg[] = L"--restore-session";
inline constexpr wchar_t kRestartUserSidArg[] = L"--user-sid";

// instance a restart waits for, its id alone may name another process once it exited
struct RestartParent
{
    DWORD pid = 0;
    ULONGLONG created = 0;
};

bool ArgTakesValue(const std::wstring& arg);
bool IsInternalRestartArg(const std::wstring& arg);
std::wstring RestartDataDir(const std::vector<std::wstring>& args);
bool RestoreSessionRequested();
std::wstring RestartUserSid();
RestartParent RestartParentOf(const std::vector<std::wstring>& args);

std::wstring RestartArguments(const wchar_t* target_arg, bool restore_session = true);
std::wstring RestartArguments(const wchar_t* target_arg, const RestartParent& parent, const std::vector<std::wstring>& original_args);

}
