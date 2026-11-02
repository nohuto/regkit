// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/system/process_token.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>

namespace regkit::util
{

using win32ui::AccountName;
using win32ui::IsProcessPrivileged;
using win32ui::IsUacEnabled;
using win32ui::IsWritableByNonAdmins;
using win32ui::OpenShellToken;
using win32ui::PrivilegeScope;
using win32ui::ShellUserMatchesProcess;

// HKEY_CURRENT_USER follows the signed in user while enabled
std::wstring GetCurrentUserSidString();
std::wstring GetClassesUserSidString();
std::wstring GetShellUserSidString();
bool ShellUserDiffers();
void SetCurrentUserFollowsShell(bool enable);
bool CurrentUserFollowsShell();
HKEY MapCurrentUserRoot(HKEY root);
struct ServiceState
{
    DWORD state = 0;
    DWORD start_type = 0;
};
LONG QueryRemoteRegistryService(const std::wstring& machine, ServiceState* state);
LONG StartRemoteRegistryService(const std::wstring& machine, bool enable);
LONG StopRemoteRegistryService(const std::wstring& machine, DWORD start_type);

}
