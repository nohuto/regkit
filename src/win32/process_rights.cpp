// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/process_rights.h"
#include "win32/restart.h"
#include "win32/text_transform.h"

#include <atomic>
#include <vector>

#include <win32ui/handle_owner.h>
#include <win32ui/system/service_control.h>

namespace regkit::util
{
namespace
{

// signed in desktop user hives, set only when process user differs
struct ShellUserRoots
{
    std::wstring sid;
    win32ui::UniqueHKey user;
    win32ui::UniqueHKey classes;
};

const ShellUserRoots& ShellRoots()
{
    static const ShellUserRoots roots = [] {
        ShellUserRoots result;
        const win32ui::UniqueHandle token = OpenShellToken(TOKEN_QUERY);
        result.sid = token ? win32ui::TokenUserSid(token.get()) : win32::RestartUserSid();
        if (result.sid.empty() || EqualsInsensitive(result.sid, win32ui::ProcessUserSid()) ||
            RegOpenKeyExW(HKEY_USERS, result.sid.c_str(), 0, MAXIMUM_ALLOWED, result.user.put()) != ERROR_SUCCESS)
        {
            result.sid.clear();
            return result;
        }
        if (token)
        {
            RegOpenUserClassesRoot(token.get(), 0, MAXIMUM_ALLOWED, result.classes.put());
        }
        return result;
    }();
    return roots;
}

std::atomic_bool g_follow_shell_user{false};

win32ui::UniqueService OpenRemoteRegistry(const std::wstring& machine, DWORD access, LONG* error)
{
    win32ui::UniqueService scm(OpenSCManagerW(machine.c_str(), nullptr, SC_MANAGER_CONNECT));
    win32ui::UniqueService service(scm ? OpenServiceW(scm.get(), L"RemoteRegistry", access) : nullptr);
    *error = service ? ERROR_SUCCESS : static_cast<LONG>(GetLastError());
    return service;
}

}

std::wstring GetCurrentUserSidString()
{
    return CurrentUserFollowsShell() ? ShellRoots().sid : win32ui::ProcessUserSid();
}

std::wstring GetClassesUserSidString()
{
    return CurrentUserFollowsShell() && ShellRoots().classes ? ShellRoots().sid : win32ui::ProcessUserSid();
}

std::wstring GetShellUserSidString()
{
    return ShellUserDiffers() ? ShellRoots().sid : win32ui::ProcessUserSid();
}

bool ShellUserDiffers()
{
    return !ShellRoots().sid.empty();
}

void SetCurrentUserFollowsShell(bool enable)
{
    g_follow_shell_user = enable;
}

bool CurrentUserFollowsShell()
{
    return g_follow_shell_user && ShellUserDiffers();
}

HKEY MapCurrentUserRoot(HKEY root)
{
    if (!CurrentUserFollowsShell())
    {
        return root;
    }
    const ShellUserRoots& roots = ShellRoots();
    return root == HKEY_CURRENT_USER ? roots.user.get() : root == HKEY_CLASSES_ROOT && roots.classes ? roots.classes.get()
                                                                                                     : root;
}

LONG QueryRemoteRegistryService(const std::wstring& machine, ServiceState* state)
{
    LONG error = ERROR_SUCCESS;
    const win32ui::UniqueService service = OpenRemoteRegistry(machine, SERVICE_QUERY_STATUS | SERVICE_QUERY_CONFIG, &error);
    SERVICE_STATUS_PROCESS status = {};
    DWORD size = 0;
    QueryServiceConfigW(service.get(), nullptr, 0, &size);
    std::vector<BYTE> config(size);
    if (!service || !win32ui::QueryServiceProcess(service.get(), &status) || config.empty() ||
        !QueryServiceConfigW(service.get(), reinterpret_cast<QUERY_SERVICE_CONFIGW*>(config.data()), size, &size))
    {
        return service ? static_cast<LONG>(GetLastError()) : error;
    }
    state->state = status.dwCurrentState;
    state->start_type = reinterpret_cast<const QUERY_SERVICE_CONFIGW*>(config.data())->dwStartType;
    return ERROR_SUCCESS;
}

LONG StartRemoteRegistryService(const std::wstring& machine, bool enable)
{
    LONG error = ERROR_SUCCESS;
    const win32ui::UniqueService service = OpenRemoteRegistry(machine, SERVICE_QUERY_STATUS | SERVICE_START | SERVICE_STOP | SERVICE_CHANGE_CONFIG, &error);
    if (!service)
    {
        return error;
    }
    if (enable && !ChangeServiceConfigW(service.get(), SERVICE_NO_CHANGE, SERVICE_DEMAND_START, SERVICE_NO_CHANGE, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr))
    {
        return static_cast<LONG>(GetLastError());
    }
    SERVICE_STATUS_PROCESS status = {};
    LONG result = ERROR_SUCCESS;
    const bool accepted = StartServiceW(service.get(), 0, nullptr) != FALSE;
    if ((!accepted && GetLastError() != ERROR_SERVICE_ALREADY_RUNNING) || !win32ui::QueryServiceProcess(service.get(), &status) ||
        !win32ui::WaitWhileServicePending(service.get(), SERVICE_START_PENDING, &status))
    {
        result = static_cast<LONG>(GetLastError());
    }
    else if (status.dwCurrentState != SERVICE_RUNNING)
    {
        result = ERROR_SERVICE_NOT_ACTIVE;
    }
    // stop failed service starts and restore disabled state if previously disabled
    if (result != ERROR_SUCCESS && accepted)
    {
        SERVICE_STATUS stopped = {};
        ControlService(service.get(), SERVICE_CONTROL_STOP, &stopped);
    }
    if (result != ERROR_SUCCESS && enable)
    {
        ChangeServiceConfigW(service.get(), SERVICE_NO_CHANGE, SERVICE_DISABLED, SERVICE_NO_CHANGE, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr);
    }
    return result;
}

LONG StopRemoteRegistryService(const std::wstring& machine, DWORD start_type)
{
    LONG error = ERROR_SUCCESS;
    const win32ui::UniqueService service = OpenRemoteRegistry(machine, SERVICE_QUERY_STATUS | SERVICE_STOP | SERVICE_CHANGE_CONFIG, &error);
    SERVICE_STATUS stopped = {};
    if (!service || (!ControlService(service.get(), SERVICE_CONTROL_STOP, &stopped) && GetLastError() != ERROR_SERVICE_NOT_ACTIVE) ||
        !ChangeServiceConfigW(service.get(), SERVICE_NO_CHANGE, start_type, SERVICE_NO_CHANGE, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr))
    {
        return service ? static_cast<LONG>(GetLastError()) : error;
    }
    return ERROR_SUCCESS;
}

}
