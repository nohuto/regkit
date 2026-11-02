// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/restart.h"
#include "win32/process_rights.h"
#include "win32/text_transform.h"

#include "win32/shell_paths.h"

#include <win32ui/system/process_launch.h>

#include <algorithm>

namespace regkit::win32
{

namespace
{

std::wstring BaseArguments(const wchar_t* target_arg, const RestartParent& parent, bool restore_session)
{
    std::wstring arguments;
    if (target_arg)
    {
        win32ui::AppendArgument(&arguments, target_arg);
    }
    win32ui::AppendArgument(&arguments, kRestartParentArg);
    win32ui::AppendArgument(&arguments, std::to_wstring(parent.pid));
    if (parent.created != 0)
    {
        win32ui::AppendArgument(&arguments, kRestartParentTimeArg);
        win32ui::AppendArgument(&arguments, std::to_wstring(parent.created));
    }
    const std::wstring data_dir = util::GetAppDataFolder();
    if (!data_dir.empty())
    {
        win32ui::AppendArgument(&arguments, kRestartDataDirArg);
        win32ui::AppendArgument(&arguments, data_dir);
        if (restore_session)
        {
            win32ui::AppendArgument(&arguments, kRestartSessionArg);
        }
    }
    const std::wstring user_sid = util::GetShellUserSidString();
    if (!user_sid.empty())
    {
        win32ui::AppendArgument(&arguments, kRestartUserSidArg);
        win32ui::AppendArgument(&arguments, user_sid);
    }
    return arguments;
}

RestartParent ThisProcess()
{
    const DWORD pid = GetCurrentProcessId();
    return {pid, win32ui::ProcessCreationTime(pid)};
}

}

bool ArgTakesValue(const std::wstring& arg)
{
    for (const wchar_t* flag : {kRestartParentArg, kRestartParentTimeArg, kRestartDataDirArg, kRestartUserSidArg})
    {
        if (util::EqualsInsensitive(arg, flag))
        {
            return true;
        }
    }
    return false;
}

bool IsInternalRestartArg(const std::wstring& arg)
{
    for (const wchar_t* flag :
         {kRestartSystemArg, kRestartTiArg, kRestartUserArg, kRestartAdminArg, kRestartSessionArg})
    {
        if (util::EqualsInsensitive(arg, flag))
        {
            return true;
        }
    }
    return ArgTakesValue(arg);
}

std::wstring RestartDataDir(const std::vector<std::wstring>& args)
{
    for (size_t i = 0; i + 1 < args.size(); ++i)
    {
        if (util::EqualsInsensitive(args[i], kRestartDataDirArg))
        {
            return args[i + 1];
        }
    }
    return L"";
}

bool RestoreSessionRequested()
{
    const std::vector<std::wstring> args = win32ui::CommandLineArgs();
    return std::any_of(args.begin(), args.end(), [](const std::wstring& arg) { return util::EqualsInsensitive(arg, kRestartSessionArg); });
}

std::wstring RestartUserSid()
{
    const std::vector<std::wstring> args = win32ui::CommandLineArgs();
    for (size_t i = 0; i + 1 < args.size(); ++i)
    {
        if (util::EqualsInsensitive(args[i], kRestartUserSidArg) && util::StartsWithInsensitive(args[i + 1], L"S-") &&
            args[i + 1].find(L'\\') == std::wstring::npos)
        {
            return args[i + 1];
        }
    }
    return {};
}

RestartParent RestartParentOf(const std::vector<std::wstring>& args)
{
    RestartParent parent;
    for (size_t i = 0; i + 1 < args.size(); ++i)
    {
        if (util::EqualsInsensitive(args[i], kRestartParentArg))
        {
            const ULONGLONG pid = win32ui::ParseUnsigned(args[i + 1]).value_or(0);
            parent.pid = pid <= MAXDWORD ? static_cast<DWORD>(pid) : 0;
        }
        else if (util::EqualsInsensitive(args[i], kRestartParentTimeArg))
        {
            parent.created = win32ui::ParseUnsigned(args[i + 1]).value_or(0);
        }
    }
    return parent.pid != 0 ? parent : RestartParent{};
}

std::wstring RestartArguments(const wchar_t* target_arg, bool restore_session)
{
    return BaseArguments(target_arg, ThisProcess(), restore_session);
}

std::wstring RestartArguments(const wchar_t* target_arg, const RestartParent& parent, const std::vector<std::wstring>& original_args)
{
    std::wstring arguments = BaseArguments(target_arg, parent.pid != 0 ? parent : ThisProcess(), true);
    for (size_t i = 0; i < original_args.size(); ++i)
    {
        const std::wstring& arg = original_args[i];
        if (IsInternalRestartArg(arg))
        {
            if (ArgTakesValue(arg))
            {
                ++i;
            }
            continue;
        }
        win32ui::AppendArgument(&arguments, arg);
    }
    return arguments;
}

}
