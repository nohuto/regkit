// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include <win32ui/windows_config.h>

#include <windows.h>

#include <algorithm>
#include <cwctype>
#include <optional>
#include <string>
#include <vector>

#include <win32ui/dialog/dialog_window.h>
#include <win32ui/handle_owner.h>
#include <win32ui/list_view/list_view.h>
#include <win32ui/system/app_runtime.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/paths.h>
#include <win32ui/system/process_launch.h>
#include <win32ui/system/process_token.h>
#include <win32ui/system/registry_value.h>
#include <win32ui/system/single_instance.h>
#include <win32ui/text/text.h>
#include <win32ui/window.h>

#include "cli/reg_command.h"
#include "frame/main_window.h"
#include "frame/window/message_ids.h"
#include "regfile/registry_transfer.h"
#include "registry/hive_files.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "resource.h"
#include "ui/dialog_layout.h"
#include "ui/feedback.h"
#include "ui/presets.h"
#include "ui/theme.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/restart.h"
#include "win32/shell_integration.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"
#include "workspace/settings.h"

namespace regkit
{
namespace
{

using frame::message_id::kEditRegFileCopyDataId;
using frame::message_id::kExternalJumpCopyDataId;
using frame::message_id::kRegKitWindowProperty;
using win32::kRestartAdminArg;
using win32::kRestartSystemArg;
using win32::kRestartTiArg;
using win32::kRestartUserArg;
constexpr wchar_t kEditRegFileArg[] = L"--edit-reg";
constexpr wchar_t kInstallEditContextMenuArg[] = L"--install-edit-context-menu";
constexpr wchar_t kUninstallEditContextMenuArg[] = L"--uninstall-edit-context-menu";
constexpr wchar_t kInstallRegEditReplacementArg[] = L"--install-regedit-replacement";
constexpr wchar_t kUninstallRegEditReplacementArg[] = L"--uninstall-regedit-replacement";
constexpr wchar_t kOverrideArg[] = L"--override";
constexpr wchar_t kSettingArg[] = L"--setting";

constexpr const wchar_t* kRegEditNames[] = {L"regedit.exe", L"regedit", L"regedt32.exe", L"regedt32"};

void ApplyDataDirOverride(const std::vector<std::wstring>& args)
{
    const std::wstring dir = win32::RestartDataDir(args);
    if (dir.empty())
    {
        return;
    }
    win32ui::EnsureDirectory(dir.c_str());
    const std::wstring probe = util::JoinPath(dir, L"session" + util::RandomFileSuffix(L".probe"));
    const win32ui::UniqueHandle handle(
        CreateFileW(probe.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE | FILE_FLAG_OPEN_REPARSE_POINT, nullptr)
    );
    if (handle)
    {
        SetEnvironmentVariableW(L"REGKIT_DATA_DIR", dir.c_str());
    }
}

bool HasCommandLineArg(const std::vector<std::wstring>& args, const wchar_t* arg)
{
    return std::any_of(args.begin(), args.end(), [&](const std::wstring& entry) { return util::EqualsInsensitive(entry, arg); });
}

bool IsRegEditLaunchArg(const std::wstring& arg)
{
    const bool drive_absolute =
        arg.size() >= 3 && iswalpha(arg[0]) && arg[1] == L':' && (arg[2] == L'\\' || arg[2] == L'/');
    const bool unc_absolute =
        arg.size() >= 3 && ((arg[0] == L'\\' && arg[1] == L'\\') || (arg[0] == L'/' && arg[1] == L'/'));
    const std::wstring name = util::FileName(arg);
    return (drive_absolute || unc_absolute) &&
           std::any_of(std::begin(kRegEditNames), std::end(kRegEditNames), [&](const wchar_t* regedit) { return util::EqualsInsensitive(name, regedit); });
}

bool IsInterceptedRegEditLaunch(const std::vector<std::wstring>& args)
{
    return std::any_of(args.begin(), args.end(), IsRegEditLaunchArg);
}

std::vector<std::wstring> StripRegEditLaunchArg(const std::vector<std::wstring>& args)
{
    std::vector<std::wstring> stripped;
    stripped.reserve(args.size());
    for (const auto& arg : args)
    {
        if (!IsRegEditLaunchArg(arg))
        {
            stripped.push_back(arg);
        }
    }
    return stripped;
}

bool LooksLikeRegistryPath(const std::wstring& arg)
{
    RegistryNode node;
    return registry_path::ParseRoot(arg, &node);
}

// before window creation only Win32 roots can be checked, assume other paths exist
bool JumpTargetFound(const std::wstring& target)
{
    RegistryNode node;
    if (!registry_path::ParseRoot(target, &node) || !node.root)
    {
        return true;
    }
    const std::wstring sid = util::GetCurrentUserSidString();
    std::wstring key_path;
    std::wstring value_name;
    bool value_missing = false;
    return registry_path::ResolveJumpTarget(
               target,
               [&](const std::wstring& path) { return registry_path::Normalize(path, sid); },
               [](const std::wstring& path, RegistryNode* key) { return registry_path::ParseRoot(path, key) && key->root && RegistryStore::KeyExists(*key); },
               &key_path,
               &value_name,
               &value_missing
           ) &&
           !value_missing;
}

struct LaunchArgs
{
    std::wstring jump_target;
    std::vector<std::wstring> reg_files;
    std::vector<std::wstring> hive_files;
    bool edit_reg = false;
    std::wstring error;
};

LaunchArgs ParseLaunchArgs(const std::vector<std::wstring>& args)
{
    LaunchArgs launch;
    bool intercepted_regedit = false;
    std::wstring unquoted;
    for (size_t index = 0; index < args.size(); ++index)
    {
        const std::wstring& arg = args[index];
        if (arg.empty())
        {
            continue;
        }
        if (IsRegEditLaunchArg(arg))
        {
            intercepted_regedit = true;
            continue;
        }
        const bool goto_arg = util::EqualsInsensitive(arg, L"--goto") || util::EqualsInsensitive(arg, L"/goto");
        if (goto_arg || win32::ArgTakesValue(arg))
        {
            if (++index >= args.size())
            {
                launch.error = util::TrLabel(L"Missing argument", arg);
                return launch;
            }
            if (goto_arg)
            {
                launch.jump_target = args[index];
            }
            continue;
        }
        if (arg[0] == L'-' || arg[0] == L'/')
        {
            // accept and ignore RegEdit switches
            const std::wstring_view name = std::wstring_view(arg).substr(1);
            const bool regedit_switch = intercepted_regedit || util::EqualsInsensitive(name, L"m") || util::EqualsInsensitive(name, L"c") ||
                                        util::StartsWithInsensitive(name, L"l:") || util::StartsWithInsensitive(name, L"r:");
            launch.edit_reg = launch.edit_reg || util::EqualsInsensitive(arg, kEditRegFileArg);
            if (!regedit_switch && !launch.edit_reg && !win32::IsInternalRestartArg(arg))
            {
                launch.error = util::TrLabel(L"Invalid option", arg);
                return launch;
            }
            continue;
        }
        if (LooksLikeRegistryPath(arg))
        {
            launch.jump_target = arg;
            continue;
        }
        const bool reg_file = util::HasFileExtension(arg, L".reg") && GetFileAttributesW(arg.c_str()) != INVALID_FILE_ATTRIBUTES;
        if (reg_file || IsHiveFile(arg))
        {
            (reg_file ? launch.reg_files : launch.hive_files).push_back(arg);
            continue;
        }
        if (!launch.jump_target.empty())
        {
            unquoted = (unquoted.empty() ? launch.jump_target : unquoted) + L" " + arg;
            continue;
        }
        launch.error = util::StartsWithInsensitive(arg, L"HK")                      ? util::TrDetail(L"Registry path not found.", arg)
                       : GetFileAttributesW(arg.c_str()) == INVALID_FILE_ATTRIBUTES ? util::TrDetail(L"Registry file not found.", arg)
                                                                                    : util::TrLabel(L"Invalid argument", arg);
        return launch;
    }
    if (!unquoted.empty())
    {
        launch.error = util::TrDetail(L"Registry paths with spaces must be in quotes.", unquoted);
    }
    else if (launch.edit_reg && launch.reg_files.empty() && launch.hive_files.empty())
    {
        launch.error = util::TrLabel(L"Missing argument", kEditRegFileArg);
    }
    std::wstring last_key;
    // RegEdit opens its last key unless deleted
    if (launch.jump_target.empty() && intercepted_regedit &&
        win32ui::ReadRegistryString(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\RegEdit", L"LastKey", &last_key, RRF_NOEXPAND) ==
            ERROR_SUCCESS)
    {
        const size_t root_end = last_key.find(L'\\');
        if (!LooksLikeRegistryPath(last_key) && root_end != std::wstring::npos)
        {
            last_key.erase(0, root_end + 1);
        }
        if (LooksLikeRegistryPath(last_key) && JumpTargetFound(last_key))
        {
            launch.jump_target = std::move(last_key);
        }
    }
    return launch;
}

workspace::Settings LoadStartupSettings()
{
    workspace::Settings settings;
    const std::wstring folder = util::GetAppDataFolder();
    if (!folder.empty())
    {
        workspace::LoadSettings(util::JoinPath(folder, L"settings.ini"), &settings);
    }
    return settings;
}

void ApplyStartupTheme(const workspace::Settings& settings)
{
    win32ui::ThemeMode mode = win32ui::ThemeMode::system;
    win32ui::ParseThemeMode(settings.theme_mode, &mode);
    std::vector<ThemePreset> presets;
    std::wstring active = settings.theme_preset;
    if (mode == win32ui::ThemeMode::custom)
    {
        ThemePresetStore::Load(&presets, &active);
    }
    const ThemePreset* preset = mode == win32ui::ThemeMode::custom ? FindThemePreset(presets, active) : nullptr;
    win32ui::SetThemeMode(mode, preset ? &preset->colors : nullptr);
}

struct RestartTarget
{
    const wchar_t* arg;
    bool (*is_current)();
    DWORD (*launch)(const std::wstring&, const std::wstring&, DWORD*);
    const wchar_t* request_failure;
    const wchar_t* launch_failure;
    const wchar_t* title;
};

constexpr RestartTarget kSystemTarget = {
    kRestartSystemArg,
    win32ui::IsProcessSystem,
    win32ui::LaunchAsSystem,
    util::TrNoop(L"Failed to request SYSTEM restart."),
    util::TrNoop(L"Failed to restart with SYSTEM rights."),
    util::TrNoop(L"Restart as SYSTEM")
};
constexpr RestartTarget kTrustedInstallerTarget = {
    kRestartTiArg,
    win32ui::IsProcessTrustedInstaller,
    win32ui::LaunchAsTrustedInstaller,
    util::TrNoop(L"Failed to request TrustedInstaller restart."),
    util::TrNoop(L"Failed to restart with TrustedInstaller rights."),
    util::TrNoop(L"Restart as TrustedInstaller")
};

void ShowStartupError(const std::wstring& message)
{
    if (!cli::PrintErrorToTerminal(message))
    {
        ui::ShowError(nullptr, message);
    }
}

std::wstring WithErrorDetail(const wchar_t* message, DWORD error)
{
    const std::wstring detail = win32ui::FormatWin32Error(error);
    return detail.empty() ? util::Tr(message) : util::TrDetail(message, detail);
}

bool RestartAs(const RestartTarget& target, const win32::RestartParent& parent, const std::vector<std::wstring>& original_args, int* exit_code)
{
    if (target.is_current())
    {
        return false;
    }
    const std::wstring exe_path = util::ModulePath();
    if (exe_path.empty())
    {
        ShowStartupError(util::Tr(L"Failed to locate the executable path."));
        return false;
    }
    // keep user arguments while replacing old internal restart flags
    const std::wstring arguments = win32::RestartArguments(target.arg, parent, original_args);
    *exit_code = 0;
    if (!win32ui::IsProcessElevated())
    {
        const DWORD error = win32ui::LaunchElevated(nullptr, exe_path, arguments);
        if (error == ERROR_SUCCESS)
        {
            return true;
        }
        if (error != ERROR_CANCELLED)
        {
            ShowStartupError(WithErrorDetail(target.request_failure, error));
        }
        return false;
    }
    if (!ui::ConfirmPrivilegedRestart(nullptr, util::Tr(target.title)))
    {
        return false;
    }
    DWORD restore_error = ERROR_SUCCESS;
    const DWORD error = target.launch(exe_path, arguments, &restore_error);
    if (restore_error != ERROR_SUCCESS)
    {
        ShowStartupError(util::Tr(L"RegKit couldn't restore its own security context and must close now."));
        *exit_code = error == ERROR_SUCCESS ? 0 : 1;
        return true;
    }
    if (error != ERROR_SUCCESS)
    {
        ShowStartupError(WithErrorDetail(target.launch_failure, error));
    }
    return error == ERROR_SUCCESS;
}

std::optional<int> RunShellIntegrationCommand(const std::vector<std::wstring>& args)
{
    const bool install_menu = HasCommandLineArg(args, kInstallEditContextMenuArg);
    const bool uninstall_menu = HasCommandLineArg(args, kUninstallEditContextMenuArg);
    const bool install_replacement = HasCommandLineArg(args, kInstallRegEditReplacementArg);
    if (!install_menu && !uninstall_menu && !install_replacement && !HasCommandLineArg(args, kUninstallRegEditReplacementArg))
    {
        return std::nullopt;
    }
    const std::wstring exe_path = util::ModulePath();
    if (exe_path.empty())
    {
        return 1;
    }
    LONG result = ERROR_SUCCESS;
    if (install_menu)
    {
        result = win32::SetEditMenu(exe_path, true);
    }
    else if (uninstall_menu)
    {
        result = win32::RemoveEditMenuIfOwned(exe_path);
    }
    else
    {
        bool conflict = false;
        result = win32::SetRegEditReplacement(exe_path, install_replacement, &conflict, HasCommandLineArg(args, kOverrideArg));
        if (result != ERROR_SUCCESS && conflict)
        {
            return 2;
        }
    }
    return result == ERROR_SUCCESS ? 0 : 1;
}

std::optional<int> RunSettingCommand(const std::vector<std::wstring>& args)
{
    std::wstring lines;
    for (size_t index = 0; index + 1 < args.size(); ++index)
    {
        if (util::EqualsInsensitive(args[index], kSettingArg))
        {
            lines.append(args[++index]).append(L"\n");
        }
    }
    if (lines.empty())
    {
        return std::nullopt;
    }
    const std::wstring folder = util::GetAppDataFolder();
    if (folder.empty())
    {
        return 1;
    }
    const std::wstring path = util::JoinPath(folder, L"settings.ini");
    workspace::Settings settings;
    std::wstring error;
    // unreadable settings file stays as it is
    if (!workspace::LoadSettings(path, &settings, &error) && !error.empty())
    {
        return 1;
    }
    const workspace::Settings changed = workspace::ParseSettings(lines, settings, &error);
    return error.empty() && workspace::SaveSettings(path, changed) ? 0 : 1;
}

int MergeRegFiles(const std::vector<std::wstring>& reg_files)
{
    for (const auto& path : reg_files)
    {
        if (!ui::ConfirmRegFileMerge(nullptr, path))
        {
            return 0;
        }
        std::wstring error;
        if (!ImportRegFileFromPath(path, &error))
        {
            ui::ShowRegFileMergeFailed(nullptr, path, error);
            return 1;
        }
        ui::ShowRegFileMergeSucceeded(nullptr, path);
    }
    return 0;
}

// instance a restart replaces is closing and never a target
bool HandOffToRunningInstance(const std::wstring& jump_target, const std::vector<std::wstring>& edit_files, const win32::RestartParent& replaced, DWORD wait_ms)
{
    const HWND existing = win32ui::FindInstanceWindow(kRegKitWindowProperty, wait_ms, replaced.pid, replaced.created);
    if (!existing)
    {
        return false;
    }
    // stop waiting when the existing instance is frozen
    constexpr DWORD kSendTimeoutMs = 1500;
    bool handed_off = jump_target.empty() || win32ui::SendInstanceText(existing, kExternalJumpCopyDataId, jump_target, kSendTimeoutMs) == ERROR_SUCCESS;
    for (const auto& path : edit_files)
    {
        // other instance has its own current directory
        handed_off = win32ui::SendInstanceText(existing, kEditRegFileCopyDataId, win32ui::FullPath(path), kSendTimeoutMs) == ERROR_SUCCESS && handed_off;
    }
    if (handed_off)
    {
        win32ui::BringToFront(existing);
    }
    return handed_off;
}

int RunMessageLoop(MainWindow& window)
{
    MSG msg = {};
    for (BOOL available; (available = GetMessageW(&msg, nullptr, 0, 0)) != 0;)
    {
        if (available == -1)
        {
            ui::ShowError(nullptr, util::Tr(L"Message loop failed unexpectedly."));
            return 1;
        }
        if (!window.TranslateAccelerator(msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    return static_cast<int>(msg.wParam);
}

int Run(HINSTANCE instance, int cmd_show)
{
    win32ui::HardenProcess();
    const auto args = win32ui::CommandLineArgs();
    if (const std::optional<int> exit_code = RunShellIntegrationCommand(args))
    {
        return *exit_code;
    }
    if (const std::optional<int> exit_code = RunSettingCommand(args))
    {
        return *exit_code;
    }

    const win32ui::GuiRuntime runtime;
    if (!runtime.com_ok())
    {
        ShowStartupError(L"COM initialization failed.");
        return 1;
    }
    win32ui::RegisterDialogClass(instance, L"RegKit.Dialog", appearance::DialogThemeHooks(), IDI_APPICON);
    win32ui::SetListGridIcons(IDI_ICON_DARK_GRID, IDI_ICON_LIGHT_GRID);

    ApplyDataDirOverride(args);
    int cli_exit = 0;
    if (cli::Execute(IsInterceptedRegEditLaunch(args) ? StripRegEditLaunchArg(args) : args, &cli_exit))
    {
        return cli_exit;
    }
    const workspace::Settings startup_settings = LoadStartupSettings();
    if (util::LoadLanguage(startup_settings.language))
    {
        win32ui::SetTextHook(win32ui::Translate);
    }
    ApplyStartupTheme(startup_settings);
    LaunchArgs launch = ParseLaunchArgs(args);
    if (!launch.error.empty())
    {
        ShowStartupError(launch.error);
        return 1;
    }
    // without a terminal, window offers the nearest key instead
    if (!launch.jump_target.empty() && !JumpTargetFound(launch.jump_target) &&
        cli::PrintErrorToTerminal(util::TrDetail(L"Registry path not found.", launch.jump_target)))
    {
        return 1;
    }
    const bool stay_as_user = HasCommandLineArg(args, kRestartUserArg);
    const win32::RestartParent restart_parent = win32::RestartParentOf(args);
    const RestartTarget* restart_target =
        HasCommandLineArg(args, kRestartTiArg)       ? &kTrustedInstallerTarget
        : HasCommandLineArg(args, kRestartSystemArg) ? &kSystemTarget
        : !stay_as_user && startup_settings.always_run_as_trustedinstaller && !win32ui::IsProcessTrustedInstaller()
            ? &kTrustedInstallerTarget
        : !stay_as_user && startup_settings.always_run_as_system && !win32ui::IsProcessSystem() ? &kSystemTarget
                                                                                                : nullptr;
    int restart_exit = 0;
    if (restart_target)
    {
        if (RestartAs(*restart_target, restart_parent, args, &restart_exit))
        {
            return restart_exit;
        }
    }
    else if ((HasCommandLineArg(args, kRestartAdminArg) || (!stay_as_user && startup_settings.always_run_as_admin)) &&
             !win32ui::IsProcessElevated())
    {
        const std::wstring exe_path = util::ModulePath();
        if (exe_path.empty())
        {
            ShowStartupError(util::Tr(L"Failed to locate the executable path."));
        }
        else if (const DWORD error = win32ui::LaunchElevated(nullptr, exe_path, win32::RestartArguments(nullptr, restart_parent, args));
                 error == ERROR_SUCCESS)
        {
            return 0;
        }
        else
        {
            ShowStartupError(error == ERROR_CANCELLED ? util::Tr(L"Administrator restart was cancelled.") : WithErrorDetail(util::TrNoop(L"Failed to restart with administrator rights."), error));
        }
    }

    if (!launch.edit_reg && !launch.reg_files.empty())
    {
        return MergeRegFiles(launch.reg_files);
    }

    // without its creation time the id may name another process by now
    if (restart_parent.created != 0)
    {
        win32ui::WaitForProcessExit(restart_parent.pid, restart_parent.created, 30000);
    }

    std::vector<std::wstring>& edit_files = launch.hive_files;
    if (launch.edit_reg)
    {
        edit_files.insert(edit_files.begin(), launch.reg_files.begin(), launch.reg_files.end());
    }
    win32ui::UniqueHandle instance_mutex;
    if (startup_settings.single_instance)
    {
        instance_mutex.reset(CreateMutexW(nullptr, TRUE, L"RegKit.SingleInstance"));
        const DWORD mutex_error = GetLastError();
        // first instance started at the same time holds the mutex before its window exists
        if ((mutex_error == ERROR_ALREADY_EXISTS || mutex_error == ERROR_ACCESS_DENIED) &&
            HandOffToRunningInstance(launch.jump_target, edit_files, restart_parent, restart_parent.pid == 0 ? 3000 : 0))
        {
            return 0;
        }
    }

    MainWindow window;
    if (!window.Create(instance))
    {
        ShowStartupError(util::Tr(L"Failed to create the main window."));
        return 1;
    }
    if (!launch.jump_target.empty())
    {
        window.QueueExternalJump(launch.jump_target);
    }
    for (const auto& path : edit_files)
    {
        window.OpenFile(path);
    }
    window.Show(cmd_show);
    return RunMessageLoop(window);
}

}
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int cmd_show)
{
    return regkit::Run(instance, cmd_show);
}
