// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

#include "win32/shell_integration.h"

namespace regkit
{
using namespace window_detail;

void MainWindow::Impl::ShowPermissionsDialog(const RegistryNode& node)
{
    ShowRegistryPermissions(hwnd_, node);
}

namespace
{

bool RestartExePath(HWND owner, std::wstring* exe_path)
{
    *exe_path = util::GetModulePath();
    if (!exe_path->empty())
    {
        return true;
    }
    ui::ShowError(owner, L"Failed to locate the executable path.");
    return false;
}

std::wstring WithErrorDetail(std::wstring message, DWORD error)
{
    const std::wstring detail = FormatWin32Error(error);
    return detail.empty() ? message : message + L"\n" + detail;
}

bool BeginRestart(HWND owner, const wchar_t* target_arg, const wchar_t* failure)
{
    std::wstring exe_path;
    if (!RestartExePath(owner, &exe_path))
    {
        return false;
    }
    // pass current PID so the replacement waits for this instance to exit
    const HRESULT hr =
        win32::LaunchElevated(owner, exe_path, win32::RestartArguments(target_arg, GetCurrentProcessId()));
    if (FAILED(hr))
    {
        if (!win32::DialogCancelled(hr))
        {
            ui::ShowError(owner, std::wstring(failure) + L"\n" + win32::FormatDialogError(hr));
        }
        return false;
    }
    PostMessageW(owner, WM_CLOSE, 0, 0);
    return true;
}

bool BrokerRestart(HWND owner, const wchar_t* target_arg, const wchar_t* failure, bool (*launch)(const std::wstring&, const std::wstring&, DWORD*, bool*))
{
    std::wstring exe_path;
    if (!RestartExePath(owner, &exe_path))
    {
        return false;
    }
    const std::wstring command_line =
        L"\"" + exe_path + L"\" " + win32::RestartArguments(target_arg, GetCurrentProcessId());
    DWORD error = 0;
    bool impersonation_lost = false;
    const bool launched = launch(command_line, L"", &error, &impersonation_lost);
    if (impersonation_lost)
    {
        ui::ShowError(owner, WithErrorDetail(L"RegKit couldn't restore its own security context and must close now.", error));
        ExitProcess(launched ? 0u : 1u);
    }
    if (!launched)
    {
        ui::ShowError(owner, WithErrorDetail(failure, error));
        return false;
    }
    PostMessageW(owner, WM_CLOSE, 0, 0);
    return true;
}

} // namespace

void MainWindow::Impl::PrepareSessionHandover()
{
    CaptureRegistryTabState(tab_ ? TabCtrl_GetCurSel(tab_) : -1);
    SaveSessionTabs();
    SaveSettings();
}

bool MainWindow::Impl::SaveSessionForRestart()
{
    CaptureRegistryTabState(tab_ ? TabCtrl_GetCurSel(tab_) : -1);
    if (SaveSessionTabs())
    {
        return true;
    }
    ui::ShowError(hwnd_, L"The current session couldn't be saved for the restart.");
    return false;
}

bool MainWindow::Impl::LaunchRestart(bool restore_session)
{
    if (ui::LaunchNewInstance(win32::RestartArguments(nullptr, GetCurrentProcessId(), restore_session)))
    {
        return true;
    }
    ui::ShowError(hwnd_, L"RegKit couldn't be restarted.");
    return false;
}

bool MainWindow::Impl::RestartCurrentInstance()
{
    if (!SaveSessionForRestart())
    {
        return false;
    }
    SaveSettings();
    return LaunchRestart(true);
}

bool MainWindow::Impl::RestartAfterCacheClear(CacheKind kind)
{
    // dont restore tab data when its cache was cleared
    const bool restore_session = kind != CacheKind::kAll && kind != CacheKind::kTabs;
    // tabs carry their own tree state
    if (kind == CacheKind::kTreeState)
    {
        for (TabEntry& entry : tabs_)
        {
            entry.selected_path.clear();
            entry.expanded_paths.clear();
        }
        ResetRegistryTreeState();
    }
    if (restore_session && !SaveSessionForRestart())
    {
        return false;
    }
    SaveSettings();
    if (!ClearCache(kind, false))
    {
        // continue tree state saving when the restart doesnt complete
        if ((kind == CacheKind::kAll || kind == CacheKind::kTreeState) && save_tree_state_)
        {
            StartTreeStateWorker();
        }
        BuildMenus();
        ui::ShowError(hwnd_, L"One or more cache files couldn't be removed.");
        return false;
    }
    if (!LaunchRestart(restore_session))
    {
        if ((kind == CacheKind::kAll || kind == CacheKind::kTreeState) && save_tree_state_)
        {
            StartTreeStateWorker();
        }
        BuildMenus();
        return false;
    }
    restart_on_close_ = true;
    return true;
}

bool MainWindow::Impl::RestartAfterSettingsReset()
{
    if (!SaveSessionForRestart())
    {
        return false;
    }
    const std::wstring path = SettingsPath();
    if (path.empty())
    {
        ui::ShowError(hwnd_, L"Failed to find the settings file.");
        return false;
    }
    if (DeleteFileW(path.c_str()) == 0)
    {
        const DWORD error = GetLastError();
        if (error != ERROR_FILE_NOT_FOUND && error != ERROR_PATH_NOT_FOUND)
        {
            ui::ShowError(hwnd_, L"The settings file couldn't be removed.\n" + FormatWin32Error(error));
            return false;
        }
    }
    if (!LaunchRestart(true))
    {
        // recreate settings when the replacement process couldnt start
        SaveSettings();
        return false;
    }
    return true;
}

bool MainWindow::Impl::RestartAsAdmin()
{
    PrepareSessionHandover();
    if (util::IsProcessSystem() || util::IsProcessTrustedInstaller())
    {
        // return through the signed in shell before requesting admin access
        return BrokerRestart(hwnd_, kRestartAdminArg, L"Failed to restart with administrator rights.", util::LaunchProcessAsShellUser);
    }
    return BeginRestart(hwnd_, nullptr, L"Failed to restart with administrator rights.");
}

bool MainWindow::Impl::RestartAsUser()
{
    PrepareSessionHandover();
    return BrokerRestart(hwnd_, kRestartUserArg, L"Failed to restart as the signed-in user.", util::LaunchProcessAsShellUser);
}

bool MainWindow::Impl::RestartAsSystem()
{
    PrepareSessionHandover();
    if (!util::IsProcessElevated())
    {
        return BeginRestart(hwnd_, kRestartSystemArg, L"Failed to request SYSTEM restart.");
    }
    return BrokerRestart(hwnd_, kRestartSystemArg, L"Failed to restart with SYSTEM rights.", util::LaunchProcessAsSystem);
}

bool MainWindow::Impl::RestartAsTrustedInstaller()
{
    PrepareSessionHandover();
    if (!util::IsProcessElevated())
    {
        return BeginRestart(hwnd_, kRestartTiArg, L"Failed to request TrustedInstaller restart.");
    }
    return BrokerRestart(hwnd_, kRestartTiArg, L"Failed to restart with TrustedInstaller rights.", util::LaunchProcessAsTrustedInstaller);
}

void MainWindow::Impl::SyncReplaceRegEditState()
{
    replace_regedit_ = win32::IsRegEditReplacementRegistered(util::GetModulePath());
}

void MainWindow::Impl::ReplaceRegEdit(bool enable)
{
    std::wstring exe_path = util::GetModulePath();
    if (exe_path.empty())
    {
        ui::ShowError(hwnd_, L"Failed to locate the executable path.");
        return;
    }
    // reject machine wide redirection to an executable another user can replace
    const bool writable_location = enable && util::IsWritableByNonAdmins(exe_path);
    if (writable_location &&
        ui::PromptKeyChoice(
            hwnd_,
            L"Replacing RegEdit registers this executable for every account on the machine.\n\n"
            L"RegKit is running from a location that non administrators can write to, so a program "
            L"without administrator rights could replace it and run whenever anyone starts RegEdit. "
            L"Install RegKit for all users first, or move it somewhere only administrators can write.\n\n"
            L"Replace anyway to apply it from this location.",
            exe_path,
            L"Replace RegEdit",
            L"Replace Anyway",
            L"",
            L"Cancel",
            {110, 70, 70}
        ) != IDYES)
    {
        SyncReplaceRegEditState();
        BuildMenus();
        return;
    }

    bool conflict = false;
    LONG result = win32::SetRegEditReplacement(exe_path, enable, &conflict, false, writable_location);
    // dont overwrite another debugger registration without approval
    if (result != ERROR_SUCCESS && conflict && enable)
    {
        const int choice = ui::PromptChoice(hwnd_, L"RegEdit already has a Debugger entry owned by another program.\n\n"
                                                   L"Override the existing entry?",
                                            L"Replace RegEdit",
                                            L"Override",
                                            L"",
                                            L"Cancel",
                                            {80, 70, 70});
        if (choice == IDYES)
        {
            result = win32::SetRegEditReplacement(exe_path, true, nullptr, true, writable_location);
            conflict = false;
        }
        else
        {
            result = ERROR_CANCELLED;
        }
    }
    if (result != ERROR_SUCCESS)
    {
        if (result != ERROR_CANCELLED)
        {
            ui::ShowError(hwnd_, FormatWin32Error(result));
        }
    }
    SyncReplaceRegEditState();
    BuildMenus();
}

void MainWindow::Impl::SyncEditContextMenuState()
{
    const std::wstring exe_path = util::GetModulePath();
    edit_context_menu_ = win32::IsRegFileEditMenuRegistered(exe_path);
}

void MainWindow::Impl::SetEditContextMenu(bool enable)
{
    LONG cleanup_result = ERROR_SUCCESS;
    const LONG result = win32::SetRegFileEditMenu(util::GetModulePath(), enable, &cleanup_result);
    if (result != ERROR_SUCCESS)
    {
        std::wstring message = FormatWin32Error(result);
        if (cleanup_result != ERROR_SUCCESS)
        {
            message += L"\nThe incomplete context menu entry couldn't be removed:\n";
            message += FormatWin32Error(cleanup_result);
        }
        ui::ShowError(hwnd_, message);
    }
    BuildMenus();
}

std::wstring MainWindow::Impl::ResolveSelectedHiveFilePath()
{
    if (registry_mode_ == RegistryMode::kRemote)
    {
        return L"";
    }
    RegistryNode* node = browse_.current_node();
    if (!node && browse_.tree().hwnd())
    {
        HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
        if (selected)
        {
            node = browse_.tree().NodeFromItem(selected);
        }
    }
    if (!node)
    {
        return L"";
    }
    RegistryNode target = *node;
    int index = browse_.values().hwnd() ? ListView_GetNextItem(browse_.values().hwnd(), -1, LVNI_SELECTED) : -1;
    if (index >= 0)
    {
        const ListRow* row = browse_.values().RowAt(index);
        if (row && row->kind == rowkind::kKey && !row->extra.empty())
        {
            target = ChildNode(*node, row->extra);
        }
    }
    return LookupHivePath(target, nullptr);
}

void MainWindow::Impl::OpenHiveFileDir()
{
    if (registry_mode_ == RegistryMode::kRemote)
    {
        ui::ShowError(hwnd_, L"Hive files aren't available for remote registries.");
        return;
    }
    std::wstring hive_path = ResolveSelectedHiveFilePath();
    if (hive_path.empty())
    {
        ui::ShowError(hwnd_, L"No hive file was found for this key.");
        return;
    }
    const HRESULT hr = win32::RevealInExplorer(hive_path);
    if (FAILED(hr))
    {
        ui::ShowError(hwnd_, win32::FormatDialogError(hr));
    }
}

LOGFONTW MainWindow::Impl::DefaultLogFont() const
{
    return ui::DefaultUIFontLogFont();
}

} // namespace regkit
