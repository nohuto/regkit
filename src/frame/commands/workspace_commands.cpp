// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"
#include "frame/tools/research_links.h"
#include "win32/shell_integration.h"
#include "win32/translation.h"

#include <win32ui/dialog/autocomplete.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/process_token.h>
#include <win32ui/system/shell_open.h>

namespace regkit
{
using namespace command_detail;

bool MainWindow::Impl::HandleWorkspaceAppearanceCommand(int command_id)
{
    if (const auto* link = frame::ResearchLinkForCommand(command_id))
    {
        ui::ReportFileDialogResult(hwnd_, win32ui::ShellOpen(hwnd_, link->url));
        return true;
    }
    if (command_id >= cmd::kTabClose && command_id <= cmd::kTabSelectMax)
    {
        return HandleTabCommand(command_id);
    }
    if (command_id >= cmd::kOptionsLanguageAuto && command_id <= cmd::kOptionsLanguageMax)
    {
        const size_t index = static_cast<size_t>(command_id - cmd::kOptionsLanguageBase);
        const std::wstring code = command_id == cmd::kOptionsLanguageAuto ? L"" : index < language_packs_.size() ? language_packs_[index].code
                                                                                                                 : settings_.language;
        if (code != settings_.language)
        {
            settings_.language = code;
            SaveSettings();
            BuildMenus();
            if (ui::PromptChoice(hwnd_, util::Tr(L"Restart RegKit now to apply the language?"), util::Tr(L"Language"), util::Tr(L"Restart"), L"", util::Tr(L"Later")) == IDYES)
            {
                restart_on_close_ = true;
                PostMessageW(hwnd_, WM_CLOSE, 0, 0);
            }
        }
        return true;
    }
    switch (command_id)
    {
    case cmd::kWindowNew:
    case cmd::kWindowClose:
    case cmd::kWindowAlwaysOnTop:
    case cmd::kOptionsThemeSystem:
    case cmd::kOptionsThemeLight:
    case cmd::kOptionsThemeDark:
    case cmd::kOptionsThemeCustom:
    case cmd::kOptionsThemePresets:
    case cmd::kOptionsIconSetPhosphor:
    case cmd::kOptionsIconSetClassic:
    case cmd::kOptionsIconSetCustom:
        return HandleWindowAppearanceCommand(command_id);
    case cmd::kOptionsRestartAdmin:
    case cmd::kOptionsRestartUser:
    case cmd::kOptionsAlwaysRunAdmin:
    case cmd::kOptionsRestartSystem:
    case cmd::kOptionsAlwaysRunSystem:
    case cmd::kOptionsRestartTrustedInstaller:
    case cmd::kOptionsAlwaysRunTrustedInstaller:
    case cmd::kOptionsReplaceRegEdit:
    case cmd::kOptionsEditContextMenu:
    case cmd::kOptionsSingleInstance:
    case cmd::kOptionsAutoComplete:
    case cmd::kOptionsHkcuFollowsUser:
    case cmd::kOptionsBackupRestore:
    case cmd::kOptionsHiveFileDir:
    case cmd::kOptionsResetSettings:
    case cmd::kHelpAbout:
    case cmd::kHelpContents:
    case cmd::kHelpDiscord:
    case cmd::kHelpCheckUpdates:
    case cmd::kHelpAutoCheckUpdates:
        return HandleLaunchHelpCommand(command_id);
    case cmd::kFavoritesAdd:
    case cmd::kFavoritesRemove:
    case cmd::kFavoritesEdit:
    case cmd::kFavoritesImport:
    case cmd::kFavoritesImportRegEdit:
    case cmd::kFavoritesExport:
        return HandleFavoritesCommand(command_id);
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleWindowAppearanceCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kWindowNew:
        if (const DWORD error = ui::LaunchNewInstance(hwnd_); error != ERROR_SUCCESS)
        {
            ui::ShowError(hwnd_, util::TrDetail(L"A new RegKit window couldn't be opened.", win32ui::FormatWin32Error(error)));
        }
        return true;
    case cmd::kWindowClose:
        PostMessageW(hwnd_, WM_CLOSE, 0, 0);
        return true;
    case cmd::kWindowAlwaysOnTop:
        settings_.always_on_top = !settings_.always_on_top;
        ApplyAlwaysOnTop();
        SaveSettings();
        return true;
    case cmd::kOptionsThemeSystem:
        theme_mode_ = win32ui::ThemeMode::system;
        win32ui::SetThemeMode(theme_mode_);
        ApplySystemTheme();
        SaveSettings();
        return true;
    case cmd::kOptionsThemeLight:
        theme_mode_ = win32ui::ThemeMode::light;
        win32ui::SetThemeMode(theme_mode_);
        ApplySystemTheme();
        SaveSettings();
        return true;
    case cmd::kOptionsThemeDark:
        theme_mode_ = win32ui::ThemeMode::dark;
        win32ui::SetThemeMode(theme_mode_);
        ApplySystemTheme();
        SaveSettings();
        return true;
    case cmd::kOptionsThemeCustom:
        ApplyThemePresetByName(settings_.theme_preset, true);
        return true;
    case cmd::kOptionsThemePresets:
        ShowThemePresetsDialog();
        return true;
    case cmd::kOptionsIconSetPhosphor:
        settings_.icon_set = kIconSetPhosphor;
        ReloadThemeIcons();
        SaveSettings();
        return true;
    case cmd::kOptionsIconSetClassic:
        settings_.icon_set = kIconSetClassic;
        ReloadThemeIcons();
        SaveSettings();
        return true;
    case cmd::kOptionsIconSetCustom:
        settings_.icon_set = kIconSetCustom;
        ReloadThemeIcons();
        SaveSettings();
        return true;
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleLaunchHelpCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kOptionsRestartAdmin:
        RestartAsAdmin();
        return true;
    case cmd::kOptionsRestartUser:
        RestartAsUser();
        return true;
    case cmd::kOptionsAlwaysRunAdmin:
        settings_.always_run_as_admin = !settings_.always_run_as_admin;
        if (settings_.always_run_as_admin)
        {
            settings_.always_run_as_system = false;
            settings_.always_run_as_trustedinstaller = false;
        }
        SaveSettings();
        if (settings_.always_run_as_admin && !win32ui::IsProcessElevated())
        {
            RestartAsAdmin();
        }
        return true;
    case cmd::kOptionsRestartSystem:
        RestartAsSystem();
        return true;
    case cmd::kOptionsAlwaysRunSystem:
        settings_.always_run_as_system = !settings_.always_run_as_system;
        if (settings_.always_run_as_system)
        {
            settings_.always_run_as_admin = false;
            settings_.always_run_as_trustedinstaller = false;
        }
        SaveSettings();
        if (settings_.always_run_as_system && !win32ui::IsProcessSystem())
        {
            RestartAsSystem();
        }
        return true;
    case cmd::kOptionsRestartTrustedInstaller:
        RestartAsTrustedInstaller();
        return true;
    case cmd::kOptionsAlwaysRunTrustedInstaller:
        settings_.always_run_as_trustedinstaller = !settings_.always_run_as_trustedinstaller;
        if (settings_.always_run_as_trustedinstaller)
        {
            settings_.always_run_as_admin = false;
            settings_.always_run_as_system = false;
        }
        SaveSettings();
        if (settings_.always_run_as_trustedinstaller && !win32ui::IsProcessTrustedInstaller())
        {
            RestartAsTrustedInstaller();
        }
        return true;
    case cmd::kOptionsReplaceRegEdit:
        ReplaceRegEdit(!win32::IsRegEditReplacementRegistered(util::ModulePath()));
        return true;
    case cmd::kOptionsEditContextMenu:
        SetEditContextMenu(!win32::IsEditMenuRegistered(util::ModulePath()));
        return true;
    case cmd::kOptionsSingleInstance:
        settings_.single_instance = !settings_.single_instance;
        SaveSettings();
        return true;
    case cmd::kOptionsHkcuFollowsUser:
        // running workers would read or change part of HKCU as the other user
        if (util::ShellUserDiffers())
        {
            key_copy_session_.CancelAndJoin();
            StopReplace();
            if (compare_session_.running())
            {
                compare_session_.CancelAndJoin();
                SetStatusMessage(L"");
            }
            for (SearchTab& tab : search_tabs_)
            {
                if (tab.run)
                {
                    // results found so far stay, the status says the search was stopped
                    tab.run->Stop();
                    tab.stopped = true;
                    FinishSearch(&tab);
                }
            }
        }
        settings_.hkcu_follows_shell_user = !settings_.hkcu_follows_shell_user;
        util::SetCurrentUserFollowsShell(settings_.hkcu_follows_shell_user);
        SaveSettings();
        ReloadLocalRoots();
        UpdateStatus();
        return true;
    case cmd::kOptionsBackupRestore:
        // privileges stay enabled only while this mode is on, never save the mode
        if (backup_privileges_)
        {
            backup_privileges_.reset();
        }
        else
        {
            backup_privileges_ = std::make_unique<util::PrivilegeScope>(std::initializer_list<const wchar_t*>{SE_BACKUP_NAME, SE_RESTORE_NAME});
            if (!backup_privileges_->held())
            {
                backup_privileges_.reset();
                ui::ShowError(hwnd_, util::Tr(L"Backup/restore mode needs the backup and restore privileges. Run RegKit elevated."));
            }
        }
        util::SetBackupRestoreMode(backup_privileges_ != nullptr);
        RegistryStore::NoteKeyChange();
        RefreshWholeTree();
        UpdateValueListForNode(browse_.current_node());
        UpdateStatus();
        return true;
    case cmd::kOptionsAutoComplete:
        settings_.autocomplete = !settings_.autocomplete;
        win32ui::SetAutoCompleteEnabled(settings_.autocomplete);
        SaveSettings();
        return true;
    case cmd::kOptionsHiveFileDir:
        OpenHiveFileDir();
        return true;
    case cmd::kOptionsResetSettings:
        if (ui::PromptChoice(hwnd_, util::Tr(L"Reset all settings and restart RegKit?"), util::Tr(L"Reset Settings"), util::Tr(L"Reset"), L"", util::Tr(L"Cancel")) == IDYES)
        {
            reset_settings_on_close_ = true;
            PostMessageW(hwnd_, WM_CLOSE, 0, 0);
        }
        return true;
    case cmd::kHelpAbout:
        ui::ShowAbout(hwnd_);
        return true;
    case cmd::kHelpContents:
        ui::ReportFileDialogResult(hwnd_, win32ui::ShellOpen(hwnd_, kHelpUrl));
        return true;
    case cmd::kHelpDiscord:
        ui::ReportFileDialogResult(hwnd_, win32ui::ShellOpen(hwnd_, kDiscordUrl));
        return true;
    case cmd::kHelpCheckUpdates:
        updates_.Check(false);
        return true;
    case cmd::kHelpAutoCheckUpdates:
        settings_.auto_check_updates = !settings_.auto_check_updates;
        SaveSettings();
        return true;
    default:
        return false;
    }
}

}
