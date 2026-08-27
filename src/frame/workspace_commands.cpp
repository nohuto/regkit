// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"
#include "frame/research_links.h"

namespace regkit {
using namespace command_detail;

bool MainWindow::Impl::HandleWorkspaceAppearanceCommand(int command_id) {
  if (const auto* link = frame::ResearchLinkForCommand(command_id)) {
    HINSTANCE result = ShellExecuteW(hwnd_, L"open", link->url, nullptr,
                                     nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
      ui::ShowError(hwnd_, L"Failed to open the research link.");
    }
    return true;
  }
  switch (command_id) {
  case cmd::kWindowNew:
  case cmd::kWindowClose:
  case cmd::kWindowAlwaysOnTop:
  case cmd::kOptionsThemeSystem:
  case cmd::kOptionsThemeLight:
  case cmd::kOptionsThemeDark:
  case cmd::kOptionsThemeCustom:
  case cmd::kOptionsThemePresets:
  case cmd::kOptionsIconSetDefault:
  case cmd::kOptionsIconSetPhosphor:
  case cmd::kOptionsIconSetLucide:
  case cmd::kOptionsIconSetMaterialSymbols:
  case cmd::kOptionsIconSetCustom:
    return HandleWindowAppearanceCommand(command_id);
  case cmd::kOptionsRestartAdmin:
  case cmd::kOptionsAlwaysRunAdmin:
  case cmd::kOptionsRestartSystem:
  case cmd::kOptionsAlwaysRunSystem:
  case cmd::kOptionsRestartTrustedInstaller:
  case cmd::kOptionsAlwaysRunTrustedInstaller:
  case cmd::kOptionsOpenDefaultRegedit:
  case cmd::kOptionsReplaceRegedit:
  case cmd::kOptionsSingleInstance:
  case cmd::kOptionsHiveFileDir:
  case cmd::kHelpAbout:
  case cmd::kHelpContents:
    return HandleLaunchHelpCommand(command_id);
  case cmd::kFavoritesAdd:
  case cmd::kFavoritesRemove:
  case cmd::kFavoritesEdit:
  case cmd::kFavoritesImport:
  case cmd::kFavoritesImportRegedit:
  case cmd::kFavoritesExport:
    return HandleFavoritesCommand(command_id);
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleWindowAppearanceCommand(int command_id) {
  switch (command_id) {
  case cmd::kWindowNew:
    ui::LaunchNewInstance();
    return true;
  case cmd::kWindowClose:
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  case cmd::kWindowAlwaysOnTop:
    always_on_top_ = !always_on_top_;
    ApplyAlwaysOnTop();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeSystem:
    theme_mode_ = ThemeMode::kSystem;
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeLight:
    theme_mode_ = ThemeMode::kLight;
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeDark:
    theme_mode_ = ThemeMode::kDark;
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeCustom:
    ApplyThemePresetByName(active_theme_preset_, true);
    return true;
  case cmd::kOptionsThemePresets:
    ShowThemePresetsDialog();
    return true;
  case cmd::kOptionsIconSetDefault:
    icon_set_ = kIconSetDefault;
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetPhosphor:
    icon_set_ = kIconSetPhosphor;
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetLucide:
    icon_set_ = kIconSetLucide;
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetMaterialSymbols:
    icon_set_ = kIconSetMaterialSymbols;
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetCustom:
    icon_set_ = kIconSetCustom;
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  default:
    return false;
  }
}

bool MainWindow::Impl::HandleLaunchHelpCommand(int command_id) {
  switch (command_id) {
  case cmd::kOptionsRestartAdmin:
    RestartAsAdmin();
    return true;
  case cmd::kOptionsAlwaysRunAdmin:
    always_run_as_admin_ = !always_run_as_admin_;
    if (always_run_as_admin_) {
      always_run_as_system_ = false;
      always_run_as_trustedinstaller_ = false;
    }
    SaveSettings();
    BuildMenus();
    if (always_run_as_admin_ && !util::IsProcessElevated()) {
      RestartAsAdmin();
    }
    return true;
  case cmd::kOptionsRestartSystem:
    RestartAsSystem();
    return true;
  case cmd::kOptionsAlwaysRunSystem:
    always_run_as_system_ = !always_run_as_system_;
    if (always_run_as_system_) {
      always_run_as_admin_ = false;
      always_run_as_trustedinstaller_ = false;
    }
    SaveSettings();
    BuildMenus();
    if (always_run_as_system_ && !util::IsProcessSystem()) {
      RestartAsSystem();
    }
    return true;
  case cmd::kOptionsRestartTrustedInstaller:
    RestartAsTrustedInstaller();
    return true;
  case cmd::kOptionsAlwaysRunTrustedInstaller:
    always_run_as_trustedinstaller_ = !always_run_as_trustedinstaller_;
    if (always_run_as_trustedinstaller_) {
      always_run_as_admin_ = false;
      always_run_as_system_ = false;
    }
    SaveSettings();
    BuildMenus();
    if (always_run_as_trustedinstaller_ &&
        !util::IsProcessTrustedInstaller()) {
      RestartAsTrustedInstaller();
    }
    return true;
  case cmd::kOptionsOpenDefaultRegedit:
    OpenDefaultRegedit();
    return true;
  case cmd::kOptionsReplaceRegedit:
    ReplaceRegedit(!replace_regedit_);
    return true;
  case cmd::kOptionsSingleInstance:
    single_instance_ = !single_instance_;
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsHiveFileDir:
    OpenHiveFileDir();
    return true;
  case cmd::kHelpAbout:
    ui::ShowAbout(hwnd_);
    return true;
  case cmd::kHelpContents:
    ShellExecuteW(hwnd_, L"open", kHelpUrl, nullptr, nullptr, SW_SHOWNORMAL);
    return true;
  default:
    return false;
  }
}

} // namespace regkit
