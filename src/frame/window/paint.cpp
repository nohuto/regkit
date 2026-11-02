// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

#include <win32ui/dialog/autocomplete.h>
#include <win32ui/dialog/edit.h>
#include <win32ui/gdi/fonts.h>
#include <win32ui/window_metrics.h>

namespace regkit
{
using namespace window_detail;

void MainWindow::Impl::OnSize(int width, int height)
{
    LayoutControls(width, height);
}

void MainWindow::Impl::ApplyThemeToChildren()
{
    menu_bar_.ApplyTheme();
    win32ui::ReapplyTheme(hwnd_);

    if (browse_.address())
    {
        win32ui::SetEditPadding(browse_.address(), 6, 2);
    }
    if (browse_.filter())
    {
        win32ui::SetEditPadding(browse_.filter(), 6, 2);
    }
    if (tree_header_)
    {
        SetWindowTheme(tree_header_, L"", L"");
    }
    win32ui::ApplyAutoCompleteTheme();
    DrawMenuBar(hwnd_);
}

void MainWindow::Impl::ApplySystemTheme()
{
    if (applying_theme_)
    {
        return;
    }
    applying_theme_ = true;
    ApplyThemeToChildren();
    ReloadThemeIcons();
    if (hwnd_)
    {
        InvalidateRect(hwnd_, nullptr, TRUE);
    }
    applying_theme_ = false;
}

void MainWindow::Impl::LoadThemePresets()
{
    std::wstring load_error;
    // file that cant be used stays as it is until the user saves presets
    const win32ui::ThemePresetLoad load = ThemePresetStore::Load(&theme_presets_, &settings_.theme_preset, &load_error);
    if (!load_error.empty())
    {
        ReportFileError(ThemePresetStore::Path(), load_error);
    }
    if (load.save_needed)
    {
        SaveThemePresets();
    }
}

void MainWindow::Impl::SaveThemePresets() const
{
    ThemePresetStore::Save(theme_presets_, nullptr);
}

bool MainWindow::Impl::ApplyThemePresetByName(const std::wstring& name, bool persist)
{
    if (theme_presets_.empty())
    {
        return false;
    }
    const ThemePreset* it = FindThemePreset(theme_presets_, name);
    theme_mode_ = win32ui::ThemeMode::custom;
    settings_.theme_preset = it->name;
    win32ui::SetThemeMode(theme_mode_, &it->colors);
    ApplySystemTheme();
    if (persist)
    {
        SaveSettings();
    }
    return true;
}

void MainWindow::Impl::UpdateThemePresets(const std::vector<ThemePreset>& presets, const std::wstring& active_name, bool apply_now)
{
    theme_presets_ = presets;
    settings_.theme_preset = active_name;
    SaveThemePresets();
    if (apply_now)
    {
        ApplyThemePresetByName(settings_.theme_preset, true);
    }
    else
    {
        SaveSettings();
    }
}

void MainWindow::Impl::ApplyAlwaysOnTop()
{
    if (!hwnd_)
    {
        return;
    }
    SetWindowPos(hwnd_, settings_.always_on_top ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
}

void MainWindow::Impl::UpdateUIFont()
{
    win32ui::SetCustomFont(settings_.use_custom_font ? &custom_font_ : nullptr);
    const LOGFONTW lf = win32ui::UIFontLogFont(win32ui::DpiForWindow(hwnd_));
    HFONT next_font = CreateFontIndirectW(&lf);
    bool next_owned = next_font != nullptr;
    if (!next_font)
    {
        next_font = CreateUIFont();
        next_owned = false;
    }
    if (ui_font_ && ui_font_owned_)
    {
        DeleteObject(ui_font_);
    }
    ui_font_ = next_font;
    ui_font_owned_ = next_owned;
    ApplyUIFontToControls();
}

void MainWindow::Impl::ApplyUIFontToControls()
{
    if (!ui_font_)
    {
        return;
    }
    ApplyFont(toolbar_, ui_font_);
    ApplyFont(browse_.address(), ui_font_);
    ApplyFont(browse_.go_button(), ui_font_);
    ApplyFont(browse_.filter(), ui_font_);
    ApplyFont(tab_, ui_font_);
    ApplyFont(tree_header_, ui_font_);
    ApplyFont(tree_close_btn_, ui_font_);
    for (const auto& tree : browse_.trees())
    {
        ApplyFont(tree->hwnd(), ui_font_);
    }
    ApplyFont(browse_.values().hwnd(), ui_font_);
    ApplyFont(history_close_btn_, ui_font_);
    ApplyFont(history_label_, ui_font_);
    ApplyFont(history_list_, ui_font_);
    ApplyFont(status_bar_, ui_font_);
    ApplyFont(search_results_list_, ui_font_);
    if (tab_)
    {
        win32ui::RefreshTabStrip(tab_);
        tab_height_ = win32ui::TabStripRowHeight(tab_);
    }
    win32ui::RefreshStatusBar(status_bar_);
    menu_bar_.ApplyTheme();
    InvalidateRect(hwnd_, nullptr, TRUE);
    if (hwnd_)
    {
        RECT rect = {};
        GetClientRect(hwnd_, &rect);
        LayoutControls(rect.right, rect.bottom);
    }
}

}
