// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_impl.h"

#include "ui/feedback.h"
#include "win32/file_dialog.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/dialog/text_prompt.h>
#include <win32ui/system/shell_open.h>
#include <win32ui/theme/preset_editor.h>

namespace regkit
{

namespace
{

constexpr wchar_t kThemeFilter[] = L"RegKit Theme Presets (*.rktheme)\0*.rktheme\0All Files (*.*)\0*.*\0";

}

void MainWindow::Impl::ShowThemePresetsDialog()
{
    if (theme_presets_.empty())
    {
        LoadThemePresets();
    }

    win32ui::PresetEditorHooks hooks;
    hooks.prompt_name = [](HWND owner, const wchar_t* title, const std::wstring& initial, std::wstring* name) {
        std::wstring text = initial;
        if (win32ui::PromptText(owner, {.title = title, .label = util::Tr(L"Preset name:")}, &text) != 0)
        {
            return false;
        }
        // preset reader trims names
        std::wstring trimmed = util::TrimWhitespace(text);
        if (trimmed.empty())
        {
            ui::ShowError(owner, util::Tr(L"Preset name can't be empty."));
            return false;
        }
        *name = std::move(trimmed);
        return true;
    };
    hooks.confirm_delete = [](HWND owner, const std::wstring& name) { return ui::ConfirmDelete(owner, util::Tr(L"Delete Preset"), name); };
    hooks.import_presets = [](HWND owner, std::vector<ThemePreset>* presets) {
        std::wstring path;
        if (!ui::PromptOpenFile(owner, kThemeFilter, &path))
        {
            return false;
        }
        std::wstring error;
        if (!ThemePresetStore::ImportFromFile(path, presets, &error))
        {
            ui::ShowError(owner, error.empty() ? util::Tr(L"Failed to import theme presets.") : error);
            return false;
        }
        return true;
    };
    hooks.export_presets = [](HWND owner, const std::vector<ThemePreset>& presets) {
        std::wstring path;
        win32::OpenAfter open_after = win32::OpenAfter::kNone;
        if (!ui::ReportFileDialogResult(owner, win32::ChooseFileToSave(owner, kThemeFilter, nullptr, &path, &open_after)))
        {
            return;
        }
        std::wstring error;
        if (!ThemePresetStore::ExportToFile(path, presets, &error))
        {
            ui::ShowError(owner, error.empty() ? util::Tr(L"Failed to export theme presets.") : error);
        }
        else if (open_after == win32::OpenAfter::kEditor)
        {
            ui::ReportFileDialogResult(owner, win32ui::ShellOpen(owner, path.c_str(), L".txt"));
        }
    };
    hooks.apply = [](void* context, const std::vector<ThemePreset>& presets, const std::wstring& active_name) {
        static_cast<MainWindow::Impl*>(context)->UpdateThemePresets(presets, active_name, true);
    };
    hooks.context = this;
    win32ui::ShowThemePresetEditor(hwnd_, theme_presets_, settings_.theme_preset, &hooks);
}

}
