// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "ui/presets.h"

#include "win32/shell_paths.h"
#include "win32/translation.h"

namespace regkit
{

namespace
{

std::wstring PresetsPath()
{
    const std::wstring folder = util::GetAppDataFolder();
    return folder.empty() ? std::wstring() : util::JoinPath(folder, L"theme_presets.rktheme");
}

std::wstring ProblemText(const win32ui::ThemePresetResult& result)
{
    const std::wstring text = win32ui::ThemePresetErrorText(result);
    const bool parse = result.error == win32ui::ThemePresetError::bad_line || result.error == win32ui::ThemePresetError::missing_name;
    return parse ? util::TrDetail(L"The theme preset file contains an entry RegKit can't parse.", text) : text;
}

bool WritePresetFile(const std::wstring& path, const std::vector<ThemePreset>& presets, std::wstring* error)
{
    const win32ui::ThemePresetResult result = win32ui::WriteThemePresets(path, presets);
    if (result.error == win32ui::ThemePresetError::none)
    {
        return true;
    }
    if (error)
    {
        *error = util::TrDetail(L"Failed to write the theme preset file.", win32ui::ThemePresetErrorText(result));
    }
    return false;
}

}

std::wstring ThemePresetStore::Path()
{
    return PresetsPath();
}

win32ui::ThemePresetLoad ThemePresetStore::Load(std::vector<ThemePreset>* presets, std::wstring* active_name, std::wstring* error)
{
    const win32ui::ThemePresetLoad load = win32ui::LoadThemePresetFile(PresetsPath(), presets, active_name);
    if (error)
    {
        *error = load.result.error == win32ui::ThemePresetError::none ? std::wstring() : ProblemText(load.result);
    }
    return load;
}

bool ThemePresetStore::Save(const std::vector<ThemePreset>& presets, std::wstring* error)
{
    return WritePresetFile(PresetsPath(), presets, error);
}

bool ThemePresetStore::ImportFromFile(const std::wstring& path, std::vector<ThemePreset>* presets, std::wstring* error)
{
    const win32ui::ThemePresetResult result = win32ui::ReadThemePresets(path, presets);
    if (result.error == win32ui::ThemePresetError::none && !presets->empty())
    {
        return true;
    }
    if (error)
    {
        switch (result.error)
        {
        case win32ui::ThemePresetError::none:
            *error = util::Tr(L"No theme presets were found in the file.");
            break;
        case win32ui::ThemePresetError::read_failed:
        case win32ui::ThemePresetError::too_large:
        case win32ui::ThemePresetError::bad_encoding:
            *error = util::TrDetail(L"Failed to open the theme preset file.", ProblemText(result));
            break;
        default:
            *error = ProblemText(result);
            break;
        }
    }
    return false;
}

bool ThemePresetStore::ExportToFile(const std::wstring& path, const std::vector<ThemePreset>& presets, std::wstring* error)
{
    return WritePresetFile(path, presets, error);
}

}
