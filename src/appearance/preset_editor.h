// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>
#include <vector>

#include "appearance/presets.h"

namespace regkit::appearance {

using ThemePresetApply = void (*)(void* context,
                                  const std::vector<ThemePreset>& presets,
                                  const std::wstring& active_name);
using ThemePresetNamePrompt = bool (*)(void* context,
                                       HWND owner,
                                       const wchar_t* title,
                                       const std::wstring& initial,
                                       std::wstring* name);

void ShowThemePresetEditor(HWND owner,
                           const std::vector<ThemePreset>& presets,
                           const std::wstring& active_name,
                           ThemePresetApply apply,
                           ThemePresetNamePrompt prompt_name,
                           void* context);

} // namespace regkit::appearance
