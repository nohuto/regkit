// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>
#include <vector>

#include "appearance/theme.h"

namespace regkit {

struct ThemePreset {
  std::wstring name;
  ThemeColors colors;
  bool is_dark = true;
};

class ThemePresetStore {
public:
  static std::wstring PresetsPath();
  static std::vector<ThemePreset> BuiltInPresets();
  static bool Load(std::vector<ThemePreset>* presets, std::wstring* error = nullptr);
  static bool Save(const std::vector<ThemePreset>& presets, std::wstring* error = nullptr);
  static bool ImportFromFile(const std::wstring& path, std::vector<ThemePreset>* presets, std::wstring* error = nullptr);
  static bool ExportToFile(const std::wstring& path, const std::vector<ThemePreset>& presets, std::wstring* error = nullptr);
};

std::wstring FormatColorHex(COLORREF color);
bool ParseColorHex(const std::wstring& text, COLORREF* color);

} // namespace regkit
