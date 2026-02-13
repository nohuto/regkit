// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#pragma once

#include <string>
#include <vector>

#include "app/theme.h"

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
