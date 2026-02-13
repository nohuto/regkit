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

#include <windows.h>

#include "win32/win32_helpers.h"

namespace regkit {

enum class ThemeMode {
  kSystem = 0,
  kLight,
  kDark,
  kCustom,
};

struct ThemeColors {
  COLORREF background = RGB(0, 0, 0);
  COLORREF panel = RGB(0, 0, 0);
  COLORREF surface = RGB(0, 0, 0);
  COLORREF header = RGB(0, 0, 0);
  COLORREF border = RGB(0, 0, 0);
  COLORREF text = RGB(0, 0, 0);
  COLORREF muted_text = RGB(0, 0, 0);
  COLORREF accent = RGB(0, 0, 0);
  COLORREF selection = RGB(0, 0, 0);
  COLORREF selection_text = RGB(0, 0, 0);
  COLORREF hover = RGB(0, 0, 0);
  COLORREF focus = RGB(0, 0, 0);
};

class Theme {
public:
  static Theme& Dark();
  static Theme& Light();
  static Theme& Custom();
  static Theme& Current();
  static void SetCustomColors(const ThemeColors& colors, bool is_dark);
  static void SetMode(ThemeMode mode);
  static ThemeMode Mode();
  static bool UseDarkMode();
  static bool IsSystemDarkMode();
  static bool UpdateFromSystem();
  static void InitializeDarkModeSupport();

  void ApplyToWindow(HWND hwnd) const;
  void ApplyToChildren(HWND hwnd) const;
  void ApplyToTreeView(HWND hwnd) const;
  void ApplyToListView(HWND hwnd) const;
  void ApplyToTabControl(HWND hwnd) const;
  void ApplyToToolbar(HWND hwnd) const;
  void ApplyToComboBox(HWND hwnd) const;
  void ApplyToStatusBar(HWND hwnd) const;

  HBRUSH ControlColor(HDC hdc, HWND target, int type) const;

  HBRUSH BackgroundBrush() const;
  HBRUSH PanelBrush() const;
  HBRUSH SurfaceBrush() const;
  HBRUSH HeaderBrush() const;

  COLORREF BackgroundColor() const;
  COLORREF PanelColor() const;
  COLORREF SurfaceColor() const;
  COLORREF HeaderColor() const;
  COLORREF BorderColor() const;
  COLORREF TextColor() const;
  COLORREF MutedTextColor() const;
  COLORREF SelectionColor() const;
  COLORREF SelectionTextColor() const;
  COLORREF HoverColor() const;
  COLORREF FocusColor() const;

private:
  explicit Theme(const ThemeColors& colors, bool is_dark);

  ThemeColors colors_;
  bool is_dark_ = true;
  util::UniqueGdiObject<HBRUSH> background_brush_;
  util::UniqueGdiObject<HBRUSH> panel_brush_;
  util::UniqueGdiObject<HBRUSH> surface_brush_;
  util::UniqueGdiObject<HBRUSH> header_brush_;
};

void EnableImmersiveDarkMode(HWND hwnd, bool enabled);
void AllowDarkModeForWindow(HWND hwnd, bool enabled);

} // namespace regkit
