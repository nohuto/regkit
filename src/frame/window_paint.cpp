// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::OnSize(int width, int height) {
  LayoutControls(width, height);
}

void MainWindow::Impl::OnPaint() {
  PAINTSTRUCT ps = {};
  HDC hdc = BeginPaint(hwnd_, &ps);
  RECT client = {};
  GetClientRect(hwnd_, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  if (width <= 0 || height <= 0) {
    EndPaint(hwnd_, &ps);
    return;
  }

  const Theme& theme = Theme::Current();
  HDC mem_dc = CreateCompatibleDC(hdc);
  HBITMAP buffer = CreateCompatibleBitmap(hdc, width, height);
  HGDIOBJ old_bitmap = SelectObject(mem_dc, buffer);

  FillRect(mem_dc, &client, theme.BackgroundBrush());

  HPEN pen = appearance::CachedPen(theme.BorderColor(), 1);
  HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(mem_dc, pen));
  HBRUSH old_brush = reinterpret_cast<HBRUSH>(SelectObject(mem_dc, GetStockObject(NULL_BRUSH)));

  RECT rect = {};
  auto draw_border = [&](HWND child) {
    if (!GetChildRectInParent(hwnd_, child, &rect)) {
      return;
    }
    DrawOutlineRect(mem_dc, rect, kBorderInflate);
  };
  auto draw_panel = [&](HWND header, HWND body) {
    if (!header || !body) {
      return;
    }
    RECT header_rect = {};
    RECT body_rect = {};
    if (!GetChildRectInParent(hwnd_, header, &header_rect)) {
      return;
    }
    if (!GetChildRectInParent(hwnd_, body, &body_rect)) {
      return;
    }
    RECT combined = {};
    combined.left = std::min(header_rect.left, body_rect.left);
    combined.top = std::min(header_rect.top, body_rect.top);
    combined.right = std::max(header_rect.right, body_rect.right);
    combined.bottom = std::max(header_rect.bottom, body_rect.bottom);
    DrawOutlineRect(mem_dc, combined, kBorderInflate);
    MoveToEx(mem_dc, combined.left, header_rect.bottom, nullptr);
    LineTo(mem_dc, combined.right, header_rect.bottom);
  };

  bool show_search = IsSearchTabSelected();
  if (show_value_ && !show_search) {
    draw_border(browse_.values().hwnd());
  }
  if (show_tree_ && !show_search) {
    draw_panel(tree_header_, browse_.tree().hwnd());
  }
  if (show_history_ && !show_search) {
    draw_panel(history_label_, history_list_);
  }
  if (show_tree_ && show_value_ && splitter_rect_.right > splitter_rect_.left) {
    RECT split = splitter_rect_;
    FillRect(mem_dc, &split, theme.PanelBrush());
    int mid_x = (split.left + split.right) / 2;
    MoveToEx(mem_dc, mid_x, split.top + 4, nullptr);
    LineTo(mem_dc, mid_x, split.bottom - 4);
  }
  if (show_history_ && history_splitter_rect_.bottom > history_splitter_rect_.top) {
    RECT split = history_splitter_rect_;
    FillRect(mem_dc, &split, theme.PanelBrush());
    int mid_y = (split.top + split.bottom) / 2;
    MoveToEx(mem_dc, split.left + 4, mid_y, nullptr);
    LineTo(mem_dc, split.right - 4, mid_y);
  }

  if (browse_.address() && browse_.go_button()) {
    RECT left = {};
    RECT right = {};
    if (GetChildRectInParent(hwnd_, browse_.address(), &left) && GetChildRectInParent(hwnd_, browse_.go_button(), &right)) {
      RECT combined = left;
      combined.right = right.right;
      DrawOutlineRect(mem_dc, combined, kBorderInflate);
    }
  }
  if (browse_.filter() && IsWindowVisible(browse_.filter())) {
    RECT filter_rect = {};
    if (GetChildRectInParent(hwnd_, browse_.filter(), &filter_rect)) {
      DrawOutlineRect(mem_dc, filter_rect, kBorderInflate);
    }
  }

  HPEN top_pen = appearance::CachedPen(theme.BorderColor(), 1);
  HGDIOBJ old_top = SelectObject(mem_dc, top_pen);
  MoveToEx(mem_dc, 0, 0, nullptr);
  LineTo(mem_dc, client.right, 0);
  SelectObject(mem_dc, old_top);

  SelectObject(mem_dc, old_brush);
  SelectObject(mem_dc, old_pen);

  BitBlt(hdc, 0, 0, width, height, mem_dc, 0, 0, SRCCOPY);
  SelectObject(mem_dc, old_bitmap);
  DeleteObject(buffer);
  DeleteDC(mem_dc);

  EndPaint(hwnd_, &ps);
}

void MainWindow::Impl::PaintMenuBarSeparator() {
  if (!hwnd_ || !GetMenu(hwnd_)) {
    return;
  }

  MENUBARINFO menu_info = {};
  menu_info.cbSize = sizeof(menu_info);
  if (!GetMenuBarInfo(hwnd_, OBJID_MENU, 0, &menu_info)) {
    return;
  }

  RECT window_rect = {};
  if (!GetWindowRect(hwnd_, &window_rect)) {
    return;
  }

  RECT separator = menu_info.rcBar;
  OffsetRect(&separator, -window_rect.left, -window_rect.top);
  separator.top = separator.bottom - 1;
  separator.bottom += 1;
  if (separator.bottom <= separator.top) {
    return;
  }

  HDC hdc = GetWindowDC(hwnd_);
  if (!hdc) {
    return;
  }
  FillRect(hdc, &separator, Theme::Current().BackgroundBrush());
  ReleaseDC(hwnd_, hdc);
}

void MainWindow::Impl::ApplyThemeToChildren() {
  const Theme& theme = Theme::Current();

  theme.ApplyToToolbar(toolbar_.hwnd());
  theme.ApplyToTreeView(browse_.tree().hwnd());
  theme.ApplyToTreeView(regedit_compat_tree_);
  theme.ApplyToListView(browse_.values().hwnd());
  theme.ApplyToListView(regedit_compat_list_);
  theme.ApplyToListView(history_list_);
  theme.ApplyToListView(search_results_list_);
  theme.ApplyToTabControl(tab_);
  theme.ApplyToStatusBar(status_bar_);

  if (browse_.address()) {
    SetWindowTheme(browse_.address(), Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(browse_.address(), 6, 6);
    SetEditVerticalRect(browse_.address(), ui_font_, 2, 6, 6);
  }
  if (regedit_compat_edit_) {
    SetWindowTheme(regedit_compat_edit_, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(regedit_compat_edit_, 6, 6);
    SetEditVerticalRect(regedit_compat_edit_, ui_font_, 2, 6, 6);
  }
  if (browse_.filter()) {
    SetWindowTheme(browse_.filter(), Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(browse_.filter(), 6, 6);
    SetEditVerticalRect(browse_.filter(), ui_font_, 2, 6, 6);
  }
  if (tree_header_) {
    SetWindowTheme(tree_header_, L"", L"");
  }
  ApplyAutoCompleteTheme();
  DrawMenuBar(hwnd_);
}

void MainWindow::Impl::ApplySystemTheme() {
  if (applying_theme_) {
    return;
  }
  applying_theme_ = true;
  Theme::UpdateFromSystem();
  Theme::Current().ApplyToWindow(hwnd_);
  ApplyThemeToChildren();
  ReloadThemeIcons();
  if (hwnd_) {
    InvalidateRect(hwnd_, nullptr, TRUE);
  }
  applying_theme_ = false;
}

void MainWindow::Impl::LoadThemePresets() {
  std::vector<ThemePreset> presets;
  bool loaded = ThemePresetStore::Load(&presets);
  bool updated_builtins = false;
  if (!loaded || presets.empty()) {
    presets = ThemePresetStore::BuiltInPresets();
  } else {
    std::vector<ThemePreset> builtins = ThemePresetStore::BuiltInPresets();
    auto same_colors = [](const ThemeColors& left, const ThemeColors& right) { return left.background == right.background && left.panel == right.panel && left.surface == right.surface && left.header == right.header && left.border == right.border && left.text == right.text && left.muted_text == right.muted_text && left.accent == right.accent && left.selection == right.selection && left.selection_text == right.selection_text && left.hover == right.hover && left.focus == right.focus; };
    auto same_preset = [&](const ThemePreset& left, const ThemePreset& right) { return left.is_dark == right.is_dark && same_colors(left.colors, right.colors); };
    for (const auto& builtin : builtins) {
      auto it = std::find_if(presets.begin(), presets.end(), [&](const ThemePreset& existing) { return _wcsicmp(existing.name.c_str(), builtin.name.c_str()) == 0; });
      if (it == presets.end()) {
        presets.push_back(builtin);
        updated_builtins = true;
      } else if (!same_preset(*it, builtin)) {
        *it = builtin;
        updated_builtins = true;
      }
    }
  }
  theme_presets_ = std::move(presets);
  if (theme_presets_.empty()) {
    return;
  }
  if (active_theme_preset_.empty()) {
    active_theme_preset_ = theme_presets_.front().name;
  }
  auto it = std::find_if(theme_presets_.begin(), theme_presets_.end(), [&](const ThemePreset& preset) { return _wcsicmp(preset.name.c_str(), active_theme_preset_.c_str()) == 0; });
  if (it == theme_presets_.end()) {
    active_theme_preset_ = theme_presets_.front().name;
  }
  if (!loaded || updated_builtins) {
    SaveThemePresets();
  }
}

void MainWindow::Impl::SaveThemePresets() const {
  ThemePresetStore::Save(theme_presets_, nullptr);
}

bool MainWindow::Impl::ApplyThemePresetByName(const std::wstring& name, bool persist) {
  if (theme_presets_.empty()) {
    return false;
  }
  auto it = std::find_if(theme_presets_.begin(), theme_presets_.end(), [&](const ThemePreset& preset) { return _wcsicmp(preset.name.c_str(), name.c_str()) == 0; });
  if (it == theme_presets_.end()) {
    it = theme_presets_.begin();
  }
  Theme::SetCustomColors(it->colors, it->is_dark);
  theme_mode_ = ThemeMode::kCustom;
  active_theme_preset_ = it->name;
  Theme::SetMode(theme_mode_);
  ApplySystemTheme();
  if (persist) {
    SaveSettings();
    BuildMenus();
  }
  return true;
}

void MainWindow::Impl::UpdateThemePresets(const std::vector<ThemePreset>& presets, const std::wstring& active_name, bool apply_now) {
  theme_presets_ = presets;
  active_theme_preset_ = active_name;
  SaveThemePresets();
  if (apply_now) {
    ApplyThemePresetByName(active_theme_preset_, true);
  } else {
    SaveSettings();
    BuildMenus();
  }
}

void MainWindow::Impl::ApplyAlwaysOnTop() {
  if (!hwnd_) {
    return;
  }
  SetWindowPos(hwnd_, always_on_top_ ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
}

void MainWindow::Impl::UpdateUIFont() {
  HFONT next_font = nullptr;
  bool next_owned = false;
  if (use_custom_font_) {
    next_font = CreateFontIndirectW(&custom_font_);
    next_owned = next_font != nullptr;
  } else {
    LOGFONTW lf = DefaultLogFont();
    next_font = CreateFontIndirectW(&lf);
    next_owned = next_font != nullptr;
  }
  if (!next_font) {
    next_font = CreateUIFont();
    next_owned = false;
  }
  if (ui_font_ && ui_font_owned_) {
    DeleteObject(ui_font_);
  }
  ui_font_ = next_font;
  ui_font_owned_ = next_owned;
  ApplyUIFontToControls();
}

void MainWindow::Impl::ApplyUIFontToControls() {
  if (!ui_font_) {
    return;
  }
  ApplyFont(toolbar_.hwnd(), ui_font_);
  ApplyFont(browse_.address(), ui_font_);
  ApplyFont(regedit_compat_edit_, ui_font_);
  ApplyFont(browse_.go_button(), ui_font_);
  ApplyFont(browse_.filter(), ui_font_);
  ApplyFont(tab_, ui_font_);
  ApplyFont(tree_header_, ui_font_);
  ApplyFont(tree_close_btn_, ui_font_);
  ApplyFont(browse_.tree().hwnd(), ui_font_);
  ApplyFont(regedit_compat_tree_, ui_font_);
  ApplyFont(browse_.values().hwnd(), ui_font_);
  ApplyFont(regedit_compat_list_, ui_font_);
  ApplyFont(history_close_btn_, ui_font_);
  ApplyFont(history_label_, ui_font_);
  ApplyFont(history_list_, ui_font_);
  ApplyFont(status_bar_, ui_font_);
  ApplyFont(search_results_list_, ui_font_);
  UpdateTabWidth();
  if (hwnd_) {
    DrawMenuBar(hwnd_);
  }
  InvalidateRect(hwnd_, nullptr, TRUE);
  if (hwnd_) {
    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    LayoutControls(rect.right, rect.bottom);
  }
}

} // namespace regkit
