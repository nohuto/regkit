// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"

#include <array>

namespace regkit {
using namespace command_detail;

namespace {
HFONT SubmenuArrowFont(int size) {
  static std::array<HFONT, 12> fonts = {};
  if (size < 1 || size > static_cast<int>(fonts.size())) {
    return nullptr;
  }
  HFONT& font = fonts[static_cast<size_t>(size) - 1];
  if (!font) {
    font = CreateFontW(-size, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_DONTCARE, L"Marlett");
  }
  return font;
}
} // namespace

void MainWindow::Impl::PrepareMenusForOwnerDraw(HMENU menu, bool is_menu_bar) {
  if (!menu) {
    return;
  }
  HDC hdc = GetDC(hwnd_);
  HFONT old_font = nullptr;
  if (hdc && ui_font_) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }

  auto prepare = [&](auto&& self, HMENU current, bool menu_bar) -> void {
    if (menu_bar) {
      MENUINFO menu_info = {};
      menu_info.cbSize = sizeof(menu_info);
      menu_info.fMask = MIM_BACKGROUND;
      menu_info.hbrBack = Theme::Current().BackgroundBrush();
      SetMenuInfo(current, &menu_info);
    }

    int count = GetMenuItemCount(current);
    for (int i = 0; i < count; ++i) {
      MENUITEMINFOW info = {};
      info.cbSize = sizeof(info);
      info.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_SUBMENU | MIIM_ID;
      wchar_t text[256] = {};
      info.dwTypeData = text;
      info.cch = static_cast<UINT>(_countof(text));
      if (!GetMenuItemInfoW(current, i, TRUE, &info)) {
        continue;
      }

      if (menu_bar) {
        auto data = std::make_unique<MenuItemData>();
        data->text = text;
        size_t tab_pos = data->text.find(L'\t');
        if (tab_pos != std::wstring::npos) {
          data->left_text = data->text.substr(0, tab_pos);
          data->right_text = data->text.substr(tab_pos + 1);
        } else {
          data->left_text = data->text;
        }
        data->separator = (info.fType & MFT_SEPARATOR) != 0;
        data->has_submenu = info.hSubMenu != nullptr;
        data->is_menu_bar = menu_bar;
        if (data->separator) {
          data->width = 4;
          data->height = 8;
        } else {
          SIZE left_size = {};
          SIZE right_size = {};
          if (hdc) {
            GetTextExtentPoint32W(hdc, data->left_text.c_str(), static_cast<int>(data->left_text.size()), &left_size);
            if (!data->right_text.empty()) {
              GetTextExtentPoint32W(hdc, data->right_text.c_str(), static_cast<int>(data->right_text.size()), &right_size);
            }
          }
          int height = data->is_menu_bar ? 18 : 22;
          int padding = data->is_menu_bar ? 2 : 12;
          int shortcut_gap = (!data->is_menu_bar && !data->right_text.empty()) ? 24 : 0;
          int extra = (!data->is_menu_bar && data->has_submenu) ? 12 : 6;
          data->height = height;
          data->width = static_cast<int>(left_size.cx) + static_cast<int>(right_size.cx) + padding + shortcut_gap + extra;
        }

        MenuItemData* raw = data.get();
        menu_items_.push_back(std::move(data));

        info.fMask = MIIM_FTYPE | MIIM_DATA;
        info.fType |= MFT_OWNERDRAW;
        info.dwItemData = reinterpret_cast<ULONG_PTR>(raw);
        SetMenuItemInfoW(current, i, TRUE, &info);
      }

      if (info.hSubMenu) {
        if (menu_bar) {
          self(self, info.hSubMenu, false);
        }
      }
    }
  };

  prepare(prepare, menu, is_menu_bar);

  if (hdc && old_font) {
    SelectObject(hdc, old_font);
  }
  if (hdc) {
    ReleaseDC(hwnd_, hdc);
  }
}

void MainWindow::Impl::OnMeasureMenuItem(MEASUREITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  auto* data = reinterpret_cast<MenuItemData*>(info->itemData);
  if (!data) {
    return;
  }
  if (data->width > 0 && data->height > 0) {
    info->itemWidth = static_cast<UINT>(data->width);
    info->itemHeight = static_cast<UINT>(data->height);
    return;
  }
  if (data->separator) {
    info->itemHeight = 8;
    info->itemWidth = 4;
    return;
  }

  HDC hdc = GetDC(hwnd_);
  HFONT old = nullptr;
  if (ui_font_) {
    old = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }
  SIZE size = {};
  GetTextExtentPoint32W(hdc, data->text.c_str(), static_cast<int>(data->text.size()), &size);
  if (old) {
    SelectObject(hdc, old);
  }
  ReleaseDC(hwnd_, hdc);

  int height = data->is_menu_bar ? 18 : 22;
  int padding = data->is_menu_bar ? 2 : 12;
  int extra = (!data->is_menu_bar && data->has_submenu) ? 12 : 6;
  info->itemHeight = height;
  info->itemWidth = size.cx + padding + extra;
}

void MainWindow::Impl::OnDrawMenuItem(const DRAWITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  auto* data = reinterpret_cast<MenuItemData*>(info->itemData);
  if (!data) {
    return;
  }
  const Theme& theme = Theme::Current();
  HDC hdc = info->hDC;
  RECT rect = info->rcItem;

  if (data->separator) {
    HPEN pen = appearance::CachedPen(theme.BorderColor());
    HPEN old = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
    int y = (rect.top + rect.bottom) / 2;
    MoveToEx(hdc, rect.left + 8, y, nullptr);
    LineTo(hdc, rect.right - 8, y);
    SelectObject(hdc, old);
    return;
  }

  bool selected = (info->itemState & (ODS_SELECTED | ODS_HOTLIGHT)) != 0;
  bool disabled = (info->itemState & ODS_DISABLED) != 0;
  bool checked = (info->itemState & ODS_CHECKED) != 0;
  COLORREF bg = data->is_menu_bar ? theme.BackgroundColor() : theme.PanelColor();
  COLORREF fg = theme.TextColor();
  if (selected) {
    if (data->is_menu_bar) {
      bg = theme.HoverColor();
    } else {
      bg = theme.SelectionColor();
      fg = theme.SelectionTextColor();
    }
  } else if (disabled) {
    fg = theme.MutedTextColor();
  }

  HBRUSH bg_brush = nullptr;
  if (selected) {
    bg_brush = appearance::CachedBrush(bg);
  } else {
    bg_brush = data->is_menu_bar ? theme.BackgroundBrush() : theme.PanelBrush();
  }
  FillRect(hdc, &rect, bg_brush);

  SetBkMode(hdc, TRANSPARENT);
  SetTextColor(hdc, fg);
  HFONT old_font = nullptr;
  if (ui_font_) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }
  RECT text_rect = rect;
  int left_padding = data->is_menu_bar ? 0 : 12;
  int right_padding = data->is_menu_bar ? 0 : (data->has_submenu ? 12 : 6);
  text_rect.left += left_padding;
  text_rect.right -= right_padding;
  if (checked && !data->is_menu_bar) {
    int mid_y = (rect.top + rect.bottom) / 2;
    HPEN pen = appearance::CachedPen(fg, 1);
    HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
    MoveToEx(hdc, rect.left + 8, mid_y, nullptr);
    LineTo(hdc, rect.left + 11, mid_y + 3);
    LineTo(hdc, rect.left + 16, mid_y - 3);
    SelectObject(hdc, old_pen);
  }
  UINT format = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS;
  if (data->is_menu_bar) {
    format |= DT_CENTER;
  }
  if (!data->right_text.empty() && !data->is_menu_bar) {
    SIZE right_size = {};
    GetTextExtentPoint32W(hdc, data->right_text.c_str(), static_cast<int>(data->right_text.size()), &right_size);
    RECT right_rect = rect;
    right_rect.right -= right_padding;
    right_rect.left = right_rect.right - right_size.cx;
    RECT left_rect = text_rect;
    left_rect.right = right_rect.left - 12;
    DrawTextW(hdc, data->left_text.c_str(), -1, &left_rect, format);
    DrawTextW(hdc, data->right_text.c_str(), -1, &right_rect, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_RIGHT);
  } else {
    DrawTextW(hdc, data->left_text.c_str(), -1, &text_rect, format);
  }
  if (old_font) {
    SelectObject(hdc, old_font);
  }

  if (data->has_submenu && !data->is_menu_bar) {
    int rect_h = rect.bottom - rect.top;
    int arrow_size = rect_h - 8;
    if (arrow_size < 6) {
      arrow_size = 6;
    } else if (arrow_size > 10) {
      arrow_size = 10;
    }
    RECT arrow_rect = rect;
    arrow_rect.right = rect.right - 6;
    arrow_rect.left = arrow_rect.right - arrow_size;
    arrow_rect.top = rect.top + (rect_h - arrow_size) / 2;
    arrow_rect.bottom = arrow_rect.top + arrow_size;
    HFONT arrow_font = SubmenuArrowFont(arrow_size);
    HFONT old_arrow = nullptr;
    if (arrow_font) {
      old_arrow = reinterpret_cast<HFONT>(SelectObject(hdc, arrow_font));
    }
    COLORREF arrow_color = disabled ? theme.MutedTextColor() : fg;
    SetTextColor(hdc, arrow_color);
    DrawTextW(hdc, L"8", -1, &arrow_rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX | DT_NOCLIP);
    if (old_arrow) {
      SelectObject(hdc, old_arrow);
    }
  }
}

} // namespace regkit
