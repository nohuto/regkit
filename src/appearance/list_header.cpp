// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "appearance/list_header.h"

#include "appearance/gdi_cache.h"
#include "appearance/theme.h"

#include <commctrl.h>
#include <uxtheme.h>
#include <vsstyle.h>

namespace regkit::appearance {

namespace {

constexpr wchar_t kHeaderThemeProp[] = L"RegKitHeaderTheme";
constexpr int kHeaderTextPadding = 8;

HTHEME HeaderTheme(HWND header) {
  HTHEME cached = reinterpret_cast<HTHEME>(GetPropW(header, kHeaderThemeProp));
  if (!cached) {
    cached = OpenThemeData(header, VSCLASS_HEADER);
    SetPropW(header, kHeaderThemeProp, cached);
  }
  return cached;
}

HBRUSH ListSurfaceBrush(HWND header) {
  HWND list = GetParent(header);
  const COLORREF color = list ? ListView_GetBkColor(list) : CLR_NONE;
  if (color == CLR_NONE || color == CLR_DEFAULT) {
    return Theme::Current().PanelBrush();
  }
  return CachedBrush(color);
}

} // namespace

void PaintListHeader(HWND header, HFONT font) {
  if (!header) {
    return;
  }
  PAINTSTRUCT ps = {};
  HDC target = BeginPaint(header, &ps);
  if (!target) {
    return;
  }
  HDC hdc = nullptr;
  HPAINTBUFFER buffer = BeginBufferedPaint(target, &ps.rcPaint, BPBF_COMPATIBLEBITMAP, nullptr, &hdc);
  if (!hdc) {
    hdc = target;
  }

  const Theme& theme = Theme::Current();
  HBRUSH surface = ListSurfaceBrush(header);
  RECT client = {};
  GetClientRect(header, &client);
  FillRect(hdc, &client, surface);

  if (!font) {
    font = reinterpret_cast<HFONT>(SendMessageW(header, WM_GETFONT, 0, 0));
  }
  HFONT old_font = font ? reinterpret_cast<HFONT>(SelectObject(hdc, font)) : nullptr;

  HTHEME header_theme = HeaderTheme(header);
  SIZE arrow_size = {0, 0};
  if (header_theme) {
    GetThemePartSize(header_theme, hdc, HP_HEADERSORTARROW, HSAS_SORTEDUP, nullptr, TS_TRUE, &arrow_size);
  }
  if (arrow_size.cx <= 0 || arrow_size.cy <= 0) {
    arrow_size.cx = 8;
    arrow_size.cy = 8;
  }

  const int count = Header_GetItemCount(header);
  for (int i = 0; i < count; ++i) {
    RECT rect = {};
    if (!Header_GetItemRect(header, i, &rect)) {
      continue;
    }
    RECT visible = {};
    if (!IntersectRect(&visible, &rect, &ps.rcPaint)) {
      continue;
    }

    wchar_t text[128] = {};
    HDITEMW item = {};
    item.mask = HDI_TEXT | HDI_FORMAT;
    item.pszText = text;
    item.cchTextMax = static_cast<int>(_countof(text));
    Header_GetItem(header, i, &item);

    const bool sorted_up = (item.fmt & HDF_SORTUP) != 0;
    const bool sorted_down = (item.fmt & HDF_SORTDOWN) != 0;

    FillRect(hdc, &rect, surface);

    RECT divider = {rect.right - 1, rect.top, rect.right, rect.bottom};
    FillRect(hdc, &divider, CachedBrush(theme.BorderColor()));

    RECT text_rect = rect;
    text_rect.left += kHeaderTextPadding;
    text_rect.right -= kHeaderTextPadding;

    UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
    if (item.fmt & HDF_RIGHT) {
      format |= DT_RIGHT;
    } else if (item.fmt & HDF_CENTER) {
      format |= DT_CENTER;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, theme.TextColor());
    DrawTextW(hdc, text, -1, &text_rect, format);

    if ((sorted_up || sorted_down) && header_theme) {
      RECT arrow_rect = rect;
      arrow_rect.right -= 1;
      arrow_rect.bottom = rect.top + arrow_size.cy;
      const int arrow_state = sorted_up ? HSAS_SORTEDUP : HSAS_SORTEDDOWN;
      DrawThemeBackground(header_theme, hdc, HP_HEADERSORTARROW, arrow_state, &arrow_rect, nullptr);
    }
  }

  if (old_font) {
    SelectObject(hdc, old_font);
  }
  if (buffer) {
    EndBufferedPaint(buffer, TRUE);
  }
  EndPaint(header, &ps);
}

void ReleaseListHeaderTheme(HWND header) {
  if (!header) {
    return;
  }
  if (HTHEME cached = reinterpret_cast<HTHEME>(GetPropW(header, kHeaderThemeProp))) {
    CloseThemeData(cached);
  }
  RemovePropW(header, kHeaderThemeProp);
}

} // namespace regkit::appearance
