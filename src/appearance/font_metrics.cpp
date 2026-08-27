// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "appearance/font_metrics.h"

namespace regkit::appearance {
namespace {

int VerticalDpi() {
  HDC dc = GetDC(nullptr);
  const int dpi = dc ? GetDeviceCaps(dc, LOGPIXELSY) : 96;
  if (dc) {
    ReleaseDC(nullptr, dc);
  }
  return dpi > 0 ? dpi : 96;
}

} // namespace

int FontPointSize(const LOGFONTW& font, int zero_height_fallback) {
  if (font.lfHeight == 0) {
    return zero_height_fallback;
  }
  return MulDiv(-font.lfHeight, 72, VerticalDpi());
}

int FontHeight(int point_size) {
  return -MulDiv(point_size, VerticalDpi(), 72);
}

} // namespace regkit::appearance
