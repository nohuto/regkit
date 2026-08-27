// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/window_metrics.h"

namespace regkit::win32 {

UINT DpiForWindow(HWND window) {
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (user32) {
    auto get_window_dpi = reinterpret_cast<UINT(WINAPI*)(HWND)>(
        GetProcAddress(user32, "GetDpiForWindow"));
    if (get_window_dpi && window) {
      return get_window_dpi(window);
    }
    auto get_system_dpi = reinterpret_cast<UINT(WINAPI*)()>(
        GetProcAddress(user32, "GetDpiForSystem"));
    if (get_system_dpi) {
      return get_system_dpi();
    }
  }

  HDC dc = GetDC(window);
  const int dpi = dc ? GetDeviceCaps(dc, LOGPIXELSX) : 96;
  if (dc) {
    ReleaseDC(window, dc);
  }
  return dpi > 0 ? static_cast<UINT>(dpi) : 96;
}

} // namespace regkit::win32
