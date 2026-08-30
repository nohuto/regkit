// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/window_metrics.h"

#include <algorithm>

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

void ClampToWorkArea(RECT* rect) {
  if (!rect || rect->right <= rect->left || rect->bottom <= rect->top) {
    return;
  }

  RECT work = {};
  MONITORINFO info = {};
  info.cbSize = sizeof(info);
  const HMONITOR monitor = MonitorFromRect(rect, MONITOR_DEFAULTTONEAREST);
  if (monitor && GetMonitorInfoW(monitor, &info)) {
    work = info.rcWork;
  } else if (!SystemParametersInfoW(SPI_GETWORKAREA, 0, &work, 0)) {
    return;
  }
  if (work.right <= work.left || work.bottom <= work.top) {
    return;
  }

  const LONG width = std::min(rect->right - rect->left, work.right - work.left);
  const LONG height =
      std::min(rect->bottom - rect->top, work.bottom - work.top);
  const LONG left = std::clamp(rect->left, work.left, work.right - width);
  const LONG top = std::clamp(rect->top, work.top, work.bottom - height);
  *rect = {left, top, left + width, top + height};
}

} // namespace regkit::win32
