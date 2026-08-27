// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "appearance/dialog_layout.h"

#include <algorithm>

namespace regkit::appearance {

void SetControlFont(HWND control, HFONT font) {
  if (control && font) {
    SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
  }
}

void RestoreDialogOwner(HWND owner, bool* restored) {
  if (!owner || !restored || *restored) {
    return;
  }
  EnableWindow(owner, TRUE);
  SetActiveWindow(owner);
  SetForegroundWindow(owner);
  *restored = true;
}

void PositionDialog(HWND dialog, HWND owner, int width, int height) {
  RECT owner_rect = {};
  if (owner && GetWindowRect(owner, &owner_rect)) {
    const int owner_width = owner_rect.right - owner_rect.left;
    const int owner_height = owner_rect.bottom - owner_rect.top;
    const int x = owner_rect.left + std::max(0, (owner_width - width) / 2);
    const int y = owner_rect.top + std::max(0, (owner_height - height) / 2);
    SetWindowPos(dialog, nullptr, x, y, width, height,
                 SWP_NOZORDER | SWP_NOACTIVATE);
    return;
  }
  SetWindowPos(dialog, nullptr, CW_USEDEFAULT, CW_USEDEFAULT, width, height,
               SWP_NOZORDER | SWP_NOACTIVATE);
}

} // namespace regkit::appearance
