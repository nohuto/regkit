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

void DialogResizer::Attach(HWND dialog, std::initializer_list<AnchorRule> rules) {
  items_.clear();
  client_ = {};
  min_window_ = {};
  if (!dialog) {
    return;
  }

  RECT client = {};
  RECT window = {};
  if (!GetClientRect(dialog, &client) || !GetWindowRect(dialog, &window)) {
    return;
  }
  client_ = {client.right - client.left, client.bottom - client.top};
  min_window_ = {window.right - window.left, window.bottom - window.top};

  items_.reserve(rules.size());
  for (const AnchorRule& rule : rules) {
    HWND control = GetDlgItem(dialog, rule.id);
    RECT rect = {};
    if (!control || !GetWindowRect(control, &rect)) {
      continue;
    }
    MapWindowPoints(nullptr, dialog, reinterpret_cast<POINT*>(&rect), 2);
    items_.push_back({rule.id, rule.anchors, rect});
  }
}

void DialogResizer::Apply(HWND dialog) const {
  if (!dialog || items_.empty() || client_.cx <= 0 || client_.cy <= 0) {
    return;
  }
  RECT client = {};
  if (!GetClientRect(dialog, &client)) {
    return;
  }
  const LONG dx = (client.right - client.left) - client_.cx;
  const LONG dy = (client.bottom - client.top) - client_.cy;

  HDWP defer = BeginDeferWindowPos(static_cast<int>(items_.size()));
  for (const Item& item : items_) {
    HWND control = GetDlgItem(dialog, item.id);
    if (!control) {
      continue;
    }
    const LONG left = (item.anchors & kAnchorLeft) ? item.rect.left : item.rect.left + dx;
    const LONG right = (item.anchors & kAnchorRight) ? item.rect.right + dx : item.rect.right;
    const LONG top = (item.anchors & kAnchorTop) ? item.rect.top : item.rect.top + dy;
    const LONG bottom = (item.anchors & kAnchorBottom) ? item.rect.bottom + dy : item.rect.bottom;
    const int width = static_cast<int>(std::max<LONG>(0, right - left));
    const int height = static_cast<int>(std::max<LONG>(0, bottom - top));
    if (defer) {
      defer = DeferWindowPos(defer, control, nullptr, static_cast<int>(left), static_cast<int>(top), width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    } else {
      SetWindowPos(control, nullptr, static_cast<int>(left), static_cast<int>(top), width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    }
  }
  if (defer) {
    EndDeferWindowPos(defer);
  }
  InvalidateRect(dialog, nullptr, TRUE);
}

void DialogResizer::ClampMinSize(MINMAXINFO* info) const {
  if (!info || min_window_.cx <= 0 || min_window_.cy <= 0) {
    return;
  }
  info->ptMinTrackSize.x = std::max<LONG>(info->ptMinTrackSize.x, min_window_.cx);
  info->ptMinTrackSize.y = std::max<LONG>(info->ptMinTrackSize.y, min_window_.cy);
}

bool DialogResizer::attached() const {
  return !items_.empty();
}

} // namespace regkit::appearance
