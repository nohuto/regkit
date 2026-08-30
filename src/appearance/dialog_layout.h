// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <initializer_list>
#include <vector>

namespace regkit::appearance {

void SetControlFont(HWND control, HFONT font);
void RestoreDialogOwner(HWND owner, bool* restored);
void PositionDialog(HWND dialog, HWND owner, int width, int height);

enum AnchorFlags : unsigned {
  kAnchorLeft = 1u,
  kAnchorTop = 2u,
  kAnchorRight = 4u,
  kAnchorBottom = 8u,
};

struct AnchorRule {
  int id = 0;
  unsigned anchors = kAnchorLeft | kAnchorTop;
};

class DialogResizer {
public:
  void Attach(HWND dialog, std::initializer_list<AnchorRule> rules);
  void Apply(HWND dialog) const;
  void ClampMinSize(MINMAXINFO* info) const;
  bool attached() const;

private:
  struct Item {
    int id = 0;
    unsigned anchors = 0;
    RECT rect = {};
  };

  std::vector<Item> items_;
  SIZE client_ = {};
  SIZE min_window_ = {};
};

} // namespace regkit::appearance
