// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>
#include <commctrl.h>

#include <string>
#include <vector>

namespace regkit {

struct ToolbarIcon {
  int resource_id = 0;
  std::wstring path;
};

class Toolbar {
public:
  void Create(HWND parent, HINSTANCE instance, int control_id);
  HWND hwnd() const;
  void LoadIcons(const std::vector<ToolbarIcon>& icons, int size, int glyph_size);
  void AddButtons(const std::vector<TBBUTTON>& buttons);

private:
  HWND hwnd_ = nullptr;
  HIMAGELIST image_list_ = nullptr;
};

} // namespace regkit
