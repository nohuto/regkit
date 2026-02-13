// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#pragma once

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
