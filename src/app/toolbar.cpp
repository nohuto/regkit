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

#include "app/toolbar.h"

#include "win32/icon_resources.h"

namespace regkit {

namespace {

UINT GetWindowDpi(HWND hwnd) {
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (user32) {
    auto get_dpi_for_window = reinterpret_cast<UINT(WINAPI*)(HWND)>(GetProcAddress(user32, "GetDpiForWindow"));
    if (get_dpi_for_window && hwnd) {
      return get_dpi_for_window(hwnd);
    }
    auto get_dpi_for_system = reinterpret_cast<UINT(WINAPI*)()>(GetProcAddress(user32, "GetDpiForSystem"));
    if (get_dpi_for_system) {
      return get_dpi_for_system();
    }
  }
  HDC hdc = GetDC(hwnd);
  int dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSX) : 96;
  if (hdc) {
    ReleaseDC(hwnd, hdc);
  }
  return dpi > 0 ? static_cast<UINT>(dpi) : 96;
}

} // namespace

void Toolbar::Create(HWND parent, HINSTANCE instance, int control_id) {
  hwnd_ = CreateWindowExW(0, TOOLBARCLASSNAMEW, nullptr, WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER, 0, 0, 0, 0, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)), instance, nullptr);
  SendMessageW(hwnd_, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
  SendMessageW(hwnd_, TB_SETMAXTEXTROWS, 0, 0);
  SendMessageW(hwnd_, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DOUBLEBUFFER);
}

HWND Toolbar::hwnd() const {
  return hwnd_;
}

void Toolbar::LoadIcons(const std::vector<ToolbarIcon>& icons, int size, int glyph_size) {
  if (image_list_) {
    ImageList_Destroy(image_list_);
    image_list_ = nullptr;
  }

  UINT dpi = GetWindowDpi(hwnd_);
  int base_icon_size = (glyph_size > 0) ? glyph_size : size;
  int icon_size = util::ScaleForDpi(base_icon_size, dpi);
  int button_padding = util::ScaleForDpi(6, dpi);

  image_list_ = ImageList_Create(icon_size, icon_size, ILC_COLOR32, static_cast<int>(icons.size()), 0);
  if (image_list_) {
    int list_cx = 0;
    int list_cy = 0;
    if (ImageList_GetIconSize(image_list_, &list_cx, &list_cy) && (list_cx != icon_size || list_cy != icon_size)) {
      ImageList_Destroy(image_list_);
      image_list_ = ImageList_Create(icon_size, icon_size, ILC_COLOR32, static_cast<int>(icons.size()), 0);
    }
  }
  ImageList_SetBkColor(image_list_, CLR_NONE);
  SendMessageW(hwnd_, TB_SETBITMAPSIZE, 0, MAKELPARAM(icon_size, icon_size));
  SendMessageW(hwnd_, TB_SETBUTTONSIZE, 0, MAKELPARAM(icon_size + button_padding, icon_size + button_padding));

  for (const auto& icon : icons) {
    HICON hicon = nullptr;
    if (!icon.path.empty()) {
      hicon = util::LoadIconFromFile(icon.path, base_icon_size, dpi);
    }
    if (!hicon && icon.resource_id != 0) {
      hicon = util::LoadIconResource(icon.resource_id, base_icon_size, dpi);
    }
    if (hicon) {
      ImageList_AddIcon(image_list_, hicon);
      DestroyIcon(hicon);
    }
  }

  SendMessageW(hwnd_, TB_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(image_list_));
}

void Toolbar::AddButtons(const std::vector<TBBUTTON>& buttons) {
  if (!buttons.empty()) {
    SendMessageW(hwnd_, TB_ADDBUTTONSW, static_cast<WPARAM>(buttons.size()), reinterpret_cast<LPARAM>(buttons.data()));
    SendMessageW(hwnd_, TB_AUTOSIZE, 0, 0);
  }
}

} // namespace regkit
