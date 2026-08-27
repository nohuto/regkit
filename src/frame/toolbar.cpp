// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/toolbar.h"

#include "appearance/icon_loader.h"
#include "win32/window_metrics.h"

namespace regkit {

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

  UINT dpi = win32::DpiForWindow(hwnd_);
  int base_icon_size = (glyph_size > 0) ? glyph_size : size;
  int icon_size = util::ScaleForDpi(base_icon_size, dpi);
  int button_padding = util::ScaleForDpi(6, dpi);

  image_list_ = ImageList_Create(icon_size, icon_size, ILC_COLOR32, static_cast<int>(icons.size()), 0);
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
    util::ImageListAddOrBlank(image_list_, hicon, icon_size);
    if (hicon) {
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
