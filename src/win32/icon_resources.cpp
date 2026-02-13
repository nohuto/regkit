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

#include "win32/icon_resources.h"

#include <commctrl.h>

namespace util {

namespace {

UINT ResolveDpi(UINT dpi) {
  if (dpi != 0) {
    return dpi;
  }
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (user32) {
    auto get_system_dpi = reinterpret_cast<UINT(WINAPI*)()>(GetProcAddress(user32, "GetDpiForSystem"));
    if (get_system_dpi) {
      return get_system_dpi();
    }
  }
  return 96;
}

HICON LoadIconWithScaleDownIfAvailable(HINSTANCE instance, int resource_id, int size) {
  if (size <= 0) {
    return nullptr;
  }
  HMODULE comctl = GetModuleHandleW(L"comctl32.dll");
  if (!comctl) {
    return nullptr;
  }
  using LoadIconWithScaleDownFn = HRESULT(WINAPI*)(HINSTANCE, PCWSTR, int, int, HICON*);
  auto load_icon = reinterpret_cast<LoadIconWithScaleDownFn>(GetProcAddress(comctl, "LoadIconWithScaleDown"));
  if (!load_icon) {
    return nullptr;
  }
  HICON icon = nullptr;
  if (SUCCEEDED(load_icon(instance, MAKEINTRESOURCEW(resource_id), size, size, &icon))) {
    return icon;
  }
  return nullptr;
}

bool GetIconSize(HICON icon, int* width, int* height) {
  if (!icon) {
    return false;
  }
  ICONINFO info = {};
  if (!GetIconInfo(icon, &info)) {
    return false;
  }
  BITMAP bmp = {};
  int w = 0;
  int h = 0;
  if (info.hbmColor && GetObject(info.hbmColor, sizeof(bmp), &bmp) == sizeof(bmp)) {
    w = bmp.bmWidth;
    h = bmp.bmHeight;
  } else if (info.hbmMask && GetObject(info.hbmMask, sizeof(bmp), &bmp) == sizeof(bmp)) {
    w = bmp.bmWidth;
    h = bmp.bmHeight / 2;
  }
  if (info.hbmColor) {
    DeleteObject(info.hbmColor);
  }
  if (info.hbmMask) {
    DeleteObject(info.hbmMask);
  }
  if (w <= 0 || h <= 0) {
    return false;
  }
  if (width) {
    *width = w;
  }
  if (height) {
    *height = h;
  }
  return true;
}

HICON EnsureIconSize(HICON icon, int size) {
  if (!icon || size <= 0) {
    return icon;
  }
  int width = 0;
  int height = 0;
  if (GetIconSize(icon, &width, &height) && width == size && height == size) {
    return icon;
  }
  HICON resized = static_cast<HICON>(CopyImage(icon, IMAGE_ICON, size, size, LR_COPYFROMRESOURCE));
  if (!resized) {
    resized = static_cast<HICON>(CopyImage(icon, IMAGE_ICON, size, size, 0));
  }
  if (resized) {
    DestroyIcon(icon);
    return resized;
  }
  return icon;
}

} // namespace

int ScaleForDpi(int size, UINT dpi) {
  if (size <= 0) {
    return size;
  }
  dpi = ResolveDpi(dpi);
  if (dpi <= 96) {
    return size;
  }
  int scaled = MulDiv(size, static_cast<int>(dpi), 96);
  if (scaled <= 0) {
    scaled = size;
  }
  return scaled;
}

HICON LoadIconResource(int resource_id, int size, UINT dpi) {
  if (resource_id == 0 || size <= 0) {
    return nullptr;
  }
  HINSTANCE instance = GetModuleHandleW(nullptr);
  int scaled = ScaleForDpi(size, dpi);

  if (HICON icon = LoadIconWithScaleDownIfAvailable(instance, resource_id, scaled)) {
    return EnsureIconSize(icon, scaled);
  }
  HICON icon = static_cast<HICON>(LoadImageW(instance, MAKEINTRESOURCEW(resource_id), IMAGE_ICON, scaled, scaled, LR_DEFAULTCOLOR));
  if (!icon && scaled != size) {
    icon = static_cast<HICON>(LoadImageW(instance, MAKEINTRESOURCEW(resource_id), IMAGE_ICON, size, size, LR_DEFAULTCOLOR));
  }
  return EnsureIconSize(icon, scaled);
}

HICON LoadIconFromFile(const std::wstring& path, int size, UINT dpi) {
  if (path.empty() || size <= 0) {
    return nullptr;
  }
  int scaled = ScaleForDpi(size, dpi);
  UINT flags = LR_LOADFROMFILE | LR_DEFAULTCOLOR;
  HICON icon = static_cast<HICON>(LoadImageW(nullptr, path.c_str(), IMAGE_ICON, scaled, scaled, flags));
  if (!icon && scaled != size) {
    icon = static_cast<HICON>(LoadImageW(nullptr, path.c_str(), IMAGE_ICON, size, size, flags));
  }
  return EnsureIconSize(icon, scaled);
}

} // namespace util
