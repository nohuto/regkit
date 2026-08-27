// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "appearance/icon_loader.h"

#include <commctrl.h>

namespace util {

namespace {

using LoadIconWithScaleDownFn = HRESULT(WINAPI*)(HINSTANCE, PCWSTR, int, int, HICON*);

LoadIconWithScaleDownFn ScaleDownLoader() {
  static LoadIconWithScaleDownFn fn = [] {
    HMODULE comctl = GetModuleHandleW(L"comctl32.dll");
    return comctl ? reinterpret_cast<LoadIconWithScaleDownFn>(GetProcAddress(comctl, "LoadIconWithScaleDown")) : nullptr;
  }();
  return fn;
}

HICON LoadScaledDown(HINSTANCE instance, PCWSTR name, int size) {
  LoadIconWithScaleDownFn fn = ScaleDownLoader();
  HICON icon = nullptr;
  if (fn && SUCCEEDED(fn(instance, name, size, size, &icon))) {
    return icon;
  }
  return nullptr;
}

UINT ResolveDpi(UINT dpi) {
  if (dpi != 0) {
    return dpi;
  }
  static UINT(WINAPI * get_system_dpi)() = [] {
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    return user32 ? reinterpret_cast<UINT(WINAPI*)()>(GetProcAddress(user32, "GetDpiForSystem")) : nullptr;
  }();
  return get_system_dpi ? get_system_dpi() : 96;
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
  return scaled > 0 ? scaled : size;
}

HICON LoadIconResource(int resource_id, int size, UINT dpi) {
  if (resource_id == 0 || size <= 0) {
    return nullptr;
  }
  HINSTANCE instance = GetModuleHandleW(nullptr);
  int scaled = ScaleForDpi(size, dpi);
  HICON icon = LoadScaledDown(instance, MAKEINTRESOURCEW(resource_id), scaled);
  if (!icon) {
    icon = static_cast<HICON>(LoadImageW(instance, MAKEINTRESOURCEW(resource_id), IMAGE_ICON, scaled, scaled, LR_DEFAULTCOLOR));
  }
  return EnsureIconSize(icon, scaled);
}

HICON LoadIconFromFile(const std::wstring& path, int size, UINT dpi) {
  if (path.empty() || size <= 0) {
    return nullptr;
  }
  int scaled = ScaleForDpi(size, dpi);
  HICON icon = LoadScaledDown(nullptr, path.c_str(), scaled);
  if (!icon) {
    icon = static_cast<HICON>(LoadImageW(nullptr, path.c_str(), IMAGE_ICON, scaled, scaled, LR_LOADFROMFILE | LR_DEFAULTCOLOR));
  }
  return EnsureIconSize(icon, scaled);
}

void ImageListAddOrBlank(HIMAGELIST list, HICON icon, int size) {
  if (!list) {
    return;
  }
  if (icon) {
    ImageList_AddIcon(list, icon);
    return;
  }
  BITMAPINFO info = {};
  info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
  info.bmiHeader.biWidth = size;
  info.bmiHeader.biHeight = -size;
  info.bmiHeader.biPlanes = 1;
  info.bmiHeader.biBitCount = 32;
  info.bmiHeader.biCompression = BI_RGB;
  void* bits = nullptr;
  if (HBITMAP blank = CreateDIBSection(nullptr, &info, DIB_RGB_COLORS, &bits, nullptr, 0)) {
    ImageList_Add(list, blank, nullptr);
    DeleteObject(blank);
  }
}

} // namespace util
