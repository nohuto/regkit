// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>
#include <commctrl.h>

#include <string>

namespace util {

int ScaleForDpi(int size, UINT dpi);
HICON LoadIconResource(int resource_id, int size, UINT dpi);
HICON LoadIconFromFile(const std::wstring& path, int size, UINT dpi);
void ImageListAddOrBlank(HIMAGELIST list, HICON icon, int size);

} // namespace util
