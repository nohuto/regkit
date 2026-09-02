// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

namespace regkit::appearance {

void PaintListHeader(HWND header, HFONT font, int reserved_right = 0);
void ReleaseListHeaderTheme(HWND header);

} // namespace regkit::appearance
