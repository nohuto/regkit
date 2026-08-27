// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

namespace regkit::appearance {

void SetControlFont(HWND control, HFONT font);
void RestoreDialogOwner(HWND owner, bool* restored);
void PositionDialog(HWND dialog, HWND owner, int width, int height);

} // namespace regkit::appearance
