// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include <initializer_list>
#include <string>

namespace regkit::editors::dialog_support {

void Initialize(HWND dialog, HFONT* owned_font,
                std::initializer_list<int> bordered_edits);
void AllowNewlines(HWND dialog, int control_id);
void ReleaseFont(HFONT* font);
bool HandleThemeMessage(HWND dialog, UINT message, WPARAM wparam,
                        LPARAM lparam, INT_PTR* result);
std::wstring ReadText(HWND dialog, int control_id);

} // namespace regkit::editors::dialog_support
