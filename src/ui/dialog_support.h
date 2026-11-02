// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include "win32/text_transform.h"

#include <string>

namespace regkit::editors::dialog_support
{

inline std::wstring ReadText(HWND dialog, int control_id)
{
    return util::DialogItemText(dialog, control_id);
}

bool Matches(const std::wstring& text, const std::wstring& filter);

}
