// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

namespace regkit {

struct FontDialogResult {
  bool use_default = true;
  LOGFONTW font = {};
};

bool ShowFontDialog(HWND owner, const LOGFONTW& default_font, bool use_default, const LOGFONTW& current, FontDialogResult* out);

} // namespace regkit
