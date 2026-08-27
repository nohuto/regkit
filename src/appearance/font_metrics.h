// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

namespace regkit::appearance {

int FontPointSize(const LOGFONTW& font, int zero_height_fallback = 0);
int FontHeight(int point_size);

} // namespace regkit::appearance
