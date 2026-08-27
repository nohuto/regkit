// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"
#include "search/replace.h"

#include <windows.h>

#include <string>

namespace regkit {

using ReplaceDialogResult = search::ReplaceOptions;

bool ShowReplaceDialog(HWND owner, ReplaceDialogResult* result);

} // namespace regkit
