// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/dialog/dialog_layout.h>
#include <win32ui/status_bar/status_bar.h>
#include <win32ui/tab_strip/tab_strip.h>
#include <win32ui/windows_config.h>

#include <windows.h>

namespace regkit::appearance
{

// lives for the whole run
const win32ui::DialogTheme* DialogThemeHooks();
win32ui::StatusBarHooks StatusBarThemeHooks();
win32ui::TabStripHooks TabStripThemeHooks();

}
