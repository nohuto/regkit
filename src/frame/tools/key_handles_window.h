// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <functional>
#include <string>

namespace regkit
{

using KeyHandlesNavigate = std::function<void(const std::wstring& path, bool new_tab)>;

HWND ShowKeyHandlesWindow(HWND owner, KeyHandlesNavigate navigate);

}
