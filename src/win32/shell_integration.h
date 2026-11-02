// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>

namespace regkit::win32
{

bool IsEditMenuRegistered(const std::wstring& exe_path);
LONG SetEditMenu(const std::wstring& exe_path, bool enable, LONG* cleanup_error = nullptr);
LONG RemoveEditMenuIfOwned(const std::wstring& exe_path);
bool IsRegEditReplacementRegistered(const std::wstring& exe_path);
LONG SetRegEditReplacement(const std::wstring& exe_path, bool enable, bool* conflict = nullptr, bool overwrite_existing = false, bool allow_writable_location = false);

}
