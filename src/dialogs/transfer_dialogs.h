// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/file_dialog.h"

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <vector>

namespace regkit
{

bool ExportRegFile(HWND owner, const std::wstring& key_path, REGSAM view, bool allow_hive, std::wstring* error, std::wstring* saved_path, win32::OpenAfter* open_after);
bool ExportRegFileSelection(HWND owner, const std::wstring& base_key_path, REGSAM view, const std::vector<std::wstring>& value_names, const std::vector<std::wstring>& subkey_names, std::wstring* error, std::wstring* saved_path, win32::OpenAfter* open_after);
bool LoadHive(HWND owner, HKEY* root, std::wstring* error);
bool UnloadHive(HWND owner, HKEY root, const std::wstring& subkey, std::wstring* error);

}
