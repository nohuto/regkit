// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>
#include <vector>

namespace regkit {

bool ImportRegFileFromPath(const std::wstring& path, std::wstring* error);
bool ExportRegFile(HWND owner, const std::wstring& key_path, std::wstring* error);
bool ExportRegFileSelection(HWND owner, const std::wstring& base_key_path, const std::vector<std::wstring>& value_names, const std::vector<std::wstring>& subkey_names, std::wstring* error);
bool LoadHive(HWND owner, HKEY* root, std::wstring* error);
bool UnloadHive(HWND owner, HKEY root, const std::wstring& subkey, std::wstring* error);
bool IsMountedHive(HKEY root, const std::wstring& subkey);

} // namespace regkit
