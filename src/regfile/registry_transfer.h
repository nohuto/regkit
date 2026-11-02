// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/registry_view.h"

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>

namespace regkit
{

std::wstring ExportFileName(const std::wstring& name, const wchar_t* extension);
std::wstring DefaultExportPath(const std::wstring& key_path, const wchar_t* extension);
bool ImportRegFileFromPath(const std::wstring& path, std::wstring* error, REGSAM view = win32::kDefaultRegistryView, const std::wstring& current_user_sid = {});
bool IsMountedHive(HKEY root, const std::wstring& subkey);
LONG SaveKeyToHive(HKEY root, const std::wstring& subkey, REGSAM view, const std::wstring& path);
LONG RestoreKeyFromHive(HKEY root, const std::wstring& subkey, REGSAM view, const std::wstring& path);
std::wstring HiveTransferError(LONG status, const std::wstring& path);

}
