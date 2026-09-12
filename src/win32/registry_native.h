// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/handle_owner.h"

#include <string>

namespace util {

UniqueHKey OpenNativeRegistryKey(const std::wstring& path, REGSAM access, bool open_link = false);
UniqueHKey OpenNativeRegistryRoot();
bool DeleteNativeRegistryKey(HKEY key);
bool ReadRegistryString(HKEY root, const wchar_t* subkey, const wchar_t* value_name, std::wstring* value);
LONG WriteRegistryString(HKEY key, const wchar_t* value_name, const std::wstring& value);

} // namespace util
