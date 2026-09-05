// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/handle_owner.h"

#include <string>

namespace util {

UniqueHKey OpenNativeRegistryKey(const std::wstring& path, REGSAM access,
                                 bool open_link = false);
UniqueHKey OpenNativeRegistryRoot();
bool DeleteNativeRegistryKey(HKEY key);

} // namespace util
