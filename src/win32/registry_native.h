// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "win32/win32_helpers.h"

namespace util {

UniqueHKey OpenNativeRegistryKey(const std::wstring& path, REGSAM access,
                                 bool open_link = false);
UniqueHKey OpenNativeRegistryRoot();

} // namespace util
