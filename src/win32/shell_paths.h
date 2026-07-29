// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <string>

namespace util {

std::wstring GetModuleDirectory();
std::wstring GetModulePath();
std::wstring JoinPath(const std::wstring& left, const std::wstring& right);
std::wstring GetAppDataFolder();

} // namespace util

