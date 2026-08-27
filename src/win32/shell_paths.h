// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>

namespace util {

std::wstring GetModuleDirectory();
std::wstring GetModulePath();
std::wstring JoinPath(const std::wstring& left, const std::wstring& right);
std::wstring GetAppDataFolder();

} // namespace util
