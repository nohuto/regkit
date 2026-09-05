// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include <functional>
#include <string>

namespace regkit::editors {

struct LoadHiveResult {
  std::wstring file;
  std::wstring key_name;
  HKEY root = HKEY_LOCAL_MACHINE;
};

bool ChooseHiveToLoad(HWND owner, LoadHiveResult* result);

struct SymbolicLinkResult {
  std::wstring name;
  std::wstring target;
};

using BrowseKeyCallback = std::function<bool(HWND, std::wstring*)>;

bool PromptSymbolicLink(HWND owner, const std::wstring& suggested_name,
                        const BrowseKeyCallback& browse,
                        SymbolicLinkResult* result);

} // namespace regkit::editors
