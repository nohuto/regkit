// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#pragma once

#include <windows.h>

#include <string>

namespace regkit {

struct ReplaceDialogResult {
  std::wstring find_text;
  std::wstring replace_text;
  std::wstring start_key;
  bool recursive = true;
  bool match_case = false;
  bool match_whole = false;
  bool use_regex = false;
};

bool ShowReplaceDialog(HWND owner, ReplaceDialogResult* result);

} // namespace regkit
