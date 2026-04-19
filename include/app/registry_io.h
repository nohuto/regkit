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
#include <vector>

namespace regkit {

bool ImportRegFile(HWND owner, std::wstring* error);
bool ImportRegFileFromPath(const std::wstring& path, std::wstring* error);
bool ExportRegFile(HWND owner, const std::wstring& key_path, std::wstring* error);
bool ExportRegFileSelection(HWND owner, const std::wstring& base_key_path, const std::vector<std::wstring>& value_names, const std::vector<std::wstring>& subkey_names, std::wstring* error);
bool LoadHive(HWND owner, HKEY root, std::wstring* error);
bool UnloadHive(HWND owner, HKEY root, const std::wstring& subkey, std::wstring* error);

} // namespace regkit
