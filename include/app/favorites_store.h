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

#include <string>
#include <vector>

namespace regkit {

class FavoritesStore {
public:
  static std::wstring FavoritesPath();
  static bool Load(std::vector<std::wstring>* favorites);
  static bool Save(const std::vector<std::wstring>& favorites);
  static bool Add(const std::wstring& path);
  static bool Remove(const std::wstring& path);
  static bool ImportFromFile(const std::wstring& path);
  static bool ExportToFile(const std::wstring& path);
  static bool ImportFromRegedit(size_t* imported_count, std::wstring* error);
};

} // namespace regkit
