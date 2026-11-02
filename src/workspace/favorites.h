// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>
#include <vector>

namespace regkit::workspace
{

struct NamedFavorite
{
    std::wstring name;
    std::wstring path;
};

class FavoritesStore
{
  public:
    static std::wstring FavoritesPath();
    static bool Load(std::vector<std::wstring>* favorites);
    static bool Save(const std::vector<std::wstring>& favorites);
    static bool Add(const std::wstring& path);
    static bool Remove(const std::wstring& path);
    static bool ImportFromFile(const std::wstring& path, size_t* imported_count);
    static bool ExportToFile(const std::wstring& path);
    static bool LoadRegEdit(std::vector<NamedFavorite>* favorites, std::wstring* error = nullptr);
    static bool ImportFromRegEdit(size_t* imported_count, std::wstring* error);
};

}
