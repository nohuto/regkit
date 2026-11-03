// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/translation.h"
#include "workspace/favorites.h"

#include "records/escaped_fields.h"
#include "registry/key_algorithms.h"
#include "win32/file_text.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"

#include <win32ui/system/error_text.h>

#include <algorithm>

namespace regkit::workspace
{

namespace
{

constexpr std::wstring_view kHeader = L"#regkit-favorites-2";

bool LoadFromFile(const std::wstring& path, std::vector<std::wstring>* favorites)
{
    std::wstring content;
    const bool loaded = !path.empty() && win32ui::ReadTextFile(path.c_str(), &content, util::kMaxStateFileBytes) == ERROR_SUCCESS;
    const std::vector<std::wstring_view> lines = win32ui::LineViews(content);
    if (lines.empty() || lines.front() != kHeader)
    {
        *favorites = util::SplitLines(content);
        return loaded;
    }
    favorites->clear();
    for (size_t i = 1; i < lines.size(); ++i)
    {
        if (!lines[i].empty())
        {
            favorites->push_back(record_fields::Unescape(lines[i]));
        }
    }
    return loaded;
}

bool SaveToFile(const std::wstring& path, const std::vector<std::wstring>& favorites)
{
    std::wstring content(kHeader);
    content.push_back(L'\n');
    for (const std::wstring& favorite : favorites)
    {
        content.append(record_fields::Escape(favorite)).push_back(L'\n');
    }
    return !path.empty() && win32ui::WriteTextFileAtomic(path.c_str(), content) == ERROR_SUCCESS;
}

size_t MergeUnique(std::vector<std::wstring>* favorites, const std::vector<std::wstring>& additions)
{
    const size_t before = favorites->size();
    for (const std::wstring& entry : additions)
    {
        const bool present = std::any_of(favorites->begin(), favorites->end(), [&](const std::wstring& existing) {
            return util::EqualsInsensitive(existing, entry);
        });
        if (!entry.empty() && !present)
        {
            favorites->push_back(entry);
        }
    }
    return favorites->size() - before;
}

}

std::wstring FavoritesStore::FavoritesPath()
{
    const std::wstring folder = util::GetAppDataFolder();
    return folder.empty() ? std::wstring() : util::JoinPath(folder, L"favorites.txt");
}

bool FavoritesStore::Load(std::vector<std::wstring>* favorites)
{
    const std::wstring path = FavoritesPath();
    return LoadFromFile(path, favorites) || (!path.empty() && util::IsMissing(path.c_str()));
}

bool FavoritesStore::Save(const std::vector<std::wstring>& favorites)
{
    return SaveToFile(FavoritesPath(), favorites);
}

bool FavoritesStore::Add(const std::wstring& path)
{
    std::vector<std::wstring> favorites;
    return !path.empty() && Load(&favorites) && (MergeUnique(&favorites, {path}) == 0 || Save(favorites));
}

bool FavoritesStore::Remove(const std::wstring& path)
{
    std::vector<std::wstring> favorites;
    if (path.empty() || !Load(&favorites))
    {
        return false;
    }
    const size_t removed =
        std::erase_if(favorites, [&](const std::wstring& entry) { return util::EqualsInsensitive(entry, path); });
    return removed == 0 || Save(favorites);
}

bool FavoritesStore::ImportFromFile(const std::wstring& path, size_t* imported_count)
{
    std::vector<std::wstring> imported;
    std::vector<std::wstring> favorites;
    if (!LoadFromFile(path, &imported) || !Load(&favorites))
    {
        return false;
    }
    *imported_count = MergeUnique(&favorites, imported);
    return *imported_count == 0 || Save(favorites);
}

bool FavoritesStore::ExportToFile(const std::wstring& path)
{
    std::vector<std::wstring> favorites;
    return Load(&favorites) && SaveToFile(path, favorites);
}

bool FavoritesStore::ImportFromRegEdit(size_t* imported_count, std::wstring* error)
{
    if (imported_count)
    {
        *imported_count = 0;
    }
    std::vector<NamedFavorite> named;
    if (!LoadRegEdit(&named, error))
    {
        return false;
    }
    std::vector<std::wstring> imported;
    imported.reserve(named.size());
    for (auto& favorite : named)
    {
        imported.push_back(std::move(favorite.path));
    }
    std::vector<std::wstring> favorites;
    if (!Load(&favorites))
    {
        if (error)
        {
            *error = util::Tr(L"Failed to import RegEdit favorites.");
        }
        return false;
    }
    const size_t added = MergeUnique(&favorites, imported);
    if (added != 0 && !Save(favorites))
    {
        if (error)
        {
            *error = util::Tr(L"Failed to save favorites.");
        }
        return false;
    }
    if (imported_count)
    {
        *imported_count = added;
    }
    return true;
}

bool FavoritesStore::LoadRegEdit(std::vector<NamedFavorite>* favorites, std::wstring* error)
{
    favorites->clear();
    if (error)
    {
        error->clear();
    }
    registry_backend::KeyContents contents;
    const LONG result = registry_backend::ReadKeyContents(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\RegEdit\\Favorites", 0, true, &contents);
    if (result == ERROR_FILE_NOT_FOUND || result == ERROR_PATH_NOT_FOUND)
    {
        return true;
    }
    if (result != ERROR_SUCCESS)
    {
        if (error)
        {
            *error = win32ui::FormatWin32Error(result);
        }
        return false;
    }
    for (const RegistryValue& entry : contents.values)
    {
        if (entry.type != REG_SZ && entry.type != REG_EXPAND_SZ)
        {
            continue;
        }
        std::wstring value(reinterpret_cast<const wchar_t*>(entry.data.data()), entry.data.size() / sizeof(wchar_t));
        value.resize(wcsnlen_s(value.c_str(), value.size()));
        if (entry.type == REG_EXPAND_SZ)
        {
            std::wstring expanded = win32ui::ExpandEnvironment(value.c_str());
            if (!expanded.empty())
            {
                value = std::move(expanded);
            }
        }
        if (!value.empty())
        {
            favorites->push_back({entry.name, std::move(value)});
        }
    }
    return true;
}

}
