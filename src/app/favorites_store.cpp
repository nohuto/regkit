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

#include "app/favorites_store.h"

#include <algorithm>
#include <cwchar>
#include <fstream>
#include <vector>

#include <windows.h>

#include <shlobj.h>

#include "win32/win32_helpers.h"

namespace regkit {

namespace {

std::wstring FormatWin32Error(DWORD code) {
  if (code == 0) {
    return L"";
  }
  wchar_t buffer[512] = {};
  DWORD len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, code, 0, buffer, static_cast<DWORD>(_countof(buffer)), nullptr);
  if (len == 0) {
    return L"Unknown error.";
  }
  return buffer;
}

bool EnsureDirectory(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  DWORD attrs = GetFileAttributesW(path.c_str());
  if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY)) {
    return true;
  }
  return CreateDirectoryW(path.c_str(), nullptr) != 0;
}

bool LoadFromFile(const std::wstring& path, std::vector<std::wstring>* favorites) {
  if (!favorites) {
    return false;
  }
  favorites->clear();
  if (path.empty()) {
    return false;
  }
  std::wifstream file(path);
  if (!file.is_open()) {
    return false;
  }
  std::wstring line;
  while (std::getline(file, line)) {
    if (!line.empty()) {
      favorites->push_back(line);
    }
  }
  return true;
}

bool SaveToFile(const std::wstring& path, const std::vector<std::wstring>& favorites) {
  if (path.empty()) {
    return false;
  }
  std::wofstream file(path, std::ios::trunc);
  if (!file.is_open()) {
    return false;
  }
  for (const auto& entry : favorites) {
    file << entry << L"\n";
  }
  return true;
}

} // namespace

std::wstring FavoritesStore::FavoritesPath() {
  PWSTR appdata = nullptr;
  if (FAILED(SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_CREATE, nullptr, &appdata))) {
    return L"";
  }
  std::wstring base(appdata);
  CoTaskMemFree(appdata);
  std::wstring dir = util::JoinPath(base, L"Noverse\\RegKit");
  EnsureDirectory(dir);
  return util::JoinPath(dir, L"favorites.txt");
}

bool FavoritesStore::Load(std::vector<std::wstring>* favorites) {
  std::wstring path = FavoritesPath();
  if (path.empty()) {
    return false;
  }
  if (LoadFromFile(path, favorites)) {
    return true;
  }
  if (favorites) {
    favorites->clear();
  }
  return true;
}

bool FavoritesStore::Save(const std::vector<std::wstring>& favorites) {
  std::wstring path = FavoritesPath();
  return SaveToFile(path, favorites);
}

bool FavoritesStore::Add(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::vector<std::wstring> favorites;
  Load(&favorites);
  auto it = std::find_if(favorites.begin(), favorites.end(), [&](const std::wstring& entry) { return _wcsicmp(entry.c_str(), path.c_str()) == 0; });
  if (it == favorites.end()) {
    favorites.push_back(path);
    return Save(favorites);
  }
  return true;
}

bool FavoritesStore::Remove(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::vector<std::wstring> favorites;
  Load(&favorites);
  auto end = std::remove_if(favorites.begin(), favorites.end(), [&](const std::wstring& entry) { return _wcsicmp(entry.c_str(), path.c_str()) == 0; });
  if (end != favorites.end()) {
    favorites.erase(end, favorites.end());
    return Save(favorites);
  }
  return true;
}

bool FavoritesStore::ImportFromFile(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::vector<std::wstring> imported;
  if (!LoadFromFile(path, &imported)) {
    return false;
  }
  if (imported.empty()) {
    return true;
  }
  std::vector<std::wstring> favorites;
  Load(&favorites);
  for (const auto& entry : imported) {
    if (entry.empty()) {
      continue;
    }
    auto it = std::find_if(favorites.begin(), favorites.end(), [&](const std::wstring& existing) { return _wcsicmp(existing.c_str(), entry.c_str()) == 0; });
    if (it == favorites.end()) {
      favorites.push_back(entry);
    }
  }
  return Save(favorites);
}

bool FavoritesStore::ExportToFile(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::vector<std::wstring> favorites;
  Load(&favorites);
  return SaveToFile(path, favorites);
}

bool FavoritesStore::ImportFromRegedit(size_t* imported_count, std::wstring* error) {
  if (imported_count) {
    *imported_count = 0;
  }
  if (error) {
    error->clear();
  }

  const wchar_t* key_path = L"Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit\\Favorites";
  HKEY key = nullptr;
  LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, key_path, 0, KEY_READ, &key);
  if (result == ERROR_FILE_NOT_FOUND) {
    return true;
  }
  if (result != ERROR_SUCCESS) {
    if (error) {
      *error = FormatWin32Error(result);
    }
    return false;
  }

  DWORD value_count = 0;
  DWORD max_value_name = 0;
  DWORD max_value_len = 0;
  result = RegQueryInfoKeyW(key, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, &value_count, &max_value_name, &max_value_len, nullptr, nullptr);
  if (result != ERROR_SUCCESS) {
    RegCloseKey(key);
    if (error) {
      *error = FormatWin32Error(result);
    }
    return false;
  }

  std::vector<std::wstring> imported;
  imported.reserve(value_count);
  std::wstring name(max_value_name + 1, L'\0');
  std::vector<BYTE> data(max_value_len + sizeof(wchar_t));

  for (DWORD i = 0; i < value_count; ++i) {
    DWORD name_len = max_value_name + 1;
    DWORD data_len = max_value_len;
    DWORD type = 0;
    LONG enum_result = RegEnumValueW(key, i, name.data(), &name_len, nullptr, &type, data.data(), &data_len);
    if (enum_result == ERROR_MORE_DATA) {
      name.resize(name_len + 1);
      data.resize(data_len + sizeof(wchar_t));
      enum_result = RegEnumValueW(key, i, name.data(), &name_len, nullptr, &type, data.data(), &data_len);
    }
    if (enum_result != ERROR_SUCCESS) {
      continue;
    }
    if (type != REG_SZ && type != REG_EXPAND_SZ) {
      continue;
    }
    if (data_len == 0) {
      continue;
    }
    size_t wchar_count = data_len / sizeof(wchar_t);
    std::wstring value(reinterpret_cast<const wchar_t*>(data.data()), wchar_count);
    while (!value.empty() && value.back() == L'\0') {
      value.pop_back();
    }
    if (value.empty()) {
      continue;
    }
    if (type == REG_EXPAND_SZ) {
      DWORD expanded_len = ExpandEnvironmentStringsW(value.c_str(), nullptr, 0);
      if (expanded_len > 0) {
        std::wstring expanded(expanded_len, L'\0');
        DWORD written = ExpandEnvironmentStringsW(value.c_str(), expanded.data(), expanded_len);
        if (written > 0 && !expanded.empty()) {
          if (expanded.back() == L'\0') {
            expanded.pop_back();
          }
          value = std::move(expanded);
        }
      }
    }
    imported.push_back(value);
  }

  RegCloseKey(key);
  if (imported.empty()) {
    return true;
  }

  std::vector<std::wstring> favorites;
  Load(&favorites);
  size_t before = favorites.size();
  for (const auto& entry : imported) {
    if (entry.empty()) {
      continue;
    }
    auto it = std::find_if(favorites.begin(), favorites.end(), [&](const std::wstring& existing) { return _wcsicmp(existing.c_str(), entry.c_str()) == 0; });
    if (it == favorites.end()) {
      favorites.push_back(entry);
    }
  }
  if (!Save(favorites)) {
    if (error) {
      *error = L"Failed to save favorites.";
    }
    return false;
  }
  if (imported_count) {
    *imported_count = favorites.size() - before;
  }
  return true;
}

} // namespace regkit
