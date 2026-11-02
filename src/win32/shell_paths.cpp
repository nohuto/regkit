// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/shell_paths.h"

#include <shlobj.h>

namespace regkit::util
{

std::wstring GetAppDataFolder()
{
    const DWORD override_size = GetEnvironmentVariableW(L"REGKIT_DATA_DIR", nullptr, 0);
    if (override_size > 1)
    {
        std::wstring path(override_size, L'\0');
        const DWORD written = GetEnvironmentVariableW(L"REGKIT_DATA_DIR", path.data(), override_size);
        if (written > 0 && written < override_size)
        {
            path.resize(written);
            win32ui::EnsureDirectory(path.c_str());
            return path;
        }
    }

    const std::wstring base = win32ui::KnownFolderPath(FOLDERID_LocalAppData);
    if (base.empty())
    {
        return {};
    }
    std::wstring folder = JoinPath(base, L"Noverse\\RegKit");
    win32ui::EnsureDirectory(folder.c_str());
    return folder;
}

std::wstring GetCacheFolder()
{
    std::wstring folder = GetAppDataFolder();
    if (folder.empty())
    {
        return {};
    }
    folder = JoinPath(folder, L"cache");
    win32ui::EnsureDirectory(folder.c_str());
    return folder;
}

}
