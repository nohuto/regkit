// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/file_text.h"

#include "win32/translation.h"

namespace regkit::util
{

bool ReadStateFile(const std::wstring& path, std::wstring* output, std::wstring* error, uint64_t max_bytes)
{
    const DWORD code = win32ui::ReadTextFile(path.c_str(), output, max_bytes);
    if (code == ERROR_SUCCESS || code == ERROR_FILE_NOT_FOUND || code == ERROR_PATH_NOT_FOUND)
    {
        return true;
    }
    if (error)
    {
        *error = win32ui::FileReadErrorText(code);
    }
    return false;
}

std::wstring BadLineReason(size_t line)
{
    return TrLabel(L"Line", std::to_wstring(line));
}

}
