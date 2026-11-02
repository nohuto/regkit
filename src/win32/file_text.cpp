// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/file_text.h"

#include "win32/translation.h"

#include <win32ui/system/error_text.h>

namespace regkit::util
{

bool ReadTextFile(const std::wstring& path, std::wstring* output, uint64_t max_bytes)
{
    return win32ui::ReadTextFile(path.c_str(), output, max_bytes) == ERROR_SUCCESS && !output->empty();
}

bool ReadStateFile(const std::wstring& path, std::wstring* output, std::wstring* error, uint64_t max_bytes)
{
    std::wstring reason;
    switch (const DWORD code = win32ui::ReadTextFile(path.c_str(), output, max_bytes))
    {
    case ERROR_SUCCESS:
    case ERROR_FILE_NOT_FOUND:
    case ERROR_PATH_NOT_FOUND:
        return true;
    case ERROR_FILE_TOO_LARGE:
        reason = Tr(L"The file is too large.");
        break;
    case ERROR_INVALID_DATA:
        reason = Tr(L"The file isn't valid text.");
        break;
    default:
        reason = win32ui::FormatWin32Error(code);
        break;
    }
    if (error)
    {
        *error = std::move(reason);
    }
    return false;
}

std::wstring BadLineReason(size_t line)
{
    return TrLabel(L"Line", std::to_wstring(line));
}

}
