// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>

namespace regkit::win32
{

enum class OpenAfter
{
    kNone,
    kEditor,
    kRegKit,
};

DWORD ChooseFileToSave(HWND owner, const wchar_t* filter, const wchar_t* suggested_name, std::wstring* path, OpenAfter* open_after = nullptr, bool regkit = false);
std::wstring FormatDialogError(DWORD error);

}
