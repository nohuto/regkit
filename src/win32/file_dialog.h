// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>

namespace regkit::win32 {

HRESULT ChooseFileToOpen(HWND owner, const wchar_t* filter, std::wstring* path);
HRESULT ChooseFileToSave(HWND owner, const wchar_t* filter,
                         const wchar_t* default_extension,
                         const wchar_t* suggested_name, std::wstring* path);
HRESULT ChooseFolder(HWND owner, std::wstring* path);

bool DialogCancelled(HRESULT hr);
std::wstring FormatDialogError(HRESULT hr);

HRESULT ShellOpen(HWND owner, const wchar_t* target);
HRESULT RevealInExplorer(const std::wstring& path);


} // namespace regkit::win32
