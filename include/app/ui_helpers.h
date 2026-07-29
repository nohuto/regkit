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

#include "../win32/windows_config.h"

#include <windows.h>
#include <commctrl.h>

#include <string>

namespace regkit { namespace ui {

LOGFONTW DefaultUIFontLogFont();
HFONT DefaultUIFont();

LRESULT HandleThemedListViewCustomDraw(HWND list, NMLVCUSTOMDRAW* draw);

bool CopyTextToClipboard(HWND owner, const std::wstring& text);
void ShowError(HWND owner, const std::wstring& message);
void ShowWarning(HWND owner, const std::wstring& message);
void ShowInfo(HWND owner, const std::wstring& message);
void ShowAbout(HWND owner);
bool ConfirmRegFileMerge(HWND owner, const std::wstring& path);
void ShowRegFileMergeSucceeded(HWND owner, const std::wstring& path);
void ShowRegFileMergeFailed(HWND owner, const std::wstring& path, const std::wstring& detail);
bool ConfirmDelete(HWND owner, const std::wstring& title, const std::wstring& name);
int PromptChoice(HWND owner, const std::wstring& message, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label);
bool LaunchNewInstance();

} } // namespace regkit::ui
