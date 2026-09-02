// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>
#include <commctrl.h>

#include <string>
#include <vector>

namespace regkit {
namespace ui {

LRESULT HandleThemedListViewCustomDraw(HWND list, NMLVCUSTOMDRAW* draw);
void PaintThemedListViewGrid(HWND list, HDC hdc);
bool ListViewItemSelected(HWND list, int item_index);

bool CopyTextToClipboard(HWND owner, const std::wstring& text);
void ShowError(HWND owner, const std::wstring& message);
void ShowWarning(HWND owner, const std::wstring& message);
void ShowInfo(HWND owner, const std::wstring& message);
void ShowAbout(HWND owner);
bool ConfirmRegFileMerge(HWND owner, const std::wstring& path);
void ShowRegFileMergeSucceeded(HWND owner, const std::wstring& path);
void ShowRegFileMergeFailed(HWND owner, const std::wstring& path, const std::wstring& detail);
bool ConfirmDelete(HWND owner, const std::wstring& title, const std::wstring& name);
bool ConfirmDelete(HWND owner, const std::wstring& title, const std::vector<std::wstring>& names);
int PromptKeyChoice(HWND owner, const std::wstring& message, const std::wstring& key_path, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, int yes_button_width_dlu = 0);
int PromptChoice(HWND owner, const std::wstring& message, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, int button_width_dlu = 50, int width = 420);
bool ReportFileDialogResult(HWND owner, HRESULT hr);
bool LaunchNewInstance();

} // namespace ui
} // namespace regkit
