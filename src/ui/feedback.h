// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/file_dialog.h"

#include <win32ui/windows_config.h>

#include <windows.h>

#include <commctrl.h>

#include <string>
#include <string_view>
#include <vector>

namespace regkit
{
namespace ui
{

void ShowError(HWND owner, const std::wstring& message);
void ShowStartError(HWND owner);
void CopyText(HWND owner, std::wstring_view text);
// background saves run often, the first failure of each file per session is shown
void ReportSaveFailure(HWND owner, const std::wstring& path, DWORD error);
void ShowWarning(HWND owner, const std::wstring& message);
void ShowInfo(HWND owner, const std::wstring& message);
void RevealFile(HWND owner, const std::wstring& path);
void ShowAbout(HWND owner);
bool ConfirmRegFileMerge(HWND owner, const std::wstring& path);
void ShowRegFileMergeSucceeded(HWND owner, const std::wstring& path);
void ShowRegFileMergeFailed(HWND owner, const std::wstring& path, const std::wstring& detail);
bool ConfirmOverwrite(HWND owner, const std::wstring& path, const std::wstring& confirmed_path = std::wstring());
bool ConfirmConversionSkips(HWND owner, const std::vector<std::wstring>& skipped, const wchar_t* action = nullptr);
void ShowConversionSucceeded(HWND owner, const std::wstring& path);
bool ConfirmDelete(HWND owner, const std::wstring& title, const std::wstring& name, const std::wstring& message = std::wstring());
bool ConfirmDelete(HWND owner, const std::wstring& title, const std::vector<std::wstring>& names, const std::wstring& message = std::wstring());
struct ChoiceButtonWidths
{
    int yes = 70;
    int no = 70;
    int cancel = 70;
};

int PromptKeyChoice(HWND owner, const std::wstring& message, const std::wstring& key_path, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, ChoiceButtonWidths widths = {});
int PromptChoice(HWND owner, const std::wstring& message, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, ChoiceButtonWidths widths = {}, int width = 420);
bool ConfirmPrivilegedRestart(HWND owner, const std::wstring& title);
bool ReportFileDialogResult(HWND owner, DWORD error);
bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path);
bool PromptSaveFile(HWND owner, const wchar_t* filter, std::wstring* path, win32::OpenAfter* open_after = nullptr, bool regkit = false);

inline constexpr wchar_t kRegFileFilter[] = L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0";
inline constexpr wchar_t kHiveFileFilter[] = L"Hive Files (*.*)\0*.*\0";
inline constexpr wchar_t kImportFileFilter[] = L"Registry Files (*.reg)\0*.reg\0Registry Hive Files\0*.dat;*.hiv;*.hive;*.sav\0All Files (*.*)\0*.*\0\0";
DWORD LaunchNewInstance(HWND owner, const std::wstring& arguments = std::wstring());

}
}
