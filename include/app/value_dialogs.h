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

#include <windows.h>

#include <string>
#include <vector>

namespace regkit {

struct ExportOptions {
  std::wstring path;
  bool include_subkeys = true;
  bool open_after = false;
};

bool PromptForValueText(HWND owner, const std::wstring& value_name, const wchar_t* title, const wchar_t* label, std::wstring* text, const wchar_t* value_type = nullptr);
bool PromptForBinary(HWND owner, const std::wstring& value_name, const std::vector<BYTE>& initial, std::vector<BYTE>* out, const wchar_t* value_type = L"REG_BINARY");
bool PromptForNumber(HWND owner, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& initial, std::vector<BYTE>* out);
bool PromptForMultiString(HWND owner, const std::wstring& value_name, const std::vector<BYTE>& initial, std::vector<BYTE>* out);
bool PromptForMultiLineText(HWND owner, const wchar_t* title, const wchar_t* label, std::wstring* text);
bool PromptForComment(HWND owner, const std::wstring& initial, bool apply_all, std::wstring* out_text, bool* out_apply_all);
bool PromptForCustomValue(HWND owner, const std::wstring& value_name, DWORD* type, std::vector<BYTE>* out);
bool PromptForFlaggedValue(HWND owner, const std::wstring& value_name, DWORD base_type, const std::vector<BYTE>& initial, const std::wstring& value_type, std::vector<BYTE>* out);
bool PromptForExportOptions(HWND owner, const std::wstring& default_path, ExportOptions* options);

} // namespace regkit
