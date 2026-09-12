// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <span>
#include <string>
#include <vector>

namespace regkit::editors {

using BrowseText = bool (*)(HWND owner, std::wstring* text);

struct TextRequest {
  std::wstring title;
  std::wstring label;
  std::wstring text;
  bool multiline = false;
  BrowseText browse = nullptr;
};

struct TextResult {
  std::wstring text;
};

struct CustomValueRequest {
  std::wstring value_name;
  DWORD type = REG_SZ;
  std::vector<BYTE> data;
  bool read_only = false;
};

struct CustomValueResult {
  DWORD type = REG_SZ;
  std::vector<BYTE> data;
};

struct FlaggedValueRequest {
  std::wstring value_name;
  DWORD base_type = REG_SZ;
  std::span<const BYTE> data;
  bool read_only = false;
};

struct FlaggedValueResult {
  std::vector<BYTE> data;
};

bool EditText(HWND owner, const TextRequest& request, TextResult* result);
bool EditCustomValue(HWND owner, const CustomValueRequest& request, CustomValueResult* result);
bool EditFlaggedValue(HWND owner, const FlaggedValueRequest& request, FlaggedValueResult* result);

} // namespace regkit::editors
