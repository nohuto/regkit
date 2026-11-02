// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/registry_view.h"

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <string_view>
#include <vector>

namespace regkit::reg_exe
{

// switch meaning depends on verb, /f forces add/delete but sets query pattern, /d supplies add data
enum class Verb
{
    kAdd,
    kQuery,
    kOther,
};

struct Options
{
    std::wstring value_name;
    std::wstring data;
    std::wstring type_text;
    std::wstring separator = L"\\0";
    bool has_value = false;
    bool default_value = false;
    bool all_values = false;
    bool recurse = false;
    bool force = false;
    bool has_data = false;
    std::wstring find;
    bool has_find = false;
    bool keys_only = false;
    bool data_only = false;
    bool value_names = false;
    bool case_sensitive = false;
    bool exact = false;
    bool verbose = false;
    wchar_t compare_output = L'd';
    REGSAM view = win32::kDefaultRegistryView;
    std::vector<std::wstring> switches;
};

bool IsSwitch(std::wstring_view text, std::wstring_view name);
bool ParseOptions(const std::vector<std::wstring>& args, size_t first, Options* options, std::vector<std::wstring>* positional, Verb verb, std::wstring* error);
bool ParseType(std::wstring_view text, DWORD* type);
std::wstring TypeName(DWORD type);
bool BuildData(DWORD type, std::wstring_view text, std::wstring_view separator, std::vector<BYTE>* data, std::wstring* error);

}
