// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "search/compare.h"

#include <windows.h>

#include <string>

namespace regkit::command_detail
{

enum class CompareSourceType
{
    kRegistry = 0,
    kRegFile = 1,
    kOfflineHive = 2,
    kNetwork = 3,
};


struct CompareDialogSelection
{
    CompareSourceType type = CompareSourceType::kRegistry;
    std::wstring file_path;
    std::wstring key_path;
    bool recursive = true;
};

struct CompareDialogDefaults
{
    CompareDialogSelection left;
    CompareDialogSelection right;
    search::compare::RowFilter filter = search::compare::RowFilter::kDifferences;
};

struct CompareDialogResult
{
    CompareDialogSelection left;
    CompareDialogSelection right;
    search::compare::RowFilter filter = search::compare::RowFilter::kDifferences;
};
bool ShowCompareDialog(HWND owner, const CompareDialogDefaults& defaults, CompareDialogResult* out);

} // namespace regkit::command_detail
