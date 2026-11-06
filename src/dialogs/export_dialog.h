// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "regfile/script_convert.h"
#include "win32/file_dialog.h"

#include <windows.h>

#include <string>

namespace regkit::editors
{

struct ExportRequest
{
    std::wstring path;
    bool include_subkeys = true;
    bool allow_hive = false;
    bool hive = false;
    regfile::Format format = regfile::Format::kReg;
    win32::OpenAfter open_after = win32::OpenAfter::kNone;
};

using ExportResult = ExportRequest;

bool ChooseExport(HWND owner, const ExportRequest& request, ExportResult* result);

}
