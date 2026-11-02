// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "regfile/script_convert.h"

#include <windows.h>

#include <string>

namespace regkit::workspace
{
class DialogFields;
}

namespace regkit
{

enum class ConvertSource
{
    kRegistry,
    kRegFile,
    kBatch,
    kPowerShell,
};

struct ConvertSettings
{
    ConvertSource source = ConvertSource::kRegistry;
    std::wstring input_path;
    std::wstring key_path;
    bool recursive = true;
    regfile::Format format = regfile::Format::kPowerShell;
    std::wstring output_path;
    bool admin_check = true;
    bool open_in_editor = false;
};

void ShowConvertDialog(HWND owner, ConvertSettings* settings);
void ConvertDialogFields(workspace::DialogFields& fields, ConvertSettings* settings);

}
