// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "search/replace.h"

#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>

namespace regkit::workspace
{
class DialogFields;
}

namespace regkit
{

using ReplaceDialogResult = search::ReplaceOptions;

bool ShowReplaceDialog(HWND owner, ReplaceDialogResult* result);
void ReplaceDialogFields(workspace::DialogFields& fields, ReplaceDialogResult* result);

}
