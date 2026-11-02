// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/list_view/sort.h>
#include <win32ui/system/strings.h>
#include <win32ui/window.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <string_view>

namespace regkit::util
{

using win32ui::CompareInsensitive;
using win32ui::CompareListText;
using win32ui::CompareSorted;
using win32ui::ContainsInsensitive;
using win32ui::DialogItemText;
using win32ui::EndsWithInsensitive;
using win32ui::EqualsInsensitive;
using win32ui::HexDigitValue;
using win32ui::IsBlank;
using win32ui::JoinLines;
using win32ui::ParseBool;
using win32ui::SortKeys;
using win32ui::SplitLines;
using win32ui::StartsWithInsensitive;
using win32ui::ToLower;
using win32ui::TrimWhitespace;
using win32ui::WindowText;

}
