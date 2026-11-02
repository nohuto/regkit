// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/dialog/table_dialog.h>

#include <windows.h>

#include <optional>
#include <vector>

namespace regkit::resource_list
{

bool IsResourceType(DWORD type) noexcept;
std::optional<std::vector<win32ui::TextTable>> Decode(DWORD type, const BYTE* data, size_t size, bool translated = false);

}
