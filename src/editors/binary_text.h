// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <span>
#include <string>

namespace regkit::editors::binary_text {

std::wstring Hex(std::span<const BYTE> data);
std::wstring Preview(std::span<const BYTE> data, int group_bytes,
                     bool unicode);

} // namespace regkit::editors::binary_text
