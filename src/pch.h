// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

// rand_s must be enabled before CRT headers below pull in stdlib
#define _CRT_RAND_S

#include "win32/windows_config.h"

#include <windows.h>

#include <commctrl.h>

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <cwchar>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>
