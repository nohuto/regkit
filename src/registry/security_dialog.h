// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include "registry/registry_store.h"

namespace regkit {

bool ShowRegistryPermissions(HWND owner, const RegistryNode& node);

} // namespace regkit
