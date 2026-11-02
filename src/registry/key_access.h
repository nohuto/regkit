// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <optional>
#include <string>
#include <vector>

namespace regkit::key_access
{

struct RunLevel
{
    const wchar_t* name = nullptr;
    std::optional<ACCESS_MASK> granted; // empty when authz could not evaluate it
};

std::optional<ACCESS_MASK> ForSid(PSECURITY_DESCRIPTOR descriptor, PSID sid);
// SYSTEM and TrustedInstaller access is modeled from group membership
std::vector<RunLevel> ForRunLevels(PSECURITY_DESCRIPTOR descriptor);
std::wstring Describe(ACCESS_MASK granted);
bool CanWrite(ACCESS_MASK granted) noexcept;

}
