// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_store.h"

#include <string>
#include <vector>

namespace regkit::changes
{

struct KeySnapshot
{
    std::wstring name;
    std::wstring class_name;
    std::wstring link_target;
    std::vector<BYTE> security;
    SECURITY_INFORMATION security_parts = 0;
    bool is_volatile = false;
    FILETIME last_write = {};
    std::vector<RegistryValue> values;
    std::vector<KeySnapshot> children;
    bool complete = true;
};

KeySnapshot CaptureKey(const RegistryNode& node, bool exact = true);
bool RestoreKey(const RegistryNode& parent, const KeySnapshot& snapshot);
bool ReplaceKey(const RegistryNode& node, const KeySnapshot& snapshot);

}
