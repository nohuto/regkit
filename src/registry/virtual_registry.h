// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_value.h"

#include <memory>
#include <shared_mutex>
#include <string>
#include <unordered_map>
#include <vector>

namespace regkit
{

struct VirtualRegistryKey
{
    std::wstring name;
    std::unordered_map<std::wstring, RegistryValue> values;
    std::unordered_map<std::wstring, std::unique_ptr<VirtualRegistryKey>> children;
    std::vector<std::wstring> removed_values;
    bool removed = false;
    bool recreated = false;
};

struct VirtualRegistryData
{
    std::wstring root_name;
    std::unique_ptr<VirtualRegistryKey> root;
    std::shared_ptr<std::shared_mutex> mutex = std::make_shared<std::shared_mutex>();
};

}
