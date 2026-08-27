// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_value.h"

#include <memory>
#include <string>
#include <unordered_map>

namespace regkit {

using VirtualRegistryValue = RegistryValue;

struct VirtualRegistryKey {
  std::wstring name;
  std::unordered_map<std::wstring, RegistryValue> values;
  std::unordered_map<std::wstring, std::unique_ptr<VirtualRegistryKey>>
      children;
};

struct VirtualRegistryData {
  std::wstring root_name;
  std::unique_ptr<VirtualRegistryKey> root;
};

} // namespace regkit
