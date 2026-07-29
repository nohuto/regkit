// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

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
