// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "registry/registry_provider.h"

#include <string>
#include <vector>

namespace regkit::changes {

struct KeySnapshot {
  std::wstring name;
  std::vector<ValueEntry> values;
  std::vector<KeySnapshot> children;
};

KeySnapshot CaptureKey(const RegistryNode& node);
bool RestoreKey(const RegistryNode& parent, const KeySnapshot& snapshot);

} // namespace regkit::changes
