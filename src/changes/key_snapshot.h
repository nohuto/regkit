// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_store.h"

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
