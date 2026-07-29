// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "changes/key_snapshot.h"

#include "registry/registry_path.h"

namespace regkit::changes {
namespace {

RegistryNode ChildNode(const RegistryNode& parent, const std::wstring& name) {
  RegistryNode child = parent;
  child.subkey = parent.subkey.empty() ? name
                                      : parent.subkey + L"\\" + name;
  child.children_loaded = false;
  return child;
}

} // namespace

KeySnapshot CaptureKey(const RegistryNode& node) {
  KeySnapshot snapshot;
  snapshot.name = registry_path::Leaf(node.subkey);
  RegistryProvider::KeyEnumResult result;
  bool reserved = false;
  RegistryProvider::EnumKeyStreaming(
      node, true, true, false, &result,
      [&](const ValueInfo& info, const BYTE* data, DWORD size) {
        if (!reserved) {
          if (result.info_valid) {
            snapshot.values.reserve(result.info.value_count);
          }
          reserved = true;
        }
        ValueEntry value;
        value.name = info.name;
        value.type = info.type;
        if (data && size > 0) {
          value.data.assign(data, data + size);
        }
        snapshot.values.push_back(std::move(value));
        return true;
      },
      {});

  const auto children = RegistryProvider::EnumSubKeyNames(node, false);
  snapshot.children.reserve(children.size());
  for (const std::wstring& name : children) {
    snapshot.children.push_back(CaptureKey(ChildNode(node, name)));
  }
  return snapshot;
}

bool RestoreKey(const RegistryNode& parent, const KeySnapshot& snapshot) {
  if (snapshot.name.empty() ||
      !RegistryProvider::CreateKey(parent, snapshot.name)) {
    return false;
  }
  const RegistryNode node = ChildNode(parent, snapshot.name);
  for (const ValueEntry& value : snapshot.values) {
    if (!RegistryProvider::SetValue(node, value.name, value.type,
                                    value.data)) {
      return false;
    }
  }
  for (const KeySnapshot& child : snapshot.children) {
    if (!RestoreKey(node, child)) {
      return false;
    }
  }
  return true;
}

} // namespace regkit::changes
