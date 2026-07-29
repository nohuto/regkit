// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "registry/registry_provider.h"
#include "search/search.h"

#include <atomic>
#include <functional>
#include <string>
#include <unordered_map>
#include <vector>

namespace regkit::search::compare {

struct Value {
  std::wstring name;
  DWORD type = REG_NONE;
  std::vector<BYTE> data;
};

struct Key {
  std::wstring relative_path;
  std::unordered_map<std::wstring, Value> values;
};

struct Snapshot {
  std::wstring label;
  std::wstring base_path;
  std::unordered_map<std::wstring, Key> keys;
};

using NormalizePath =
    std::function<std::wstring(const std::wstring& path)>;

bool CaptureRegistry(const std::wstring& base_path,
                     const RegistryNode& base_node, bool recursive,
                     Snapshot* snapshot, std::atomic_bool* cancel = nullptr);

bool LoadRegFile(const std::wstring& file_path,
                 const std::wstring& base_path, bool recursive,
                 const NormalizePath& normalize, Snapshot* snapshot,
                 std::wstring* error,
                 std::atomic_bool* cancel = nullptr);

std::vector<Result> Diff(const Snapshot& first, const Snapshot& second,
                         std::atomic_bool* cancel = nullptr);

} // namespace regkit::search::compare
