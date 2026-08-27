// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_store.h"
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
