// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <windows.h>

#include <functional>
#include <memory>
#include <shared_mutex>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace regkit::defaults {

struct Value {
  DWORD type = REG_NONE;
  std::wstring data;
};

struct Key {
  std::unordered_map<std::wstring, Value> values;
};

struct Data {
  std::unordered_map<std::wstring, Key> values_by_key;
  std::shared_ptr<std::shared_mutex> mutex =
      std::make_shared<std::shared_mutex>();
};

struct Entry {
  std::wstring source_path;
  std::wstring key_path;
  bool has_value = false;
  std::wstring value_name;
  DWORD type = REG_NONE;
  std::wstring data;
};

using AliasPath =
    std::function<std::wstring(const std::wstring& path)>;

void Merge(Data* data, const std::vector<Entry>& entries,
           const AliasPath& alias,
           std::unordered_set<std::wstring>* affected_keys = nullptr);

} // namespace regkit::defaults
