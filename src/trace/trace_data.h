// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <memory>
#include <shared_mutex>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace regkit::trace {

struct Entry {
  std::wstring key_path;
  std::wstring display_path;
  bool has_value = false;
  std::wstring value_name;
};

struct Selection {
  bool select_all = false;
  bool recursive = true;
  std::vector<std::wstring> key_paths;
  std::unordered_map<std::wstring, std::unordered_set<std::wstring>>
      values_by_key;
};

struct KeyValues {
  std::unordered_set<std::wstring> values_lower;
  std::vector<std::wstring> values_display;
};

struct Data {
  std::wstring label;
  std::wstring source_path;
  std::unordered_map<std::wstring, KeyValues> values_by_key;
  std::unordered_map<std::wstring, std::vector<std::wstring>>
      children_by_key;
  std::vector<std::wstring> key_paths;
  std::vector<std::wstring> display_key_paths;
  std::unordered_map<std::wstring, std::wstring> display_to_key;
  std::shared_ptr<std::shared_mutex> mutex =
      std::make_shared<std::shared_mutex>();
};

bool IncludesKey(const Selection& selection,
                 const std::wstring& key_lower);
bool IncludesValue(const Selection& selection,
                   const std::wstring& key_lower,
                   const std::wstring& value_lower);
void NormalizeSelection(const Data& data, Selection* selection);
void Merge(Data* data, const std::vector<Entry>& entries,
           std::unordered_set<std::wstring>* affected_keys = nullptr);
void Sort(Data* data);

} // namespace regkit::trace
