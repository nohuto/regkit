// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "trace/trace_data.h"

#include "registry/registry_path.h"

#include <algorithm>
#include <cwctype>
#include <mutex>
#include <utility>

namespace regkit::trace {

namespace {

std::wstring Lower(std::wstring text) {
  std::transform(text.begin(), text.end(), text.begin(),
                 [](wchar_t ch) {
                   return static_cast<wchar_t>(towlower(ch));
                 });
  return text;
}

bool IsChild(const std::wstring& path, const std::wstring& parent) {
  return path.size() > parent.size() &&
         path.compare(0, parent.size(), parent) == 0 &&
         path[parent.size()] == L'\\';
}

} // namespace

bool IncludesKey(const Selection& selection,
                 const std::wstring& key_lower) {
  if (selection.select_all || key_lower.empty()) {
    return true;
  }
  if (!selection.key_paths.empty()) {
    for (const auto& path : selection.key_paths) {
      const std::wstring selected = Lower(path);
      if (selected.empty()) {
        continue;
      }
      if (key_lower == selected || IsChild(selected, key_lower) ||
          (selection.recursive && IsChild(key_lower, selected))) {
        return true;
      }
    }
    return false;
  }
  if (!selection.values_by_key.empty()) {
    return selection.values_by_key.find(key_lower) !=
           selection.values_by_key.end();
  }
  return true;
}

bool IncludesValue(const Selection& selection,
                   const std::wstring& key_lower,
                   const std::wstring& value_lower) {
  if (selection.select_all) {
    return true;
  }
  const auto key = selection.values_by_key.find(key_lower);
  return key == selection.values_by_key.end() || key->second.empty() ||
         key->second.find(value_lower) != key->second.end();
}

void NormalizeSelection(const Data& data, Selection* selection) {
  if (!selection || selection->select_all) {
    return;
  }
  auto resolve_key = [&](const std::wstring& key) {
    const auto display = data.display_to_key.find(Lower(key));
    return display == data.display_to_key.end() ? key : display->second;
  };

  std::unordered_map<std::wstring, std::wstring> keys;
  keys.reserve(data.key_paths.size());
  for (const auto& path : data.key_paths) {
    keys.emplace(Lower(path), path);
  }

  std::vector<std::wstring> normalized_keys;
  normalized_keys.reserve(selection->key_paths.size());
  std::unordered_set<std::wstring> seen;
  for (const auto& selected : selection->key_paths) {
    const std::wstring resolved = resolve_key(selected);
    const std::wstring lower = Lower(resolved);
    const auto key = keys.find(lower);
    if (key != keys.end() && seen.insert(lower).second) {
      normalized_keys.push_back(key->second);
    }
  }

  std::unordered_map<std::wstring, std::unordered_set<std::wstring>>
      normalized_values;
  for (const auto& pair : selection->values_by_key) {
    const std::wstring key = Lower(resolve_key(pair.first));
    if (!key.empty()) {
      normalized_values[key].insert(pair.second.begin(), pair.second.end());
    }
  }

  selection->key_paths = std::move(normalized_keys);
  selection->values_by_key = std::move(normalized_values);
  if (selection->key_paths.empty() && selection->values_by_key.empty()) {
    selection->select_all = true;
  }
}

void Merge(Data* data, const std::vector<Entry>& entries,
           std::unordered_set<std::wstring>* affected_keys) {
  if (!data || entries.empty()) {
    return;
  }
  std::unique_lock<std::shared_mutex> lock(*data->mutex);
  for (const auto& entry : entries) {
    if (entry.key_path.empty()) {
      continue;
    }
    const std::wstring key_lower = Lower(entry.key_path);
    if (affected_keys) {
      affected_keys->insert(key_lower);
    }
    auto [key, inserted] = data->values_by_key.try_emplace(key_lower);
    if (inserted) {
      data->key_paths.push_back(entry.key_path);
      const auto parts = registry_path::Split(entry.key_path);
      if (parts.size() > 1) {
        std::wstring parent = parts.front();
        for (size_t index = 1; index < parts.size(); ++index) {
          data->children_by_key[Lower(parent)].push_back(parts[index]);
          parent += L"\\" + parts[index];
        }
      }
    }
    if (!entry.display_path.empty()) {
      const std::wstring display_lower = Lower(entry.display_path);
      if (data->display_to_key
              .try_emplace(display_lower, entry.key_path)
              .second) {
        data->display_key_paths.push_back(entry.display_path);
      }
    }
    if (entry.has_value) {
      const std::wstring value_lower = Lower(entry.value_name);
      if (key->second.values_lower.insert(value_lower).second) {
        key->second.values_display.push_back(entry.value_name);
      }
    }
  }
}

void Sort(Data* data) {
  if (!data) {
    return;
  }
  auto less = [](const std::wstring& left, const std::wstring& right) {
    return _wcsicmp(left.c_str(), right.c_str()) < 0;
  };
  std::unique_lock<std::shared_mutex> lock(*data->mutex);
  std::sort(data->key_paths.begin(), data->key_paths.end(), less);
  std::sort(data->display_key_paths.begin(),
            data->display_key_paths.end(), less);
}

} // namespace regkit::trace
