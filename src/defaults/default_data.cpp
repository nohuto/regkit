// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "defaults/default_data.h"

#include <algorithm>
#include <cwctype>
#include <mutex>
#include <utility>

namespace regkit::defaults {

namespace {

std::wstring Lower(std::wstring text) {
  std::transform(text.begin(), text.end(), text.begin(),
                 [](wchar_t ch) {
                   return static_cast<wchar_t>(towlower(ch));
                 });
  return text;
}

} // namespace

void Merge(Data* data, const std::vector<Entry>& entries,
           const AliasPath& alias,
           std::unordered_set<std::wstring>* affected_keys) {
  if (!data || entries.empty()) {
    return;
  }
  std::unique_lock<std::shared_mutex> lock(*data->mutex);
  for (const auto& entry : entries) {
    if (entry.key_path.empty()) {
      continue;
    }
    const std::wstring key = Lower(entry.key_path);
    if (affected_keys) {
      affected_keys->insert(key);
    }
    if (!entry.has_value) {
      continue;
    }
    Value value;
    value.type = entry.type;
    value.data = entry.data;
    const std::wstring name = Lower(entry.value_name);
    data->values_by_key[key].values[name] = value;

    const std::wstring alias_path = alias ? alias(entry.key_path) : L"";
    if (!alias_path.empty()) {
      const std::wstring alias_key = Lower(alias_path);
      data->values_by_key[alias_key].values[name] = std::move(value);
      if (affected_keys) {
        affected_keys->insert(alias_key);
      }
    }
  }
}

} // namespace regkit::defaults
