// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "defaults/default_loader.h"

#include "regfile/reg_file.h"
#include "registry/value_format.h"

#include <algorithm>
#include <cwctype>
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

bool Load(const std::wstring& path, const NormalizePath& normalize,
          Data* data, std::vector<Entry>* entries,
          std::wstring* error, const std::atomic_bool* cancel) {
  if ((!data && !entries) || !normalize) {
    return false;
  }
  regfile::Document document;
  bool cancelled = false;
  if (!regfile::Load(path, &document, error, cancel, &cancelled)) {
    return false;
  }

  Data loaded;
  bool saw_key = false;
  std::vector<Entry> parsed_entries;
  if (entries) {
    parsed_entries.reserve(document.key_order.size());
  }
  for (const auto& source_path : document.key_order) {
    if (cancel && cancel->load()) {
      return false;
    }
    std::wstring key_path = normalize(source_path);
    if (key_path.empty()) {
      key_path = source_path;
    }
    if (key_path.empty()) {
      continue;
    }
    const auto source = document.keys.find(Lower(source_path));
    if (source == document.keys.end()) {
      continue;
    }

    saw_key = true;
    Key* target = nullptr;
    if (data) {
      target = &loaded.values_by_key[Lower(key_path)];
      target->values.reserve(source->second.values.size());
    }
    if (entries) {
      Entry key_entry;
      key_entry.source_path = source_path;
      key_entry.key_path = key_path;
      parsed_entries.push_back(std::move(key_entry));
    }
    for (const auto& pair : source->second.values) {
      if (cancel && cancel->load()) {
        return false;
      }
      Value value;
      value.type = pair.second.type;
      value.raw = pair.second.data;
      value.data = value_format::DisplayData(
          pair.second.type,
          pair.second.data.empty() ? nullptr : pair.second.data.data(),
          static_cast<DWORD>(pair.second.data.size()));
      if (target) {
        auto& target_value =
            target->values[Lower(pair.second.name)];
        if (entries) {
          target_value = value;
        } else {
          target_value = std::move(value);
        }
      }
      if (entries) {
        Entry entry;
        entry.source_path = source_path;
        entry.key_path = key_path;
        entry.has_value = true;
        entry.value_name = pair.second.name;
        entry.type = value.type;
        entry.raw = value.raw;
        entry.data = std::move(value.data);
        parsed_entries.push_back(std::move(entry));
      }
    }
  }

  if (!saw_key) {
    if (error) {
      *error = L"Default file contains no usable entries.";
    }
    return false;
  }
  if (data) {
    *data = std::move(loaded);
  }
  if (entries) {
    *entries = std::move(parsed_entries);
  }
  return true;
}

} // namespace regkit::defaults
