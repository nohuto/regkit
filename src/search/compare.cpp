// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/compare.h"

#include "regfile/reg_file.h"
#include "registry/value_format.h"

#include <algorithm>
#include <cwctype>
#include <filesystem>
#include <unordered_set>
#include <utility>

namespace regkit::search::compare {

namespace {

bool Cancelled(const std::atomic_bool* cancel) {
  return cancel && cancel->load();
}

std::wstring Lower(std::wstring text) {
  std::transform(text.begin(), text.end(), text.begin(),
                 [](wchar_t ch) {
                   return static_cast<wchar_t>(towlower(ch));
                 });
  return text;
}

bool EqualsInsensitive(const std::wstring& left,
                       const std::wstring& right) {
  return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

bool IsWithin(const std::wstring& path, const std::wstring& base,
              bool recursive) {
  if (EqualsInsensitive(path, base)) {
    return true;
  }
  if (!recursive || path.size() <= base.size() ||
      _wcsnicmp(path.c_str(), base.c_str(), base.size()) != 0) {
    return false;
  }
  return path[base.size()] == L'\\';
}

std::wstring Combine(const std::wstring& base,
                     const std::wstring& relative) {
  if (relative.empty()) {
    return base;
  }
  if (base.empty()) {
    return relative;
  }
  return base + L"\\" + relative;
}

std::wstring DataText(const Value& value) {
  if (value.data.empty()) {
    return L"";
  }
  return value_format::DisplayData(
      value.type, value.data.data(),
      static_cast<DWORD>(value.data.size()));
}

std::wstring EntryText(const Value* value) {
  if (!value) {
    return L"(Missing)";
  }
  const std::wstring type = value_format::TypeName(value->type);
  const std::wstring data = DataText(*value);
  return data.empty() ? type : type + L": " + data;
}

template <typename Map>
void AppendKeys(const Map& source, std::unordered_set<std::wstring>* seen,
                std::vector<std::wstring>* target) {
  for (const auto& pair : source) {
    if (seen->insert(pair.first).second) {
      target->push_back(pair.first);
    }
  }
}

} // namespace

bool CaptureRegistry(const std::wstring& base_path,
                     const RegistryNode& base_node, bool recursive,
                     Snapshot* snapshot, std::atomic_bool* cancel) {
  if (!snapshot) {
    return false;
  }
  snapshot->label = base_path;
  snapshot->base_path = base_path;
  snapshot->keys.clear();

  std::vector<std::pair<RegistryNode, std::wstring>> stack;
  stack.emplace_back(base_node, L"");
  while (!stack.empty()) {
    if (Cancelled(cancel)) {
      return false;
    }
    RegistryNode node = std::move(stack.back().first);
    std::wstring relative = std::move(stack.back().second);
    stack.pop_back();

    Key key;
    key.relative_path = relative;
    RegistryStore::KeyEnumResult enumeration;
    bool reserved = false;
    RegistryStore::EnumKeyStreaming(
        node, true, true, false, &enumeration,
        [&](const ValueInfo& value, const BYTE* data, DWORD size) {
          if (Cancelled(cancel)) {
            return false;
          }
          if (!reserved) {
            if (enumeration.info_valid) {
              key.values.reserve(enumeration.info.value_count);
            }
            reserved = true;
          }
          Value captured;
          captured.name = value.name;
          captured.type = value.type;
          if (data && size > 0) {
            captured.data.assign(data, data + size);
          }
          key.values[Lower(captured.name)] = std::move(captured);
          return true;
        },
        {});
    if (Cancelled(cancel)) {
      return false;
    }
    snapshot->keys[Lower(relative)] = std::move(key);

    if (!recursive) {
      continue;
    }
    for (const auto& name :
         RegistryStore::EnumSubKeyNames(node, false)) {
      RegistryNode child = node;
      child.subkey =
          node.subkey.empty() ? name : node.subkey + L"\\" + name;
      const std::wstring child_relative =
          relative.empty() ? name : relative + L"\\" + name;
      stack.emplace_back(std::move(child), child_relative);
    }
  }
  return true;
}

bool LoadRegFile(const std::wstring& file_path,
                 const std::wstring& base_path, bool recursive,
                 const NormalizePath& normalize, Snapshot* snapshot,
                 std::wstring* error, std::atomic_bool* cancel) {
  if (!snapshot || !normalize) {
    return false;
  }
  regfile::Document document;
  bool cancelled = false;
  if (!regfile::Load(file_path, &document, error, cancel, &cancelled)) {
    return false;
  }
  if (document.keys.empty()) {
    if (error) {
      *error = L"No registry keys were found in the .reg file.";
    }
    return false;
  }

  snapshot->base_path = base_path;
  snapshot->label =
      std::filesystem::path(file_path).filename().wstring();
  if (!base_path.empty()) {
    snapshot->label += L": " + base_path;
  }
  snapshot->keys.clear();

  bool matched = false;
  for (const auto& original_path : document.key_order) {
    if (Cancelled(cancel)) {
      return false;
    }
    const std::wstring normalized = normalize(original_path);
    if (normalized.empty() ||
        !IsWithin(normalized, base_path, recursive)) {
      continue;
    }
    matched = true;
    std::wstring relative;
    if (normalized.size() > base_path.size()) {
      relative = normalized.substr(base_path.size() + 1);
    }
    auto source = document.keys.find(Lower(original_path));
    if (source == document.keys.end()) {
      source = document.keys.find(Lower(normalized));
    }

    Key key;
    key.relative_path = relative;
    if (source != document.keys.end()) {
      key.values.reserve(source->second.values.size());
      for (const auto& pair : source->second.values) {
        Value value;
        value.name = pair.second.name;
        value.type = pair.second.type;
        value.data = pair.second.data;
        key.values[Lower(value.name)] = std::move(value);
      }
    }
    snapshot->keys[Lower(relative)] = std::move(key);
  }

  if (!matched) {
    if (error) {
      *error = L"No matching keys were found for the selected path.";
    }
    return false;
  }
  return true;
}

void SortRows(std::vector<Row>* rows, int column, bool ascending) {
  if (!rows || rows->size() < 2) {
    return;
  }
  auto field = [column](const Row& row) -> const std::wstring& {
    switch (column) {
    case 1:
      return row.value_name;
    case 2:
      return row.first_text;
    case 3:
      return row.second_text;
    default:
      return row.key_path;
    }
  };
  std::stable_sort(rows->begin(), rows->end(),
                   [&](const Row& left, const Row& right) {
                     const int result =
                         _wcsicmp(field(left).c_str(), field(right).c_str());
                     return result != 0 && (ascending ? result < 0 : result > 0);
                   });
}

std::vector<Row> Diff(const Snapshot& first, const Snapshot& second,
                         std::atomic_bool* cancel) {
  std::vector<std::wstring> keys;
  keys.reserve(first.keys.size() + second.keys.size());
  std::unordered_set<std::wstring> seen;
  seen.reserve(keys.capacity());
  AppendKeys(first.keys, &seen, &keys);
  AppendKeys(second.keys, &seen, &keys);

  auto key_display = [&](const std::wstring& lower) -> const std::wstring& {
    const auto first_key = first.keys.find(lower);
    if (first_key != first.keys.end()) {
      return first_key->second.relative_path;
    }
    return second.keys.find(lower)->second.relative_path;
  };
  std::sort(keys.begin(), keys.end(),
            [&](const std::wstring& left, const std::wstring& right) {
              return _wcsicmp(key_display(left).c_str(),
                              key_display(right).c_str()) < 0;
            });

  std::vector<Row> results;
  for (const auto& key_name : keys) {
    if (Cancelled(cancel)) {
      break;
    }
    const auto first_it = first.keys.find(key_name);
    const auto second_it = second.keys.find(key_name);
    const Key* first_key =
        first_it == first.keys.end() ? nullptr : &first_it->second;
    const Key* second_key =
        second_it == second.keys.end() ? nullptr : &second_it->second;
    const std::wstring relative = key_display(key_name);
    const std::wstring first_path = Combine(first.base_path, relative);
    const std::wstring second_path = Combine(second.base_path, relative);

    if (!first_key || !second_key) {
      Row result;
      result.is_key = true;
      result.key_path = first_key ? first_path : second_path;
      result.first_text = first_key ? L"Present" : L"(Missing)";
      result.second_text = second_key ? L"Present" : L"(Missing)";
      results.push_back(std::move(result));
      continue;
    }

    std::vector<std::wstring> values;
    values.reserve(first_key->values.size() + second_key->values.size());
    std::unordered_set<std::wstring> seen_values;
    seen_values.reserve(values.capacity());
    AppendKeys(first_key->values, &seen_values, &values);
    AppendKeys(second_key->values, &seen_values, &values);
    std::sort(values.begin(), values.end(),
              [](const std::wstring& left, const std::wstring& right) {
                return _wcsicmp(left.c_str(), right.c_str()) < 0;
              });

    for (const auto& value_name : values) {
      if (Cancelled(cancel)) {
        return results;
      }
      const auto first_value = first_key->values.find(value_name);
      const auto second_value = second_key->values.find(value_name);
      const Value* left =
          first_value == first_key->values.end()
              ? nullptr
              : &first_value->second;
      const Value* right =
          second_value == second_key->values.end()
              ? nullptr
              : &second_value->second;
      if (left && right && left->type == right->type &&
          left->data == right->data) {
        continue;
      }

      Row result;
      result.key_path = first_path;
      result.value_name = left ? left->name : right->name;
      result.first_text = EntryText(left);
      result.second_text = EntryText(right);
      results.push_back(std::move(result));
    }
  }
  return results;
}

} // namespace regkit::search::compare
