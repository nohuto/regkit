// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "trace/trace_parser.h"

#include "registry/registry_path.h"
#include "win32/file_text.h"

#include <algorithm>
#include <cwctype>
#include <unordered_map>
#include <utility>

namespace regkit::trace {

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

std::wstring Trim(std::wstring text) {
  const auto first =
      std::find_if_not(text.begin(), text.end(), iswspace);
  const auto last =
      std::find_if_not(text.rbegin(), text.rend(), iswspace).base();
  return first < last ? std::wstring(first, last) : std::wstring();
}

bool EqualsInsensitive(const std::wstring& left,
                       const wchar_t* right) {
  return _wcsicmp(left.c_str(), right) == 0;
}

bool Decode(std::string_view buffer, std::wstring* content,
            std::wstring* error) {
  if (!content) {
    return false;
  }
  if (buffer.empty()) {
    if (error) {
      *error = L"Trace file is empty or too large to load.";
    }
    return false;
  }
  if (buffer.size() >= 3 &&
      static_cast<unsigned char>(buffer[0]) == 0xEF &&
      static_cast<unsigned char>(buffer[1]) == 0xBB &&
      static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.remove_prefix(3);
  }
  *content = util::Utf8ToWide(buffer);
  if (content->empty()) {
    if (error) {
      *error = L"Trace file has no readable entries.";
    }
    return false;
  }
  return true;
}

void Finalize(Data* data) {
  auto less = [](const std::wstring& left, const std::wstring& right) {
    return _wcsicmp(left.c_str(), right.c_str()) < 0;
  };
  std::sort(data->key_paths.begin(), data->key_paths.end(), less);
  std::sort(data->display_key_paths.begin(),
            data->display_key_paths.end(), less);

  std::unordered_map<
      std::wstring,
      std::unordered_map<std::wstring, std::wstring>>
      children;
  children.reserve(data->key_paths.size());
  for (const auto& path : data->key_paths) {
    const auto parts = registry_path::Split(path);
    if (parts.size() < 2) {
      continue;
    }
    std::wstring parent = parts.front();
    for (size_t index = 1; index < parts.size(); ++index) {
      children[Lower(parent)].try_emplace(Lower(parts[index]),
                                          parts[index]);
      parent += L"\\" + parts[index];
    }
  }
  data->children_by_key.reserve(children.size());
  for (auto& pair : children) {
    std::vector<std::wstring> names;
    names.reserve(pair.second.size());
    for (auto& child : pair.second) {
      names.push_back(std::move(child.second));
    }
    std::sort(names.begin(), names.end(), less);
    data->children_by_key.emplace(std::move(pair.first),
                                  std::move(names));
  }
}

} // namespace

bool ParseEntries(std::string_view buffer,
                  const Normalizers& normalizers,
                  const EntryCallback& callback, std::wstring* error,
                  const std::atomic_bool* cancel) {
  if (!normalizers.key || !normalizers.display || !callback) {
    return false;
  }
  std::wstring content;
  if (!Decode(buffer, &content, error)) {
    return false;
  }

  bool saw_entry = false;
  size_t start = 0;
  while (start < content.size()) {
    if (Cancelled(cancel)) {
      return false;
    }
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    start = end + 1;
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    line = Trim(std::move(line));
    if (line.empty()) {
      continue;
    }

    size_t separator = line.rfind(L" : ");
    size_t separator_size = 3;
    if (separator == std::wstring::npos) {
      separator = line.rfind(L':');
      separator_size = 1;
    }
    if (separator == std::wstring::npos) {
      continue;
    }

    const std::wstring source_key =
        Trim(line.substr(0, separator));
    if (source_key.empty()) {
      continue;
    }
    Entry entry;
    entry.display_path = normalizers.display(source_key);
    if (entry.display_path.empty()) {
      continue;
    }
    entry.key_path = normalizers.key(source_key);
    if (entry.key_path.empty()) {
      entry.key_path = entry.display_path;
    }
    entry.value_name =
        Trim(line.substr(separator + separator_size));
    entry.has_value = true;
    if (EqualsInsensitive(entry.value_name, L"(Default)")) {
      entry.value_name.clear();
    }
    saw_entry = true;
    if (!callback(std::move(entry))) {
      return false;
    }
  }

  if (!saw_entry) {
    if (error) {
      *error = L"Trace file contains no usable entries.";
    }
    return false;
  }
  return true;
}

bool Parse(const std::wstring& label, const std::wstring& source,
           std::string_view buffer, const Normalizers& normalizers,
           Data* data, std::wstring* error,
           const std::atomic_bool* cancel) {
  if (!data) {
    return false;
  }
  Data parsed;
  parsed.label = label;
  parsed.source_path = source;
  const bool ok = ParseEntries(
      buffer, normalizers,
      [&](Entry&& entry) {
        const std::wstring key_lower = Lower(entry.key_path);
        auto [key, inserted] =
            parsed.values_by_key.try_emplace(key_lower);
        if (inserted) {
          parsed.key_paths.push_back(entry.key_path);
        }
        const std::wstring display_lower = Lower(entry.display_path);
        if (parsed.display_to_key
                .try_emplace(display_lower, entry.key_path)
                .second) {
          parsed.display_key_paths.push_back(entry.display_path);
        }
        if (entry.has_value) {
          const std::wstring value_lower = Lower(entry.value_name);
          if (key->second.values_lower.insert(value_lower).second) {
            key->second.values_display.push_back(
                std::move(entry.value_name));
          }
        }
        return !Cancelled(cancel);
      },
      error, cancel);
  if (!ok) {
    return false;
  }
  Finalize(&parsed);
  *data = std::move(parsed);
  return true;
}

} // namespace regkit::trace
