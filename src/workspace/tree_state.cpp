// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "workspace/tree_state.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/win32_helpers.h"

#include <algorithm>
#include <unordered_set>

namespace regkit::workspace {

void TreeState::Clear() {
  selected_path.clear();
  expanded_paths.clear();
}

void TreeState::Normalize() {
  std::vector<std::wstring> normalized;
  normalized.reserve(expanded_paths.size());
  std::unordered_set<std::wstring> seen;
  seen.reserve(expanded_paths.size());
  for (const std::wstring& path : expanded_paths) {
    if (!path.empty() && seen.insert(util::ToLower(path)).second) {
      normalized.push_back(path);
    }
  }
  std::sort(normalized.begin(), normalized.end(),
            [](const std::wstring& left, const std::wstring& right) {
              if (left.size() != right.size()) {
                return left.size() < right.size();
              }
              return _wcsicmp(left.c_str(), right.c_str()) < 0;
            });
  expanded_paths.swap(normalized);
}

TreeState ParseTreeState(const std::wstring& content) {
  TreeState state;
  for (const std::wstring& line : record_fields::Lines(content)) {
    if (line.empty() || line.front() == L'#') {
      continue;
    }
    const size_t separator = line.find(L'=');
    if (separator == std::wstring::npos) {
      continue;
    }
    const std::wstring key =
        util::TrimWhitespace(line.substr(0, separator));
    const std::wstring value =
        record_fields::Unescape(line.substr(separator + 1));
    if (_wcsicmp(key.c_str(), L"selected") == 0) {
      state.selected_path = value;
    } else if (_wcsicmp(key.c_str(), L"expanded") == 0 && !value.empty()) {
      state.expanded_paths.push_back(value);
    }
  }
  return state;
}

std::wstring SerializeTreeState(const TreeState& state) {
  std::wstring content;
  if (!state.selected_path.empty()) {
    content.append(L"selected=");
    content.append(record_fields::Escape(state.selected_path));
    content.push_back(L'\n');
  }
  for (const std::wstring& path : state.expanded_paths) {
    if (!path.empty()) {
      content.append(L"expanded=");
      content.append(record_fields::Escape(path));
      content.push_back(L'\n');
    }
  }
  return content;
}

bool LoadTreeState(const std::wstring& path, TreeState* state) {
  if (!state) {
    return false;
  }
  state->Clear();
  std::wstring content;
  if (!util::ReadTextFile(path, &content)) {
    return false;
  }
  *state = ParseTreeState(content);
  return true;
}

bool SaveTreeState(const std::wstring& path, const TreeState& state) {
  return !path.empty() &&
         util::WriteTextFile(path, SerializeTreeState(state), false);
}

} // namespace regkit::workspace
