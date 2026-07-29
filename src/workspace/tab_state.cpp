// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "workspace/tab_state.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"

namespace regkit::workspace {

TabState ParseTabs(const std::wstring& content) {
  TabState state;
  for (const std::wstring& line : record_fields::Lines(content)) {
    if (line.empty()) {
      continue;
    }
    if (line.rfind(L"active=", 0) == 0) {
      state.active_index = _wtoi(line.substr(7).c_str());
      continue;
    }
    const auto fields = record_fields::Split(line);
    if (fields.size() < 3 || _wcsicmp(fields[0].c_str(), L"tab") != 0) {
      continue;
    }
    PersistedTab tab;
    tab.label = record_fields::Unescape(fields[2]);
    if (_wcsicmp(fields[1].c_str(), L"registry") == 0) {
      tab.kind = PersistedTab::Kind::kRegistry;
      if (fields.size() >= 4) {
        tab.selected_path = record_fields::Unescape(fields[3]);
      }
      for (size_t index = 4; index < fields.size(); ++index) {
        std::wstring path = record_fields::Unescape(fields[index]);
        if (!path.empty()) {
          tab.expanded_paths.push_back(std::move(path));
        }
      }
    } else if (_wcsicmp(fields[1].c_str(), L"search") == 0 &&
               fields.size() >= 4) {
      tab.kind = PersistedTab::Kind::kSearch;
      tab.search_cache_file = record_fields::Unescape(fields[3]);
    } else {
      continue;
    }
    state.tabs.push_back(std::move(tab));
  }
  return state;
}

std::wstring SerializeTabs(const TabState& state) {
  std::wstring content =
      L"active=" + std::to_wstring(state.active_index) + L"\n";
  for (const PersistedTab& tab : state.tabs) {
    content.append(L"tab\t");
    content.append(tab.kind == PersistedTab::Kind::kSearch ? L"search\t"
                                                           : L"registry\t");
    content.append(record_fields::Escape(tab.label));
    content.push_back(L'\t');
    if (tab.kind == PersistedTab::Kind::kSearch) {
      content.append(record_fields::Escape(tab.search_cache_file));
    } else {
      content.append(record_fields::Escape(tab.selected_path));
      for (const std::wstring& path : tab.expanded_paths) {
        content.push_back(L'\t');
        content.append(record_fields::Escape(path));
      }
    }
    content.push_back(L'\n');
  }
  return content;
}

bool LoadTabs(const std::wstring& path, TabState* state) {
  if (!state) {
    return false;
  }
  std::wstring content;
  if (!util::ReadTextFile(path, &content)) {
    return false;
  }
  *state = ParseTabs(content);
  return true;
}

bool SaveTabs(const std::wstring& path, const TabState& state) {
  return !path.empty() &&
         util::WriteTextFile(path, SerializeTabs(state), false);
}

} // namespace regkit::workspace
