// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <string>
#include <vector>

namespace regkit::workspace {

struct PersistedTab {
  enum class Kind {
    kRegistry,
    kSearch,
  };

  Kind kind = Kind::kRegistry;
  std::wstring label;
  std::wstring selected_path;
  std::vector<std::wstring> expanded_paths;
  std::wstring search_cache_file;
};

struct TabState {
  static constexpr int kCurrentVersion = 1;
  int source_version = 1;
  int active_index = 0;
  std::vector<PersistedTab> tabs;
};

TabState ParseTabs(const std::wstring& content);
std::wstring SerializeTabs(const TabState& state);
bool LoadTabs(const std::wstring& path, TabState* state);
bool SaveTabs(const std::wstring& path, const TabState& state);

} // namespace regkit::workspace
