// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#pragma once

#include <windows.h>

#include <string>
#include <vector>

#include "registry/search_engine.h"

namespace regkit {

enum class SearchScope {
  kEntireRegistry,
  kCurrentKey,
};

enum class SearchResultMode {
  kReuseTab,
  kNewTab,
};

struct SearchDialogResult {
  SearchCriteria criteria;
  std::wstring start_key;
  std::vector<std::wstring> exclude_paths;
  std::vector<std::wstring> root_paths;
  bool search_standard_hives = true;
  bool search_registry_root = true;
  bool search_trace_values = true;
  SearchScope scope = SearchScope::kEntireRegistry;
  SearchResultMode result_mode = SearchResultMode::kNewTab;
};

bool ShowSearchDialog(HWND owner, SearchDialogResult* result, bool trace_available, bool registry_available);
bool ShowBrowseKeyDialog(HWND owner, std::wstring* selected_path);

} // namespace regkit
