// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>
#include <vector>

#include "search/search.h"

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
  search::Criteria criteria;
  std::wstring start_key;
  std::vector<std::wstring> exclude_paths;
  std::vector<std::wstring> root_paths;
  bool search_standard_hives = true;
  bool search_registry_root = true;
  bool search_trace_values = true;
  SearchScope scope = SearchScope::kEntireRegistry;
  SearchResultMode result_mode = SearchResultMode::kNewTab;
  bool open_in_new_tab = false;
};

bool ShowSearchDialog(HWND owner, SearchDialogResult* result, bool trace_available, bool registry_available);
bool ShowBrowseKeyDialog(HWND owner, std::wstring* selected_path);

} // namespace regkit
