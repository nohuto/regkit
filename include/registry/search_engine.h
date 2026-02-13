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

#include <atomic>
#include <cstdint>
#include <functional>
#include <string>
#include <vector>

#include "registry/registry_provider.h"

namespace regkit {

struct SearchCriteria {
  std::wstring query;
  bool search_keys = true;
  bool search_values = true;
  bool search_data = true;
  bool match_case = false;
  bool match_whole = false;
  bool use_regex = false;
  bool recursive = true;
  bool use_min_size = false;
  uint64_t min_size = 0;
  bool use_max_size = false;
  uint64_t max_size = 0;
  bool use_modified_from = false;
  FILETIME modified_from = {};
  bool use_modified_to = false;
  FILETIME modified_to = {};
  std::vector<DWORD> allowed_types;
  std::vector<RegistryNode> start_nodes;
  std::vector<std::wstring> exclude_paths;
};

enum class SearchMatchField {
  kNone,
  kPath,
  kName,
  kData,
};

struct SearchResult {
  std::wstring key_path;
  std::wstring key_name;
  std::wstring value_name;
  std::wstring display_name;
  std::wstring type_text;
  DWORD type = 0;
  std::wstring data;
  std::wstring size_text;
  std::wstring date_text;
  std::wstring comment;
  bool is_key = false;
  SearchMatchField match_field = SearchMatchField::kNone;
  int match_start = -1;
  int match_length = 0;
};

using SearchProgressCallback = std::function<void(uint64_t searched, uint64_t total)>;

bool SearchRegistryStreaming(const SearchCriteria& criteria, std::atomic_bool* cancel_flag, const std::function<bool(const SearchResult&)>& callback, const SearchProgressCallback& progress, bool stop_on_first);

} // namespace regkit
