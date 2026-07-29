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
#include <regex>
#include <string>
#include <string_view>
#include <vector>

#include "registry/registry_provider.h"

namespace regkit::search {

struct TextOptions {
  std::wstring query;
  bool match_case = false;
  bool match_whole = false;
  bool use_regex = false;
};

struct Match {
  bool matched = false;
  size_t start = std::wstring::npos;
  size_t length = 0;
};

class Matcher {
public:
  explicit Matcher(const TextOptions& options);

  bool valid() const noexcept;
  Match Find(std::wstring_view text) const;

private:
  std::wstring query_;
  std::wregex regex_;
  bool use_regex_ = false;
  bool match_case_ = false;
  bool match_whole_ = false;
  bool valid_ = true;
};

struct Criteria {
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

enum class MatchField {
  kNone,
  kPath,
  kName,
  kData,
};

struct Result {
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
  bool data_loaded = true;
  MatchField match_field = MatchField::kNone;
  int match_start = -1;
  int match_length = 0;
};

using ProgressCallback =
    std::function<void(uint64_t searched, uint64_t total)>;
using ResultCallback = std::function<bool(Result&& result)>;

bool Run(const Criteria& criteria, std::atomic_bool* cancel_flag,
         const ResultCallback& callback,
         const ProgressCallback& progress, bool stop_on_first);

void SortResults(std::vector<Result>* results, int column,
                 bool ascending, bool compare);

} // namespace regkit::search
