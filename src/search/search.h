// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <atomic>
#include <cstdint>
#include <functional>
#include <regex>
#include <string>
#include <string_view>
#include <vector>

#include "registry/registry_store.h"

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

enum class Provider : uint8_t {
  kLocal,
  kRemote,
  kOffline,
  kVirtual,
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
  Provider provider = Provider::kLocal;
};

enum class MatchField : uint8_t {
  kNone,
  kPath,
  kName,
  kData,
};

enum class ResultKind : uint8_t {
  kKey,
  kValue,
  kTraceKey,
  kTraceValue,
};

enum class DataState : uint8_t {
  kNotApplicable,
  kNotLoaded,
  kLoaded,
};

struct Result {
  std::wstring key_path;
  std::wstring value_name;
  std::wstring data_text;
  DWORD type = 0;
  DWORD data_size = 0;
  FILETIME modified = {};
  uint64_t row_id = 0;
  uint32_t match_start = 0;
  uint32_t match_length = 0;
  MatchField match_field = MatchField::kNone;
  ResultKind kind = ResultKind::kValue;
  DataState data_state = DataState::kNotApplicable;
};

bool IsKeyRow(const Result& result) noexcept;
std::wstring_view DisplayName(const Result& result) noexcept;
std::wstring TypeText(const Result& result);
std::wstring SizeText(const Result& result);
std::wstring DateText(const Result& result);
std::wstring_view KeyLeaf(const Result& result) noexcept;

using ProgressCallback =
    std::function<void(uint64_t searched, uint64_t total)>;
using ResultBatch = std::vector<Result>;
using BatchCallback = std::function<bool(ResultBatch&&)>;

bool Run(const Criteria& criteria, std::atomic_bool* cancel_flag,
         const BatchCallback& publish, const ProgressCallback& progress);

void SortResults(std::vector<Result>* results, int column, bool ascending);

} // namespace regkit::search
