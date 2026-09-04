// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/search.h"

#include "registry/registry_path.h"
#include "registry/value_format.h"

#include <algorithm>
#include <functional>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <iterator>
#include <map>
#include <mutex>
#include <regex>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

#include <windows.h>

namespace regkit::search {

Matcher::Matcher(const TextOptions& options)
    : query_(options.query), use_regex_(options.use_regex),
      match_case_(options.match_case), match_whole_(options.match_whole),
      valid_(!query_.empty()) {
  if (!valid_ || !use_regex_) {
    return;
  }
  try {
    auto flags = std::regex_constants::ECMAScript;
    if (!match_case_) {
      flags |= std::regex_constants::icase;
    }
    regex_ = std::wregex(query_, flags);
  } catch (const std::regex_error&) {
    valid_ = false;
  }
}

bool Matcher::valid() const noexcept {
  return valid_;
}

Match Matcher::Find(std::wstring_view text) const {
  Match location;
  if (!valid_ || text.empty()) {
    return location;
  }
  if (use_regex_) {
    std::match_results<std::wstring_view::const_iterator> match;
    if (match_whole_) {
      if (std::regex_match(text.begin(), text.end(), match, regex_)) {
        location.matched = true;
        location.start = 0;
        location.length = static_cast<size_t>(match.length());
      }
    } else if (std::regex_search(text.begin(), text.end(), match, regex_)) {
      location.matched = true;
      location.start =
          static_cast<size_t>(std::distance(text.begin(), match[0].first));
      location.length = static_cast<size_t>(match.length());
    }
    return location;
  }

  if (match_whole_) {
    const bool matched =
        match_case_
            ? text == query_
            : CompareStringOrdinal(
                  text.data(), static_cast<int>(text.size()),
                  query_.c_str(), static_cast<int>(query_.size()),
                  TRUE) == CSTR_EQUAL;
    if (matched) {
      location.matched = true;
      location.start = 0;
      location.length = text.size();
    }
    return location;
  }

  if (match_case_) {
    const size_t position = text.find(query_);
    if (position != std::wstring::npos) {
      location.matched = true;
      location.start = position;
      location.length = query_.size();
    }
    return location;
  }

  const int position = FindStringOrdinal(
      FIND_FROMSTART, text.data(), static_cast<int>(text.size()),
      query_.c_str(), static_cast<int>(query_.size()), TRUE);
  if (position >= 0) {
    location.matched = true;
    location.start = static_cast<size_t>(position);
    location.length = query_.size();
  }
  return location;
}

namespace {

std::wstring FormatFileTime(const FILETIME& filetime) {
  if (filetime.dwLowDateTime == 0 && filetime.dwHighDateTime == 0) {
    return L"";
  }
  FILETIME local = {};
  SYSTEMTIME st = {};
  if (!FileTimeToLocalFileTime(&filetime, &local) ||
      !FileTimeToSystemTime(&local, &st)) {
    return L"";
  }
  wchar_t buffer[64] = {};
  swprintf_s(buffer, L"%d/%d/%d %d:%02d", st.wMonth, st.wDay, st.wYear,
             st.wHour, st.wMinute);
  return buffer;
}

} // namespace

bool IsKeyRow(const Result& result) noexcept {
  return result.kind == ResultKind::kKey || result.kind == ResultKind::kTraceKey;
}

std::wstring_view DisplayName(const Result& result) noexcept {
  if (IsKeyRow(result)) {
    return std::wstring_view();
  }
  return result.value_name.empty() ? std::wstring_view(L"(Default)")
                                   : std::wstring_view(result.value_name);
}

std::wstring TypeText(const Result& result) {
  if (IsKeyRow(result)) {
    return L"Key";
  }
  if (result.kind == ResultKind::kTraceValue) {
    return L"TRACE";
  }
  return value_format::TypeName(result.type);
}

std::wstring SizeText(const Result& result) {
  if (IsKeyRow(result) || result.kind == ResultKind::kTraceValue) {
    return std::wstring();
  }
  return std::to_wstring(result.data_size);
}

std::wstring DateText(const Result& result) {
  return FormatFileTime(result.modified);
}

std::wstring_view KeyLeaf(const Result& result) noexcept {
  const std::wstring& path = result.key_path;
  const size_t slash = path.find_last_of(L'\\');
  if (slash == std::wstring::npos) {
    return std::wstring_view(path);
  }
  return std::wstring_view(path).substr(slash + 1);
}

namespace {

int CompareText(std::wstring_view left, std::wstring_view right) {
  if (left.empty()) {
    return right.empty() ? 0 : 1;
  }
  if (right.empty()) {
    return -1;
  }
  const int result = CompareStringOrdinal(
      left.data(), static_cast<int>(left.size()), right.data(),
      static_cast<int>(right.size()), TRUE);
  if (result == CSTR_LESS_THAN) {
    return -1;
  }
  if (result == CSTR_GREATER_THAN) {
    return 1;
  }
  return 0;
}

int CompareNumeric(uint64_t left, uint64_t right) {
  if (left == right) {
    return 0;
  }
  return left < right ? -1 : 1;
}

int CompareResult(const Result& left, const Result& right, int column) {
  switch (column) {
  case 0:
    return CompareText(left.key_path, right.key_path);
  case 1:
    return CompareText(DisplayName(left), DisplayName(right));
  case 3:
    return CompareText(left.data_text, right.data_text);
  case 4:
    return CompareNumeric(left.data_size, right.data_size);
  case 5:
    return CompareNumeric(
        (static_cast<uint64_t>(left.modified.dwHighDateTime) << 32) |
            left.modified.dwLowDateTime,
        (static_cast<uint64_t>(right.modified.dwHighDateTime) << 32) |
            right.modified.dwLowDateTime);
  default:
    return CompareText(left.key_path, right.key_path);
  }
}

} // namespace

void SortResults(std::vector<Result>* results, int column, bool ascending) {
  if (!results || results->size() < 2) {
    return;
  }
  if (column == 2) {
    // One formatted type string per distinct kind and registry type.
    std::map<std::pair<ResultKind, DWORD>, std::wstring> labels;
    auto label_of = [&labels](const Result& row) -> const std::wstring& {
      const auto key = std::make_pair(row.kind, row.type);
      auto found = labels.find(key);
      if (found == labels.end()) {
        found = labels.emplace(key, TypeText(row)).first;
      }
      return found->second;
    };
    for (const auto& row : *results) {
      label_of(row);
    }
    std::stable_sort(results->begin(), results->end(),
                     [&labels, ascending](const Result& left, const Result& right) {
                       const std::wstring& left_text =
                           labels.find(std::make_pair(left.kind, left.type))->second;
                       const std::wstring& right_text =
                           labels.find(std::make_pair(right.kind, right.type))->second;
                       const int result = CompareText(left_text, right_text);
                       return result != 0 && (ascending ? result < 0 : result > 0);
                     });
    return;
  }
  std::stable_sort(results->begin(), results->end(),
                   [column, ascending](const Result& left, const Result& right) {
                     const int result = CompareResult(left, right, column);
                     return result != 0 && (ascending ? result < 0 : result > 0);
                   });
}

namespace {

// One owned subkey string per queued key. The root display name is owned once
// per search by RootContext, and full display paths are built only for matches.
struct RootContext {
  HKEY root = nullptr;
  std::wstring root_name;
  std::wstring display_root;
};

struct NodeTask {
  const RootContext* context = nullptr;
  std::wstring subkey;
};

std::wstring BuildDisplayPath(const NodeTask& task) {
  if (!task.context) {
    return task.subkey;
  }
  if (task.subkey.empty()) {
    return task.context->display_root;
  }
  if (task.context->display_root.empty()) {
    return task.subkey;
  }
  std::wstring path;
  path.reserve(task.context->display_root.size() + task.subkey.size() + 1);
  path.append(task.context->display_root);
  path.push_back(L'\\');
  path.append(task.subkey);
  return path;
}

std::wstring_view TaskLeaf(const NodeTask& task) {
  if (task.subkey.empty()) {
    return task.context ? std::wstring_view(task.context->display_root)
                        : std::wstring_view();
  }
  const size_t slash = task.subkey.find_last_of(L'\\');
  if (slash == std::wstring::npos) {
    return std::wstring_view(task.subkey);
  }
  return std::wstring_view(task.subkey).substr(slash + 1);
}

RegistryNode TaskNode(const NodeTask& task) {
  RegistryNode node;
  if (task.context) {
    node.root = task.context->root;
    node.root_name = task.context->root_name;
  }
  node.subkey = task.subkey;
  return node;
}

bool IsExcludedPath(const std::wstring& path, const std::vector<std::wstring>& excludes) {
  if (excludes.empty()) {
    return false;
  }
  for (const auto& exclude : excludes) {
    if (!exclude.empty() && FindStringOrdinal(FIND_FROMSTART, path.c_str(), static_cast<int>(path.size()), exclude.c_str(), static_cast<int>(exclude.size()), TRUE) >= 0) {
      return true;
    }
  }
  return false;
}

struct HexQuery {
  bool hex_only = false;
  bool parsed = false;
  bool digits_only = false;
  std::vector<BYTE> bytes;
};

HexQuery ParseHexQuery(const std::wstring& query) {
  HexQuery result;
  std::wstring digits;
  digits.reserve(query.size());
  bool digits_only = true;
  size_t start = 0;
  if (query.size() >= 2 && query[0] == L'0' && (query[1] == L'x' || query[1] == L'X')) {
    start = 2;
  }
  for (size_t i = start; i < query.size(); ++i) {
    wchar_t ch = query[i];
    if (iswxdigit(ch)) {
      digits.push_back(ch);
      if (!iswdigit(ch)) {
        digits_only = false;
      }
    } else if (ch == L' ' || ch == L'\t' || ch == L',' || ch == L';' || ch == L'-' || ch == L':') {
      continue;
    } else {
      return result;
    }
  }
  result.hex_only = true;
  result.digits_only = digits_only;
  if (digits.empty()) {
    return result;
  }
  if ((digits.size() % 2) != 0) {
    digits.insert(digits.begin(), L'0');
  }
  result.bytes.reserve(digits.size() / 2);
  auto hex_value = [](wchar_t ch) -> int {
    if (ch >= L'0' && ch <= L'9') {
      return ch - L'0';
    }
    if (ch >= L'a' && ch <= L'f') {
      return ch - L'a' + 10;
    }
    if (ch >= L'A' && ch <= L'F') {
      return ch - L'A' + 10;
    }
    return -1;
  };
  for (size_t i = 0; i < digits.size(); i += 2) {
    int hi = hex_value(digits[i]);
    int lo = hex_value(digits[i + 1]);
    if (hi < 0 || lo < 0) {
      result.bytes.clear();
      result.hex_only = false;
      return result;
    }
    result.bytes.push_back(static_cast<BYTE>((hi << 4) | lo));
  }
  result.parsed = !result.bytes.empty();
  return result;
}

bool BuildStringView(const BYTE* data, DWORD size, std::wstring_view* view) {
  if (!view) {
    return false;
  }
  *view = std::wstring_view();
  if (!data || size < sizeof(wchar_t)) {
    return false;
  }
  size_t count = size / sizeof(wchar_t);
  if (count == 0) {
    return false;
  }
  const wchar_t* text = reinterpret_cast<const wchar_t*>(data);
  while (count > 0 && text[count - 1] == L'\0') {
    --count;
  }
  *view = std::wstring_view(text, count);
  return true;
}

bool IsBinaryType(DWORD base_type) {
  return base_type == REG_BINARY || base_type == REG_RESOURCE_LIST || base_type == REG_FULL_RESOURCE_DESCRIPTOR || base_type == REG_RESOURCE_REQUIREMENTS_LIST || base_type == REG_NONE || base_type == REG_DWORD_BIG_ENDIAN;
}

struct DataMatch {
  bool matched = false;
  Match match;
  std::wstring data_text;
};

DataMatch MatchValueData(const Matcher& matcher,
                         const HexQuery& hex_query, DWORD type,
                         const BYTE* data, DWORD size,
                         std::wstring* scratch) {
  DataMatch result;
  if (!data || size == 0) {
    return result;
  }

  DWORD base_type = value_format::NormalizeType(type);
  if (base_type == REG_SZ || base_type == REG_EXPAND_SZ || base_type == REG_LINK || base_type == REG_MULTI_SZ) {
    std::wstring_view view;
    if (!BuildStringView(data, size, &view)) {
      return result;
    }
    Match match = matcher.Find(view);
    if (!match.matched) {
      return result;
    }
    result.matched = true;
    result.match = match;
    result.data_text = value_format::DisplayData(type, data, size);
    return result;
  }

  if (IsBinaryType(base_type)) {
    if (hex_query.hex_only && hex_query.parsed && !hex_query.bytes.empty()) {
      const size_t needle = hex_query.bytes.size();
      if (needle <= size) {
        const BYTE* begin = data;
        const BYTE* end = data + size;
        const BYTE* hit = nullptr;
        if (needle <= 2) {
          for (const BYTE* p = begin; p + needle <= end; ++p) {
            if (memcmp(p, hex_query.bytes.data(), needle) == 0) {
              hit = p;
              break;
            }
          }
        } else {
          const std::boyer_moore_horspool_searcher searcher(
              hex_query.bytes.begin(), hex_query.bytes.end());
          const BYTE* found = std::search(begin, end, searcher);
          hit = found == end ? nullptr : found;
        }
        if (hit) {
          const size_t offset = static_cast<size_t>(hit - begin);
          result.matched = true;
          result.data_text = value_format::Data(type, data, size);
          result.match.matched = true;
          result.match.start = offset * 3;
          result.match.length = needle * 3 - 1;
          return result;
        }
      }
      if (!hex_query.digits_only) {
        return result;
      }
    }
    // Byte-to-wide compatibility scan reuses the worker's widening buffer.
    if (scratch) {
      scratch->assign(size, L'\0');
      for (DWORD i = 0; i < size; ++i) {
        (*scratch)[i] = static_cast<wchar_t>(data[i]);
      }
      if (matcher.Find(*scratch).matched) {
        result.matched = true;
        result.data_text = value_format::Data(type, data, size);
        return result;
      }
    }
    // UTF-16 payloads are matched in place when the buffer is aligned.
    if (size >= sizeof(wchar_t) && (size % sizeof(wchar_t)) == 0 &&
        (reinterpret_cast<uintptr_t>(data) % alignof(wchar_t)) == 0) {
      const std::wstring_view wide(reinterpret_cast<const wchar_t*>(data),
                                   size / sizeof(wchar_t));
      if (matcher.Find(wide).matched) {
        result.matched = true;
        result.data_text = value_format::Data(type, data, size);
        return result;
      }
    }
    return result;
  }

  std::wstring text = value_format::DisplayData(type, data, size);
  Match match = matcher.Find(text);
  if (!match.matched) {
    return result;
  }
  result.matched = true;
  result.match = match;
  result.data_text = std::move(text);
  return result;
}

bool IsTypeAllowed(const Criteria& criteria, DWORD type) {
  if (criteria.allowed_types.empty()) {
    return true;
  }
  for (DWORD allowed : criteria.allowed_types) {
    DWORD allowed_base = value_format::NormalizeType(allowed);
    if (allowed_base != allowed) {
      if (allowed == type) {
        return true;
      }
      continue;
    }
    if (value_format::NormalizeType(type) == allowed_base) {
      return true;
    }
  }
  return false;
}

bool IsSizeAllowed(const Criteria& criteria, DWORD size) {
  if (criteria.use_min_size && size < criteria.min_size) {
    return false;
  }
  if (criteria.use_max_size && size > criteria.max_size) {
    return false;
  }
  return true;
}

bool IsKeyInRange(const Criteria& criteria, const FILETIME& last_write) {
  if (!criteria.use_modified_from && !criteria.use_modified_to) {
    return true;
  }
  if (last_write.dwLowDateTime == 0 && last_write.dwHighDateTime == 0) {
    return false;
  }
  if (criteria.use_modified_from) {
    if (CompareFileTime(&last_write, &criteria.modified_from) < 0) {
      return false;
    }
  }
  if (criteria.use_modified_to) {
    if (CompareFileTime(&last_write, &criteria.modified_to) > 0) {
      return false;
    }
  }
  return true;
}

} // namespace



namespace {

constexpr size_t kResultBatchSize = 128;
constexpr size_t kNodeChunkSize = 24;
constexpr uint64_t kProgressKeyInterval = 256;
constexpr uint64_t kProgressTickInterval = 200;
constexpr unsigned int kLocalWorkerLimit = 4;
constexpr unsigned int kRemoteWorkerLimit = 2;
constexpr unsigned int kOfflineWorkerLimit = 4;

// Provider and scope decide the worker cap, not the number of start roots.
unsigned int WorkerPolicy(const Criteria& criteria) {
  if (!criteria.recursive) {
    return 1u;
  }
  Provider provider = criteria.provider;
  if (provider == Provider::kLocal) {
    for (const auto& node : criteria.start_nodes) {
      if (RegistryStore::IsVirtualRoot(node.root)) {
        provider = Provider::kVirtual;
        break;
      }
      if (RegistryStore::IsOfflineRoot(node.root)) {
        provider = Provider::kOffline;
        break;
      }
    }
  }
  switch (provider) {
  case Provider::kRemote:
    return kRemoteWorkerLimit;
  case Provider::kOffline:
  case Provider::kVirtual:
    return kOfflineWorkerLimit;
  case Provider::kLocal:
  default:
    return kLocalWorkerLimit;
  }
}

} // namespace

bool Run(const Criteria& criteria, std::atomic_bool* cancel_flag,
         const BatchCallback& publish, const ProgressCallback& progress) {
  if (criteria.query.empty() || criteria.start_nodes.empty()) {
    return false;
  }

  TextOptions match_options;
  match_options.query = criteria.query;
  match_options.match_case = criteria.match_case;
  match_options.match_whole = criteria.match_whole;
  match_options.use_regex = criteria.use_regex;
  const Matcher matcher(match_options);
  if (!matcher.valid()) {
    return false;
  }
  const HexQuery hex_query = ParseHexQuery(criteria.query);
  const bool has_excludes = !criteria.exclude_paths.empty();
  const bool want_values = criteria.search_values || criteria.search_data;
  const bool want_subkeys = criteria.recursive;
  const DWORD enum_max_data =
      criteria.use_max_size && criteria.max_size < MAXDWORD
          ? static_cast<DWORD>(criteria.max_size)
          : MAXDWORD;

  std::vector<std::unique_ptr<RootContext>> contexts;
  contexts.reserve(criteria.start_nodes.size());
  std::mutex mutex;
  std::condition_variable cv;
  std::vector<NodeTask> stack;
  stack.reserve(criteria.start_nodes.size());
  for (const auto& node : criteria.start_nodes) {
    auto context = std::make_unique<RootContext>();
    context->root = node.root;
    context->root_name = node.root_name;
    RegistryNode root_only = node;
    root_only.subkey.clear();
    context->display_root = registry_path::Build(root_only);
    NodeTask task;
    task.context = context.get();
    task.subkey = node.subkey;
    stack.push_back(std::move(task));
    contexts.push_back(std::move(context));
  }

  std::atomic<uint64_t> searched_keys(0);
  std::atomic<uint64_t> total_keys(stack.size());
  std::atomic<uint64_t> last_reported(0);
  std::atomic<uint64_t> last_reported_tick(0);
  int active = 0;
  bool done = false;
  std::atomic_bool stop(false);
  std::mutex publish_mutex;

  auto should_stop = [&]() -> bool {
    return stop.load() || (cancel_flag && cancel_flag->load());
  };

  auto request_stop = [&]() {
    stop.store(true);
    cv.notify_all();
  };

  auto report_progress = [&](bool force) {
    if (!progress) {
      return;
    }
    uint64_t searched = searched_keys.load();
    const uint64_t total = total_keys.load();
    uint64_t last = last_reported.load();
    const uint64_t now = GetTickCount64();
    const uint64_t last_tick = last_reported_tick.load();
    if (!force && searched - last < kProgressKeyInterval &&
        now - last_tick < kProgressTickInterval && total != searched) {
      return;
    }
    if (last_reported.compare_exchange_strong(last, searched)) {
      last_reported_tick.store(now);
      progress(searched, total);
    }
  };

  // One move per result: worker batch to sink.
  auto publish_batch = [&](ResultBatch& batch) -> bool {
    if (batch.empty()) {
      return true;
    }
    bool accepted = true;
    {
      std::lock_guard<std::mutex> lock(publish_mutex);
      accepted = !publish || publish(std::move(batch));
    }
    batch.clear();
    batch.reserve(kResultBatchSize);
    if (!accepted) {
      request_stop();
    }
    return accepted;
  };

  auto worker = [&]() {
    EnumerationScratch scratch;
    std::wstring widen_scratch;
    ResultBatch batch;
    batch.reserve(kResultBatchSize);
    std::vector<NodeTask> local;
    std::vector<NodeTask> children;
    uint64_t local_searched = 0;
    uint64_t local_tick = GetTickCount64();

    auto flush_progress = [&](bool force) {
      if (local_searched == 0 && !force) {
        return;
      }
      const uint64_t now = GetTickCount64();
      if (!force && local_searched < kProgressKeyInterval &&
          now - local_tick < kProgressTickInterval) {
        return;
      }
      searched_keys.fetch_add(local_searched);
      local_searched = 0;
      local_tick = now;
      report_progress(force);
    };

    for (;;) {
      if (local.empty()) {
        std::unique_lock<std::mutex> lock(mutex);
        cv.wait(lock, [&]() { return done || should_stop() || !stack.empty(); });
        if (done || should_stop()) {
          break;
        }
        if (stack.empty()) {
          continue;
        }
        // Transfer a chunk of nodes under one lock instead of one per key.
        const size_t take = std::min(kNodeChunkSize, stack.size());
        local.insert(local.end(),
                     std::make_move_iterator(stack.end() - take),
                     std::make_move_iterator(stack.end()));
        stack.erase(stack.end() - take, stack.end());
        active += 1;
      }

      while (!local.empty() && !should_stop()) {
        NodeTask entry = std::move(local.back());
        local.pop_back();
        ++local_searched;
        flush_progress(false);

        std::wstring display_path;
        auto path_text = [&]() -> const std::wstring& {
          if (display_path.empty()) {
            display_path = BuildDisplayPath(entry);
          }
          return display_path;
        };

        if (has_excludes &&
            IsExcludedPath(path_text(), criteria.exclude_paths)) {
          continue;
        }

        RegistryStore::KeyEnumResult enum_result;
        bool key_range_checked = false;
        bool key_in_range = true;

        auto is_key_in_range = [&]() -> bool {
          if (!criteria.use_modified_from && !criteria.use_modified_to) {
            return true;
          }
          if (!key_range_checked) {
            key_range_checked = true;
            key_in_range = enum_result.info_valid &&
                           IsKeyInRange(criteria, enum_result.info.last_write);
          }
          return key_in_range;
        };

        children.clear();
        auto value_cb = [&](const ValueInfo& value, const BYTE* data,
                            DWORD data_size) -> bool {
          if (should_stop()) {
            return false;
          }
          if (!IsTypeAllowed(criteria, value.type) ||
              !IsSizeAllowed(criteria, data_size) || !is_key_in_range()) {
            return true;
          }
          const std::wstring display_name =
              value.name.empty() ? std::wstring(L"(Default)") : value.name;
          Match name_match;
          if (criteria.search_values) {
            name_match = matcher.Find(display_name);
          }
          // Data matching cannot change emission once the name matched.
          DataMatch data_match;
          if (!name_match.matched && criteria.search_data) {
            data_match = MatchValueData(matcher, hex_query, value.type, data,
                                        data_size, &widen_scratch);
          }
          if (!name_match.matched && !data_match.matched) {
            return true;
          }

          Result result;
          result.key_path = path_text();
          result.value_name = value.name;
          result.type = value.type;
          result.data_size = data_size;
          result.kind = ResultKind::kValue;
          if (enum_result.info_valid) {
            result.modified = enum_result.info.last_write;
          }
          if (data_match.matched) {
            result.data_text = std::move(data_match.data_text);
            result.data_state = DataState::kLoaded;
            if (data_match.match.matched) {
              result.match_field = MatchField::kData;
              result.match_start = static_cast<uint32_t>(data_match.match.start);
              result.match_length =
                  static_cast<uint32_t>(data_match.match.length);
            }
          } else {
            // A name match defers data to the background hydrator.
            result.data_state = DataState::kNotLoaded;
          }
          if (name_match.matched) {
            result.match_field = MatchField::kName;
            result.match_start = static_cast<uint32_t>(name_match.start);
            result.match_length = static_cast<uint32_t>(name_match.length);
          }
          batch.push_back(std::move(result));
          if (batch.size() >= kResultBatchSize && !publish_batch(batch)) {
            return false;
          }
          return true;
        };

        auto subkey_cb = [&](const std::wstring& name) -> bool {
          if (should_stop()) {
            return false;
          }
          NodeTask child;
          child.context = entry.context;
          child.subkey.reserve(entry.subkey.size() + name.size() + 1);
          child.subkey.append(entry.subkey);
          if (!entry.subkey.empty()) {
            child.subkey.push_back(L'\\');
          }
          child.subkey.append(name);
          children.push_back(std::move(child));
          return true;
        };

        const RegistryNode node = TaskNode(entry);
        const bool enumerated = RegistryStore::EnumKeyStreaming(
            node, want_values, criteria.search_data, want_subkeys, &enum_result,
            want_values ? RegistryStore::ValueStreamCallback(value_cb)
                        : RegistryStore::ValueStreamCallback(),
            want_subkeys ? RegistryStore::SubkeyStreamCallback(subkey_cb)
                         : RegistryStore::SubkeyStreamCallback(),
            enum_max_data, &scratch, false);

        // A key that could not be opened must not produce a key result.
        if (enumerated && enum_result.info_valid && criteria.search_keys &&
            is_key_in_range()) {
          const std::wstring_view leaf = TaskLeaf(entry);
          const Match key_match = matcher.Find(leaf);
          if (key_match.matched) {
            Result result;
            result.key_path = path_text();
            result.kind = ResultKind::kKey;
            result.data_state = DataState::kNotApplicable;
            result.modified = enum_result.info.last_write;
            const size_t path_start = result.key_path.size() >= leaf.size()
                                          ? result.key_path.size() - leaf.size()
                                          : 0;
            result.match_field = MatchField::kPath;
            result.match_start =
                static_cast<uint32_t>(path_start + key_match.start);
            result.match_length = static_cast<uint32_t>(key_match.length);
            batch.push_back(std::move(result));
            if (batch.size() >= kResultBatchSize && !publish_batch(batch)) {
              break;
            }
          }
        }

        if (!should_stop() && want_subkeys && !children.empty()) {
          std::lock_guard<std::mutex> lock(mutex);
          stack.insert(stack.end(), std::make_move_iterator(children.begin()),
                       std::make_move_iterator(children.end()));
          total_keys.fetch_add(static_cast<uint64_t>(children.size()));
          cv.notify_all();
        }
      }

      publish_batch(batch);
      flush_progress(false);

      {
        std::lock_guard<std::mutex> lock(mutex);
        active -= 1;
        if (stack.empty() && active == 0) {
          done = true;
          cv.notify_all();
        }
      }
      if (should_stop()) {
        break;
      }
    }

    publish_batch(batch);
    flush_progress(true);
  };

  unsigned int worker_count = std::thread::hardware_concurrency();
  if (worker_count == 0) {
    worker_count = 1;
  }
  worker_count = std::min(worker_count, WorkerPolicy(criteria));
  worker_count = std::min<unsigned int>(
      worker_count, std::max<size_t>(1, criteria.start_nodes.size() * 4));

  // The calling thread is worker zero, so no idle coordinator is created.
  std::vector<std::thread> workers;
  workers.reserve(worker_count > 0 ? worker_count - 1 : 0);
  for (unsigned int i = 1; i < worker_count; ++i) {
    workers.emplace_back(worker);
  }
  worker();
  for (auto& thread : workers) {
    thread.join();
  }

  report_progress(true);
  return true;
}

} // namespace regkit::search
