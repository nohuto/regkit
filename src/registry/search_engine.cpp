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

#include "registry/search_engine.h"

#include <algorithm>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <mutex>
#include <regex>
#include <string_view>
#include <thread>
#include <vector>

#include <windows.h>

namespace regkit {

namespace {

struct SearchNode {
  RegistryNode node;
  std::wstring path;
  std::wstring key_name;
};

std::wstring KeyLeafName(const RegistryNode& node) {
  if (node.subkey.empty()) {
    return node.root_name.empty() ? RegistryProvider::RootName(node.root) : node.root_name;
  }
  size_t pos = node.subkey.rfind(L'\\');
  if (pos == std::wstring::npos) {
    return node.subkey;
  }
  return node.subkey.substr(pos + 1);
}

SearchNode MakeSearchNode(const RegistryNode& node) {
  SearchNode entry;
  entry.node = node;
  entry.path = RegistryProvider::BuildPath(node);
  entry.key_name = KeyLeafName(node);
  return entry;
}

SearchNode MakeChildNode(const SearchNode& parent, const std::wstring& name) {
  SearchNode child;
  child.node.root = parent.node.root;
  child.node.root_name = parent.node.root_name;
  child.node.subkey = parent.node.subkey.empty() ? name : parent.node.subkey + L"\\" + name;
  child.path = parent.path.empty() ? name : parent.path + L"\\" + name;
  child.key_name = name;
  return child;
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

std::wstring FormatFileTime(const FILETIME& filetime) {
  if (filetime.dwLowDateTime == 0 && filetime.dwHighDateTime == 0) {
    return L"";
  }
  FILETIME local = {};
  SYSTEMTIME st = {};
  if (!FileTimeToLocalFileTime(&filetime, &local) || !FileTimeToSystemTime(&local, &st)) {
    return L"";
  }
  wchar_t buffer[64] = {};
  swprintf_s(buffer, L"%d/%d/%d %d:%02d", st.wMonth, st.wDay, st.wYear, st.wHour, st.wMinute);
  return buffer;
}

struct MatchLocation {
  bool matched = false;
  size_t start = std::wstring::npos;
  size_t length = 0;
};

class Matcher {
public:
  Matcher(const SearchCriteria& criteria, bool* ok) : query_(criteria.query), use_regex_(criteria.use_regex), match_case_(criteria.match_case), match_whole_(criteria.match_whole) {
    if (use_regex_) {
      try {
        auto flags = std::regex_constants::ECMAScript;
        if (!match_case_) {
          flags |= std::regex_constants::icase;
        }
        regex_ = std::wregex(query_, flags);
      } catch (const std::regex_error&) {
        if (ok) {
          *ok = false;
        }
      }
    }
  }

  MatchLocation MatchView(std::wstring_view text) const {
    MatchLocation location;
    if (text.empty()) {
      return location;
    }
    if (use_regex_) {
      std::wstring temp(text);
      std::wsmatch match;
      if (match_whole_) {
        if (std::regex_match(temp, match, regex_)) {
          location.matched = true;
          location.start = 0;
          location.length = match.length();
        }
      } else if (std::regex_search(temp, match, regex_)) {
        location.matched = true;
        location.start = match.position();
        location.length = match.length();
      }
      return location;
    }

    if (match_whole_) {
      if (match_case_) {
        if (text == query_) {
          location.matched = true;
          location.start = 0;
          location.length = text.size();
        }
      } else if (CompareStringOrdinal(text.data(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE) == CSTR_EQUAL) {
        location.matched = true;
        location.start = 0;
        location.length = text.size();
      }
      return location;
    }

    if (match_case_) {
      size_t pos = text.find(query_);
      if (pos != std::wstring::npos) {
        location.matched = true;
        location.start = pos;
        location.length = query_.size();
      }
    } else {
      int pos = FindStringOrdinal(FIND_FROMSTART, text.data(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE);
      if (pos >= 0) {
        location.matched = true;
        location.start = static_cast<size_t>(pos);
        location.length = query_.size();
      }
    }
    return location;
  }

private:
  std::wstring query_;
  bool use_regex_ = false;
  bool match_case_ = false;
  bool match_whole_ = false;
  std::wregex regex_;
};

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
  MatchLocation match;
  std::wstring data_text;
};

DataMatch MatchValueData(const Matcher& matcher, const HexQuery& hex_query, DWORD type, const BYTE* data, DWORD size) {
  DataMatch result;
  if (!data || size == 0) {
    return result;
  }

  DWORD base_type = RegistryProvider::NormalizeValueType(type);
  if (base_type == REG_SZ || base_type == REG_EXPAND_SZ || base_type == REG_LINK || base_type == REG_MULTI_SZ) {
    std::wstring_view view;
    if (!BuildStringView(data, size, &view)) {
      return result;
    }
    MatchLocation match = matcher.MatchView(view);
    if (!match.matched) {
      return result;
    }
    result.matched = true;
    result.match = match;
    result.data_text = RegistryProvider::FormatValueDataForDisplay(type, data, size);
    return result;
  }

  if (IsBinaryType(base_type)) {
    if (hex_query.hex_only && hex_query.parsed && !hex_query.bytes.empty()) {
      size_t needle = hex_query.bytes.size();
      if (needle <= size) {
        for (size_t i = 0; i + needle <= size; ++i) {
          if (memcmp(data + i, hex_query.bytes.data(), needle) == 0) {
            result.matched = true;
            result.data_text = RegistryProvider::FormatValueData(type, data, size);
            constexpr size_t kPreviewBytes = 32;
            size_t preview = std::min<size_t>(size, kPreviewBytes);
            if (i < preview) {
              result.match.matched = true;
              result.match.start = i * 3;
              result.match.length = needle * 3 - 1;
            }
            return result;
          }
        }
      }
      if (!hex_query.digits_only) {
        return result;
      }
    }
    std::wstring ascii;
    ascii.reserve(size);
    for (DWORD i = 0; i < size; ++i) {
      ascii.push_back(static_cast<wchar_t>(data[i]));
    }
    MatchLocation ascii_match = matcher.MatchView(ascii);
    if (ascii_match.matched) {
      result.matched = true;
      result.data_text = RegistryProvider::FormatValueData(type, data, size);
      return result;
    }
    if (size >= sizeof(wchar_t) && (size % sizeof(wchar_t)) == 0) {
      size_t wchar_count = size / sizeof(wchar_t);
      std::wstring wide;
      wide.resize(wchar_count);
      memcpy(wide.data(), data, size);
      MatchLocation wide_match = matcher.MatchView(wide);
      if (wide_match.matched) {
        result.matched = true;
        result.data_text = RegistryProvider::FormatValueData(type, data, size);
        return result;
      }
    }
    return result;
  }

  std::wstring text = RegistryProvider::FormatValueDataForDisplay(type, data, size);
  MatchLocation match = matcher.MatchView(text);
  if (!match.matched) {
    return result;
  }
  result.matched = true;
  result.match = match;
  result.data_text = std::move(text);
  return result;
}

bool IsTypeAllowed(const SearchCriteria& criteria, DWORD type) {
  if (criteria.allowed_types.empty()) {
    return true;
  }
  for (DWORD allowed : criteria.allowed_types) {
    DWORD allowed_base = RegistryProvider::NormalizeValueType(allowed);
    if (allowed_base != allowed) {
      if (allowed == type) {
        return true;
      }
      continue;
    }
    if (RegistryProvider::NormalizeValueType(type) == allowed_base) {
      return true;
    }
  }
  return false;
}

bool IsSizeAllowed(const SearchCriteria& criteria, DWORD size) {
  if (criteria.use_min_size && size < criteria.min_size) {
    return false;
  }
  if (criteria.use_max_size && size > criteria.max_size) {
    return false;
  }
  return true;
}

bool IsKeyInRange(const SearchCriteria& criteria, const FILETIME& last_write) {
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

bool SearchRegistryStreaming(const SearchCriteria& criteria, std::atomic_bool* cancel_flag, const std::function<bool(const SearchResult&)>& callback, const SearchProgressCallback& progress, bool stop_on_first) {
  if (criteria.query.empty() || criteria.start_nodes.empty()) {
    return false;
  }

  bool regex_ok = true;
  Matcher matcher(criteria, &regex_ok);
  if (!regex_ok) {
    return false;
  }
  HexQuery hex_query = ParseHexQuery(criteria.query);

  std::mutex mutex;
  std::condition_variable cv;
  std::vector<SearchNode> stack;
  stack.reserve(criteria.start_nodes.size());
  for (const auto& node : criteria.start_nodes) {
    stack.push_back(MakeSearchNode(node));
  }
  std::atomic<uint64_t> searched_keys(0);
  std::atomic<uint64_t> total_keys(stack.size());
  std::atomic<uint64_t> last_reported(0);
  std::atomic<uint64_t> last_reported_tick(0);
  int active = 0;
  bool done = false;
  std::atomic_bool stop(false);

  auto should_stop = [&]() -> bool {
    if (stop.load()) {
      return true;
    }
    return cancel_flag && cancel_flag->load();
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
    uint64_t total = total_keys.load();
    uint64_t last = last_reported.load();
    uint64_t now = GetTickCount64();
    uint64_t last_tick = last_reported_tick.load();
    if (!force) {
      if (searched - last < 512 && now - last_tick < 200 && total != searched) {
        return;
      }
    }
    if (last_reported.compare_exchange_strong(last, searched)) {
      last_reported_tick.store(now);
      progress(searched, total);
    }
  };

  auto emit = [&](SearchResult&& result) -> bool {
    if (should_stop()) {
      return false;
    }
    if (callback && !callback(result)) {
      request_stop();
      return false;
    }
    if (stop_on_first) {
      request_stop();
      return false;
    }
    return true;
  };

  bool has_excludes = !criteria.exclude_paths.empty();

  auto worker = [&]() {
    SetThreadPriority(GetCurrentThread(), THREAD_MODE_BACKGROUND_BEGIN);
    for (;;) {
      SearchNode entry;
      {
        std::unique_lock<std::mutex> lock(mutex);
        cv.wait(lock, [&]() { return done || should_stop() || !stack.empty(); });
        if (done || should_stop()) {
          return;
        }
        if (stack.empty()) {
          continue;
        }
        entry = std::move(stack.back());
        stack.pop_back();
        ++active;
      }

      searched_keys.fetch_add(1);
      report_progress(false);

      if (!should_stop()) {
        if (!has_excludes || !IsExcludedPath(entry.path, criteria.exclude_paths)) {
          RegistryProvider::KeyEnumResult enum_result;
          std::wstring date_text;
          bool key_range_checked = false;
          bool key_in_range = true;

          auto is_key_in_range = [&]() -> bool {
            if (!criteria.use_modified_from && !criteria.use_modified_to) {
              return true;
            }
            if (!key_range_checked) {
              key_range_checked = true;
              if (!enum_result.info_valid) {
                key_in_range = false;
              } else {
                key_in_range = IsKeyInRange(criteria, enum_result.info.last_write);
              }
            }
            return key_in_range;
          };

          auto get_date_text = [&]() -> const std::wstring& {
            if (enum_result.info_valid && date_text.empty()) {
              date_text = FormatFileTime(enum_result.info.last_write);
            }
            return date_text;
          };

          std::vector<std::wstring> pending_subkeys;
          bool want_values = criteria.search_values || criteria.search_data;
          bool want_subkeys = criteria.recursive;
          auto value_cb = [&](const ValueInfo& value, const BYTE* data, DWORD data_size) -> bool {
            if (should_stop()) {
              return false;
            }
            if (!IsTypeAllowed(criteria, value.type)) {
              return true;
            }
            if (!IsSizeAllowed(criteria, data_size)) {
              return true;
            }
            if (!is_key_in_range()) {
              return true;
            }

            std::wstring display_name = value.name.empty() ? L"(Default)" : value.name;
            MatchLocation name_match;
            if (criteria.search_values) {
              name_match = matcher.MatchView(display_name);
            }
            DataMatch data_match;
            if (criteria.search_data) {
              data_match = MatchValueData(matcher, hex_query, value.type, data, data_size);
            }

            if (name_match.matched || data_match.matched) {
              SearchResult result;
              result.key_path = entry.path;
              result.key_name = entry.key_name;
              result.value_name = value.name;
              result.display_name = display_name;
              result.type = value.type;
              result.type_text = RegistryProvider::FormatValueType(value.type);
              if (criteria.search_data) {
                if (data_match.matched) {
                  result.data = std::move(data_match.data_text);
                } else if (name_match.matched) {
                  result.data = RegistryProvider::FormatValueDataForDisplay(value.type, data, data_size);
                }
              } else if (name_match.matched) {
                constexpr DWORD kMaxDisplaySize = 1024 * 1024;
                if (data_size > kMaxDisplaySize) {
                  result.data.clear();
                } else {
                  ValueEntry entry_value;
                  if (RegistryProvider::QueryValue(entry.node, value.name, &entry_value)) {
                    result.data = RegistryProvider::FormatValueDataForDisplay(entry_value.type, entry_value.data.data(), static_cast<DWORD>(entry_value.data.size()));
                    result.size_text = std::to_wstring(entry_value.data.size());
                    result.type = entry_value.type;
                    result.type_text = RegistryProvider::FormatValueType(result.type);
                  }
                }
              }
              if (result.size_text.empty()) {
                result.size_text = std::to_wstring(data_size);
              }
              result.date_text = get_date_text();
              result.is_key = false;
              if (name_match.matched) {
                result.match_field = SearchMatchField::kName;
                result.match_start = static_cast<int>(name_match.start);
                result.match_length = static_cast<int>(name_match.length);
              } else if (data_match.matched && data_match.match.matched) {
                result.match_field = SearchMatchField::kData;
                result.match_start = static_cast<int>(data_match.match.start);
                result.match_length = static_cast<int>(data_match.match.length);
              }
              if (!emit(std::move(result))) {
                request_stop();
                return false;
              }
            }
            return true;
          };

          auto subkey_cb = [&](const std::wstring& name) -> bool {
            if (should_stop()) {
              return false;
            }
            pending_subkeys.push_back(name);
            return true;
          };

          RegistryProvider::EnumKeyStreaming(entry.node, want_values, criteria.search_data, want_subkeys, &enum_result, want_values ? value_cb : RegistryProvider::ValueStreamCallback(), want_subkeys ? subkey_cb : RegistryProvider::SubkeyStreamCallback());

          if (criteria.search_keys && is_key_in_range()) {
            MatchLocation key_match = matcher.MatchView(entry.key_name);
            if (key_match.matched) {
              SearchResult result;
              result.key_path = entry.path;
              result.key_name = entry.key_name;
              result.type_text = L"Key";
              result.is_key = true;
              result.date_text = get_date_text();
              size_t path_start = entry.path.size() >= entry.key_name.size() ? entry.path.size() - entry.key_name.size() : 0;
              result.match_field = SearchMatchField::kPath;
              result.match_start = static_cast<int>(path_start + key_match.start);
              result.match_length = static_cast<int>(key_match.length);
              if (!emit(std::move(result))) {
                request_stop();
              }
            }
          }

          if (!should_stop() && want_subkeys && !pending_subkeys.empty()) {
            std::lock_guard<std::mutex> lock(mutex);
            for (const auto& name : pending_subkeys) {
              stack.push_back(MakeChildNode(entry, name));
            }
            total_keys.fetch_add(static_cast<uint64_t>(pending_subkeys.size()));
            report_progress(false);
            cv.notify_all();
          }
        }
      }

      {
        std::lock_guard<std::mutex> lock(mutex);
        --active;
        if (stack.empty() && active == 0) {
          done = true;
          cv.notify_all();
        }
      }
    }
  };

  unsigned int worker_count = std::thread::hardware_concurrency();
  if (worker_count == 0) {
    worker_count = 1;
  }
  unsigned int max_workers = criteria.start_nodes.size() > 1 ? 8u : 4u;
  worker_count = std::min(worker_count, max_workers);
  std::vector<std::thread> workers;
  workers.reserve(worker_count);
  for (unsigned int i = 0; i < worker_count; ++i) {
    workers.emplace_back(worker);
  }
  for (auto& thread : workers) {
    thread.join();
  }

  report_progress(true);
  return true;
}

} // namespace regkit
