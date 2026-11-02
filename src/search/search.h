// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <atomic>
#include <cstdint>
#include <functional>
#include <string>
#include <string_view>
#include <vector>

#include "registry/registry_store.h"
#include "search/regex.h"
#include "win32/text_transform.h"

namespace regkit::search
{

struct TextOptions
{
    std::wstring query;
    bool match_case = false;
    bool match_whole = false;
    bool use_regex = false;
    bool match_all = false;
};

struct Match
{
    bool matched = false;
    size_t start = std::wstring::npos;
    size_t length = 0;
    regex::Status status = regex::Status::kNoMatch;
};

class Matcher
{
  public:
    explicit Matcher(const TextOptions& options);
    Matcher(const Matcher& other);
    Matcher(Matcher&&) noexcept = default;

    bool valid() const noexcept;
    const regex::Error& error() const noexcept;
    Match Find(std::wstring_view text) const;
    regex::Status Replace(std::wstring_view text, const std::wstring& replacement, std::wstring* out) const;

  private:
    std::wstring query_;
    regex::PatternRef pattern_;
    regex::Session session_;
    regex::Error error_;
    bool use_regex_ = false;
    bool match_case_ = false;
    bool match_whole_ = false;
    bool match_all_ = false;
    bool valid_ = true;
};

enum Anomaly : uint32_t
{
    kAnomalyNulName = 1u << 0,
    kAnomalyOddName = 1u << 1,
    kAnomalyIntegerSize = 1u << 2,
    kAnomalyStringEnd = 1u << 3,
    kAnomalyMultiString = 1u << 4,
    kAnomalyUnknownType = 1u << 5,
    kAnomalyBrokenLink = 1u << 6,
    kAnomalyVirtualStore = 1u << 7,
};
inline constexpr uint32_t kValueAnomalies = kAnomalyNulName | kAnomalyOddName | kAnomalyIntegerSize | kAnomalyStringEnd | kAnomalyMultiString | kAnomalyUnknownType;
inline constexpr uint32_t kKeyAnomalies = kAnomalyNulName | kAnomalyOddName | kAnomalyBrokenLink | kAnomalyVirtualStore;
inline constexpr uint32_t kDataAnomalies = kAnomalyStringEnd | kAnomalyMultiString;

enum class Provider : uint8_t
{
    kLocal,
    kRemote,
    kOffline,
    kVirtual,
};

struct Source
{
    enum class Kind : uint8_t
    {
        kLocal = 0,
        kRemote = 1,
        kOffline = 2,
        kRegFile = 3,
    };
    Kind kind = Kind::kLocal;
    std::wstring name;
    REGSAM view = 0;
};

bool SameSource(const Source& first, const Source& second) noexcept;
std::wstring SourceLabel(const Source& source);

struct StartNode
{
    RegistryNode node;
    uint16_t source = 0;
};

struct Criteria
{
    std::wstring query;
    bool search_keys = true;
    bool search_values = true;
    bool search_data = true;
    bool search_comments = false;
    bool match_case = false;
    bool match_whole = false;
    bool use_regex = false;
    bool recursive = true;
    bool skip_links = false;
    bool use_min_size = false;
    uint64_t min_size = 0;
    bool use_max_size = false;
    uint64_t max_size = 0;
    bool use_modified_from = false;
    FILETIME modified_from = {};
    bool use_modified_to = false;
    FILETIME modified_to = {};
    std::vector<DWORD> allowed_types;
    uint32_t anomalies = 0;
    std::vector<StartNode> start_nodes;
    std::vector<std::wstring> exclude_paths;
    uint64_t max_results = 1000;
    Provider provider = Provider::kLocal;
    std::shared_ptr<const Matcher> matcher;
    std::function<std::wstring(const std::wstring& key_path, const std::wstring* value_name, DWORD type, DWORD data_size)> comment_text;
    std::function<std::wstring(const std::wstring& key_path, const std::wstring& value_name)> default_text;
};

enum class MatchField : uint8_t
{
    kNone,
    kPath,
    kName,
    kData,
    kComment,
    kDefault,
};

const wchar_t* MatchFieldLabel(MatchField field) noexcept;

enum class ResultKind : uint8_t
{
    kKey,
    kValue,
    kTraceKey,
    kTraceValue,
};

enum class DataState : uint8_t
{
    kNotApplicable,
    kNotLoaded,
    kLoaded,
};

struct Result
{
    std::wstring key_path;
    std::wstring value_name;
    std::wstring data_text;
    DWORD type = 0;
    DWORD data_size = 0;
    FILETIME modified = {};
    uint64_t row_id = 0;
    uint32_t match_start = 0;
    uint16_t source = 0;
    uint32_t match_length = 0;
    MatchField match_field = MatchField::kNone;
    ResultKind kind = ResultKind::kValue;
    DataState data_state = DataState::kNotApplicable;
};

bool IsExcludedPath(const std::wstring& path, const std::vector<std::wstring>& excludes);
bool IsKeyRow(const Result& result) noexcept;
std::wstring_view DisplayName(const Result& result) noexcept;
std::wstring TypeText(const Result& result);
std::wstring SourceText(const Result& result, const std::vector<Source>& sources);

using ProgressCallback = std::function<void(uint64_t searched, uint64_t total)>;
using ResultBatch = std::vector<Result>;
using BatchCallback = std::function<bool(ResultBatch&&)>;

bool Run(const Criteria& criteria, std::atomic_bool* cancel_flag, const BatchCallback& publish, const ProgressCallback& progress, regex::Status* status = nullptr);

void SortResults(std::vector<Result>* results, int column, bool ascending, const util::SortKeys& earlier = {}, const std::vector<Source>& sources = {});

}
