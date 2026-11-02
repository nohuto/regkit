// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/search.h"

#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/registry_native.h"
#include "win32/registry_view.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/date_time.h>
#include <win32ui/system/worker.h>

#include <algorithm>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <exception>
#include <functional>
#include <iterator>
#include <map>
#include <mutex>
#include <string_view>
#include <thread>
#include <utility>
#include <vector>

#include <windows.h>

namespace regkit::search
{

bool SameSource(const Source& first, const Source& second) noexcept
{
    return first.kind == second.kind && first.view == second.view && util::EqualsInsensitive(first.name, second.name);
}

std::wstring SourceLabel(const Source& source)
{
    switch (source.kind)
    {
    case Source::Kind::kRemote:
        return source.name.empty() ? util::Tr(L"Network Registry") : source.name;
    case Source::Kind::kOffline:
    case Source::Kind::kRegFile:
        return util::FileName(source.name);
    default:
        break;
    }
    return source.view ? util::Tr(L"Local Registry (32-bit)") : util::Tr(L"Local Registry");
}

const wchar_t* MatchFieldLabel(MatchField field) noexcept
{
    switch (field)
    {
    case MatchField::kPath:
        return util::Tr(L"Key");
    case MatchField::kName:
        return util::Tr(L"Value");
    case MatchField::kData:
        return util::Tr(L"Data");
    case MatchField::kComment:
        return util::Tr(L"Comment");
    case MatchField::kDefault:
        return util::Tr(L"Default");
    default:
        break;
    }
    return nullptr;
}

Matcher::Matcher(const TextOptions& options)
    : query_(options.query), use_regex_(options.use_regex), match_case_(options.match_case),
      match_whole_(options.match_whole), match_all_(options.match_all), valid_(match_all_ || !query_.empty())
{
    if (match_all_ || !valid_ || !use_regex_)
    {
        return;
    }
    regex::Options regex_options;
    regex_options.ignore_case = !match_case_;
    regex_options.whole = match_whole_;
    pattern_ = regex::Compile(query_, regex_options, &error_);
    session_ = regex::Session(pattern_);
    valid_ = session_.valid();
}

Matcher::Matcher(const Matcher& other)
    : query_(other.query_), pattern_(other.pattern_), session_(other.pattern_), error_(other.error_),
      use_regex_(other.use_regex_), match_case_(other.match_case_), match_whole_(other.match_whole_),
      match_all_(other.match_all_), valid_(other.valid_)
{
}

bool Matcher::valid() const noexcept
{
    return valid_;
}

const regex::Error& Matcher::error() const noexcept
{
    return error_;
}

Match Matcher::Find(std::wstring_view text) const
{
    Match location;
    if (!valid_)
    {
        location.status = regex::Status::kFailed;
        return location;
    }
    if (match_all_)
    {
        location.matched = true;
        location.start = 0;
        location.status = regex::Status::kMatch;
        return location;
    }
    if (use_regex_)
    {
        const regex::Found found = session_.Find(text);
        location.status = found.status;
        location.matched = found.status == regex::Status::kMatch;
        location.start = found.start;
        location.length = found.length;
        return location;
    }
    if (text.empty())
    {
        return location;
    }

    if (match_whole_)
    {
        const bool matched = match_case_
                                 ? text == query_
                                 : CompareStringOrdinal(text.data(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE) == CSTR_EQUAL;
        if (matched)
        {
            location.matched = true;
            location.start = 0;
            location.length = text.size();
        }
        return location;
    }

    if (match_case_)
    {
        const size_t position = text.find(query_);
        if (position != std::wstring::npos)
        {
            location.matched = true;
            location.start = position;
            location.length = query_.size();
        }
        return location;
    }

    const size_t position = win32ui::FindInsensitive(text, query_);
    if (position != std::wstring_view::npos)
    {
        location.matched = true;
        location.start = position;
        location.length = query_.size();
    }
    return location;
}

regex::Status Matcher::Replace(std::wstring_view text, const std::wstring& replacement, std::wstring* out) const
{
    if (!out || !valid_ || match_all_)
    {
        return regex::Status::kFailed;
    }
    if (use_regex_)
    {
        return session_.Replace(text, replacement, out, nullptr);
    }
    std::wstring replaced;
    size_t cursor = 0;
    bool matched = false;
    while (cursor < text.size())
    {
        const Match match = Find(text.substr(cursor));
        if (!match.matched)
        {
            break;
        }
        matched = true;
        replaced.append(text.substr(cursor, match.start)).append(replacement);
        cursor += match.start + match.length;
    }
    if (!matched)
    {
        return regex::Status::kNoMatch;
    }
    replaced.append(text.substr(cursor));
    *out = std::move(replaced);
    return regex::Status::kMatch;
}

bool IsExcludedPath(const std::wstring& path, const std::vector<std::wstring>& excludes)
{
    if (excludes.empty())
    {
        return false;
    }
    for (const auto& exclude : excludes)
    {
        if (!exclude.empty() && util::StartsWithInsensitive(path, exclude) &&
            (path.size() == exclude.size() || path[exclude.size()] == L'\\'))
        {
            return true;
        }
    }
    return false;
}

bool IsKeyRow(const Result& result) noexcept
{
    return result.kind == ResultKind::kKey || result.kind == ResultKind::kTraceKey;
}

std::wstring_view DisplayName(const Result& result) noexcept
{
    if (IsKeyRow(result))
    {
        return std::wstring_view();
    }
    return result.value_name.empty() ? std::wstring_view(util::Tr(L"(Default)")) : std::wstring_view(result.value_name);
}

std::wstring TypeText(const Result& result)
{
    if (IsKeyRow(result))
    {
        return util::Tr(L"Key");
    }
    if (result.kind == ResultKind::kTraceValue)
    {
        return L"TRACE";
    }
    return value_format::TypeName(result.type);
}

std::wstring SourceText(const Result& result, const std::vector<Source>& sources)
{
    if (result.source >= sources.size())
    {
        return {};
    }
    std::wstring label = SourceLabel(sources[result.source]);
    if (const wchar_t* field = MatchFieldLabel(result.match_field))
    {
        label.append(L" (").append(field).append(L")");
    }
    return label;
}

namespace
{

int CompareNumeric(uint64_t left, uint64_t right)
{
    if (left == right)
    {
        return 0;
    }
    return left < right ? -1 : 1;
}

int CompareLocation(const Result& left, const Result& right)
{
    int result = util::CompareListText(left.key_path, right.key_path);
    if (result == 0)
    {
        result = CompareNumeric(!IsKeyRow(left), !IsKeyRow(right));
    }
    return result != 0 ? result : util::CompareInsensitive(left.value_name, right.value_name);
}

int CompareResult(const Result& left, const Result& right, int column)
{
    switch (column)
    {
    case -1:
        return CompareLocation(left, right);
    case 1:
        return util::CompareListText(DisplayName(left), DisplayName(right));
    case 3:
        return util::CompareListText(left.data_text, right.data_text);
    case 4:
        return CompareNumeric(left.data_size, right.data_size);
    case 5:
        return CompareNumeric(win32ui::FileTimeTicks(left.modified), win32ui::FileTimeTicks(right.modified));
    default:
        return util::CompareListText(left.key_path, right.key_path);
    }
}

}

void SortResults(std::vector<Result>* results, int column, bool ascending, const util::SortKeys& earlier, const std::vector<Source>& sources)
{
    if (!results || results->size() < 2)
    {
        return;
    }
    // cache labels because sorting calls the comparator repeatedly
    std::map<std::pair<ResultKind, DWORD>, std::wstring> labels;
    auto label_of = [&labels](const Result& row) -> const std::wstring& {
        const auto key = std::make_pair(row.kind, row.type);
        auto found = labels.find(key);
        if (found == labels.end())
        {
            found = labels.emplace(key, TypeText(row)).first;
        }
        return found->second;
    };
    std::map<std::pair<uint16_t, MatchField>, std::wstring> source_labels;
    auto source_of = [&](const Result& row) -> const std::wstring& {
        const auto key = std::make_pair(row.source, row.match_field);
        auto found = source_labels.find(key);
        if (found == source_labels.end())
        {
            found = source_labels.emplace(key, SourceText(row, sources)).first;
        }
        return found->second;
    };
    std::stable_sort(results->begin(), results->end(), [&](const Result& left, const Result& right) {
        return util::CompareSorted(column, ascending, earlier, [&](int key) {
                   return key == 2   ? util::CompareListText(label_of(left), label_of(right))
                          : key == 6 ? util::CompareListText(source_of(left), source_of(right))
                                     : CompareResult(left, right, key);
               }) < 0;
    });
}

namespace
{

struct RootContext
{
    HKEY root = nullptr;
    REGSAM view = 0;
    uint16_t source = 0;
    std::wstring root_name;
    std::wstring display_root;
    std::wstring base_subkey;
    std::wstring mirror_root;
    std::wstring mirror_prefix;
    bool machine_classes = false;
};

struct NodeTask
{
    const RootContext* context = nullptr;
    std::wstring subkey;
    bool no_store_copy = false;
};

std::wstring BuildDisplayPath(const NodeTask& task)
{
    if (!task.context)
    {
        return task.subkey;
    }
    if (task.subkey.empty())
    {
        return task.context->display_root;
    }
    if (task.context->display_root.empty())
    {
        return registry_path::DisplayName(task.subkey);
    }
    std::wstring path;
    path.reserve(task.context->display_root.size() + task.subkey.size() + 1);
    path.append(task.context->display_root);
    path.push_back(L'\\');
    path.append(registry_path::DisplayName(task.subkey));
    return path;
}

std::wstring_view TaskLeaf(const NodeTask& task)
{
    if (task.subkey.empty())
    {
        return task.context ? std::wstring_view(task.context->display_root) : std::wstring_view();
    }
    const size_t slash = task.subkey.find_last_of(L'\\');
    if (slash == std::wstring::npos)
    {
        return std::wstring_view(task.subkey);
    }
    return std::wstring_view(task.subkey).substr(slash + 1);
}

std::wstring BuildMirrorPath(const NodeTask& task)
{
    const RootContext* context = task.context;
    if (!context || context->mirror_root.empty())
    {
        return std::wstring();
    }
    const std::wstring& prefix = context->mirror_prefix;
    const std::wstring& subkey = task.subkey;
    if (!util::StartsWithInsensitive(subkey, prefix) ||
        (subkey.size() > prefix.size() && subkey[prefix.size()] != L'\\'))
    {
        return std::wstring();
    }
    std::wstring path = context->mirror_root;
    if (subkey.size() > prefix.size())
    {
        path.push_back(L'\\');
        path.append(registry_path::DisplayName(std::wstring_view(subkey).substr(prefix.size() + 1)));
    }
    return path;
}

bool UserClassesOverrides(const std::wstring& relative, const std::wstring* value_name)
{
    std::wstring path = L"SOFTWARE\\Classes";
    if (!relative.empty())
    {
        path.push_back(L'\\');
        path.append(relative);
    }
    win32ui::UniqueHKey key;
    if (util::OpenRegistryPath(HKEY_CURRENT_USER, path, KEY_QUERY_VALUE | win32::kDefaultRegistryView, false, &key) !=
        ERROR_SUCCESS)
    {
        return false;
    }
    return !value_name ||
           RegQueryValueExW(key.get(), value_name->c_str(), nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS;
}

RegistryNode TaskNode(const NodeTask& task)
{
    RegistryNode node;
    if (task.context)
    {
        node.root = task.context->root;
        node.view = task.context->view;
        node.root_name = task.context->root_name;
        const std::wstring& base = task.context->base_subkey;
        if (!base.empty())
        {
            node.subkey.reserve(base.size() + task.subkey.size() + 1);
            node.subkey.append(base);
            if (!task.subkey.empty())
            {
                node.subkey.push_back(L'\\');
                node.subkey.append(task.subkey);
            }
            return node;
        }
    }
    node.subkey = task.subkey;
    return node;
}

struct HexQuery
{
    bool hex_only = false;
    bool parsed = false;
    std::vector<BYTE> bytes;
};
HexQuery ParseHexQuery(const std::wstring& query)
{
    HexQuery result;
    std::wstring digits;
    digits.reserve(query.size());
    size_t start = 0;
    if (query.size() >= 2 && query[0] == L'0' && (query[1] == L'x' || query[1] == L'X'))
    {
        start = 2;
    }
    for (size_t i = start; i < query.size(); ++i)
    {
        wchar_t ch = query[i];
        if (util::HexDigitValue(ch) >= 0)
        {
            digits.push_back(ch);
        }
        else if (ch == L' ' || ch == L'\t' || ch == L',' || ch == L';' || ch == L'-' || ch == L':')
        {
            continue;
        }
        else
        {
            return result;
        }
    }
    result.hex_only = true;
    if (digits.empty())
    {
        return result;
    }
    if ((digits.size() % 2) != 0)
    {
        digits.insert(digits.begin(), L'0');
    }
    result.bytes.reserve(digits.size() / 2);
    for (size_t i = 0; i < digits.size(); i += 2)
    {
        result.bytes.push_back(
            static_cast<BYTE>((util::HexDigitValue(digits[i]) << 4) | util::HexDigitValue(digits[i + 1]))
        );
    }
    result.parsed = !result.bytes.empty();
    return result;
}

bool BuildStringView(const BYTE* data, DWORD size, std::wstring_view* view)
{
    if (!view)
    {
        return false;
    }
    *view = std::wstring_view();
    if (!data || size < sizeof(wchar_t))
    {
        return false;
    }
    size_t count = size / sizeof(wchar_t);
    if (count == 0)
    {
        return false;
    }
    const wchar_t* text = reinterpret_cast<const wchar_t*>(data);
    while (count > 0 && text[count - 1] == L'\0')
    {
        --count;
    }
    *view = std::wstring_view(text, count);
    return true;
}

bool IsBinaryType(DWORD base_type)
{
    return base_type == REG_BINARY || base_type == REG_RESOURCE_LIST || base_type == REG_FULL_RESOURCE_DESCRIPTOR ||
           base_type == REG_RESOURCE_REQUIREMENTS_LIST || base_type == REG_NONE;
}

struct DataMatch
{
    bool matched = false;
    Match match;
    std::wstring data_text;
};

DataMatch MatchValueData(const Matcher& matcher, const HexQuery& hex_query, DWORD type, const BYTE* data, DWORD size, std::wstring* scratch)
{
    DataMatch result;
    if (!data && size != 0)
    {
        return result;
    }
    const auto note = [&](regex::Status status) {
        if (status != regex::Status::kMatch && status != regex::Status::kNoMatch)
        {
            result.match.status = status;
        }
    };

    DWORD base_type = value_format::NormalizeType(type);
    if (base_type == REG_MULTI_SZ)
    {
        size_t display_start = 0;
        for (const std::wstring& item : value_format::MultiStringItems({data, size - size % sizeof(wchar_t)}))
        {
            const Match match = matcher.Find(item);
            note(match.status);
            if (match.matched)
            {
                result.matched = true;
                result.match = match;
                result.match.start = display_start + match.start;
                result.data_text = value_format::DisplayData(type, data, size, false);
                return result;
            }
            display_start += item.size() + 1;
        }
        return result;
    }
    if (base_type == REG_SZ || base_type == REG_EXPAND_SZ || base_type == REG_LINK)
    {
        // search strings in place, format only matches
        std::wstring_view view;
        BuildStringView(data, size, &view);
        Match match = matcher.Find(view);
        if (!match.matched)
        {
            result.match.status = match.status;
            return result;
        }
        result.matched = true;
        result.match = match;
        std::wstring display = value_format::DisplayData(type, data, size, false);
        // show the text that matched so highlight offsets stay valid
        if (display == view)
        {
            result.data_text = std::move(display);
            return result;
        }
        const Match display_match = matcher.Find(display);
        if (display_match.matched)
        {
            result.match = display_match;
            result.data_text = std::move(display);
        }
        else
        {
            note(display_match.status);
            result.data_text = std::wstring(view);
        }
        return result;
    }

    if (IsBinaryType(base_type))
    {
        if (hex_query.hex_only && hex_query.parsed && !hex_query.bytes.empty())
        {
            const size_t needle = hex_query.bytes.size();
            if (needle <= size)
            {
                const BYTE* begin = data;
                const BYTE* end = data + size;
                const auto first = hex_query.bytes.begin();
                const auto last = hex_query.bytes.end();
                const BYTE* hit = needle <= 2 ? std::search(begin, end, first, last) : std::search(begin, end, std::boyer_moore_horspool_searcher(first, last));
                if (hit != end)
                {
                    const size_t offset = static_cast<size_t>(hit - begin);
                    result.matched = true;
                    result.data_text = value_format::Data(type, data, size);
                    result.match.matched = true;
                    result.match.start = offset * 3;
                    result.match.length = needle * 3 - 1;
                    return result;
                }
            }
        }

        auto accept_bytes = [&](size_t byte_start, size_t byte_length) {
            result.matched = true;
            result.data_text = value_format::Data(type, data, size);
            if (byte_length > 0)
            {
                result.match.matched = true;
                result.match.start = byte_start * 3;
                result.match.length = byte_length * 3 - 1;
            }
        };

        // trim trailing NULs from views, matching string value behavior
        DWORD bytes = size;
        while (bytes > 0 && data[bytes - 1] == 0)
        {
            --bytes;
        }
        if (scratch)
        {
            scratch->assign(data, data + bytes);
            const Match match = matcher.Find(*scratch);
            if (match.matched)
            {
                accept_bytes(match.start, match.length);
                return result;
            }
            note(match.status);
        }

        if (size >= sizeof(wchar_t) && (size % sizeof(wchar_t)) == 0 &&
            (reinterpret_cast<uintptr_t>(data) % alignof(wchar_t)) == 0)
        {
            const std::wstring_view wide(reinterpret_cast<const wchar_t*>(data), (bytes + 1) / sizeof(wchar_t));
            const Match match = matcher.Find(wide);
            if (match.matched)
            {
                accept_bytes(match.start * sizeof(wchar_t), match.length * sizeof(wchar_t));
                return result;
            }
            note(match.status);
        }
        return result;
    }

    std::wstring text = value_format::DisplayData(type, data, size, false);
    Match match = matcher.Find(text);
    if (!match.matched)
    {
        note(match.status);
        return result;
    }
    result.matched = true;
    result.match = match;
    result.data_text = std::move(text);
    return result;
}

std::wstring NameAnomaly(uint32_t kinds, std::wstring_view name)
{
    if ((kinds & kAnomalyNulName) && name.find(L'\0') != std::wstring_view::npos)
    {
        return util::Tr(L"Null character in the name");
    }
    const auto control = [](wchar_t c) { return (c > 0 && c < 0x20) || c == 0x7F; };
    if ((kinds & kAnomalyOddName) && !name.empty() && (name.front() == L' ' || name.back() == L' ' || std::any_of(name.begin(), name.end(), control)))
    {
        return util::Tr(L"Edge spaces or control characters in the name");
    }
    return {};
}

std::wstring ValueAnomaly(uint32_t kinds, const ValueInfo& value, const BYTE* data, DWORD size)
{
    std::wstring reason = NameAnomaly(kinds, value.name);
    const DWORD type = value.type;
    if (!reason.empty())
    {
        return reason;
    }
    if ((kinds & kAnomalyUnknownType) && type > REG_QWORD)
    {
        return util::Tr(L"Unknown type");
    }
    if ((kinds & kAnomalyIntegerSize) && (((type == REG_DWORD || type == REG_DWORD_BIG_ENDIAN) && size != sizeof(DWORD)) || (type == REG_QWORD && size != sizeof(ULONGLONG))))
    {
        return util::TrLabel(L"Unexpected size", value_format::ByteCount(size));
    }
    const bool string = type == REG_SZ || type == REG_EXPAND_SZ;
    const bool multi = type == REG_MULTI_SZ;
    if (!(kinds & (string ? kAnomalyStringEnd : multi ? kAnomalyMultiString
                                                      : 0)) ||
        size == 0)
    {
        return {};
    }
    if (size % sizeof(wchar_t))
    {
        return util::TrLabel(L"Odd size", value_format::ByteCount(size));
    }
    if (!data)
    {
        return {};
    }
    const std::wstring_view text(reinterpret_cast<const wchar_t*>(data), size / sizeof(wchar_t));
    if (string)
    {
        return text.back() ? util::Tr(L"No terminating null") : std::wstring();
    }
    if (text.back() || (text.size() >= 2 && text[text.size() - 2]))
    {
        return util::Tr(L"No final double null");
    }
    if (text.size() > 2 && text.substr(0, text.size() - 1).find(std::wstring_view(L"\0\0", 2)) != std::wstring_view::npos)
    {
        return util::Tr(L"Empty item");
    }
    return {};
}

bool VirtualStoreCopy(const RegistryNode& node, std::wstring* store)
{
    std::wstring native = node.view ? std::wstring() : registry_path::BuildNative(node);
    KeyDetails details;
    if (native.empty() && RegistryStore::QueryKeyDetails(node, &details, true))
    {
        native = std::move(details.native.native_name);
    }
    *store = registry_path::VirtualStorePath(native);
    RegistryNode copy;
    KeyInfo info = {};
    return !store->empty() && registry_path::ParseRoot(*store, &copy) && RegistryStore::QueryKeyInfo(copy, &info);
}

bool IsTypeAllowed(const Criteria& criteria, DWORD type)
{
    if (criteria.allowed_types.empty())
    {
        return true;
    }
    return std::find(criteria.allowed_types.begin(), criteria.allowed_types.end(), type) !=
           criteria.allowed_types.end();
}

bool IsSizeAllowed(const Criteria& criteria, DWORD size)
{
    if (criteria.use_min_size && size < criteria.min_size)
    {
        return false;
    }
    if (criteria.use_max_size && size > criteria.max_size)
    {
        return false;
    }
    return true;
}

bool IsKeyInRange(const Criteria& criteria, const FILETIME& last_write)
{
    if (!criteria.use_modified_from && !criteria.use_modified_to)
    {
        return true;
    }
    if (last_write.dwLowDateTime == 0 && last_write.dwHighDateTime == 0)
    {
        return false;
    }
    if (criteria.use_modified_from)
    {
        if (CompareFileTime(&last_write, &criteria.modified_from) < 0)
        {
            return false;
        }
    }
    if (criteria.use_modified_to)
    {
        if (CompareFileTime(&last_write, &criteria.modified_to) > 0)
        {
            return false;
        }
    }
    return true;
}

}

namespace
{

constexpr size_t kResultBatchSize = 128;
constexpr size_t kNodeChunkSize = 24;
constexpr uint64_t kProgressKeyInterval = 256;
constexpr uint64_t kProgressTickInterval = 200;
constexpr unsigned int kLocalWorkerLimit = 4;
constexpr unsigned int kRemoteWorkerLimit = 2;
constexpr unsigned int kOfflineWorkerLimit = 4;

unsigned int WorkerPolicy(const Criteria& criteria)
{
    if (!criteria.recursive)
    {
        return 1u;
    }
    Provider provider = criteria.provider;
    if (provider == Provider::kLocal)
    {
        for (const auto& start : criteria.start_nodes)
        {
            const RegistryNode& node = start.node;
            if (RegistryStore::IsVirtualRoot(node.root))
            {
                provider = Provider::kVirtual;
                break;
            }
            if (RegistryStore::IsOfflineRoot(node.root))
            {
                provider = Provider::kOffline;
                break;
            }
        }
    }
    switch (provider)
    {
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

}

bool Run(const Criteria& criteria, std::atomic_bool* cancel_flag, const BatchCallback& publish, const ProgressCallback& progress, regex::Status* status)
{
    if ((criteria.query.empty() && !criteria.anomalies) || criteria.start_nodes.empty())
    {
        return false;
    }

    TextOptions match_options;
    match_options.query = criteria.query;
    match_options.match_case = criteria.match_case;
    match_options.match_whole = criteria.match_whole;
    match_options.use_regex = criteria.use_regex;
    const bool match_all = criteria.query.empty();
    match_options.match_all = match_all;
    const Matcher base_matcher = criteria.matcher ? Matcher(*criteria.matcher) : Matcher(match_options);
    if (!base_matcher.valid())
    {
        return false;
    }
    std::atomic<int> worst_status(static_cast<int>(regex::Status::kNoMatch));
    const HexQuery hex_query = criteria.use_regex || criteria.query.empty() ? HexQuery() : ParseHexQuery(criteria.query);
    const bool has_excludes = !criteria.exclude_paths.empty();
    const uint32_t value_anomalies = criteria.anomalies & kValueAnomalies;
    const uint32_t key_anomalies = criteria.anomalies & kKeyAnomalies;
    const bool watch_links = (key_anomalies & kAnomalyBrokenLink) && criteria.skip_links;
    // skip values/subkeys not needed by selected search fields
    const bool want_values = criteria.search_values || criteria.search_data || criteria.comment_text || criteria.default_text || value_anomalies || watch_links;
    const bool want_data = criteria.search_data || (criteria.anomalies & kDataAnomalies);
    const bool want_subkeys = criteria.recursive;
    // values above the size limit need metadata only
    const DWORD enum_max_data =
        criteria.use_max_size && criteria.max_size < MAXDWORD ? static_cast<DWORD>(criteria.max_size) : MAXDWORD;

    std::vector<std::unique_ptr<RootContext>> contexts;
    contexts.reserve(criteria.start_nodes.size());
    std::mutex mutex;
    std::condition_variable cv;
    std::vector<NodeTask> stack;
    stack.reserve(criteria.start_nodes.size());

    const wchar_t kClassesSubkey[] = L"SOFTWARE\\Classes";
    std::wstring merged_classes_root;
    bool machine_in_scope = false;
    bool user_in_scope = false;
    if (criteria.provider == Provider::kLocal)
    {
        for (const auto& start : criteria.start_nodes)
        {
            const RegistryNode& node = start.node;
            if (!node.subkey.empty())
            {
                continue;
            }
            if (node.root == HKEY_CLASSES_ROOT)
            {
                RegistryNode root_only = node;
                merged_classes_root = registry_path::Build(root_only);
            }
            else if (node.root == HKEY_LOCAL_MACHINE)
            {
                machine_in_scope = true;
            }
            else if (node.root == HKEY_CURRENT_USER)
            {
                user_in_scope = true;
            }
        }
    }
    const bool mirror_classes = !merged_classes_root.empty();

    for (const auto& start : criteria.start_nodes)
    {
        const RegistryNode& node = start.node;
        RegistryNode root_only = node;
        root_only.subkey.clear();
        const std::wstring display_root = registry_path::Build(root_only);
        const bool whole_root = node.subkey.empty();

        struct Store
        {
            HKEY root;
            const wchar_t* base;
        };
        Store stores[2] = {{node.root, nullptr}, {nullptr, nullptr}};
        size_t store_count = 1;
        // search HKCR through backing stores not already covered by another root
        if (mirror_classes && whole_root && node.root == HKEY_CLASSES_ROOT)
        {
            store_count = 0;
            if (!machine_in_scope)
            {
                stores[store_count++] = {HKEY_LOCAL_MACHINE, kClassesSubkey};
            }
            if (!user_in_scope)
            {
                stores[store_count++] = {HKEY_CURRENT_USER, kClassesSubkey};
            }
        }

        for (size_t i = 0; i < store_count; ++i)
        {
            auto context = std::make_unique<RootContext>();
            context->root = stores[i].root;
            context->view = node.view;
            context->source = start.source;
            context->display_root = display_root;
            if (stores[i].base)
            {
                context->base_subkey = stores[i].base;
                context->machine_classes = stores[i].root == HKEY_LOCAL_MACHINE;
            }
            else
            {
                context->root_name = node.root_name;
                if (mirror_classes && whole_root && (node.root == HKEY_LOCAL_MACHINE || node.root == HKEY_CURRENT_USER))
                {
                    context->mirror_root = merged_classes_root;
                    context->mirror_prefix = kClassesSubkey;
                    context->machine_classes = node.root == HKEY_LOCAL_MACHINE;
                }
            }
            NodeTask task;
            task.context = context.get();
            task.subkey = node.subkey;
            stack.push_back(std::move(task));
            contexts.push_back(std::move(context));
        }
    }

    std::reverse(stack.begin(), stack.end());
    std::atomic<uint64_t> searched_keys(0);
    std::atomic<uint64_t> total_keys(stack.size());
    std::atomic<uint64_t> last_reported(0);
    std::atomic<uint64_t> last_reported_tick(0);
    int active = 0;
    bool done = false;
    std::atomic_bool stop(false);
    std::mutex publish_mutex;

    auto should_stop = [&]() -> bool { return stop.load() || (cancel_flag && cancel_flag->load()); };

    auto request_stop = [&]() {
        stop.store(true);
        cv.notify_all();
    };

    auto report_progress = [&](bool force) {
        if (!progress)
        {
            return;
        }
        uint64_t searched = searched_keys.load();
        const uint64_t total = total_keys.load();
        uint64_t last = last_reported.load();
        const uint64_t now = GetTickCount64();
        const uint64_t last_tick = last_reported_tick.load();
        if (!force && searched - last < kProgressKeyInterval && now - last_tick < kProgressTickInterval &&
            total != searched)
        {
            return;
        }
        if (last_reported.compare_exchange_strong(last, searched))
        {
            last_reported_tick.store(now);
            progress(searched, total);
        }
    };

    const uint64_t max_results = criteria.max_results;
    const auto rank_less = [](const Result& left, const Result& right) { return CompareLocation(left, right) < 0; };
    std::vector<Result> ranked;
    std::mutex cutoff_mutex;
    std::atomic<uint64_t> cutoff_version{0};
    std::wstring cutoff_path;

    auto publish_batch = [&](ResultBatch& batch) -> bool {
        if (batch.empty())
        {
            return true;
        }
        bool accepted = true;
        {
            std::lock_guard<std::mutex> lock(publish_mutex);
            if (max_results > 0)
            {
                std::erase_if(batch, [&](const Result& result) {
                    if (ranked.size() == max_results)
                    {
                        if (!rank_less(result, ranked.front()))
                        {
                            return true;
                        }
                        std::pop_heap(ranked.begin(), ranked.end(), rank_less);
                        ranked.pop_back();
                    }
                    Result rank;
                    rank.key_path = result.key_path;
                    rank.value_name = result.value_name;
                    rank.kind = result.kind;
                    ranked.push_back(std::move(rank));
                    std::push_heap(ranked.begin(), ranked.end(), rank_less);
                    return false;
                });
                if (ranked.size() == max_results && cutoff_path != ranked.front().key_path)
                {
                    std::lock_guard<std::mutex> cutoff_lock(cutoff_mutex);
                    cutoff_path = ranked.front().key_path;
                    cutoff_version.fetch_add(1);
                }
            }
            if (!batch.empty())
            {
                accepted = !publish || publish(std::move(batch));
            }
        }
        batch.clear();
        batch.reserve(kResultBatchSize);
        if (!accepted)
        {
            request_stop();
        }
        return accepted;
    };

    auto worker = [&]() {
        // one match session per worker
        const Matcher matcher(base_matcher);
        auto record_status = [&](regex::Status found) {
            if (found != regex::Status::kMatch && found != regex::Status::kNoMatch)
            {
                int previous = worst_status.load();
                while (previous == static_cast<int>(regex::Status::kNoMatch) &&
                       !worst_status.compare_exchange_weak(previous, static_cast<int>(found)))
                {
                }
            }
        };
        // reuse worker buffers across keys
        EnumerationScratch scratch;
        std::wstring widen_scratch;
        ResultBatch batch;
        batch.reserve(kResultBatchSize);
        std::vector<NodeTask> local;
        std::vector<NodeTask> children;
        uint64_t local_searched = 0;
        uint64_t local_tick = GetTickCount64();
        std::wstring local_cutoff;
        uint64_t local_cutoff_version = 0;

        auto flush_progress = [&](bool force) {
            if (local_searched == 0 && !force)
            {
                return;
            }
            const uint64_t now = GetTickCount64();
            if (!force && local_searched < kProgressKeyInterval && now - local_tick < kProgressTickInterval)
            {
                return;
            }
            searched_keys.fetch_add(local_searched);
            local_searched = 0;
            local_tick = now;
            report_progress(force);
        };

        for (;;)
        {
            if (local.empty())
            {
                std::unique_lock<std::mutex> lock(mutex);
                cv.wait(lock, [&]() { return done || should_stop() || !stack.empty(); });
                if (done || should_stop())
                {
                    break;
                }
                if (stack.empty())
                {
                    continue;
                }

                const size_t take = std::min(kNodeChunkSize, stack.size());
                local.insert(local.end(), std::make_move_iterator(stack.end() - take), std::make_move_iterator(stack.end()));
                stack.erase(stack.end() - take, stack.end());
                active += 1;
            }

            while (!local.empty() && !should_stop())
            {
                NodeTask entry = std::move(local.back());
                local.pop_back();
                ++local_searched;
                flush_progress(false);

                std::wstring display_path;
                // build display paths only for filters and matches
                auto path_text = [&]() -> const std::wstring& {
                    if (display_path.empty())
                    {
                        display_path = BuildDisplayPath(entry);
                    }
                    return display_path;
                };

                const bool classes_direct =
                    entry.context && entry.context->machine_classes && !entry.context->base_subkey.empty();
                auto classes_shadowed = [&](const std::wstring* value_name) -> bool {
                    if (!entry.context || !entry.context->machine_classes)
                    {
                        return false;
                    }
                    if (classes_direct)
                    {
                        return UserClassesOverrides(entry.subkey, value_name);
                    }
                    const size_t prefix = entry.context->mirror_prefix.size();
                    return UserClassesOverrides(
                        entry.subkey.size() > prefix ? entry.subkey.substr(prefix + 1) : std::wstring(),
                        value_name
                    );
                };

                std::wstring mirror_path;
                bool mirror_built = false;
                auto mirror_text = [&]() -> const std::wstring& {
                    if (!mirror_built)
                    {
                        mirror_built = true;
                        mirror_path = BuildMirrorPath(entry);
                        if (!mirror_path.empty() && has_excludes && IsExcludedPath(mirror_path, criteria.exclude_paths))
                        {
                            mirror_path.clear();
                        }
                    }
                    return mirror_path;
                };

                if (has_excludes && IsExcludedPath(path_text(), criteria.exclude_paths))
                {
                    continue;
                }
                if (max_results > 0 && cutoff_version.load(std::memory_order_relaxed) != local_cutoff_version)
                {
                    std::lock_guard<std::mutex> lock(cutoff_mutex);
                    local_cutoff = cutoff_path;
                    local_cutoff_version = cutoff_version.load();
                }
                if (!local_cutoff.empty() && util::CompareListText(path_text(), local_cutoff) > 0 &&
                    (mirror_text().empty() || util::CompareListText(mirror_text(), local_cutoff) > 0))
                {
                    continue;
                }

                RegistryStore::KeyEnumResult enum_result;
                bool key_range_checked = false;
                bool key_in_range = true;

                auto is_key_in_range = [&]() -> bool {
                    if (!criteria.use_modified_from && !criteria.use_modified_to)
                    {
                        return true;
                    }
                    if (!key_range_checked)
                    {
                        key_range_checked = true;
                        key_in_range = enum_result.info_valid && IsKeyInRange(criteria, enum_result.info.last_write);
                    }
                    return key_in_range;
                };

                children.clear();
                bool saw_link = false;
                auto value_cb = [&](const ValueInfo& value, const BYTE* data, DWORD data_size) -> bool {
                    if (should_stop())
                    {
                        return false;
                    }
                    saw_link = saw_link || (value.type == REG_LINK && util::EqualsInsensitive(value.name, L"SymbolicLinkValue"));
                    // reject metadata mismatches before searching names/data
                    if ((criteria.skip_links && value.type == REG_LINK) || !IsTypeAllowed(criteria, value.type) ||
                        !IsSizeAllowed(criteria, data_size) || !is_key_in_range())
                    {
                        return true;
                    }
                    std::wstring anomaly;
                    if (criteria.anomalies)
                    {
                        anomaly = value_anomalies ? ValueAnomaly(value_anomalies, value, data, data_size) : std::wstring();
                        if (anomaly.empty())
                        {
                            return true;
                        }
                    }
                    const std::wstring display_name = value.name.empty() ? std::wstring(util::Tr(L"(Default)")) : value.name;
                    Match name_match;
                    if (criteria.search_values && !match_all)
                    {
                        name_match = matcher.Find(display_name);
                        record_status(name_match.status);
                    }

                    DataMatch data_match;
                    // matching names need no data search
                    if (!name_match.matched && criteria.search_data && !match_all)
                    {
                        data_match = MatchValueData(matcher, hex_query, value.type, data, data_size, &widen_scratch);
                        record_status(data_match.match.status);
                    }
                    MatchField extra_field = MatchField::kNone;
                    if (!name_match.matched && !data_match.matched && !match_all)
                    {
                        const auto extra = [&](const std::wstring& text, MatchField field) {
                            const Match match = matcher.Find(text);
                            record_status(match.status);
                            extra_field = match.matched ? field : MatchField::kNone;
                        };
                        if (criteria.comment_text)
                        {
                            extra(criteria.comment_text(path_text(), &value.name, value.type, data_size), MatchField::kComment);
                        }
                        if (extra_field == MatchField::kNone && criteria.default_text)
                        {
                            extra(criteria.default_text(path_text(), value.name), MatchField::kDefault);
                        }
                        if (extra_field == MatchField::kNone)
                        {
                            return true;
                        }
                    }

                    Result result;
                    result.source = entry.context ? entry.context->source : 0;
                    result.key_path = path_text();
                    result.value_name = value.name;
                    result.type = value.type;
                    result.data_size = data_size;
                    result.kind = ResultKind::kValue;
                    if (enum_result.info_valid)
                    {
                        result.modified = enum_result.info.last_write;
                    }
                    if (data_match.matched)
                    {
                        result.data_text = std::move(data_match.data_text);
                        result.data_state = DataState::kLoaded;
                        result.match_field = MatchField::kData;
                        if (data_match.match.matched)
                        {
                            result.match_start = static_cast<uint32_t>(data_match.match.start);
                            result.match_length = static_cast<uint32_t>(data_match.match.length);
                        }
                    }
                    else
                    {
                        // defer loading name match data until requested by the UI
                        result.data_state = DataState::kNotLoaded;
                        result.match_field = extra_field;
                    }
                    if (name_match.matched)
                    {
                        result.match_field = MatchField::kName;
                        result.match_start = static_cast<uint32_t>(name_match.start);
                        result.match_length = static_cast<uint32_t>(name_match.length);
                    }
                    else if (!anomaly.empty())
                    {
                        result.match_field = MatchField::kData;
                        result.match_start = 0;
                        result.match_length = 0;
                    }
                    if (!anomaly.empty())
                    {
                        result.data_text = std::move(anomaly);
                        result.data_state = DataState::kLoaded;
                    }

                    const bool shadowed = (classes_direct || !mirror_text().empty()) && classes_shadowed(&value.name);
                    if (shadowed && classes_direct)
                    {
                        return true;
                    }
                    if (!shadowed && !mirror_text().empty())
                    {
                        Result merged = result;
                        merged.key_path = mirror_text();
                        batch.push_back(std::move(merged));
                    }
                    batch.push_back(std::move(result));
                    if (batch.size() >= kResultBatchSize && !publish_batch(batch))
                    {
                        return false;
                    }
                    return true;
                };

                auto subkey_cb = [&](const std::wstring& name) -> bool {
                    if (should_stop())
                    {
                        return false;
                    }
                    NodeTask child;
                    child.context = entry.context;
                    child.subkey.reserve(entry.subkey.size() + name.size() + 1);
                    child.subkey.append(entry.subkey);
                    if (!entry.subkey.empty())
                    {
                        child.subkey.push_back(L'\\');
                    }
                    child.subkey.append(name);
                    children.push_back(std::move(child));
                    return true;
                };

                const RegistryNode node = TaskNode(entry);
                const bool enumerated = RegistryStore::EnumKeyStreaming(
                    node,
                    want_values,
                    want_data,
                    want_subkeys,
                    &enum_result,
                    want_values ? RegistryStore::ValueStreamCallback(value_cb) : RegistryStore::ValueStreamCallback(),
                    want_subkeys ? RegistryStore::SubkeyStreamCallback(subkey_cb)
                                 : RegistryStore::SubkeyStreamCallback(),
                    enum_max_data,
                    &scratch,
                    false,
                    criteria.skip_links
                );

                std::wstring key_anomaly;
                if (key_anomalies)
                {
                    std::wstring store;
                    if (key_anomalies & kAnomalyVirtualStore)
                    {
                        const bool machine = node.root == HKEY_LOCAL_MACHINE;
                        const bool in_scope = machine && registry_path::HasComponentPrefix(node.subkey, L"SOFTWARE");
                        const bool copy = in_scope && !entry.no_store_copy && VirtualStoreCopy(node, &store);
                        const bool prune = !machine || (!node.subkey.empty() && !copy);
                        for (NodeTask& child : children)
                        {
                            child.no_store_copy = prune;
                        }
                        if (copy)
                        {
                            key_anomaly = util::TrLabel(L"Virtual store copy", store);
                        }
                    }
                    const std::wstring name_anomaly = NameAnomaly(key_anomalies, TaskLeaf(entry));
                    std::wstring target;
                    if (!name_anomaly.empty())
                    {
                        key_anomaly = name_anomaly;
                    }
                    else if ((key_anomalies & kAnomalyBrokenLink) && (!enumerated || saw_link) && RegistryStore::IsBrokenLink(node, &target))
                    {
                        key_anomaly = util::TrLabel(L"Link target missing", registry_path::DisplayName(target));
                    }
                }
                const bool key_wanted = criteria.anomalies ? !key_anomaly.empty() : enumerated && enum_result.info_valid;
                if (key_wanted && (criteria.search_keys || criteria.comment_text || match_all) && is_key_in_range())
                {
                    const std::wstring leaf = registry_path::DisplayName(TaskLeaf(entry));
                    Match key_match;
                    if (criteria.search_keys && !match_all)
                    {
                        key_match = matcher.Find(leaf);
                        record_status(key_match.status);
                    }
                    Match comment_match;
                    if (!key_match.matched && criteria.comment_text && !match_all)
                    {
                        comment_match = matcher.Find(criteria.comment_text(path_text(), nullptr, 0, 0));
                        record_status(comment_match.status);
                    }
                    if (key_match.matched || comment_match.matched || match_all)
                    {
                        Result result;
                        result.source = entry.context ? entry.context->source : 0;
                        result.key_path = path_text();
                        result.kind = ResultKind::kKey;
                        result.data_state = DataState::kNotApplicable;
                        result.data_text = std::move(key_anomaly);
                        result.modified = enum_result.info_valid ? enum_result.info.last_write : FILETIME{};
                        const size_t path_start =
                            result.key_path.size() >= leaf.size() ? result.key_path.size() - leaf.size() : 0;
                        result.match_field = key_match.matched ? MatchField::kPath : comment_match.matched ? MatchField::kComment
                                                                                                           : MatchField::kData;
                        result.match_start = key_match.matched ? static_cast<uint32_t>(path_start + key_match.start) : 0u;
                        result.match_length = static_cast<uint32_t>(key_match.length);

                        const bool shadowed = (classes_direct || !mirror_text().empty()) && classes_shadowed(nullptr);
                        if (!shadowed && !mirror_text().empty())
                        {
                            Result merged = result;
                            merged.key_path = mirror_text();
                            const size_t merged_start =
                                merged.key_path.size() >= leaf.size() ? merged.key_path.size() - leaf.size() : 0;
                            merged.match_start = key_match.matched ? static_cast<uint32_t>(merged_start + key_match.start) : 0u;
                            batch.push_back(std::move(merged));
                        }
                        if (!shadowed || !classes_direct)
                        {
                            batch.push_back(std::move(result));
                        }
                        if (batch.size() >= kResultBatchSize && !publish_batch(batch))
                        {
                            break;
                        }
                    }
                }

                if (!should_stop() && want_subkeys && !children.empty())
                {
                    std::lock_guard<std::mutex> lock(mutex);
                    stack.insert(stack.end(), std::make_move_iterator(children.rbegin()), std::make_move_iterator(children.rend()));
                    total_keys.fetch_add(static_cast<uint64_t>(children.size()));
                    cv.notify_all();
                }
            }

            publish_batch(batch);
            flush_progress(false);

            {
                std::lock_guard<std::mutex> lock(mutex);
                active -= 1;
                if (stack.empty() && active == 0)
                {
                    done = true;
                    cv.notify_all();
                }
            }
            if (should_stop())
            {
                break;
            }
        }

        publish_batch(batch);
        flush_progress(true);
    };

    unsigned int worker_count = std::thread::hardware_concurrency();
    if (worker_count == 0)
    {
        worker_count = 1;
    }
    worker_count = std::min(worker_count, WorkerPolicy(criteria));
    // avoid extra threads when the search starts from only a few roots
    worker_count =
        static_cast<unsigned int>(std::min<size_t>(worker_count, std::max<size_t>(1, criteria.start_nodes.size() * 4)));

    const auto guarded_worker = [&]() {
        try
        {
            worker();
        }
        catch (...)
        {
            worst_status.store(static_cast<int>(regex::Status::kFailed));
            request_stop();
        }
    };
    std::vector<std::thread> workers;
    workers.reserve(worker_count > 0 ? worker_count - 1 : 0);
    for (unsigned int i = 1; i < worker_count; ++i)
    {
        try
        {
            win32ui::NameThread(workers.emplace_back(guarded_worker).native_handle(), L"SearchWorkerThread");
        }
        // std::system_error from CreateThread, std::bad_alloc from the state of the thread, go on with the started
        // workers
        catch (const std::exception&)
        {
            break;
        }
    }
    guarded_worker();
    for (auto& thread : workers)
    {
        thread.join();
    }

    report_progress(true);
    if (status)
    {
        *status = static_cast<regex::Status>(worst_status.load());
    }
    return true;
}

}
