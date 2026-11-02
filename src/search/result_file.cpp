// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/result_file.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"

#include <win32ui/system/date_time.h>

#include <cstdint>
#include <string_view>
#include <utility>

namespace regkit::search
{

namespace
{

constexpr uint64_t kMaxResultFileBytes = 256ull * 1024 * 1024;

bool ParseVersionedRecord(std::vector<std::wstring>&& fields, Result* out)
{
    uint64_t type = 0;
    uint64_t data_size = 0;
    uint64_t modified = 0;
    uint64_t match_field = 0;
    uint64_t match_start = 0;
    uint64_t match_length = 0;
    uint64_t kind = 0;
    uint64_t state = 0;
    uint64_t source = 0;
    if (!record_fields::ParseUnsigned(fields[3], MAXDWORD, &type) ||
        !record_fields::ParseUnsigned(fields[4], MAXDWORD, &data_size) ||
        !record_fields::ParseUnsigned(fields[5], UINT64_MAX, &modified) ||
        !record_fields::ParseUnsigned(fields[6], static_cast<uint64_t>(MatchField::kDefault), &match_field) ||
        !record_fields::ParseUnsigned(fields[7], UINT32_MAX, &match_start) ||
        !record_fields::ParseUnsigned(fields[8], UINT32_MAX, &match_length) ||
        !record_fields::ParseUnsigned(fields[9], static_cast<uint64_t>(ResultKind::kTraceValue), &kind) ||
        !record_fields::ParseUnsigned(fields[10], static_cast<uint64_t>(DataState::kLoaded), &state) ||
        !record_fields::ParseUnsigned(fields[11], UINT16_MAX, &source))
    {
        return false;
    }
    Result result;
    result.key_path = std::move(fields[0]);
    result.value_name = std::move(fields[1]);
    result.data_text = std::move(fields[2]);
    result.type = static_cast<DWORD>(type);
    result.data_size = static_cast<DWORD>(data_size);
    result.modified = win32ui::FileTimeFromTicks(modified);
    result.match_field = static_cast<MatchField>(match_field);
    result.match_start = static_cast<uint32_t>(match_start);
    result.match_length = static_cast<uint32_t>(match_length);
    result.kind = static_cast<ResultKind>(kind);
    result.data_state = static_cast<DataState>(state);
    result.source = static_cast<uint16_t>(source);
    *out = std::move(result);
    return true;
}

}

bool ParseResults(const std::wstring& content, std::vector<Result>* out, std::wstring* error)
{
    if (!out)
    {
        return false;
    }
    std::vector<Result> results;
    size_t line_number = 0;
    for (const std::wstring_view line : win32ui::LineViews(content))
    {
        ++line_number;
        if (line.empty())
        {
            continue;
        }
        auto fields = record_fields::DecodeRecord(line);
        Result record;
        if (fields.size() != 12 || !ParseVersionedRecord(std::move(fields), &record))
        {
            if (error)
            {
                *error = util::BadLineReason(line_number);
            }
            return false;
        }
        results.push_back(std::move(record));
    }
    *out = std::move(results);
    return true;
}

std::wstring SerializeResults(const std::vector<Result>& results)
{
    std::wstring content;
    for (const Result& result : results)
    {
        const uint64_t modified = win32ui::FileTimeTicks(result.modified);
        record_fields::AppendRecord(
            &content,
            {result.key_path, result.value_name, result.data_text, std::to_wstring(result.type), std::to_wstring(result.data_size), std::to_wstring(modified), std::to_wstring(static_cast<int>(result.match_field)), std::to_wstring(result.match_start), std::to_wstring(result.match_length), std::to_wstring(static_cast<int>(result.kind)), std::to_wstring(static_cast<int>(result.data_state)), std::to_wstring(result.source)}
        );
    }
    return content;
}

bool LoadResults(const std::wstring& path, std::vector<Result>* results, std::wstring* error)
{
    std::wstring content;
    return results && !path.empty() && util::ReadStateFile(path, &content, error, kMaxResultFileBytes) && !content.empty() && ParseResults(content, results, error);
}

bool SaveResults(const std::wstring& path, const std::vector<Result>& results)
{
    return !path.empty() && win32ui::WriteTextFileAtomic(path.c_str(), SerializeResults(results)) == ERROR_SUCCESS;
}

}
