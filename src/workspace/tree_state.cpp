// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "workspace/tree_state.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <unordered_set>

namespace regkit::workspace
{

void TreeState::Clear()
{
    selected_path.clear();
    expanded_paths.clear();
}

void TreeState::Normalize()
{
    std::vector<std::wstring> normalized;
    normalized.reserve(expanded_paths.size());
    std::unordered_set<std::wstring> seen;
    seen.reserve(expanded_paths.size());
    for (const std::wstring& path : expanded_paths)
    {
        if (!path.empty() && seen.insert(util::ToLower(path)).second)
        {
            normalized.push_back(path);
        }
    }
    std::sort(normalized.begin(), normalized.end(), [](const std::wstring& left, const std::wstring& right) {
        if (left.size() != right.size())
        {
            return left.size() < right.size();
        }
        return util::CompareInsensitive(left, right) < 0;
    });
    expanded_paths.swap(normalized);
}

TreeState ParseTreeState(const std::wstring& content, std::wstring* error)
{
    TreeState state;
    size_t line_number = 0;
    for (const std::wstring_view line : win32ui::LineViews(content))
    {
        ++line_number;
        if (line.empty())
        {
            continue;
        }
        const size_t separator = line.find(L'=');
        const std::wstring_view name = separator == std::wstring::npos ? std::wstring_view() : line.substr(0, separator);
        if (name != L"selected" && name != L"expanded")
        {
            if (error)
            {
                *error = util::BadLineReason(line_number);
            }
            return {};
        }
        const std::wstring key = util::TrimWhitespace(line.substr(0, separator));
        const std::wstring value = record_fields::Unescape(line.substr(separator + 1));
        if (util::EqualsInsensitive(key, L"selected"))
        {
            state.selected_path = value;
        }
        else if (util::EqualsInsensitive(key, L"expanded") && !value.empty())
        {
            state.expanded_paths.push_back(value);
        }
    }
    return state;
}

std::wstring SerializeTreeState(const TreeState& state)
{
    std::wstring content;
    if (!state.selected_path.empty())
    {
        content.append(L"selected=");
        content.append(record_fields::Escape(state.selected_path));
        content.push_back(L'\n');
    }
    for (const std::wstring& path : state.expanded_paths)
    {
        if (!path.empty())
        {
            content.append(L"expanded=");
            content.append(record_fields::Escape(path));
            content.push_back(L'\n');
        }
    }
    return content;
}

bool LoadTreeState(const std::wstring& path, TreeState* state, std::wstring* error)
{
    if (!state)
    {
        return false;
    }
    state->Clear();
    std::wstring content;
    std::wstring reason;
    if (util::ReadStateFile(path, &content, &reason) && !content.empty())
    {
        TreeState parsed = ParseTreeState(content, &reason);
        if (reason.empty())
        {
            *state = std::move(parsed);
            return true;
        }
    }
    if (error)
    {
        *error = std::move(reason);
    }
    return false;
}

bool SaveTreeState(const std::wstring& path, const TreeState& state)
{
    return !path.empty() && win32ui::WriteTextFileAtomic(path.c_str(), SerializeTreeState(state)) == ERROR_SUCCESS;
}

}
