// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "workspace/dialog_state.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"

namespace regkit::workspace
{
std::wstring DialogStatePath()
{
    const std::wstring folder = util::GetCacheFolder();
    return folder.empty() ? std::wstring() : util::JoinPath(folder, L"dialog_state.ini");
}

DialogState LoadDialogState(std::wstring* error)
{
    DialogState state;
    std::wstring content;
    const std::wstring path = DialogStatePath();
    if (path.empty() || !util::ReadStateFile(path, &content, error))
    {
        return state;
    }
    size_t line_number = 0;
    for (const std::wstring_view line : win32ui::LineViews(content))
    {
        ++line_number;
        const size_t separator = line.find(L'=');
        if (separator == std::wstring_view::npos && !line.empty())
        {
            if (error)
            {
                *error = util::BadLineReason(line_number);
            }
            return {};
        }
        if (separator != std::wstring_view::npos)
        {
            state.insert_or_assign(util::ToLower(util::TrimWhitespace(line.substr(0, separator))), std::wstring(line.substr(separator + 1)));
        }
    }
    return state;
}

void SaveDialogState(const DialogState& state)
{
    std::wstring content;
    for (const auto& [key, value] : state)
    {
        content.append(key).append(L"=").append(value).append(L"\n");
    }
    const std::wstring path = DialogStatePath();
    if (!path.empty())
    {
        win32ui::WriteTextFileAtomic(path.c_str(), content);
    }
}

void DialogFields::Field(std::wstring_view key, std::wstring* value)
{
    if (write_)
    {
        state_->insert_or_assign(std::wstring(key), *value);
    }
    else if (const auto found = state_->find(key); found != state_->end())
    {
        *value = found->second;
    }
}

void DialogFields::Field(std::wstring_view key, bool* value)
{
    std::wstring text = *value ? L"1" : L"0";
    Field(key, &text);
    *value = util::ParseBool(text);
}

void DialogFields::Field(std::wstring_view key, uint64_t* value)
{
    std::wstring text = std::to_wstring(*value);
    Field(key, &text);
    record_fields::ParseUnsigned(text, UINT64_MAX, value);
}

}
