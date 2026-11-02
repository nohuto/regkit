// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "workspace/tab_state.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"

#include <climits>

namespace regkit::workspace
{

namespace
{

struct KindName
{
    PersistedTab::Kind kind;
    const wchar_t* tag;
};

constexpr KindName kKindNames[] = {
    {PersistedTab::Kind::kRegistry, L"registry"},
    {PersistedTab::Kind::kSearch, L"search"},
    {PersistedTab::Kind::kRegFile, L"regfile"},
};

const wchar_t* KindTag(PersistedTab::Kind kind)
{
    for (const KindName& name : kKindNames)
    {
        if (name.kind == kind)
        {
            return name.tag;
        }
    }
    return kKindNames[0].tag;
}

const KindName* FindKind(std::wstring_view tag)
{
    for (const KindName& name : kKindNames)
    {
        if (util::EqualsInsensitive(tag, name.tag))
        {
            return &name;
        }
    }
    return nullptr;
}

void AppendField(std::vector<std::wstring>* fields, const wchar_t* key, const std::wstring& value)
{
    if (!value.empty())
    {
        fields->push_back(key + value);
    }
}

void AppendNumber(std::vector<std::wstring>* fields, const wchar_t* key, int value)
{
    if (value != 0)
    {
        fields->push_back(key + std::to_wstring(value));
    }
}

int Number(std::wstring_view text)
{
    uint64_t value = 0;
    return record_fields::ParseUnsigned(text, INT_MAX, &value) ? static_cast<int>(value) : 0;
}

bool ParseTaggedFields(const std::vector<std::wstring>& fields, PersistedTab* tab)
{
    for (size_t index = 3; index < fields.size(); ++index)
    {
        const std::wstring_view field = fields[index];
        const size_t separator = field.find(L'=');
        if (separator == std::wstring_view::npos)
        {
            return false;
        }
        const std::wstring_view key = field.substr(0, separator);
        const std::wstring value(field.substr(separator + 1));
        if (key == L"path")
        {
            tab->selected_path = value;
        }
        else if (key == L"sel")
        {
            tab->selected_values.push_back(value);
        }
        else if (key == L"exp")
        {
            tab->expanded_paths.push_back(value);
        }
        else if (key == L"cache")
        {
            tab->search_cache_file = value;
        }
        else if (key == L"ccache")
        {
            tab->compare_cache_file = value;
        }
        else if (key == L"src")
        {
            tab->source_path = value;
        }
        else if (key == L"machine")
        {
            tab->remote_machine = value;
        }
        else if (key == L"mode")
        {
            tab->registry_mode = Number(value);
        }
        else if (key == L"view")
        {
            tab->registry_view = Number(value);
        }
        else if (key == L"top")
        {
            tab->value_top_index = Number(value);
        }
        else if (key == L"cmp")
        {
            tab->is_compare = Number(value) != 0;
        }
        else if (key == L"cmpf")
        {
            tab->compare_filter = Number(value);
        }
        else if (key == L"srck")
        {
            tab->source_kinds.push_back(Number(value));
        }
        else if (key == L"srcn")
        {
            tab->source_names.push_back(value);
        }
        else if (key == L"srcv")
        {
            tab->source_views.push_back(Number(value));
        }
        else
        {
            return false;
        }
    }
    return true;
}

}

TabState ParseTabs(const std::wstring& content, std::wstring* error)
{
    TabState state;
    size_t line_number = 0;
    for (const std::wstring_view line : win32ui::LineViews(content))
    {
        ++line_number;
        if (line.empty())
        {
            continue;
        }
        if (line.starts_with(L"active="))
        {
            state.active_index = Number(line.substr(7));
            continue;
        }
        const auto fields = record_fields::DecodeRecord(line);
        const KindName* kind =
            fields.size() >= 3 && util::EqualsInsensitive(fields[0], L"tab") ? FindKind(fields[1]) : nullptr;
        PersistedTab tab;
        if (!kind || !ParseTaggedFields(fields, &tab))
        {
            if (error)
            {
                *error = util::BadLineReason(line_number);
            }
            return {};
        }
        tab.kind = kind->kind;
        tab.label = fields[2];
        if (tab.kind == PersistedTab::Kind::kRegFile && tab.source_path.empty())
        {
            continue;
        }
        state.tabs.push_back(std::move(tab));
    }
    return state;
}

std::wstring SerializeTabs(const TabState& state)
{
    std::wstring content = L"active=" + std::to_wstring(state.active_index) + L"\n";
    for (const PersistedTab& tab : state.tabs)
    {
        std::vector<std::wstring> fields = {L"tab", KindTag(tab.kind), tab.label};
        AppendField(&fields, L"path=", tab.selected_path);
        for (const std::wstring& value : tab.selected_values)
        {
            fields.push_back(L"sel=" + value);
        }
        for (const std::wstring& path : tab.expanded_paths)
        {
            AppendField(&fields, L"exp=", path);
        }
        AppendField(&fields, L"cache=", tab.search_cache_file);
        AppendField(&fields, L"ccache=", tab.compare_cache_file);
        AppendField(&fields, L"src=", tab.source_path);
        AppendField(&fields, L"machine=", tab.remote_machine);
        AppendNumber(&fields, L"mode=", tab.registry_mode);
        AppendNumber(&fields, L"view=", tab.registry_view);
        AppendNumber(&fields, L"top=", tab.value_top_index);
        AppendNumber(&fields, L"cmp=", tab.is_compare ? 1 : 0);
        AppendNumber(&fields, L"cmpf=", tab.compare_filter);
        for (size_t i = 0; i < tab.source_kinds.size(); ++i)
        {
            fields.push_back(L"srck=" + std::to_wstring(tab.source_kinds[i]));
            fields.push_back(L"srcn=" + (i < tab.source_names.size() ? tab.source_names[i] : std::wstring()));
            fields.push_back(L"srcv=" + std::to_wstring(i < tab.source_views.size() ? tab.source_views[i] : 0));
        }
        record_fields::AppendRecord(&content, fields);
    }
    return content;
}

bool LoadTabs(const std::wstring& path, TabState* state, std::wstring* error)
{
    std::wstring content;
    std::wstring reason;
    if (state && util::ReadStateFile(path, &content, &reason) && !content.empty())
    {
        TabState parsed = ParseTabs(content, &reason);
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

bool SaveTabs(const std::wstring& path, const TabState& state)
{
    return !path.empty() && win32ui::WriteTextFileAtomic(path.c_str(), SerializeTabs(state)) == ERROR_SUCCESS;
}

}
