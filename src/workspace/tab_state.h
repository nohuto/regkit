// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>
#include <vector>

namespace regkit::workspace
{

struct PersistedTab
{
    enum class Kind
    {
        kRegistry,
        kSearch,
        kRegFile,
    };

    Kind kind = Kind::kRegistry;
    std::wstring label;
    std::wstring selected_path;
    std::vector<std::wstring> selected_values;
    std::vector<std::wstring> expanded_paths;
    std::wstring search_cache_file;
    std::wstring compare_cache_file;
    std::wstring source_path;
    std::wstring remote_machine;
    int registry_mode = 0;
    int registry_view = 0;
    int value_top_index = 0;
    bool is_compare = false;
    int compare_filter = 0;
    std::vector<int> source_kinds;
    std::vector<std::wstring> source_names;
    std::vector<int> source_views;
};

struct TabState
{
    int active_index = 0;
    std::vector<PersistedTab> tabs;
};

TabState ParseTabs(const std::wstring& content, std::wstring* error = nullptr);
std::wstring SerializeTabs(const TabState& state);
bool LoadTabs(const std::wstring& path, TabState* state, std::wstring* error = nullptr);
bool SaveTabs(const std::wstring& path, const TabState& state);

}
