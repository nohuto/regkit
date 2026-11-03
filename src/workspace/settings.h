// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>
#include <vector>

namespace regkit::workspace
{

inline constexpr int kSaveTabsLocal = 1;
inline constexpr int kSaveTabsOffline = 2;
inline constexpr int kSaveTabsRemote = 4;
inline constexpr int kSaveTabsSearch = 8;
inline constexpr int kSaveTabsCompare = 16;
inline constexpr int kSaveTabsRegFile = 32;
inline constexpr int kSaveTabsAll =
    kSaveTabsLocal | kSaveTabsOffline | kSaveTabsRemote | kSaveTabsSearch | kSaveTabsCompare | kSaveTabsRegFile;

struct Settings
{
    bool clear_history_on_exit = false;
    bool clear_tabs_on_exit = false;
    bool show_toolbar = true;
    bool show_address_bar = true;
    bool show_filter_bar = true;
    bool show_tab_control = true;
    bool show_tree = true;
    bool show_history = true;
    bool show_status_bar = true;
    bool show_keys_in_list = false;
    bool show_simulated_keys = true;
    bool show_extra_hives = false;
    bool show_value_grid = false;
    bool auto_refresh = false;
    bool simple_mode = false;
    bool hkcu_follows_shell_user = true;
    bool save_tree_state = true;
    bool save_tabs = true;
    int save_tab_kinds = kSaveTabsAll;
    bool always_run_as_admin = false;
    bool always_run_as_system = false;
    bool always_run_as_trustedinstaller = false;
    bool always_on_top = false;
    bool single_instance = true;
    bool autocomplete = true;
    bool read_only = false;
    bool auto_check_updates = false;
    bool default_reset_enabled = false;

    bool window_placement_present = false;
    int window_x = 0;
    int window_y = 0;
    int window_width = 0;
    int window_height = 0;
    int window_dpi = 0;
    bool window_maximized = false;
    int tree_width = 260;
    int history_height = 160;

    std::wstring theme_mode = L"system";
    std::wstring theme_preset;
    std::wstring icon_set = L"phosphor";
    std::wstring language;

    bool use_custom_font = false;
    std::wstring font_face;
    int font_size = 0;
    int font_weight = 400;
    bool font_italic = false;

    std::vector<std::wstring> recent_traces;
    std::vector<std::wstring> recent_defaults;
    std::vector<int> value_column_widths;
    std::vector<bool> value_column_visible;
};

Settings ParseSettings(const std::wstring& content, Settings settings = {}, std::wstring* error = nullptr);
std::wstring SerializeSettings(const Settings& settings);
Settings DefaultOptions(const Settings& settings);
bool LoadSettings(const std::wstring& path, Settings* settings, std::wstring* error = nullptr);
DWORD SaveSettings(const std::wstring& path, const Settings& settings);

}
