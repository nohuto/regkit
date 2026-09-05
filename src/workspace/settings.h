// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>
#include <vector>

namespace regkit::workspace {

struct Settings {
  static constexpr int kCurrentVersion = 1;
  int source_version = 1;

  bool clear_history_on_exit = false;
  bool clear_tabs_on_exit = false;
  bool show_toolbar = true;
  bool show_address_bar = true;
  bool show_filter_bar = true;
  bool show_tab_control = true;
  bool show_tree = true;
  bool show_history = true;
  bool show_status_bar = true;
  bool show_keys_in_list = true;
  bool show_simulated_keys = true;
  bool show_extra_hives = false;
  bool show_value_grid = false;
  bool save_tree_state = true;
  bool save_tabs = true;
  bool always_run_as_admin = false;
  bool always_run_as_system = false;
  bool always_run_as_trustedinstaller = false;
  bool always_on_top = false;
  bool single_instance = true;
  bool read_only = false;
  bool auto_check_updates = false;

  bool window_placement_present = false;
  int window_x = 0;
  int window_y = 0;
  int window_width = 0;
  int window_height = 0;
  bool window_maximized = false;
  int tree_width = 260;
  int history_height = 160;

  std::wstring theme_mode = L"system";
  std::wstring theme_preset;
  std::wstring icon_set = L"phosphor-regedit";

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

Settings ParseSettings(const std::wstring& content,
                       Settings settings = {});
std::wstring SerializeSettings(const Settings& settings);
bool LoadSettings(const std::wstring& path, Settings* settings);
bool SaveSettings(const std::wstring& path, const Settings& settings);

} // namespace regkit::workspace
