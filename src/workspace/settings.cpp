// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "workspace/settings.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <cwchar>
#include <limits>

namespace regkit::workspace {
namespace {

bool Boolean(const std::wstring& value) {
  return _wtoi(value.c_str()) != 0 ||
         _wcsicmp(value.c_str(), L"true") == 0 ||
         _wcsicmp(value.c_str(), L"yes") == 0;
}

bool Indexed(const std::wstring& key, const wchar_t* prefix, int* index) {
  const size_t length = wcslen(prefix);
  if (!index || key.size() <= length ||
      _wcsnicmp(key.c_str(), prefix, length) != 0) {
    return false;
  }
  const wchar_t* start = key.c_str() + length;
  wchar_t* end = nullptr;
  const long value = wcstol(start, &end, 10);
  if (end == start || *end != L'\0' || value < 0 ||
      value > std::numeric_limits<int>::max()) {
    return false;
  }
  *index = static_cast<int>(value);
  return true;
}

void BooleanLine(std::wstring* output, const wchar_t* key, bool value) {
  output->append(key);
  output->append(value ? L"=1\n" : L"=0\n");
}

void TextLine(std::wstring* output, const wchar_t* key,
              const std::wstring& value) {
  output->append(key);
  output->push_back(L'=');
  output->append(value);
  output->push_back(L'\n');
}

void NumberLine(std::wstring* output, const wchar_t* key, int value) {
  TextLine(output, key, std::to_wstring(value));
}

} // namespace

Settings ParseSettings(const std::wstring& content, Settings settings) {
  for (const std::wstring& line : record_fields::Lines(content)) {
    const size_t separator = line.find(L'=');
    if (separator == std::wstring::npos) {
      continue;
    }
    const std::wstring key =
        util::TrimWhitespace(line.substr(0, separator));
    const std::wstring value =
        util::TrimWhitespace(line.substr(separator + 1));
    int index = -1;
#define REGKIT_BOOL(name, member)         \
  if (_wcsicmp(key.c_str(), name) == 0) { \
    settings.member = Boolean(value);     \
  }
    REGKIT_BOOL(L"clear_history_on_exit", clear_history_on_exit)
    else REGKIT_BOOL(L"clear_tabs_on_exit",             clear_tabs_on_exit)
    else REGKIT_BOOL(L"view_toolbar",                   show_toolbar)
    else REGKIT_BOOL(L"view_address_bar",               show_address_bar)
    else REGKIT_BOOL(L"view_filter_bar",                show_filter_bar)
    else REGKIT_BOOL(L"view_tab_control",               show_tab_control)
    else REGKIT_BOOL(L"view_tree",                      show_tree)
    else REGKIT_BOOL(L"view_history",                   show_history)
    else REGKIT_BOOL(L"view_status_bar",                show_status_bar)
    else REGKIT_BOOL(L"view_keys_in_list",              show_keys_in_list)
    else REGKIT_BOOL(L"view_simulated_keys",            show_simulated_keys)
    else REGKIT_BOOL(L"view_extra_hives",               show_extra_hives)
    else REGKIT_BOOL(L"view_value_grid",                show_value_grid)
    else REGKIT_BOOL(L"save_tree_state",                save_tree_state)
    else REGKIT_BOOL(L"save_tabs",                      save_tabs)
    else REGKIT_BOOL(L"always_run_as_admin",            always_run_as_admin)
    else REGKIT_BOOL(L"always_run_as_system",           always_run_as_system)
    else REGKIT_BOOL(L"always_run_as_trustedinstaller", always_run_as_trustedinstaller)
    else REGKIT_BOOL(L"always_on_top",                  always_on_top)
    else REGKIT_BOOL(L"single_instance",                single_instance)
    else REGKIT_BOOL(L"read_only",                      read_only)
    else REGKIT_BOOL(L"auto_check_updates",             auto_check_updates)
    else REGKIT_BOOL(L"default_reset_enabled",          default_reset_enabled)
#undef REGKIT_BOOL
        else if (_wcsicmp(key.c_str(), L"window_x") == 0) {
      settings.window_x = _wtoi(value.c_str());
      settings.window_placement_present = true;
    }
    else if (_wcsicmp(key.c_str(), L"window_y") == 0) {
      settings.window_y = _wtoi(value.c_str());
      settings.window_placement_present = true;
    }
    else if (_wcsicmp(key.c_str(), L"window_width") == 0) {
      settings.window_width = _wtoi(value.c_str());
      settings.window_placement_present = true;
    }
    else if (_wcsicmp(key.c_str(), L"window_height") == 0) {
      settings.window_height = _wtoi(value.c_str());
      settings.window_placement_present = true;
    }
    else if (_wcsicmp(key.c_str(), L"window_maximized") == 0) {
      settings.window_maximized = Boolean(value);
      settings.window_placement_present = true;
    }
    else if (_wcsicmp(key.c_str(), L"tree_width") == 0) {
      const int width = _wtoi(value.c_str());
      if (width > 0) {
        settings.tree_width = width;
      }
    }
    else if (_wcsicmp(key.c_str(), L"history_height") == 0) {
      const int height = _wtoi(value.c_str());
      if (height > 0) {
        settings.history_height = height;
      }
    }
    else if (_wcsicmp(key.c_str(), L"theme_mode") == 0) {
      settings.theme_mode = value;
    }
    else if (_wcsicmp(key.c_str(), L"theme_preset") == 0) {
      settings.theme_preset = value;
    }
    else if (_wcsicmp(key.c_str(), L"icon_set") == 0) {
      settings.icon_set = value;
    }
    else if (_wcsicmp(key.c_str(), L"font_use_default") == 0) {
      settings.use_custom_font = !Boolean(value);
    }
    else if (_wcsicmp(key.c_str(), L"font_face") == 0) {
      settings.font_face = value;
    }
    else if (_wcsicmp(key.c_str(), L"font_size") == 0) {
      const int size = _wtoi(value.c_str());
      if (size > 0) {
        settings.font_size = size;
      }
    }
    else if (_wcsicmp(key.c_str(), L"font_weight") == 0) {
      const int weight = _wtoi(value.c_str());
      if (weight > 0) {
        settings.font_weight = weight;
      }
    }
    else if (_wcsicmp(key.c_str(), L"font_italic") == 0) {
      settings.font_italic = Boolean(value);
    }
    else if (Indexed(key, L"trace_recent_", &index)) {
      if (static_cast<size_t>(index) >= settings.recent_traces.size()) {
        settings.recent_traces.resize(static_cast<size_t>(index) + 1);
      }
      settings.recent_traces[static_cast<size_t>(index)] = value;
    }
    else if (Indexed(key, L"default_recent_", &index)) {
      if (static_cast<size_t>(index) >= settings.recent_defaults.size()) {
        settings.recent_defaults.resize(static_cast<size_t>(index) + 1);
      }
      settings.recent_defaults[static_cast<size_t>(index)] = value;
    }
    else if (Indexed(key, L"value_column_width_", &index)) {
      if (static_cast<size_t>(index) >=
          settings.value_column_widths.size()) {
        settings.value_column_widths.resize(static_cast<size_t>(index) + 1);
      }
      settings.value_column_widths[static_cast<size_t>(index)] =
          _wtoi(value.c_str());
    }
    else if (Indexed(key, L"value_column_visible_", &index)) {
      if (static_cast<size_t>(index) >=
          settings.value_column_visible.size()) {
        settings.value_column_visible.resize(static_cast<size_t>(index) + 1,
                                             true);
      }
      settings.value_column_visible[static_cast<size_t>(index)] =
          Boolean(value);
    }
  }
  if (settings.always_run_as_trustedinstaller) {
    settings.always_run_as_system = false;
    settings.always_run_as_admin = false;
  } else if (settings.always_run_as_system) {
    settings.always_run_as_admin = false;
  }
  return settings;
}

std::wstring SerializeSettings(const Settings& settings) {
  std::wstring content;
  BooleanLine(&content, L"clear_history_on_exit",
              settings.clear_history_on_exit);
  BooleanLine(&content, L"clear_tabs_on_exit", settings.clear_tabs_on_exit);
  BooleanLine(&content, L"view_toolbar", settings.show_toolbar);
  BooleanLine(&content, L"view_address_bar", settings.show_address_bar);
  BooleanLine(&content, L"view_filter_bar", settings.show_filter_bar);
  BooleanLine(&content, L"view_tab_control", settings.show_tab_control);
  BooleanLine(&content, L"view_tree", settings.show_tree);
  BooleanLine(&content, L"view_history", settings.show_history);
  BooleanLine(&content, L"view_status_bar", settings.show_status_bar);
  BooleanLine(&content, L"view_keys_in_list", settings.show_keys_in_list);
  BooleanLine(&content, L"view_simulated_keys",
              settings.show_simulated_keys);
  BooleanLine(&content, L"view_extra_hives", settings.show_extra_hives);
  BooleanLine(&content, L"view_value_grid", settings.show_value_grid);
  BooleanLine(&content, L"save_tree_state", settings.save_tree_state);
  BooleanLine(&content, L"save_tabs", settings.save_tabs);
  BooleanLine(&content, L"auto_check_updates", settings.auto_check_updates);
  BooleanLine(&content, L"default_reset_enabled",
              settings.default_reset_enabled);
  BooleanLine(&content, L"always_run_as_admin",
              settings.always_run_as_admin);
  BooleanLine(&content, L"always_run_as_system",
              settings.always_run_as_system);
  BooleanLine(&content, L"always_run_as_trustedinstaller",
              settings.always_run_as_trustedinstaller);
  if (settings.window_width > 0 && settings.window_height > 0) {
    NumberLine(&content, L"window_x", settings.window_x);
    NumberLine(&content, L"window_y", settings.window_y);
    NumberLine(&content, L"window_width", settings.window_width);
    NumberLine(&content, L"window_height", settings.window_height);
    BooleanLine(&content, L"window_maximized", settings.window_maximized);
  }
  BooleanLine(&content, L"always_on_top", settings.always_on_top);
  BooleanLine(&content, L"single_instance", settings.single_instance);
  BooleanLine(&content, L"read_only", settings.read_only);
  TextLine(&content, L"theme_mode", settings.theme_mode);
  TextLine(&content, L"theme_preset", settings.theme_preset);
  TextLine(&content, L"icon_set", settings.icon_set);
  NumberLine(&content, L"tree_width", settings.tree_width);
  NumberLine(&content, L"history_height", settings.history_height);
  BooleanLine(&content, L"font_use_default", !settings.use_custom_font);
  if (!settings.font_face.empty()) {
    TextLine(&content, L"font_face", settings.font_face);
  }
  if (settings.font_size > 0) {
    NumberLine(&content, L"font_size", settings.font_size);
  }
  NumberLine(&content, L"font_weight", settings.font_weight);
  BooleanLine(&content, L"font_italic", settings.font_italic);
  for (size_t index = 0; index < settings.recent_traces.size(); ++index) {
    if (!settings.recent_traces[index].empty()) {
      TextLine(&content,
               (L"trace_recent_" + std::to_wstring(index)).c_str(),
               settings.recent_traces[index]);
    }
  }
  for (size_t index = 0; index < settings.recent_defaults.size(); ++index) {
    if (!settings.recent_defaults[index].empty()) {
      TextLine(&content,
               (L"default_recent_" + std::to_wstring(index)).c_str(),
               settings.recent_defaults[index]);
    }
  }
  const size_t columns =
      std::max(settings.value_column_widths.size(),
               settings.value_column_visible.size());
  for (size_t index = 0; index < columns; ++index) {
    const int width = index < settings.value_column_widths.size()
                          ? settings.value_column_widths[index]
                          : 0;
    const bool visible = index < settings.value_column_visible.size()
                             ? settings.value_column_visible[index]
                             : true;
    NumberLine(&content,
               (L"value_column_width_" + std::to_wstring(index)).c_str(),
               width);
    BooleanLine(
        &content,
        (L"value_column_visible_" + std::to_wstring(index)).c_str(),
        visible);
  }
  return content;
}

bool LoadSettings(const std::wstring& path, Settings* settings) {
  if (!settings) {
    return false;
  }
  std::wstring content;
  if (!util::ReadTextFile(path, &content)) {
    return false;
  }
  *settings = ParseSettings(content, std::move(*settings));
  return true;
}

bool SaveSettings(const std::wstring& path, const Settings& settings) {
  return !path.empty() &&
         util::WriteTextFile(path, SerializeSettings(settings), false);
}

} // namespace regkit::workspace
