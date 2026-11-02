// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "workspace/settings.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <cwchar>
#include <span>

namespace regkit::workspace
{
namespace
{

struct BoolField
{
    const wchar_t* key;
    bool Settings::* member;
};

struct IntField
{
    const wchar_t* key;
    int Settings::* member;
};

struct TextField
{
    const wchar_t* key;
    std::wstring Settings::* member;
};

constexpr BoolField kBoolFields[] = {
    {L"clear_history_on_exit", &Settings::clear_history_on_exit},     // Options > Clear History on Exit
    {L"clear_tabs_on_exit", &Settings::clear_tabs_on_exit},           // Options > Clear Tabs on Exit
    {L"view_toolbar", &Settings::show_toolbar},                       // View > Toolbar
    {L"view_address_bar", &Settings::show_address_bar},               // View > Address Bar
    {L"view_filter_bar", &Settings::show_filter_bar},                 // View > Filter Bar
    {L"view_tab_control", &Settings::show_tab_control},               // View > Tabs
    {L"view_tree", &Settings::show_tree},                             // View > Key Tree
    {L"view_history", &Settings::show_history},                       // View > History
    {L"view_status_bar", &Settings::show_status_bar},                 // View > Status Bar
    {L"view_keys_in_list", &Settings::show_keys_in_list},             // View > Keys in List
    {L"view_simulated_keys", &Settings::show_simulated_keys},         // View > Simulated Keys
    {L"view_extra_hives", &Settings::show_extra_hives},               // View > Show Extra Root Keys
    {L"view_value_grid", &Settings::show_value_grid},                 // View > Grid Lines
    {L"view_auto_refresh", &Settings::auto_refresh},                  // View > Auto Refresh
    {L"view_simple_mode", &Settings::simple_mode},                    // View > Simple Mode
    {L"hkcu_follows_shell_user", &Settings::hkcu_follows_shell_user}, // Options > Run As > HKCU Follows Signed-In User
    {L"save_tree_state", &Settings::save_tree_state},                 // Options > Save Previous Tree State
    {L"auto_check_updates", &Settings::auto_check_updates},           // Help > Check for Updates Automatically
    {L"default_reset_enabled", &Settings::default_reset_enabled},     // Default > Enable Context Menu (risky)
    {L"always_run_as_admin", &Settings::always_run_as_admin},         // Options > Run As > Always Run as Admin
    {L"always_run_as_system", &Settings::always_run_as_system},       // Options > Run As > Always Run as SYSTEM
    {L"always_run_as_trustedinstaller",
     &Settings::always_run_as_trustedinstaller},      // Options > Run As > Always Run as TrustedInstaller
    {L"always_on_top", &Settings::always_on_top},     // Window > Always on Top
    {L"single_instance", &Settings::single_instance}, // Options > Single Instance
    {L"autocomplete", &Settings::autocomplete},       // Options > Autocomplete Key Paths
    {L"read_only", &Settings::read_only},             // Options > Read Only Mode
    {L"font_italic", &Settings::font_italic},         // Options > Font
};

constexpr IntField kPositiveFields[] = {
    {L"tree_width", &Settings::tree_width},
    {L"history_height", &Settings::history_height},
    {L"font_size", &Settings::font_size},     // Options > Font
    {L"font_weight", &Settings::font_weight}, // Options > Font
};

constexpr IntField kPlacementFields[] = {
    {L"window_x", &Settings::window_x},
    {L"window_y", &Settings::window_y},
    {L"window_width", &Settings::window_width},
    {L"window_height", &Settings::window_height},
};

constexpr TextField kTextFields[] = {
    {L"theme_mode", &Settings::theme_mode},     // Options > Theme
    {L"theme_preset", &Settings::theme_preset}, // Options > Theme > Theme Presets
    {L"icon_set", &Settings::icon_set},         // Options > Icons
    {L"font_face", &Settings::font_face},       // Options > Font
    {L"language", &Settings::language},         // Options > Language
};

template <typename Field>
const Field* FindField(const std::wstring& key, std::span<const Field> fields)
{
    for (const Field& field : fields)
    {
        if (util::EqualsInsensitive(key, field.key))
        {
            return &field;
        }
    }
    return nullptr;
}

bool Indexed(const std::wstring& key, std::wstring_view prefix, size_t* index)
{
    uint64_t value = 0;
    if (!util::StartsWithInsensitive(key, prefix) ||
        !record_fields::ParseUnsigned(std::wstring_view(key).substr(prefix.size()), 4096, &value))
    {
        return false;
    }
    *index = static_cast<size_t>(value);
    return true;
}

template <typename T>
void SetIndexed(std::vector<T>* items, size_t index, T value, T fill = T())
{
    if (index >= items->size())
    {
        items->resize(index + 1, fill);
    }
    (*items)[index] = value;
}

void Line(std::wstring* output, std::wstring_view key, std::wstring_view value)
{
    output->append(key).append(L"=").append(value).append(L"\n");
}

}

Settings ParseSettings(const std::wstring& content, Settings settings, std::wstring* error)
{
    const Settings original = settings;
    size_t line_number = 0;
    for (const std::wstring_view line : win32ui::LineViews(content))
    {
        ++line_number;
        const size_t separator = line.find(L'=');
        if (separator == std::wstring::npos)
        {
            if (util::IsBlank(line))
            {
                continue;
            }
            if (error)
            {
                *error = util::BadLineReason(line_number);
            }
            return original;
        }
        const std::wstring key = util::TrimWhitespace(std::wstring_view(line).substr(0, separator));
        const std::wstring value = util::TrimWhitespace(std::wstring_view(line).substr(separator + 1));
        const int number = _wtoi(value.c_str());
        size_t index = 0;
        if (const auto* bool_field = FindField<BoolField>(key, kBoolFields))
        {
            settings.*(bool_field->member) = util::ParseBool(value);
        }
        else if (const auto* positive_field = FindField<IntField>(key, kPositiveFields))
        {
            if (number > 0)
            {
                settings.*(positive_field->member) = number;
            }
        }
        else if (const auto* placement_field = FindField<IntField>(key, kPlacementFields))
        {
            settings.*(placement_field->member) = number;
            settings.window_placement_present = true;
        }
        else if (const auto* text_field = FindField<TextField>(key, kTextFields))
        {
            settings.*(text_field->member) = value;
        }
        else if (util::EqualsInsensitive(key, L"window_maximized"))
        {
            settings.window_maximized = util::ParseBool(value);
            settings.window_placement_present = true;
        }
        else if (util::EqualsInsensitive(key, L"save_tabs"))
        { // Options > Save Tabs
            settings.save_tabs = util::ParseBool(value);
            settings.save_tab_kinds = settings.save_tabs ? kSaveTabsAll : 0;
        }
        else if (util::EqualsInsensitive(key, L"save_tab_types"))
        { // Options > Save Tabs > tab types
            settings.save_tab_kinds = number & kSaveTabsAll;
            settings.save_tabs = settings.save_tab_kinds != 0;
        }
        else if (util::EqualsInsensitive(key, L"font_use_default"))
        { // Options > Font
            settings.use_custom_font = !util::ParseBool(value);
        }
        else if (Indexed(key, L"trace_recent_", &index))
        { // Trace > recent trace files
            SetIndexed(&settings.recent_traces, index, value);
        }
        else if (Indexed(key, L"default_recent_", &index))
        { // Default > recent default files
            SetIndexed(&settings.recent_defaults, index, value);
        }
        else if (Indexed(key, L"value_column_width_", &index))
        {
            SetIndexed(&settings.value_column_widths, index, number);
        }
        else if (Indexed(key, L"value_column_visible_", &index))
        {
            SetIndexed(&settings.value_column_visible, index, util::ParseBool(value), true);
        }
    }
    if (settings.always_run_as_trustedinstaller)
    {
        settings.always_run_as_system = false;
        settings.always_run_as_admin = false;
    }
    else if (settings.always_run_as_system)
    {
        settings.always_run_as_admin = false;
    }
    return settings;
}

std::wstring SerializeSettings(const Settings& settings)
{
    std::wstring content;
    for (const BoolField& field : kBoolFields)
    {
        Line(&content, field.key, settings.*(field.member) ? L"1" : L"0");
    }
    for (const IntField& field : kPositiveFields)
    {
        if (settings.*(field.member) > 0)
        {
            Line(&content, field.key, std::to_wstring(settings.*(field.member)));
        }
    }
    for (const TextField& field : kTextFields)
    {
        Line(&content, field.key, settings.*(field.member));
    }
    if (settings.window_width > 0 && settings.window_height > 0)
    {
        for (const IntField& field : kPlacementFields)
        {
            Line(&content, field.key, std::to_wstring(settings.*(field.member)));
        }
        Line(&content, L"window_maximized", settings.window_maximized ? L"1" : L"0");
    }
    Line(&content, L"save_tabs", settings.save_tabs ? L"1" : L"0");
    Line(&content, L"save_tab_types", std::to_wstring(settings.save_tab_kinds));
    Line(&content, L"font_use_default", settings.use_custom_font ? L"0" : L"1");
    for (size_t index = 0; index < settings.recent_traces.size(); ++index)
    {
        if (!settings.recent_traces[index].empty())
        {
            Line(&content, L"trace_recent_" + std::to_wstring(index), settings.recent_traces[index]);
        }
    }
    for (size_t index = 0; index < settings.recent_defaults.size(); ++index)
    {
        if (!settings.recent_defaults[index].empty())
        {
            Line(&content, L"default_recent_" + std::to_wstring(index), settings.recent_defaults[index]);
        }
    }
    const size_t columns = std::max(settings.value_column_widths.size(), settings.value_column_visible.size());
    for (size_t index = 0; index < columns; ++index)
    {
        const bool visible = index >= settings.value_column_visible.size() || settings.value_column_visible[index];
        Line(&content, L"value_column_width_" + std::to_wstring(index), std::to_wstring(index < settings.value_column_widths.size() ? settings.value_column_widths[index] : 0));
        Line(&content, L"value_column_visible_" + std::to_wstring(index), visible ? L"1" : L"0");
    }
    return content;
}

Settings DefaultOptions(const Settings& settings)
{
    Settings reset;
    reset.window_placement_present = settings.window_placement_present;
    reset.window_x = settings.window_x;
    reset.window_y = settings.window_y;
    reset.window_width = settings.window_width;
    reset.window_height = settings.window_height;
    reset.window_maximized = settings.window_maximized;
    reset.tree_width = settings.tree_width;
    reset.history_height = settings.history_height;
    reset.recent_traces = settings.recent_traces;
    reset.recent_defaults = settings.recent_defaults;
    reset.value_column_widths = settings.value_column_widths;
    reset.value_column_visible = settings.value_column_visible;
    if (!settings.use_custom_font)
    {
        reset.font_face = settings.font_face;
        reset.font_size = settings.font_size;
        reset.font_weight = settings.font_weight;
        reset.font_italic = settings.font_italic;
    }
    return reset;
}

bool LoadSettings(const std::wstring& path, Settings* settings, std::wstring* error)
{
    std::wstring content;
    std::wstring reason;
    if (settings && util::ReadStateFile(path, &content, &reason) && !content.empty())
    {
        Settings parsed = ParseSettings(content, *settings, &reason);
        if (reason.empty())
        {
            *settings = std::move(parsed);
            return true;
        }
    }
    if (error)
    {
        *error = std::move(reason);
    }
    return false;
}

bool SaveSettings(const std::wstring& path, const Settings& settings)
{
    return !path.empty() && win32ui::WriteTextFileAtomic(path.c_str(), SerializeSettings(settings)) == ERROR_SUCCESS;
}

}
