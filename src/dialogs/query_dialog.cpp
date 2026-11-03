// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/query_dialog.h"

#include "dialogs/query_prompts.h"
#include "win32/text_transform.h"

#include <win32ui/dialog/autocomplete.h>
#include <win32ui/dialog/combo_box.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/text_prompt.h>
#include <win32ui/dialog/tooltip.h>
#include <win32ui/menu/popup_menu.h>
#include <win32ui/system/date_time.h>
#include <win32ui/system/recent_list.h>
#include <win32ui/system/strings.h>
#include <win32ui/theme/theme.h>

#include <algorithm>
#include <cerrno>
#include <limits>
#include <vector>

#include <commctrl.h>
#include <windowsx.h>

#include "records/escaped_fields.h"
#include "registry/registry_store.h"
#include "resource.h"
#include "ui/autocomplete.h"
#include "ui/feedback.h"
#include "win32/file_text.h"
#include "win32/shell_paths.h"
#include "win32/translation.h"
#include "workspace/dialog_state.h"

namespace regkit
{

namespace
{

struct SearchDialogState
{
    HWND dialog = nullptr;
    SearchDialogResult* out = nullptr;
    SearchSources sources;
    bool recursive = true;
    win32ui::RecentList history{20};
    std::vector<std::wstring> root_names;
    std::vector<bool> root_selected;
    std::vector<DWORD> data_types;
};

HWND Item(const SearchDialogState* state, int id)
{
    return GetDlgItem(state->dialog, id);
}

constexpr std::pair<int, uint32_t> kAnomalyBoxes[] = {
    {IDC_FIND_NUL_NAMES, search::kAnomalyNulName},
    {IDC_FIND_ODD_NAMES, search::kAnomalyOddName},
    {IDC_FIND_INTEGER_SIZE, search::kAnomalyIntegerSize},
    {IDC_FIND_STRING_END, search::kAnomalyStringEnd},
    {IDC_FIND_MULTI_STRING, search::kAnomalyMultiString},
    {IDC_FIND_UNKNOWN_TYPE, search::kAnomalyUnknownType},
    {IDC_FIND_BROKEN_LINKS, search::kAnomalyBrokenLink},
    {IDC_FIND_VIRTUAL_STORE, search::kAnomalyVirtualStore},
};

std::wstring SearchHistoryPath()
{
    std::wstring folder = util::GetCacheFolder();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"search_history.txt");
}

std::vector<std::wstring> LoadSearchHistory(HWND owner)
{
    std::vector<std::wstring> items;
    std::wstring content;
    std::wstring error;
    const std::wstring path = SearchHistoryPath();
    if (!path.empty() && !util::ReadStateFile(path, &content, &error))
    {
        ui::ShowError(owner, std::wstring(util::Tr(L"RegKit couldn't read this file and didn't use it.")) + L"\n" + path + L"\n" + error);
    }
    else if (!path.empty())
    {
        for (const std::wstring_view line : win32ui::LineViews(content))
        {
            if (!line.empty())
            {
                items.emplace_back(line);
            }
        }
    }
    return items;
}

void SaveSearchHistory(HWND owner, const std::vector<std::wstring>& items)
{
    std::wstring content;
    for (const auto& item : items)
    {
        content.append(item).append(L"\n");
    }
    const std::wstring path = SearchHistoryPath();
    if (!content.empty())
    {
        ui::ReportSaveFailure(owner, path, path.empty() ? ERROR_PATH_NOT_FOUND : win32ui::WriteTextFileAtomic(path.c_str(), content));
    }
}

void UpdateScopeComboText(SearchDialogState* state)
{
    if (!state || !Item(state, IDC_FIND_SCOPE_ROOTS))
    {
        return;
    }
    size_t total = state->root_selected.size();
    size_t selected = 0;
    std::wstring first;
    for (size_t i = 0; i < state->root_selected.size(); ++i)
    {
        if (state->root_selected[i])
        {
            ++selected;
            if (first.empty() && i < state->root_names.size())
            {
                first = state->root_names[i];
            }
        }
    }
    std::wstring text;
    if (selected == 0)
    {
        text = util::Tr(L"No top level keys");
    }
    else if (selected == total)
    {
        text = util::Tr(L"All top level keys");
    }
    else if (selected == 1)
    {
        text = first;
    }
    else
    {
        text = util::Tr(L"Multiple keys");
    }
    win32ui::SetComboItems(Item(state, IDC_FIND_SCOPE_ROOTS), {text});
    win32ui::SetComboIndex(Item(state, IDC_FIND_SCOPE_ROOTS), 0);
}

void ShowRootSelectionMenu(HWND owner, SearchDialogState* state)
{
    if (!owner || !state || !Item(state, IDC_FIND_SCOPE_ROOTS))
    {
        return;
    }
    RECT rect = {};
    GetWindowRect(Item(state, IDC_FIND_SCOPE_ROOTS), &rect);
    HMENU menu = CreatePopupMenu();
    for (size_t i = 0; i < state->root_names.size(); ++i)
    {
        UINT flags = MF_STRING;
        if (i < state->root_selected.size() && state->root_selected[i])
        {
            flags |= MF_CHECKED;
        }
        AppendMenuW(menu, flags, static_cast<UINT>(1000 + i), state->root_names[i].c_str());
    }
    const int cmd = win32ui::TrackDropdownMenu(menu, rect, owner);
    DestroyMenu(menu);
    if (cmd >= 1000)
    {
        size_t index = static_cast<size_t>(cmd - 1000);
        if (index < state->root_selected.size())
        {
            state->root_selected[index] = !state->root_selected[index];
            UpdateScopeComboText(state);
        }
    }
}

// typed numbers, surrounding spaces and a 0x prefix allowed
bool ParseUint64(const std::wstring& text, uint64_t* out)
{
    const std::optional<uint64_t> value = win32ui::ParseUnsigned(text, 10, true);
    *out = value.value_or(0);
    return value.has_value();
}

bool GetDateTimeValue(HWND control, FILETIME* out)
{
    if (!control || !out)
    {
        return false;
    }
    SYSTEMTIME local = {};
    DWORD result = static_cast<DWORD>(SendMessageW(control, DTM_GETSYSTEMTIME, 0, reinterpret_cast<LPARAM>(&local)));
    return result == GDT_VALID && win32ui::LocalToFileTime(local, out);
}

void SetDateTimeValue(HWND control, const FILETIME& value)
{
    if (!control)
    {
        return;
    }
    SYSTEMTIME local = {};
    if (!win32ui::FileTimeToLocal(value, &local))
    {
        return;
    }
    SendMessageW(control, DTM_SETSYSTEMTIME, GDT_VALID, reinterpret_cast<LPARAM>(&local));
}

std::vector<std::wstring> SplitExcludePaths(const std::wstring& text)
{
    std::vector<std::wstring> items;
    std::wstring current;
    bool quoted = false;
    bool kept = false;
    auto finish = [&] {
        if (!kept)
        {
            current.erase(current.begin(), std::find_if(current.begin(), current.end(), [](wchar_t c) { return c != L' '; }));
            while (!current.empty() && current.back() == L' ')
            {
                current.pop_back();
            }
        }
        if (!current.empty())
        {
            items.push_back(std::move(current));
        }
        current.clear();
        kept = false;
    };
    for (size_t i = 0; i < text.size(); ++i)
    {
        const wchar_t ch = text[i];
        if (quoted)
        {
            if (ch != L'"')
            {
                current.push_back(ch);
            }
            else if (i + 1 < text.size() && text[i + 1] == L'"')
            {
                current.push_back(L'"');
                ++i;
            }
            else
            {
                quoted = false;
            }
        }
        else if (ch == L'\r' || ch == L'\n' || ch == L';' || ch == L',')
        {
            finish();
        }
        else if (ch == L'"' && !kept && std::all_of(current.begin(), current.end(), [](wchar_t c) { return c == L' '; }))
        {
            current.clear();
            quoted = kept = true;
        }
        else if (!kept || ch != L' ')
        {
            current.push_back(ch);
        }
    }
    finish();
    return items;
}

std::wstring JoinExcludePaths(const std::vector<std::wstring>& items)
{
    std::wstring out;
    for (const auto& item : items)
    {
        if (item.empty())
        {
            continue;
        }
        if (!out.empty())
        {
            out.append(L", ");
        }
        if (item.find_first_of(L",;\"\r\n") == std::wstring::npos && !item.starts_with(L' ') && !item.ends_with(L' '))
        {
            out.append(item);
            continue;
        }
        out.push_back(L'"');
        for (const wchar_t ch : item)
        {
            if (ch == L'"')
            {
                out.push_back(L'"');
            }
            out.push_back(ch);
        }
        out.push_back(L'"');
    }
    return out;
}

bool IsChecked(HWND button)
{
    return Button_GetCheck(button) == BST_CHECKED;
}

void SetChecked(HWND button, bool checked)
{
    Button_SetCheck(button, checked ? BST_CHECKED : BST_UNCHECKED);
}

void UpdateDialogEnableState(SearchDialogState* state)
{
    if (!state)
    {
        return;
    }

    bool scope_top = IsChecked(Item(state, IDC_FIND_SCOPE_TOP));
    bool scope_key = IsChecked(Item(state, IDC_FIND_SCOPE_KEY));
    bool standard_roots = IsChecked(Item(state, IDC_FIND_ROOT_KEYS));
    bool enable_roots = scope_top;
    EnableWindow(Item(state, IDC_FIND_SCOPE_ROOTS), enable_roots && standard_roots);
    EnableWindow(Item(state, IDC_FIND_SCOPE_EDIT), scope_key);
    EnableWindow(Item(state, IDC_FIND_BROWSE), scope_key);
    EnableWindow(Item(state, IDC_FIND_RECURSIVE), scope_key);

    bool search_data = IsChecked(Item(state, IDC_FIND_DATA));
    EnableWindow(Item(state, IDC_FIND_DATA_TYPES), search_data);
    EnableWindow(Item(state, IDC_FIND_MIN_SIZE), search_data);
    EnableWindow(Item(state, IDC_FIND_MAX_SIZE), search_data);

    bool min_checked = IsChecked(Item(state, IDC_FIND_MIN_SIZE));
    bool max_checked = IsChecked(Item(state, IDC_FIND_MAX_SIZE));
    EnableWindow(Item(state, IDC_FIND_MIN_SIZE_EDIT), search_data && min_checked);
    EnableWindow(Item(state, IDC_FIND_MAX_SIZE_EDIT), search_data && max_checked);

    bool exclude_checked = IsChecked(Item(state, IDC_FIND_EXCLUDE));
    EnableWindow(Item(state, IDC_FIND_EXCLUDE_EDIT), exclude_checked);

    bool limit_checked = IsChecked(Item(state, IDC_FIND_LIMIT));
    EnableWindow(Item(state, IDC_FIND_LIMIT_EDIT), limit_checked);
    EnableWindow(Item(state, IDC_FIND_EXCLUDE_BUTTON), exclude_checked);

    EnableWindow(Item(state, IDC_FIND_TRACE), state->sources.traces);
    EnableWindow(Item(state, IDC_FIND_DEFAULTS), state->sources.defaults);
    EnableWindow(Item(state, IDC_FIND_REGISTRY), state->sources.registry_root);
    EnableWindow(Item(state, IDC_FIND_OFFLINE), state->sources.offline);
    EnableWindow(Item(state, IDC_FIND_REG_FILES), state->sources.reg_files);
    EnableWindow(Item(state, IDC_FIND_REMOTE), state->sources.remote);
}

void LoadInitialState(SearchDialogState* state)
{
    state->history.Assign(LoadSearchHistory(state->dialog));
    win32ui::SetComboItems(Item(state, IDC_FIND_WHAT), state->history.items());
    if (state->out && (!state->out->criteria.query.empty() || state->out->criteria.anomalies))
    {
        SetWindowTextW(Item(state, IDC_FIND_WHAT), state->out->criteria.query.c_str());
    }
    else if (!state->history.items().empty())
    {
        SetWindowTextW(Item(state, IDC_FIND_WHAT), state->history.items().front().c_str());
    }

    auto roots = RegistryStore::DefaultRoots(state->sources.extra_hives);
    state->root_names.clear();
    state->root_selected.clear();
    state->root_names.reserve(roots.size());
    for (const auto& root : roots)
    {
        state->root_names.push_back(root.path_name);
    }
    state->root_selected.assign(state->root_names.size(), true);
    if (state->out && !state->out->root_paths.empty())
    {
        state->root_selected.assign(state->root_names.size(), false);
        for (size_t i = 0; i < state->root_names.size(); ++i)
        {
            for (const auto& path : state->out->root_paths)
            {
                if (util::EqualsInsensitive(state->root_names[i], path))
                {
                    state->root_selected[i] = true;
                    break;
                }
            }
        }
    }
    UpdateScopeComboText(state);

    if (state->out)
    {
        state->data_types = state->out->criteria.allowed_types;
        state->recursive = state->out->criteria.recursive;
    }
    const SearchDialogResult* initial = state->out;
    if (initial)
    {
        SetChecked(Item(state, IDC_FIND_KEYS), initial->criteria.search_keys);
        SetChecked(Item(state, IDC_FIND_VALUES), initial->criteria.search_values);
        SetChecked(Item(state, IDC_FIND_DATA), initial->criteria.search_data);
        SetChecked(Item(state, IDC_FIND_COMMENTS), initial->criteria.search_comments);
        SetChecked(Item(state, IDC_FIND_CASE), initial->criteria.match_case);
        SetChecked(Item(state, IDC_FIND_WHOLE), initial->criteria.match_whole);
        SetChecked(Item(state, IDC_FIND_REGEX), initial->criteria.use_regex);
        SetChecked(Item(state, IDC_FIND_SKIP_LINKS), initial->criteria.skip_links);
        for (const auto& [id, kind] : kAnomalyBoxes)
        {
            SetChecked(Item(state, id), state->sources.anomalies && (initial->criteria.anomalies & kind) != 0);
        }
        SetChecked(Item(state, IDC_FIND_MIN_SIZE), initial->criteria.use_min_size);
        SetChecked(Item(state, IDC_FIND_MAX_SIZE), initial->criteria.use_max_size);
        if (initial->criteria.use_min_size || initial->criteria.min_size)
        {
            SetWindowTextW(Item(state, IDC_FIND_MIN_SIZE_EDIT), std::to_wstring(initial->criteria.min_size).c_str());
        }
        if (initial->criteria.use_max_size || initial->criteria.max_size)
        {
            SetWindowTextW(Item(state, IDC_FIND_MAX_SIZE_EDIT), std::to_wstring(initial->criteria.max_size).c_str());
        }
        if (initial->criteria.use_modified_from)
        {
            SetDateTimeValue(Item(state, IDC_FIND_MODIFIED_FROM), initial->criteria.modified_from);
        }
        if (initial->criteria.use_modified_to)
        {
            SetDateTimeValue(Item(state, IDC_FIND_MODIFIED_TO), initial->criteria.modified_to);
        }
        SetChecked(Item(state, IDC_FIND_ROOT_KEYS), initial->search_standard_hives);
        SetChecked(Item(state, IDC_FIND_REGISTRY), initial->search_registry_root && state->sources.registry_root);
        SetChecked(Item(state, IDC_FIND_TRACE), initial->search_trace_values && state->sources.traces);
        SetChecked(Item(state, IDC_FIND_DEFAULTS), initial->search_default_data && state->sources.defaults);
        SetChecked(Item(state, IDC_FIND_OFFLINE), initial->search_offline_hives && state->sources.offline);
        SetChecked(Item(state, IDC_FIND_REG_FILES), initial->search_reg_files && state->sources.reg_files);
        SetChecked(Item(state, IDC_FIND_REMOTE), initial->search_remote_registry && state->sources.remote);
        bool scope_top = initial->scope == SearchScope::kEntireRegistry;
        SetChecked(Item(state, IDC_FIND_SCOPE_TOP), scope_top);
        SetChecked(Item(state, IDC_FIND_SCOPE_KEY), !scope_top);
        if (!initial->start_key.empty())
        {
            SetWindowTextW(Item(state, IDC_FIND_SCOPE_EDIT), initial->start_key.c_str());
        }
        bool new_tab = initial->result_mode == SearchResultMode::kNewTab;
        SetChecked(Item(state, IDC_FIND_RESULT_REUSE), !new_tab);
        SetChecked(Item(state, IDC_FIND_RESULT_NEW), new_tab);
        SetChecked(Item(state, IDC_FIND_RESULT_TAB), initial->open_in_new_tab);
        SetChecked(Item(state, IDC_FIND_EXCLUDE), initial->use_exclude);
        SetWindowTextW(Item(state, IDC_FIND_EXCLUDE_EDIT), JoinExcludePaths(initial->criteria.exclude_paths).c_str());
        const bool limited = initial->criteria.max_results > 0;
        SetChecked(Item(state, IDC_FIND_LIMIT), limited);
        SetWindowTextW(Item(state, IDC_FIND_LIMIT_EDIT), std::to_wstring(limited ? initial->criteria.max_results : 1000).c_str());
        SendMessageW(Item(state, IDC_FIND_LIMIT_EDIT), EM_SETSEL, 0, 0);
        EnableWindow(Item(state, IDC_FIND_LIMIT_EDIT), limited);
    }
    else
    {
        SetChecked(Item(state, IDC_FIND_KEYS), false);
        SetChecked(Item(state, IDC_FIND_VALUES), true);
        SetChecked(Item(state, IDC_FIND_DATA), true);
        SetChecked(Item(state, IDC_FIND_ROOT_KEYS), true);
        SetChecked(Item(state, IDC_FIND_REGISTRY), false);
        SetChecked(Item(state, IDC_FIND_TRACE), state->sources.traces);
        SetChecked(Item(state, IDC_FIND_SCOPE_TOP), true);
        SetChecked(Item(state, IDC_FIND_RESULT_REUSE), true);
        SetChecked(Item(state, IDC_FIND_LIMIT), true);
    }
    SetChecked(Item(state, IDC_FIND_RECURSIVE), state->recursive);
}

bool ReadSearchResult(HWND hwnd, SearchDialogState* state, SearchDialogResult* out)
{
    std::wstring query_text = util::WindowText(Item(state, IDC_FIND_WHAT));
    uint32_t anomalies = 0;
    for (const auto& [id, kind] : kAnomalyBoxes)
    {
        anomalies |= IsChecked(Item(state, id)) ? kind : 0u;
    }
    if (query_text.empty() && !anomalies)
    {
        ui::ShowWarning(hwnd, util::Tr(L"Enter a search term."));
        return false;
    }
    if (IsChecked(Item(state, IDC_FIND_REGEX)) && query_text.size() > search::regex::kMaxPatternLength)
    {
        ui::ShowWarning(hwnd, util::Tr(L"The regular expression is too long."));
        return false;
    }
    bool keys = IsChecked(Item(state, IDC_FIND_KEYS));
    bool values = IsChecked(Item(state, IDC_FIND_VALUES));
    bool data = IsChecked(Item(state, IDC_FIND_DATA));
    bool comments = IsChecked(Item(state, IDC_FIND_COMMENTS));
    bool default_data = state->sources.defaults && IsChecked(Item(state, IDC_FIND_DEFAULTS));
    if (!keys && !values && !data && !comments && !default_data && !anomalies)
    {
        ui::ShowWarning(hwnd, util::Tr(L"Select at least one search option."));
        return false;
    }
    bool standard_hives = IsChecked(Item(state, IDC_FIND_ROOT_KEYS));
    bool registry_root = state->sources.registry_root && IsChecked(Item(state, IDC_FIND_REGISTRY));
    bool trace_values = state->sources.traces && IsChecked(Item(state, IDC_FIND_TRACE));
    bool offline_hives = state->sources.offline && IsChecked(Item(state, IDC_FIND_OFFLINE));
    bool reg_files = state->sources.reg_files && IsChecked(Item(state, IDC_FIND_REG_FILES));
    bool remote_registry = state->sources.remote && IsChecked(Item(state, IDC_FIND_REMOTE));
    if (!standard_hives && !registry_root && !trace_values && !offline_hives && !reg_files && !remote_registry)
    {
        ui::ShowWarning(hwnd, util::Tr(L"Select at least one search source."));
        return false;
    }

    SearchDialogResult& result = *out;
    result.criteria.query = query_text;
    result.criteria.search_keys = keys;
    result.criteria.search_values = values;
    result.criteria.search_data = data;
    result.criteria.search_comments = comments;
    result.criteria.match_case = IsChecked(Item(state, IDC_FIND_CASE));
    result.criteria.match_whole = IsChecked(Item(state, IDC_FIND_WHOLE));
    result.criteria.use_regex = IsChecked(Item(state, IDC_FIND_REGEX));
    result.criteria.skip_links = IsChecked(Item(state, IDC_FIND_SKIP_LINKS));
    result.criteria.anomalies = anomalies;
    // retain unused data filters for next time, StartSearch applies them only with Search data
    result.criteria.allowed_types = state->data_types;
    result.criteria.use_min_size = IsChecked(Item(state, IDC_FIND_MIN_SIZE));
    result.criteria.use_max_size = IsChecked(Item(state, IDC_FIND_MAX_SIZE));
    if (!ParseUint64(util::WindowText(Item(state, IDC_FIND_MIN_SIZE_EDIT)), &result.criteria.min_size) && data && result.criteria.use_min_size)
    {
        ui::ShowWarning(hwnd, util::Tr(L"Enter a valid minimum data size."));
        return false;
    }
    if (!ParseUint64(util::WindowText(Item(state, IDC_FIND_MAX_SIZE_EDIT)), &result.criteria.max_size) && data && result.criteria.use_max_size)
    {
        ui::ShowWarning(hwnd, util::Tr(L"Enter a valid maximum data size."));
        return false;
    }
    FILETIME modified_from = {};
    FILETIME modified_to = {};
    bool has_modified_from = GetDateTimeValue(Item(state, IDC_FIND_MODIFIED_FROM), &modified_from);
    bool has_modified_to = GetDateTimeValue(Item(state, IDC_FIND_MODIFIED_TO), &modified_to);
    if (has_modified_from)
    {
        result.criteria.use_modified_from = true;
        result.criteria.modified_from = modified_from;
    }
    if (has_modified_to)
    {
        result.criteria.use_modified_to = true;
        result.criteria.modified_to = modified_to;
    }
    if (data && result.criteria.use_min_size && result.criteria.use_max_size &&
        result.criteria.min_size > result.criteria.max_size)
    {
        ui::ShowWarning(hwnd, util::Tr(L"Minimum data size can't exceed maximum data size."));
        return false;
    }
    if (has_modified_from && has_modified_to && CompareFileTime(&modified_from, &modified_to) > 0)
    {
        ui::ShowWarning(hwnd, util::Tr(L"Modified date range is invalid."));
        return false;
    }
    result.search_standard_hives = standard_hives;
    result.search_registry_root = registry_root;
    result.search_trace_values = trace_values;
    result.search_default_data = default_data;
    result.search_offline_hives = offline_hives;
    result.search_reg_files = reg_files;
    result.search_remote_registry = remote_registry;

    bool scope_top = IsChecked(Item(state, IDC_FIND_SCOPE_TOP));
    result.scope = scope_top ? SearchScope::kEntireRegistry : SearchScope::kCurrentKey;
    state->recursive = IsChecked(Item(state, IDC_FIND_RECURSIVE));
    result.criteria.recursive = scope_top ? true : state->recursive;
    result.result_mode = IsChecked(Item(state, IDC_FIND_RESULT_NEW)) ? SearchResultMode::kNewTab : SearchResultMode::kReuseTab;
    result.open_in_new_tab = IsChecked(Item(state, IDC_FIND_RESULT_TAB));
    if (IsChecked(Item(state, IDC_FIND_LIMIT)))
    {
        uint64_t limit = 0;
        if (!ParseUint64(util::WindowText(Item(state, IDC_FIND_LIMIT_EDIT)), &limit) || limit == 0)
        {
            ui::ShowWarning(hwnd, util::Tr(L"Enter a valid result limit."));
            return false;
        }
        result.criteria.max_results = limit;
    }
    else
    {
        result.criteria.max_results = 0;
    }

    result.use_exclude = IsChecked(Item(state, IDC_FIND_EXCLUDE));
    result.criteria.exclude_paths = SplitExcludePaths(util::WindowText(Item(state, IDC_FIND_EXCLUDE_EDIT)));

    result.root_paths.clear();
    result.start_key.clear();
    if (scope_top)
    {
        if (standard_hives)
        {
            for (size_t i = 0; i < state->root_selected.size(); ++i)
            {
                if (state->root_selected[i] && i < state->root_names.size())
                {
                    result.root_paths.push_back(state->root_names[i]);
                }
            }
            if (result.root_paths.empty())
            {
                ui::ShowWarning(hwnd, util::Tr(L"Select at least one top level key."));
                return false;
            }
        }
    }
    else
    {
        result.start_key = util::WindowText(Item(state, IDC_FIND_SCOPE_EDIT));
    }

    return true;
}

INT_PTR CALLBACK SearchDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<SearchDialogState*>(GetWindowLongPtrW(hwnd, DWLP_USER));
    if (msg == WM_INITDIALOG)
    {
        state = reinterpret_cast<SearchDialogState*>(lparam);
        SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        state->dialog = hwnd;
        const std::wstring time_format = win32ui::DateTimePattern();
        for (const int id : {IDC_FIND_MODIFIED_FROM, IDC_FIND_MODIFIED_TO})
        {
            SendDlgItemMessageW(hwnd, id, DTM_SETFORMAT, 0, reinterpret_cast<LPARAM>(time_format.c_str()));
            SendDlgItemMessageW(hwnd, id, DTM_SETSYSTEMTIME, GDT_NONE, 0);
        }
        win32ui::AttachAutoComplete(Item(state, IDC_FIND_SCOPE_EDIT), appearance::SuggestKeys);
        win32ui::AttachAutoComplete(Item(state, IDC_FIND_EXCLUDE_EDIT), [](const std::wstring& text) {
            const size_t split = text.find_first_not_of(L' ', std::min(text.size(), text.find_last_of(L",;") + 1));
            const std::wstring head = text.substr(0, std::min(split, text.size()));
            std::vector<std::wstring> items = appearance::SuggestKeys(text.substr(head.size()));
            for (auto& item : items)
            {
                item.insert(0, head);
            }
            return items;
        });
        win32ui::ApplyThemeToWindow(win32ui::AddTooltip(
            hwnd,
            Item(state, IDC_FIND_REGEX),
            util::Tr(L"PCRE syntax: ^ $ anchors, character classes, greedy, lazy (*?) and possessive (*+) quantifiers,\n"
                     L"(?<name>...) groups, lookaround (?=...) (?<=...), backreferences \\1 and Unicode classes \\p{L}, \\w, "
                     L"\\X.\n"
                     L"Matching is unicode aware and ignores case unless 'Match case' is set.")
        ));
        SetDlgItemTextW(hwnd, IDC_FIND_LIMIT_EDIT, L"1000");
        LoadInitialState(state);
        UpdateDialogEnableState(state);
        if (!state->sources.anomalies)
        {
            win32ui::CollapseSection(hwnd, IDC_FIND_NUL_NAMES);
        }
        SetFocus(Item(state, IDC_FIND_WHAT));
        SendDlgItemMessageW(hwnd, IDC_FIND_WHAT, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
        return FALSE;
    }
    if (!state)
    {
        return FALSE;
    }
    switch (msg)
    {
    case WM_COMMAND:
        {
            if (HIWORD(wparam) == CBN_DROPDOWN && LOWORD(wparam) == IDC_FIND_SCOPE_ROOTS)
            {
                ShowRootSelectionMenu(hwnd, state);
                SendMessageW(Item(state, IDC_FIND_SCOPE_ROOTS), CB_SHOWDROPDOWN, FALSE, 0);
                return TRUE;
            }
            if (HIWORD(wparam) == BN_CLICKED)
            {
                switch (LOWORD(wparam))
                {
                case IDC_FIND_SCOPE_TOP:
                case IDC_FIND_SCOPE_KEY:
                case IDC_FIND_DATA:
                case IDC_FIND_MIN_SIZE:
                case IDC_FIND_MAX_SIZE:
                case IDC_FIND_ROOT_KEYS:
                case IDC_FIND_REGISTRY:
                case IDC_FIND_TRACE:
                case IDC_FIND_OFFLINE:
                case IDC_FIND_REG_FILES:
                case IDC_FIND_REMOTE:
                case IDC_FIND_EXCLUDE:
                case IDC_FIND_LIMIT:
                    UpdateDialogEnableState(state);
                    break;
                default:
                    break;
                }
            }
            switch (LOWORD(wparam))
            {
            case IDC_FIND_BROWSE:
                {
                    std::wstring selected;
                    if (ShowBrowseKeyDialog(hwnd, &selected))
                    {
                        if (!selected.empty())
                        {
                            SetWindowTextW(Item(state, IDC_FIND_SCOPE_EDIT), selected.c_str());
                        }
                        SetChecked(Item(state, IDC_FIND_SCOPE_KEY), true);
                        SetChecked(Item(state, IDC_FIND_SCOPE_TOP), false);
                        UpdateDialogEnableState(state);
                    }
                    return TRUE;
                }
            case IDC_FIND_DATA_TYPES:
                query_prompts::ShowDataTypes(hwnd, &state->data_types);
                return TRUE;
            case IDC_FIND_EXCLUDE_BUTTON:
                {
                    std::wstring keys = util::JoinLines(SplitExcludePaths(util::WindowText(Item(state, IDC_FIND_EXCLUDE_EDIT))));
                    const win32ui::TextPrompt prompt = {
                        .title = util::Tr(L"Exclude Keys"),
                        .label = util::Tr(L"Each line should include one key."),
                        .multiline = true,
                        .browse = [](void*, HWND owner, std::wstring* line) { return ShowBrowseKeyDialog(owner, line); },
                    };
                    if (win32ui::PromptText(hwnd, prompt, &keys) == 0)
                    {
                        SetWindowTextW(Item(state, IDC_FIND_EXCLUDE_EDIT), JoinExcludePaths(util::SplitLines(keys)).c_str());
                    }
                    return TRUE;
                }
            case IDOK:
                {
                    SearchDialogResult result;
                    if (!ReadSearchResult(hwnd, state, &result))
                    {
                        return TRUE;
                    }
                    state->history.Add(result.criteria.query);
                    SaveSearchHistory(hwnd, state->history.items());

                    *state->out = std::move(result);
                    win32ui::EndDialogWindow(hwnd, IDOK);
                    return TRUE;
                }
            case IDCANCEL:
                win32ui::EndDialogWindow(hwnd, IDCANCEL);
                return TRUE;
            default:
                break;
            }
            break;
        }
    default:
        break;
    }
    return FALSE;
}

}

bool ShowBrowseKeyDialog(HWND owner, std::wstring* selected_path)
{
    return query_prompts::ShowRegistryKey(owner, selected_path);
}

bool ShowSearchDialog(HWND owner, SearchDialogResult* result, const SearchSources& available)
{
    SearchDialogState state;
    state.out = result;
    state.sources = available;
    return result && win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_FIND), SearchDialogProc, reinterpret_cast<LPARAM>(&state)) == IDOK;
}

void SearchDialogFields(workspace::DialogFields& fields, SearchDialogResult* result)
{
    search::Criteria& criteria = result->criteria;
    constexpr std::pair<const wchar_t*, bool search::Criteria::*> kCriteriaChecks[] = {
        {L"find_keys", &search::Criteria::search_keys},
        {L"find_values", &search::Criteria::search_values},
        {L"find_data", &search::Criteria::search_data},
        {L"find_comments", &search::Criteria::search_comments},
        {L"find_match_case", &search::Criteria::match_case},
        {L"find_match_whole", &search::Criteria::match_whole},
        {L"find_regex", &search::Criteria::use_regex},
        {L"find_recursive", &search::Criteria::recursive},
        {L"find_skip_links", &search::Criteria::skip_links},
        {L"find_use_min_size", &search::Criteria::use_min_size},
        {L"find_use_max_size", &search::Criteria::use_max_size},
        {L"find_use_modified_from", &search::Criteria::use_modified_from},
        {L"find_use_modified_to", &search::Criteria::use_modified_to},
    };
    constexpr std::pair<const wchar_t*, bool SearchDialogResult::*> kChecks[] = {
        {L"find_root_keys", &SearchDialogResult::search_standard_hives},
        {L"find_registry", &SearchDialogResult::search_registry_root},
        {L"find_trace_values", &SearchDialogResult::search_trace_values},
        {L"find_default_data", &SearchDialogResult::search_default_data},
        {L"find_offline_hives", &SearchDialogResult::search_offline_hives},
        {L"find_reg_files", &SearchDialogResult::search_reg_files},
        {L"find_network", &SearchDialogResult::search_remote_registry},
        {L"find_result_in_new_tab", &SearchDialogResult::open_in_new_tab},
        {L"find_use_exclude", &SearchDialogResult::use_exclude},
    };
    for (const auto& [key, member] : kCriteriaChecks)
    {
        fields.Field(key, &(criteria.*member));
    }
    for (const auto& [key, member] : kChecks)
    {
        fields.Field(key, &(result->*member));
    }
    fields.Field(L"find_min_size", &criteria.min_size);
    fields.Field(L"find_max_size", &criteria.max_size);
    fields.Field(L"find_max_results", &criteria.max_results);
    fields.Field(L"find_anomalies", &criteria.anomalies, UINT32_MAX);
    fields.Field(L"find_scope", &result->scope, SearchScope::kCurrentKey);
    fields.Field(L"find_result_mode", &result->result_mode, SearchResultMode::kNewTab);
    fields.Field(L"find_start_key", &result->start_key);
    for (const auto& [key, time] : {std::pair{L"find_modified_from", &criteria.modified_from}, std::pair{L"find_modified_to", &criteria.modified_to}})
    {
        uint64_t ticks = win32ui::FileTimeTicks(*time);
        fields.Field(key, &ticks);
        *time = win32ui::FileTimeFromTicks(ticks);
    }
    // lists are stored as one comma separated value
    std::wstring exclude = JoinExcludePaths(criteria.exclude_paths);
    std::wstring roots = JoinExcludePaths(result->root_paths);
    std::wstring types;
    for (const DWORD type : criteria.allowed_types)
    {
        types.append(types.empty() ? L"" : L",").append(std::to_wstring(type));
    }
    fields.Field(L"find_exclude", &exclude);
    fields.Field(L"find_root_paths", &roots);
    fields.Field(L"find_data_types", &types);
    criteria.exclude_paths = SplitExcludePaths(exclude);
    result->root_paths = SplitExcludePaths(roots);
    criteria.allowed_types.clear();
    for (const std::wstring& item : SplitExcludePaths(types))
    {
        uint64_t type = 0;
        if (record_fields::ParseUnsigned(item, MAXDWORD, &type))
        {
            criteria.allowed_types.push_back(static_cast<DWORD>(type));
        }
    }
}

}
