// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "win32/text_transform.h"

namespace regkit
{

using namespace window_detail;

std::wstring MainWindow::Impl::NormalizeRegistryPath(const std::wstring& input) const
{
    const std::wstring sid = util::GetCurrentUserSidString();
    std::wstring path = registry_path::Normalize(input, sid);
    auto strip_context = [&](const std::wstring& label) {
        if (label.empty())
        {
            return;
        }
        const std::wstring prefix = label + L"\\";
        if (util::StartsWithInsensitive(path, prefix))
        {
            path.erase(0, prefix.size());
        }
    };
    strip_context(TreeRootLabel());
    if (registry_mode_ == RegistryMode::kRemote)
    {
        strip_context(StripMachinePrefix(remote_machine_));
    }
    return registry_path::Normalize(path, sid);
}

std::wstring MainWindow::Impl::FormatRegistryPath(const std::wstring& path, RegistryPathFormat format) const
{
    const std::wstring normalized = NormalizeRegistryPath(path);
    if (normalized.empty())
    {
        return {};
    }
    std::wstring tree_root = registry_mode_ == RegistryMode::kLocal ? L"Computer" : TreeRootLabel();
    registry_path::Style style = registry_path::Style::kFull;
    switch (format)
    {
    case RegistryPathFormat::kAbbrev:
        style = registry_path::Style::kAbbreviated;
        break;
    case RegistryPathFormat::kRegEdit:
        style = registry_path::Style::kRegEditAddress;
        break;
    case RegistryPathFormat::kRegFile:
        style = registry_path::Style::kRegFileHeader;
        break;
    case RegistryPathFormat::kPowerShellDrive:
        style = registry_path::Style::kPowerShellDrive;
        break;
    case RegistryPathFormat::kPowerShellProvider:
        style = registry_path::Style::kPowerShellProvider;
        break;
    case RegistryPathFormat::kEscaped:
        style = registry_path::Style::kEscaped;
        break;
    case RegistryPathFormat::kFull:
        break;
    }
    return registry_path::Format(normalized, style, tree_root);
}
bool MainWindow::Impl::FindNearestExistingPath(const std::wstring& path, std::wstring* nearest_path) const
{
    return changes::FindNearestExistingPath(
        path,
        [this](const std::wstring& candidate) {
            RegistryNode node;
            KeyInfo info = {};
            return ResolvePathToNode(candidate, &node) && RegistryStore::QueryKeyInfo(node, &info);
        },
        nearest_path
    );
}

bool MainWindow::Impl::CreateRegistryPath(const std::wstring& path)
{
    RegistryNode node;
    if (!ResolvePathToNode(path, &node))
    {
        return false;
    }
    if (node.subkey.empty())
    {
        return true;
    }
    const std::vector<std::wstring> parts = registry_path::Split(node.subkey);
    RegistryNode current = node;
    current.subkey.clear();
    bool created = false;
    for (const auto& part : parts)
    {
        RegistryNode child = registry_path::ChildNode(current, part);
        KeyInfo info = {};
        if (!RegistryStore::QueryKeyInfo(child, &info))
        {
            if (!RegistryStore::CreateKey(current, part))
            {
                return false;
            }
            created = true;
        }
        current = std::move(child);
    }
    if (created)
    {
        MarkOfflineDirty();
    }
    return true;
}

void MainWindow::Impl::SetStatusMessage(const std::wstring& text)
{
    status_message_ = text;
    UpdateStatus();
    if (hwnd_)
    {
        KillTimer(hwnd_, kStatusMessageTimerId);
        if (!text.empty())
        {
            SetTimer(hwnd_, kStatusMessageTimerId, 8000, nullptr);
        }
    }
}

void MainWindow::Impl::UpdateStatus()
{
    if (!status_bar_)
    {
        return;
    }
    RECT rc = {};
    GetClientRect(status_bar_, &rc);
    int total_width = rc.right - rc.left;
    if (total_width < 0)
    {
        total_width = 0;
    }
    LONG_PTR sb_style = GetWindowLongPtrW(status_bar_, GWL_STYLE);
    if (sb_style & SBARS_SIZEGRIP)
    {
        int grip = GetSystemMetrics(SM_CXVSCROLL);
        total_width = std::max(total_width - grip, 0);
    }
    auto measure_text = [&](HDC hdc, const std::wstring& text) -> int {
        if (!hdc || text.empty())
        {
            return 0;
        }
        SIZE size = {};
        GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &size);
        return size.cx + 20;
    };
    if (IsSearchTabSelected())
    {
        bool compare_selected = IsCompareTabSelected();
        int sel = TabCtrl_GetCurSel(tab_);
        int tab_index = SearchIndexFromTab(sel);
        size_t count = 0;
        if (tab_index >= 0 && static_cast<size_t>(tab_index) < search_tabs_.size())
        {
            count = SearchRowCount(tab_index);
        }
        unsigned long long count_value = static_cast<unsigned long long>(count);
        wchar_t buffer[256] = {};
        if (compare_selected)
        {
            swprintf_s(buffer, L"Results: %llu", count_value);
        }
        else if (search_running_)
        {
            uint64_t searched = search_progress_searched_.load();
            if (searched > 0)
            {
                swprintf_s(buffer, L"Searching... Results: ~%llu | Scanned: %llu", count_value, searched);
            }
            else
            {
                swprintf_s(buffer, L"Searching... Results: ~%llu", count_value);
            }
        }
        else if (search_duration_valid_ && search_duration_ms_ > 0)
        {
            double seconds = static_cast<double>(search_duration_ms_) / 1000.0;
            swprintf_s(buffer, L"Results: %llu (%.2fs)", count_value, seconds);
        }
        else
        {
            swprintf_s(buffer, L"Results: %llu", count_value);
        }
        int part = total_width;
        SendMessageW(status_bar_, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(&part));
        SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(status_message_.empty() ? buffer : status_message_.c_str()));
        return;
    }
    if (IsRegFileTabSelected())
    {
        int sel = TabCtrl_GetCurSel(tab_);
        if (IsRegFileTabIndex(sel) && static_cast<size_t>(sel) < tabs_.size())
        {
            const TabEntry& entry = tabs_[static_cast<size_t>(sel)];
            if (entry.reg_file_loading)
            {
                std::wstring label = entry.reg_file_label.empty() ? L"registry file" : entry.reg_file_label;
                std::wstring text = L"Loading " + label + L"...";
                int part = total_width;
                SendMessageW(status_bar_, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(&part));
                SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(text.c_str()));
                return;
            }
        }
    }

    int selected = ListView_GetSelectedCount(browse_.values().hwnd());
    wchar_t buffer[256] = {};
    std::wstring keys_text;
    std::wstring values_text;
    std::wstring selected_text;
    std::wstring path_text;
    if (!status_message_.empty())
    {
        path_text = status_message_;
    }
    else if (browse_.current_node())
    {
        path_text = registry_path::Build(*browse_.current_node());
    }
    swprintf_s(buffer, L"Keys: %d", current_key_count_);
    keys_text = buffer;
    swprintf_s(buffer, L"Values: %d", current_value_count_);
    values_text = buffer;
    swprintf_s(buffer, L"Selected: %d", selected);
    selected_text = buffer;

    HDC hdc = GetDC(status_bar_);
    HFONT old_font = nullptr;
    if (hdc && ui_font_)
    {
        old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
    }
    int values_width = measure_text(hdc, values_text);
    int selected_width = measure_text(hdc, selected_text);
    int keys_width = measure_text(hdc, keys_text);
    if (old_font)
    {
        SelectObject(hdc, old_font);
    }
    if (hdc)
    {
        ReleaseDC(status_bar_, hdc);
    }

    int part3 = total_width;
    int part2 = std::max(part3 - keys_width, 0);
    int part1 = std::max(part2 - selected_width, 0);
    int part0 = std::max(part1 - values_width, 0);
    int parts[4] = {part0, part1, part2, part3};
    SendMessageW(status_bar_, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));
    SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(path_text.c_str()));
    SendMessageW(status_bar_, SB_SETTEXTW, 1, reinterpret_cast<LPARAM>(values_text.c_str()));
    SendMessageW(status_bar_, SB_SETTEXTW, 2, reinterpret_cast<LPARAM>(selected_text.c_str()));
    SendMessageW(status_bar_, SB_SETTEXTW, 3, reinterpret_cast<LPARAM>(keys_text.c_str()));
}

bool MainWindow::Impl::IsSearchTabSelected() const
{
    if (!tab_)
    {
        return false;
    }
    int index = TabCtrl_GetCurSel(tab_);
    return IsSearchTabIndex(index);
}

bool MainWindow::Impl::IsRegFileTabSelected() const
{
    if (!tab_)
    {
        return false;
    }
    int index = TabCtrl_GetCurSel(tab_);
    return IsRegFileTabIndex(index);
}

bool MainWindow::Impl::IsCompareTabSelected() const
{
    if (!tab_)
    {
        return false;
    }
    int index = TabCtrl_GetCurSel(tab_);
    if (!IsSearchTabIndex(index))
    {
        return false;
    }
    int search_index = SearchIndexFromTab(index);
    if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size())
    {
        return false;
    }
    return search_tabs_[static_cast<size_t>(search_index)].is_compare;
}

bool MainWindow::Impl::IsCompareResultColumnAvailable() const
{
    if (!tab_)
    {
        return false;
    }
    const int search_index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
    return search_index >= 0 && static_cast<size_t>(search_index) < search_tabs_.size() &&
           search_tabs_[static_cast<size_t>(search_index)].is_compare &&
           search_tabs_[static_cast<size_t>(search_index)].compare_filter == search::compare::RowFilter::kAll;
}

bool MainWindow::Impl::IsSearchTabIndex(int index) const
{
    if (index < 0)
    {
        return false;
    }
    if (static_cast<size_t>(index) >= tabs_.size())
    {
        return false;
    }
    return tabs_[static_cast<size_t>(index)].kind == TabEntry::Kind::kSearch;
}

bool MainWindow::Impl::IsRegFileTabIndex(int index) const
{
    if (index < 0)
    {
        return false;
    }
    if (static_cast<size_t>(index) >= tabs_.size())
    {
        return false;
    }
    return tabs_[static_cast<size_t>(index)].kind == TabEntry::Kind::kRegFile;
}

int MainWindow::Impl::SearchIndexFromTab(int index) const
{
    if (!IsSearchTabIndex(index))
    {
        return -1;
    }
    return tabs_[static_cast<size_t>(index)].search_index;
}

bool MainWindow::Impl::IsLocalRegistryTabIndex(int index) const
{
    if (index < 0 || static_cast<size_t>(index) >= tabs_.size())
    {
        return false;
    }
    const TabEntry& entry = tabs_[static_cast<size_t>(index)];
    return entry.kind == TabEntry::Kind::kRegistry && entry.registry_mode == RegistryMode::kLocal;
}

int MainWindow::Impl::FindLocalRegistryTabIndex() const
{
    for (size_t i = 0; i < tabs_.size(); ++i)
    {
        if (IsLocalRegistryTabIndex(static_cast<int>(i)))
        {
            return static_cast<int>(i);
        }
    }
    return -1;
}

int MainWindow::Impl::FindFirstRegistryTabIndex() const
{
    for (size_t i = 0; i < tabs_.size(); ++i)
    {
        if (tabs_[i].kind == TabEntry::Kind::kRegistry)
        {
            return static_cast<int>(i);
        }
    }
    return -1;
}

void MainWindow::Impl::SyncRegFileTabSelection()
{
    if (!tab_)
    {
        return;
    }
    int index = TabCtrl_GetCurSel(tab_);
    if (!IsRegFileTabIndex(index))
    {
        return;
    }
    if (static_cast<size_t>(index) >= tabs_.size())
    {
        return;
    }
    TabEntry& entry = tabs_[static_cast<size_t>(index)];
    if (entry.reg_file_roots.empty() && !entry.reg_file_path.empty() && !entry.reg_file_loading)
    {
        if (entry.reg_file_session_key.empty())
        {
            entry.reg_file_session_key =
                ToLower(entry.reg_file_path) + L"|" + std::to_wstring(++reg_file_session_serial_);
        }
        entry.reg_file_loading = true;
        StartRegFileParse(entry.reg_file_path, entry.reg_file_session_key);
    }
    registry_mode_ = RegistryMode::kLocal;
    std::vector<RegistryRootEntry> roots;
    roots.reserve(entry.reg_file_roots.size());
    for (const auto& root : entry.reg_file_roots)
    {
        if (!root.root)
        {
            continue;
        }
        RegistryRootEntry reg_root;
        reg_root.root = root.root;
        reg_root.display_name = root.name;
        reg_root.path_name = root.name;
        reg_root.subkey_prefix = L"";
        reg_root.group = RegistryRootGroup::kStandard;
        roots.push_back(std::move(reg_root));
    }
    ApplyRegistryRoots(roots);
    RestoreRegistryTabState(index);
    if (!pending_compare_key_path_.empty() && !entry.reg_file_roots.empty())
    {
        const std::wstring path = std::move(pending_compare_key_path_);
        const std::wstring value_name = std::move(pending_compare_value_name_);
        pending_compare_key_path_.clear();
        pending_compare_value_name_.clear();
        SelectTreePath(path);
        if (!value_name.empty())
        {
            SelectValueWhenReady(value_name);
        }
    }
}

void MainWindow::Impl::UpdateSearchResultsView()
{
    if (!search_results_list_)
    {
        return;
    }
    int sel = TabCtrl_GetCurSel(tab_);
    if (!IsSearchTabIndex(sel))
    {
        return;
    }
    int search_index = SearchIndexFromTab(sel);
    if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size())
    {
        return;
    }
    EnsureSearchTabResultsLoaded(search_index);
    bool force_redraw = (search_results_view_tab_index_ != sel);
    search_results_view_tab_index_ = sel;
    auto& tab = search_tabs_[static_cast<size_t>(search_index)];
    bool compare = tab.is_compare;
    const bool show_result = compare && tab.compare_filter == search::compare::RowFilter::kAll;
    if (compare != compare_columns_active_ || show_result != compare_result_column_active_)
    {
        ApplySearchColumns(compare);
        force_redraw = true;
    }
    else if (compare && force_redraw)
    {
        RefreshCompareColumnTitles();
    }
    int max_sort_col = compare ? (show_result ? 4 : 3) : 5;
    if (tab.sort_column > max_sort_col)
    {
        tab.sort_column = -1;
    }
    appearance::UpdateListViewSort(search_results_list_, tab.sort_column, tab.sort_ascending);
    size_t count = compare ? tab.compare_rows.size() : tab.results.size();
    size_t old_count = tab.last_ui_count;
    if (force_redraw || count != old_count)
    {
        ListView_SetItemCountEx(search_results_list_, static_cast<int>(count), LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
        if (force_redraw || count < old_count)
        {
            RedrawWindow(search_results_list_, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
        }
        else if (count > old_count)
        {
            int first = static_cast<int>(old_count);
            int last = static_cast<int>(count - 1);
            ListView_RedrawItems(search_results_list_, first, last);
        }
        tab.last_ui_count = count;
    }
}

void MainWindow::Impl::StartSearch(const SearchDialogResult& options)
{
    if (options.criteria.query.empty())
    {
        ui::ShowWarning(hwnd_, L"Enter text to find.");
        return;
    }

    search::TextOptions match_options;
    match_options.query = options.criteria.query;
    match_options.match_case = options.criteria.match_case;
    match_options.match_whole = options.criteria.match_whole;
    match_options.use_regex = options.criteria.use_regex;
    auto matcher = std::make_shared<const search::Matcher>(match_options);
    if (!matcher->valid())
    {
        ui::ShowError(hwnd_, search::regex::ErrorText(matcher->error()));
        return;
    }

    bool want_registry = options.search_standard_hives || options.search_registry_root ||
                         options.search_offline_hives || options.search_reg_files || options.search_remote_registry;
    bool want_trace = options.search_trace_values && !active_traces_.empty();
    std::wstring registry_scope_path;
    std::wstring scope_path;
    if (options.scope == SearchScope::kCurrentKey)
    {
        if (!options.start_key.empty())
        {
            registry_scope_path = options.start_key;
            scope_path = NormalizeRegistryPath(options.start_key);
        }
        else if (browse_.current_node())
        {
            registry_scope_path = registry_path::Build(*browse_.current_node());
            scope_path = NormalizeRegistryPath(registry_scope_path);
        }
        else
        {
            ui::ShowError(hwnd_, L"Select a starting key first.");
            return;
        }
    }

    std::vector<search::StartNode> start_nodes;
    std::vector<search::Source> sources(1);
    bool remote_nodes = false;
    auto source_index = [&](search::Source::Kind kind, const std::wstring& name) -> uint16_t {
        const search::Source wanted{kind, name};
        for (size_t i = 0; i < sources.size(); ++i)
        {
            if (search::SameSource(sources[i], wanted))
            {
                return static_cast<uint16_t>(i);
            }
        }
        sources.push_back(wanted);
        return static_cast<uint16_t>(sources.size() - 1);
    };
    if (want_registry)
    {
        if (options.scope == SearchScope::kCurrentKey)
        {
            const search::Source current = CurrentTabSource();
            const uint16_t source = source_index(current.kind, current.name);
            if (!registry_scope_path.empty())
            {
                RegistryNode node;
                if (ResolvePathToNode(registry_scope_path, &node))
                {
                    start_nodes.push_back({node, source});
                }
                else
                {
                    std::wstring normalized = NormalizeRegistryPath(registry_scope_path);
                    if (!normalized.empty() && ResolvePathToNode(normalized, &node))
                    {
                        start_nodes.push_back({node, source});
                    }
                    else
                    {
                        ui::ShowError(hwnd_, L"Starting key path wasn't found.");
                        return;
                    }
                }
            }
            else if (browse_.current_node())
            {
                start_nodes.push_back({*browse_.current_node(), source});
            }
            else
            {
                ui::ShowError(hwnd_, L"Select a starting key first.");
                return;
            }
        }
        else
        {
            std::unordered_set<std::wstring> seen;
            auto add_root = [&](const RegistryRootEntry& entry, uint16_t source) {
                std::wstring key = ToLower(entry.path_name.empty() ? entry.display_name : entry.path_name);
                if (key.empty())
                {
                    return;
                }
                key.append(L"|").append(std::to_wstring(reinterpret_cast<uintptr_t>(entry.root)));
                if (!seen.insert(key).second)
                {
                    return;
                }
                RegistryNode node;
                node.root = entry.root;
                node.root_name = entry.path_name;
                node.subkey = entry.subkey_prefix;
                start_nodes.push_back({std::move(node), source});
            };

            std::vector<RegistryRootEntry> local_roots = RegistryStore::DefaultRoots(show_extra_hives_);
            AppendRealRegistryRoot(&local_roots);
            if (options.search_standard_hives)
            {
                for (const auto& path : options.root_paths)
                {
                    for (const auto& root : local_roots)
                    {
                        if (util::EqualsInsensitive(root.path_name, path) ||
                            util::EqualsInsensitive(root.display_name, path))
                        {
                            add_root(root, 0);
                            break;
                        }
                    }
                }
                if (start_nodes.empty())
                {
                    for (const auto& root : local_roots)
                    {
                        if (root.group == RegistryRootGroup::kStandard)
                        {
                            add_root(root, 0);
                        }
                    }
                }
            }
            if (options.search_registry_root)
            {
                for (const auto& root : local_roots)
                {
                    if (root.group == RegistryRootGroup::kReal)
                    {
                        add_root(root, 0);
                        break;
                    }
                }
            }
            if (options.search_offline_hives && !offline_roots_.empty())
            {
                std::wstring offline_path;
                for (const auto& tab : tabs_)
                {
                    if (tab.kind == TabEntry::Kind::kRegistry && tab.registry_mode == RegistryMode::kOffline)
                    {
                        offline_path = tab.offline_path;
                        break;
                    }
                }
                const uint16_t source = source_index(search::Source::Kind::kOffline, offline_path);
                for (size_t i = 0; i < offline_roots_.size(); ++i)
                {
                    RegistryRootEntry entry;
                    entry.root = offline_roots_[i];
                    entry.display_name = i < offline_root_labels_.size() ? offline_root_labels_[i] : L"OfflineHive";
                    entry.path_name = offline_root_name_ + L"\\" + entry.display_name;
                    add_root(entry, source);
                }
            }
            if (options.search_reg_files)
            {
                for (const auto& tab : tabs_)
                {
                    if (tab.kind != TabEntry::Kind::kRegFile || tab.reg_file_roots.empty())
                    {
                        continue;
                    }
                    const uint16_t source = source_index(search::Source::Kind::kRegFile, tab.reg_file_path);
                    for (const auto& root : tab.reg_file_roots)
                    {
                        if (!root.root)
                        {
                            continue;
                        }
                        RegistryRootEntry entry;
                        entry.root = root.root;
                        entry.display_name = root.name;
                        entry.path_name = root.name;
                        add_root(entry, source);
                    }
                }
            }
            if (options.search_remote_registry && remote_hklm_)
            {
                const std::wstring prefix = remote_machine_ + L"\\";
                const uint16_t source = source_index(search::Source::Kind::kRemote, remote_machine_);
                remote_nodes = true;
                add_root({remote_hklm_, L"HKEY_LOCAL_MACHINE", prefix + L"HKEY_LOCAL_MACHINE", L""}, source);
                if (remote_hku_)
                {
                    add_root({remote_hku_, L"HKEY_USERS", prefix + L"HKEY_USERS", L""}, source);
                }
            }
        }
    }

    if (want_registry && start_nodes.empty())
    {
        ui::ShowError(hwnd_, L"No keys to search in the selected sources.");
        return;
    }
    if (!want_registry && !want_trace)
    {
        return;
    }
    if (!tab_)
    {
        return;
    }

    CancelSearch();

    search::Criteria criteria = options.criteria;
    criteria.matcher = matcher;
    criteria.start_nodes = start_nodes;
    if (criteria.search_comments)
    {
        const auto comments = std::make_shared<const std::pair<changes::ValueComments, changes::ValueComments>>(value_comments_, default_comments_);
        criteria.comment_text = [comments](const std::wstring& path, const std::wstring* name, DWORD type, DWORD size) {
            return changes::ResolveComment(comments->first, comments->second, {path, name ? *name : std::wstring(), type, size, !name}).text;
        };
    }
    if (options.search_default_data && !active_defaults_.empty())
    {
        criteria.default_text = [defaults = active_defaults_](const std::wstring& path, const std::wstring& name) {
            thread_local std::wstring last_path;
            thread_local std::wstring key_lower;
            if (path != last_path)
            {
                last_path = path;
                const std::wstring normalized = NormalizeTraceKeyPathBasic(path);
                key_lower = ToLower(normalized.empty() ? path : normalized);
            }
            const std::wstring value_lower = ToLower(name);
            std::wstring text;
            for (const auto& set : defaults)
            {
                if (!set.data || !set.selection || !trace::IncludesKey(*set.selection, key_lower) || !trace::IncludesValue(*set.selection, key_lower, value_lower))
                {
                    continue;
                }
                std::shared_lock<std::shared_mutex> lock(*set.data->mutex);
                const auto key = set.data->values_by_key.find(key_lower);
                if (key == set.data->values_by_key.end())
                {
                    continue;
                }
                const auto value = key->second.values.find(value_lower);
                if (value != key->second.values.end())
                {
                    text.append(value->second.data).push_back(L'\n');
                }
            }
            return text;
        };
    }

    std::wstring label = L"Find";
    if (!criteria.query.empty())
    {
        label = L"Find: " + criteria.query;
        constexpr size_t kMaxLabel = 48;
        if (label.size() > kMaxLabel)
        {
            label.resize(kMaxLabel - 3);
            label.append(L"...");
        }
    }

    int tab_index = -1;
    int search_index = -1;
    bool reuse_tab = options.result_mode == SearchResultMode::kReuseTab;
    if (reuse_tab)
    {
        int sel = TabCtrl_GetCurSel(tab_);
        int candidate = IsSearchTabIndex(sel) ? sel : active_search_tab_index_;
        if (IsSearchTabIndex(candidate))
        {
            int index = SearchIndexFromTab(candidate);
            if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size() &&
                !search_tabs_[static_cast<size_t>(index)].is_compare)
            {
                tab_index = candidate;
                search_index = index;
            }
        }
    }

    if (search_index >= 0)
    {
        SearchTab& tab = search_tabs_[static_cast<size_t>(search_index)];
        tab.label = label;
        tab.results.clear();
        tab.last_ui_count = 0;
        tab.is_compare = false;
        tab.sort_dirty = false;
        tab.sources = sources;
        tab.open_in_new_tab = options.open_in_new_tab;
        TCITEMW item = {};
        item.mask = TCIF_TEXT;
        item.pszText = const_cast<wchar_t*>(tab.label.c_str());
        TabCtrl_SetItem(tab_, tab_index, &item);
    }
    else
    {
        SearchTab tab;
        tab.label = label;
        tab.is_compare = false;
        tab.sources = sources;
        tab.open_in_new_tab = options.open_in_new_tab;
        search_tabs_.push_back(std::move(tab));
        search_index = static_cast<int>(search_tabs_.size() - 1);
        TCITEMW item = {};
        item.mask = TCIF_TEXT;
        item.pszText = const_cast<wchar_t*>(search_tabs_.back().label.c_str());
        tab_index = TabCtrl_GetItemCount(tab_);
        TabCtrl_InsertItem(tab_, tab_index, &item);
        tabs_.push_back({TabEntry::Kind::kSearch, search_index});
    }

    UpdateTabWidth();
    SelectTabIndex(tab_index);
    active_search_tab_index_ = tab_index;
    search_results_view_tab_index_ = -1;
    search_progress_searched_.store(0);
    search_progress_posted_.store(false);
    search_posted_.store(false);
    {
        std::lock_guard<std::mutex> lock(search_mutex_);
        search_pending_batches_.clear();
        search_pending_rows_ = 0;
        search_producer_done_ = false;
    }
    search_last_refresh_tick_ = 0;
    search_start_tick_ = GetTickCount64();
    search_duration_ms_ = 0;
    search_duration_valid_ = false;
    search_running_ = true;

    if (search_progress_)
    {
        SendMessageW(search_progress_, PBM_SETMARQUEE, TRUE, 30);
    }

    ApplyViewVisibility();
    UpdateSearchResultsView();
    UpdateStatus();

    std::vector<ActiveTrace> traces = active_traces_;
    if (options.scope == SearchScope::kEntireRegistry)
    {
        criteria.provider = remote_nodes ? search::Provider::kRemote : search::Provider::kLocal;
    }
    else
    {
        criteria.provider = registry_mode_ == RegistryMode::kRemote    ? search::Provider::kRemote
                            : registry_mode_ == RegistryMode::kOffline ? search::Provider::kOffline
                                                                       : search::Provider::kLocal;
    }
    std::vector<std::wstring> exclude_paths = criteria.exclude_paths;
    std::wstring scope_lower = ToLower(scope_path);
    bool scope_recursive = criteria.recursive;
    bool trace_enabled = want_trace;
    bool registry_enabled = want_registry && !criteria.start_nodes.empty();

    const uint64_t generation = search_session_.Start([this, criteria, traces, exclude_paths, scope_lower, scope_recursive, trace_enabled, registry_enabled, matcher](uint64_t generation, std::atomic_bool& cancel) mutable {
        auto should_stop = [&]() { return cancel.load(); };

        auto publish_batch = [&](search::ResultBatch&& rows) -> bool {
            if (rows.empty())
            {
                return !cancel.load();
            }
            PendingSearchBatch pending;
            pending.generation = generation;
            pending.rows = std::move(rows);
            const size_t added = pending.rows.size();
            {
                std::unique_lock<std::mutex> lock(search_mutex_);

                search_queue_space_.wait(
                    lock,
                    [&]() { return cancel.load() || search_pending_rows_ < kSearchPendingRowLimit; }
                );
                if (cancel.load())
                {
                    return false;
                }
                search_pending_batches_.push_back(std::move(pending));
                search_pending_rows_ += added;
            }
            if (!search_posted_.exchange(true))
            {
                if (!PostMessageW(hwnd_, frame::message_id::kSearchResults, static_cast<WPARAM>(generation), 0))
                {
                    search_posted_.store(false);
                }
            }
            return !cancel.load();
        };

        search::ResultBatch trace_batch;
        trace_batch.reserve(kSearchQueueBatch);
        auto queue_result = [&](search::Result&& result) {
            trace_batch.push_back(std::move(result));
            if (trace_batch.size() >= kSearchQueueBatch)
            {
                publish_batch(std::move(trace_batch));
                trace_batch.clear();
                trace_batch.reserve(kSearchQueueBatch);
            }
        };

        auto flush = [&]() {
            if (!trace_batch.empty())
            {
                publish_batch(std::move(trace_batch));
                trace_batch.clear();
                trace_batch.reserve(kSearchQueueBatch);
            }
        };

        auto is_excluded = [&](const std::wstring& path) { return search::IsExcludedPath(path, exclude_paths); };

        auto key_in_scope = [&](const std::wstring& key_lower) {
            if (scope_lower.empty())
            {
                return true;
            }
            if (key_lower == scope_lower)
            {
                return true;
            }
            if (!scope_recursive)
            {
                return false;
            }
            if (key_lower.size() <= scope_lower.size())
            {
                return false;
            }
            if (key_lower.compare(0, scope_lower.size(), scope_lower) != 0)
            {
                return false;
            }
            return key_lower[scope_lower.size()] == L'\\';
        };

        if (trace_enabled)
        {
            for (const auto& trace : traces)
            {
                if (should_stop())
                {
                    break;
                }
                if (!trace.data)
                {
                    continue;
                }
                std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
                for (const auto& key_path : trace.data->key_paths)
                {
                    if (should_stop())
                    {
                        break;
                    }
                    if (key_path.empty())
                    {
                        continue;
                    }
                    if (is_excluded(key_path))
                    {
                        continue;
                    }
                    std::wstring key_lower = ToLower(key_path);
                    if (!trace.selection || !trace::IncludesKey(*trace.selection, key_lower))
                    {
                        continue;
                    }
                    if (!key_in_scope(key_lower))
                    {
                        continue;
                    }
                    std::wstring key_name = registry_path::Leaf(key_path);

                    if (criteria.search_keys)
                    {
                        const search::Match match = matcher->Find(key_name);
                        if (match.matched)
                        {
                            search::Result result;
                            result.key_path = key_path;
                            result.kind = search::ResultKind::kTraceKey;
                            result.data_state = search::DataState::kNotApplicable;
                            const size_t path_start =
                                key_path.size() >= key_name.size() ? key_path.size() - key_name.size() : 0;
                            result.match_field = search::MatchField::kPath;
                            result.match_start = static_cast<uint32_t>(path_start + match.start);
                            result.match_length = static_cast<uint32_t>(match.length);
                            queue_result(std::move(result));
                        }
                    }

                    if (criteria.search_values)
                    {
                        auto it = trace.data->values_by_key.find(key_lower);
                        if (it != trace.data->values_by_key.end())
                        {
                            for (const auto& value_name : it->second.values_display)
                            {
                                if (should_stop())
                                {
                                    break;
                                }
                                std::wstring value_lower = ToLower(value_name);
                                if (!trace.selection || !trace::IncludesValue(*trace.selection, key_lower, value_lower))
                                {
                                    continue;
                                }
                                const std::wstring display =
                                    value_name.empty() ? std::wstring(L"(Default)") : value_name;
                                const search::Match match = matcher->Find(display);
                                if (!match.matched)
                                {
                                    continue;
                                }
                                search::Result result;
                                result.key_path = key_path;
                                result.value_name = value_name;
                                result.kind = search::ResultKind::kTraceValue;
                                result.data_state = search::DataState::kNotApplicable;
                                result.match_field = search::MatchField::kName;
                                result.match_start = static_cast<uint32_t>(match.start);
                                result.match_length = static_cast<uint32_t>(match.length);
                                queue_result(std::move(result));
                            }
                        }
                    }
                }
            }
            flush();
        }

        if (!should_stop() && registry_enabled)
        {
            std::atomic<uint64_t> last_progress_tick{0};
            auto progress_cb = [&](uint64_t searched, uint64_t total) {
                search_progress_searched_.store(searched);
                uint64_t now = GetTickCount64();
                uint64_t last = last_progress_tick.load();
                if (now - last < kSearchProgressUiMs && searched < total)
                {
                    return;
                }
                if (last_progress_tick.compare_exchange_strong(last, now))
                {
                    if (!search_progress_posted_.exchange(true))
                    {
                        PostMessageW(hwnd_, frame::message_id::kSearchProgress, static_cast<WPARAM>(generation), 0);
                    }
                }
            };
            search::regex::Status regex_status = search::regex::Status::kNoMatch;
            const bool ok = search::Run(
                criteria,
                &cancel,
                [&](search::ResultBatch&& rows) -> bool {
                    if (should_stop())
                    {
                        return false;
                    }
                    return publish_batch(std::move(rows));
                },
                progress_cb,
                &regex_status
            );
            flush();
            if (!ok || regex_status != search::regex::Status::kNoMatch)
            {
                PostMessageW(hwnd_, frame::message_id::kSearchFailed, static_cast<WPARAM>(generation), static_cast<LPARAM>(regex_status));
                return;
            }
        }

        flush();
        {
            std::lock_guard<std::mutex> lock(search_mutex_);
            search_producer_done_ = true;
        }
        if (!search_posted_.exchange(true))
        {
            if (!PostMessageW(hwnd_, frame::message_id::kSearchResults, static_cast<WPARAM>(generation), 0))
            {
                search_posted_.store(false);
            }
        }
    });
    search_tabs_[static_cast<size_t>(search_index)].generation = generation;
}

namespace
{

enum class DataReplace
{
    kUnchanged,
    kChanged,
    kRejected
};

std::wstring PaddedHex(uint64_t value, size_t width)
{
    static constexpr wchar_t kDigits[] = L"0123456789ABCDEF";
    std::wstring text(width * 2, L'0');
    for (size_t index = 0; index < width * 2; ++index)
    {
        text[width * 2 - 1 - index] = kDigits[(value >> (index * 4)) & 0xF];
    }
    return text;
}

bool ParseHexBytesStrict(const std::wstring& text, std::vector<BYTE>* out)
{
    out->clear();
    auto separator = [](wchar_t c) { return c == L' ' || c == L'\t' || c == L','; };
    size_t index = 0;
    while (index < text.size())
    {
        while (index < text.size() && separator(text[index]))
        {
            ++index;
        }
        if (index >= text.size())
        {
            break;
        }
        const size_t start = index;
        while (index < text.size() && !separator(text[index]))
        {
            ++index;
        }
        if (index - start != 2)
        {
            return false;
        }
        int value = 0;
        for (size_t offset = 0; offset < 2; ++offset)
        {
            const int digit = util::HexDigitValue(text[start + offset]);
            if (digit < 0)
            {
                return false;
            }
            value = value * 16 + digit;
        }
        out->push_back(static_cast<BYTE>(value));
    }
    return true;
}

uint64_t ReadNumber(const std::vector<BYTE>& data, size_t width, bool big_endian)
{
    uint64_t value = 0;
    for (size_t i = 0; i < width; ++i)
    {
        const uint64_t byte = data[big_endian ? i : width - 1 - i];
        value = (value << 8) | byte;
    }
    return value;
}

void WriteNumber(uint64_t value, size_t width, bool big_endian, std::vector<BYTE>* out)
{
    out->assign(width, 0);
    for (size_t i = 0; i < width; ++i)
    {
        const BYTE byte = static_cast<BYTE>((value >> (8 * i)) & 0xFF);
        (*out)[big_endian ? width - 1 - i : i] = byte;
    }
}

DataReplace ApplyReplace(const search::Replacer& matcher, const std::wstring& text, std::wstring* updated)
{
    const search::regex::Status status = matcher.Replace(text, updated);
    if (status == search::regex::Status::kMatch)
    {
        return *updated == text ? DataReplace::kUnchanged : DataReplace::kChanged;
    }
    return status == search::regex::Status::kNoMatch ? DataReplace::kUnchanged : DataReplace::kRejected;
}

DataReplace ReplaceValueData(const search::Replacer& matcher, DWORD type, const std::vector<BYTE>& data, bool number_decimal, bool number_hex, std::vector<BYTE>* out)
{
    const DWORD base = value_format::NormalizeType(type);
    switch (base)
    {
    case REG_SZ:
    case REG_EXPAND_SZ:
    case REG_LINK:
        {
            const std::wstring text = value_format::Data(type, data.data(), static_cast<DWORD>(data.size()));
            std::wstring updated;
            const DataReplace applied = ApplyReplace(matcher, text, &updated);
            if (applied != DataReplace::kChanged)
            {
                return applied;
            }
            *out = value_format::StringData(updated);
            return DataReplace::kChanged;
        }
    case REG_MULTI_SZ:
        {
            std::vector<std::wstring> parts = value_format::MultiStringItems(data);
            bool changed = false;
            for (auto& part : parts)
            {
                std::wstring updated;
                const DataReplace applied = ApplyReplace(matcher, part, &updated);
                if (applied == DataReplace::kRejected)
                {
                    return applied;
                }
                if (applied == DataReplace::kChanged)
                {
                    part = std::move(updated);
                    changed = true;
                }
            }
            if (!changed)
            {
                return DataReplace::kUnchanged;
            }
            *out = value_format::MultiStringData(parts);
            return DataReplace::kChanged;
        }
    case REG_DWORD:
    case REG_DWORD_BIG_ENDIAN:
    case REG_QWORD:
        {
            const bool big_endian = base == REG_DWORD_BIG_ENDIAN;
            const size_t width = base == REG_QWORD ? sizeof(uint64_t) : sizeof(DWORD);
            if (data.size() < width)
            {
                return DataReplace::kUnchanged;
            }
            const uint64_t number = ReadNumber(data, width, big_endian);
            const std::wstring decimal = std::to_wstring(number);
            const std::wstring bare_hex = PaddedHex(number, width);
            const std::wstring prefixed_hex = L"0x" + bare_hex;
            struct NumberForm
            {
                const std::wstring* text;
                int base;
            };
            std::vector<NumberForm> forms;
            if (number_decimal)
            {
                forms.push_back({&decimal, 10});
            }
            if (number_hex)
            {
                forms.push_back({&prefixed_hex, 16});
                forms.push_back({&bare_hex, 16});
            }

            for (const auto& form : forms)
            {
                std::wstring updated;
                const DataReplace applied = ApplyReplace(matcher, *form.text, &updated);
                if (applied == DataReplace::kRejected)
                {
                    return applied;
                }
                if (applied == DataReplace::kUnchanged)
                {
                    continue;
                }
                uint64_t parsed = 0;
                if (!util::ParseUnsignedNumber(updated, form.base, &parsed))
                {
                    return DataReplace::kRejected;
                }
                if (width == sizeof(DWORD) && parsed > MAXDWORD)
                {
                    return DataReplace::kRejected;
                }
                WriteNumber(parsed, width, big_endian, out);
                return DataReplace::kChanged;
            }
            return DataReplace::kUnchanged;
        }
    default:
        {
            const std::wstring text = util::ToHex(data);
            std::wstring updated;
            const DataReplace applied = ApplyReplace(matcher, text, &updated);
            if (applied != DataReplace::kChanged)
            {
                return applied;
            }
            std::vector<BYTE> bytes;
            if (!ParseHexBytesStrict(updated, &bytes))
            {
                return DataReplace::kRejected;
            }
            *out = std::move(bytes);
            return DataReplace::kChanged;
        }
    }
}

} // namespace

void MainWindow::Impl::StartReplace(const ReplaceDialogResult& options)
{
    if (read_only_)
    {
        ui::ShowWarning(hwnd_, L"Read only mode is enabled.");
        return;
    }
    if (options.find_text.empty())
    {
        return;
    }

    RegistryNode start;
    if (!options.start_key.empty())
    {
        if (!ResolvePathToNode(options.start_key, &start))
        {
            ui::ShowError(hwnd_, L"Starting key path wasn't found.");
            return;
        }
    }
    else if (browse_.current_node())
    {
        start = *browse_.current_node();
    }
    else
    {
        ui::ShowError(hwnd_, L"Select a starting key first.");
        return;
    }

    search::Replacer matcher(options);
    if (!matcher.valid())
    {
        ui::ShowError(hwnd_, search::regex::ErrorText(matcher.error()));
        return;
    }

    if (replace_result_pending_)
    {
        ui::ShowWarning(hwnd_, L"Replace is already running.");
        return;
    }

    const HWND hwnd = hwnd_;
    replace_result_pending_ = true;
    replace_session_.Start(
        [this, start, options, matcher, hwnd](uint64_t generation, std::atomic_bool& cancel) mutable {
            auto payload = std::make_unique<ReplacePayload>();
            payload->generation = generation;
            std::vector<RegistryNode> stack;
            std::vector<std::pair<RegistryNode, std::wstring>> key_renames;
            stack.push_back(start);

            while (!stack.empty() && !cancel.load())
            {
                RegistryNode node = std::move(stack.back());
                stack.pop_back();

                std::vector<ValueEntry> values;
                RegistryStore::KeyEnumResult enum_result;
                bool values_reserved = false;
                RegistryStore::EnumKeyStreaming(node, true, true, false, &enum_result, [&](const ValueInfo& info, const BYTE* data, DWORD data_size) {
                    if (!values_reserved)
                    {
                        if (enum_result.info_valid)
                        {
                            values.reserve(enum_result.info.value_count);
                        }
                        values_reserved = true;
                    }
                    ValueEntry value;
                    value.name = info.name;
                    value.type = info.type;
                    if (data_size > 0 && data)
                    {
                        value.data.assign(data, data + data_size);
                    }
                    values.push_back(std::move(value));
                    return !cancel.load();
                },
                                                {});

                for (const auto& value : values)
                {
                    if (cancel.load())
                    {
                        break;
                    }

                    std::wstring current_name = value.name;
                    std::wstring replaced_name;
                    if (options.replace_values && !current_name.empty() &&
                        matcher.Replace(current_name, &replaced_name) == search::regex::Status::kMatch &&
                        replaced_name != current_name)
                    {
                        if (replaced_name.empty())
                        {
                            continue;
                        }
                        std::wstring unique = MakeUniqueValueName(node, replaced_name);
                        bool both_names_left = false;
                        if (!RegistryStore::RenameValue(node, current_name, unique, &both_names_left))
                        {
                            ++payload->failures;
                            if (both_names_left)
                            {
                                ++payload->partial_renames;
                            }
                        }
                        else
                        {
                            ReplacePayload::Change change;
                            change.undo.type = changes::UndoOperation::Type::kRenameValue;
                            change.undo.node = node;
                            change.undo.name = current_name;
                            change.undo.new_name = unique;
                            change.history.action = L"Rename value " + current_name;
                            change.history.old_data = current_name;
                            change.history.new_data = unique;
                            change.history.key_path = registry_path::Build(node);
                            change.history.value_name = unique;
                            payload->changes.push_back(std::move(change));
                            current_name = unique;
                        }
                    }

                    if (!options.replace_data || value.data.empty())
                    {
                        continue;
                    }

                    std::vector<BYTE> new_data;
                    const DataReplace outcome = ReplaceValueData(matcher, value.type, value.data, options.number_decimal, options.number_hex, &new_data);
                    if (outcome == DataReplace::kRejected)
                    {
                        ++payload->rejected;
                        continue;
                    }
                    if (outcome == DataReplace::kUnchanged)
                    {
                        continue;
                    }
                    if (!RegistryStore::SetValue(node, current_name, value.type, new_data))
                    {
                        ++payload->failures;
                        continue;
                    }

                    ValueEntry old_value = value;
                    old_value.name = current_name;
                    ValueEntry new_value = value;
                    new_value.name = current_name;
                    new_value.data = new_data;
                    ReplacePayload::Change change;
                    change.undo.type = changes::UndoOperation::Type::kModifyValue;
                    change.undo.node = node;
                    change.undo.old_value = old_value;
                    change.undo.new_value = new_value;
                    change.history.action = L"Modify value " + current_name;
                    change.history.old_data =
                        value_format::Data(value.type, value.data.data(), static_cast<DWORD>(value.data.size()));
                    change.history.new_data =
                        value_format::Data(value.type, new_data.data(), static_cast<DWORD>(new_data.size()));
                    change.history.key_path = registry_path::Build(node);
                    change.history.value_name = current_name;
                    change.history.revert_kind = HistoryEntry::RevertKind::kSetValue;
                    change.history.revert_value = std::move(old_value);
                    payload->changes.push_back(std::move(change));
                }

                if (options.replace_keys && !node.subkey.empty())
                {
                    const std::wstring leaf = LeafName(node);
                    std::wstring renamed;
                    if (!leaf.empty() && matcher.Replace(leaf, &renamed) == search::regex::Status::kMatch &&
                        renamed != leaf && !renamed.empty())
                    {
                        key_renames.emplace_back(node, renamed);
                    }
                }

                if (options.recursive && !cancel.load())
                {
                    auto subkeys = RegistryStore::EnumSubKeyNames(node, false);
                    for (const auto& name : subkeys)
                    {
                        stack.push_back(ChildNode(node, name));
                    }
                }
            }

            std::stable_sort(key_renames.begin(), key_renames.end(), [](const auto& left, const auto& right) {
                return std::count(left.first.subkey.begin(), left.first.subkey.end(), L'\\') >
                       std::count(right.first.subkey.begin(), right.first.subkey.end(), L'\\');
            });
            for (const auto& rename : key_renames)
            {
                if (cancel.load())
                {
                    break;
                }
                const std::wstring leaf = LeafName(rename.first);
                if (!RegistryStore::RenameKey(rename.first, rename.second))
                {
                    ++payload->failures;
                    continue;
                }
                RegistryNode parent = rename.first;
                const size_t split = parent.subkey.rfind(L'\\');
                parent.subkey = split == std::wstring::npos ? std::wstring() : parent.subkey.substr(0, split);
                ReplacePayload::Change change;
                change.undo.type = changes::UndoOperation::Type::kRenameKey;
                change.undo.node = parent;
                change.undo.name = leaf;
                change.undo.new_name = rename.second;
                change.history.action = L"Rename key " + leaf;
                change.history.old_data = leaf;
                change.history.new_data = rename.second;
                change.history.key_path = registry_path::Build(parent);
                payload->changes.push_back(std::move(change));
            }

            payload->cancelled = cancel.load();
            if (hwnd && IsWindow(hwnd) &&
                PostMessageW(hwnd, frame::message_id::kReplaceReady, static_cast<WPARAM>(generation), reinterpret_cast<LPARAM>(payload.get())))
            {
                ReleasePostedPayload(payload);
            }
        }
    );
}

void MainWindow::Impl::ApplyReplacePayload(ReplacePayload* payload)
{
    if (!payload)
    {
        return;
    }
    std::unique_ptr<ReplacePayload> owned(payload);
    if (!replace_session_.IsCurrent(owned->generation))
    {
        return;
    }
    replace_session_.Join();
    replace_result_pending_ = false;
    CommitReplacePayload(std::move(owned), true);
}

void MainWindow::Impl::CommitReplacePayload(std::unique_ptr<ReplacePayload> payload, bool show_failures)
{
    if (!payload)
    {
        return;
    }
    for (auto& change : payload->changes)
    {
        PushUndo(std::move(change.undo));
        AppendHistoryEntry(std::move(change.history));
    }
    if (!payload->changes.empty() || payload->partial_renames > 0)
    {
        MarkOfflineDirty();
    }
    if (browse_.current_node())
    {
        UpdateValueListForNode(browse_.current_node());
    }
    if (show_failures)
    {
        std::wstring summary = L"Replaced " + std::to_wstring(payload->changes.size()) +
                               (payload->changes.size() == 1 ? L" entry" : L" entries");
        if (payload->failures > 0)
        {
            summary += L", " + std::to_wstring(payload->failures) + L" failed";
        }
        if (payload->rejected > 0)
        {
            summary += L", " + std::to_wstring(payload->rejected) + L" skipped";
        }
        if (payload->cancelled)
        {
            summary += L" (cancelled)";
        }
        SetStatusMessage(summary);
    }
    if (show_failures && (payload->failures > 0 || payload->rejected > 0))
    {
        std::wstring message = L"Replace finished with some failures.\nReplaced: " +
                               std::to_wstring(payload->changes.size()) + L"\nFailed: " +
                               std::to_wstring(payload->failures);
        if (payload->rejected > 0)
        {
            message += L"\nSkipped: " + std::to_wstring(payload->rejected) +
                       L" value(s) because the replacement wasn't valid for the "
                       L"value type.";
        }
        if (payload->partial_renames > 0)
        {
            message += L"\n" + std::to_wstring(payload->partial_renames) +
                       L" value(s) were copied to the new name but the old name "
                       L"couldn't be removed. Both names now exist.";
        }
        ui::ShowError(hwnd_, message);
    }
}

void MainWindow::Impl::StopReplace()
{
    replace_session_.CancelAndJoin();
    MSG message = {};
    while (PeekMessageW(&message, hwnd_, frame::message_id::kReplaceReady, frame::message_id::kReplaceReady, PM_REMOVE))
    {
        CommitReplacePayload(std::unique_ptr<ReplacePayload>(reinterpret_cast<ReplacePayload*>(message.lParam)), false);
    }
    replace_result_pending_ = false;
}

void MainWindow::Impl::CancelSearch()
{
    search_session_.Cancel();
    search_queue_space_.notify_all();
    search_session_.Join();
    search_running_ = false;
    search_preview_request_posted_ = false;
    search_start_tick_ = 0;
    search_duration_ms_ = 0;
    search_duration_valid_ = false;
    search_progress_searched_.store(0);
    search_progress_posted_.store(false);
    search_posted_.store(false);
    {
        std::lock_guard<std::mutex> lock(search_mutex_);
        search_pending_batches_.clear();
        search_pending_rows_ = 0;
        search_producer_done_ = false;
    }
    if (search_progress_)
    {
        SendMessageW(search_progress_, PBM_SETMARQUEE, FALSE, 0);
    }
    ApplyViewVisibility();
    UpdateStatus();
}

void MainWindow::Impl::CloseSearchTab(int tab_index)
{
    if (!tab_ || !IsSearchTabIndex(tab_index))
    {
        return;
    }
    int count = TabCtrl_GetItemCount(tab_);
    if (tab_index >= count)
    {
        return;
    }
    if (search_running_ && active_search_tab_index_ == tab_index)
    {
        CancelSearch();
    }
    int search_index = SearchIndexFromTab(tab_index);
    if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size())
    {
        return;
    }

    const int previous_index = TabCtrl_GetCurSel(tab_);

    search_tabs_.erase(search_tabs_.begin() + search_index);
    tabs_.erase(tabs_.begin() + tab_index);
    for (auto& entry : tabs_)
    {
        if (entry.kind == TabEntry::Kind::kSearch && entry.search_index > search_index)
        {
            --entry.search_index;
        }
    }
    TabCtrl_DeleteItem(tab_, tab_index);
    if (active_search_tab_index_ == tab_index)
    {
        active_search_tab_index_ = -1;
    }
    else if (active_search_tab_index_ > tab_index)
    {
        --active_search_tab_index_;
    }

    SelectTabAfterClose(tab_index, previous_index);
    UpdateTabWidth();
    UpdateSearchResultsView();
    ApplyViewVisibility();
    UpdateStatus();
}

} // namespace regkit
