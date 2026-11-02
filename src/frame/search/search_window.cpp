// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "win32/text_transform.h"
#include "win32/translation.h"
#include "trace/trace_paths.h"

#include <win32ui/system/error_text.h>
#include <win32ui/system/process_token.h>

namespace regkit
{

using namespace window_detail;

std::vector<std::wstring> MainWindow::Impl::RegistryPathContexts() const
{
    std::vector<std::wstring> contexts = {TreeRootLabel()};
    if (session_->mode == RegistryMode::kRemote)
    {
        contexts.push_back(StripMachinePrefix(session_->remote_machine));
    }
    return contexts;
}

std::wstring MainWindow::Impl::NormalizeRegistryPath(const std::wstring& input) const
{
    return NormalizeRegistryPath(input, util::GetCurrentUserSidString(), RegistryPathContexts());
}

std::wstring MainWindow::Impl::NormalizeRegistryPath(const std::wstring& input, const std::wstring& sid, const std::vector<std::wstring>& contexts)
{
    std::wstring path = registry_path::Normalize(input, sid);
    for (const std::wstring& label : contexts)
    {
        if (!label.empty() && util::StartsWithInsensitive(path, label + L"\\"))
        {
            path.erase(0, label.size() + 1);
        }
    }
    return registry_path::Normalize(path, sid);
}

std::wstring MainWindow::Impl::FormatRegistryPath(const std::wstring& path, registry_path::Style style) const
{
    const std::wstring normalized = NormalizeRegistryPath(path);
    if (normalized.empty())
    {
        return {};
    }
    std::wstring tree_root = session_->mode == RegistryMode::kLocal ? L"Computer" : TreeRootLabel();
    return registry_path::Format(normalized, style, tree_root);
}
bool MainWindow::Impl::KeyPathExists(const std::wstring& path, RegistryNode* node) const
{
    return !path.empty() && ResolvePathToNode(path, node) && RegistryStore::KeyExists(*node);
}

bool MainWindow::Impl::FindNearestExistingPath(const std::wstring& path, std::wstring* nearest_path) const
{
    return changes::FindNearestExistingPath(
        path,
        [this](const std::wstring& candidate) {
            RegistryNode node;
            return KeyPathExists(candidate, &node);
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

void MainWindow::Impl::SetStatusMessage(const std::wstring& text, UINT ms)
{
    // keep compare message until result arrives
    win32ui::ShowStatusMessage(status_bar_, text, compare_session_.running() ? 0 : ms);
    UpdateStatus();
}

void MainWindow::Impl::UpdateStatus()
{
    UpdateToolbarStates();
    if (!status_bar_)
    {
        return;
    }
    const std::wstring message = win32ui::StatusMessage(status_bar_);
    // search tabs and loading .reg files use one text part, running searches add progress
    auto show_one = [this](const std::wstring& text, bool searching) {
        std::vector<win32ui::StatusPart> parts = {{L"", 0, true}};
        if (searching)
        {
            parts.push_back({L"", 120});
        }
        win32ui::SetStatusParts(status_bar_, parts);
        win32ui::ShowStatusProgress(status_bar_, 1, searching);
        win32ui::SetStatusText(status_bar_, 0, text);
        if (searching)
        {
            win32ui::SetStatusText(status_bar_, 1, L"");
        }
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
        bool searching = false;
        if (compare_selected)
        {
            swprintf_s(buffer, util::Tr(L"Results: %llu"), count_value);
        }
        else if (const SearchTab* shown = ShownSearchTab(); shown && shown->run)
        {
            searching = true;
            uint64_t searched = shown->run->searched.load();
            if (searched > 0)
            {
                swprintf_s(buffer, util::Tr(L"Searching... Results: ~%llu | Scanned: %llu"), count_value, searched);
            }
            else
            {
                swprintf_s(buffer, util::Tr(L"Searching... Results: ~%llu"), count_value);
            }
        }
        else if (shown && shown->duration_ms > 0)
        {
            double seconds = static_cast<double>(shown->duration_ms) / 1000.0;
            swprintf_s(buffer, util::Tr(L"Results: %llu (%.2fs)"), count_value, seconds);
        }
        else
        {
            swprintf_s(buffer, util::Tr(L"Results: %llu"), count_value);
        }
        show_one(message.empty() ? std::wstring(buffer) : message, searching);
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
                show_one(util::TrLabel(L"Loading", entry.reg_file_label), false);
                return;
            }
        }
    }

    win32ui::ShowStatusProgress(status_bar_, 1, false);
    int selected = ListView_GetSelectedCount(browse_.values().hwnd());
    wchar_t buffer[256] = {};
    std::wstring path_text = message;
    if (path_text.empty() && value_list_error_ != ERROR_SUCCESS && browse_.current_node())
    {
        // reason replaces path until the key is read or left
        path_text = std::wstring(util::Tr(L"The contents of this key couldn't be read.")) + L" " + win32ui::FormatWin32Error(static_cast<DWORD>(value_list_error_));
    }
    if (path_text.empty() && browse_.current_node())
    {
        path_text = registry_path::Build(*browse_.current_node());
    }
    std::wstring mode_text;
    if (session_->mode == RegistryMode::kLocal && (util::ShellUserDiffers() || win32ui::IsProcessSystem()))
    {
        const std::wstring sid = util::GetCurrentUserSidString();
        const std::wstring classes_sid = util::GetClassesUserSidString();
        if (sid != status_account_sid_ || classes_sid != status_classes_sid_)
        {
            status_account_sid_ = sid;
            status_classes_sid_ = classes_sid;
            status_account_ = util::TrLabel(L"HKCU", util::AccountName(sid));
            if (classes_sid != sid)
            {
                status_account_.append(L", ").append(util::TrLabel(L"HKCR", util::AccountName(classes_sid)));
            }
        }
        mode_text = status_account_;
    }
    if (backup_privileges_)
    {
        mode_text.append(mode_text.empty() ? L"" : L", ").append(util::Tr(L"Backup/restore mode"));
    }
    swprintf_s(buffer, util::Tr(L"Keys: %d"), current_key_count_);
    std::wstring keys_text = buffer;
    swprintf_s(buffer, util::Tr(L"Values: %d"), current_value_count_);
    std::wstring values_text = buffer;
    swprintf_s(buffer, util::Tr(L"Selected: %d"), selected);
    std::wstring selected_text = buffer;

    // path fills remaining space, mode part exists only when needed, counts fit five digits and grow for longer values
    std::vector<win32ui::StatusPart> parts = {{L"", 0, true}};
    if (!mode_text.empty())
    {
        parts.push_back({mode_text});
    }
    for (const auto& [format, count] : {std::pair{util::Tr(L"Values: %d"), current_value_count_}, std::pair{util::Tr(L"Selected: %d"), selected}, std::pair{util::Tr(L"Keys: %d"), current_key_count_}})
    {
        swprintf_s(buffer, format, std::max(count, 99999));
        parts.push_back({buffer});
    }
    win32ui::SetStatusParts(status_bar_, parts);
    int part = 0;
    for (const std::wstring* text : {&path_text, &mode_text, &values_text, &selected_text, &keys_text})
    {
        if (text != &mode_text || !mode_text.empty())
        {
            win32ui::SetStatusText(status_bar_, part++, *text);
        }
    }
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
    return entry.kind == TabEntry::Kind::kRegistry && entry.registry_mode == RegistryMode::kLocal && !entry.registry_view;
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
    const bool fresh = ActivateTabTree(index);
    // own session keeps roots and undo out of registry tab sessions
    if (!entry.session)
    {
        entry.session = std::make_shared<RegistrySession>();
    }
    session_ = entry.session;
    UpdateToolbarStates();
    std::vector<RegistryRootEntry>& roots = session_->roots;
    roots.clear();
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
    const bool shown = !fresh && roots.size() == browse_.roots().size() &&
                       std::equal(roots.begin(), roots.end(), browse_.roots().begin(), [](const auto& left, const auto& right) { return left.root == right.root; });
    if (shown)
    {
        ResumeTabTree(index);
    }
    else
    {
        ApplyRegistryRoots(roots);
        RestoreRegistryTabState(index);
    }
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
        ApplySearchColumns(true);
    }
    int max_sort_col = compare ? (show_result ? 4 : 3) : 6;
    if (tab.sort_column > max_sort_col)
    {
        tab.sort_column = -1;
    }
    win32ui::UpdateListViewSort(search_results_list_, tab.sort_column, tab.sort_ascending);
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

bool MainWindow::Impl::CollectSearchStartNodes(const SearchDialogResult& options, const std::wstring& registry_scope_path, std::vector<search::StartNode>* out_nodes, std::vector<search::Source>* out_sources, bool* out_remote)
{
    std::vector<search::StartNode>& start_nodes = *out_nodes;
    std::vector<search::Source>& sources = *out_sources;
    bool& remote_nodes = *out_remote;
    auto source_index = [&](const search::Source& wanted) -> uint16_t {
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
    if (options.scope == SearchScope::kCurrentKey)
    {
        const uint16_t source = source_index(CurrentTabSource());
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
                    ui::ShowError(hwnd_, util::Tr(L"Starting key path wasn't found."));
                    return false;
                }
            }
        }
        else if (browse_.current_node())
        {
            start_nodes.push_back({*browse_.current_node(), source});
        }
        else
        {
            ui::ShowError(hwnd_, util::Tr(L"Select a starting key first."));
            return false;
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
            node.view = entry.view;
            node.root_name = entry.path_name;
            node.subkey = entry.subkey_prefix;
            start_nodes.push_back({std::move(node), source});
        };

        std::vector<RegistryRootEntry> local_roots = RegistryStore::DefaultRoots(settings_.show_extra_hives);
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
        if (options.search_offline_hives && !session_->offline_roots.empty())
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
            const uint16_t source = source_index({search::Source::Kind::kOffline, offline_path});
            for (size_t i = 0; i < session_->offline_roots.size(); ++i)
            {
                RegistryRootEntry entry;
                entry.root = session_->offline_roots[i];
                entry.display_name = i < session_->offline_root_labels.size() ? session_->offline_root_labels[i] : L"OfflineHive";
                entry.path_name = session_->offline_root_name + L"\\" + entry.display_name;
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
                const uint16_t source = source_index({search::Source::Kind::kRegFile, tab.reg_file_path});
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
        if (options.search_remote_registry && session_->remote_hklm)
        {
            const std::wstring prefix = session_->remote_machine + L"\\";
            const uint16_t source = source_index({search::Source::Kind::kRemote, session_->remote_machine});
            remote_nodes = true;
            add_root({session_->remote_hklm, L"HKEY_LOCAL_MACHINE", prefix + L"HKEY_LOCAL_MACHINE", L""}, source);
            if (session_->remote_hku)
            {
                add_root({session_->remote_hku, L"HKEY_USERS", prefix + L"HKEY_USERS", L""}, source);
            }
        }
    }
    return true;
}

std::function<std::wstring(const std::wstring&, const std::wstring&)> MainWindow::Impl::DefaultDataLookup(std::vector<ActiveDefault> defaults)
{
    return [defaults = std::move(defaults)](const std::wstring& path, const std::wstring& name) {
        thread_local std::wstring last_path;
        thread_local std::wstring key_lower;
        if (path != last_path)
        {
            last_path = path;
            const std::wstring normalized = trace::NormalizeKeyPathBasic(path);
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

void MainWindow::Impl::StartSearch(const SearchDialogResult& options)
{
    const bool anomalies = options.criteria.anomalies != 0;
    if (options.criteria.query.empty() && !anomalies)
    {
        ui::ShowWarning(hwnd_, util::Tr(L"Enter text to find."));
        return;
    }

    search::TextOptions match_options;
    match_options.query = options.criteria.query;
    match_options.match_case = options.criteria.match_case;
    match_options.match_whole = options.criteria.match_whole;
    match_options.use_regex = options.criteria.use_regex;
    match_options.match_all = options.criteria.query.empty();
    auto matcher = std::make_shared<const search::Matcher>(match_options);
    if (!matcher->valid())
    {
        ui::ShowError(hwnd_, search::regex::ErrorText(matcher->error()));
        return;
    }

    bool want_registry = options.search_standard_hives || options.search_registry_root ||
                         options.search_offline_hives || options.search_reg_files || options.search_remote_registry;
    bool want_trace = options.search_trace_values && !active_traces_.empty() && !anomalies;
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
            ui::ShowError(hwnd_, util::Tr(L"Select a starting key first."));
            return;
        }
    }

    std::vector<search::StartNode> start_nodes;
    std::vector<search::Source> sources(1);
    bool remote_nodes = false;
    if (want_registry && !CollectSearchStartNodes(options, registry_scope_path, &start_nodes, &sources, &remote_nodes))
    {
        return;
    }

    if (want_registry && start_nodes.empty())
    {
        ui::ShowError(hwnd_, util::Tr(L"No keys to search in the selected sources."));
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

    search::Criteria criteria = options.criteria;
    criteria.matcher = matcher;
    criteria.start_nodes = start_nodes;
    // dialog keeps unticked filters for next time
    if (!options.use_exclude)
    {
        criteria.exclude_paths.clear();
    }
    if (!criteria.search_data)
    {
        criteria.allowed_types.clear();
        criteria.use_min_size = false;
        criteria.use_max_size = false;
    }
    for (std::wstring& path : criteria.exclude_paths)
    {
        path = NormalizeRegistryPath(path);
        while (path.ends_with(L'\\'))
        {
            path.pop_back();
        }
    }
    if (criteria.search_comments)
    {
        const auto comments = std::make_shared<const std::pair<changes::ValueComments, changes::ValueComments>>(value_comments_, default_comments_);
        criteria.comment_text = [comments](const std::wstring& path, const std::wstring* name, DWORD type, DWORD size) {
            return changes::ResolveComment(comments->first, comments->second, {path, name ? *name : std::wstring(), type, size, !name}).text;
        };
    }
    if (options.search_default_data && !active_defaults_.empty() && !anomalies)
    {
        criteria.default_text = DefaultDataLookup(active_defaults_);
    }

    std::wstring label = anomalies && criteria.query.empty() ? util::Tr(L"Anomalies") : util::Tr(L"Find");
    if (!criteria.query.empty())
    {
        label = util::TrLabel(L"Find", criteria.query);
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
        CancelSearch(&tab);
        tab.label = label;
        tab.results.clear();
        tab.last_ui_count = 0;
        tab.is_compare = false;
        tab.sort_dirty = false;
        tab.sources = sources;
        tab.max_results = criteria.max_results;
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
        tab.max_results = criteria.max_results;
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
    search_last_refresh_tick_ = 0;
    const auto run = std::make_shared<SearchRun>();
    run->start_tick = GetTickCount64();
    const uint64_t generation = ++search_generation_;
    SearchTab& started = search_tabs_[static_cast<size_t>(search_index)];
    started.run = run;
    started.generation = generation;
    started.duration_ms = 0;

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
        criteria.provider = session_->mode == RegistryMode::kRemote    ? search::Provider::kRemote
                            : session_->mode == RegistryMode::kOffline ? search::Provider::kOffline
                                                                       : search::Provider::kLocal;
    }
    std::vector<std::wstring> exclude_paths = criteria.exclude_paths;
    std::wstring scope_lower = ToLower(scope_path);
    bool scope_recursive = criteria.recursive;
    bool trace_enabled = want_trace;
    bool registry_enabled = want_registry && !criteria.start_nodes.empty();

    const uint64_t thread = run->session.Start(L"SearchThread", [hwnd = hwnd_, run = run.get(), generation, criteria, traces, exclude_paths, scope_lower, scope_recursive, trace_enabled, registry_enabled, matcher](uint64_t, std::atomic_bool& cancel) mutable {
        auto should_stop = [&]() { return cancel.load(); };

        auto publish_batch = [&](search::ResultBatch&& rows) -> bool {
            if (rows.empty())
            {
                return !cancel.load();
            }
            {
                std::unique_lock<std::mutex> lock(run->mutex);
                run->space.wait(lock, [&]() { return cancel.load() || run->pending_rows < kSearchPendingRowLimit; });
                if (cancel.load())
                {
                    return false;
                }
                run->pending_rows += rows.size();
                run->batches.push_back(std::move(rows));
            }
            if (!run->posted.exchange(true) && !PostMessageW(hwnd, frame::message_id::kSearchResults, static_cast<WPARAM>(generation), 0))
            {
                run->posted.store(false);
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
                                    value_name.empty() ? std::wstring(util::Tr(L"(Default)")) : value_name;
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
                run->searched.store(searched);
                uint64_t now = GetTickCount64();
                uint64_t last = last_progress_tick.load();
                if (now - last < kSearchProgressUiMs && searched < total)
                {
                    return;
                }
                if (last_progress_tick.compare_exchange_strong(last, now))
                {
                    if (!run->progress_posted.exchange(true))
                    {
                        PostMessageW(hwnd, frame::message_id::kSearchProgress, static_cast<WPARAM>(generation), 0);
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
                PostMessageW(hwnd, frame::message_id::kSearchFailed, static_cast<WPARAM>(generation), static_cast<LPARAM>(regex_status));
                return;
            }
        }

        flush();
        {
            std::lock_guard<std::mutex> lock(run->mutex);
            run->producer_done = true;
        }
        if (!run->posted.exchange(true) && !PostMessageW(hwnd, frame::message_id::kSearchResults, static_cast<WPARAM>(generation), 0))
        {
            run->posted.store(false);
        }
    });
    // like a failed search, the tab stops searching
    if (!thread)
    {
        CancelSearch(&started);
        ui::ShowStartError(hwnd_);
    }
}

namespace
{

enum class DataReplace
{
    kUnchanged,
    kChanged,
    kRejected
};

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
    case REG_MULTI_SZ:
        {
            const bool multi = base == REG_MULTI_SZ;
            const std::wstring text(reinterpret_cast<const wchar_t*>(data.data()), data.size() / sizeof(wchar_t));
            const size_t body = multi ? text.size() : text.find_last_not_of(L'\0') + 1;
            std::wstring updated;
            bool changed = false;
            for (size_t start = 0; start <= body;)
            {
                const size_t end = multi ? std::min(text.find(L'\0', start), body) : body;
                const std::wstring part = text.substr(start, end - start);
                if (multi && part.empty())
                {
                    updated += text.substr(start);
                    break;
                }
                std::wstring replaced;
                const DataReplace applied = part.empty() ? DataReplace::kUnchanged : ApplyReplace(matcher, part, &replaced);
                if (applied == DataReplace::kRejected)
                {
                    return applied;
                }
                changed = changed || applied == DataReplace::kChanged;
                updated += applied == DataReplace::kChanged ? replaced : part;
                if (end < body)
                {
                    updated.push_back(L'\0');
                }
                start = end + 1;
            }
            if (!changed)
            {
                return DataReplace::kUnchanged;
            }
            if (!multi)
            {
                updated += text.substr(body);
            }
            out->assign(reinterpret_cast<const BYTE*>(updated.data()), reinterpret_cast<const BYTE*>(updated.data() + updated.size()));
            if (data.size() % sizeof(wchar_t))
            {
                out->push_back(data.back());
            }
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
            const uint64_t number = value_format::ReadUnsigned(data, width, big_endian);
            const std::wstring decimal = std::to_wstring(number);
            const std::wstring prefixed_hex = win32ui::FormatHex(number, static_cast<int>(width * 2));
            const std::wstring bare_hex = prefixed_hex.substr(2);
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
                const std::optional<uint64_t> parsed = win32ui::ParseUnsigned(updated, form.base, true);
                if (!parsed || (width == sizeof(DWORD) && *parsed > MAXDWORD))
                {
                    return DataReplace::kRejected;
                }
                *out = value_format::UnsignedBytes(*parsed, width, big_endian);
                return DataReplace::kChanged;
            }
            return DataReplace::kUnchanged;
        }
    default:
        {
            const std::wstring text = win32ui::HexBytes(data);
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

}

void MainWindow::Impl::StartReplace(const ReplaceDialogResult& options)
{
    if (settings_.read_only)
    {
        ui::ShowWarning(hwnd_, util::Tr(L"Read only mode is enabled."));
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
            ui::ShowError(hwnd_, util::Tr(L"Starting key path wasn't found."));
            return;
        }
    }
    else if (browse_.current_node())
    {
        start = *browse_.current_node();
    }
    else
    {
        ui::ShowError(hwnd_, util::Tr(L"Select a starting key first."));
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
        ui::ShowWarning(hwnd_, util::Tr(L"Replace is already running."));
        return;
    }

    const HWND hwnd = hwnd_;
    const uint64_t started = replace_session_.Start(
        L"ReplaceThread",
        [this, start, options, matcher, hwnd, session = std::weak_ptr<RegistrySession>(session_), source = HistorySource()](uint64_t generation, std::atomic_bool& cancel) mutable {
            auto payload = std::make_unique<ReplacePayload>();
            payload->generation = generation;
            payload->session = session;
            payload->source = std::move(source);
            std::vector<RegistryNode> stack;
            std::vector<std::pair<RegistryNode, std::wstring>> key_renames;
            stack.push_back(start);

            while (!stack.empty() && !cancel.load())
            {
                RegistryNode node = std::move(stack.back());
                stack.pop_back();

                std::vector<RegistryValue> values;
                RegistryStore::KeyEnumResult enum_result;
                bool values_reserved = false;
                const bool listed = RegistryStore::EnumKeyStreaming(node, true, true, false, &enum_result, [&](const ValueInfo& info, const BYTE* data, DWORD data_size) {
                    if (!values_reserved)
                    {
                        if (enum_result.info_valid)
                        {
                            values.reserve(enum_result.info.value_count);
                        }
                        values_reserved = true;
                    }
                    RegistryValue value;
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
                if (!cancel.load() && (!listed || enum_result.error != ERROR_SUCCESS))
                {
                    ++payload->failures;
                }

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
                        const std::wstring unique = MakeUniqueValueName(node, replaced_name).value_or(L"");
                        bool both_names_left = false;
                        if (unique.empty() || !RegistryStore::RenameValue(node, current_name, unique, &both_names_left))
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

                    RegistryValue old_value = value;
                    old_value.name = current_name;
                    RegistryValue new_value = value;
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
                parent.subkey = registry_path::Parent(rename.first.subkey);
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
            if (hwnd && IsWindow(hwnd))
            {
                win32ui::PostPayload(hwnd, frame::message_id::kReplaceReady, static_cast<WPARAM>(generation), payload);
            }
        }
    );
    if (!started)
    {
        ui::ShowStartError(hwnd_);
        return;
    }
    // result arrives as a message, after these are set
    replace_result_pending_ = true;
    replace_target_ = session_;
}

void MainWindow::Impl::ApplyReplacePayload(std::unique_ptr<ReplacePayload> owned)
{
    if (!owned)
    {
        return;
    }
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
    std::vector<changes::UndoOperation> steps;
    std::vector<HistoryEntry> history;
    const bool renamed_keys = std::any_of(payload->changes.begin(), payload->changes.end(), [](const ReplacePayload::Change& change) {
        return change.undo.type == changes::UndoOperation::Type::kRenameKey;
    });
    steps.reserve(payload->changes.size());
    history.reserve(payload->changes.size());
    for (auto& change : payload->changes)
    {
        steps.push_back(std::move(change.undo));
        history.push_back(std::move(change.history));
    }
    AppendHistoryEntries(std::move(history), payload->source);
    const auto session = payload->session.lock();
    if (session)
    {
        session->undo.Push(std::move(steps));
        UpdateToolbarStates();
    }
    // mark changes in the replace session, closed sources have nothing left to mark
    if (session && (!payload->changes.empty() || payload->partial_renames > 0))
    {
        MarkSessionDirty(*session);
    }
    // renamed keys can lie anywhere in the visible tree, not just under selection
    if (renamed_keys)
    {
        RefreshWholeTree();
    }
    if (browse_.current_node())
    {
        UpdateValueListForNode(browse_.current_node());
    }
    if (show_failures)
    {
        std::wstring summary = util::TrLabel(L"Replaced", std::to_wstring(payload->changes.size()));
        if (payload->failures > 0)
        {
            summary += L", " + util::TrLabel(L"Failed", std::to_wstring(payload->failures));
        }
        if (payload->rejected > 0)
        {
            summary += L", " + util::TrLabel(L"Skipped", std::to_wstring(payload->rejected));
        }
        if (payload->cancelled)
        {
            summary += L" (" + std::wstring(util::Tr(L"cancelled")) + L")";
        }
        SetStatusMessage(summary);
    }
    if (show_failures && (payload->failures > 0 || payload->rejected > 0))
    {
        std::wstring message = util::TrDetail(L"Replace finished with some failures.", util::TrLabel(L"Replaced", std::to_wstring(payload->changes.size()))) + L"\n" +
                               util::TrLabel(L"Failed", std::to_wstring(payload->failures));
        if (payload->rejected > 0)
        {
            message += L"\n" + util::TrLabel(L"Skipped, the replacement isn't valid for the value type", std::to_wstring(payload->rejected));
        }
        if (payload->partial_renames > 0)
        {
            message += L"\n" + util::TrLabel(L"Copied to the new name, the old name couldn't be removed", std::to_wstring(payload->partial_renames));
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

void MainWindow::Impl::CancelSearch(SearchTab* tab)
{
    if (!tab || !tab->run)
    {
        return;
    }
    tab->run.reset();
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
    int search_index = SearchIndexFromTab(tab_index);
    if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size())
    {
        return;
    }
    CancelSearch(&search_tabs_[static_cast<size_t>(search_index)]);

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

}
