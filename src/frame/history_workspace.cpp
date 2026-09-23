// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

namespace regkit
{

namespace
{

search::Source::Kind ToSourceKind(int value)
{
    return value < 0 || value > static_cast<int>(search::Source::Kind::kRegFile)
               ? search::Source::Kind::kLocal
               : static_cast<search::Source::Kind>(value);
}

search::Source::Kind LegacyCompareKind(int value)
{
    switch (value)
    {
    case 1:
        return search::Source::Kind::kRegFile;
    case 2:
        return search::Source::Kind::kOffline;
    default:
        return search::Source::Kind::kLocal;
    }
}

bool DeleteCacheFile(const std::wstring& path)
{
    if (path.empty() || DeleteFileW(path.c_str()) != 0)
    {
        return true;
    }
    const DWORD error = GetLastError();
    return error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND;
}

bool DeleteCacheFiles(const std::wstring& folder, const wchar_t* pattern_name)
{
    if (folder.empty() || !pattern_name || !*pattern_name)
    {
        return false;
    }
    const std::wstring pattern = util::JoinPath(folder, pattern_name);
    WIN32_FIND_DATAW data = {};
    HANDLE find = FindFirstFileW(pattern.c_str(), &data);
    if (find == INVALID_HANDLE_VALUE)
    {
        const DWORD error = GetLastError();
        return error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND;
    }
    bool deleted = true;
    do
    {
        if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
        {
            deleted = DeleteCacheFile(util::JoinPath(folder, data.cFileName)) && deleted;
        }
    } while (FindNextFileW(find, &data) != 0);
    const DWORD error = GetLastError();
    FindClose(find);
    return deleted && error == ERROR_NO_MORE_FILES;
}

} // namespace
using namespace window_detail;

void MainWindow::Impl::AppendHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data)
{
    HistoryEntry entry;
    entry.action = action;
    entry.old_data = old_data;
    entry.new_data = new_data;
    if (browse_.current_node())
    {
        entry.key_path = registry_path::Build(*browse_.current_node());
    }
    AppendHistoryEntry(std::move(entry));
}

void MainWindow::Impl::AppendValueHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data, const RegistryNode& node, const std::wstring& value_name, HistoryEntry::RevertKind revert_kind, const ValueEntry* revert_value)
{
    HistoryEntry entry;
    entry.action = action;
    entry.old_data = old_data;
    entry.new_data = new_data;
    entry.key_path = registry_path::Build(node);
    entry.value_name = value_name;
    entry.revert_kind = revert_kind;
    if (revert_value)
    {
        entry.revert_value = *revert_value;
    }
    AppendHistoryEntry(std::move(entry));
}

void MainWindow::Impl::AppendHistoryEntry(HistoryEntry entry)
{
    if (!history_list_)
    {
        return;
    }
    entry.action = registry_path::DisplayName(entry.action);
    entry.old_data = registry_path::DisplayName(entry.old_data);
    entry.new_data = registry_path::DisplayName(entry.new_data);

    const HistoryEntry appended = change_history_.Append(std::move(entry), static_cast<size_t>(history_max_rows_));
    if (history_loaded_)
    {
        AppendHistoryCache(appended);
    }
    change_history_.Sort(history_sort_column_, history_sort_ascending_);
    RebuildHistoryList();
}

bool MainWindow::Impl::PrepareHistoryRevert(const HistoryEntry& entry, HistoryEntry* prepared) const
{
    auto query_value = [this](const std::wstring& path, const std::wstring& name, ValueEntry* value) {
        RegistryNode node;
        return ResolvePathToNode(path, &node) && RegistryStore::QueryValue(node, name, value);
    };
    if (!changes::PrepareRevert(entry, query_value, prepared))
    {
        return false;
    }
    if (prepared->revert_kind != HistoryEntry::RevertKind::kDeleteValue)
    {
        return true;
    }

    std::vector<const HistoryEntry*> later_entries;
    later_entries.reserve(change_history_.entries().size());
    for (const HistoryEntry& candidate : change_history_.entries())
    {
        if (candidate.timestamp > entry.timestamp && EqualsInsensitive(candidate.key_path, entry.key_path) &&
            StartsWithInsensitive(candidate.action, L"Rename value "))
        {
            later_entries.push_back(&candidate);
        }
    }
    std::stable_sort(
        later_entries.begin(),
        later_entries.end(),
        [](const HistoryEntry* left, const HistoryEntry* right) { return left->timestamp < right->timestamp; }
    );
    for (const HistoryEntry* candidate : later_entries)
    {
        if (!EqualsInsensitive(candidate->old_data, prepared->value_name))
        {
            continue;
        }
        prepared->value_name = candidate->value_name.empty() ? candidate->new_data : candidate->value_name;
    }

    ValueEntry current;
    return query_value(prepared->key_path, prepared->value_name, &current);
}

bool MainWindow::Impl::OpenHistoryTarget(const HistoryEntry& entry)
{
    if (entry.key_path.empty())
    {
        return false;
    }
    RegistryNode node;
    std::wstring target = entry.key_path;
    KeyInfo info = {};
    if (!ResolvePathToNode(target, &node) || !RegistryStore::QueryKeyInfo(node, &info))
    {
        if (!FindNearestExistingPath(target, &target) || target.empty())
        {
            return false;
        }
        return NavigateToResolvedExternalJump(target, L"");
    }
    return NavigateToResolvedExternalJump(target, entry.value_name);
}

bool MainWindow::Impl::RevertHistoryEntry(const HistoryEntry& entry)
{
    HistoryEntry prepared;
    if (!EnsureWritable() || !PrepareHistoryRevert(entry, &prepared))
    {
        return false;
    }
    if (util::IsProcessPrivileged() &&
        ui::PromptKeyChoice(hwnd_, L"Revert this change with elevated rights?", prepared.value_name.empty() ? prepared.key_path : prepared.key_path + L"\\" + prepared.value_name, L"Revert", L"Revert", L"", L"Cancel") != IDYES)
    {
        return false;
    }

    bool ok = false;
    is_replaying_ = true;
    switch (prepared.revert_kind)
    {
    case HistoryEntry::RevertKind::kSetValue:
        {
            RegistryNode node;
            if (ResolvePathToNode(prepared.key_path, &node))
            {
                ok = RegistryStore::SetValue(node, prepared.revert_value.name, prepared.revert_value.type, prepared.revert_value.data);
            }
            break;
        }
    case HistoryEntry::RevertKind::kDeleteValue:
        {
            RegistryNode node;
            if (ResolvePathToNode(prepared.key_path, &node))
            {
                ok = RegistryStore::DeleteValue(node, prepared.value_name);
            }
            break;
        }
    case HistoryEntry::RevertKind::kDeleteKey:
        {
            RegistryNode node;
            if (ResolvePathToNode(prepared.key_path, &node))
            {
                std::wstring name = LeafName(node);
                if (!name.empty() && ui::ConfirmDelete(hwnd_, L"Revert Key Creation", registry_path::DisplayName(name)))
                {
                    ok = RegistryStore::DeleteKey(node);
                }
            }
            break;
        }
    default:
        break;
    }
    is_replaying_ = false;

    if (!ok)
    {
        ui::ShowError(hwnd_, L"Failed to revert history entry.");
        return false;
    }

    MarkOfflineDirty();
    HistoryEntry revert_entry;
    revert_entry.action = L"Revert: " + entry.action;
    revert_entry.key_path = prepared.key_path;
    revert_entry.value_name = prepared.value_name;
    AppendHistoryEntry(std::move(revert_entry));
    RefreshTreeSelection();
    if (browse_.current_node())
    {
        UpdateValueListForNode(browse_.current_node());
    }
    return true;
}

bool MainWindow::Impl::HistoryStaysInMemory() const
{
    return util::IsProcessSystem() || util::IsProcessTrustedInstaller();
}

bool MainWindow::Impl::AppendHistoryCache(const HistoryEntry& entry)
{
    if (HistoryStaysInMemory())
    {
        return true;
    }
    if (changes::AppendHistoryFile(HistoryCachePath(), entry))
    {
        history_cache_failed_ = false;
        return true;
    }
    if (!history_cache_failed_)
    {
        history_cache_failed_ = true;
        ui::ShowError(hwnd_, L"The history couldn't be written to disk. It is "
                             L"kept for this session only.");
    }
    return false;
}

std::wstring MainWindow::Impl::CacheFolderPath() const
{
    return util::GetCacheFolder();
}

std::wstring MainWindow::Impl::HistoryCachePath() const
{
    std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"history.tsv");
}

std::wstring MainWindow::Impl::TabsCachePath() const
{
    std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"tabs.ini");
}

std::wstring MainWindow::Impl::SessionCachePath() const
{
    std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"session.ini");
}

std::wstring MainWindow::Impl::SearchTabCachePath(const std::wstring& file) const
{
    std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return L"";
    }
    if (file.empty() || file.find_first_of(L"\\/:") != std::wstring::npos)
    {
        return L"";
    }
    return util::JoinPath(folder, file);
}

bool MainWindow::Impl::EnsureSearchTabResultsLoaded(int search_index)
{
    if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size())
    {
        return false;
    }
    SearchTab& tab = search_tabs_[static_cast<size_t>(search_index)];
    if (tab.results_loaded)
    {
        return true;
    }
    tab.last_ui_count = 0;
    if (tab.is_compare)
    {
        tab.results_loaded = true;
        if (!tab.compare_cache_file.empty())
        {
            search::compare::LoadRows(SearchTabCachePath(tab.compare_cache_file), &tab.compare_rows);
        }
        return true;
    }
    if (tab.cache_file.empty())
    {
        tab.results_loaded = true;
        return true;
    }

    if (tab.load_pending)
    {
        return false;
    }
    tab.load_pending = true;
    StartSearchTabLoadWorker();
    auto task = std::make_unique<SearchTabLoadTask>();
    task->generation = ++search_tab_load_generation_;
    task->tab_index = search_index;
    task->path = SearchTabCachePath(tab.cache_file);
    task->sort_column = tab.sort_column;
    task->sort_ascending = tab.sort_ascending;
    task->hwnd = hwnd_;
    tab.load_generation = task->generation;

    std::unique_ptr<SearchTabLoadTask> displaced = search_tab_loader_.Submit(std::move(task));
    if (displaced && displaced->tab_index >= 0 && static_cast<size_t>(displaced->tab_index) < search_tabs_.size())
    {
        SearchTab& stale = search_tabs_[static_cast<size_t>(displaced->tab_index)];
        if (stale.load_generation == displaced->generation)
        {
            stale.load_pending = false;
        }
    }
    return false;
}

void MainWindow::Impl::ApplySearchTabLoad(SearchTabLoadPayload* payload)
{
    if (!payload)
    {
        return;
    }
    std::unique_ptr<SearchTabLoadPayload> owned(payload);
    if (owned->tab_index < 0 || static_cast<size_t>(owned->tab_index) >= search_tabs_.size())
    {
        return;
    }
    SearchTab& tab = search_tabs_[static_cast<size_t>(owned->tab_index)];
    if (tab.load_generation != owned->generation)
    {
        return;
    }
    tab.load_pending = false;
    tab.results = std::move(owned->rows);
    for (auto& row : tab.results)
    {
        row.row_id = tab.next_row_id++;
    }
    tab.sort_dirty = false;
    tab.results_loaded = true;
    if (SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) == owned->tab_index)
    {
        UpdateSearchResultsView();
        UpdateStatus();
    }
}

bool MainWindow::Impl::ClearTabsCache()
{
    bool cleared = true;
    for (const std::wstring& path : {TabsCachePath(), SessionCachePath()})
    {
        cleared = DeleteCacheFile(path) && cleared;
    }
    std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return false;
    }
    for (const wchar_t* pattern_name : {L"search_*.tsv", L"compare_*.tsv"})
    {
        cleared = DeleteCacheFiles(folder, pattern_name) && cleared;
    }
    return cleared;
}

bool MainWindow::Impl::ClearCache(CacheKind kind, bool resume_tree_worker)
{
    bool cleared = true;
    const bool all = kind == CacheKind::kAll;
    bool restart_tree_worker = false;
    const std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return false;
    }
    if (all || kind == CacheKind::kTabs)
    {
        cleared = ClearTabsCache() && cleared;
    }
    if (all || kind == CacheKind::kHistory)
    {
        ClearHistoryItems(false);
        history_cache_failed_ = false;
        cleared = DeleteCacheFile(HistoryCachePath()) && cleared;
    }
    if (all || kind == CacheKind::kSearchHistory)
    {
        cleared = DeleteCacheFile(util::JoinPath(folder, L"search_history.txt")) && cleared;
    }
    if (all || kind == CacheKind::kTreeState)
    {
        KillTimer(hwnd_, kTreeStateTimerId);
        StopTreeStateWorker();
        saved_tree_state_.Clear();
        tree_state_restored_ = false;
        cleared = DeleteCacheFile(TreeStatePath()) && cleared;
        restart_tree_worker = resume_tree_worker && save_tree_state_;
    }
    if (all || kind == CacheKind::kTemporary)
    {
        cleared = DeleteCacheFiles(folder, L"export_*.reg") && cleared;
    }
    if (all)
    {
        cleared = DeleteCacheFiles(folder, L"*") && cleared;
    }
    if (restart_tree_worker)
    {
        StartTreeStateWorker();
    }
    return cleared;
}

void MainWindow::Impl::LoadTabs()
{
    if (!tab_)
    {
        return;
    }
    tabs_.clear();
    search_tabs_.clear();
    active_search_tab_index_ = -1;
    TabCtrl_DeleteAllItems(tab_);

    int active_index = 0;
    bool loaded = false;
    const bool restore_session = win32::RestoreSessionRequested();
    const std::wstring session_path = SessionCachePath();
    workspace::TabState state;
    if (restore_session && workspace::LoadTabs(session_path, &state))
    {
        loaded = true;
        DeleteFileW(session_path.c_str());
    }
    else if (save_tab_kinds_ != 0)
    {
        loaded = workspace::LoadTabs(TabsCachePath(), &state);
    }
    if (loaded)
    {
        active_index = state.active_index;
        for (workspace::PersistedTab& saved : state.tabs)
        {
            std::wstring label = std::move(saved.label);
            if (saved.kind == workspace::PersistedTab::Kind::kSearch)
            {
                SearchTab search_tab;
                search_tab.label = label.empty() ? L"Find" : std::move(label);
                search_tab.cache_file = std::move(saved.search_cache_file);
                search_tab.compare_cache_file = std::move(saved.compare_cache_file);
                search_tab.is_compare = saved.is_compare || StartsWithInsensitive(search_tab.label, L"Compare:");
                if (saved.compare_filter >= 0 &&
                    saved.compare_filter <= static_cast<int>(search::compare::RowFilter::kAll))
                {
                    search_tab.compare_filter = static_cast<search::compare::RowFilter>(saved.compare_filter);
                }
                search_tab.results_loaded =
                    search_tab.is_compare ? search_tab.compare_cache_file.empty() : search_tab.cache_file.empty();
                for (size_t s = 0; s < saved.source_kinds.size(); ++s)
                {
                    search::Source source;
                    source.kind = ToSourceKind(saved.source_kinds[s]);
                    if (s < saved.source_names.size())
                    {
                        source.name = std::move(saved.source_names[s]);
                    }
                    search_tab.sources.push_back(std::move(source));
                }
                if (search_tab.sources.empty() && search_tab.is_compare)
                {
                    search_tab.sources = {
                        {LegacyCompareKind(saved.first_source_kind), std::move(saved.first_source_file)},
                        {LegacyCompareKind(saved.second_source_kind), std::move(saved.second_source_file)}
                    };
                }
                search_tabs_.push_back(std::move(search_tab));
                const int search_index = static_cast<int>(search_tabs_.size() - 1);
                TCITEMW item = {};
                item.mask = TCIF_TEXT;
                item.pszText = search_tabs_.back().label.data();
                TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
                tabs_.push_back({TabEntry::Kind::kSearch, search_index});
                continue;
            }
            if (saved.kind == workspace::PersistedTab::Kind::kRegFile)
            {
                if (label.empty())
                {
                    label = FileNameOnly(saved.source_path);
                }
                TCITEMW item = {};
                item.mask = TCIF_TEXT;
                item.pszText = label.data();
                TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
                TabEntry entry;
                entry.kind = TabEntry::Kind::kRegFile;
                entry.reg_file_path = std::move(saved.source_path);
                entry.reg_file_label = label;
                entry.selected_path = std::move(saved.selected_path);
                entry.selected_value = std::move(saved.selected_value);
                entry.selected_values = std::move(saved.selected_values);
                entry.value_top_index = saved.value_top_index;
                entry.expanded_paths = std::move(saved.expanded_paths);
                tabs_.push_back(std::move(entry));
                continue;
            }
            if (label.empty())
            {
                label = L"Local Registry";
            }
            TCITEMW item = {};
            item.mask = TCIF_TEXT;
            item.pszText = label.data();
            TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
            TabEntry entry;
            entry.kind = TabEntry::Kind::kRegistry;
            entry.registry_mode = static_cast<RegistryMode>(saved.registry_mode);
            entry.offline_path =
                entry.registry_mode == RegistryMode::kOffline ? std::move(saved.source_path) : std::wstring();
            entry.remote_machine = std::move(saved.remote_machine);
            entry.selected_path = std::move(saved.selected_path);
            entry.selected_value = std::move(saved.selected_value);
            entry.selected_values = std::move(saved.selected_values);
            entry.value_top_index = saved.value_top_index;
            entry.expanded_paths = std::move(saved.expanded_paths);
            tabs_.push_back(std::move(entry));
        }
    }

    if (!loaded || tabs_.empty())
    {
        TCITEMW item = {};
        item.mask = TCIF_TEXT;
        item.pszText = const_cast<wchar_t*>(L"Local Registry");
        TabCtrl_InsertItem(tab_, 0, &item);
        TabEntry entry;
        entry.kind = TabEntry::Kind::kRegistry;
        entry.registry_mode = RegistryMode::kLocal;
        tabs_.push_back(std::move(entry));
        active_index = 0;
    }

    int count = TabCtrl_GetItemCount(tab_);
    if (count > 0)
    {
        int sel = ClampValue(active_index, 0, count - 1);
        TabCtrl_SetCurSel(tab_, sel);
        if (IsSearchTabIndex(sel))
        {
            active_search_tab_index_ = sel;
        }
    }
    RefreshRegistryTabLabels();
}

bool MainWindow::Impl::SaveTabs()
{
    return SaveTabState(TabsCachePath(), save_tab_kinds_);
}

bool MainWindow::Impl::SaveSessionTabs()
{
    return SaveTabState(SessionCachePath(), workspace::kSaveTabsAll);
}

int MainWindow::Impl::TabSaveKind(const TabEntry& entry) const
{
    if (entry.kind == TabEntry::Kind::kRegFile)
    {
        return workspace::kSaveTabsRegFile;
    }
    if (entry.kind == TabEntry::Kind::kSearch)
    {
        const size_t index = static_cast<size_t>(entry.search_index);
        return entry.search_index >= 0 && index < search_tabs_.size() && search_tabs_[index].is_compare
                   ? workspace::kSaveTabsCompare
                   : workspace::kSaveTabsSearch;
    }
    return entry.registry_mode == RegistryMode::kOffline  ? workspace::kSaveTabsOffline
           : entry.registry_mode == RegistryMode::kRemote ? workspace::kSaveTabsRemote
                                                          : workspace::kSaveTabsLocal;
}

bool MainWindow::Impl::SaveTabState(const std::wstring& path, int kinds)
{
    if (!tab_ || path.empty())
    {
        return true;
    }
    CaptureRegistryTabState(TabCtrl_GetCurSel(tab_));
    std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return false;
    }

    std::unordered_set<std::wstring> referenced_files;
    std::unordered_set<std::wstring> reserved_files;
    for (const auto& search_tab : search_tabs_)
    {
        if (!search_tab.cache_file.empty())
        {
            reserved_files.insert(search_tab.cache_file);
        }
    }
    workspace::TabState state;
    bool saved_all = true;
    int active_index = TabCtrl_GetCurSel(tab_);
    int saved_active_index = -1;

    int search_file_index = 0;
    int tab_count = TabCtrl_GetItemCount(tab_);
    int saved_index = 0;
    for (int i = 0; i < tab_count; ++i)
    {
        if (static_cast<size_t>(i) >= tabs_.size())
        {
            break;
        }
        const auto& entry = tabs_[static_cast<size_t>(i)];
        const int entry_kind = TabSaveKind(entry);
        if ((kinds & entry_kind) == 0)
        {
            continue;
        }
        if (i == active_index)
        {
            saved_active_index = saved_index;
        }
        wchar_t text[256] = {};
        TCITEMW item = {};
        item.mask = TCIF_TEXT;
        item.pszText = text;
        item.cchTextMax = static_cast<int>(_countof(text));
        std::wstring label;
        if (TabCtrl_GetItem(tab_, i, &item))
        {
            label = text;
        }
        if (entry.kind == TabEntry::Kind::kSearch)
        {
            int search_index = entry.search_index;
            if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size())
            {
                continue;
            }
            SearchTab& search_tab = search_tabs_[static_cast<size_t>(search_index)];
            std::wstring& stored_name = search_tab.is_compare ? search_tab.compare_cache_file : search_tab.cache_file;
            std::wstring file_name = stored_name;
            if (file_name.empty())
            {
                const wchar_t* prefix = search_tab.is_compare ? L"compare_" : L"search_";
                do
                {
                    file_name = prefix + std::to_wstring(search_file_index++) + L".tsv";
                } while (referenced_files.find(file_name) != referenced_files.end() ||
                         reserved_files.find(file_name) != reserved_files.end());
                stored_name = file_name;
                reserved_files.insert(file_name);
            }
            if (search_tab.results_loaded)
            {
                const bool written =
                    search_tab.is_compare
                        ? search::compare::SaveRows(SearchTabCachePath(file_name), search_tab.compare_rows)
                        : search::SaveResults(SearchTabCachePath(file_name), search_tab.results);
                if (!written)
                {
                    saved_all = false;
                    continue;
                }
            }
            referenced_files.insert(file_name);
            if (label.empty())
            {
                label = search_tab.label;
            }
            workspace::PersistedTab saved;
            saved.kind = workspace::PersistedTab::Kind::kSearch;
            saved.label = std::move(label);
            saved.is_compare = search_tab.is_compare;
            saved.compare_filter = static_cast<int>(search_tab.compare_filter);
            for (const search::Source& source : search_tab.sources)
            {
                saved.source_kinds.push_back(static_cast<int>(source.kind));
                saved.source_names.push_back(source.name);
            }
            if (search_tab.is_compare)
            {
                saved.compare_cache_file = std::move(file_name);
            }
            else
            {
                saved.search_cache_file = std::move(file_name);
            }
            state.tabs.push_back(std::move(saved));
        }
        else if (entry.kind == TabEntry::Kind::kRegFile)
        {
            if (entry.reg_file_path.empty())
            {
                continue;
            }
            workspace::PersistedTab saved;
            saved.kind = workspace::PersistedTab::Kind::kRegFile;
            saved.label = label.empty() ? entry.reg_file_label : std::move(label);
            saved.source_path = entry.reg_file_path;
            saved.selected_path = entry.selected_path;
            saved.selected_value = entry.selected_value;
            saved.selected_values = entry.selected_values;
            saved.value_top_index = entry.value_top_index;
            saved.expanded_paths = entry.expanded_paths;
            state.tabs.push_back(std::move(saved));
        }
        else
        {
            const TabEntry& registry_entry = tabs_[static_cast<size_t>(i)];
            if (label.empty())
            {
                if (registry_entry.registry_mode == RegistryMode::kLocal)
                {
                    label = LocalRegistryTabLabel(i);
                }
                else
                {
                    label = L"Local Registry";
                }
            }
            workspace::PersistedTab saved;
            saved.kind = workspace::PersistedTab::Kind::kRegistry;
            saved.label = std::move(label);
            saved.selected_path = registry_entry.selected_path;
            saved.selected_value = registry_entry.selected_value;
            saved.selected_values = registry_entry.selected_values;
            saved.value_top_index = registry_entry.value_top_index;
            saved.expanded_paths = registry_entry.expanded_paths;
            saved.registry_mode = static_cast<int>(registry_entry.registry_mode);
            saved.source_path = registry_entry.offline_path;
            saved.remote_machine = registry_entry.remote_machine;
            state.tabs.push_back(std::move(saved));
        }
        ++saved_index;
    }
    if (saved_active_index < 0)
    {
        saved_active_index = 0;
    }
    state.active_index = saved_active_index;
    const bool saved = workspace::SaveTabs(path, state) && saved_all;

    for (const auto& search_tab : search_tabs_)
    {
        if (!search_tab.cache_file.empty())
        {
            referenced_files.insert(search_tab.cache_file);
        }
        if (!search_tab.compare_cache_file.empty())
        {
            referenced_files.insert(search_tab.compare_cache_file);
        }
    }
    for (const wchar_t* pattern_name : {L"search_*.tsv", L"compare_*.tsv"})
    {
        std::wstring pattern = util::JoinPath(folder, pattern_name);
        WIN32_FIND_DATAW data = {};
        HANDLE find = FindFirstFileW(pattern.c_str(), &data);
        if (find == INVALID_HANDLE_VALUE)
        {
            continue;
        }
        do
        {
            if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
            {
                continue;
            }
            if (referenced_files.find(data.cFileName) != referenced_files.end())
            {
                continue;
            }
            std::wstring stale = util::JoinPath(folder, data.cFileName);
            DeleteFileW(stale.c_str());
        } while (FindNextFileW(find, &data) != 0);
        FindClose(find);
    }
    return saved;
}

std::wstring MainWindow::Impl::CommentsPath() const
{
    std::wstring folder = util::GetAppDataFolder();
    if (folder.empty())
    {
        return L"";
    }
    const std::wstring path = util::JoinPath(folder, L"comments.jsonc");
    // keep using the previous file until the next save renames it
    const std::wstring legacy = util::JoinPath(folder, L"comments.json");
    return IsFilePath(path) || !IsFilePath(legacy) ? path : legacy;
}

std::wstring MainWindow::Impl::CommentKeyPath(const RegistryNode& node) const
{
    const std::wstring path = registry_path::Build(node);
    return registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty()
               ? L"\\\\" + StripMachinePrefix(remote_machine_) + L"\\" + path
               : path;
}

bool MainWindow::Impl::SaveComments() const
{
    return !comments_unreadable_ && value_comments_.Save(CommentsPath());
}

bool MainWindow::Impl::ImportCommentsFromFile(const std::wstring& path)
{
    std::wstring content;
    std::vector<changes::CommentRule> rules;
    if (!util::ReadTextFile(path, &content, nullptr, util::kMaxCommentFileBytes) ||
        !changes::ParseComments(content, &rules))
    {
        return false;
    }
    value_comments_.Merge(rules);
    if (!SaveComments())
    {
        ui::ShowError(hwnd_, L"Comments were imported but couldn't be saved.");
    }
    RefreshValueListComments();
    return true;
}

bool MainWindow::Impl::ExportCommentsToFile(const std::wstring& path) const
{
    return value_comments_.Save(path);
}

void MainWindow::Impl::RefreshValueListComments()
{
    if (!browse_.current_node())
    {
        return;
    }
    const std::wstring path = CommentKeyPath(*browse_.current_node());
    bool changed = false;
    for (auto& row : browse_.values().rows())
    {
        std::wstring display;
        if (row.kind == rowkind::kValue)
        {
            display =
                FormatCommentDisplay(changes::ResolveComment(value_comments_, default_comments_, {path, row.extra, row.value_type, row.value_data_size})
                                         .text);
        }
        else if (row.kind == rowkind::kKey && !row.extra.empty())
        {
            display = FormatCommentDisplay(
                changes::ResolveComment(value_comments_, default_comments_, {registry_path::JoinSubkey(path, row.extra), {}, 0, 0, true})
                    .text
            );
        }
        if (row.comment != display)
        {
            row.comment = std::move(display);
            changed = true;
        }
    }
    if (browse_.columns().sort_column == kValueColComment)
    {
        SortValueRows(&browse_.values().rows(), browse_.columns().sort_column, browse_.columns().sort_ascending);
        changed = true;
    }
    if (changed)
    {
        browse_.values().InvalidateFilterCache();
    }
    if (browse_.values().HasFilter())
    {
        browse_.values().RebuildFilter();
    }
    else if (changed && browse_.values().hwnd())
    {
        RedrawWindow(browse_.values().hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
    }
}

bool MainWindow::Impl::EditComments(const std::vector<changes::CommentTarget>& targets)
{
    if (targets.empty())
    {
        return false;
    }
    const changes::ValueComments none;
    const changes::CommentTarget& first = targets.front();
    const changes::ResolvedComment shown = changes::ResolveComment(value_comments_, default_comments_, first);
    bool same_text = true;
    bool same_type = true;
    bool same_size = true;
    bool can_restore = false;
    for (const changes::CommentTarget& target : targets)
    {
        const changes::ResolvedComment resolved = changes::ResolveComment(value_comments_, default_comments_, target);
        same_text = same_text && resolved.text == shown.text;
        same_type = same_type && target.type == first.type;
        same_size = same_size && target.data_size == first.data_size;
        can_restore =
            can_restore || (resolved.source == changes::CommentSource::kUser && default_comments_.Match(target));
    }

    const bool edits_rule = same_text && shown.source == changes::CommentSource::kUser;
    const changes::CommentRule value_rule = changes::ValueRule(first);
    const bool broad =
        edits_rule && (shown.rule.key_scope != value_rule.key_scope || shown.rule.type != value_rule.type ||
                       shown.rule.data_size || !util::EqualsInsensitive(shown.rule.key_path, value_rule.key_path));
    editors::CommentRequest request;
    request.key = first.key;
    request.text = same_text ? shown.text : std::wstring();
    request.scope.key_path = first.path;
    if (broad)
    {
        request.scope.rule = true;
        request.scope.same_type = shown.rule.type.has_value();
        request.scope.same_size = shown.rule.data_size.has_value();
        request.scope.in_key = shown.rule.key_scope != changes::CommentKeyScope::kAny;
        request.scope.include_subkeys = shown.rule.key_scope == changes::CommentKeyScope::kRecursive;
        if (request.scope.in_key)
        {
            request.scope.key_path = shown.rule.key_path;
        }
    }
    request.name = L"\"" + (first.name.empty() ? std::wstring(L"(Default)") : first.name) + L"\"";
    request.type = same_type ? value_format::TypeName(first.type) : L"Different types";
    request.size = same_size ? std::to_wstring(first.data_size) + (first.data_size == 1 ? L" byte" : L" bytes")
                             : L"Different lengths";
    request.multiple = targets.size() > 1;
    request.can_restore = can_restore;
    editors::CommentResult result;
    if (!editors::EditComment(hwnd_, request, &result))
    {
        return false;
    }

    const std::wstring text = util::IsBlank(result.text) ? std::wstring() : std::move(result.text);
    for (const changes::CommentTarget& target : targets)
    {
        if (result.restore_default)
        {
            const changes::ResolvedComment user = changes::ResolveComment(value_comments_, none, target);
            if (user.source == changes::CommentSource::kUser)
            {
                value_comments_.Erase(user.rule);
            }
            continue;
        }
        changes::CommentRule rule = changes::ValueRule(target);
        if (result.scope.rule)
        {
            value_comments_.Erase(rule);
            if (broad && targets.size() == 1)
            {
                value_comments_.Erase(shown.rule);
            }
            rule.type = result.scope.same_type ? std::optional<DWORD>(target.type) : std::nullopt;
            rule.data_size = result.scope.same_size ? std::optional<uint64_t>(target.data_size) : std::nullopt;
            rule.key_path = result.scope.in_key ? util::TrimWhitespace(result.scope.key_path) : std::wstring();
            rule.key_scope = rule.key_path.empty()          ? changes::CommentKeyScope::kAny
                             : result.scope.include_subkeys ? changes::CommentKeyScope::kRecursive
                                                            : changes::CommentKeyScope::kExact;
        }
        rule.text = text;
        if (text.empty() && !default_comments_.Match(target))
        {
            value_comments_.Erase(rule);
        }
        else
        {
            value_comments_.Set(std::move(rule));
        }
    }
    if (!SaveComments())
    {
        ui::ShowError(hwnd_, L"The comment couldn't be saved.");
    }
    RefreshValueListComments();
    return true;
}

void MainWindow::Impl::LoadSettings()
{
    workspace::Settings settings;
    settings.clear_history_on_exit = clear_history_on_exit_;
    settings.clear_tabs_on_exit = clear_tabs_on_exit_;
    settings.show_toolbar = show_toolbar_;
    settings.show_address_bar = show_address_bar_;
    settings.show_filter_bar = show_filter_bar_;
    settings.show_tab_control = show_tab_control_;
    settings.show_tree = show_tree_;
    settings.show_history = show_history_;
    settings.show_status_bar = show_status_bar_;
    settings.show_keys_in_list = show_keys_in_list_;
    settings.show_simulated_keys = show_simulated_keys_;
    settings.show_extra_hives = show_extra_hives_;
    settings.show_value_grid = show_value_grid_;
    settings.save_tree_state = save_tree_state_;
    settings.save_tab_kinds = save_tab_kinds_;
    settings.save_tabs = save_tab_kinds_ != 0;
    settings.always_run_as_admin = always_run_as_admin_;
    settings.always_run_as_system = always_run_as_system_;
    settings.always_run_as_trustedinstaller = always_run_as_trustedinstaller_;
    settings.always_on_top = always_on_top_;
    settings.single_instance = single_instance_;
    settings.read_only = read_only_;
    settings.auto_check_updates = auto_check_updates_;
    settings.default_reset_enabled = default_reset_enabled_;
    settings.window_x = window_x_;
    settings.window_y = window_y_;
    settings.window_width = window_width_;
    settings.window_height = window_height_;
    settings.window_maximized = window_maximized_;
    settings.tree_width = tree_width_;
    settings.history_height = history_height_;
    settings.theme_preset = active_theme_preset_;
    settings.icon_set = icon_set_;
    settings.use_custom_font = use_custom_font_;
    settings.font_face = custom_font_.lfFaceName;
    settings.font_size = appearance::FontPointSize(custom_font_, 9);
    settings.font_weight = custom_font_.lfWeight;
    settings.font_italic = custom_font_.lfItalic != FALSE;
    settings.value_column_widths = browse_.columns().saved_widths;
    settings.value_column_visible = browse_.columns().saved_visible;

    if (!workspace::LoadSettings(SettingsPath(), &settings))
    {
        return;
    }

    clear_history_on_exit_ = settings.clear_history_on_exit;
    clear_tabs_on_exit_ = settings.clear_tabs_on_exit;
    show_toolbar_ = settings.show_toolbar;
    show_address_bar_ = settings.show_address_bar;
    show_filter_bar_ = settings.show_filter_bar;
    show_tab_control_ = settings.show_tab_control;
    show_tree_ = settings.show_tree;
    show_history_ = settings.show_history;
    show_status_bar_ = settings.show_status_bar;
    show_keys_in_list_ = settings.show_keys_in_list;
    show_simulated_keys_ = settings.show_simulated_keys;
    show_extra_hives_ = settings.show_extra_hives;
    show_value_grid_ = settings.show_value_grid;
    save_tree_state_ = settings.save_tree_state;
    save_tab_kinds_ = settings.save_tab_kinds;
    always_run_as_admin_ = settings.always_run_as_admin;
    always_run_as_system_ = settings.always_run_as_system;
    always_run_as_trustedinstaller_ = settings.always_run_as_trustedinstaller;
    always_on_top_ = settings.always_on_top;
    single_instance_ = settings.single_instance;
    read_only_ = settings.read_only;
    auto_check_updates_ = settings.auto_check_updates;
    default_reset_enabled_ = settings.default_reset_enabled;
    window_placement_loaded_ = settings.window_placement_present;
    window_x_ = settings.window_x;
    window_y_ = settings.window_y;
    window_width_ = settings.window_width;
    window_height_ = settings.window_height;
    window_maximized_ = settings.window_maximized;
    tree_width_ = settings.tree_width;
    history_height_ = settings.history_height;
    theme_mode_ = ParseThemeMode(settings.theme_mode);
    active_theme_preset_ = std::move(settings.theme_preset);
    icon_set_ = IsKnownIconSetName(settings.icon_set) ? std::move(settings.icon_set) : kIconSetPhosphor;
    use_custom_font_ = settings.use_custom_font;
    if (!settings.font_face.empty())
    {
        wcsncpy_s(custom_font_.lfFaceName, settings.font_face.c_str(), _TRUNCATE);
    }
    if (settings.font_size > 0)
    {
        custom_font_.lfHeight = appearance::FontHeight(settings.font_size);
    }
    custom_font_.lfWeight = settings.font_weight;
    custom_font_.lfItalic = settings.font_italic ? TRUE : FALSE;
    recent_trace_paths_.Replace(std::move(settings.recent_traces));
    recent_default_paths_.Replace(std::move(settings.recent_defaults));
    browse_.columns().saved_widths = std::move(settings.value_column_widths);
    browse_.columns().saved_visible = std::move(settings.value_column_visible);
    browse_.columns().saved = !browse_.columns().saved_widths.empty() || !browse_.columns().saved_visible.empty();
    if (!save_tree_state_)
    {
        saved_tree_state_.Clear();
    }
}
void MainWindow::Impl::SaveSettings() const
{
    workspace::Settings settings;
    settings.clear_history_on_exit = clear_history_on_exit_;
    settings.clear_tabs_on_exit = clear_tabs_on_exit_;
    settings.show_toolbar = show_toolbar_;
    settings.show_address_bar = show_address_bar_;
    settings.show_filter_bar = show_filter_bar_;
    settings.show_tab_control = show_tab_control_;
    settings.show_tree = show_tree_;
    settings.show_history = show_history_;
    settings.show_status_bar = show_status_bar_;
    settings.show_keys_in_list = show_keys_in_list_;
    settings.show_simulated_keys = show_simulated_keys_;
    settings.show_extra_hives = show_extra_hives_;
    settings.show_value_grid = show_value_grid_;
    settings.save_tree_state = save_tree_state_;
    settings.save_tab_kinds = save_tab_kinds_;
    settings.save_tabs = save_tab_kinds_ != 0;
    settings.always_run_as_admin = always_run_as_admin_;
    settings.always_run_as_system = always_run_as_system_;
    settings.always_run_as_trustedinstaller = always_run_as_trustedinstaller_;
    settings.always_on_top = always_on_top_;
    settings.single_instance = single_instance_;
    settings.read_only = read_only_;
    settings.auto_check_updates = auto_check_updates_;
    settings.default_reset_enabled = default_reset_enabled_;
    settings.window_x = window_x_;
    settings.window_y = window_y_;
    settings.window_width = window_width_;
    settings.window_height = window_height_;
    settings.window_maximized = window_maximized_;
    if (hwnd_ && IsWindow(hwnd_))
    {
        WINDOWPLACEMENT placement = {};
        placement.length = sizeof(placement);
        if (GetWindowPlacement(hwnd_, &placement))
        {
            const RECT& normal = placement.rcNormalPosition;
            const int width = normal.right - normal.left;
            const int height = normal.bottom - normal.top;
            if (width > 0 && height > 0)
            {
                settings.window_x = normal.left;
                settings.window_y = normal.top;
                settings.window_width = width;
                settings.window_height = height;
            }
            settings.window_maximized = placement.showCmd == SW_SHOWMAXIMIZED;
        }
    }
    settings.tree_width = tree_width_;
    settings.history_height = history_height_;
    settings.theme_mode = ThemeModeName(theme_mode_);
    settings.theme_preset = active_theme_preset_;
    settings.icon_set = IsKnownIconSetName(icon_set_) ? icon_set_ : kIconSetPhosphor;
    settings.use_custom_font = use_custom_font_;
    settings.font_face = custom_font_.lfFaceName;
    settings.font_size = appearance::FontPointSize(custom_font_, 9);
    settings.font_weight = custom_font_.lfWeight;
    settings.font_italic = custom_font_.lfItalic != FALSE;
    settings.recent_traces = recent_trace_paths_.items();
    settings.recent_defaults = recent_default_paths_.items();
    settings.value_column_widths = browse_.columns().widths;
    settings.value_column_visible = browse_.columns().visible;
    settings.value_column_widths.resize(browse_.columns().items.size(), 0);
    settings.value_column_visible.resize(browse_.columns().items.size(), true);
    workspace::SaveSettings(SettingsPath(), settings);
}
std::wstring MainWindow::Impl::SettingsPath() const
{
    std::wstring folder = util::GetAppDataFolder();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"settings.ini");
}

std::wstring MainWindow::Impl::TreeStatePath() const
{
    std::wstring folder = CacheFolderPath();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"tree_state.ini");
}

void MainWindow::Impl::LoadTreeState()
{
    saved_tree_state_.Clear();
    if (!save_tree_state_)
    {
        return;
    }
    workspace::LoadTreeState(TreeStatePath(), &saved_tree_state_);
}

void MainWindow::Impl::StartTreeStateWorker()
{
    if (!save_tree_state_ || tree_state_saver_.running())
    {
        return;
    }
    tree_state_saver_.Start(std::chrono::seconds(2), [this](workspace::TreeState state) {
        SaveTreeStateFile(state.selected_path, state.expanded_paths);
    });
}

void MainWindow::Impl::StopTreeStateWorker()
{
    tree_state_saver_.Stop();
    if (save_tree_state_ && browse_.tree().hwnd() && IsWindow(browse_.tree().hwnd()))
    {
        std::wstring selected;
        std::vector<std::wstring> expanded;
        CaptureTreeState(&selected, &expanded);
        SaveTreeStateFile(selected, expanded);
    }
}

} // namespace regkit
