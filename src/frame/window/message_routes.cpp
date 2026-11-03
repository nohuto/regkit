// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

#include "registry/hive_files.h"
#include "win32/text_transform.h"
#include "win32/translation.h"
#include "trace/trace_paths.h"

#include <win32ui/focus.h>
#include <win32ui/gdi/buffered_paint.h>
#include <win32ui/system/paths.h>
#include <win32ui/system/single_instance.h>
#include <win32ui/window.h>
#include <win32ui/window_metrics.h>

#include <utility>

namespace regkit
{

using namespace window_detail;

namespace
{

struct HandoffPayload : win32ui::PostedPayload
{
    std::wstring target;
};

bool AcceptHandoffFile(HWND owner, const std::wstring& path)
{
    if (!util::IsProcessPrivileged() || !util::IsUacEnabled())
    {
        return true;
    }
    bool local = win32ui::OnFixedDrive(path);
    DWORD attributes = FILE_ATTRIBUTE_DIRECTORY;
    for (std::wstring component = path; local && component.size() > 3; component.resize(component.find_last_of(L"\\/")))
    {
        const DWORD component_attributes = GetFileAttributesW(component.c_str());
        local = component_attributes != INVALID_FILE_ATTRIBUTES &&
                (component_attributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0;
        attributes = component.size() == path.size() ? component_attributes : attributes;
    }
    if (!local || (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        ui::ShowWarning(
            owner,
            util::Tr(L"RegKit is running with elevated rights and only opens local files handed to it by another instance.")
        );
        return false;
    }
    return ui::PromptKeyChoice(owner, util::Tr(L"Another RegKit instance asked this elevated window to open a .reg file.\n\nOpen it?"), path, util::Tr(L"Open .reg File"), util::Tr(L"Open"), L"", util::Tr(L"Cancel")) == IDYES;
}

// entries added during startup replace loaded entries for the same source
template <typename Active>
bool MergeAdded(std::vector<Active>* loaded, std::vector<Active>&& added)
{
    for (Active& entry : added)
    {
        std::erase_if(*loaded, [&](const Active& old) { return util::EqualsInsensitive(old.source_path, entry.source_path); });
        loaded->push_back(std::move(entry));
    }
    return !added.empty();
}

}

LRESULT MainWindow::Impl::HandleMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    std::optional<LRESULT> result;
    switch (frame::ClassifyMessage(message))
    {
    case frame::MessageArea::kLifecycle:
        result = HandleLifecycleMessage(message, wparam, lparam);
        break;
    case frame::MessageArea::kLayoutInput:
        result = HandleLayoutInputMessage(message, wparam, lparam);
        break;
    case frame::MessageArea::kWorker:
        result = HandleWorkerMessage(message, wparam, lparam);
        break;
    case frame::MessageArea::kExternal:
        result = HandleExternalMessage(message, wparam, lparam);
        break;
    case frame::MessageArea::kAppearance:
        result = HandleAppearanceMessage(message, wparam, lparam);
        break;
    case frame::MessageArea::kBrowse:
        result = HandleBrowseMessage(message, wparam, lparam);
        break;
    case frame::MessageArea::kUnknown:
        break;
    }
    return result ? *result : DefWindowProcW(hwnd_, message, wparam, lparam);
}

std::optional<LRESULT> MainWindow::Impl::HandleLifecycleMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    switch (message)
    {
    case WM_CREATE:
        return OnCreate() ? 0 : -1;
    case WM_DESTROY:
        OnDestroy();
        PostQuitMessage(0);
        return 0;
    case WM_NCDESTROY:
        ReleaseDrawingResources();
        return std::nullopt;
    case WM_NCPAINT:
        {
            LRESULT result = DefWindowProcW(hwnd_, message, wparam, lparam);
            menu_bar_.PaintSeparator();
            return result;
        }
    case WM_NCACTIVATE:
        {
            LRESULT result = DefWindowProcW(hwnd_, message, wparam, lparam);
            menu_bar_.PaintSeparator();
            return result;
        }
    case WM_GETMINMAXINFO:
        {
            auto* info = reinterpret_cast<MINMAXINFO*>(lparam);
            if (info)
            {
                const UINT dpi = win32ui::DpiForWindow(hwnd_);
                win32ui::RaiseMinTrackSize(hwnd_, info, {400, 200}, dpi, false);
                // room above/below panels plus minimum history height
                RECT window = {};
                if (content_rect_.right > 0 && GetWindowRect(hwnd_, &window))
                {
                    const int panels = TextRowHeight(kPanelHeaderHeight) +
                                       (settings_.show_history && !settings_.simple_mode && !IsSearchTabSelected() ? win32ui::Scale(kMinHistoryHeight + kHistoryGap + kHistorySplitterHeight, dpi) : 0);
                    const int chrome = (window.bottom - window.top) - (content_rect_.bottom - content_rect_.top);
                    info->ptMinTrackSize.y = std::max<LONG>(info->ptMinTrackSize.y, chrome + panels);
                }
            }
            return 0;
        }
    case WM_SIZE:
        OnSize(LOWORD(lparam), HIWORD(lparam));
        return 0;
    case WM_DPICHANGED:
        win32ui::ApplySuggestedRect(hwnd_, reinterpret_cast<const RECT*>(lparam));
        // toolbar first, UpdateUIFont runs the layout, list columns follow in RefreshListView
        ReloadThemeIcons();
        UpdateUIFont();
        return 0;
    case WM_DPICHANGED_AFTERPARENT:
        ReloadThemeIcons();
        UpdateUIFont();
        return 0;
    case WM_ACTIVATE:
        if (const HWND last = win32ui::ActivationFocus(hwnd_, wparam, &last_focus_))
        {
            FocusPane(last);
            return 0;
        }
        if (LOWORD(wparam) == WA_ACTIVE && browse_.tree().hwnd())
        {
            if (last_focus_)
            {
                break;
            }
            int sel = TabCtrl_GetCurSel(tab_);
            if (sel >= 0 && static_cast<size_t>(sel) < tabs_.size() &&
                tabs_[static_cast<size_t>(sel)].kind == TabEntry::Kind::kRegistry)
            {
                SetFocus(browse_.tree().hwnd());
            }
        }
        break;
    case WM_CLOSE:
        {
            if (!std::exchange(close_confirmed_, false) && !ConfirmCloseAllTabs())
            {
                restart_on_close_ = false;
                reset_settings_on_close_ = false;
                cache_clear_on_close_.reset();
                updates_.DropSetup();
                return 0;
            }
            // stop services only after all tabs agree to close
            for (TabEntry& entry : tabs_)
            {
                if (entry.session)
                {
                    OfferRemoteServiceRestore(*entry.session, false);
                }
            }
            if (reset_settings_on_close_)
            {
                if (!RestartAfterSettingsReset())
                {
                    reset_settings_on_close_ = false;
                    return 0;
                }
            }
            else if (cache_clear_on_close_)
            {
                if (!RestartAfterCacheClear(*cache_clear_on_close_))
                {
                    cache_clear_on_close_.reset();
                    return 0;
                }
            }
            else if (restart_on_close_)
            {
                if (!RestartCurrentInstance())
                {
                    restart_on_close_ = false;
                    return 0;
                }
            }
            else if (updates_.SetupPending())
            {
                // keep RegKit running until setup sends WM_ENDSESSION
                if (PrepareSessionHandover() && updates_.StartSetup())
                {
                    EnableWindow(hwnd_, FALSE);
                }
                updates_.DropSetup();
                return 0;
            }
            else
            {
                SaveSettings();
            }
            DestroyWindow(hwnd_);
            return 0;
        }
    case WM_ENDSESSION:
        if (wparam && updates_.SetupRunning())
        {
            restart_on_close_ = true;
            DestroyWindow(hwnd_);
        }
        return 0;
    default:
        return std::nullopt;
    }
    return std::nullopt;
}

std::optional<LRESULT> MainWindow::Impl::HandleLayoutInputMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    (void)wparam;
    switch (message)
    {
    case WM_SETCURSOR:
        {
            if (compare_session_.running() && LOWORD(lparam) == HTCLIENT)
            {
                SetCursor(LoadCursorW(nullptr, IDC_APPSTARTING));
                return TRUE;
            }
            break;
        }
    default:
        return std::nullopt;
    }
    return std::nullopt;
}

std::optional<LRESULT> MainWindow::Impl::HandleWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    switch (message)
    {
    case frame::message_id::kSearchResults:
    case frame::message_id::kSearchPreviewRequest:
    case frame::message_id::kSearchPreviewReady:
    case frame::message_id::kSearchSortReady:
    case frame::message_id::kSearchTabLoadReady:
    case frame::message_id::kSearchProgress:
    case frame::message_id::kSearchFailed:
    case frame::message_id::kReplaceReady:
    case frame::message_id::kCompareReady:
        return HandleSearchWorkerMessage(message, wparam, lparam);
    case frame::message_id::kLoadTraces:
    case frame::message_id::kLoadDefaults:
    case frame::message_id::kTraceLoadReady:
    case frame::message_id::kDefaultLoadReady:
    case frame::message_id::kDeferredStartup:
    case frame::message_id::kStartupCacheReady:
    case frame::message_id::kUpdateCheckReady:
    case frame::message_id::kShowFileErrors:
        return HandleLoadWorkerMessage(message, wparam, lparam);
    case frame::message_id::kRegFileLoadReady:
        return HandleRegFileWorkerMessage(message, wparam, lparam);
    case frame::message_id::kTraceParseBatch:
        return HandleTraceWorkerMessage(message, wparam, lparam);
    case frame::message_id::kDefaultParseBatch:
        return HandleDefaultWorkerMessage(message, wparam, lparam);
    case frame::message_id::kValueListReady:
    case frame::message_id::kValuePreviewReady:
    case frame::message_id::kValuePreviewRequest:
        return HandleValueWorkerMessage(message, wparam, lparam);
    default:
        return std::nullopt;
    }
}

MainWindow::Impl::SearchTab* MainWindow::Impl::SearchTabByGeneration(uint64_t generation)
{
    for (auto& tab : search_tabs_)
    {
        if (tab.run && tab.generation == generation)
        {
            return &tab;
        }
    }
    return nullptr;
}

MainWindow::Impl::SearchTab* MainWindow::Impl::ShownSearchTab()
{
    const int index = tab_ ? SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) : -1;
    return index >= 0 && static_cast<size_t>(index) < search_tabs_.size() ? &search_tabs_[static_cast<size_t>(index)] : nullptr;
}

void MainWindow::Impl::FinishSearch(SearchTab* tab)
{
    tab->run->session.Join();
    tab->duration_ms = std::max<uint64_t>(GetTickCount64() - tab->run->start_tick, 1);
    tab->run.reset();
    if (tab->sort_dirty)
    {
        SortSearchTabResults(tab);
    }
    if (ShownSearchTab() == tab)
    {
        search_last_refresh_tick_ = GetTickCount64();
        UpdateSearchResultsView();
    }
    ApplyViewVisibility();
    UpdateStatus();
}

std::optional<LRESULT> MainWindow::Impl::HandleSearchWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    switch (message)
    {
    case frame::message_id::kSearchResults:
        {
            SearchTab* tab = SearchTabByGeneration(static_cast<uint64_t>(wparam));
            if (!tab)
            {
                return 0;
            }
            SearchRun& run = *tab->run;
            run.posted.store(false);
            const uint64_t start_tick = GetTickCount64();
            bool appended = false;
            bool more_pending = false;
            bool producer_done = false;
            for (;;)
            {
                std::vector<search::Result> rows;
                {
                    std::lock_guard<std::mutex> lock(run.mutex);
                    producer_done = run.producer_done;
                    if (run.batches.empty())
                    {
                        break;
                    }
                    rows = std::move(run.batches.front());
                    run.batches.pop_front();
                    run.pending_rows -= rows.size();
                    more_pending = !run.batches.empty();
                }
                run.space.notify_all();
                for (auto& row : rows)
                {
                    row.row_id = tab->next_row_id++;
                }
                tab->results.insert(tab->results.end(), std::make_move_iterator(rows.begin()), std::make_move_iterator(rows.end()));
                appended = true;
                if (GetTickCount64() - start_tick >= kSearchResultsMaxMs)
                {
                    break;
                }
            }
            tab->sort_dirty = tab->sort_dirty || appended;
            if (more_pending && !run.posted.exchange(true) && !PostMessageW(hwnd_, frame::message_id::kSearchResults, wparam, 0))
            {
                run.posted.store(false);
            }
            if (appended && ShownSearchTab() == tab)
            {
                const uint64_t now = GetTickCount64();
                if (now - search_last_refresh_tick_ >= kSearchResultsRefreshMs)
                {
                    search_last_refresh_tick_ = now;
                    UpdateSearchResultsView();
                    UpdateStatus();
                }
            }
            if (!more_pending && producer_done)
            {
                FinishSearch(tab);
            }
            return 0;
        }
    case frame::message_id::kSearchPreviewRequest:
        {
            search_preview_request_posted_ = false;
            if (!search_results_list_)
            {
                return 0;
            }
            const int top = ListView_GetTopIndex(search_results_list_);
            const int page = ListView_GetCountPerPage(search_results_list_);
            const int count = ListView_GetItemCount(search_results_list_);
            QueueSearchPreviews(std::max(0, top), std::min(count - 1, top + std::max(page, 1)));
            return 0;
        }
    case frame::message_id::kSearchPreviewReady:
        {
            auto owned = win32ui::TakePayload<SearchPreviewPayload>(lparam);
            if (!owned)
            {
                return 0;
            }
            if (owned->tab_index < 0 || static_cast<size_t>(owned->tab_index) >= search_tabs_.size())
            {
                return 0;
            }
            SearchTab& tab = search_tabs_[static_cast<size_t>(owned->tab_index)];
            if (tab.generation != owned->generation)
            {
                return 0;
            }
            int first = -1;
            int last = -1;
            for (auto& item : owned->items)
            {
                int index = -1;
                if (item.index >= 0 && static_cast<size_t>(item.index) < tab.results.size() &&
                    tab.results[static_cast<size_t>(item.index)].row_id == item.row_id)
                {
                    index = item.index;
                }
                else
                {
                    for (size_t i = 0; i < tab.results.size(); ++i)
                    {
                        if (tab.results[i].row_id == item.row_id)
                        {
                            index = static_cast<int>(i);
                            break;
                        }
                    }
                }
                if (index < 0)
                {
                    continue;
                }
                search::Result& row = tab.results[static_cast<size_t>(index)];
                row.data_text = std::move(item.preview);
                row.type = item.type;
                row.data_size = item.data_size;
                row.data_state = search::DataState::kLoaded;
                first = first < 0 ? index : std::min(first, index);
                last = std::max(last, index);
            }
            if (first >= 0 && search_results_list_ && SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) == owned->tab_index)
            {
                ListView_RedrawItems(search_results_list_, first, last);
            }
            return 0;
        }
    case frame::message_id::kSearchSortReady:
        {
            auto owned = win32ui::TakePayload<SearchSortPayload>(lparam);
            if (!owned)
            {
                return 0;
            }
            if (owned->tab_index < 0 || static_cast<size_t>(owned->tab_index) >= search_tabs_.size())
            {
                return 0;
            }
            SearchTab& tab = search_tabs_[static_cast<size_t>(owned->tab_index)];
            if (tab.generation != owned->generation || tab.results.size() != owned->rows.size() || tab.sort_revision != owned->sort_revision)
            {
                return 0;
            }
            ApplySortedSearchRows(&tab, std::move(owned->rows), SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) == owned->tab_index);
            return 0;
        }
    case frame::message_id::kSearchTabLoadReady:
        ApplySearchTabLoad(win32ui::TakePayload<SearchTabLoadPayload>(lparam));
        return 0;
    case frame::message_id::kSearchProgress:
        if (SearchTab* tab = SearchTabByGeneration(static_cast<uint64_t>(wparam)))
        {
            tab->run->progress_posted.store(false);
            if (ShownSearchTab() == tab)
            {
                UpdateStatus();
            }
        }
        return 0;
    case frame::message_id::kSearchFailed:
        if (SearchTab* tab = SearchTabByGeneration(static_cast<uint64_t>(wparam)))
        {
            CancelSearch(tab);
            const std::wstring detail = search::regex::StatusText(static_cast<search::regex::Status>(lparam));
            ui::ShowError(hwnd_, detail.empty() ? std::wstring(util::Tr(L"The find text isn't a valid regular expression.")) : detail);
        }
        return 0;
    case frame::message_id::kReplaceReady:
        ApplyReplacePayload(win32ui::TakePayload<ReplacePayload>(lparam));
        return 0;
    case frame::message_id::kCompareReady:
        ApplyComparePayload(win32ui::TakePayload<ComparePayload>(lparam));
        return 0;
    default:
        return std::nullopt;
    }
}

std::optional<LRESULT> MainWindow::Impl::HandleLoadWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    (void)wparam;
    switch (message)
    {
    case frame::message_id::kLoadTraces:
        StartTraceLoadWorker();
        return 0;
    case frame::message_id::kLoadDefaults:
        StartDefaultLoadWorker();
        return 0;
    case frame::message_id::kTraceLoadReady:
        {
            auto owned = win32ui::TakePayload<TraceLoadPayload>(lparam);
            if (!owned)
            {
                return 0;
            }
            if (!trace_load_session_.IsCurrent(owned->generation))
            {
                return 0;
            }
            trace_load_session_.Join();
            for (const ActiveTrace& added : active_traces_)
            {
                if (const auto cached = trace_selection_cache_.find(ToLower(added.source_path)); cached != trace_selection_cache_.end())
                {
                    owned->selection_cache[cached->first] = cached->second;
                }
            }
            const bool edited = MergeAdded(&owned->traces, std::move(active_traces_));
            active_traces_ = std::move(owned->traces);
            trace_selection_cache_ = std::move(owned->selection_cache);
            if (edited)
            {
                SaveActiveTraces();
            }
            RefreshTreeSelection();
            UpdateValueListForNode(browse_.current_node());
            return 0;
        }
    case frame::message_id::kDefaultLoadReady:
        {
            auto owned = win32ui::TakePayload<DefaultLoadPayload>(lparam);
            if (!owned)
            {
                return 0;
            }
            if (!default_load_session_.IsCurrent(owned->generation))
            {
                return 0;
            }
            default_load_session_.Join();
            const bool edited = MergeAdded(&owned->defaults, std::move(active_defaults_));
            active_defaults_ = std::move(owned->defaults);
            if (edited)
            {
                SaveActiveDefaults();
            }
            UpdateValueListForNode(browse_.current_node());
            return 0;
        }
    case frame::message_id::kDeferredStartup:
        RunDeferredStartup();
        return 0;
    case frame::message_id::kStartupCacheReady:
        ApplyStartupCachePayload(win32ui::TakePayload<StartupCachePayload>(lparam));
        return 0;
    case frame::message_id::kShowFileErrors:
        ShowFileErrors();
        return 0;
    case frame::message_id::kUpdateCheckReady:
        {
            updates_.Apply(win32ui::TakePayload<frame::UpdateCheckPayload>(lparam).get());
            return 0;
        }
    default:
        return std::nullopt;
    }
}

std::optional<LRESULT> MainWindow::Impl::HandleRegFileWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    (void)wparam;
    switch (message)
    {
    case frame::message_id::kRegFileLoadReady:
        {
            auto owned = win32ui::TakePayload<RegFileParsePayload>(lparam);
            if (!owned)
            {
                return 0;
            }
            auto session_it = reg_file_parse_sessions_.find(owned->source_lower);
            if (session_it == reg_file_parse_sessions_.end())
            {
                return 0;
            }
            if (!session_it->second || !session_it->second->work.IsCurrent(owned->generation))
            {
                return 0;
            }
            session_it->second->work.Join();
            reg_file_parse_sessions_.erase(session_it);
            if (owned->cancelled)
            {
                return 0;
            }

            int tab_index = -1;
            for (size_t i = 0; i < tabs_.size(); ++i)
            {
                const TabEntry& entry = tabs_[i];
                if (entry.kind != TabEntry::Kind::kRegFile)
                {
                    continue;
                }
                if (entry.reg_file_session_key == owned->source_lower)
                {
                    tab_index = static_cast<int>(i);
                    break;
                }
            }
            if (tab_index < 0 || static_cast<size_t>(tab_index) >= tabs_.size())
            {
                return 0;
            }
            TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
            entry.reg_file_loading = false;
            if (!owned->error.empty())
            {
                ui::ShowError(hwnd_, owned->error.c_str());
                UpdateStatus();
                return 0;
            }
            if (entry.reg_file_dirty)
            {
                UpdateStatus();
                return 0;
            }

            ReleaseRegFileRoots(&entry);
            std::vector<TabEntry::RegFileRoot> roots;
            roots.reserve(owned->roots.size());
            for (auto& parsed : owned->roots)
            {
                if (!parsed.data)
                {
                    continue;
                }
                TabEntry::RegFileRoot root;
                root.name = parsed.name;
                root.data = parsed.data;
                root.root = RegistryStore::RegisterVirtualRoot(root.name, root.data);
                if (root.root)
                {
                    roots.push_back(std::move(root));
                }
            }
            entry.reg_file_roots = std::move(roots);
            entry.reg_file_dirty = false;
            if (tab_ && TabCtrl_GetCurSel(tab_) == tab_index)
            {
                SyncRegFileTabSelection();
                ApplyViewVisibility();
                UpdateStatus();
            }
            return 0;
        }
    default:
        return std::nullopt;
    }
}

std::optional<LRESULT> MainWindow::Impl::HandleTraceWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    (void)wparam;
    switch (message)
    {
    case frame::message_id::kTraceParseBatch:
        {
            auto owned = win32ui::TakePayload<ParseBatch>(lparam);
            if (!owned)
            {
                return 0;
            }
            auto it = trace_parse_sessions_.find(owned->source_lower);
            if (it == trace_parse_sessions_.end())
            {
                return 0;
            }
            TraceParseSession* session = it->second.get();
            if (!session || !session->work.IsCurrent(owned->generation))
            {
                return 0;
            }
            const bool touches_current =
                browse_.current_node() &&
                owned->affected_keys.find(TracePathLowerForNode(*browse_.current_node())) != owned->affected_keys.end();
            if (session->dialog && IsWindow(session->dialog) && !owned->entries.empty())
            {
                TraceDialogPostEntries(session->dialog, std::move(owned->entries));
            }
            if (session->added_to_active && touches_current && browse_.current_node())
            {
                uint64_t now = GetTickCount64();
                if (owned->done || (now - last_trace_refresh_tick_) >= 100)
                {
                    last_trace_refresh_tick_ = now;
                    UpdateValueListForNode(browse_.current_node());
                }
            }
            if (owned->done)
            {
                session->parsing_done = true;
                if (session->data)
                {
                    std::unique_lock<std::shared_mutex> data_lock(*session->data->mutex);
                    trace::Selection normalized = session->selection;
                    trace::NormalizeSelection(*session->data, &normalized);
                    session->selection = normalized;
                    trace_selection_cache_[session->source_lower] = normalized;
                    if (session->added_to_active)
                    {
                        for (auto& trace : active_traces_)
                        {
                            if (EqualsInsensitive(trace.source_path, session->source_path))
                            {
                                trace.selection = std::make_shared<trace::Selection>(normalized);
                                break;
                            }
                        }
                    }
                }
                if (!owned->error.empty())
                {
                    HWND error_owner = session->dialog && IsWindow(session->dialog) ? session->dialog : hwnd_;
                    ui::ShowError(error_owner, owned->error.c_str());
                    if (session->dialog && IsWindow(session->dialog))
                    {
                        PostMessageW(session->dialog, WM_CLOSE, 0, 0);
                    }
                    if (session->added_to_active)
                    {
                        active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(), [&](const ActiveTrace& trace) {
                                                 return EqualsInsensitive(trace.source_path, session->source_path);
                                             }),
                                             active_traces_.end());
                        trace_selection_cache_.erase(session->source_lower);
                        SaveActiveTraces();
                        SaveTraceSettings();
                        RefreshTreeSelection();
                        UpdateValueListForNode(browse_.current_node());
                        SaveSettings();
                        session->added_to_active = false;
                    }
                }
                else if (session->dialog && IsWindow(session->dialog))
                {
                    TraceDialogPostDone(session->dialog, true);
                }
                if (session->added_to_active && browse_.current_node())
                {
                    UpdateValueListForNode(browse_.current_node());
                }
                session->work.Join();
                if (!session->dialog || !IsWindow(session->dialog))
                {
                    trace_parse_sessions_.erase(it);
                }
            }
            return 0;
        }
    default:
        return std::nullopt;
    }
}

std::optional<LRESULT> MainWindow::Impl::HandleDefaultWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    (void)wparam;
    switch (message)
    {
    case frame::message_id::kDefaultParseBatch:
        {
            auto owned = win32ui::TakePayload<ParseBatch>(lparam);
            if (!owned)
            {
                return 0;
            }
            auto it = default_parse_sessions_.find(owned->source_lower);
            if (it == default_parse_sessions_.end())
            {
                return 0;
            }
            DefaultParseSession* session = it->second.get();
            if (!session || !session->work.IsCurrent(owned->generation))
            {
                return 0;
            }
            bool touches_current = false;
            if (browse_.current_node())
            {
                std::wstring path = SourceLookupPath(*browse_.current_node());
                std::wstring normalized = trace::NormalizeKeyPathBasic(path);
                if (normalized.empty())
                {
                    normalized = path;
                }
                touches_current = owned->affected_keys.find(ToLower(normalized)) != owned->affected_keys.end();
            }
            if (session->dialog && IsWindow(session->dialog) && !owned->entries.empty())
            {
                TraceDialogPostEntries(session->dialog, std::move(owned->entries));
            }
            if (session->added_to_active && touches_current && browse_.current_node())
            {
                uint64_t now = GetTickCount64();
                if (owned->done || (now - last_default_refresh_tick_) >= 100)
                {
                    last_default_refresh_tick_ = now;
                    UpdateValueListForNode(browse_.current_node());
                }
            }
            if (owned->done)
            {
                session->parsing_done = true;
                if (!owned->error.empty())
                {
                    if (session->show_errors)
                    {
                        HWND error_owner = session->dialog && IsWindow(session->dialog) ? session->dialog : hwnd_;
                        ui::ShowError(error_owner, owned->error.c_str());
                    }
                    if (session->dialog && IsWindow(session->dialog))
                    {
                        PostMessageW(session->dialog, WM_CLOSE, 0, 0);
                    }
                    if (session->added_to_active)
                    {
                        active_defaults_.erase(std::remove_if(active_defaults_.begin(), active_defaults_.end(), [&](const ActiveDefault& defaults) {
                                                   return EqualsInsensitive(defaults.source_path, session->source_path);
                                               }),
                                               active_defaults_.end());
                        SaveActiveDefaults();
                        UpdateValueListForNode(browse_.current_node());
                        SaveSettings();
                        session->added_to_active = false;
                    }
                }
                else if (session->dialog && IsWindow(session->dialog))
                {
                    TraceDialogPostDone(session->dialog, true);
                }
                if (session->added_to_active && browse_.current_node())
                {
                    UpdateValueListForNode(browse_.current_node());
                }
                session->work.Join();
                if (!session->dialog || !IsWindow(session->dialog))
                {
                    default_parse_sessions_.erase(it);
                }
            }
            return 0;
        }
    default:
        return std::nullopt;
    }
}

std::optional<LRESULT> MainWindow::Impl::HandleValueWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    (void)wparam;
    switch (message)
    {
    case frame::message_id::kValuePreviewRequest:
        {
            value_preview_request_posted_ = false;
            HWND list = browse_.values().hwnd();
            if (!list)
            {
                return 0;
            }
            const int top = ListView_GetTopIndex(list);
            const int page = ListView_GetCountPerPage(list);
            const int count = ListView_GetItemCount(list);
            QueueValuePreviews(std::max(0, top), std::min(count - 1, top + std::max(page, 1)));
            return 0;
        }
    case frame::message_id::kValuePreviewReady:
        {
            auto owned = win32ui::TakePayload<ValuePreviewPayload>(lparam);
            if (!owned)
            {
                return 0;
            }
            if (owned->generation != value_list_generation_.load())
            {
                return 0;
            }
            int first = -1;
            int last = -1;
            for (const auto& item : owned->items)
            {
                int index = item.index;
                ListRow* row = browse_.values().MutableRowAt(index);
                if (!row || row->extra != item.name)
                {
                    row = nullptr;
                    const int count = static_cast<int>(browse_.values().RowCount());
                    for (int i = 0; i < count; ++i)
                    {
                        ListRow* candidate = browse_.values().MutableRowAt(i);
                        if (candidate && candidate->kind == rowkind::kValue && candidate->extra == item.name)
                        {
                            row = candidate;
                            index = i;
                            break;
                        }
                    }
                }
                if (!row)
                {
                    continue;
                }
                row->data = item.preview;
                row->value_type = item.type;
                row->value_data_size = item.size;
                if (row->type.empty())
                {
                    row->type = value_format::TypeName(item.type);
                }
                row->size_value = item.size;
                row->has_size = true;
                if (row->size.empty() && item.size > 0)
                {
                    row->size = std::to_wstring(item.size);
                }
                row->data_ready = true;
                row->data_preview = item.size > kValuePreviewBytes || item.preview.size() >= kValuePreviewLimit;
                browse_.values().InvalidateFilterCache(row);
                first = first < 0 ? index : std::min(first, index);
                last = std::max(last, index);
            }
            browse_.values().RefreshFilter();
            if (first >= 0 && browse_.values().hwnd())
            {
                ListView_RedrawItems(browse_.values().hwnd(), first, last);
            }
            return 0;
        }
    case frame::message_id::kValueListReady:
        {
            auto owned = win32ui::TakePayload<ValueListPayload>(lparam);
            if (!owned)
            {
                return 0;
            }
            if (owned->generation != value_list_generation_.load())
            {
                return 0;
            }
            HWND list_hwnd = browse_.values().hwnd();
            const bool merge = owned->generation == merge_value_list_generation_ && browse_.current_node();
            if (list_hwnd && !merge)
            {
                SendMessageW(list_hwnd, WM_SETREDRAW, FALSE, 0);
            }
            bool merged = false;
            if (merge)
            {
                ApplyValueComments(&owned->rows);
                merged = browse_.values().MergeRows(std::move(owned->rows));
            }
            else
            {
                browse_.values().SetRows(std::move(owned->rows));
            }
            if (!merged)
            {
                RefreshValueListComments();
            }
            current_key_count_ = owned->key_count;
            current_value_count_ = owned->value_count;
            value_list_error_ = owned->error;
            if (!merged && list_hwnd && !jump_ui_batch_active_)
            {
                SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
                RedrawWindow(list_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
            }
            value_list_loading_ = false;
            UpdateStatus();
            StartPendingValueListRename();
            if (!merge && !retained_value_key_path_.empty() && browse_.current_node() &&
                EqualsInsensitive(registry_path::Build(*browse_.current_node()), retained_value_key_path_) &&
                !SelectValueByName(retained_value_name_))
            {
                SelectListRowAtIndex(browse_.values().hwnd(), retained_value_index_);
            }
            retained_value_name_.clear();
            retained_value_key_path_.clear();
            retained_value_index_ = -1;
            if (!pending_value_selection_key_.empty() && browse_.current_node() &&
                EqualsInsensitive(registry_path::Build(*browse_.current_node()), pending_value_selection_key_))
            {
                RestoreValueSelection();
            }
            if (!pending_value_name_.empty())
            {
                SelectValueByName(pending_value_name_);
                pending_value_name_.clear();
            }
            if (!pending_external_value_name_.empty() && browse_.current_node())
            {
                std::wstring current_path = registry_path::Build(*browse_.current_node());
                if (EqualsInsensitive(current_path, pending_external_value_key_path_))
                {
                    const bool selected = SelectValueByName(pending_external_value_name_);
                    pending_external_value_key_path_.clear();
                    pending_external_value_name_.clear();
                    const int command = pending_value_command_;
                    pending_value_command_ = 0;
                    if (selected && command != 0)
                    {
                        FocusPane(browse_.values().hwnd());
                        PostMessageW(hwnd_, WM_COMMAND, MAKEWPARAM(command, 0), 0);
                    }
                }
            }
            return 0;
        }
    default:
        return std::nullopt;
    }
}

std::optional<LRESULT> MainWindow::Impl::HandleExternalMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    switch (message)
    {
    case WM_DROPFILES:
        {
            HDROP drop = reinterpret_cast<HDROP>(wparam);
            UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
            std::vector<std::wstring> reg_paths;
            std::wstring offline_candidate;
            for (UINT index = 0; index < count; ++index)
            {
                UINT path_len = DragQueryFileW(drop, index, nullptr, 0);
                if (path_len == 0)
                {
                    continue;
                }
                std::wstring path(path_len + 1, L'\0');
                if (DragQueryFileW(drop, index, path.data(), static_cast<UINT>(path.size())) == 0)
                {
                    continue;
                }
                path.resize(path_len);
                if (util::HasFileExtension(path, L".reg"))
                {
                    reg_paths.push_back(path);
                }
                else if (offline_candidate.empty())
                {
                    offline_candidate = path;
                }
            }
            DragFinish(drop);
            if (!reg_paths.empty())
            {
                for (const auto& path : reg_paths)
                {
                    OpenRegFileTab(path);
                }
            }
            if (!offline_candidate.empty())
            {
                LoadOfflineRegistryFromPath(offline_candidate, true);
            }
            return 0;
        }
    case frame::message_id::kTreeRedraw:
        FlushTreeRedraw();
        return 0;
    case frame::message_id::kStatusUpdate:
        status_update_pending_ = false;
        UpdateStatus();
        return 0;
    case frame::message_id::kRegistryChanged:
        if (wparam == key_watcher_.generation())
        {
            const ULONGLONG elapsed = GetTickCount64() - auto_refresh_tick_;
            SetTimer(hwnd_, kAutoRefreshTimerId, elapsed >= 750 ? 250 : static_cast<UINT>(1000 - elapsed), nullptr);
        }
        return 0;
    case WM_TIMER:
        if (wparam == kAutoRefreshTimerId)
        {
            ApplyAutoRefresh();
            return 0;
        }
        if (wparam == kCompatJumpTimerId)
        {
            FlushExternalNavigation();
            return 0;
        }
        if (wparam == kTreeStateTimerId)
        {
            KillTimer(hwnd_, kTreeStateTimerId);
            CaptureTreeStateNow();
            return 0;
        }
        break;
    case WM_COPYDATA:
        {
            ULONG_PTR id = 0;
            auto payload = std::make_unique<HandoffPayload>();
            const std::wstring& target = payload->target;
            if (!win32ui::ReadInstanceText(wparam, lparam, kExternalMessageMaxBytes / sizeof(wchar_t) - 1, &id, &payload->target) ||
                (id != kExternalJumpCopyDataId && id != kEditRegFileCopyDataId) ||
                (id == kEditRegFileCopyDataId && !util::HasFileExtension(target, L".reg") && !IsHiveFile(target)))
            {
                return 0;
            }
            return win32ui::PostPayload(hwnd_, frame::message_id::kExternalHandoff, id, payload) ? TRUE : 0;
        }
    case frame::message_id::kExternalHandoff:
        {
            const std::unique_ptr<HandoffPayload> payload = win32ui::TakePayload<HandoffPayload>(lparam);
            const std::wstring& target = payload->target;
            win32ui::BringToFront(hwnd_);
            if (wparam == kEditRegFileCopyDataId)
            {
                if (AcceptHandoffFile(hwnd_, target))
                {
                    OpenFile(target);
                }
                return 0;
            }
            if (deferred_startup_complete_)
            {
                if (!NavigateToExternalJump(target))
                {
                    ui::ShowWarning(hwnd_, util::TrDetail(L"Registry path not found.", target));
                }
            }
            else
            {
                QueueExternalJump(target);
            }
            FocusAddressBarForExternalJump(true);
            return 0;
        }
    case WM_SETFOCUS:
        if (last_focus_ && IsWindow(last_focus_) && IsChild(hwnd_, last_focus_) && IsWindowVisible(last_focus_) &&
            IsWindowEnabled(last_focus_))
        {
            FocusPane(last_focus_);
            return 0;
        }
        break;
    case frame::message_id::kAddressEnter:
        NavigateToAddress();
        return 0;
    case frame::message_id::kFocusAddressBar:
        FocusAddressBarForExternalJump(false);
        return 0;
    default:
        return std::nullopt;
    }
    return std::nullopt;
}

std::optional<LRESULT> MainWindow::Impl::HandleAppearanceMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    switch (message)
    {
    case WM_ERASEBKGND:
    case WM_PAINT:
    case WM_PRINTCLIENT:
        return win32ui::HandleFillPaint(hwnd_, message, wparam, win32ui::ThemeBrush(win32ui::ThemeFill::background));
    case WM_SETTINGCHANGE:
    case WM_SYSCOLORCHANGE:
    case WM_THEMECHANGED:
        // WM_THEMECHANGED from our own reapply needs no system reread
        if (!applying_theme_ && win32ui::ThemeMessage(message, lparam))
        {
            ApplySystemTheme();
        }
        // Windows text size and menu font changed
        if (message == WM_SETTINGCHANGE && wparam == SPI_SETNONCLIENTMETRICS)
        {
            UpdateUIFont();
        }
        return 0;
    case WM_CTLCOLORSTATIC:
        {
            HDC hdc = reinterpret_cast<HDC>(wparam);
            HWND target = reinterpret_cast<HWND>(lparam);
            const win32ui::ThemeColors& colors = win32ui::CurrentThemeColors();
            COLORREF background = colors.panel;
            HBRUSH brush = win32ui::ThemeBrush(win32ui::ThemeFill::panel);
            if (target == history_label_ || target == tree_header_)
            {
                background = colors.header;
                brush = win32ui::ThemeBrush(win32ui::ThemeFill::header);
            }
            SetTextColor(hdc, colors.text);
            SetBkColor(hdc, background);
            return reinterpret_cast<LRESULT>(brush);
        }
    case WM_INITMENUPOPUP:
        // Windows manages window menu state
        if (!HIWORD(lparam))
        {
            UpdateMenuState(reinterpret_cast<HMENU>(wparam));
        }
        return 0;
    case WM_CTLCOLOREDIT:
        {
            return reinterpret_cast<LRESULT>(win32ui::ThemeControlColor(message, reinterpret_cast<HDC>(wparam), reinterpret_cast<HWND>(lparam)));
        }
    case WM_DRAWITEM:
        {
            auto* draw = reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
            if (menu_bar_.Draw(draw))
            {
                return TRUE;
            }
            if (draw && draw->CtlType == ODT_BUTTON)
            {
                DrawPanelButton(draw);
                return TRUE;
            }
            if (draw && draw->CtlType == ODT_STATIC && (draw->CtlID == kTreeHeaderId || draw->CtlID == kHistoryLabelId))
            {
                HDC hdc = draw->hDC;
                RECT rect = draw->rcItem;
                UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
                FillRect(hdc, &rect, win32ui::ThemeBrush(win32ui::ThemeFill::header));

                wchar_t text[128] = {};
                GetWindowTextW(draw->hwndItem, text, static_cast<int>(_countof(text)));
                HFONT old_font = nullptr;
                if (ui_font_)
                {
                    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
                }
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, win32ui::CurrentThemeColors().text);
                const UINT dpi = win32ui::DpiForWindow(hwnd_);
                const int padding = win32ui::Scale(kHeaderTextPadding, dpi);
                RECT text_rect = rect;
                text_rect.left += padding;
                // close button covers the right end of the header
                text_rect.right -= padding + win32ui::Scale(kPanelCloseInset + kPanelCloseSize, dpi);
                DrawTextW(hdc, text, -1, &text_rect, format);
                if (old_font)
                {
                    SelectObject(hdc, old_font);
                }
                return TRUE;
            }
            break;
        }
    case WM_MEASUREITEM:
        {
            auto* measure = reinterpret_cast<MEASUREITEMSTRUCT*>(lparam);
            if (menu_bar_.Measure(measure))
            {
                return TRUE;
            }
            break;
        }
    default:
        return std::nullopt;
    }
    return std::nullopt;
}

std::optional<LRESULT> MainWindow::Impl::HandleBrowseMessage(UINT message, WPARAM wparam, LPARAM lparam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kTreeHeaderCloseId)
            {
                settings_.show_tree = false;
                SaveSettings();
                ApplyViewVisibility();
                return 0;
            }
            if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kHistoryHeaderCloseId)
            {
                settings_.show_history = false;
                SaveSettings();
                ApplyViewVisibility();
                return 0;
            }
            if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kAddressGoId)
            {
                NavigateToAddress();
                return 0;
            }
            if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kFilterClearId)
            {
                ClearValueFilter(true);
                return 0;
            }
            if (HIWORD(wparam) == EN_CHANGE && LOWORD(wparam) == kAddressEditId)
            {
                UpdateGoButtonState();
                return 0;
            }
            if (HIWORD(wparam) == EN_CHANGE && LOWORD(wparam) == kFilterEditId)
            {
                const std::wstring buffer = util::WindowText(browse_.filter());
                bool needs_full_data = !buffer.empty();
                if (needs_full_data)
                {
                    needs_full_data =
                        std::any_of(browse_.values().rows().begin(), browse_.values().rows().end(), [](const ListRow& row) { return row.kind == rowkind::kValue && (!row.data_ready || row.data_preview); });
                }
                browse_.values().SetFilter(buffer);
                if (needs_full_data && browse_.current_node())
                {
                    UpdateValueListForNode(browse_.current_node());
                }
                UpdateStatus();
                return 0;
            }
            if (HIWORD(wparam) <= 1 && HandleMenuCommand(LOWORD(wparam)))
            {
                return 0;
            }
            return 0;
        }
    case WM_CONTEXTMENU:
        {
            HWND source = reinterpret_cast<HWND>(wparam);
            POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
            if (source == browse_.tree().hwnd())
            {
                const HTREEITEM item = win32ui::ContextMenuTreeItem(source, lparam, &screen_pt);
                ShowTreeContextMenu(item, screen_pt);
                return 0;
            }
            if (source != browse_.values().hwnd() && source != history_list_ && source != search_results_list_)
            {
                break;
            }
            // right click outside selection selects one row, empty space keeps selection
            const int row = win32ui::ContextMenuRow(source, lparam, &screen_pt, true);
            if (source == browse_.values().hwnd())
            {
                ShowValueContextMenu(row, screen_pt);
            }
            else if (source == history_list_)
            {
                ShowHistoryContextMenu(row, screen_pt);
            }
            else
            {
                ShowSearchResultContextMenu(row, screen_pt);
            }
            return 0;
        }
    case WM_NOTIFY:
        return HandleNotification(lparam);
    default:
        return std::nullopt;
    }
    return std::nullopt;
}

}
