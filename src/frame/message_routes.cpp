// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

#include "regfile/registry_transfer.h"

namespace regkit {
using namespace window_detail;

LRESULT MainWindow::Impl::HandleMessage(UINT message, WPARAM wparam, LPARAM lparam) {
  std::optional<LRESULT> result;
  switch (frame::ClassifyMessage(message)) {
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

std::optional<LRESULT> MainWindow::Impl::HandleLifecycleMessage(UINT message,
                                                                WPARAM wparam,
                                                                LPARAM lparam) {
  switch (message) {
  case WM_CREATE:
    return OnCreate() ? 0 : -1;
  case WM_DESTROY:
    OnDestroy();
    PostQuitMessage(0);
    return 0;
  case WM_NCPAINT: {
    LRESULT result = DefWindowProcW(hwnd_, message, wparam, lparam);
    PaintMenuBarSeparator();
    return result;
  }
  case WM_NCACTIVATE: {
    LRESULT result = DefWindowProcW(hwnd_, message, wparam, lparam);
    PaintMenuBarSeparator();
    return result;
  }
  case WM_GETMINMAXINFO: {
    auto* info = reinterpret_cast<MINMAXINFO*>(lparam);
    if (info) {
      info->ptMinTrackSize.x = std::max<LONG>(info->ptMinTrackSize.x, 400);
      info->ptMinTrackSize.y = std::max<LONG>(info->ptMinTrackSize.y, 200);
    }
    return 0;
  }
  case WM_SIZE:
    OnSize(LOWORD(lparam), HIWORD(lparam));
    return 0;
  case WM_DPICHANGED: {
    const RECT* suggested = reinterpret_cast<const RECT*>(lparam);
    if (suggested) {
      SetWindowPos(hwnd_, nullptr, suggested->left, suggested->top, suggested->right - suggested->left, suggested->bottom - suggested->top, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    UpdateUIFont();
    ReloadThemeIcons();
    return 0;
  }
  case WM_DPICHANGED_AFTERPARENT:
    UpdateUIFont();
    ReloadThemeIcons();
    return 0;
  case WM_ACTIVATE:
    if (LOWORD(wparam) == WA_ACTIVE && browse_.tree().hwnd()) {
      if (last_focus_) {
        break;
      }
      int sel = TabCtrl_GetCurSel(tab_);
      if (sel >= 0 && static_cast<size_t>(sel) < tabs_.size() && tabs_[static_cast<size_t>(sel)].kind == TabEntry::Kind::kRegistry) {
        SetFocus(browse_.tree().hwnd());
      }
    }
    break;
  case WM_CLOSE: {
    SaveSettings();
    DestroyWindow(hwnd_);
    return 0;
  }
  default:
    return std::nullopt;
  }
  return std::nullopt;
}

std::optional<LRESULT> MainWindow::Impl::HandleLayoutInputMessage(UINT message,
                                                                  WPARAM wparam,
                                                                  LPARAM lparam) {
  (void)wparam;
  switch (message) {
  case WM_LBUTTONDOWN: {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    if (show_tree_ && PtInRect(&splitter_rect_, pt)) {
      splitter_start_x_ = pt.x;
      splitter_start_width_ = tree_width_;
      BeginSplitterDrag();
      return 0;
    }
    if (show_history_ && PtInRect(&history_splitter_rect_, pt)) {
      history_splitter_start_y_ = pt.y;
      history_splitter_start_height_ = history_height_;
      BeginHistorySplitterDrag();
      return 0;
    }
    break;
  }
  case WM_LBUTTONUP:
    if (splitter_dragging_) {
      EndSplitterDrag();
      return 0;
    }
    if (history_splitter_dragging_) {
      EndHistorySplitterDrag();
      return 0;
    }
    break;
  case WM_MOUSEMOVE: {
    if (splitter_dragging_) {
      POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      UpdateSplitterTrack(pt.x);
      return 0;
    }
    if (history_splitter_dragging_) {
      POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      UpdateHistorySplitterTrack(pt.y);
      return 0;
    }
    break;
  }
  case WM_CAPTURECHANGED:
    if (splitter_dragging_) {
      EndSplitterDrag();
      return 0;
    }
    if (history_splitter_dragging_) {
      EndHistorySplitterDrag();
      return 0;
    }
    break;
  case WM_SETCURSOR: {
    if (splitter_dragging_) {
      SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
      return TRUE;
    }
    if (history_splitter_dragging_) {
      SetCursor(LoadCursorW(nullptr, IDC_SIZENS));
      return TRUE;
    }
    if (show_tree_) {
      POINT pt = {};
      GetCursorPos(&pt);
      ScreenToClient(hwnd_, &pt);
      if (PtInRect(&splitter_rect_, pt)) {
        SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
        return TRUE;
      }
    }
    if (show_history_) {
      POINT pt = {};
      GetCursorPos(&pt);
      ScreenToClient(hwnd_, &pt);
      if (PtInRect(&history_splitter_rect_, pt)) {
        SetCursor(LoadCursorW(nullptr, IDC_SIZENS));
        return TRUE;
      }
    }
    break;
  }
  default:
    return std::nullopt;
  }
  return std::nullopt;
}

std::optional<LRESULT> MainWindow::Impl::HandleWorkerMessage(UINT message,
                                                             WPARAM wparam,
                                                             LPARAM lparam) {
  switch (message) {
  case frame::message_id::kSearchResults:
  case frame::message_id::kSearchPreviewRequest:
  case frame::message_id::kSearchPreviewReady:
  case frame::message_id::kSearchSortReady:
  case frame::message_id::kSearchTabLoadReady:
  case frame::message_id::kSearchProgress:
  case frame::message_id::kSearchFailed:
  case frame::message_id::kReplaceReady:
    return HandleSearchWorkerMessage(message, wparam, lparam);
  case frame::message_id::kLoadTraces:
  case frame::message_id::kLoadDefaults:
  case frame::message_id::kTraceLoadReady:
  case frame::message_id::kDefaultLoadReady:
  case frame::message_id::kDeferredStartup:
  case frame::message_id::kStartupCacheReady:
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

void MainWindow::Impl::FinishSearchSession(uint64_t generation) {
  if (!search_running_ || !search_session_.IsCurrent(generation)) {
    return;
  }
  search_session_.Join();
  search_running_ = false;
  for (auto& search_tab : search_tabs_) {
    if (search_tab.generation == generation && search_tab.sort_dirty) {
      SortSearchTabResults(&search_tab);
      break;
    }
  }
  if (search_start_tick_ != 0) {
    search_duration_ms_ = GetTickCount64() - search_start_tick_;
    search_duration_valid_ = true;
  } else {
    search_duration_ms_ = 0;
    search_duration_valid_ = false;
  }
  if (IsSearchTabIndex(TabCtrl_GetCurSel(tab_))) {
    search_last_refresh_tick_ = GetTickCount64();
    UpdateSearchResultsView();
  }
  ApplyViewVisibility();
  UpdateStatus();
}

std::optional<LRESULT> MainWindow::Impl::HandleSearchWorkerMessage(UINT message,
                                                                   WPARAM wparam,
                                                                   LPARAM lparam) {
  switch (message) {
  case frame::message_id::kSearchResults: {
    const uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }

    search_posted_.store(false);
    const int index = IsSearchTabIndex(active_search_tab_index_)
                          ? SearchIndexFromTab(active_search_tab_index_)
                          : -1;
    SearchTab* tab = index >= 0 && static_cast<size_t>(index) < search_tabs_.size()
                         ? &search_tabs_[static_cast<size_t>(index)]
                         : nullptr;

    const uint64_t start_tick = GetTickCount64();
    bool appended = false;
    for (;;) {
      PendingSearchBatch batch;
      {
        std::lock_guard<std::mutex> lock(search_mutex_);
        if (search_pending_batches_.empty()) {
          break;
        }
        batch = std::move(search_pending_batches_.front());
        search_pending_batches_.pop_front();
        search_pending_rows_ -= batch.rows.size();
      }
      search_queue_space_.notify_all();
      if (batch.generation != generation || !tab) {
        continue;
      }
      for (auto& row : batch.rows) {
        row.row_id = tab->next_row_id++;
      }

      if (tab->results.empty()) {
        tab->results = std::move(batch.rows);
      } else {
        tab->results.insert(tab->results.end(),
                            std::make_move_iterator(batch.rows.begin()),
                            std::make_move_iterator(batch.rows.end()));
      }
      appended = true;
      if ((GetTickCount64() - start_tick) >= kSearchResultsMaxMs) {
        break;
      }
    }

    if (appended && tab && tab->sort_column >= 0) {
      tab->sort_dirty = true;
    }

    bool more_pending = false;
    bool producer_done = false;
    {
      std::lock_guard<std::mutex> lock(search_mutex_);
      more_pending = !search_pending_batches_.empty();
      producer_done = search_producer_done_;
    }

    if (more_pending && !search_posted_.exchange(true)) {
      if (!PostMessageW(hwnd_, frame::message_id::kSearchResults,
                        static_cast<WPARAM>(generation), 0)) {
        search_posted_.store(false);
      }
    }

    if (appended && TabCtrl_GetCurSel(tab_) == active_search_tab_index_) {
      const uint64_t now = GetTickCount64();
      if (now - search_last_refresh_tick_ >= kSearchResultsRefreshMs) {
        search_last_refresh_tick_ = now;
        UpdateSearchResultsView();
        UpdateStatus();
      }
    }


    if (!more_pending && producer_done) {
      FinishSearchSession(generation);
    }
    return 0;
  }
  case frame::message_id::kSearchPreviewRequest: {
    search_preview_request_posted_ = false;
    if (!search_results_list_) {
      return 0;
    }
    const int top = ListView_GetTopIndex(search_results_list_);
    const int page = ListView_GetCountPerPage(search_results_list_);
    const int count = ListView_GetItemCount(search_results_list_);
    QueueSearchPreviews(std::max(0, top),
                        std::min(count - 1, top + std::max(page, 1)));
    return 0;
  }
  case frame::message_id::kSearchPreviewReady: {
    auto* raw = reinterpret_cast<SearchPreviewPayload*>(lparam);
    if (!raw) {
      return 0;
    }
    std::unique_ptr<SearchPreviewPayload> owned(raw);
    if (owned->tab_index < 0 ||
        static_cast<size_t>(owned->tab_index) >= search_tabs_.size()) {
      return 0;
    }
    SearchTab& tab = search_tabs_[static_cast<size_t>(owned->tab_index)];
    if (tab.generation != owned->generation) {
      return 0;
    }
    int first = -1;
    int last = -1;
    for (auto& item : owned->items) {


      int index = -1;
      if (item.index >= 0 && static_cast<size_t>(item.index) < tab.results.size() &&
          tab.results[static_cast<size_t>(item.index)].row_id == item.row_id) {
        index = item.index;
      } else {
        for (size_t i = 0; i < tab.results.size(); ++i) {
          if (tab.results[i].row_id == item.row_id) {
            index = static_cast<int>(i);
            break;
          }
        }
      }
      if (index < 0) {
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
    if (first >= 0 && search_results_list_ &&
        SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) == owned->tab_index) {
      ListView_RedrawItems(search_results_list_, first, last);
    }
    return 0;
  }
  case frame::message_id::kSearchSortReady: {
    auto* raw = reinterpret_cast<SearchSortPayload*>(lparam);
    if (!raw) {
      return 0;
    }
    std::unique_ptr<SearchSortPayload> owned(raw);
    if (owned->tab_index < 0 ||
        static_cast<size_t>(owned->tab_index) >= search_tabs_.size()) {
      return 0;
    }
    SearchTab& tab = search_tabs_[static_cast<size_t>(owned->tab_index)];
    if (tab.generation != owned->generation ||
        tab.results.size() != owned->rows.size()) {
      return 0;
    }
    tab.results = std::move(owned->rows);
    if (SearchIndexFromTab(TabCtrl_GetCurSel(tab_)) == owned->tab_index) {
      UpdateSearchResultsView();
    }
    return 0;
  }
  case frame::message_id::kSearchTabLoadReady:
    ApplySearchTabLoad(reinterpret_cast<SearchTabLoadPayload*>(lparam));
    return 0;
  case frame::message_id::kSearchProgress: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    search_progress_posted_.store(false);
    UpdateStatus();
    return 0;
  }
  case frame::message_id::kSearchFailed: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    search_session_.Join();
    search_running_ = false;
    search_duration_ms_ = 0;
    search_duration_valid_ = false;
    ui::ShowError(hwnd_, L"Invalid regex.");
    ApplyViewVisibility();
    UpdateStatus();
    return 0;
  }
  case frame::message_id::kReplaceReady:
    ApplyReplacePayload(reinterpret_cast<ReplacePayload*>(lparam));
    return 0;
  default:
    return std::nullopt;
  }
}

std::optional<LRESULT> MainWindow::Impl::HandleLoadWorkerMessage(UINT message,
                                                                 WPARAM wparam,
                                                                 LPARAM lparam) {
  (void)wparam;
  switch (message) {
  case frame::message_id::kLoadTraces:
    StartTraceLoadWorker();
    return 0;
  case frame::message_id::kLoadDefaults:
    StartDefaultLoadWorker();
    return 0;
  case frame::message_id::kTraceLoadReady: {
    auto* payload = reinterpret_cast<TraceLoadPayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<TraceLoadPayload> owned(payload);
    if (!trace_load_session_.IsCurrent(owned->generation)) {
      return 0;
    }
    trace_load_session_.Join();
    active_traces_ = std::move(owned->traces);
    trace_selection_cache_ = std::move(owned->selection_cache);
    BuildMenus();
    RefreshTreeSelection();
    UpdateValueListForNode(browse_.current_node());
    return 0;
  }
  case frame::message_id::kDefaultLoadReady: {
    auto* payload = reinterpret_cast<DefaultLoadPayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<DefaultLoadPayload> owned(payload);
    if (!default_load_session_.IsCurrent(owned->generation)) {
      return 0;
    }
    default_load_session_.Join();
    active_defaults_ = std::move(owned->defaults);
    BuildMenus();
    UpdateValueListForNode(browse_.current_node());
    return 0;
  }
  case frame::message_id::kDeferredStartup:
    RunDeferredStartup();
    return 0;
  case frame::message_id::kStartupCacheReady:
    ApplyStartupCachePayload(reinterpret_cast<StartupCachePayload*>(lparam));
    return 0;
  default:
    return std::nullopt;
  }
}

std::optional<LRESULT> MainWindow::Impl::HandleRegFileWorkerMessage(UINT message,
                                                                    WPARAM wparam,
                                                                    LPARAM lparam) {
  (void)wparam;
  switch (message) {
  case frame::message_id::kRegFileLoadReady: {
    auto* payload = reinterpret_cast<RegFileParsePayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<RegFileParsePayload> owned(payload);
    auto session_it = reg_file_parse_sessions_.find(owned->source_lower);
    if (session_it == reg_file_parse_sessions_.end()) {
      return 0;
    }
    if (!session_it->second ||
        !session_it->second->work.IsCurrent(owned->generation)) {
      return 0;
    }
    session_it->second->work.Join();
    reg_file_parse_sessions_.erase(session_it);
    if (owned->cancelled) {
      return 0;
    }

    int tab_index = -1;
    for (size_t i = 0; i < tabs_.size(); ++i) {
      const TabEntry& entry = tabs_[i];
      if (entry.kind != TabEntry::Kind::kRegFile) {
        continue;
      }
      if (EqualsInsensitive(entry.reg_file_path, owned->source_path)) {
        tab_index = static_cast<int>(i);
        break;
      }
    }
    if (tab_index < 0 || static_cast<size_t>(tab_index) >= tabs_.size()) {
      return 0;
    }
    TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
    entry.reg_file_loading = false;
    if (!owned->error.empty()) {
      ui::ShowError(hwnd_, owned->error.c_str());
      UpdateStatus();
      return 0;
    }
    if (entry.reg_file_dirty) {
      UpdateStatus();
      return 0;
    }

    ReleaseRegFileRoots(&entry);
    std::vector<TabEntry::RegFileRoot> roots;
    roots.reserve(owned->roots.size());
    for (auto& parsed : owned->roots) {
      if (!parsed.data) {
        continue;
      }
      TabEntry::RegFileRoot root;
      root.name = parsed.name;
      root.data = parsed.data;
      root.root = RegistryStore::RegisterVirtualRoot(root.name, root.data);
      if (root.root) {
        roots.push_back(std::move(root));
      }
    }
    entry.reg_file_roots = std::move(roots);
    entry.reg_file_dirty = false;
    if (tab_ && TabCtrl_GetCurSel(tab_) == tab_index) {
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

std::optional<LRESULT> MainWindow::Impl::HandleTraceWorkerMessage(UINT message,
                                                                  WPARAM wparam,
                                                                  LPARAM lparam) {
  (void)wparam;
  switch (message) {
  case frame::message_id::kTraceParseBatch: {
    auto* payload = reinterpret_cast<TraceParseBatch*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<TraceParseBatch> owned(payload);
    auto it = trace_parse_sessions_.find(owned->source_lower);
    if (it == trace_parse_sessions_.end()) {
      return 0;
    }
    TraceParseSession* session = it->second.get();
    if (!session || !session->work.IsCurrent(owned->generation)) {
      return 0;
    }
    const bool touches_current =
        browse_.current_node() &&
        owned->affected_keys.find(TracePathLowerForNode(*browse_.current_node())) !=
            owned->affected_keys.end();
    if (session->dialog && IsWindow(session->dialog) && !owned->entries.empty()) {
      auto dialog_entries = std::make_unique<std::vector<KeyValueDialogEntry>>(std::move(owned->entries));
      TraceDialogPostEntries(session->dialog, dialog_entries.release());
    }
    if (session->added_to_active && touches_current && browse_.current_node()) {
      uint64_t now = GetTickCount64();
      if (owned->done || (now - last_trace_refresh_tick_) >= 100) {
        last_trace_refresh_tick_ = now;
        UpdateValueListForNode(browse_.current_node());
      }
    }
    if (owned->done) {
      session->parsing_done = true;
      if (session->data) {
        std::unique_lock<std::shared_mutex> data_lock(*session->data->mutex);
        trace::Selection normalized = session->selection;
        trace::NormalizeSelection(*session->data, &normalized);
        session->selection = normalized;
        trace_selection_cache_[session->source_lower] = normalized;
        if (session->added_to_active) {
          for (auto& trace : active_traces_) {
            if (EqualsInsensitive(trace.source_path, session->source_path)) {
              trace.selection =
                  std::make_shared<trace::Selection>(normalized);
              break;
            }
          }
        }
      }
      if (!owned->error.empty()) {
        HWND error_owner = session->dialog && IsWindow(session->dialog) ? session->dialog : hwnd_;
        ui::ShowError(error_owner, owned->error.c_str());
        if (session->dialog && IsWindow(session->dialog)) {
          PostMessageW(session->dialog, WM_CLOSE, 0, 0);
        }
        if (session->added_to_active) {
          active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(), [&](const ActiveTrace& trace) { return EqualsInsensitive(trace.source_path, session->source_path); }), active_traces_.end());
          trace_selection_cache_.erase(session->source_lower);
          SaveActiveTraces();
          SaveTraceSettings();
          BuildMenus();
          RefreshTreeSelection();
          UpdateValueListForNode(browse_.current_node());
          SaveSettings();
          session->added_to_active = false;
        }
      } else if (session->dialog && IsWindow(session->dialog)) {
        TraceDialogPostDone(session->dialog, true);
      }
      if (session->added_to_active && browse_.current_node()) {
        UpdateValueListForNode(browse_.current_node());
      }
      session->work.Join();
      if (!session->dialog || !IsWindow(session->dialog)) {
        trace_parse_sessions_.erase(it);
      }
    }
    return 0;
  }
  default:
    return std::nullopt;
  }
}

std::optional<LRESULT> MainWindow::Impl::HandleDefaultWorkerMessage(UINT message,
                                                                    WPARAM wparam,
                                                                    LPARAM lparam) {
  (void)wparam;
  switch (message) {
  case frame::message_id::kDefaultParseBatch: {
    auto* payload = reinterpret_cast<DefaultParseBatch*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<DefaultParseBatch> owned(payload);
    auto it = default_parse_sessions_.find(owned->source_lower);
    if (it == default_parse_sessions_.end()) {
      return 0;
    }
    DefaultParseSession* session = it->second.get();
    if (!session || !session->work.IsCurrent(owned->generation)) {
      return 0;
    }
    bool touches_current = false;
    if (browse_.current_node()) {
      std::wstring path = registry_path::Build(*browse_.current_node());
      std::wstring normalized = NormalizeTraceKeyPathBasic(path);
      if (normalized.empty()) {
        normalized = path;
      }
      touches_current =
          owned->affected_keys.find(ToLower(normalized)) !=
          owned->affected_keys.end();
    }
    if (session->dialog && IsWindow(session->dialog) && !owned->entries.empty()) {
      auto dialog_entries = std::make_unique<std::vector<KeyValueDialogEntry>>(std::move(owned->entries));
      TraceDialogPostEntries(session->dialog, dialog_entries.release());
    }
    if (session->added_to_active && touches_current && browse_.current_node()) {
      uint64_t now = GetTickCount64();
      if (owned->done || (now - last_default_refresh_tick_) >= 100) {
        last_default_refresh_tick_ = now;
        UpdateValueListForNode(browse_.current_node());
      }
    }
    if (owned->done) {
      session->parsing_done = true;
      if (!owned->error.empty()) {
        if (session->show_errors) {
          HWND error_owner = session->dialog && IsWindow(session->dialog) ? session->dialog : hwnd_;
          ui::ShowError(error_owner, owned->error.c_str());
        }
        if (session->dialog && IsWindow(session->dialog)) {
          PostMessageW(session->dialog, WM_CLOSE, 0, 0);
        }
        if (session->added_to_active) {
          active_defaults_.erase(std::remove_if(active_defaults_.begin(), active_defaults_.end(), [&](const ActiveDefault& defaults) { return EqualsInsensitive(defaults.source_path, session->source_path); }), active_defaults_.end());
          SaveActiveDefaults();
          BuildMenus();
          UpdateValueListForNode(browse_.current_node());
          SaveSettings();
          session->added_to_active = false;
        }
      } else if (session->dialog && IsWindow(session->dialog)) {
        TraceDialogPostDone(session->dialog, true);
      }
      if (session->added_to_active && browse_.current_node()) {
        UpdateValueListForNode(browse_.current_node());
      }
      session->work.Join();
      if (!session->dialog || !IsWindow(session->dialog)) {
        default_parse_sessions_.erase(it);
      }
    }
    return 0;
  }
  default:
    return std::nullopt;
  }
}

std::optional<LRESULT> MainWindow::Impl::HandleValueWorkerMessage(UINT message,
                                                                  WPARAM wparam,
                                                                  LPARAM lparam) {
  (void)wparam;
  switch (message) {
  case frame::message_id::kValuePreviewRequest: {
    value_preview_request_posted_ = false;
    HWND list = browse_.values().hwnd();
    if (!list) {
      return 0;
    }
    const int top = ListView_GetTopIndex(list);
    const int page = ListView_GetCountPerPage(list);
    const int count = ListView_GetItemCount(list);
    QueueValuePreviews(std::max(0, top),
                       std::min(count - 1, top + std::max(page, 1)));
    return 0;
  }
  case frame::message_id::kValuePreviewReady: {
    auto* raw = reinterpret_cast<ValuePreviewPayload*>(lparam);
    if (!raw) {
      return 0;
    }
    std::unique_ptr<ValuePreviewPayload> owned(raw);
    if (owned->generation != value_list_generation_.load()) {
      return 0;
    }
    int first = -1;
    int last = -1;
    for (const auto& item : owned->items) {
      int index = item.index;
      ListRow* row = browse_.values().MutableRowAt(index);
      if (!row || row->extra != item.name) {
        row = nullptr;
        const int count = static_cast<int>(browse_.values().RowCount());
        for (int i = 0; i < count; ++i) {
          ListRow* candidate = browse_.values().MutableRowAt(i);
          if (candidate && candidate->kind == rowkind::kValue &&
              candidate->extra == item.name) {
            row = candidate;
            index = i;
            break;
          }
        }
      }
      if (!row) {
        continue;
      }
      row->data = item.preview;
      row->value_type = item.type;
      row->value_data_size = item.size;
      if (row->type.empty()) {
        row->type = value_format::TypeName(item.type);
      }
      row->size_value = item.size;
      row->has_size = true;
      if (row->size.empty() && item.size > 0) {
        row->size = std::to_wstring(item.size);
      }
      row->data_ready = true;
      browse_.values().InvalidateFilterCache(row);
      first = first < 0 ? index : std::min(first, index);
      last = std::max(last, index);
    }
    if (first >= 0 && browse_.values().hwnd()) {
      ListView_RedrawItems(browse_.values().hwnd(), first, last);
    }
    return 0;
  }
  case frame::message_id::kValueListReady: {
    auto* payload = reinterpret_cast<ValueListPayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<ValueListPayload> owned(payload);
    if (payload->generation != value_list_generation_.load()) {
      return 0;
    }
    HWND list_hwnd = browse_.values().hwnd();
    if (list_hwnd) {
      SendMessageW(list_hwnd, WM_SETREDRAW, FALSE, 0);
    }
    browse_.values().SetRows(std::move(payload->rows));
    RefreshValueListComments();
    current_key_count_ = payload->key_count;
    current_value_count_ = payload->value_count;
    if (list_hwnd && !jump_ui_batch_active_) {
      SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
      RedrawWindow(list_hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
    }
    value_list_loading_ = false;
    UpdateStatus();
    StartPendingValueListRename();
    if (!retained_value_key_path_.empty() && browse_.current_node() &&
        EqualsInsensitive(registry_path::Build(*browse_.current_node()),
                          retained_value_key_path_)) {
      SelectValueByName(retained_value_name_);
    }
    retained_value_name_.clear();
    retained_value_key_path_.clear();
    if (!pending_external_value_name_.empty() && browse_.current_node()) {
      std::wstring current_path = registry_path::Build(*browse_.current_node());
      if (EqualsInsensitive(current_path, pending_external_value_key_path_)) {
        const bool selected = SelectValueByName(pending_external_value_name_);
        pending_external_value_key_path_.clear();
        pending_external_value_name_.clear();
        const int command = pending_value_command_;
        pending_value_command_ = 0;
        if (selected && command != 0) {
          if (browse_.values().hwnd()) {
            SetFocus(browse_.values().hwnd());
          }
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

std::optional<LRESULT> MainWindow::Impl::HandleExternalMessage(UINT message,
                                                               WPARAM wparam,
                                                               LPARAM lparam) {
  switch (message) {
  case WM_DROPFILES: {
    HDROP drop = reinterpret_cast<HDROP>(wparam);
    UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
    std::vector<std::wstring> reg_paths;
    std::wstring offline_candidate;
    for (UINT index = 0; index < count; ++index) {
      UINT path_len = DragQueryFileW(drop, index, nullptr, 0);
      if (path_len == 0) {
        continue;
      }
      std::wstring path(path_len + 1, L'\0');
      if (DragQueryFileW(drop, index, path.data(), static_cast<UINT>(path.size())) == 0) {
        continue;
      }
      path.resize(path_len);
      if (HasRegExtension(path)) {
        reg_paths.push_back(path);
      } else if (offline_candidate.empty()) {
        offline_candidate = path;
      }
    }
    DragFinish(drop);
    if (!reg_paths.empty()) {
      for (const auto& path : reg_paths) {
        OpenRegFileTab(path);
      }
    }
    if (!offline_candidate.empty()) {
      LoadOfflineRegistryFromPath(offline_candidate, true);
    }
    return 0;
  }
  case WM_TIMER:
    break;
  case WM_COPYDATA: {
    auto* data = reinterpret_cast<const COPYDATASTRUCT*>(lparam);
    if (!data) {
      return 0;
    }
    if (data->dwData != kExternalJumpCopyDataId || !data->lpData || data->cbData < sizeof(wchar_t)) {
      return 0;
    }
    size_t length = data->cbData / sizeof(wchar_t);
    const wchar_t* text = reinterpret_cast<const wchar_t*>(data->lpData);
    std::wstring target(text, text + length);
    while (!target.empty() && target.back() == L'\0') {
      target.pop_back();
    }
    if (target.empty()) {
      return 0;
    }
    if (deferred_startup_complete_) {
      NavigateToExternalJump(target);
    } else {
      QueueExternalJump(target);
    }
    ShowWindow(hwnd_, SW_RESTORE);
    SetForegroundWindow(hwnd_);
    FocusAddressBarForExternalJump(true);
    return TRUE;
  }
  case WM_SETFOCUS:
    if (last_focus_ && IsWindow(last_focus_) && IsChild(hwnd_, last_focus_) &&
        IsWindowVisible(last_focus_) && IsWindowEnabled(last_focus_)) {
      SetFocus(last_focus_);
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

std::optional<LRESULT> MainWindow::Impl::HandleAppearanceMessage(UINT message,
                                                                 WPARAM wparam,
                                                                 LPARAM lparam) {
  switch (message) {
  case WM_ERASEBKGND: {
    return 1;
  }
  case WM_PAINT:
    OnPaint();
    return 0;
  case WM_SETTINGCHANGE: {
    if (applying_theme_ || theme_mode_ != ThemeMode::kSystem) {
      return 0;
    }
    if (!Theme::UpdateFromSystem()) {
      return 0;
    }
    applying_theme_ = true;
    Theme::Current().ApplyToWindow(hwnd_);
    ApplyThemeToChildren();
    ReloadThemeIcons();
    if (hwnd_) {
      InvalidateRect(hwnd_, nullptr, TRUE);
    }
    applying_theme_ = false;
    return 0;
  }
  case WM_THEMECHANGED:
    if (applying_theme_ || theme_mode_ != ThemeMode::kSystem) {
      return 0;
    }
    if (!Theme::UpdateFromSystem()) {
      return 0;
    }
    applying_theme_ = true;
    Theme::Current().ApplyToWindow(hwnd_);
    ApplyThemeToChildren();
    ReloadThemeIcons();
    if (hwnd_) {
      InvalidateRect(hwnd_, nullptr, TRUE);
    }
    applying_theme_ = false;
    return 0;
  case WM_CTLCOLORSTATIC: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    const Theme& theme = Theme::Current();
    COLORREF color = theme.TextColor();
    COLORREF background = theme.PanelColor();
    HBRUSH brush = theme.PanelBrush();
    if (target == history_label_ || target == tree_header_) {
      background = theme.HeaderColor();
      brush = theme.HeaderBrush();
    }
    SetTextColor(hdc, color);
    SetBkColor(hdc, background);
    return reinterpret_cast<LRESULT>(brush);
  }
  case WM_INITMENUPOPUP: {
    HMENU menu = reinterpret_cast<HMENU>(wparam);
    CheckMenuItem(menu, cmd::kViewGridLines,
                  MF_BYCOMMAND | (show_value_grid_ ? MF_CHECKED : MF_UNCHECKED));
    UINT state = browse_.current_node() ? MF_ENABLED : MF_GRAYED;
    EnableMenuItem(menu, cmd::kEditPermissions, MF_BYCOMMAND | state);
    const int selected_count = browse_.values().hwnd() ? ListView_GetSelectedCount(browse_.values().hwnd()) : 0;
    const int selected_index = selected_count == 1 ? ListView_GetNextItem(browse_.values().hwnd(), -1, LVNI_SELECTED) : -1;
    const ListRow* selected_row = selected_index >= 0 ? browse_.values().RowAt(selected_index) : nullptr;
    const bool can_modify_value = !read_only_ && selected_row &&
                                  selected_row->kind == rowkind::kValue &&
                                  !selected_row->simulated;
    const UINT modify_state = can_modify_value ? MF_ENABLED : MF_GRAYED;
    EnableMenuItem(menu, cmd::kEditModify, MF_BYCOMMAND | modify_state);
    EnableMenuItem(menu, cmd::kEditModifyBinary, MF_BYCOMMAND | modify_state);
    const bool hives_allowed = !read_only_ && registry_mode_ != RegistryMode::kRemote;
    const RegistryNode* hive_node = browse_.current_node();
    const bool hive_selected =
        hives_allowed && hive_node &&
        IsMountedHive(hive_node->root, hive_node->subkey);
    EnableMenuItem(menu, cmd::kFileLoadHive,
                   MF_BYCOMMAND | (hives_allowed ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(menu, cmd::kFileUnloadHive,
                   MF_BYCOMMAND | (hive_selected ? MF_ENABLED : MF_GRAYED));
    if (GetMenuState(menu, cmd::kOptionsHiveFileDir, MF_BYCOMMAND) !=
        static_cast<UINT>(-1)) {
      const UINT hive_state = ResolveSelectedHiveFilePath().empty()
                                  ? MF_GRAYED
                                  : MF_ENABLED;
      EnableMenuItem(menu, cmd::kOptionsHiveFileDir,
                     MF_BYCOMMAND | hive_state);
    }
    return 0;
  }
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    SetTextColor(hdc, Theme::Current().TextColor());
    SetBkColor(hdc, Theme::Current().FieldColor());
    return reinterpret_cast<LRESULT>(Theme::Current().FieldBrush());
  }
  case WM_DRAWITEM: {
    auto* draw = reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
    if (draw && draw->CtlType == ODT_MENU) {
      OnDrawMenuItem(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_BUTTON && draw->CtlID == kAddressGoId) {
      DrawAddressButton(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_BUTTON && draw->CtlID == kTreeHeaderCloseId) {
      DrawHeaderCloseButton(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_BUTTON && draw->CtlID == kHistoryHeaderCloseId) {
      DrawHeaderCloseButton(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_STATIC && (draw->CtlID == kTreeHeaderId || draw->CtlID == kHistoryLabelId)) {
      const Theme& theme = Theme::Current();
      HDC hdc = draw->hDC;
      RECT rect = draw->rcItem;
      UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
      FillRect(hdc, &rect, theme.HeaderBrush());

      wchar_t text[128] = {};
      GetWindowTextW(draw->hwndItem, text, static_cast<int>(_countof(text)));
      HFONT old_font = nullptr;
      if (ui_font_) {
        old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
      }
      SetBkMode(hdc, TRANSPARENT);
      SetTextColor(hdc, theme.TextColor());
      RECT text_rect = rect;
      text_rect.left += kHeaderTextPadding;
      text_rect.right -= kHeaderTextPadding;
      DrawTextW(hdc, text, -1, &text_rect, format);
      if (old_font) {
        SelectObject(hdc, old_font);
      }
      return TRUE;
    }
    break;
  }
  case WM_MEASUREITEM: {
    auto* measure = reinterpret_cast<MEASUREITEMSTRUCT*>(lparam);
    if (measure && measure->CtlType == ODT_MENU) {
      OnMeasureMenuItem(measure);
      return TRUE;
    }
    break;
  }
  default:
    return std::nullopt;
  }
  return std::nullopt;
}

std::optional<LRESULT> MainWindow::Impl::HandleBrowseMessage(UINT message,
                                                             WPARAM wparam,
                                                             LPARAM lparam) {
  switch (message) {
  case WM_COMMAND: {
    if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kTreeHeaderCloseId) {
      show_tree_ = false;
      ApplyViewVisibility();
      BuildMenus();
      return 0;
    }
    if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kHistoryHeaderCloseId) {
      show_history_ = false;
      SaveSettings();
      ApplyViewVisibility();
      BuildMenus();
      return 0;
    }
    if (HIWORD(wparam) == 0 && LOWORD(wparam) == kValueGridButtonId) {
      SetValueGridEnabled(!show_value_grid_, true);
      return 0;
    }
    if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kAddressGoId) {
      NavigateToAddress();
      return 0;
    }
    if (HIWORD(wparam) == EN_CHANGE && LOWORD(wparam) == kFilterEditId) {
      const std::wstring buffer = util::WindowText(browse_.filter());
      bool needs_full_data = !buffer.empty();
      if (needs_full_data) {
        needs_full_data = std::any_of(browse_.values().rows().begin(), browse_.values().rows().end(), [](const ListRow& row) {
          return row.kind == rowkind::kValue && !row.data_ready;
        });
      }
      browse_.values().SetFilter(buffer);
      if (needs_full_data && browse_.current_node()) {
        UpdateValueListForNode(browse_.current_node());
      }
      UpdateStatus();
      return 0;
    }
    if (HIWORD(wparam) == 0 && HandleMenuCommand(LOWORD(wparam))) {
      return 0;
    }
    return 0;
  }
  case WM_CONTEXTMENU: {
    HWND source = reinterpret_cast<HWND>(wparam);
    HWND header_hwnd = ListView_GetHeader(browse_.values().hwnd());
    if (source == header_hwnd) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(header_hwnd, &rect);
        screen_pt.x = rect.left + 12;
        screen_pt.y = rect.bottom - 4;
      }
      ShowValueHeaderMenu(screen_pt);
      return 0;
    }
    if (source == browse_.tree().hwnd()) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(browse_.tree().hwnd(), &rect);
        screen_pt.x = rect.left + 16;
        screen_pt.y = rect.top + 16;
      }
      ShowTreeContextMenu(screen_pt);
      return 0;
    }
    if (source == browse_.values().hwnd()) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(browse_.values().hwnd(), &rect);
        screen_pt.x = rect.left + 24;
        screen_pt.y = rect.top + 24;
      }
      ShowValueContextMenu(screen_pt);
      return 0;
    }
    if (source == history_list_) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(history_list_, &rect);
        screen_pt.x = rect.left + 24;
        screen_pt.y = rect.top + 24;
      }
      ShowHistoryContextMenu(screen_pt);
      return 0;
    }
    if (source == search_results_list_) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(search_results_list_, &rect);
        screen_pt.x = rect.left + 24;
        screen_pt.y = rect.top + 24;
      }
      ShowSearchResultContextMenu(screen_pt);
      return 0;
    }
    break;
  }
  case WM_NOTIFY:
    return HandleNotification(lparam);
  default:
    return std::nullopt;
  }
  return std::nullopt;
}

} // namespace regkit
