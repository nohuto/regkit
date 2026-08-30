// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

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
      int sel = TabCtrl_GetCurSel(tab_);
      if (sel >= 0 && static_cast<size_t>(sel) < tabs_.size() && tabs_[static_cast<size_t>(sel)].kind == TabEntry::Kind::kRegistry) {
        if (regedit_compatibility_mode_ && regedit_compat_edit_) {
          SetFocus(regedit_compat_edit_);
          SendMessageW(regedit_compat_edit_, EM_SETSEL, 0, -1);
        } else {
          SetFocus(browse_.tree().hwnd());
        }
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
      EndSplitterDrag(true);
      return 0;
    }
    if (history_splitter_dragging_) {
      EndHistorySplitterDrag(true);
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
      EndSplitterDrag(false);
      return 0;
    }
    if (history_splitter_dragging_) {
      EndHistorySplitterDrag(false);
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
  case frame::message_id::kSearchProgress:
  case frame::message_id::kSearchFinished:
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
    return HandleValueWorkerMessage(message, wparam, lparam);
  default:
    return std::nullopt;
  }
}

std::optional<LRESULT> MainWindow::Impl::HandleSearchWorkerMessage(UINT message,
                                                                   WPARAM wparam,
                                                                   LPARAM lparam) {
  switch (message) {
  case frame::message_id::kSearchResults: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    std::vector<PendingSearchResult> pending;
    {
      std::lock_guard<std::mutex> lock(search_mutex_);
      if (search_pending_.empty()) {
        search_posted_.store(false);
        return 0;
      }
      pending.swap(search_pending_);
      search_posted_.store(false);
    }
    bool should_refresh = false;
    if (IsSearchTabIndex(active_search_tab_index_)) {
      int index = SearchIndexFromTab(active_search_tab_index_);
      if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size()) {
        uint64_t start_tick = GetTickCount64();
        size_t processed = 0;
        size_t stop_at = pending.size();
        for (size_t i = 0; i < pending.size(); ++i) {
          auto& item = pending[i];
          if (item.generation != generation) {
            continue;
          }
          search_tabs_[static_cast<size_t>(index)].results.push_back(std::move(item.result));
          ++processed;
          if (processed >= kSearchResultsBatch || (GetTickCount64() - start_tick) >= kSearchResultsMaxMs) {
            stop_at = i + 1;
            break;
          }
        }
        auto& tab = search_tabs_[static_cast<size_t>(index)];
        if (processed > 0 && tab.sort_column >= 0) {
          tab.sort_dirty = true;
        }
        if (stop_at < pending.size()) {
          std::vector<PendingSearchResult> remainder;
          remainder.reserve(pending.size() - stop_at);
          for (size_t i = stop_at; i < pending.size(); ++i) {
            if (pending[i].generation != generation) {
              continue;
            }
            remainder.push_back(std::move(pending[i]));
          }
          if (!remainder.empty()) {
            std::lock_guard<std::mutex> lock(search_mutex_);
            std::vector<PendingSearchResult> merged;
            merged.reserve(remainder.size() + search_pending_.size());
            for (auto& item : remainder) {
              merged.push_back(std::move(item));
            }
            for (auto& item : search_pending_) {
              merged.push_back(std::move(item));
            }
            search_pending_.swap(merged);
            if (!search_posted_.exchange(true)) {
              PostMessageW(hwnd_, frame::message_id::kSearchResults, static_cast<WPARAM>(generation), 0);
            }
          }
        }
        if (TabCtrl_GetCurSel(tab_) == active_search_tab_index_) {
          should_refresh = true;
        }
      }
    }
    if (should_refresh) {
      uint64_t now = GetTickCount64();
      if (now - search_last_refresh_tick_ >= kSearchResultsRefreshMs) {
        search_last_refresh_tick_ = now;
        UpdateSearchResultsView();
        UpdateStatus();
      }
    }
    return 0;
  }
  case frame::message_id::kSearchProgress: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    search_progress_posted_.store(false);
    UpdateStatus();
    return 0;
  }
  case frame::message_id::kSearchFinished: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    bool results_pending = search_posted_.load();
    {
      std::lock_guard<std::mutex> lock(search_mutex_);
      results_pending = results_pending || !search_pending_.empty();
    }
    if (results_pending) {
      PostMessageW(hwnd_, frame::message_id::kSearchFinished, wparam, 0);
      return 0;
    }
    search_session_.Join();
    search_running_ = false;
    for (auto& search_tab : search_tabs_) {
      if (search_tab.generation == generation && search_tab.sort_dirty) {
        SortSearchTabResults(&search_tab);
        break;
      }
    }
    if (search_session_.IsCurrent(generation) &&
        search_start_tick_ != 0) {
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
      InvalidateRect(list_hwnd, nullptr, TRUE);
    }
    value_list_loading_ = false;
    UpdateStatus();
    StartPendingValueListRename();
    if (!pending_external_value_name_.empty() && browse_.current_node()) {
      std::wstring current_path = registry_path::Build(*browse_.current_node());
      if (EqualsInsensitive(current_path, pending_external_value_key_path_)) {
        SelectValueByName(pending_external_value_name_);
        pending_external_value_key_path_.clear();
        pending_external_value_name_.clear();
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
    if (wparam == kRegeditCompatApplyTimerId) {
      KillTimer(hwnd_, kRegeditCompatApplyTimerId);
      ApplyPendingRegeditCompatNavigation();
      return 0;
    }
    break;
  case WM_COPYDATA: {
    auto* data = reinterpret_cast<const COPYDATASTRUCT*>(lparam);
    if (!data) {
      return 0;
    }
    if (data->dwData == kRegeditCompatActivateCopyDataId) {
      ActivateRegeditCompatibilityMode();
      ShowWindow(hwnd_, SW_RESTORE);
      SetForegroundWindow(hwnd_);
      FocusAddressBarForExternalJump(true);
      return TRUE;
    }
    if (data->dwData != kExternalJumpCopyDataId || !data->lpData || data->cbData < sizeof(wchar_t)) {
      return 0;
    }
    ActivateRegeditCompatibilityMode();
    size_t length = data->cbData / sizeof(wchar_t);
    const wchar_t* text = reinterpret_cast<const wchar_t*>(data->lpData);
    std::wstring target(text, text + length);
    while (!target.empty() && target.back() == L'\0') {
      target.pop_back();
    }
    if (target.empty()) {
      return 0;
    }
    // External jump requests are authoritative; clear any in-flight compat edit debounce.
    KillTimer(hwnd_, kRegeditCompatApplyTimerId);
    regedit_compat_pending_key_path_.clear();
    regedit_compat_pending_value_name_.clear();
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
    if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kAddressGoId) {
      NavigateToAddress();
      return 0;
    }
    if (HIWORD(wparam) == EN_CHANGE && LOWORD(wparam) == kFilterEditId) {
      wchar_t buffer[256] = {};
      GetWindowTextW(browse_.filter(), buffer, static_cast<int>(_countof(buffer)));
      bool needs_full_data = buffer[0] != L'\0';
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
