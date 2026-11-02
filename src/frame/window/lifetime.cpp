// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "frame/window/shortcut_bindings.h"

#include "ui/dialog_layout.h"
#include "win32/translation.h"

#include <win32ui/dialog/edit.h>
#include <win32ui/dialog/file_dialog.h>
#include <win32ui/dialog/tooltip.h>

#include <dwmapi.h>

namespace regkit
{
using namespace window_detail;

bool MainWindow::Impl::OnCreate()
{
    if (util::IsProcessPrivileged() && util::IsUacEnabled())
    {
        ChangeWindowMessageFilterEx(hwnd_, WM_COPYDATA, MSGFLT_ALLOW, nullptr);
    }
    updates_.Attach(hwnd_, [this](const std::wstring& text) { SetStatusMessage(text); });
    win32ui::SetMissingDesktopPrompt([](HWND owner, const wchar_t* path) {
        return ui::PromptKeyChoice(owner, util::Tr(L"SYSTEM has no Desktop folder, so Windows reports \"Location is not available\" in file dialogs. Create this folder to prevent the error?"), path, util::Tr(L"Create Folder"), util::Tr(L"Create"), L"", util::Tr(L"Cancel")) == IDYES;
    });
    ui_font_ = CreateUIFont();
    custom_font_ = win32ui::MessageFontLogFont(0);
    LoadSettings();
    if (theme_mode_ == win32ui::ThemeMode::custom)
    {
        LoadThemePresets();
    }
    ApplySavedWindowPlacement();
    if (theme_mode_ != win32ui::ThemeMode::custom || !ApplyThemePresetByName(settings_.theme_preset, false))
    {
        win32ui::SetThemeMode(theme_mode_);
        ApplySystemTheme();
    }
    UpdateUIFont();
    BuildMenus();
    BuildAccelerators();

    toolbar_hooks_.frame = appearance::DialogThemeHooks();
    toolbar_hooks_.icon = [](void* context, int command, UINT dpi) { return static_cast<const MainWindow::Impl*>(context)->ToolbarIconFor(command, dpi); };
    toolbar_hooks_.tip = [](void* context, int command) { return static_cast<const MainWindow::Impl*>(context)->CommandTooltipText(command); };
    toolbar_hooks_.shortcut = [](int command) { return frame::ShortcutText(command); };
    toolbar_hooks_.context = this;
    std::vector<win32ui::ToolbarButton> buttons;
    for (const int command : {cmd::kRegistryLocal, cmd::kRegistryNetwork, cmd::kRegistryOffline, 0, cmd::kEditFind, cmd::kEditReplace, 0, cmd::kEditUndo, cmd::kEditRedo, cmd::kEditCopy, cmd::kEditPaste, cmd::kEditDelete, cmd::kViewRefresh, 0, cmd::kNavBack, cmd::kNavForward, cmd::kNavUp})
    {
        // tooltip also supplies the accessible button name
        buttons.push_back({command, command ? CommandTooltipText(command) : std::wstring()});
    }
    toolbar_ = win32ui::CreateToolbar(hwnd_, kToolbarId, util::Tr(L"Toolbar"), &toolbar_hooks_, buttons);

    browse::CreateRequest browse_request;
    browse_request.parent = hwnd_;
    browse_request.instance = instance_;
    browse_request.address_id = kAddressEditId;
    browse_request.go_id = kAddressGoId;
    browse_request.filter_id = kFilterEditId;
    browse_request.tree_id = kTreeId;
    browse_request.values_id = kValueListId;
    browse_request.address_proc = AddressEditProc;
    browse_request.address_subclass_id = kAddressSubclassId;
    browse_request.filter_proc = FilterEditProc;
    browse_request.filter_subclass_id = kFilterSubclassId;
    browse_request.tree_proc = TreeViewProc;
    browse_request.tree_subclass_id = kTreeViewSubclassId;
    browse_request.callback_context = reinterpret_cast<DWORD_PTR>(this);
    if (!browse_.Create(browse_request))
    {
        return false;
    }
    // one line each, so a paste with line breaks isnt refused
    win32ui::KeepSingleLine(browse_.address());
    win32ui::KeepSingleLine(browse_.filter());
    win32ui::SelectAllOnFocus(browse_.address());
    AttachMainList(browse_.values().hwnd());
    // hidden RegEdit tree supports tools navigating through tree messages
    regedit_compat_tree_.Create(hwnd_, instance_, kRegEditCompatTreeId, false, false);
    if (!regedit_compat_tree_.hwnd())
    {
        return false;
    }
    regedit_compat_tree_.SetRegEditLayout(true);
    regedit_compat_tree_.SetRootLabel(L"Computer");
    regedit_compat_tree_.PopulateRoots(RegistryStore::DefaultRoots(false));
    if (!SetWindowSubclass(regedit_compat_tree_.hwnd(), TreeViewProc, kTreeViewSubclassId, reinterpret_cast<DWORD_PTR>(this)))
    {
        return false;
    }
    TreeView_SelectItem(regedit_compat_tree_.hwnd(), TreeView_GetRoot(regedit_compat_tree_.hwnd()));
    SetWindowPos(regedit_compat_tree_.hwnd(), HWND_TOP, -32000, -32000, 1, 1, SWP_NOACTIVATE);
    ShowWindow(regedit_compat_tree_.hwnd(), SW_HIDE);

    value_tooltip_ = win32ui::AddTooltip(hwnd_, browse_.go_button(), util::Tr(L"Go"));
    // owner drawn buttons show only their glyph, the text names them for screen readers
    SetWindowTextW(browse_.go_button(), util::Tr(L"Go"));
    win32ui::ApplyThemeToWindow(value_tooltip_);

    tab_hooks_ = appearance::TabStripThemeHooks();
    tab_hooks_.closable = [](void* context, int index) { return static_cast<MainWindow::Impl*>(context)->CanCloseTab(index); };
    tab_hooks_.close = [](void* context, int index) { static_cast<MainWindow::Impl*>(context)->CloseTab(index); };
    tab_hooks_.context = this;
    tab_ = win32ui::CreateTabStrip(hwnd_, kTabId, util::Tr(L"Tabs"), &tab_hooks_);
    ApplyFont(tab_, ui_font_);

    tree_header_ = CreateWindowExW(0, L"STATIC", util::Tr(L"Key Tree"), WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | SS_LEFT | SS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTreeHeaderId)), instance_, nullptr);
    tree_close_btn_ = CreateWindowExW(0, L"BUTTON", util::Tr(L"Close"), WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTreeHeaderCloseId)), instance_, nullptr);
    filter_clear_btn_ =
        CreateWindowExW(0, L"BUTTON", util::Tr(L"Clear Filter"), WS_CHILD | WS_CLIPSIBLINGS | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kFilterClearId)), instance_, nullptr);
    SetWindowPos(tree_close_btn_, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

    ConfigureTree(browse_.tree());
    search_results_list_ = CreateWindowExW(
        0,
        WC_LISTVIEWW,
        L"",
        WS_CHILD | WS_CLIPSIBLINGS | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(kSearchResultsListId)),
        instance_,
        nullptr
    );
    LoadTabs();

    history_label_ = CreateWindowExW(
        0,
        L"STATIC",
        util::Tr(L"History"),
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | SS_LEFT | SS_OWNERDRAW,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryLabelId)),
        instance_,
        nullptr
    );
    history_close_btn_ =
        CreateWindowExW(0, L"BUTTON", util::Tr(L"Close"), WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryHeaderCloseId)), instance_, nullptr);
    SetWindowPos(history_close_btn_, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    for (HWND button : {tree_close_btn_, filter_clear_btn_, history_close_btn_})
    {
        SetWindowSubclass(button, CloseButtonProc, kCloseButtonSubclassId, reinterpret_cast<DWORD_PTR>(this));
    }
    splitter_hooks_.size = [](void* context, HWND splitter) {
        const auto* impl = static_cast<MainWindow::Impl*>(context);
        return splitter == impl->tree_splitter_ ? impl->tree_shown_width_ : impl->history_shown_height_;
    };
    splitter_hooks_.resize = [](void* context, HWND splitter, int size, bool done) { static_cast<MainWindow::Impl*>(context)->ResizePane(splitter, size, done); };
    splitter_hooks_.context = this;
    tree_splitter_ = win32ui::CreateSplitter(hwnd_, kTreeSplitterId, util::Tr(L"Key Tree Splitter"), &splitter_hooks_, true);
    history_splitter_ = win32ui::CreateSplitter(hwnd_, kHistorySplitterId, util::Tr(L"History Splitter"), &splitter_hooks_, false, true);
    status_hooks_ = appearance::StatusBarThemeHooks();
    status_hooks_.message_ended = [](void* context) { static_cast<MainWindow::Impl*>(context)->UpdateStatus(); };
    status_hooks_.context = this;
    status_bar_ = win32ui::CreateStatusBar(hwnd_, kStatusBarId, util::Tr(L"Status Bar"), &status_hooks_);
    history_list_ = CreateWindowExW(
        0,
        WC_LISTVIEWW,
        L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryListId)),
        instance_,
        nullptr
    );

    AttachMainList(history_list_);
    AttachMainList(search_results_list_);
    // install after library subclass to run first, update external navigation before text search
    for (HWND list : {browse_.values().hwnd(), history_list_, search_results_list_})
    {
        SetWindowSubclass(list, ListViewProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this));
    }

    win32ui::AttachThemedBorder(tree_header_);
    win32ui::AttachThemedBorder(history_label_);
    win32ui::AttachThemedBorder(browse_.address());
    win32ui::AttachThemedBorder(browse_.go_button());
    UpdateGoButtonState();
    win32ui::AttachThemedBorder(browse_.filter());

    ApplyUIFontToControls();

    CreateValueColumns();
    CreateHistoryColumns();
    CreateSearchColumns();
    ApplyThemeToChildren();
    SetValueGridEnabled(settings_.show_value_grid, false);
    // grid buttons of main and dialog lists save the setting
    win32ui::SetListGridChanged([](void* context, bool shown) { static_cast<MainWindow::Impl*>(context)->SetValueGridEnabled(shown, true); }, this);
    UpdateToolbarStates();

    browse_.roots() = LocalRoots(0);
    browse_.tree().SetRegEditLayout(false);
    browse_.tree().SetRootLabel(TreeRootLabel(), TreeRootIcon());
    browse_.tree().PopulateRoots(browse_.roots());

    int initial_tab = tab_ ? TabCtrl_GetCurSel(tab_) : -1;
    if (initial_tab >= 0 && util::IsProcessPrivileged() && !IsLocalRegistryTabIndex(initial_tab))
    {
        const int local_tab = FindLocalRegistryTabIndex();
        if (local_tab < 0)
        {
            OpenLocalRegistryTab();
        }
        else
        {
            suppress_tab_change_ = true;
            TabCtrl_SetCurSel(tab_, local_tab);
            suppress_tab_change_ = false;
        }
        initial_tab = TabCtrl_GetCurSel(tab_);
    }
    if (initial_tab >= 0)
    {
        ApplyTabSelection(initial_tab);
    }
    else
    {
        SelectDefaultTreeItem();
    }
    StartValueListWorker();

    ApplyViewVisibility();
    ApplyAlwaysOnTop();
    UpdateStatus();
    return true;
}

void MainWindow::Impl::ConfigureTree(RegistryTree& tree)
{
    tree.SetIconResolver([this](const RegistryNode& node) { return KeyIconIndex(node, nullptr, nullptr); });
    tree.SetVirtualChildProvider(
        [this](const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) { AppendTraceChildren(node, existing_lower, out); }
    );
    tree.SetExpandAllCommand(hwnd_, cmd::kTreeExpandAll);
    win32ui::AttachThemedBorder(tree.hwnd());
    win32ui::ApplyThemeToWindow(tree.hwnd());
    if (tree_images_)
    {
        tree.SetImageList(tree_images_);
    }
    if (ui_font_)
    {
        ApplyFont(tree.hwnd(), ui_font_);
    }
}

void MainWindow::Impl::RunDeferredStartup()
{
    if (deferred_startup_complete_)
    {
        return;
    }
    deferred_startup_complete_ = true;
    const bool has_external_jump = !queued_external_jump_target_.empty();
    // external jumps and saved tabs override global tree state
    startup_tree_restore_pending_ = !has_external_jump && settings_.save_tree_state;

    EnableAddressAutoComplete();
    ReloadThemeIcons();
    FlushTreeRedraw();
    // cloak until all children paint to avoid a first frame with white panes
    BOOL cloak = TRUE;
    DwmSetWindowAttribute(hwnd_, DWMWA_CLOAK, &cloak, sizeof(cloak));
    ShowWindow(hwnd_, pending_show_cmd_);
    RedrawWindow(hwnd_, nullptr, nullptr, RDW_UPDATENOW | RDW_ALLCHILDREN);
    cloak = FALSE;
    DwmSetWindowAttribute(hwnd_, DWMWA_CLOAK, &cloak, sizeof(cloak));
    if (!file_errors_.empty() && !file_errors_posted_)
    {
        file_errors_posted_ = PostMessageW(hwnd_, frame::message_id::kShowFileErrors, 0, 0) != FALSE;
    }

    UpdateSearchResultsView();
    StartStartupCacheLoad(startup_tree_restore_pending_);
    ApplyQueuedExternalJump();
    StartTreeStateWorker();
    MarkTreeStateDirty();

    BuildMenus();
    UpdateStatus();
    if (settings_.auto_check_updates)
    {
        updates_.Check(true);
    }
}

void MainWindow::Impl::StartStartupCacheLoad(bool include_tree_state)
{
    StopStartupCacheLoad();
    int history_max_rows = history_max_rows_;
    int history_sort_column = history_sort_column_;
    bool history_sort_ascending = history_sort_ascending_;
    const HWND hwnd = hwnd_;
    // runs here without a thread, history is only written once it loaded
    startup_cache_session_.StartOrRun(L"StartupCacheThread", [this, include_tree_state, history_max_rows, history_sort_column, history_sort_ascending, hwnd](uint64_t generation, const std::atomic_bool& cancel) {
        auto payload = std::make_unique<StartupCachePayload>();
        payload->generation = generation;

        std::wstring comments_content;
        std::wstring reason;
        const std::wstring defaults_path =
            util::JoinPath(util::ModuleDirectory(), L"assets\\comments\\default-comments.jsonc");
        if (!util::ReadStateFile(defaults_path, &comments_content, &reason, util::kMaxCommentFileBytes) ||
            (!comments_content.empty() && !changes::ParseComments(comments_content, &payload->default_comments, &reason)))
        {
            payload->default_comments.clear();
            payload->file_errors.emplace_back(defaults_path, std::move(reason));
        }
        else if (!changes::ValidateCatalog(payload->default_comments))
        {
            payload->default_comments.clear();
        }
        const std::wstring comments_path = CommentsPath();
        if (!comments_path.empty() &&
            (!util::ReadStateFile(comments_path, &comments_content, &payload->comments_error, util::kMaxCommentFileBytes) ||
             (!comments_content.empty() && !changes::ParseComments(comments_content, &payload->user_comments, &payload->comments_error))))
        {
            payload->comments_unreadable = true;
        }
        payload->comments_loaded = true;
        if (cancel.load())
        {
            return;
        }

        std::wstring history_path = HistoryCachePath();
        std::wstring history_content;
        std::wstring history_error;
        if (!history_path.empty() && util::ReadStateFile(history_path, &history_content, &history_error, util::kMaxHistoryFileBytes))
        {
            changes::HistoryDocument document = changes::ParseHistory(history_content, &history_error);
            changes::ChangeHistory history;
            history.Replace(std::move(document.entries), static_cast<size_t>(history_max_rows));
            history.Sort(history_sort_column, history_sort_ascending);
            payload->history_entries = std::move(history.entries());
        }
        if (!history_error.empty())
        {
            payload->history_unreadable = true;
            payload->file_errors.emplace_back(history_path, std::move(history_error));
        }
        payload->history_loaded = true;
        if (cancel.load())
        {
            return;
        }

        if (include_tree_state)
        {
            std::wstring tree_error;
            if (!workspace::LoadTreeState(TreeStatePath(), &payload->tree_state, &tree_error) && !tree_error.empty())
            {
                payload->file_errors.emplace_back(TreeStatePath(), std::move(tree_error));
            }
            payload->tree_state_loaded = true;
        }

        if (cancel.load())
        {
            return;
        }
        if (hwnd && IsWindow(hwnd))
        {
            win32ui::PostPayload(hwnd, frame::message_id::kStartupCacheReady, static_cast<WPARAM>(generation), payload);
        }
    });
}

void MainWindow::Impl::StopStartupCacheLoad()
{
    startup_cache_session_.CancelAndJoin();
}

void MainWindow::Impl::ApplyStartupCachePayload(std::unique_ptr<StartupCachePayload> owned)
{
    if (!owned)
    {
        return;
    }
    if (!startup_cache_session_.IsCurrent(owned->generation))
    {
        return;
    }
    startup_cache_session_.Join();

    if (owned->comments_loaded)
    {
        default_comments_.Clear();
        default_comments_.Merge(owned->default_comments);
        changes::ValueComments loaded;
        loaded.Merge(owned->user_comments);
        // comments added during startup override older disk copies
        loaded.Merge(value_comments_.rules());
        value_comments_ = std::move(loaded);
        comments_unreadable_ = owned->comments_unreadable;
        comments_loaded_ = true;
        if (comments_unreadable_)
        {
            ReportFileError(CommentsPath(), owned->comments_error, util::Tr(L"The comments file couldn't be read, so comment changes won't be saved until it is fixed or removed."));
        }
        RefreshValueListComments();
        UpdateSearchResultsView();
    }

    for (const auto& [path, reason] : owned->file_errors)
    {
        ReportFileError(path, reason);
    }
    history_unreadable_ = history_unreadable_ || owned->history_unreadable;
    if (owned->history_loaded)
    {
        std::vector<HistoryEntry> pending_session_entries;
        if (!history_loaded_ && !change_history_.entries().empty())
        {
            pending_session_entries = change_history_.entries();
        }
        if (!change_history_.entries().empty())
        {
            owned->history_entries.insert(owned->history_entries.end(), change_history_.entries().begin(), change_history_.entries().end());
        }
        change_history_.Replace(std::move(owned->history_entries), static_cast<size_t>(history_max_rows_));
        change_history_.Sort(history_sort_column_, history_sort_ascending_, history_sort_earlier_);
        history_loaded_ = true;
        AppendHistoryCache(pending_session_entries);
        RebuildHistoryList();
    }

    if (owned->tree_state_loaded && startup_tree_restore_pending_)
    {
        applying_startup_tree_restore_ = true;
        RestoreTreeState(std::move(owned->tree_state));
        applying_startup_tree_restore_ = false;
        startup_tree_restore_pending_ = false;
    }
}

void MainWindow::Impl::OnDestroy()
{
    if (IsWindow(key_handles_window_))
    {
        DestroyWindow(key_handles_window_);
    }
    key_handles_window_ = nullptr;
    win32ui::SetListGridChanged(nullptr, nullptr);
    if (hwnd_)
    {
        RemovePropW(hwnd_, kRegKitWindowProperty);
    }
    EndJumpUiBatch();
    // stop workers before releasing controls and resources they may reference
    StopStartupCacheLoad();
    StopReplace();
    compare_session_.CancelAndJoin();
    StopTraceParseSessions();
    StopDefaultParseSessions();
    StopRegFileParseSessions();
    StopTraceLoadWorker();
    StopDefaultLoadWorker();
    StopValueListWorker();
    const bool clearing_tree_state = cache_clear_on_close_ && (*cache_clear_on_close_ == CacheKind::kAll ||
                                                               *cache_clear_on_close_ == CacheKind::kTreeState);
    if (clearing_tree_state)
    {
        tree_state_saver_.Stop();
    }
    else
    {
        StopTreeStateWorker();
    }
    for (auto& tab : search_tabs_)
    {
        tab.run.reset();
    }
    updates_.Cancel();
    DiscardWorkerMessages();
    if (!restart_on_close_ && !reset_settings_on_close_)
    {
        if (settings_.clear_tabs_on_exit)
        {
            ClearTabsCache();
        }
        else if (settings_.save_tab_kinds != 0 && !SaveTabs())
        {
            ui::ShowError(hwnd_, util::Tr(L"The open tabs couldn't be saved for the next session."));
        }
    }
    // release after saving, it clears hidden .reg trees used by capture
    for (auto& entry : tabs_)
    {
        if (entry.kind == TabEntry::Kind::kRegFile)
        {
            ReleaseRegFileRoots(&entry);
        }
    }
    ClearHistoryItems(false);
    if (settings_.clear_history_on_exit)
    {
        std::wstring history_path = HistoryCachePath();
        if (!history_path.empty())
        {
            DeleteFileW(history_path.c_str());
        }
    }
    for (TabEntry& tab : tabs_)
    {
        tab.session.reset();
    }
    session_ = local_session_;
    if (ui_font_ && ui_font_owned_)
    {
        DeleteObject(ui_font_);
    }
    ui_font_ = nullptr;
    ui_font_owned_ = false;
    if (tree_images_)
    {
        ImageList_Destroy(tree_images_);
        tree_images_ = nullptr;
    }
    if (list_images_)
    {
        ImageList_Destroy(list_images_);
        list_images_ = nullptr;
    }
    if (address_go_icon_)
    {
        DestroyIcon(address_go_icon_);
        address_go_icon_ = nullptr;
    }
}

void MainWindow::Impl::DiscardWorkerMessages()
{
    if (!hwnd_)
    {
        return;
    }
    MSG message = {};
    const UINT payload_messages[] = {frame::message_id::kTraceLoadReady, frame::message_id::kDefaultLoadReady, frame::message_id::kStartupCacheReady, frame::message_id::kRegFileLoadReady, frame::message_id::kTraceParseBatch, frame::message_id::kDefaultParseBatch, frame::message_id::kValueListReady, frame::message_id::kReplaceReady, frame::message_id::kCompareReady, frame::message_id::kValuePreviewReady, frame::message_id::kSearchPreviewReady, frame::message_id::kSearchSortReady, frame::message_id::kSearchTabLoadReady, frame::message_id::kUpdateCheckReady};
    for (const UINT id : payload_messages)
    {
        while (PeekMessageW(&message, hwnd_, id, id, PM_REMOVE))
        {
            win32ui::TakePayload<win32ui::PostedPayload>(message.lParam).reset();
        }
    }
    while (PeekMessageW(&message, hwnd_, frame::message_id::kExternalHandoff, frame::message_id::kExternalHandoff, PM_REMOVE))
    {
        delete reinterpret_cast<std::wstring*>(message.lParam);
    }
}

}
