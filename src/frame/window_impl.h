// Copyright (C) 2026 nohuto
// SPDX License Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/main_window.h"
#include "frame/tools/update_checker.h"
#include "win32/windows_config.h"

#include <windows.h>

#include <commctrl.h>
#include <uxtheme.h>

#include <atomic>
#include <condition_variable>
#include <deque>
#include <memory>
#include <mutex>
#include <optional>
#include <shared_mutex>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "browse/browse_pane.h"
#include "changes/change_history.h"
#include "changes/key_snapshot.h"
#include "changes/undo_stack.h"
#include "changes/value_comments.h"
#include "defaults/default_data.h"
#include "dialogs/query_dialog.h"
#include "dialogs/replace_dialog.h"
#include "dialogs/trace_dialog.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/virtual_registry.h"
#include "search/compare.h"
#include "search/search.h"
#include "trace/trace_data.h"
#include "ui/presets.h"
#include "ui/splitter.h"
#include "ui/tab_strip.h"
#include "ui/theme.h"
#include "ui/toolbar.h"
#include "win32/file_dialog.h"
#include "win32/handle_owner.h"
#include "win32/process_rights.h"
#include "win32/translation.h"
#include "work/key_watcher.h"
#include "work/session.h"
#include "workspace/dialog_state.h"
#include "workspace/favorites.h"
#include "workspace/recent_items.h"
#include "workspace/settings.h"
#include "workspace/tree_state.h"

namespace regkit
{

class MainWindow::Impl
{
  public:
    ~Impl();
    bool Create(HINSTANCE instance);
    void Show(int cmd_show);
    bool OpenRegFileTab(const std::wstring& path, bool force_new_tab = false);
    bool OpenFile(const std::wstring& path);
    void OpenSavedFile(const std::wstring& path, win32::OpenAfter open_after);
    void StartRegFileParse(const std::wstring& path, const std::wstring& session_key);
    bool TranslateAccelerator(const MSG& msg);
    void QueueExternalJump(const std::wstring& target);

  private:
    friend class RegistryAddressEnum;
    enum class RegistryMode
    {
        kLocal,
        kRemote,
        kOffline,
    };

    // the local registry, one remote machine or a set of offline hives; local tabs share one, the others own theirs
    struct RegistrySession
    {
        RegistrySession() = default;
        RegistrySession(const RegistrySession&) = delete;
        RegistrySession& operator=(const RegistrySession&) = delete;
        ~RegistrySession();

        RegistryMode mode = RegistryMode::kLocal;
        REGSAM view = 0;
        std::wstring remote_machine;
        HKEY remote_hklm = nullptr;
        HKEY remote_hku = nullptr;
        bool remote_service_started = false;
        DWORD remote_service_start_type = 0;
        HKEY offline_root = nullptr;
        std::vector<HKEY> offline_roots;
        std::wstring offline_mount;
        std::vector<std::wstring> offline_root_labels;
        std::vector<std::wstring> offline_root_paths;
        std::wstring offline_root_name;
        bool offline_dirty = false;
        std::vector<RegistryRootEntry> roots;
        changes::UndoStack undo;
    };
    enum class CacheKind
    {
        kAll,
        kTabs,
        kHistory,
        kSearchHistory,
        kDialogState,
        kTreeState,
        kTemporary,
    };
    struct TabEntry;
    struct SearchTab;
    struct SearchTabLoadPayload;
    struct TraceParseSession;
    struct DefaultParseSession;
    struct StartupCachePayload : work::MoveOnly
    {
        uint64_t generation = 0;
        std::vector<HistoryEntry> history_entries;
        std::vector<changes::CommentRule> user_comments;
        std::vector<changes::CommentRule> default_comments;
        bool comments_unreadable = false;
        workspace::TreeState tree_state;
        bool history_loaded = false;
        bool comments_loaded = false;
        bool tree_state_loaded = false;
    };
    struct ReplacePayload : work::MoveOnly
    {
        struct Change
        {
            changes::UndoOperation undo;
            HistoryEntry history;
        };

        uint64_t generation = 0;
        // the tab session the replace ran in, its undo goes there even after a tab switch
        std::weak_ptr<RegistrySession> session;
        std::optional<search::Source> source;
        std::vector<Change> changes;
        int failures = 0;
        int partial_renames = 0;
        int rejected = 0;
        bool cancelled = false;
    };
    struct ComparePayload : work::MoveOnly
    {
        uint64_t generation = 0;
        std::wstring error;
        std::vector<search::compare::Row> rows;
        search::compare::RowFilter filter = search::compare::RowFilter::kDifferences;
        std::vector<search::Source> sources;
        std::wstring unreadable;
    };

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam);
    static LRESULT CALLBACK AddressEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
    static LRESULT CALLBACK FilterEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
    static LRESULT CALLBACK ListViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
    static LRESULT CALLBACK TreeViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);

    LRESULT HandleMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleLifecycleMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleLayoutInputMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleSearchWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleLoadWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleRegFileWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleTraceWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleDefaultWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleValueWorkerMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleExternalMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleAppearanceMessage(UINT message, WPARAM wparam, LPARAM lparam);
    std::optional<LRESULT> HandleBrowseMessage(UINT message, WPARAM wparam, LPARAM lparam);
    LRESULT HandleNotification(LPARAM lparam);
    LRESULT HandleTooltipNotification(NMHDR* header, LPARAM lparam);
    std::wstring ListCellFieldText(HWND list, int item, int display_subitem);
    bool ListCellTooltipText(std::wstring* out);
    std::wstring SearchCellFieldText(const search::Result& result, int subitem) const;
    LRESULT HandleToolbarNotification(NMHDR* header, LPARAM lparam);
    LRESULT HandleTabNotification(NMHDR* header, LPARAM lparam);
    LRESULT HandleTreeNotification(NMHDR* header, LPARAM lparam);
    LRESULT HandleHeaderNotification(NMHDR* header, LPARAM lparam);
    LRESULT HandleValueNotification(NMHDR* header, LPARAM lparam);
    LRESULT HandleHistoryNotification(NMHDR* header, LPARAM lparam);
    LRESULT HandleSearchNotification(NMHDR* header, LPARAM lparam);
    LRESULT HandleSearchListCustomDraw(NMLVCUSTOMDRAW* draw);
    search::Result* SearchResultAt(int item);
    std::wstring SearchRowKeyPath(int tab_index, int item) const;
    size_t SearchRowCount(int tab_index) const;
    bool OnCreate();
    void RunDeferredStartup();
    void OnDestroy();
    void DiscardWorkerMessages();
    void OnSize(int width, int height);
    void OnPrintClient(HDC hdc);
    void ApplyThemeToChildren();
    void ApplySystemTheme();
    void LoadThemePresets();
    void SaveThemePresets() const;
    void UpdateThemePresets(const std::vector<ThemePreset>& presets, const std::wstring& active_name, bool apply_now);
    bool ApplyThemePresetByName(const std::wstring& name, bool persist);
    void ShowThemePresetsDialog();
    void ApplyAlwaysOnTop();
    void UpdateUIFont();
    void ApplyUIFontToControls();
    void LayoutControls(int width, int height);
    void LayoutContent(bool dragging);
    int TextRowHeight(int nominal, int padding = 6) const;
    void DragSplitter(ui::Splitter* splitter, int* size, POINT point);
    void BuildImageLists();
    void ReloadThemeIcons();
    bool ShouldUseLightIcons() const;
    std::wstring ResolveIconDir(bool use_light) const;
    std::wstring ResolveIconPath(const wchar_t* filename) const;
    void ApplyGridToolbarIcons();
    void LayoutValueGridToolbar();
    void SetValueGridEnabled(bool enabled, bool persist);
    ToolbarIcon MakeToolbarIcon(const wchar_t* filename, int resource_id) const;
    void CreateValueColumns();
    void CreateHistoryColumns();
    void CreateSearchColumns();
    void ApplyValueColumns();
    void ApplyHistoryColumns();
    void ApplySearchColumns(bool compare);
    void RefreshCompareColumnTitles();
    bool IsCompareResultColumnAvailable() const;
    void UpdateValueListForNode(RegistryNode* node);
    void AttachHeader(HWND header);
    void ResetValueFilter();
    void FocusFirstValue();
    void EnsureValueRowData(ListRow* row);
    void StartValuePreviewWorker();
    void QueueValuePreviews(int first, int last);
    void QueueSearchPreviews(int first, int last);
    void QueueSearchSort(SearchTab* tab, const util::SortKeys& earlier);
    void StartSearchSortWorker();
    void StartSearchTabLoadWorker();
    void ApplySearchTabLoad(std::unique_ptr<SearchTabLoadPayload> owned);
    void FinishSearch(SearchTab* tab);
    SearchTab* SearchTabByGeneration(uint64_t generation);
    SearchTab* SearchTabByLoadGeneration(uint64_t generation);
    SearchTab* ShownSearchTab();
    void StartSearchPreviewWorker();
    void StartValueListWorker();
    void StopValueListWorker();
    void StartTraceLoadWorker();
    void StopTraceLoadWorker();
    void StartTraceParseThread(TraceParseSession* session);
    void MergeTraceEntries(TraceParseSession* session, const std::vector<KeyValueDialogEntry>& entries, std::unordered_set<std::wstring>* affected_keys);
    void StopTraceParseSessions();
    void StartDefaultLoadWorker();
    void StopDefaultLoadWorker();
    void StartDefaultParseThread(DefaultParseSession* session);
    void MergeDefaultEntries(DefaultParseSession* session, const std::vector<KeyValueDialogEntry>& entries, std::unordered_set<std::wstring>* affected_keys);
    void StopDefaultParseSessions();
    void StopRegFileParseSessions();
    static void StartTraceDialogLoad(HWND hwnd, void* context);
    static void StartDefaultDialogLoad(HWND hwnd, void* context);
    void UpdateAddressBar(RegistryNode* node);
    void UpdateGoButtonState();
    void EnableAddressAutoComplete();
    std::vector<std::wstring> BuildAddressSuggestions(const std::wstring& input) const;
    void UpdateStatus();
    void SetStatusMessage(const std::wstring& text);
    void SortValueList(int column, bool toggle);
    void SortHistoryList(int column, bool toggle);
    void SortSearchResults(int column, bool toggle);
    void ClearHistoryItems(bool delete_cache);
    void RemoveSelectedHistoryItems();
    void RebuildHistoryList();
    void RefreshHistory();
    void ScheduleValueListRename(LPARAM kind, const std::wstring& name);
    void StartPendingValueListRename();
    bool CollectSearchStartNodes(const SearchDialogResult& options, const std::wstring& registry_scope_path, std::vector<search::StartNode>* out_nodes, std::vector<search::Source>* out_sources, bool* out_remote);
    void StartSearch(const SearchDialogResult& options);
    void StartReplace(const ReplaceDialogResult& options);
    void ApplyReplacePayload(std::unique_ptr<ReplacePayload> owned);
    void CommitReplacePayload(std::unique_ptr<ReplacePayload> payload, bool show_failures);
    void StopReplace();
    void CancelSearch(SearchTab* tab);
    bool IsSearchTabSelected() const;
    void UpdateSearchResultsView();
    void SortSearchTabResults(SearchTab* tab, bool keep_order = false);
    void CloseSearchTab(int tab_index);
    void SelectTabAfterClose(int closed_index, int previous_index);
    bool SwitchToLocalRegistry();
    bool SwitchToRemoteRegistry();
    bool ConnectRemoteRegistry(const std::wstring& machine, bool open_new_tab = false);
    bool OfferRemoteServiceStart(const std::wstring& machine, RegistrySession* session);
    void OfferRemoteServiceRestore(RegistrySession& session, bool hand_over = true);
    bool SwitchToOfflineRegistry();
    bool SaveOfflineRegistry(RegistrySession& session, bool choose_path = false);
    bool LoadOfflineRegistryFromPath(const std::wstring& path, bool open_new_tab);
    void ApplyRegistryRoots(const std::vector<RegistryRootEntry>& roots);
    void ReapplyRoots();
    std::vector<std::wstring> BuildVisibleTreePathParts(const std::wstring& path) const;
    std::wstring TreeRootLabel() const;
    int TreeRootIcon() const;
    void SelectDefaultTreeItem();
    void CaptureRegistryTabState(int index, bool tree_state);
    void ResetRegistryTreeState();
    void SuspendTreeRedraw();
    void FlushTreeRedraw();
    void RestoreRegistryTabState(int index);
    std::wstring LocalRegistryTabLabel(int index) const;
    void RefreshRegistryTabLabels();
    void ResetNavigationState();
    void UpdateTabText(const std::wstring& text);
    void UpdateTabWidth();
    void CloseTab(int tab_index);
    void ActivateTabIndex(int index);
    bool HandleTabCommand(int command_id);
    bool ConfirmCloseTab(int tab_index);
    bool ConfirmOfflineChanges(RegistrySession& session, const wchar_t* message);
    void MarkOfflineDirty();
    void MarkSessionDirty(RegistrySession& session);
    int AddRegistryTab(RegistryMode mode, const wchar_t* label);
    bool ActivateTabTree(int index);
    void ResumeTabTree(int index);
    void RestoreValueSelection(const TabEntry& entry);
    void ConfigureTree(RegistryTree& tree);
    void OpenLocalRegistryTab(REGSAM view = 0);
    void GoToOtherView();
    std::wstring VirtualStoreTarget() const;
    std::wstring GlobalKeyTarget() const;
    int CurrentRegistryTabIndex() const;
    void UpdateRegistryTabEntry(RegistryMode mode, const std::wstring& offline_path, const std::wstring& remote_machine);
    bool IsSearchTabIndex(int index) const;
    bool IsRegFileTabIndex(int index) const;
    bool IsRegFileTabSelected() const;
    int SearchIndexFromTab(int index) const;
    int FindFirstRegistryTabIndex() const;
    int FindLocalRegistryTabIndex() const;
    bool IsLocalRegistryTabIndex(int index) const;
    bool ActivateLocalRegistryTab();
    bool SearchResultOpensInNewTab() const;
    bool OpenSearchResultRow(int item, bool new_tab);
    bool OpenSelectedSearchResult(bool new_tab);
    RegistrySession* CurrentTabSession();
    void ShowSession(const std::shared_ptr<RegistrySession>& session);
    void UpdateUndoButtons();
    void NavigateToAddress();
    void PromptMissingPath(const std::wstring& path);
    void PromptMissingValue(const std::wstring& value_name);
    bool SelectTreePath(const std::wstring& path);
    std::wstring TreeNeighbourPath(HTREEITEM item);
    bool SelectChildKey(const RegistryNode& parent, const std::wstring& name);
    bool SelectValueByName(const std::wstring& name);
    void SelectValueAfterRefresh(const std::wstring& name);
    void RestoreValueSelection();
    void SelectValueWhenReady(const std::wstring& name);
    void SelectListRowAtIndex(HWND list, int index);
    void FocusAddressBarForExternalJump(bool defer_if_needed);
    void BeginJumpUiBatch();
    void EndJumpUiBatch();
    void ApplyTreeSelectionEffects(RegistryNode* node);
    void ApplyQueuedExternalJump();
    bool NavigateToResolvedExternalJump(const std::wstring& key_path, const std::wstring& value_name);
    bool NavigateToExternalJump(const std::wstring& target);
    void QueueCompatJump(const RegistryNode& node);
    void FlushExternalNavigation();
    bool ResolveJumpTarget(const std::wstring& target, std::wstring* key_path, std::wstring* value_name, bool* value_missing) const;
    bool LoadTraceFromFile(const std::wstring& label, const std::wstring& path, const trace::Selection* selection_override = nullptr);
    bool LoadBundledTrace(const std::wstring& label, const trace::Selection* selection_override = nullptr);
    std::wstring ResolveBundledTracePath(const std::wstring& label) const;
    bool LoadTraceFromPrompt();
    void ClearTrace();
    bool LoadDefaultFromFile(const std::wstring& label, const std::wstring& path);
    std::wstring ResolveBundledDefaultPath(const std::wstring& label) const;
    bool LoadDefaultFromPrompt();
    void ClearDefaults();
    void RefreshFavoritesCache();
    void RefreshRegEditFavoritesMenu();
    void RefreshBundledDefaultsCache();
    void BuildMenus();
    void UpdateMenuState(HMENU menu);
    void RefreshStorageMenuState(HMENU menu);
    void FillBitfieldMenu(HMENU menu);
    void BuildAccelerators();
    std::wstring CommandTooltipText(int command_id) const;
    bool HandleMenuCommand(int command_id);
    bool HandleDynamicCommand(int command_id);
    bool HandleFileCommand(int command_id);
    bool HandleViewCommand(int command_id);
    bool HandleTraceDefaultCommand(int command_id);
    bool HandleWorkspaceAppearanceCommand(int command_id);
    bool HandleWindowAppearanceCommand(int command_id);
    bool HandleLaunchHelpCommand(int command_id);
    bool HandleFavoritesCommand(int command_id);
    bool HandleNavigateClipboardCommand(int command_id);
    bool HandleClipboardCommand(int command_id);
    bool HandleEditToolsCommand(int command_id);
    bool HandleChangeHistoryCommand(int command_id);
    bool HandleRegistryNavigationCommand(int command_id);
    bool HandleMutationCommand(int command_id);
    bool HandleToolsCommand(int command_id);
    bool HandleCreateCommand(int command_id);
    bool HandleModifyCommand(int command_id);
    bool HandleRenameCommand(int command_id);
    bool HandleDeleteCommand(int command_id);
    bool EnsureWritable();
    bool ShowsRegFile() const;
    void PrepareMenusForOwnerDraw(HMENU menu);
    void ApplySimpleMode(HMENU menu) const;
    void OnMeasureMenuItem(MEASUREITEMSTRUCT* info);
    void OnDrawMenuItem(const DRAWITEMSTRUCT* info);
    void PaintMenuBarSeparator();
    void ShowHeaderMenu(HWND list, POINT screen_pt);
    ui::ColumnSet* ColumnSetFor(HWND list);
    void ApplyColumns(HWND list);
    void AppendHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data);
    void AppendHistoryEntry(HistoryEntry entry);
    // one cache write, sort & list update for a whole mass change
    void AppendHistoryEntries(std::vector<HistoryEntry> entries);
    void AppendHistoryEntries(std::vector<HistoryEntry> entries, const std::optional<search::Source>& source);
    std::optional<search::Source> HistorySource() const;
    static HistoryEntry ValueHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data, const RegistryNode& node, const std::wstring& value_name, HistoryEntry::RevertKind revert_kind, const RegistryValue* revert_value = nullptr);
    void AppendValueHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data, const RegistryNode& node, const std::wstring& value_name, HistoryEntry::RevertKind revert_kind, const RegistryValue* revert_value = nullptr);
    bool PrepareHistoryRevert(const HistoryEntry& entry, HistoryEntry* prepared) const;
    bool OpenHistoryTarget(const HistoryEntry& entry);
    bool RevertHistoryEntry(const HistoryEntry& entry);
    void ShowAddressContextMenu(HWND edit, POINT screen_pt);
    void ShowTreeContextMenu(HTREEITEM item, POINT screen_pt);
    bool CanOpenHiveFile(const RegistryNode& node);
    void ShowValueContextMenu(int index, POINT screen_pt);
    void ShowHistoryContextMenu(int index, POINT screen_pt);
    void ShowSearchResultContextMenu(int index, POINT screen_pt);
    void DrawPanelButton(const DRAWITEMSTRUCT* info);
    void ClearValueFilter(bool focus_values);
    void ShowPermissionsDialog(const RegistryNode& node);
    void ShowKeyInfoDialog(const RegistryNode& node);
    bool ShowResourceList(const RegistryValue& value);
    RegistryNode SelectedKeyNode() const;
    void CopyKeyPathAs(int command_id, const RegistryNode& node, const std::wstring& path);
    void WatchCurrentKey();
    void ApplyAutoRefresh();
    void RestoreHiveFile(const std::wstring& path);
    void ReplaceRegEdit(bool enable);
    void SetEditContextMenu(bool enable);
    void OpenHiveFileDir();
    std::wstring ResolveSelectedHiveFilePath();
    void RecordNavigation(const std::wstring& path);
    void NavigateBack();
    void NavigateForward();
    void NavigateUp();
    void UpdateNavigationButtons();
    void ApplyViewVisibility();
    void SelectTabIndex(int index);
    void ApplyTabSelection(int index);
    search::Source CurrentTabSource() const;
    search::Source TabSource(int index) const;
    void SyncRegFileTabSelection();
    void ResetHiveListCache();
    void EnsureHiveListLoaded();
    std::wstring LookupHivePath(const RegistryNode& node, bool* is_root);
    std::wstring LookupNativeHivePath(const std::wstring& nt_path, bool* is_root);
    int KeyIconIndex(const RegistryNode& node, bool* is_link, bool* is_hive_root);
    void AppendRealRegistryRoot(std::vector<RegistryRootEntry>* roots);
    std::vector<RegistryRootEntry> LocalRoots(REGSAM view);
    void ReloadLocalRoots();
    std::shared_ptr<RegistrySession> LocalViewSession(REGSAM view) const;
    void HandleTypeToSelectTree(wchar_t ch);
    void HandleTypeToSelectList(wchar_t ch);
    std::wstring NormalizeRegistryPath(const std::wstring& path) const;
    // the tree labels NormalizeRegistryPath strips, so a worker can normalize without the UI state
    std::vector<std::wstring> RegistryPathContexts() const;
    static std::wstring NormalizeRegistryPath(const std::wstring& path, const std::wstring& sid, const std::vector<std::wstring>& contexts);
    std::wstring FormatRegistryPath(const std::wstring& path, registry_path::Style style) const;
    bool FindNearestExistingPath(const std::wstring& path, std::wstring* nearest_path) const;
    bool CreateRegistryPath(const std::wstring& path);
    bool SelectAllInFocusedList();
    void CyclePaneFocus(bool forward);
    void FocusPane(HWND pane);
    bool InvertSelectionInFocusedList();
    bool IsCompareTabSelected() const;
    void StartCompareRegistries();
    void ApplyComparePayload(std::unique_ptr<ComparePayload> payload);
    void OpenSourceEntry(const search::Source& source, const std::wstring& path, const std::wstring& value_name, bool new_tab);
    int FindSourceTab(const search::Source& source) const;
    bool AppendHistoryCache(const std::vector<HistoryEntry>& entries);
    bool HistoryStaysInMemory() const;
    std::wstring CacheFolderPath() const;
    std::wstring HistoryCachePath() const;
    std::wstring TabsCachePath() const;
    std::wstring SearchTabCachePath(const std::wstring& file) const;
    void LoadTabs();
    bool SaveTabs();
    bool SaveSessionTabs();
    bool SaveTabState(const std::wstring& path, int kinds);
    int TabSaveKind(const TabEntry& entry) const;
    std::wstring SessionCachePath() const;
    bool ClearTabsCache();
    bool ClearCache(CacheKind kind, bool resume_tree_worker);
    bool EnsureSearchTabResultsLoaded(int search_index);
    void StartStartupCacheLoad(bool include_tree_state);
    void StopStartupCacheLoad();
    void ApplyStartupCachePayload(std::unique_ptr<StartupCachePayload> owned);
    bool SaveComments() const;
    bool ImportCommentsFromFile(const std::wstring& path, size_t* imported, std::wstring* error);
    void ShowFavoritesImported(size_t imported);
    bool ExportCommentsToFile(const std::wstring& path) const;
    void RefreshValueListComments();
    bool ApplyValueComments(std::vector<ListRow>* rows) const;
    std::wstring CommentsPath() const;
    std::wstring CommentKeyPath(const RegistryNode& node) const;
    bool EditComments(const std::vector<changes::CommentTarget>& targets);
    bool RestartAsAdmin();
    bool RestartAsUser();
    bool RestartCurrentInstance();
    bool RestartAfterCacheClear(CacheKind kind);
    bool RestartAfterSettingsReset();
    bool PrepareSessionHandover();
    bool SaveSessionForRestart();
    bool LaunchRestart(bool restore_session);
    bool RestartAsSystem();
    bool RestartAsTrustedInstaller();
    void LoadSettings();
    workspace::Settings CurrentSettings() const;
    void SaveSettings() const;
    std::wstring SettingsPath() const;
    std::wstring ActiveTracesPath() const;
    void SaveActiveTraces() const;
    std::wstring ActiveDefaultsPath() const;
    void SaveActiveDefaults() const;
    std::wstring TraceSettingsPath() const;
    void LoadTraceSettings();
    void SaveTraceSettings() const;
    bool AddTraceFromFile(const std::wstring& label, const std::wstring& path, const trace::Selection* selection_override, bool prompt_for_selection = true, bool update_ui = true);
    bool RemoveTraceByPath(const std::wstring& path);
    bool RemoveTraceByLabel(const std::wstring& label);
    bool HasActiveTraces() const;
    bool AddDefaultFromFile(const std::wstring& label, const std::wstring& path, bool show_error = true, bool prompt_for_selection = false, bool update_ui = true);
    bool SaveRegFileTab(int tab_index);
    bool ExportRegFileTab(int tab_index, const std::wstring& path);
    bool BuildRegFileContent(const TabEntry& entry, std::wstring* out) const;
    void ReleaseRegFileRoots(TabEntry* entry);
    bool RemoveDefaultByPath(const std::wstring& path);
    struct DefaultValueChoice
    {
        std::wstring label;
        bool present = false;
        DWORD type = REG_NONE;
        std::vector<BYTE> data;
    };
    std::vector<DefaultValueChoice> CollectDefaultChoices(const std::wstring& value_name) const;
    bool HandleResetDefaultCommand(int command_id);
    std::vector<DefaultValueChoice> SelectedValueDefaultChoices() const;
    HMENU BuildResetDefaultMenu(const std::vector<DefaultValueChoice>& choices) const;
    void RefreshResetDefaultMenu(HMENU menu);
    std::wstring TreeStatePath() const;
    void StartTreeStateWorker();
    void StopTreeStateWorker();
    void MarkTreeStateDirty();
    void CaptureTreeStateNow();
    bool CaptureLocalTreeState(workspace::TreeState* state, bool shown_only) const;
    void CaptureTreeState(const RegistryTree& tree, std::wstring* selected_path, std::vector<std::wstring>* expanded_paths) const;
    void RestoreTreeState(workspace::TreeState state);
    void ExpandTreePaths(const std::vector<std::wstring>& paths);
    HTREEITEM FindTreeItem(const std::wstring& path);
    void RefreshTreeItem(HTREEITEM item);
    void RefreshTreePath(const std::wstring& path);
    void RefreshTreeSelection();
    void RefreshWholeTree();
    void RefreshMatchingTreeNodes(HTREEITEM selected = nullptr);
    void UpdateSimulatedChain(HTREEITEM item);
    void ApplySavedWindowPlacement();
    LOGFONTW DefaultLogFont() const;
    void AddRecentTracePath(const std::wstring& path);
    void NormalizeRecentTraceList();
    void AddRecentDefaultPath(const std::wstring& path);
    void NormalizeRecentDefaultList();
    void AppendTraceChildren(const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) const;
    static std::wstring TracePathLowerForNode(const RegistryNode& node);
    static std::wstring SourceLookupPath(const RegistryNode& node);
    bool AllowTraceSimulation(const RegistryNode& node) const;

    struct ClipboardItem
    {
        enum class Kind
        {
            kNone,
            kValue,
            kKey,
        };

        Kind kind = Kind::kNone;
        RegistryNode source_parent;
        std::wstring name;
        RegistryValue value;
        changes::KeySnapshot key_snapshot;
    };

    void PushUndo(changes::UndoOperation operation);
    void PushUndo(std::vector<changes::UndoOperation> steps);
    void ClearRedo();
    enum class ReplayResult
    {
        kSuccess,
        kUnchanged,
        kPartial,
    };
    ReplayResult ApplyUndoOperation(const changes::UndoOperation& operation, bool redo, size_t* replayed = nullptr, bool* left_both_names = nullptr);
    bool SameNode(const RegistryNode& left, const RegistryNode& right) const;
    std::optional<std::wstring> MakeUniqueValueName(const RegistryNode& node, const std::wstring& base) const;
    std::wstring MakeUniqueKeyName(const RegistryNode& node, const std::wstring& base) const;
    bool ResolvePathToNode(const std::wstring& path, RegistryNode* node) const;
    static bool ResolvePathInRoots(const std::vector<RegistryRootEntry>& roots, const std::wstring& path, RegistryNode* node);
    std::vector<std::vector<RegistryRootEntry>> SearchSourceRoots(const SearchTab& tab);
    bool KeyPathExists(const std::wstring& path, RegistryNode* node) const;

    HINSTANCE instance_ = nullptr;
    HWND hwnd_ = nullptr;
    HFONT ui_font_ = nullptr;
    bool ui_font_owned_ = false;
    LOGFONTW custom_font_ = {};
    workspace::Settings settings_;
    HACCEL accelerators_ = nullptr;
    Toolbar toolbar_;
    HWND tab_ = nullptr;
    HWND tree_header_ = nullptr;
    HWND tree_close_btn_ = nullptr;
    HWND filter_clear_btn_ = nullptr;
    HWND history_label_ = nullptr;
    HWND history_close_btn_ = nullptr;
    HWND history_list_ = nullptr;
    HWND key_handles_window_ = nullptr;
    HWND status_bar_ = nullptr;
    HWND search_progress_ = nullptr;
    browse::Pane browse_;
    RegistryTree regedit_compat_tree_;
    HIMAGELIST tree_images_ = nullptr;
    HIMAGELIST list_images_ = nullptr;
    std::vector<int> value_column_subitems_;
    std::vector<int> search_column_subitems_;
    std::vector<int> history_column_subitems_;
    ui::ColumnSet history_columns_;
    ui::ColumnSet search_columns_;
    ui::ColumnSet compare_columns_;
    std::vector<std::wstring> compare_column_titles_;
    bool compare_columns_active_ = false;
    bool compare_result_column_active_ = false;
    int history_sort_column_ = 0;
    bool history_sort_ascending_ = false;
    util::SortKeys history_sort_earlier_;
    int history_max_rows_ = 500;
    changes::ChangeHistory change_history_;
    std::shared_ptr<RegistrySession> local_session_ = std::make_shared<RegistrySession>();
    std::shared_ptr<RegistrySession> session_ = local_session_;
    int current_key_count_ = 0;
    int current_value_count_ = 0;
    int tab_height_ = 22;
    bool suppress_tab_change_ = false;
    RECT content_rect_ = {};
    ui::Splitter tree_splitter_{true, false};
    ui::Splitter history_splitter_{false, true};
    HICON address_go_icon_ = nullptr;
    HWND value_tooltip_ = nullptr;
    HWND value_tip_list_ = nullptr;
    int value_tip_item_ = -1;
    int value_tip_subitem_ = -1;
    std::wstring value_tooltip_text_;
    std::unique_ptr<util::PrivilegeScope> backup_privileges_;
    std::wstring status_account_sid_;
    std::wstring status_classes_sid_;
    std::wstring status_account_;
    work::KeyWatcher key_watcher_;
    bool show_value_ = true;
    work::DebouncedTask<workspace::TreeState> tree_state_saver_;
    ThemeMode theme_mode_ = ThemeMode::kSystem;
    std::vector<util::LanguagePack> language_packs_;
    std::wstring icon_dir_;
    bool updating_value_list_ = false;
    bool value_list_loading_ = false;
    bool merge_value_list_ = false;
    uint64_t merge_value_list_generation_ = 0;
    ULONGLONG auto_refresh_tick_ = 0;
    std::atomic<uint64_t> value_list_generation_{0};
    bool applying_theme_ = false;
    bool restart_on_close_ = false;
    bool reset_settings_on_close_ = false;
    std::optional<CacheKind> cache_clear_on_close_;
    bool history_loaded_ = false;
    bool history_cache_failed_ = false;
    std::wstring status_message_;
    bool is_replaying_ = false;
    bool hive_list_loaded_ = false;
    std::vector<ThemePreset> theme_presets_;
    LPARAM pending_value_list_kind_ = 0;
    std::wstring pending_value_list_name_;
    std::wstring appended_value_name_;
    HWND last_focus_ = nullptr;
    int pending_value_command_ = 0;
    std::wstring retained_value_name_;
    std::wstring pending_value_name_;
    std::vector<std::wstring> pending_value_selection_;
    std::wstring pending_value_selection_key_;
    int pending_value_top_index_ = 0;
    int retained_value_index_ = -1;
    std::wstring retained_value_key_path_;
    std::wstring queued_external_jump_target_;
    std::wstring pending_compat_jump_;
    bool flushing_external_navigation_ = false;
    bool jump_ui_batch_active_ = false;
    bool tree_redraw_pending_ = false;
    // set when the tree state cache is cleared for a restart, saved tabs then carry none
    bool drop_tree_state_ = false;
    bool status_update_pending_ = false;
    bool tree_painted_ = false;
    int pending_show_cmd_ = SW_SHOWNORMAL;
    std::wstring pending_compare_key_path_;
    std::wstring pending_compare_value_name_;
    std::wstring pending_external_value_key_path_;
    std::wstring pending_external_value_name_;
    std::unordered_map<std::wstring, std::wstring> hive_list_;
    std::shared_ptr<const std::unordered_set<std::wstring>> hive_roots_;
    bool deferred_startup_complete_ = false;
    work::Session startup_cache_session_;
    bool startup_tree_restore_pending_ = false;
    bool applying_startup_tree_restore_ = false;
    bool window_placement_loaded_ = false;
    ClipboardItem clipboard_;

    struct SearchRun
    {
        ~SearchRun()
        {
            session.Cancel();
            {
                std::lock_guard<std::mutex> lock(mutex);
            }
            space.notify_all();
            session.Join();
        }
        work::Session session;
        std::mutex mutex;
        std::condition_variable space;
        std::deque<std::vector<search::Result>> batches;
        size_t pending_rows = 0;
        bool producer_done = false;
        std::atomic_bool posted{false};
        std::atomic_bool progress_posted{false};
        std::atomic<uint64_t> searched{0};
        uint64_t start_tick = 0;
    };

    struct SearchTab
    {
        std::wstring label;
        std::vector<search::Result> results;
        std::vector<search::compare::Row> compare_rows;
        uint64_t next_row_id = 1;
        std::wstring compare_cache_file;
        bool load_pending = false;
        uint64_t load_generation = 0;
        std::wstring cache_file;
        bool results_loaded = true;
        uint64_t generation = 0;
        bool is_compare = false;
        search::compare::RowFilter compare_filter = search::compare::RowFilter::kDifferences;
        std::vector<search::Source> sources;
        size_t last_ui_count = 0;
        uint64_t max_results = 0;
        std::shared_ptr<SearchRun> run;
        uint64_t duration_ms = 0;
        int sort_column = -1;
        bool sort_ascending = true;
        util::SortKeys sort_earlier;
        uint64_t sort_revision = 0;
        bool sort_dirty = false;
        bool open_in_new_tab = false;
    };

    struct TabEntry
    {
        enum class Kind
        {
            kRegistry,
            kSearch,
            kRegFile,
        };

        Kind kind = Kind::kRegistry;
        int search_index = -1;
        RegistryMode registry_mode = RegistryMode::kLocal;
        REGSAM registry_view = 0;
        std::wstring offline_path;
        std::wstring remote_machine;
        std::wstring selected_path;
        std::wstring selected_value;
        std::vector<std::wstring> selected_values;
        int value_top_index = 0;
        std::vector<std::wstring> expanded_paths;
        std::shared_ptr<RegistrySession> session;
        // owned by browse_, created the first time the tab is shown
        RegistryTree* tree = nullptr;
        std::wstring reg_file_path;
        std::wstring reg_file_label;
        std::wstring reg_file_session_key;
        struct RegFileRoot
        {
            HKEY root = nullptr;
            std::wstring name;
            std::shared_ptr<VirtualRegistryData> data;
        };
        std::vector<RegFileRoot> reg_file_roots;
        bool reg_file_dirty = false;
        bool reg_file_loading = false;
    };

    struct TraceLoadPayload;
    struct DefaultLoadPayload;

    HWND search_results_list_ = nullptr;
    std::vector<TabEntry> tabs_;
    std::vector<SearchTab> search_tabs_;
    bool search_preview_request_posted_ = false;
    bool value_preview_request_posted_ = false;
    uint64_t search_last_refresh_tick_ = 0;
    uint64_t search_generation_ = 0;
    work::Session replace_session_;
    bool replace_result_pending_ = false;
    std::weak_ptr<RegistrySession> replace_target_;
    work::Session compare_session_;
    int active_search_tab_index_ = -1;
    int search_results_view_tab_index_ = -1;
    ui::TabStrip tab_strip_;
    bool value_activate_from_key_ = false;
    bool address_autocomplete_ = false;
    struct ActiveTrace
    {
        std::wstring label;
        std::wstring source_path;
        std::shared_ptr<const trace::Data> data;
        std::shared_ptr<const trace::Selection> selection;
    };
    struct ActiveDefault
    {
        std::wstring label;
        std::wstring source_path;
        std::shared_ptr<const defaults::Data> data;
        std::shared_ptr<const trace::Selection> selection;
    };
    static std::function<std::wstring(const std::wstring&, const std::wstring&)> DefaultDataLookup(std::vector<ActiveDefault> defaults);
    static void CollectTraceChildren(const std::vector<ActiveTrace>& traces, const std::wstring& key_lower, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out);

    struct TraceLoadPayload : work::MoveOnly
    {
        uint64_t generation = 0;
        std::vector<ActiveTrace> traces;
        std::unordered_map<std::wstring, trace::Selection> selection_cache;
    };

    struct DefaultLoadPayload : work::MoveOnly
    {
        uint64_t generation = 0;
        std::vector<ActiveDefault> defaults;
    };
    struct TraceParseSession
    {
        std::wstring label;
        std::wstring source_path;
        std::wstring source_lower;
        std::shared_ptr<trace::Data> data;
        trace::Selection selection;
        work::Session work;
        HWND dialog = nullptr;
        bool added_to_active = false;
        bool parsing_done = false;
    };
    struct DefaultParseSession
    {
        std::wstring label;
        std::wstring source_path;
        std::wstring source_lower;
        std::shared_ptr<defaults::Data> data;
        trace::Selection selection;
        work::Session work;
        HWND dialog = nullptr;
        bool added_to_active = false;
        bool parsing_done = false;
        bool show_errors = true;
    };
    struct RegFileParseSession
    {
        std::wstring source_path;
        std::wstring source_lower;
        work::Session work;
    };
    struct TraceDialogStartContext
    {
        MainWindow::Impl* window = nullptr;
        TraceParseSession* session = nullptr;
    };
    struct DefaultDialogStartContext
    {
        MainWindow::Impl* window = nullptr;
        DefaultParseSession* session = nullptr;
    };
    struct ValueListTask
    {
        uint64_t generation = 0;
        RegistryNode snapshot;
        std::wstring trace_path_lower;
        std::wstring default_path_lower;
        bool include_dates = false;
        int sort_column = 0;
        bool sort_ascending = true;
        util::SortKeys sort_earlier;
        bool show_keys_in_list = false;
        bool include_details = false;
        bool show_simulated_keys = false;
        bool include_all_value_data = false;
        HWND hwnd = nullptr;
        std::vector<ActiveTrace> trace_data_list;
        std::vector<ActiveDefault> default_data_list;
        std::shared_ptr<const std::unordered_set<std::wstring>> hive_roots;
    };
    struct ValuePreviewTask
    {
        uint64_t generation = 0;
        RegistryNode snapshot;
        std::vector<int> indices;
        std::vector<std::wstring> names;
        HWND hwnd = nullptr;
    };
    struct ValuePreviewItem
    {
        int index = 0;
        std::wstring name;
        DWORD type = 0;
        DWORD size = 0;
        std::wstring preview;
    };
    struct ValuePreviewPayload : work::MoveOnly
    {
        uint64_t generation = 0;
        std::vector<ValuePreviewItem> items;
    };
    work::LatestTask<ValuePreviewTask> value_preview_loader_;
    struct SearchPreviewRequest
    {
        int index = 0;
        uint64_t row_id = 0;
        RegistryNode node;
        std::wstring value_name;
    };
    struct SearchPreviewTask
    {
        uint64_t generation = 0;
        int tab_index = -1;
        std::vector<SearchPreviewRequest> requests;
        HWND hwnd = nullptr;
    };
    struct SearchPreviewItem
    {
        int index = 0;
        uint64_t row_id = 0;
        DWORD type = 0;
        DWORD data_size = 0;
        std::wstring preview;
    };
    struct SearchPreviewPayload : work::MoveOnly
    {
        uint64_t generation = 0;
        int tab_index = -1;
        std::vector<SearchPreviewItem> items;
    };
    work::LatestTask<SearchPreviewTask> search_preview_loader_;
    struct SearchSortTask
    {
        uint64_t generation = 0;
        int tab_index = -1;
        int column = 0;
        bool ascending = true;
        util::SortKeys earlier;
        uint64_t sort_revision = 0;
        std::vector<search::Source> sources;
        std::vector<search::Result> rows;
        std::vector<RegistryNode> nodes;
        HWND hwnd = nullptr;
    };
    struct SearchSortPayload : work::MoveOnly
    {
        uint64_t generation = 0;
        int tab_index = -1;
        uint64_t sort_revision = 0;
        std::vector<search::Result> rows;
    };
    work::LatestTask<SearchSortTask> search_sort_loader_;
    struct SearchTabLoadTask
    {
        uint64_t generation = 0;
        int tab_index = -1;
        std::wstring path;
        int sort_column = -1;
        bool sort_ascending = true;
        util::SortKeys sort_earlier;
        uint64_t sort_revision = 0;
        std::vector<search::Source> sources;
        HWND hwnd = nullptr;
    };
    struct SearchTabLoadPayload : work::MoveOnly
    {
        uint64_t generation = 0;
        int tab_index = -1;
        bool ok = false;
        uint64_t sort_revision = 0;
        std::vector<search::Result> rows;
    };
    work::LatestTask<SearchTabLoadTask> search_tab_loader_;
    uint64_t search_tab_load_generation_ = 0;
    std::vector<ActiveTrace> active_traces_;
    std::unordered_map<std::wstring, trace::Selection> trace_selection_cache_;
    workspace::RecentItems recent_trace_paths_{10};
    std::vector<ActiveDefault> active_defaults_;
    workspace::RecentItems recent_default_paths_{10};
    work::LatestTask<ValueListTask> value_loader_;
    frame::UpdateChecker updates_;
    work::Session trace_load_session_;
    std::unordered_map<std::wstring, std::unique_ptr<TraceParseSession>> trace_parse_sessions_;
    work::Session default_load_session_;
    std::unordered_map<std::wstring, std::unique_ptr<DefaultParseSession>> default_parse_sessions_;
    std::unordered_map<std::wstring, std::unique_ptr<RegFileParseSession>> reg_file_parse_sessions_;
    uint64_t reg_file_session_serial_ = 0;
    uint64_t last_trace_refresh_tick_ = 0;
    uint64_t last_default_refresh_tick_ = 0;
    changes::ValueComments value_comments_;
    changes::ValueComments default_comments_;
    bool comments_unreadable_ = false;
    bool comments_loaded_ = false;
    util::UniqueHKey registry_root_;
    std::vector<std::wstring> favorites_cache_;
    std::vector<workspace::NamedFavorite> regedit_favorites_;
    HMENU regedit_favorites_menu_ = nullptr;
    int regedit_favorites_static_count_ = 0;
    bool favorites_loaded_ = false;

    struct BundledDefault
    {
        std::wstring group;
        std::wstring label;
        std::wstring path;
    };

    bool bundled_defaults_loaded_ = false;
    struct MenuItemData
    {
        std::wstring text;
        bool separator = false;
        int width = 0;
        int height = 0;
    };
    std::vector<std::unique_ptr<MenuItemData>> menu_items_;
    std::vector<BundledDefault> bundled_defaults_;
};

} // namespace regkit
