// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/main_window.h"
#include "win32/windows_config.h"

#include <windows.h>
#include <commctrl.h>
#include <uxtheme.h>

#include <atomic>
#include <memory>
#include <condition_variable>
#include <deque>
#include <mutex>
#include <optional>
#include <shared_mutex>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "search/replace_dialog.h"
#include "search/query_dialog.h"
#include "appearance/theme.h"
#include "appearance/presets.h"
#include "frame/toolbar.h"
#include "trace/trace_dialog.h"
#include "registry/registry_store.h"
#include "search/compare.h"
#include "search/search.h"
#include "defaults/default_data.h"
#include "trace/trace_data.h"
#include "browse/browse_pane.h"
#include "registry/virtual_registry.h"
#include "changes/change_history.h"
#include "changes/key_snapshot.h"
#include "changes/undo_stack.h"
#include "changes/value_comments.h"
#include "workspace/recent_items.h"
#include "workspace/tree_state.h"
#include "work/session.h"
#include "win32/handle_owner.h"

struct IAutoComplete2;
struct IEnumString;

namespace regkit {

class MainWindow::Impl {
public:
  ~Impl();
  bool Create(HINSTANCE instance);
  void Show(int cmd_show);
  bool OpenRegFileTab(const std::wstring& path);
  bool TranslateAccelerator(const MSG& msg);
  void QueueExternalJump(const std::wstring& target);

private:
  friend class RegistryAddressEnum;
  enum class RegistryMode {
    kLocal,
    kRemote,
    kOffline,
  };
  enum class RegistryPathFormat {
    kFull,
    kAbbrev,
    kRegedit,
    kRegFile,
    kPowerShellDrive,
    kPowerShellProvider,
    kEscaped,
  };
  struct TabEntry;
  struct SearchTab;
  struct SearchTabLoadPayload;
  struct TraceParseSession;
  struct DefaultParseSession;
  struct UpdateCheckPayload : work::MoveOnly {
    bool silent = false;
    bool failed = false;
    std::wstring version;
    std::wstring download_url;
  };

  struct StartupCachePayload : work::MoveOnly {
    uint64_t generation = 0;
    std::vector<HistoryEntry> history_entries;
    std::vector<changes::CommentEntry> value_comments;
    std::vector<changes::CommentEntry> name_comments;
    std::wstring tree_selected_path;
    std::vector<std::wstring> tree_expanded_paths;
    bool history_loaded = false;
    bool comments_loaded = false;
    bool tree_state_loaded = false;
  };
  struct ReplacePayload : work::MoveOnly {
    struct Change {
      changes::UndoOperation undo;
      HistoryEntry history;
    };

    uint64_t generation = 0;
    std::vector<Change> changes;
    int failures = 0;
    bool cancelled = false;
  };

  static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam);
  static LRESULT CALLBACK AddressEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK FilterEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK TabProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK BorderProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK HeaderProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK ListViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK TreeViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);

  LRESULT HandleMessage(UINT message, WPARAM wparam, LPARAM lparam);
  std::optional<LRESULT> HandleLifecycleMessage(UINT message, WPARAM wparam,
                                                LPARAM lparam);
  std::optional<LRESULT> HandleLayoutInputMessage(UINT message, WPARAM wparam,
                                                  LPARAM lparam);
  std::optional<LRESULT> HandleWorkerMessage(UINT message, WPARAM wparam,
                                             LPARAM lparam);
  std::optional<LRESULT> HandleSearchWorkerMessage(UINT message, WPARAM wparam,
                                                   LPARAM lparam);
  std::optional<LRESULT> HandleLoadWorkerMessage(UINT message, WPARAM wparam,
                                                 LPARAM lparam);
  std::optional<LRESULT> HandleRegFileWorkerMessage(UINT message, WPARAM wparam,
                                                    LPARAM lparam);
  std::optional<LRESULT> HandleTraceWorkerMessage(UINT message, WPARAM wparam,
                                                  LPARAM lparam);
  std::optional<LRESULT> HandleDefaultWorkerMessage(UINT message, WPARAM wparam,
                                                    LPARAM lparam);
  std::optional<LRESULT> HandleValueWorkerMessage(UINT message, WPARAM wparam,
                                                  LPARAM lparam);
  std::optional<LRESULT> HandleExternalMessage(UINT message, WPARAM wparam,
                                               LPARAM lparam);
  std::optional<LRESULT> HandleAppearanceMessage(UINT message, WPARAM wparam,
                                                 LPARAM lparam);
  std::optional<LRESULT> HandleBrowseMessage(UINT message, WPARAM wparam,
                                             LPARAM lparam);
  LRESULT HandleNotification(LPARAM lparam);
  LRESULT HandleTooltipNotification(NMHDR* header, LPARAM lparam);
  bool ValueCellTooltipText(std::wstring* out);
  bool SearchCellTooltipText(std::wstring* out);
  bool ListCellTooltipText(std::wstring* out);
  std::wstring SearchCellFieldText(const search::Result& result, int subitem) const;
  LRESULT HandleToolbarNotification(NMHDR* header, LPARAM lparam);
  LRESULT HandleTabNotification(NMHDR* header, LPARAM lparam);
  LRESULT HandleTreeNotification(NMHDR* header, LPARAM lparam);
  LRESULT HandleHeaderNotification(NMHDR* header, LPARAM lparam);
  bool PaintHeaderItem(HWND header, NMCUSTOMDRAW* draw);
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
  void OnPaint();
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
  void InitDragLayout();
  void ApplyDragLayout();
  void BeginSplitterDrag();
  void BeginHistorySplitterDrag();
  void UpdateSplitterTrack(int client_x);
  void UpdateHistorySplitterTrack(int client_y);
  void EndSplitterDrag();
  void EndHistorySplitterDrag();
  void ComputeSplitterLimits(int* min_width, int* max_width) const;
  void ComputeHistorySplitterLimits(int* min_height, int* max_height) const;
  void BuildImageLists();
  void ReloadThemeIcons();
  bool ShouldUseLightIcons() const;
  std::wstring ResolveIconDir(bool use_light) const;
  std::wstring ResolveIconPath(const wchar_t* filename) const;
  HICON LoadThemeIcon(const wchar_t* filename, int light_id, int dark_id, int size, UINT dpi) const;
  void EnsureValueGridToolbar();
  void ApplyGridToolbarIcons();
  void LayoutGridToolbar(HWND list, HWND toolbar);
  void EnsureGridToolbar(HWND list, HWND* toolbar_slot, int command_id);
  void ApplyGridToolbarTheme(HWND toolbar);
  HBRUSH ValueHeaderSurfaceBrush() const;
  void LayoutValueGridToolbar();
  void SetValueGridEnabled(bool enabled, bool persist);
  void PaintValueGridLines(HWND list, HDC hdc, const RECT& area, int first_line_y,
                           int row_height);
  void PaintValueGridTail(HWND list, HDC hdc);
  int ValueGridToggleWidth(HWND header) const;
  ToolbarIcon MakeToolbarIcon(const wchar_t* filename, int light_id, int dark_id, bool use_light) const;
  void CreateValueColumns();
  void CreateHistoryColumns();
  void CreateSearchColumns();
  void ApplyValueColumns();
  void ApplyHistoryColumns();
  void ApplySearchColumns(bool compare);
  void UpdateValueListForNode(RegistryNode* node);
  void AttachBorder(HWND control);
  void AttachHeader(HWND header);
  void ResetValueFilter();
  void FocusFirstValue();
  void EnsureValueRowData(ListRow* row);
  void StartValuePreviewWorker();
  void QueueValuePreviews(int first, int last);
  void QueueSearchPreviews(int first, int last);
  void QueueSearchSort(SearchTab* tab);
  void StartSearchSortWorker();
  void StartSearchTabLoadWorker();
  void ApplySearchTabLoad(SearchTabLoadPayload* payload);
  void FinishSearchSession(uint64_t generation);
  void StartSearchPreviewWorker();
  void StartValueListWorker();
  void StopValueListWorker();
  void StartTraceLoadWorker();
  void StopTraceLoadWorker();
  void StartTraceParseThread(TraceParseSession* session);
  void MergeTraceEntries(
      TraceParseSession* session,
      const std::vector<KeyValueDialogEntry>& entries,
      std::unordered_set<std::wstring>* affected_keys);
  void StopTraceParseSessions();
  void StartDefaultLoadWorker();
  void StopDefaultLoadWorker();
  void StartDefaultParseThread(DefaultParseSession* session);
  void MergeDefaultEntries(
      DefaultParseSession* session,
      const std::vector<KeyValueDialogEntry>& entries,
      std::unordered_set<std::wstring>* affected_keys);
  void StopDefaultParseSessions();
  void StopRegFileParseSessions();
  static void StartTraceDialogLoad(HWND hwnd, void* context);
  static void StartDefaultDialogLoad(HWND hwnd, void* context);
  void UpdateAddressBar(RegistryNode* node);
  void UpdateGoButtonState();
  void EnableAddressAutoComplete();
  std::vector<std::wstring> BuildAddressSuggestions(const std::wstring& input) const;
  void ApplyAutoCompleteTheme();
  void UpdateStatus();
  void SortValueList(int column, bool toggle);
  void SortHistoryList(int column, bool toggle);
  void SortSearchResults(int column, bool toggle);
  void ClearHistoryItems(bool delete_cache);
  void CheckForUpdates(bool silent);
  void ApplyUpdateCheckResult(UpdateCheckPayload* payload);
  void RemoveSelectedHistoryItems();
  void RebuildHistoryList();
  void ScheduleValueListRename(LPARAM kind, const std::wstring& name);
  void StartPendingValueListRename();
  void StartSearch(const SearchDialogResult& options);
  void StartReplace(const ReplaceDialogResult& options);
  void ApplyReplacePayload(ReplacePayload* payload);
  void CommitReplacePayload(std::unique_ptr<ReplacePayload> payload,
                            bool show_failures);
  void StopReplace();
  void CancelSearch();
  bool IsSearchTabSelected() const;
  void UpdateSearchResultsView();
  void SortSearchTabResults(SearchTab* tab);
  void CloseSearchTab(int tab_index);
  bool SwitchToLocalRegistry();
  bool SwitchToRemoteRegistry();
  bool SwitchToOfflineRegistry();
  bool SaveOfflineRegistry();
  bool LoadOfflineRegistryFromPath(const std::wstring& path, bool open_new_tab);
  void ApplyRegistryRoots(const std::vector<RegistryRootEntry>& roots);
  std::vector<std::wstring> BuildVisibleTreePathParts(const std::wstring& path) const;
  std::wstring TreeRootLabel() const;
  void SelectDefaultTreeItem();
  void CaptureRegistryTabState(int index);
  void ResetRegistryTreeState();
  void RestoreRegistryTabState(int index);
  std::wstring LocalRegistryTabLabel(int index) const;
  void RefreshRegistryTabLabels();
  void ResetNavigationState();
  void UpdateTabText(const std::wstring& text);
  void UpdateTabWidth();
  void CloseTab(int tab_index);
  bool ConfirmCloseTab(int tab_index);
  void MarkOfflineDirty();
  void ClearOfflineDirty();
  void OpenLocalRegistryTab();
  int CurrentRegistryTabIndex() const;
  void UpdateRegistryTabEntry(RegistryMode mode, const std::wstring& offline_path, const std::wstring& remote_machine);
  bool IsSearchTabIndex(int index) const;
  bool IsRegFileTabIndex(int index) const;
  bool IsRegFileTabSelected() const;
  int SearchIndexFromTab(int index) const;
  int FindFirstRegistryTabIndex() const;
  void ActivateRegistryTab();
  bool SearchResultOpensInNewTab() const;
  void UpdateTabHotState(HWND hwnd, POINT pt);
  void PaintTabControl(HWND hwnd, HDC hdc);
  void DrawTabItem(HDC hdc, int index, const RECT& item_rect, int header_bottom, bool selected);
  bool GetTabCloseRect(int index, RECT* rect) const;
  void ReleaseRemoteRegistry();
  bool UnloadOfflineRegistry(std::wstring* error);
  void NavigateToAddress();
  bool SelectTreePath(const std::wstring& path);
  bool SelectValueByName(const std::wstring& name);
  void FocusAddressBarForExternalJump(bool defer_if_needed);
  void BeginJumpUiBatch();
  void EndJumpUiBatch();
  void ApplyTreeSelectionEffects(RegistryNode* node);
  void ApplyQueuedExternalJump();
  bool NavigateToResolvedExternalJump(const std::wstring& key_path, const std::wstring& value_name);
  bool NavigateToExternalJump(const std::wstring& target);
  bool ResolveExternalJumpTarget(const std::wstring& target, std::wstring* key_path, std::wstring* value_name) const;
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
  void RefreshBundledDefaultsCache();
  void BuildMenus();
  void BuildAccelerators();
  std::wstring CommandShortcutText(int command_id) const;
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
  bool HandleCreateCommand(int command_id);
  bool HandleModifyCommand(int command_id);
  bool HandleRenameCommand(int command_id);
  bool HandleDeleteCommand(int command_id);
  bool EnsureWritable();
  void PrepareMenusForOwnerDraw(HMENU menu, bool is_menu_bar);
  void OnMeasureMenuItem(MEASUREITEMSTRUCT* info);
  void OnDrawMenuItem(const DRAWITEMSTRUCT* info);
  void PaintMenuBarSeparator();
  void ShowValueHeaderMenu(POINT screen_pt);
  void ShowHistoryHeaderMenu(POINT screen_pt);
  void ShowSearchHeaderMenu(POINT screen_pt);
  void ToggleValueColumn(int column, bool visible);
  void ToggleHistoryColumn(int column, bool visible);
  void ToggleSearchColumn(int column, bool visible);
  void AppendHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data);
  void AppendHistoryEntry(HistoryEntry entry);
  void AppendValueHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data, const RegistryNode& node, const std::wstring& value_name, HistoryEntry::RevertKind revert_kind, const ValueEntry* revert_value = nullptr);
  bool PrepareHistoryRevert(const HistoryEntry& entry, HistoryEntry* prepared) const;
  bool OpenHistoryTarget(const HistoryEntry& entry);
  bool RevertHistoryEntry(const HistoryEntry& entry);
  void ShowAddressContextMenu(HWND edit, POINT screen_pt);
  void ShowTreeContextMenu(POINT screen_pt);
  void ShowValueContextMenu(POINT screen_pt);
  void ShowHistoryContextMenu(POINT screen_pt);
  void ShowSearchResultContextMenu(POINT screen_pt);
  void DrawAddressButton(const DRAWITEMSTRUCT* info);
  void DrawHeaderCloseButton(const DRAWITEMSTRUCT* info);
  void ShowPermissionsDialog(const RegistryNode& node);
  bool OpenDefaultRegedit();
  void ReplaceRegedit(bool enable);
  void SyncReplaceRegeditState();
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
  void SyncRegFileTabSelection();
  void ResetHiveListCache();
  void EnsureHiveListLoaded();
  std::wstring LookupHivePath(const RegistryNode& node, bool* is_root);
  int KeyIconIndex(const RegistryNode& node, bool* is_link, bool* is_hive_root);
  void AppendRealRegistryRoot(std::vector<RegistryRootEntry>* roots);
  void HandleTypeToSelectTree(wchar_t ch);
  void HandleTypeToSelectList(wchar_t ch);
  std::wstring NormalizeRegistryPath(const std::wstring& path) const;
  std::wstring FormatRegistryPath(const std::wstring& path, RegistryPathFormat format) const;
  bool FindNearestExistingPath(const std::wstring& path, std::wstring* nearest_path) const;
  bool CreateRegistryPath(const std::wstring& path);
  bool SelectAllInFocusedList();
  bool InvertSelectionInFocusedList();
  bool IsCompareTabSelected() const;
  void StartCompareRegistries();
  void AppendHistoryCache(const HistoryEntry& entry);
  std::wstring CacheFolderPath() const;
  std::wstring HistoryCachePath() const;
  std::wstring TabsCachePath() const;
  std::wstring SearchTabCachePath(const std::wstring& file) const;
  void LoadTabs();
  bool SaveTabs();
  void ClearTabsCache();
  bool EnsureSearchTabResultsLoaded(int search_index);
  void StartStartupCacheLoad(bool include_tree_state);
  void StopStartupCacheLoad();
  void ApplyStartupCachePayload(StartupCachePayload* payload);
  void SaveComments() const;
  bool ImportCommentsFromFile(const std::wstring& path);
  bool ExportCommentsToFile(const std::wstring& path) const;
  void RefreshValueListComments();
  std::wstring CommentsPath() const;
  bool EditValueComments(const std::vector<ListRow>& rows);
  bool RestartAsAdmin();
  bool RestartAsSystem();
  bool RestartAsTrustedInstaller();
  void LoadSettings();
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
  struct DefaultValueChoice {
    std::wstring label;
    bool present = false;
    DWORD type = REG_NONE;
    std::vector<BYTE> data;
  };
  std::vector<DefaultValueChoice> CollectDefaultChoices(
      const std::wstring& value_name) const;
  bool HandleResetDefaultCommand(int command_id);
  std::vector<DefaultValueChoice> SelectedValueDefaultChoices() const;
  HMENU BuildResetDefaultMenu(
      const std::vector<DefaultValueChoice>& choices) const;
  void AppendResetDefaultMenu(HMENU menu);
  void RefreshResetDefaultMenu(HMENU menu);
  std::wstring TreeStatePath() const;
  void LoadTreeState();
  void StartTreeStateWorker();
  void StopTreeStateWorker();
  void MarkTreeStateDirty();
  void SaveTreeStateFile(const std::wstring& selected, const std::vector<std::wstring>& expanded) const;
  void CaptureTreeState(std::wstring* selected_path, std::vector<std::wstring>* expanded_paths) const;
  void RestoreTreeState();
  bool ExpandTreePath(const std::wstring& path);
  HTREEITEM FindTreeItem(const std::wstring& path);
  void RefreshTreeItem(HTREEITEM item);
  void RefreshTreePath(const std::wstring& path);
  void RefreshTreeSelection();
  void RefreshMatchingTreeNodes();
  void UpdateSimulatedChain(HTREEITEM item);
  void ApplySavedWindowPlacement();
  LOGFONTW DefaultLogFont() const;
  void AddRecentTracePath(const std::wstring& path);
  void NormalizeRecentTraceList();
  void AddRecentDefaultPath(const std::wstring& path);
  void NormalizeRecentDefaultList();
  void AppendTraceChildren(const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) const;
  std::wstring TracePathLowerForNode(const RegistryNode& node) const;
  bool AllowTraceSimulation(const RegistryNode& node) const;

  struct ClipboardItem {
    enum class Kind {
      kNone,
      kValue,
      kKey,
    };

    Kind kind = Kind::kNone;
    RegistryNode source_parent;
    std::wstring name;
    ValueEntry value;
    changes::KeySnapshot key_snapshot;
  };

  void PushUndo(changes::UndoOperation operation);
  void ClearRedo();
  bool ApplyUndoOperation(const changes::UndoOperation& operation, bool redo);
  bool SameNode(const RegistryNode& left, const RegistryNode& right) const;
  std::wstring MakeUniqueValueName(const RegistryNode& node, const std::wstring& base) const;
  std::wstring MakeUniqueKeyName(const RegistryNode& node, const std::wstring& base) const;
  bool ResolvePathToNode(const std::wstring& path, RegistryNode* node) const;

  HINSTANCE instance_ = nullptr;
  HWND hwnd_ = nullptr;
  HFONT ui_font_ = nullptr;
  HFONT icon_font_ = nullptr;
  bool ui_font_owned_ = false;
  bool use_custom_font_ = false;
  LOGFONTW custom_font_ = {};
  HACCEL accelerators_ = nullptr;
  Toolbar toolbar_;
  HWND tab_ = nullptr;
  HWND tree_header_ = nullptr;
  HWND tree_close_btn_ = nullptr;
  HWND history_label_ = nullptr;
  HWND history_close_btn_ = nullptr;
  HWND history_list_ = nullptr;
  HWND status_bar_ = nullptr;
  HWND search_progress_ = nullptr;
  browse::Pane browse_;
  HIMAGELIST tree_images_ = nullptr;
  HIMAGELIST list_images_ = nullptr;
  std::vector<int> value_column_subitems_;
  std::vector<int> search_column_subitems_;
  std::vector<ColumnInfo> history_columns_;
  std::vector<int> history_column_widths_;
  std::vector<bool> history_column_visible_;
  std::vector<ColumnInfo> search_columns_;
  std::vector<int> search_column_widths_;
  std::vector<bool> search_column_visible_;
  std::vector<ColumnInfo> compare_columns_;
  std::vector<int> compare_column_widths_;
  std::vector<bool> compare_column_visible_;
  bool compare_columns_active_ = false;
  int last_header_column_ = -1;
  int history_sort_column_ = 0;
  bool history_sort_ascending_ = true;
  int history_max_rows_ = 500;
  changes::ChangeHistory change_history_;
  RegistryMode registry_mode_ = RegistryMode::kLocal;
  std::wstring remote_machine_;
  HKEY remote_hklm_ = nullptr;
  HKEY remote_hku_ = nullptr;
  HKEY offline_root_ = nullptr;
  std::vector<HKEY> offline_roots_;
  std::wstring offline_mount_;
  std::vector<std::wstring> offline_root_labels_;
  std::vector<std::wstring> offline_root_paths_;
  std::wstring offline_root_name_;
  int current_key_count_ = 0;
  int current_value_count_ = 0;
  int tab_height_ = 22;
  bool suppress_tab_change_ = false;
  int tree_width_ = 260;
  int history_height_ = 160;
  RECT splitter_rect_ = {};
  bool splitter_dragging_ = false;
  int splitter_start_x_ = 0;
  int splitter_start_width_ = 0;
  int splitter_min_width_ = 0;
  int splitter_max_width_ = 0;
  RECT history_splitter_rect_ = {};
  bool history_splitter_dragging_ = false;
  int history_splitter_start_y_ = 0;
  int history_splitter_start_height_ = 0;
  int history_splitter_min_height_ = 0;
  int history_splitter_max_height_ = 0;
  bool drag_layout_valid_ = false;
  int drag_client_width_ = 0;
  int drag_client_height_ = 0;
  int drag_content_top_ = 0;
  int drag_content_left_ = 0;
  int drag_content_right_ = 0;
  int drag_status_top_ = 0;
  int drag_tree_header_height_ = 0;
  int drag_history_label_height_ = 0;
  HICON address_go_icon_ = nullptr;
  HWND value_grid_toolbar_ = nullptr;
  HWND search_grid_toolbar_ = nullptr;
  HWND value_tooltip_ = nullptr;
  HWND value_tip_list_ = nullptr;
  int value_tip_item_ = -1;
  int value_tip_subitem_ = -1;
  std::wstring value_tooltip_text_;
  HIMAGELIST value_grid_image_list_ = nullptr;
  bool show_value_grid_ = false;
  COLORREF grid_line_color_ = CLR_INVALID;
  bool auto_check_updates_ = false;
  bool default_reset_enabled_ = false;
  HMENU reset_default_menu_ = nullptr;
  bool update_check_running_ = false;
  bool show_toolbar_ = true;
  bool show_address_bar_ = true;
  bool show_filter_bar_ = true;
  bool show_tab_control_ = true;
  bool show_tree_ = true;
  bool show_history_ = true;
  bool show_value_ = true;
  bool show_status_bar_ = true;
  bool show_keys_in_list_ = true;
  bool show_extra_hives_ = false;
  bool show_simulated_keys_ = true;
  bool save_tree_state_ = true;
  work::DebouncedTask<workspace::TreeState> tree_state_saver_;
  bool always_on_top_ = false;
  bool always_run_as_admin_ = false;
  bool always_run_as_system_ = false;
  bool always_run_as_trustedinstaller_ = false;
  bool replace_regedit_ = false;
  bool single_instance_ = true;
  bool read_only_ = false;
  ThemeMode theme_mode_ = ThemeMode::kSystem;
  std::wstring icon_set_ = L"phosphor-regedit";
  std::wstring icon_dir_;
  bool updating_value_list_ = false;
  bool value_list_loading_ = false;
  std::atomic<uint64_t> value_list_generation_{0};
  bool applying_theme_ = false;
  bool history_loaded_ = false;
  bool is_replaying_ = false;
  bool clear_history_on_exit_ = false;
  bool save_tabs_ = true;
  bool clear_tabs_on_exit_ = false;
  bool hive_list_loaded_ = false;
  std::vector<ThemePreset> theme_presets_;
  std::wstring active_theme_preset_;
  LPARAM pending_value_list_kind_ = 0;
  std::wstring pending_value_list_name_;
  std::wstring appended_value_name_;
  HWND last_focus_ = nullptr;
  int pending_value_command_ = 0;
  std::wstring retained_value_name_;
  std::wstring retained_value_key_path_;
  std::wstring queued_external_jump_target_;
  bool jump_ui_batch_active_ = false;
  std::wstring pending_external_value_key_path_;
  std::wstring pending_external_value_name_;
  std::unordered_map<std::wstring, std::wstring> hive_list_;
  std::shared_ptr<const std::unordered_set<std::wstring>> hive_roots_;
  workspace::TreeState saved_tree_state_;
  bool tree_state_restored_ = false;
  bool deferred_startup_complete_ = false;
  work::Session startup_cache_session_;
  bool startup_tree_restore_pending_ = false;
  bool applying_startup_tree_restore_ = false;
  bool window_placement_loaded_ = false;
  int window_x_ = 0;
  int window_y_ = 0;
  int window_width_ = 0;
  int window_height_ = 0;
  bool window_maximized_ = false;
  ClipboardItem clipboard_;
  changes::UndoStack undo_stack_;
  ReplaceDialogResult last_replace_;
  SearchDialogResult last_search_;

  struct SearchTab {
    std::wstring label;
    std::vector<search::Result> results;
    std::vector<search::compare::Row> compare_rows;
    uint64_t next_row_id = 1;
    bool load_pending = false;
    uint64_t load_generation = 0;
    std::wstring cache_file;
    bool results_loaded = true;
    uint64_t generation = 0;
    bool is_compare = false;
    size_t last_ui_count = 0;
    int sort_column = -1;
    bool sort_ascending = true;
    bool sort_dirty = false;
    bool open_in_new_tab = false;
  };

  struct TabEntry {
    enum class Kind {
      kRegistry,
      kSearch,
      kRegFile,
    };

    Kind kind = Kind::kRegistry;
    int search_index = -1;
    RegistryMode registry_mode = RegistryMode::kLocal;
    std::wstring offline_path;
    std::wstring remote_machine;
    std::wstring selected_path;
    std::vector<std::wstring> expanded_paths;
    bool offline_dirty = false;
    std::wstring reg_file_path;
    std::wstring reg_file_label;
    struct RegFileRoot {
      HKEY root = nullptr;
      std::wstring name;
      std::shared_ptr<VirtualRegistryData> data;
    };
    std::vector<RegFileRoot> reg_file_roots;
    bool reg_file_dirty = false;
    bool reg_file_loading = false;
  };

  struct PendingSearchBatch {
    uint64_t generation = 0;
    std::vector<search::Result> rows;
  };

  struct TraceLoadPayload;
  struct DefaultLoadPayload;

  HWND search_results_list_ = nullptr;
  std::vector<TabEntry> tabs_;
  std::vector<SearchTab> search_tabs_;
  std::deque<PendingSearchBatch> search_pending_batches_;
  size_t search_pending_rows_ = 0;
  bool search_producer_done_ = false;
  std::mutex search_mutex_;
  std::condition_variable search_queue_space_;
  std::atomic_bool search_posted_{false};
  bool search_preview_request_posted_ = false;
  bool value_preview_request_posted_ = false;
  std::atomic<uint64_t> search_progress_searched_{0};
  std::atomic_bool search_progress_posted_{false};
  uint64_t search_last_refresh_tick_ = 0;
  uint64_t search_start_tick_ = 0;
  uint64_t search_duration_ms_ = 0;
  bool search_duration_valid_ = false;
  work::Session search_session_;
  work::Session replace_session_;
  bool replace_result_pending_ = false;
  bool search_running_ = false;
  int active_search_tab_index_ = -1;
  int search_results_view_tab_index_ = -1;
  int tab_hot_index_ = -1;
  int tab_close_hot_index_ = -1;
  int tab_close_down_index_ = -1;
  bool tab_mouse_tracking_ = false;
  bool value_activate_from_key_ = false;
  ::IAutoComplete2* address_autocomplete_ = nullptr;
  ::IEnumString* address_autocomplete_source_ = nullptr;
  struct ActiveTrace {
    std::wstring label;
    std::wstring source_path;
    std::shared_ptr<const trace::Data> data;
    std::shared_ptr<const trace::Selection> selection;
  };
  struct ActiveDefault {
    std::wstring label;
    std::wstring source_path;
    std::shared_ptr<const defaults::Data> data;
    std::shared_ptr<const trace::Selection> selection;
  };

  struct TraceLoadPayload : work::MoveOnly {
    uint64_t generation = 0;
    std::vector<ActiveTrace> traces;
    std::unordered_map<std::wstring, trace::Selection> selection_cache;
  };

  struct DefaultLoadPayload : work::MoveOnly {
    uint64_t generation = 0;
    std::vector<ActiveDefault> defaults;
  };
  struct TraceParseSession {
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
  struct DefaultParseSession {
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
  struct RegFileParseSession {
    std::wstring source_path;
    std::wstring source_lower;
    work::Session work;
  };
  struct TraceDialogStartContext {
    MainWindow::Impl* window = nullptr;
    TraceParseSession* session = nullptr;
  };
  struct DefaultDialogStartContext {
    MainWindow::Impl* window = nullptr;
    DefaultParseSession* session = nullptr;
  };
  struct ValueListTask {
    uint64_t generation = 0;
    RegistryNode snapshot;
    std::wstring trace_path_lower;
    std::wstring default_path_lower;
    bool include_dates = false;
    int sort_column = 0;
    bool sort_ascending = true;
    bool show_keys_in_list = false;
    bool include_details = false;
    bool show_simulated_keys = false;
    bool include_all_value_data = false;
    HWND hwnd = nullptr;
    std::vector<ActiveTrace> trace_data_list;
    std::vector<ActiveDefault> default_data_list;
    std::shared_ptr<const std::unordered_set<std::wstring>> hive_roots;
  };
  struct ValuePreviewTask {
    uint64_t generation = 0;
    RegistryNode snapshot;
    std::vector<int> indices;
    std::vector<std::wstring> names;
    HWND hwnd = nullptr;
  };
  struct ValuePreviewItem {
    int index = 0;
    std::wstring name;
    DWORD type = 0;
    DWORD size = 0;
    std::wstring preview;
  };
  struct ValuePreviewPayload {
    uint64_t generation = 0;
    std::vector<ValuePreviewItem> items;
  };
  work::LatestTask<ValuePreviewTask> value_preview_loader_;
  struct SearchPreviewRequest {
    int index = 0;
    uint64_t row_id = 0;
    RegistryNode node;
    std::wstring value_name;
  };
  struct SearchPreviewTask {
    uint64_t generation = 0;
    int tab_index = -1;
    std::vector<SearchPreviewRequest> requests;
    HWND hwnd = nullptr;
  };
  struct SearchPreviewItem {
    int index = 0;
    uint64_t row_id = 0;
    DWORD type = 0;
    DWORD data_size = 0;
    std::wstring preview;
  };
  struct SearchPreviewPayload {
    uint64_t generation = 0;
    int tab_index = -1;
    std::vector<SearchPreviewItem> items;
  };
  work::LatestTask<SearchPreviewTask> search_preview_loader_;
  struct SearchSortTask {
    uint64_t generation = 0;
    int tab_index = -1;
    int column = 0;
    bool ascending = true;
    std::vector<search::Result> rows;
    std::vector<RegistryNode> nodes;
    HWND hwnd = nullptr;
  };
  struct SearchSortPayload {
    uint64_t generation = 0;
    int tab_index = -1;
    std::vector<search::Result> rows;
  };
  work::LatestTask<SearchSortTask> search_sort_loader_;
  struct SearchTabLoadTask {
    uint64_t generation = 0;
    int tab_index = -1;
    std::wstring path;
    int sort_column = -1;
    bool sort_ascending = true;
    HWND hwnd = nullptr;
  };
  struct SearchTabLoadPayload {
    uint64_t generation = 0;
    int tab_index = -1;
    bool ok = false;
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
  work::Session trace_load_session_;
  std::unordered_map<std::wstring, std::unique_ptr<TraceParseSession>> trace_parse_sessions_;
  work::Session default_load_session_;
  std::unordered_map<std::wstring, std::unique_ptr<DefaultParseSession>> default_parse_sessions_;
  std::unordered_map<std::wstring, std::unique_ptr<RegFileParseSession>> reg_file_parse_sessions_;
  uint64_t last_trace_refresh_tick_ = 0;
  uint64_t last_default_refresh_tick_ = 0;
  changes::ValueComments value_comments_;
  util::UniqueHKey registry_root_;
  std::vector<std::wstring> favorites_cache_;
  bool favorites_loaded_ = false;

  struct BundledDefault {
    std::wstring group;
    std::wstring label;
    std::wstring path;
  };

  bool bundled_defaults_loaded_ = false;
  struct MenuItemData {
    std::wstring text;
    std::wstring left_text;
    std::wstring right_text;
    bool separator = false;
    bool has_submenu = false;
    bool is_menu_bar = false;
    int width = 0;
    int height = 0;
  };
  std::vector<std::unique_ptr<MenuItemData>> menu_items_;
  std::vector<BundledDefault> bundled_defaults_;
};

} // namespace regkit
