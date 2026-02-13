// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#pragma once

#include <windows.h>
#include <commctrl.h>

#include <atomic>
#include <condition_variable>
#include <memory>
#include <mutex>
#include <shared_mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "app/registry_tree.h"
#include "app/replace_dialog.h"
#include "app/search_dialog.h"
#include "app/theme.h"
#include "app/theme_presets.h"
#include "app/toolbar.h"
#include "app/trace_dialog.h"
#include "app/value_list.h"
#include "registry/registry_provider.h"
#include "registry/search_engine.h"
#include "win32/win32_helpers.h"

struct IAutoComplete2;
struct IEnumString;

namespace regkit {

struct HistoryEntry {
  uint64_t timestamp = 0;
  std::wstring time_text;
  std::wstring action;
  std::wstring old_data;
  std::wstring new_data;
};

class MainWindow {
public:
  ~MainWindow();
  bool Create(HINSTANCE instance);
  void Show(int cmd_show);
  bool OpenRegFileTab(const std::wstring& path);
  bool TranslateAccelerator(const MSG& msg);
  void UpdateThemePresets(const std::vector<ThemePreset>& presets, const std::wstring& active_name, bool apply_now);

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
  struct TraceData;
  struct DefaultData;
  struct TabEntry;
  struct TraceParseSession;
  struct DefaultParseSession;

  static LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam);
  static LRESULT CALLBACK AddressEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK FilterEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK TabProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK HeaderProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK ListViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);
  static LRESULT CALLBACK TreeViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR subclass_id, DWORD_PTR ref_data);

  LRESULT HandleMessage(UINT message, WPARAM wparam, LPARAM lparam);
  bool OnCreate();
  void OnDestroy();
  void OnSize(int width, int height);
  void OnPaint();
  void ApplyThemeToChildren();
  void ApplySystemTheme();
  void LoadThemePresets();
  void SaveThemePresets() const;
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
  void EndSplitterDrag(bool apply);
  void EndHistorySplitterDrag(bool apply);
  void ComputeSplitterLimits(int* min_width, int* max_width) const;
  void ComputeHistorySplitterLimits(int* min_height, int* max_height) const;
  void BuildImageLists();
  void ReloadThemeIcons();
  bool ShouldUseLightIcons() const;
  std::wstring ResolveIconPath(const wchar_t* filename, bool use_light) const;
  HICON LoadThemeIcon(const wchar_t* filename, int light_id, int dark_id, int size, UINT dpi) const;
  ToolbarIcon MakeToolbarIcon(const wchar_t* filename, int light_id, int dark_id, bool use_light) const;
  void CreateValueColumns();
  void CreateHistoryColumns();
  void CreateSearchColumns();
  void ApplyValueColumns();
  void ApplyHistoryColumns();
  void ApplySearchColumns(bool compare);
  void UpdateValueListForNode(RegistryNode* node);
  void EnsureValueRowData(ListRow* row);
  void StartValueListWorker();
  void StopValueListWorker();
  void StartTraceLoadWorker();
  void StopTraceLoadWorker();
  void StartTraceParseThread(TraceParseSession* session);
  void StopTraceParseSessions();
  void StartDefaultLoadWorker();
  void StopDefaultLoadWorker();
  void StartDefaultParseThread(DefaultParseSession* session);
  void StopDefaultParseSessions();
  void StopRegFileParseSessions();
  static void StartTraceDialogLoad(HWND hwnd, void* context);
  static void StartDefaultDialogLoad(HWND hwnd, void* context);
  void UpdateAddressBar(RegistryNode* node);
  void EnableAddressAutoComplete();
  std::vector<std::wstring> BuildAddressSuggestions(const std::wstring& input) const;
  void ApplyAutoCompleteTheme();
  void UpdateStatus();
  void SortValueList(int column, bool toggle);
  void SortHistoryList(int column, bool toggle);
  void SortSearchResults(int column, bool toggle);
  void ClearHistoryItems(bool delete_cache);
  void RebuildHistoryList();
  void ScheduleValueListRename(LPARAM kind, const std::wstring& name);
  void StartPendingValueListRename();
  void StartSearch(const SearchDialogResult& options);
  void StartReplace(const ReplaceDialogResult& options);
  void CancelSearch();
  bool IsSearchTabSelected() const;
  void UpdateSearchResultsView();
  void CloseSearchTab(int tab_index);
  bool SwitchToLocalRegistry();
  bool SwitchToRemoteRegistry();
  bool SwitchToOfflineRegistry();
  bool SaveOfflineRegistry();
  bool LoadOfflineRegistryFromPath(const std::wstring& path, bool open_new_tab);
  void ApplyRegistryRoots(const std::vector<RegistryRootEntry>& roots);
  std::wstring TreeRootLabel() const;
  void SelectDefaultTreeItem();
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
  int FindFirstSearchTabIndex() const;
  int FindFirstRegistryTabIndex() const;
  void UpdateTabHotState(HWND hwnd, POINT pt);
  void PaintTabControl(HWND hwnd, HDC hdc);
  void DrawTabItem(HDC hdc, int index, const RECT& item_rect, int header_bottom, bool selected);
  bool GetTabCloseRect(int index, RECT* rect) const;
  void ReleaseRemoteRegistry();
  bool UnloadOfflineRegistry(std::wstring* error);
  void NavigateToAddress();
  bool SelectTreePath(const std::wstring& path);
  bool SelectValueByName(const std::wstring& name);
  bool LoadTraceFromFile(const std::wstring& label, const std::wstring& path, const TraceSelection* selection_override = nullptr);
  bool LoadBundledTrace(const std::wstring& label, const TraceSelection* selection_override = nullptr);
  std::wstring ResolveBundledTracePath(const std::wstring& label) const;
  bool LoadTraceFromBuffer(const std::wstring& label, const std::wstring& source, const std::string& buffer, const TraceSelection* selection_override = nullptr);
  bool LoadTraceFromPrompt();
  void ClearTrace();
  bool LoadDefaultFromFile(const std::wstring& label, const std::wstring& path);
  bool LoadBundledDefault(const std::wstring& label);
  std::wstring ResolveBundledDefaultPath(const std::wstring& label) const;
  bool LoadDefaultFromPrompt();
  void ClearDefaults();
  bool ParseDefaultRegFile(const std::wstring& path, DefaultData* out, std::wstring* error) const;
  void BuildMenus();
  void BuildAccelerators();
  std::wstring CommandShortcutText(int command_id) const;
  std::wstring CommandTooltipText(int command_id) const;
  bool HandleMenuCommand(int command_id);
  bool EnsureWritable();
  void PrepareMenusForOwnerDraw(HMENU menu, bool is_menu_bar);
  void OnMeasureMenuItem(MEASUREITEMSTRUCT* info);
  void OnDrawMenuItem(const DRAWITEMSTRUCT* info);
  void ShowValueHeaderMenu(POINT screen_pt);
  void ShowHistoryHeaderMenu(POINT screen_pt);
  void ShowSearchHeaderMenu(POINT screen_pt);
  void ToggleValueColumn(int column, bool visible);
  void ToggleHistoryColumn(int column, bool visible);
  void ToggleSearchColumn(int column, bool visible);
  void AppendHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data);
  std::wstring ResolveSearchComment(const SearchResult& result) const;
  void ShowTreeContextMenu(POINT screen_pt);
  void ShowValueContextMenu(POINT screen_pt);
  void ShowHistoryContextMenu(POINT screen_pt);
  void ShowSearchResultContextMenu(POINT screen_pt);
  void DrawAddressButton(const DRAWITEMSTRUCT* info);
  void DrawHeaderCloseButton(const DRAWITEMSTRUCT* info);
  void ShowPermissionsDialog(const RegistryNode& node);
  void ReplaceRegedit(bool enable);
  void SyncReplaceRegeditState();
  bool OpenDefaultRegedit();
  void OpenHiveFileDir();
  void AddAddressHistory(const std::wstring& path);
  void RecordNavigation(const std::wstring& path);
  void NavigateBack();
  void NavigateForward();
  void NavigateUp();
  void UpdateNavigationButtons();
  void ApplyViewVisibility();
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
  void LoadHistoryCache();
  void AppendHistoryCache(const HistoryEntry& entry);
  std::wstring CacheFolderPath() const;
  std::wstring HistoryCachePath() const;
  std::wstring TabsCachePath() const;
  std::wstring SearchTabCachePath(const std::wstring& file) const;
  void LoadTabs();
  void SaveTabs();
  void ClearTabsCache();
  bool ReadSearchResults(const std::wstring& path, std::vector<SearchResult>* results) const;
  bool WriteSearchResults(const std::wstring& path, const std::vector<SearchResult>& results) const;
  void LoadComments();
  void SaveComments() const;
  bool ImportCommentsFromFile(const std::wstring& path);
  bool ExportCommentsToFile(const std::wstring& path) const;
  void RefreshValueListComments();
  std::wstring CommentsPath() const;
  bool EditValueComment(const ListRow& row);
  bool IsProcessElevated() const;
  bool IsProcessSystem() const;
  bool IsProcessTrustedInstaller() const;
  bool RestartAsAdmin();
  bool RestartAsSystem();
  bool RestartAsTrustedInstaller();
  void LoadSettings();
  void SaveSettings() const;
  std::wstring SettingsPath() const;
  std::wstring ActiveTracesPath() const;
  void LoadActiveTraces();
  void SaveActiveTraces() const;
  std::wstring ActiveDefaultsPath() const;
  void LoadActiveDefaults();
  void SaveActiveDefaults() const;
  std::wstring TraceSettingsPath() const;
  void LoadTraceSettings();
  void SaveTraceSettings() const;
  bool AddTraceFromFile(const std::wstring& label, const std::wstring& path, const TraceSelection* selection_override, bool prompt_for_selection = true, bool update_ui = true);
  bool AddTraceFromBuffer(const std::wstring& label, const std::wstring& source, const std::string& buffer, const TraceSelection* selection_override, bool prompt_for_selection);
  bool BuildTraceDataFromBuffer(const std::wstring& label, const std::wstring& source, const std::string& buffer, TraceData* out_data, std::wstring* error) const;
  bool RemoveTraceByPath(const std::wstring& path);
  bool RemoveTraceByLabel(const std::wstring& label);
  bool HasActiveTraces() const;
  bool AddDefaultFromFile(const std::wstring& label, const std::wstring& path, bool show_error = true, bool prompt_for_selection = false, bool update_ui = true);
  bool SaveRegFileTab(int tab_index);
  bool ExportRegFileTab(int tab_index, const std::wstring& path);
  bool BuildRegFileContent(const TabEntry& entry, std::wstring* out) const;
  void ReleaseRegFileRoots(TabEntry* entry);
  bool RemoveDefaultByPath(const std::wstring& path);
  bool RemoveDefaultByLabel(const std::wstring& label);
  bool HasActiveDefaults() const;
  std::wstring TreeStatePath() const;
  void LoadTreeState();
  void StartTreeStateWorker();
  void StopTreeStateWorker();
  void MarkTreeStateDirty();
  void SaveTreeStateFile(const std::wstring& selected, const std::vector<std::wstring>& expanded) const;
  void CaptureTreeState(std::wstring* selected_path, std::vector<std::wstring>* expanded_paths) const;
  void RestoreTreeState();
  bool ExpandTreePath(const std::wstring& path);
  void RefreshTreeSelection();
  void UpdateSimulatedChain(HTREEITEM item);
  void ApplySavedWindowPlacement();
  LOGFONTW DefaultLogFont() const;
  void AddRecentTracePath(const std::wstring& path);
  void NormalizeRecentTraceList();
  void AddRecentDefaultPath(const std::wstring& path);
  void NormalizeRecentDefaultList();
  void AppendTraceChildren(const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) const;
  std::wstring TracePathLowerForNode(const RegistryNode& node) const;
  void NormalizeSelectionForTrace(const TraceData& trace, TraceSelection* selection) const;
  bool AllowTraceSimulation(const RegistryNode& node) const;

  struct KeySnapshot {
    std::wstring name;
    std::vector<ValueEntry> values;
    std::vector<KeySnapshot> children;
  };

  struct UndoOperation {
    enum class Type {
      kCreateKey,
      kDeleteKey,
      kRenameKey,
      kCreateValue,
      kDeleteValue,
      kModifyValue,
      kRenameValue,
    };

    Type type = Type::kCreateKey;
    RegistryNode node;
    std::wstring name;
    std::wstring new_name;
    ValueEntry old_value;
    ValueEntry new_value;
    KeySnapshot key_snapshot;
  };

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
    KeySnapshot key_snapshot;
  };

  void PushUndo(UndoOperation operation);
  void ClearRedo();
  bool ApplyUndoOperation(const UndoOperation& operation, bool redo);
  KeySnapshot CaptureKeySnapshot(const RegistryNode& node);
  bool RestoreKeySnapshot(const RegistryNode& parent, const KeySnapshot& snapshot);
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
  HWND address_edit_ = nullptr;
  HWND address_go_btn_ = nullptr;
  HWND filter_edit_ = nullptr;
  HWND tab_ = nullptr;
  HWND tree_header_ = nullptr;
  HWND tree_close_btn_ = nullptr;
  HWND history_label_ = nullptr;
  HWND history_close_btn_ = nullptr;
  HWND history_list_ = nullptr;
  HWND status_bar_ = nullptr;
  HWND search_progress_ = nullptr;
  RegistryTree tree_;
  ValueList value_list_;
  HIMAGELIST tree_images_ = nullptr;
  HIMAGELIST list_images_ = nullptr;
  std::vector<ColumnInfo> value_columns_;
  std::vector<int> value_column_widths_;
  std::vector<bool> value_column_visible_;
  std::vector<int> saved_value_column_widths_;
  std::vector<bool> saved_value_column_visible_;
  bool saved_value_columns_loaded_ = false;
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
  int value_sort_column_ = 0;
  bool value_sort_ascending_ = true;
  int history_sort_column_ = 0;
  bool history_sort_ascending_ = true;
  int history_max_rows_ = 500;
  std::vector<HistoryEntry> history_entries_;
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
  std::vector<std::wstring> address_history_;
  std::vector<std::wstring> nav_history_;
  int nav_index_ = -1;
  bool nav_is_programmatic_ = false;
  bool suppress_tab_change_ = false;
  std::vector<RegistryRootEntry> roots_;
  RegistryNode* current_node_ = nullptr;
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
  std::mutex tree_state_mutex_;
  std::condition_variable tree_state_cv_;
  std::thread tree_state_thread_;
  bool tree_state_stop_ = false;
  bool tree_state_dirty_ = false;
  std::wstring tree_state_selected_;
  std::vector<std::wstring> tree_state_expanded_;
  bool always_on_top_ = false;
  bool always_run_as_admin_ = false;
  bool always_run_as_system_ = false;
  bool always_run_as_trustedinstaller_ = false;
  bool replace_regedit_ = false;
  bool single_instance_ = true;
  bool read_only_ = false;
  ThemeMode theme_mode_ = ThemeMode::kSystem;
  std::wstring icon_set_ = L"default";
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
  std::unordered_map<std::wstring, std::wstring> hive_list_;
  std::wstring saved_tree_selected_path_;
  std::vector<std::wstring> saved_tree_expanded_paths_;
  bool tree_state_restored_ = false;
  bool window_placement_loaded_ = false;
  int window_x_ = 0;
  int window_y_ = 0;
  int window_width_ = 0;
  int window_height_ = 0;
  bool window_maximized_ = false;
  ClipboardItem clipboard_;
  std::vector<UndoOperation> undo_stack_;
  std::vector<UndoOperation> redo_stack_;
  ReplaceDialogResult last_replace_;
  SearchDialogResult last_search_;
  std::vector<SearchResult> last_search_results_;
  size_t last_search_index_ = 0;

  struct SearchTab {
    std::wstring label;
    std::vector<SearchResult> results;
    uint64_t generation = 0;
    bool is_compare = false;
    size_t last_ui_count = 0;
    int sort_column = -1;
    bool sort_ascending = true;
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
    bool offline_dirty = false;
    std::wstring reg_file_path;
    std::wstring reg_file_label;
    struct RegFileRoot {
      HKEY root = nullptr;
      std::wstring name;
      std::shared_ptr<RegistryProvider::VirtualRegistryData> data;
    };
    std::vector<RegFileRoot> reg_file_roots;
    bool reg_file_dirty = false;
    bool reg_file_loading = false;
  };

  struct PendingSearchResult {
    uint64_t generation = 0;
    SearchResult result;
  };

  struct TraceKeyValues {
    std::unordered_set<std::wstring> values_lower;
    std::vector<std::wstring> values_display;
  };

  struct TraceData {
    std::wstring label;
    std::wstring source_path;
    std::unordered_map<std::wstring, TraceKeyValues> values_by_key;
    std::unordered_map<std::wstring, std::vector<std::wstring>> children_by_key;
    std::vector<std::wstring> key_paths;
    std::vector<std::wstring> display_key_paths;
    std::unordered_map<std::wstring, std::wstring> display_to_key;
    std::shared_ptr<std::shared_mutex> mutex = std::make_shared<std::shared_mutex>();
  };
  struct TraceLoadPayload;

  struct DefaultValueEntry {
    DWORD type = REG_NONE;
    std::wstring data;
  };

  struct DefaultKeyValues {
    std::unordered_map<std::wstring, DefaultValueEntry> values;
  };

  struct DefaultData {
    std::unordered_map<std::wstring, DefaultKeyValues> values_by_key;
    std::shared_ptr<std::shared_mutex> mutex = std::make_shared<std::shared_mutex>();
  };
  struct DefaultLoadPayload;

  struct CommentEntry {
    std::wstring path;
    std::wstring name;
    DWORD type = 0;
    std::wstring text;
  };

  HWND search_results_list_ = nullptr;
  std::vector<TabEntry> tabs_;
  std::vector<SearchTab> search_tabs_;
  std::vector<PendingSearchResult> search_pending_;
  std::mutex search_mutex_;
  std::atomic_bool search_posted_{false};
  std::atomic_bool search_cancel_{false};
  std::atomic<uint64_t> search_progress_searched_{0};
  std::atomic<uint64_t> search_progress_total_{0};
  std::atomic_bool search_progress_posted_{false};
  int search_progress_percent_ = 0;
  uint64_t search_last_refresh_tick_ = 0;
  uint64_t search_progress_last_tick_ = 0;
  uint64_t search_start_tick_ = 0;
  uint64_t search_duration_ms_ = 0;
  bool search_duration_valid_ = false;
  std::thread search_thread_;
  bool search_running_ = false;
  uint64_t search_generation_ = 0;
  int active_search_tab_index_ = -1;
  int search_results_view_tab_index_ = -1;
  int tab_hot_index_ = -1;
  int tab_close_hot_index_ = -1;
  int tab_close_down_index_ = -1;
  int last_tab_index_ = -1;
  bool tab_mouse_tracking_ = false;
  DWORD last_value_click_time_ = 0;
  DWORD last_value_click_delta_ = 0;
  int last_value_click_index_ = -1;
  bool last_value_click_delta_valid_ = false;
  bool value_activate_from_key_ = false;
  std::wstring type_buffer_tree_;
  std::wstring type_buffer_list_;
  DWORD type_buffer_tree_tick_ = 0;
  DWORD type_buffer_list_tick_ = 0;
  ::IAutoComplete2* address_autocomplete_ = nullptr;
  ::IEnumString* address_autocomplete_source_ = nullptr;
  struct ActiveTrace {
    std::wstring label;
    std::wstring source_path;
    std::shared_ptr<const TraceData> data;
    TraceSelection selection;
  };
  struct ActiveDefault {
    std::wstring label;
    std::wstring source_path;
    std::shared_ptr<const DefaultData> data;
    KeyValueSelection selection;
  };
  struct TraceParseSession {
    std::wstring label;
    std::wstring source_path;
    std::wstring source_lower;
    std::shared_ptr<TraceData> data;
    TraceSelection selection;
    std::thread thread;
    std::atomic_bool cancel{false};
    HWND dialog = nullptr;
    bool added_to_active = false;
    bool parsing_done = false;
  };
  struct DefaultParseSession {
    std::wstring label;
    std::wstring source_path;
    std::wstring source_lower;
    std::shared_ptr<DefaultData> data;
    KeyValueSelection selection;
    std::thread thread;
    std::atomic_bool cancel{false};
    HWND dialog = nullptr;
    bool added_to_active = false;
    bool parsing_done = false;
    bool show_errors = true;
  };
  struct RegFileParseSession {
    std::wstring source_path;
    std::wstring source_lower;
    std::thread thread;
    std::atomic_bool cancel{false};
  };
  struct TraceDialogStartContext {
    MainWindow* window = nullptr;
    TraceParseSession* session = nullptr;
  };
  struct DefaultDialogStartContext {
    MainWindow* window = nullptr;
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
    HWND hwnd = nullptr;
    std::vector<ActiveTrace> trace_data_list;
    std::vector<ActiveDefault> default_data_list;
    std::unordered_map<std::wstring, std::wstring> hive_list;
    std::unordered_map<std::wstring, CommentEntry> value_comments;
    std::unordered_map<std::wstring, CommentEntry> name_comments;
  };
  std::vector<ActiveTrace> active_traces_;
  std::unordered_map<std::wstring, TraceSelection> trace_selection_cache_;
  std::vector<std::wstring> recent_trace_paths_;
  std::vector<ActiveDefault> active_defaults_;
  std::vector<std::wstring> recent_default_paths_;
  std::mutex value_list_mutex_;
  std::condition_variable value_list_cv_;
  std::thread value_list_thread_;
  bool value_list_stop_ = false;
  bool value_list_pending_ = false;
  std::unique_ptr<ValueListTask> value_list_task_;
  std::thread trace_load_thread_;
  std::atomic_bool trace_load_stop_{false};
  std::atomic_bool trace_load_running_{false};
  std::unordered_map<std::wstring, std::unique_ptr<TraceParseSession>> trace_parse_sessions_;
  std::thread default_load_thread_;
  std::atomic_bool default_load_stop_{false};
  std::atomic_bool default_load_running_{false};
  std::unordered_map<std::wstring, std::unique_ptr<DefaultParseSession>> default_parse_sessions_;
  std::unordered_map<std::wstring, std::unique_ptr<RegFileParseSession>> reg_file_parse_sessions_;
  uint64_t last_trace_refresh_tick_ = 0;
  uint64_t last_default_refresh_tick_ = 0;
  std::unordered_map<std::wstring, CommentEntry> value_comments_;
  std::unordered_map<std::wstring, CommentEntry> name_comments_;
  util::UniqueHKey registry_root_;

  struct BundledDefault {
    std::wstring label;
    std::wstring path;
  };

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
