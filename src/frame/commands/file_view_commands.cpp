// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"
#include "regfile/script_convert.h"
#include "win32/registry_view.h"
#include "win32/translation.h"

#include <win32ui/dialog/text_prompt.h>
#include <win32ui/dialog/font_picker.h>
#include <win32ui/system/shell_open.h>

namespace regkit
{
using namespace command_detail;

void MainWindow::Impl::RestoreHiveFile(const std::wstring& path)
{
    if (!browse_.current_node())
    {
        return;
    }
    const RegistryNode node = SelectedKeyNode();
    if (session_->mode != RegistryMode::kLocal || node.root == HKEY_CLASSES_ROOT || RegistryStore::IsVirtualRoot(node.root))
    {
        ui::ShowError(hwnd_, util::Tr(L"Hive files can only be restored into the local registry, outside HKEY_CLASSES_ROOT."));
        return;
    }
    // undo and redo restore these copies in one kernel step like the restore itself
    auto before = std::make_shared<win32ui::UniqueFile>();
    std::vector<std::pair<RegistryNode, changes::KeySnapshot>> volatile_keys;
    const bool can_undo = SaveKeyCopy(node.root, node.subkey, ViewOf(node), before.get()) == ERROR_SUCCESS && changes::CaptureVolatileKeys(node, &volatile_keys);
    const std::wstring message = can_undo ? util::Tr(L"Replace all contents of this key with the hive file?")
                                          : util::Tr(L"Replace all contents of this key with the hive file? The restore can't be undone.");
    if (ui::PromptKeyChoice(hwnd_, message, registry_path::Build(node) + L"\r\n" + path, util::Tr(L"Restore Key"), util::Tr(L"Restore"), L"", util::Tr(L"Cancel")) != IDYES)
    {
        return;
    }
    const LONG status = RestoreKeyFromHive(node.root, node.subkey, ViewOf(node), path);
    if (status != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, HiveTransferError(status, path));
        return;
    }
    AppendHistoryEntry(L"Restore key from hive file " + util::FileName(path), L"", path);
    auto after = std::make_shared<win32ui::UniqueFile>();
    if (can_undo && SaveKeyCopy(node.root, node.subkey, ViewOf(node), after.get()) == ERROR_SUCCESS)
    {
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kReplaceKey;
        op.node = node;
        op.hive_before = std::move(before);
        op.hive_after = std::move(after);
        op.volatile_keys = std::move(volatile_keys);
        PushUndo(std::move(op));
    }
    RefreshTreeSelection();
    UpdateValueListForNode(browse_.current_node());
}

bool MainWindow::Impl::HandleDynamicCommand(int command_id)
{
    if (command_id >= cmd::kRegEditFavoriteBase && command_id <= cmd::kRegEditFavoriteMax)
    {
        const size_t index = static_cast<size_t>(command_id - cmd::kRegEditFavoriteBase);
        if (index < regedit_favorites_.size())
        {
            NavigateToExternalJump(regedit_favorites_[index].path);
            return true;
        }
    }
    if (command_id >= cmd::kFavoritesItemBase && command_id <= cmd::kFavoritesItemMax)
    {
        if (!favorites_loaded_)
        {
            RefreshFavoritesCache();
        }
        size_t index = static_cast<size_t>(command_id - cmd::kFavoritesItemBase);
        if (index < favorites_cache_.size())
        {
            SelectTreePath(favorites_cache_[index]);
            return true;
        }
    }
    if (command_id >= cmd::kDefaultBundledBase && command_id <= cmd::kDefaultBundledMax)
    {
        size_t index = static_cast<size_t>(command_id - cmd::kDefaultBundledBase);
        if (index < bundled_defaults_.size())
        {
            const auto& entry = bundled_defaults_[index];
            if (RemoveDefaultByPath(entry.path))
            {
                return true;
            }
            LoadDefaultFromFile(entry.label, entry.path);
            return true;
        }
    }
    if (command_id >= cmd::kDefaultRecentBase && command_id <= cmd::kDefaultRecentMax)
    {
        size_t index = static_cast<size_t>(command_id - cmd::kDefaultRecentBase);
        if (index < recent_default_paths_.items().size())
        {
            std::wstring path = recent_default_paths_.items()[index];
            std::wstring label = util::FileBaseName(path);
            if (label.empty())
            {
                label = L"Default";
            }
            if (RemoveDefaultByPath(path))
            {
                return true;
            }
            if (LoadDefaultFromFile(label, path))
            {
                AddRecentDefaultPath(path);
                SaveSettings();
            }
            return true;
        }
    }
    if (command_id >= cmd::kTraceRecentBase && command_id <= cmd::kTraceRecentMax)
    {
        size_t index = static_cast<size_t>(command_id - cmd::kTraceRecentBase);
        if (index < recent_trace_paths_.items().size())
        {
            std::wstring path = recent_trace_paths_.items()[index];
            std::wstring label = util::FileBaseName(path);
            if (label.empty())
            {
                label = L"Trace";
            }
            if (RemoveTraceByPath(path))
            {
                return true;
            }
            if (LoadTraceFromFile(label, path))
            {
                AddRecentTracePath(path);
                SaveSettings();
                SaveActiveTraces();
            }
            return true;
        }
    }

    return false;
}

bool MainWindow::Impl::HandleFileCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kFileExit:
        PostMessageW(hwnd_, WM_CLOSE, 0, 0);
        return true;
    case cmd::kFileImport:
        {
            if (session_->mode != RegistryMode::kLocal || ShowsRegFile() || !EnsureWritable())
            {
                return true;
            }
            std::wstring path;
            if (!ui::PromptOpenFile(hwnd_, ui::kImportFileFilter, &path))
            {
                return true;
            }
            if (IsHiveFile(path))
            {
                RestoreHiveFile(path);
                return true;
            }
            if (!ui::ConfirmRegFileMerge(hwnd_, path))
            {
                return true;
            }
            std::wstring error;
            const bool imported = ImportRegFileFromPath(path, &error, session_->view, util::CurrentUserFollowsShell() ? util::GetCurrentUserSidString() : std::wstring());
            RegistryStore::NoteKeyChange();
            RefreshWholeTree();
            UpdateValueListForNode(browse_.current_node());
            if (!imported)
            {
                ui::ShowRegFileMergeFailed(hwnd_, path, error);
                return true;
            }
            AppendHistoryEntry(L"Import .reg file " + util::FileName(path), L"", path);
            ui::ShowRegFileMergeSucceeded(hwnd_, path);
            return true;
        }
    case cmd::kFileOpenRegFile:
        {
            std::wstring path;
            if (!ui::PromptOpenFile(hwnd_, ui::kRegFileFilter, &path))
            {
                return true;
            }
            OpenRegFileTab(path);
            return true;
        }
    case cmd::kFileSave:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (IsRegFileTabSelected())
            {
                int tab_index = TabCtrl_GetCurSel(tab_);
                if (tab_index >= 0 && static_cast<size_t>(tab_index) < tabs_.size())
                {
                    if (tabs_[static_cast<size_t>(tab_index)].reg_file_dirty)
                    {
                        SaveRegFileTab(tab_index);
                    }
                }
                return true;
            }
            if (session_->offline_dirty)
            {
                SaveOfflineRegistry(*session_);
            }
            return true;
        }
    case cmd::kFileExport:
        {
            if (IsRegFileTabSelected())
            {
                int tab_index = TabCtrl_GetCurSel(tab_);
                std::wstring path;
                win32::OpenAfter open_after = win32::OpenAfter::kNone;
                if (!ui::PromptSaveFile(hwnd_, regfile::kConvertFilter, &path, &open_after, true))
                {
                    return true;
                }
                if (ExportRegFileTab(tab_index, &path, &open_after))
                {
                    AppendHistoryEntry(L"Export .reg tab " + util::FileName(path), L"", path);
                    OpenSavedFile(path, open_after);
                }
                return true;
            }
            if (!browse_.current_node())
            {
                return true;
            }
            if (browse_.values().hwnd() && GetFocus() == browse_.values().hwnd())
            {
                std::vector<std::wstring> selected_values;
                std::vector<std::wstring> selected_keys;
                for (const ListRow* row : SelectedListRows(browse_.values()))
                {
                    if (row->kind == rowkind::kValue)
                    {
                        selected_values.push_back(row->extra);
                    }
                    else if (row->kind == rowkind::kKey)
                    {
                        selected_keys.push_back(row->extra);
                    }
                }
                if (!selected_values.empty() || !selected_keys.empty())
                {
                    auto dedupe = [](std::vector<std::wstring>* items) {
                        if (!items)
                        {
                            return;
                        }
                        std::unordered_set<std::wstring> seen;
                        std::vector<std::wstring> unique;
                        unique.reserve(items->size());
                        for (const auto& item : *items)
                        {
                            std::wstring key = ToLower(item);
                            if (seen.insert(key).second)
                            {
                                unique.push_back(item);
                            }
                        }
                        *items = std::move(unique);
                    };
                    dedupe(&selected_values);
                    dedupe(&selected_keys);
                    std::wstring error;
                    std::wstring path = registry_path::Build(*browse_.current_node());
                    std::wstring saved_path;
                    win32::OpenAfter open_after = win32::OpenAfter::kNone;
                    if (ExportRegFileSelection(hwnd_, path, ViewOf(*browse_.current_node()), selected_values, selected_keys, &error, &saved_path, &open_after))
                    {
                        HistoryEntry entry;
                        entry.action = L"Export registry selection";
                        entry.old_data = std::to_wstring(selected_keys.size()) + L" keys, " +
                                         std::to_wstring(selected_values.size()) + L" values";
                        entry.key_path = path;
                        AppendHistoryEntry(std::move(entry));
                        OpenSavedFile(saved_path, open_after);
                    }
                    else if (!error.empty())
                    {
                        ui::ShowError(hwnd_, error);
                    }
                    return true;
                }
            }
            std::wstring error;
            std::wstring path = registry_path::Build(*browse_.current_node());
            std::wstring saved_path;
            win32::OpenAfter open_after = win32::OpenAfter::kNone;
            const bool allow_hive = session_->mode == RegistryMode::kLocal && browse_.current_node()->root != HKEY_CLASSES_ROOT;
            if (ExportRegFile(hwnd_, path, ViewOf(*browse_.current_node()), allow_hive, &error, &saved_path, &open_after))
            {
                HistoryEntry entry;
                entry.action = L"Export registry key";
                entry.key_path = path;
                entry.new_data = path;
                AppendHistoryEntry(std::move(entry));
                OpenSavedFile(saved_path, open_after);
            }
            else if (!error.empty())
            {
                ui::ShowError(hwnd_, error);
            }
            return true;
        }
    case cmd::kFileImportComments:
        {
            std::wstring path;
            if (!ui::PromptOpenFile(hwnd_, L"RegKit Comment Files (*.jsonc)\0*.jsonc;*.json\0All Files (*.*)\0*.*\0\0", &path))
            {
                return true;
            }
            size_t imported = 0;
            std::wstring error;
            if (ImportCommentsFromFile(path, &imported, &error))
            {
                AppendHistoryEntry(L"Import comments " + util::FileName(path), L"", path);
                ui::ShowInfo(hwnd_, util::TrLabel(L"Comments imported", std::to_wstring(imported)));
            }
            else
            {
                ui::ShowError(hwnd_, error.empty() ? std::wstring(util::Tr(L"Failed to import comments.")) : util::TrDetail(L"Failed to import comments.", error));
            }
            return true;
        }
    case cmd::kFileExportComments:
        {
            std::wstring path;
            win32::OpenAfter open_after = win32::OpenAfter::kNone;
            if (!ui::PromptSaveFile(hwnd_, L"RegKit Comment Files (*.jsonc)\0*.jsonc\0All Files (*.*)\0*.*\0\0", &path, &open_after))
            {
                return true;
            }
            if (ExportCommentsToFile(path))
            {
                AppendHistoryEntry(L"Export comments " + util::FileName(path), L"", path);
                OpenSavedFile(path, open_after);
            }
            else
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to export comments."));
            }
            return true;
        }
    case cmd::kFileLoadHive:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (session_->mode == RegistryMode::kRemote)
            {
                ui::ShowError(hwnd_, util::Tr(L"Loading hives isn't supported for remote registries."));
                return true;
            }
            std::wstring error;
            HKEY root = HKEY_LOCAL_MACHINE;
            if (LoadHive(hwnd_, &root, &error))
            {
                const std::wstring root_name = registry_path::RootName(root);
                AppendHistoryEntry(L"Load hive", L"", root_name);
                SelectTreePath(root_name);
                RefreshTreeSelection();
                UpdateValueListForNode(browse_.current_node());
            }
            else if (!error.empty())
            {
                ui::ShowError(hwnd_, error);
            }
            return true;
        }
    case cmd::kFileUnloadHive:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (session_->mode == RegistryMode::kRemote)
            {
                ui::ShowError(hwnd_, util::Tr(L"Unloading hives isn't supported for remote registries."));
                return true;
            }
            HKEY root = HKEY_LOCAL_MACHINE;
            std::wstring subkey;
            if (browse_.current_node() &&
                (browse_.current_node()->root == HKEY_LOCAL_MACHINE || browse_.current_node()->root == HKEY_USERS))
            {
                root = browse_.current_node()->root;
                subkey = browse_.current_node()->subkey;
            }
            if (subkey.empty() || subkey.find(L'\\') != std::wstring::npos)
            {
                ui::ShowError(hwnd_, util::Tr(L"Select a hive you loaded under HKEY_LOCAL_MACHINE or HKEY_USERS first."));
                return true;
            }
            if (ui::PromptKeyChoice(hwnd_, util::Tr(L"Unload this key and all of its subkeys?"), subkey, util::Tr(L"Unload Hive"), util::Tr(L"Unload"), L"", util::Tr(L"Cancel")) != IDYES)
            {
                return true;
            }
            std::wstring error;
            if (!UnloadHive(hwnd_, root, subkey, &error))
            {
                if (!error.empty())
                {
                    ui::ShowError(hwnd_, error);
                }
                return true;
            }
            AppendHistoryEntry(L"Unload hive", subkey, L"");
            SelectTreePath(registry_path::RootName(root));
            RefreshTreeSelection();
            UpdateValueListForNode(browse_.current_node());
            return true;
        }
    case cmd::kFileSaveOfflineHive:
        SaveOfflineRegistry(*session_, true);
        return true;
    case cmd::kFileClearHistoryOnExit:
        settings_.clear_history_on_exit = !settings_.clear_history_on_exit;
        SaveSettings();
        return true;
    case cmd::kFileClearTabsOnExit:
        settings_.clear_tabs_on_exit = !settings_.clear_tabs_on_exit;
        SaveSettings();
        return true;
    case cmd::kFileClearCacheAll:
    case cmd::kFileClearCacheTabs:
    case cmd::kFileClearCacheHistory:
    case cmd::kFileClearCacheSearchHistory:
    case cmd::kFileClearCacheDialogState:
    case cmd::kFileClearCacheTreeState:
    case cmd::kFileClearCacheTemporary:
        {
            CacheKind kind = CacheKind::kAll;
            switch (command_id)
            {
            case cmd::kFileClearCacheTabs:
                kind = CacheKind::kTabs;
                break;
            case cmd::kFileClearCacheHistory:
                kind = CacheKind::kHistory;
                break;
            case cmd::kFileClearCacheSearchHistory:
                kind = CacheKind::kSearchHistory;
                break;
            case cmd::kFileClearCacheDialogState:
                kind = CacheKind::kDialogState;
                break;
            case cmd::kFileClearCacheTreeState:
                kind = CacheKind::kTreeState;
                break;
            case cmd::kFileClearCacheTemporary:
                kind = CacheKind::kTemporary;
                break;
            default:
                break;
            }
            cache_clear_on_close_ = kind;
            PostMessageW(hwnd_, WM_CLOSE, 0, 0);
            return true;
        }
    case cmd::kFileRestart:
        restart_on_close_ = true;
        PostMessageW(hwnd_, WM_CLOSE, 0, 0);
        return true;
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleViewCommand(int command_id)
{
    for (const PaneToggle& toggle : kPaneToggles)
    {
        if (toggle.command == command_id)
        {
            settings_.*toggle.shown = !(settings_.*toggle.shown);
            SaveSettings();
            ApplyViewVisibility();
            return true;
        }
    }
    switch (command_id)
    {
    case cmd::kViewRefresh:
        RegistryStore::NoteKeyChange();
        RefreshWholeTree();
        UpdateValueListForNode(browse_.current_node());
        return true;
    case cmd::kViewFocusFilter:
        if (!settings_.show_filter_bar)
        {
            settings_.show_filter_bar = true;
            SaveSettings();
            ApplyViewVisibility();
        }
        if (browse_.filter())
        {
            SetFocus(browse_.filter());
            SendMessageW(browse_.filter(), EM_SETSEL, 0, -1);
        }
        return true;
    case cmd::kTreeToggleExpand:
        {
            if (!browse_.tree().hwnd())
            {
                return true;
            }
            HTREEITEM item = TreeView_GetSelection(browse_.tree().hwnd());
            if (!item)
            {
                return true;
            }
            TVITEMW tvi = {};
            tvi.hItem = item;
            tvi.mask = TVIF_STATE | TVIF_CHILDREN;
            tvi.stateMask = TVIS_EXPANDED;
            if (!TreeView_GetItem(browse_.tree().hwnd(), &tvi))
            {
                return true;
            }
            bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
            bool has_child = TreeView_GetChild(browse_.tree().hwnd(), item) != nullptr || tvi.cChildren != 0;
            if (!expanded && !has_child)
            {
                return true;
            }
            TreeView_Expand(browse_.tree().hwnd(), item, expanded ? TVE_COLLAPSE : TVE_EXPAND);
            return true;
        }
    case cmd::kTreeExpandAll:
        {
            HWND tree = browse_.tree().hwnd();
            HTREEITEM root = tree ? TreeView_GetSelection(tree) : nullptr;
            if (!root)
            {
                return true;
            }
            HCURSOR previous = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
            browse_.tree().SuspendRedraw();
            constexpr int kExpandAllKeyLimit = 5000;
            std::vector<HTREEITEM> pending{root};
            int expanded_keys = 0;
            while (!pending.empty() && expanded_keys < kExpandAllKeyLimit)
            {
                HTREEITEM item = pending.back();
                pending.pop_back();
                TreeView_Expand(tree, item, TVE_EXPAND);
                ++expanded_keys;
                for (HTREEITEM child = TreeView_GetChild(tree, item); child; child = TreeView_GetNextSibling(tree, child))
                {
                    pending.push_back(child);
                }
            }
            const bool truncated = !pending.empty();
            browse_.tree().ResumeRedraw();
            TreeView_EnsureVisible(tree, root);
            SetCursor(previous);
            MarkTreeStateDirty();
            if (truncated)
            {
                ui::ShowWarning(hwnd_, util::TrDetail(L"This key has too many subkeys to expand at once.", util::TrLabel(L"Expanded keys", std::to_wstring(expanded_keys))));
            }
            return true;
        }
    case cmd::kViewSelectAll:
        if (!SelectAllInFocusedList())
        {
            HWND focus = GetFocus();
            if (focus)
            {
                SendMessageW(focus, EM_SETSEL, 0, -1);
            }
        }
        return true;
    case cmd::kEditInvertSelection:
        InvertSelectionInFocusedList();
        return true;
    case cmd::kViewGridLines:
        SetValueGridEnabled(!settings_.show_value_grid, true);
        return true;
    case cmd::kViewAutoRefresh:
        settings_.auto_refresh = !settings_.auto_refresh;
        SaveSettings();
        WatchCurrentKey();
        return true;
    case cmd::kViewKeysInList:
        settings_.show_keys_in_list = !settings_.show_keys_in_list;
        UpdateValueListForNode(browse_.current_node());
        SaveSettings();
        return true;
    case cmd::kViewSimulatedKeys:
        settings_.show_simulated_keys = !settings_.show_simulated_keys;
        RefreshTreeSelection();
        UpdateValueListForNode(browse_.current_node());
        SaveSettings();
        return true;
    case cmd::kViewSimpleMode:
        settings_.simple_mode = !settings_.simple_mode;
        SaveSettings();
        BuildMenus();
        ApplyValueColumns();
        ApplyViewVisibility();
        ReloadLocalRoots();
        return true;
    case cmd::kViewExtraHives:
        settings_.show_extra_hives = !settings_.show_extra_hives;
        SaveSettings();
        ReloadLocalRoots();
        return true;
    case cmd::kViewSaveTreeState:
        settings_.save_tree_state = !settings_.save_tree_state;
        if (settings_.save_tree_state)
        {
            StartTreeStateWorker();
            MarkTreeStateDirty();
        }
        else
        {
            tree_state_saver_.Stop();
        }
        SaveSettings();
        return true;
    case cmd::kOptionsSaveTabs:
    case cmd::kOptionsSaveTabsLocal:
    case cmd::kOptionsSaveTabsOffline:
    case cmd::kOptionsSaveTabsRemote:
    case cmd::kOptionsSaveTabsSearch:
    case cmd::kOptionsSaveTabsCompare:
    case cmd::kOptionsSaveTabsRegFile:
        {
            if (command_id == cmd::kOptionsSaveTabs)
            {
                settings_.save_tab_kinds =
                    (settings_.save_tab_kinds & workspace::kSaveTabsAll) == workspace::kSaveTabsAll ? 0 : workspace::kSaveTabsAll;
            }
            else
            {
                settings_.save_tab_kinds ^= 1 << (command_id - cmd::kOptionsSaveTabsLocal);
            }
            if (settings_.save_tab_kinds == 0)
            {
                ClearTabsCache();
            }
            else
            {
                SaveTabs();
            }
            SaveSettings();
            return true;
        }
    case cmd::kOptionsReadOnly:
        settings_.read_only = !settings_.read_only;
        SaveSettings();
        UpdateToolbarStates();
        return true;
    case cmd::kOptionsCompareRegistries:
        StartCompareRegistries();
        return true;
    case cmd::kViewFont:
        {
            const LOGFONTW default_font = win32ui::MessageFontLogFont(0);
            win32ui::FontChoice result;
            if (win32ui::ChooseUiFont(hwnd_, settings_.use_custom_font ? custom_font_ : default_font, default_font, &result) == ERROR_SUCCESS)
            {
                settings_.use_custom_font = !result.use_default;
                custom_font_ = result.font;
                UpdateUIFont();
                SaveSettings();
            }
            return true;
        }
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleTraceDefaultCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kTraceLoad23H2:
        if (RemoveTraceByLabel(L"23H2"))
        {
            return true;
        }
        LoadBundledTrace(L"23H2");
        return true;
    case cmd::kTraceLoad24H2:
        if (RemoveTraceByLabel(L"24H2"))
        {
            return true;
        }
        LoadBundledTrace(L"24H2");
        return true;
    case cmd::kTraceLoad25H2:
        if (RemoveTraceByLabel(L"25H2"))
        {
            return true;
        }
        LoadBundledTrace(L"25H2");
        return true;
    case cmd::kTraceLoadCustom:
        LoadTraceFromPrompt();
        return true;
    case cmd::kTraceClear:
        ClearTrace();
        return true;
    case cmd::kTraceClearRecent:
        recent_trace_paths_.Assign({});
        SaveSettings();
        return true;
    case cmd::kDefaultLoadCustom:
        LoadDefaultFromPrompt();
        return true;
    case cmd::kDefaultClear:
        ClearDefaults();
        return true;
    case cmd::kDefaultClearRecent:
        recent_default_paths_.Assign({});
        SaveSettings();
        return true;
    case cmd::kDefaultResetEnable:
        settings_.default_reset_enabled = !settings_.default_reset_enabled;
        SaveSettings();
        return true;
    case cmd::kDefaultEditActive:
        {
            std::vector<std::wstring> active;
            active.reserve(active_defaults_.size());
            for (const auto& defaults : active_defaults_)
            {
                active.push_back(defaults.source_path);
            }
            std::wstring content = JoinLines(active);
            if (win32ui::PromptText(hwnd_, {.title = util::Tr(L"Edit Active Defaults"), .label = util::Tr(L"One default path per line."), .multiline = true}, &content) == 0)
            {
                std::vector<std::wstring> lines = SplitLines(content);
                default_load_session_.Cancel();
                // parse still running would make its line count as added
                StopDefaultParseSessions();
                active_defaults_.clear();
                for (const auto& line : lines)
                {
                    AddDefaultFromFile(L"", line, false, false, false);
                }
                SaveActiveDefaults();
                UpdateValueListForNode(browse_.current_node());
                SaveSettings();
            }
            return true;
        }
    case cmd::kTraceEditActive:
        {
            std::vector<std::wstring> active;
            active.reserve(active_traces_.size());
            for (const auto& trace : active_traces_)
            {
                active.push_back(trace.source_path);
            }
            std::wstring content = JoinLines(active);
            if (win32ui::PromptText(hwnd_, {.title = util::Tr(L"Edit Active Traces"), .label = util::Tr(L"One trace path per line."), .multiline = true}, &content) == 0)
            {
                std::vector<std::wstring> lines = SplitLines(content);
                LoadTraceSettings();
                trace_load_session_.Cancel();
                StopTraceParseSessions();
                active_traces_.clear();
                for (const auto& line : lines)
                {
                    AddTraceFromFile(L"", line, nullptr, false, false);
                }
                SaveActiveTraces();
                SaveTraceSettings();
                RefreshTreeSelection();
                UpdateValueListForNode(browse_.current_node());
                SaveSettings();
            }
            return true;
        }
    case cmd::kTraceGuide:
        ui::ReportFileDialogResult(hwnd_, win32ui::ShellOpen(hwnd_, L"https://noverse.dev/docs/regkit/guides/wpr-wpa/"));
        return true;
    case cmd::kDefaultEditRecent:
        {
            std::wstring content = JoinLines(recent_default_paths_.items());
            if (win32ui::PromptText(hwnd_, {.title = util::Tr(L"Edit Recent Defaults"), .label = util::Tr(L"One default path per line."), .multiline = true}, &content) == 0)
            {
                std::vector<std::wstring> updated = SplitLines(content);
                recent_default_paths_.Assign(std::move(updated));
                SaveSettings();
            }
            return true;
        }
    case cmd::kTraceEditRecent:
        {
            std::wstring content = JoinLines(recent_trace_paths_.items());
            if (win32ui::PromptText(hwnd_, {.title = util::Tr(L"Edit Recent Traces"), .label = util::Tr(L"One trace path per line."), .multiline = true}, &content) == 0)
            {
                std::vector<std::wstring> updated = SplitLines(content);
                recent_trace_paths_.Assign(std::move(updated));
                SaveSettings();
            }
            return true;
        }
    default:
        return false;
    }
}

}
