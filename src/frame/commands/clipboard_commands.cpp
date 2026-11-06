// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"
#include "win32/translation.h"

#include <win32ui/dialog/text_prompt.h>

namespace regkit
{
using namespace command_detail;

void MainWindow::Impl::ShowFavoritesImported(size_t imported)
{
    ui::ShowInfo(hwnd_, imported ? util::TrLabel(L"New favorites imported", std::to_wstring(imported)) : std::wstring(util::Tr(L"There are no new favorites to import.")));
}

bool MainWindow::Impl::HandleFavoritesCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kFavoritesAdd:
        {
            if (browse_.current_node())
            {
                if (FavoritesStore::Add(registry_path::Build(*browse_.current_node())))
                {
                    RefreshFavoritesCache();
                }
            }
            return true;
        }
    case cmd::kFavoritesRemove:
        {
            if (browse_.current_node())
            {
                if (FavoritesStore::Remove(registry_path::Build(*browse_.current_node())))
                {
                    RefreshFavoritesCache();
                }
            }
            return true;
        }
    case cmd::kFavoritesEdit:
        {
            if (!favorites_loaded_)
            {
                RefreshFavoritesCache();
            }
            std::wstring content = JoinLines(favorites_cache_);
            if (win32ui::PromptText(hwnd_, {.title = util::Tr(L"Edit Favorites"), .label = util::Tr(L"Each line should include one key."), .multiline = true}, &content) == 0)
            {
                std::vector<std::wstring> updated = SplitLines(content);
                if (!FavoritesStore::Save(updated))
                {
                    ui::ShowError(hwnd_, util::Tr(L"Failed to save favorites."));
                    return true;
                }
                favorites_cache_ = std::move(updated);
                favorites_loaded_ = true;
            }
            return true;
        }
    case cmd::kFavoritesImport:
        {
            std::wstring path;
            if (!ui::PromptOpenFile(hwnd_, L"Favorites Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0", &path))
            {
                return true;
            }
            size_t imported = 0;
            if (!FavoritesStore::ImportFromFile(path, &imported))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to import favorites."));
                return true;
            }
            RefreshFavoritesCache();
            AppendHistoryEntry(L"Import favorites " + util::FileName(path), L"", path);
            ShowFavoritesImported(imported);
            return true;
        }
    case cmd::kFavoritesImportRegEdit:
        {
            size_t imported = 0;
            std::wstring error;
            if (!FavoritesStore::ImportFromRegEdit(&imported, &error))
            {
                ui::ShowError(hwnd_, error.empty() ? util::Tr(L"Failed to import RegEdit favorites.") : error);
                return true;
            }
            if (imported > 0)
            {
                RefreshFavoritesCache();
            }
            AppendHistoryEntry(L"Import RegEdit favorites", L"", std::to_wstring(imported) + L" favorites");
            ShowFavoritesImported(imported);
            return true;
        }
    case cmd::kFavoritesExport:
        {
            std::wstring path;
            win32::OpenAfter open_after = win32::OpenAfter::kNone;
            if (!ui::PromptSaveFile(hwnd_, L"Favorites Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0", &path, &open_after))
            {
                return true;
            }
            if (!FavoritesStore::ExportToFile(path))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to export favorites."));
            }
            else
            {
                AppendHistoryEntry(L"Export favorites " + util::FileName(path), L"", path);
                OpenSavedFile(path, open_after);
            }
            return true;
        }
    default:
        return false;
    }
}

RegistryNode MainWindow::Impl::SelectedKeyNode() const
{
    int index = -1;
    const ListRow* row = SelectedValueRow(browse_.values(), &index);
    return row && row->kind == rowkind::kKey && !row->extra.empty() ? ChildNode(*browse_.current_node(), row->extra) : *browse_.current_node();
}

bool MainWindow::Impl::HandleNavigateClipboardCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kEditCopyKey:
    case cmd::kEditCopyValueName:
    case cmd::kEditCopyValueData:
    case cmd::kEditCopyKeyPath:
    case cmd::kEditCopyKeyPathAbbrev:
    case cmd::kEditCopyKeyPathRegEdit:
    case cmd::kEditCopyKeyPathRegFile:
    case cmd::kEditCopyKeyPathPowerShell:
    case cmd::kEditCopyKeyPathPowerShellProvider:
    case cmd::kEditCopyKeyPathEscaped:
    case cmd::kEditCopyKeyPathNative:
    case cmd::kEditCopyKeyPathNativeResolved:
    case cmd::kEditCopy:
        return HandleClipboardCommand(command_id);
    case cmd::kEditGoTo:
    case cmd::kEditPermissions:
    case cmd::kEditKeyInfo:
    case cmd::kEditOpenSourceHive:
    case cmd::kEditFind:
        return HandleEditToolsCommand(command_id);
    case cmd::kEditPaste:
    case cmd::kEditReplace:
    case cmd::kEditUndo:
    case cmd::kEditRedo:
        return HandleChangeHistoryCommand(command_id);
    case cmd::kRegistryLocal:
    case cmd::kRegistryLocal32:
    case cmd::kRegistryOtherView:
    case cmd::kEditGoToVirtualStore:
    case cmd::kEditGoToGlobalKey:
    case cmd::kRegistryNetwork:
    case cmd::kRegistryOffline:
    case cmd::kNavBack:
    case cmd::kNavForward:
    case cmd::kNavUp:
        return HandleRegistryNavigationCommand(command_id);
    default:
        return false;
    }
}

void MainWindow::Impl::CaptureKeyForPaste(const RegistryNode& node)
{
    auto hold = std::make_shared<SourceHold>();
    hold->sessions.push_back(session_);
    RegistryNode source = node;
    source.root = hold->Hold(source.root);
    // stays incomplete when the capture never ends, a failed start or a throwing capture
    auto snapshot = std::make_shared<changes::KeySnapshot>(changes::KeySnapshot{.complete = false});
    key_copy_session_.StartOrRun(L"KeyCopyThread", [source, hold, snapshot](uint64_t, std::atomic_bool& cancel) {
        *snapshot = changes::CaptureKey(source, false, &cancel);
    });
    clipboard_.kind = ClipboardItem::Kind::kKey;
    clipboard_.key_snapshot = std::move(snapshot);
}

bool MainWindow::Impl::HandleClipboardCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kEditCopyKey:
        {
            std::wstring name;
            int index = -1;
            const ListRow* row = SelectedValueRow(browse_.values(), &index);
            if (row && row->kind == rowkind::kKey)
            {
                name = row->name;
            }
            else if (browse_.current_node())
            {
                name = registry_path::DisplayName(LeafName(*browse_.current_node()));
            }
            if (!name.empty())
            {
                ui::CopyText(hwnd_, name);
            }
            return true;
        }
    case cmd::kEditCopyValueName:
        {
            const std::vector<const ListRow*> selected_rows = SelectedListRows(browse_.values());
            if (selected_rows.size() != 1 || selected_rows.front()->kind != rowkind::kValue)
            {
                return true;
            }
            const ListRow* row = selected_rows.front();
            std::wstring name = row->extra.empty() ? L"(Default)" : row->extra;
            ui::CopyText(hwnd_, name);
            return true;
        }
    case cmd::kEditCopyValueData:
        {
            if (!browse_.current_node())
            {
                return true;
            }
            const ListRow* row = SelectedValueRow(browse_.values(), nullptr);
            if (!row || row->kind != rowkind::kValue)
            {
                return true;
            }
            if (row->simulated)
            {
                return true;
            }
            RegistryValue entry;
            if (!GetValueEntry(*browse_.current_node(), row->extra, &entry))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to read value."));
                return true;
            }
            std::wstring data =
                value_format::DisplayData(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
            ui::CopyText(hwnd_, data);
            return true;
        }
    case cmd::kEditCopyKeyPath:
    case cmd::kEditCopyKeyPathAbbrev:
    case cmd::kEditCopyKeyPathRegEdit:
    case cmd::kEditCopyKeyPathRegFile:
    case cmd::kEditCopyKeyPathPowerShell:
    case cmd::kEditCopyKeyPathPowerShellProvider:
    case cmd::kEditCopyKeyPathEscaped:
    case cmd::kEditCopyKeyPathNative:
    case cmd::kEditCopyKeyPathNativeResolved:
        {
            if (!browse_.current_node())
            {
                return true;
            }
            std::wstring path = registry_path::Build(*browse_.current_node());
            const ListRow* row = SelectedValueRow(browse_.values(), nullptr);
            if (row && row->kind == rowkind::kKey && !row->extra.empty())
            {
                path += L"\\" + row->name;
            }
            CopyKeyPathAs(command_id, SelectedKeyNode(), path);
            return true;
        }
    case cmd::kEditCopy:
        {
            const auto clear_clipboard = [this] {
                key_copy_session_.Cancel();
                clipboard_ = {};
            };
            const auto capture_row = [this, &clear_clipboard](const ListRow* row) {
                const RegistryNode& node = *browse_.current_node();
                clear_clipboard();
                if (row && row->kind == rowkind::kValue)
                {
                    if (!GetValueEntry(node, row->extra, &clipboard_.value))
                    {
                        return false;
                    }
                    clipboard_.kind = ClipboardItem::Kind::kValue;
                    clipboard_.name = clipboard_.value.name;
                }
                else if (row && row->kind == rowkind::kKey)
                {
                    CaptureKeyForPaste(ChildNode(node, row->extra));
                    clipboard_.name = row->extra;
                }
                clipboard_.source_parent = node;
                return true;
            };
            HWND focus = GetFocus();
            if (focus == browse_.tree().hwnd() && browse_.current_node() && !browse_.current_node()->subkey.empty())
            {
                const RegistryNode& node = *browse_.current_node();
                RegistryNode parent = node;
                parent.subkey = registry_path::Parent(node.subkey);
                CaptureKeyForPaste(node);
                clipboard_.source_parent = parent;
                clipboard_.name = LeafName(node);
                ui::CopyText(hwnd_, registry_path::Build(node));
                return true;
            }
            if (focus == browse_.values().hwnd() || focus == search_results_list_ || focus == history_list_)
            {
                HWND list = focus;
                int selected = ListView_GetSelectedCount(list);
                if (selected > 0)
                {
                    std::wstring text = win32ui::SelectedRowsText(list);
                    if (!text.empty())
                    {
                        ui::CopyText(hwnd_, text);
                    }
                    if (list == browse_.values().hwnd())
                    {
                        clear_clipboard();
                        if (selected == 1 && browse_.current_node())
                        {
                            capture_row(SelectedValueRow(browse_.values(), nullptr));
                        }
                    }
                    return true;
                }
            }
            if (!browse_.current_node())
            {
                return true;
            }
            const ListRow* row = SelectedValueRow(browse_.values(), nullptr);
            if (!capture_row(row))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to read value."));
                return true;
            }
            const bool key_row = row && row->kind == rowkind::kKey;
            ui::CopyText(hwnd_, row && row->kind == rowkind::kValue ? row->name : registry_path::Build(key_row ? ChildNode(*browse_.current_node(), row->extra) : *browse_.current_node()));
            return true;
        }
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleEditToolsCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kEditGoTo:
        if (browse_.address())
        {
            SetFocus(browse_.address());
            SendMessageW(browse_.address(), EM_SETSEL, 0, -1);
        }
        return true;
    case cmd::kEditPermissions:
    case cmd::kEditKeyInfo:
    case cmd::kEditOpenSourceHive:
        if (command_id == cmd::kEditPermissions && !EnsureWritable())
        {
            return true;
        }
        if (browse_.current_node())
        {
            const RegistryNode node = SelectedKeyNode();
            if (command_id == cmd::kEditPermissions)
            {
                ShowPermissionsDialog(node);
            }
            else if (command_id == cmd::kEditKeyInfo)
            {
                ShowKeyInfoDialog(node);
            }
            else
            {
                KeyDetails details;
                const std::wstring source = RegistryStore::QueryKeyDetails(node, &details) ? registry_path::ClassesSourcePath(details.native.native_name, util::GetCurrentUserSidString()) : std::wstring();
                if (source.empty() || !SelectTreePath(source))
                {
                    ui::ShowError(hwnd_, util::Tr(L"The source key couldn't be opened.") + std::wstring(L"\n") + registry_path::Build(node));
                }
            }
        }
        return true;
    case cmd::kEditFind:
        {
            SearchDialogResult options;
            options.criteria.search_keys = false;
            workspace::DialogState state = LoadDialogState();
            workspace::DialogFields last(&state, false);
            SearchDialogFields(last, &options);
            SearchSources sources;
            // simple mode hides traces, defaults, REGISTRY and anomalies
            sources.traces = !settings_.simple_mode && HasActiveTraces();
            sources.defaults = !settings_.simple_mode && !active_defaults_.empty();
            sources.registry_root = !settings_.simple_mode && registry_root_.get() != nullptr;
            sources.anomalies = !settings_.simple_mode;
            sources.offline = !session_->offline_roots.empty();
            sources.reg_files = std::any_of(tabs_.begin(), tabs_.end(), [](const TabEntry& tab) {
                return tab.kind == TabEntry::Kind::kRegFile && !tab.reg_file_roots.empty();
            });
            sources.remote = session_->remote_hklm != nullptr;
            sources.extra_hives = settings_.show_extra_hives;
            if (ShowSearchDialog(hwnd_, &options, sources))
            {
                workspace::DialogFields store(&state, true);
                SearchDialogFields(store, &options);
                ui::ReportSaveFailure(hwnd_, workspace::DialogStatePath(), workspace::SaveDialogState(state));
                StartSearch(options);
            }
            return true;
        }
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleChangeHistoryCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kEditPaste:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            if (!browse_.current_node() || clipboard_.kind == ClipboardItem::Kind::kNone)
            {
                return true;
            }
            if (clipboard_.kind == ClipboardItem::Kind::kValue)
            {
                bool same_parent = SameNode(*browse_.current_node(), clipboard_.source_parent);
                std::wstring base_name = clipboard_.name;
                if (same_parent)
                {
                    if (base_name.empty())
                    {
                        base_name = L"Default - Copy";
                    }
                    else
                    {
                        base_name += L" - Copy";
                    }
                }
                const std::optional<std::wstring> free_name = MakeUniqueValueName(*browse_.current_node(), base_name);
                const std::wstring unique = free_name.value_or(L"");
                RegistryValue new_value = clipboard_.value;
                new_value.name = unique;
                if (!free_name || !RegistryStore::SetValue(*browse_.current_node(), unique, new_value.type, new_value.data))
                {
                    ui::ShowError(hwnd_, util::Tr(L"Failed to paste value."));
                }
                else
                {
                    std::wstring data_text = value_format::Data(new_value.type, new_value.data.data(), static_cast<DWORD>(new_value.data.size()));
                    AppendValueHistoryEntry(L"Create value " + unique, L"", data_text, *browse_.current_node(), unique, HistoryEntry::RevertKind::kDeleteValue);
                    MarkOfflineDirty();
                    changes::UndoOperation op;
                    op.type = changes::UndoOperation::Type::kCreateValue;
                    op.node = *browse_.current_node();
                    op.name = unique;
                    op.new_value = new_value;
                    PushUndo(std::move(op));
                    UpdateValueListForNode(browse_.current_node());
                }
                return true;
            }
            if (clipboard_.kind == ClipboardItem::Kind::kKey)
            {
                key_copy_session_.Join();
                // partial key copies would paste without unreadable data
                if (!clipboard_.key_snapshot->complete)
                {
                    ui::ShowError(hwnd_, util::TrDetail(L"The key information couldn't be read.", registry_path::Build(ChildNode(clipboard_.source_parent, clipboard_.name))));
                    return true;
                }
                bool same_parent = SameNode(*browse_.current_node(), clipboard_.source_parent);
                std::wstring base_name = clipboard_.name;
                if (same_parent && !base_name.empty())
                {
                    base_name += L" - Copy";
                }
                std::wstring unique = MakeUniqueKeyName(*browse_.current_node(), base_name);
                changes::KeySnapshot snapshot = *clipboard_.key_snapshot;
                snapshot.name = unique;
                if (!changes::RestoreKey(*browse_.current_node(), snapshot))
                {
                    RefreshTreeSelection();
                    UpdateValueListForNode(browse_.current_node());
                    ui::ShowError(hwnd_, util::Tr(L"Failed to paste key."));
                }
                else
                {
                    AppendHistoryEntry(L"Create key " + unique, L"", L"");
                    MarkOfflineDirty();
                    changes::UndoOperation op;
                    op.type = changes::UndoOperation::Type::kCreateKey;
                    op.node = *browse_.current_node();
                    op.name = unique;
                    op.key_snapshot = std::move(snapshot);
                    PushUndo(std::move(op));
                    RefreshTreeSelection();
                    UpdateValueListForNode(browse_.current_node());
                }
                return true;
            }
            return true;
        }
    case cmd::kEditReplace:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            ReplaceDialogResult options;
            workspace::DialogState state = LoadDialogState();
            workspace::DialogFields last(&state, false);
            ReplaceDialogFields(last, &options);
            if (options.start_key.empty() && browse_.current_node())
            {
                options.start_key = registry_path::Build(*browse_.current_node());
            }
            if (ShowReplaceDialog(hwnd_, &options))
            {
                workspace::DialogFields store(&state, true);
                ReplaceDialogFields(store, &options);
                ui::ReportSaveFailure(hwnd_, workspace::DialogStatePath(), workspace::SaveDialogState(state));
                StartReplace(options);
            }
            return true;
        }
    case cmd::kEditUndo:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            auto operation = session_->undo.TakeUndo();
            if (!operation)
            {
                return true;
            }
            size_t replayed = 0;
            bool left_both_names = false;
            switch (ApplyUndoOperation(*operation, false, &replayed, &left_both_names))
            {
            case ReplayResult::kSuccess:
                session_->undo.CompleteUndo(std::move(*operation));
                break;
            case ReplayResult::kUnchanged:
                session_->undo.CompleteRedo(std::move(*operation));
                ui::ShowError(hwnd_, util::Tr(L"The change couldn't be undone."));
                break;
            case ReplayResult::kPartial:
                if (operation->type == changes::UndoOperation::Type::kGroup)
                {
                    session_->undo.CompletePartial(std::move(*operation), replayed, false, left_both_names);
                }
                break;
            }
            UpdateToolbarStates();
            return true;
        }
    case cmd::kEditRedo:
        {
            if (!EnsureWritable())
            {
                return true;
            }
            auto operation = session_->undo.TakeRedo();
            if (!operation)
            {
                return true;
            }
            size_t replayed = 0;
            bool left_both_names = false;
            switch (ApplyUndoOperation(*operation, true, &replayed, &left_both_names))
            {
            case ReplayResult::kSuccess:
                session_->undo.CompleteRedo(std::move(*operation));
                break;
            case ReplayResult::kUnchanged:
                session_->undo.CompleteUndo(std::move(*operation));
                ui::ShowError(hwnd_, util::Tr(L"The change couldn't be redone."));
                break;
            case ReplayResult::kPartial:
                if (operation->type == changes::UndoOperation::Type::kGroup)
                {
                    session_->undo.CompletePartial(std::move(*operation), replayed, true, left_both_names);
                }
                break;
            }
            UpdateToolbarStates();
            return true;
        }
    default:
        return false;
    }
}

bool MainWindow::Impl::HandleRegistryNavigationCommand(int command_id)
{
    switch (command_id)
    {
    case cmd::kRegistryLocal:
        OpenLocalRegistryTab();
        return true;
    case cmd::kRegistryLocal32:
        OpenLocalRegistryTab(win32::kAlternateRegistryView);
        return true;
    case cmd::kRegistryOtherView:
        GoToOtherView();
        return true;
    case cmd::kEditGoToVirtualStore:
    case cmd::kEditGoToGlobalKey:
        {
            const std::wstring target = command_id == cmd::kEditGoToVirtualStore ? VirtualStoreTarget() : GlobalKeyTarget();
            if (!target.empty())
            {
                NavigateToResolvedExternalJump(target, L"");
            }
            return true;
        }
    case cmd::kRegistryNetwork:
        SwitchToRemoteRegistry();
        return true;
    case cmd::kRegistryOffline:
        SwitchToOfflineRegistry();
        return true;
    case cmd::kNavBack:
        NavigateBack();
        return true;
    case cmd::kNavForward:
        NavigateForward();
        return true;
    case cmd::kNavUp:
        NavigateUp();
        return true;
    default:
        return false;
    }
}

void MainWindow::Impl::CopyKeyPathAs(int command_id, const RegistryNode& node, const std::wstring& path)
{
    std::wstring text;
    if (command_id == cmd::kEditCopyKeyPathNative || command_id == cmd::kEditCopyKeyPathNativeResolved)
    {
        // HKCR is merged and the 32 bit view redirects paths, only the opened key knows its location
        const bool unresolved = command_id == cmd::kEditCopyKeyPathNative;
        text = unresolved && !node.view ? registry_path::BuildNative(node) : std::wstring();
        KeyDetails details;
        if (text.empty() && RegistryStore::QueryKeyDetails(node, &details, unresolved))
        {
            text = std::move(details.native.native_name);
        }
        if (text.empty())
        {
            text = registry_path::BuildNative(node);
        }
        text = registry_path::DisplayName(text);
    }
    else if (!path.empty())
    {
        registry_path::Style style = registry_path::Style::kFull;
        switch (command_id)
        {
        case cmd::kEditCopyKeyPathAbbrev:
            style = registry_path::Style::kAbbreviated;
            break;
        case cmd::kEditCopyKeyPathRegEdit:
            style = registry_path::Style::kRegEditAddress;
            break;
        case cmd::kEditCopyKeyPathRegFile:
            style = registry_path::Style::kRegFileHeader;
            break;
        case cmd::kEditCopyKeyPathPowerShell:
            style = registry_path::Style::kPowerShellDrive;
            break;
        case cmd::kEditCopyKeyPathPowerShellProvider:
            style = registry_path::Style::kPowerShellProvider;
            break;
        case cmd::kEditCopyKeyPathEscaped:
            style = registry_path::Style::kEscaped;
            break;
        default:
            break;
        }
        text = FormatRegistryPath(path, style);
    }
    if (!text.empty())
    {
        ui::CopyText(hwnd_, text);
    }
}

}
