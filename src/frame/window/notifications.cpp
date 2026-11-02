// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "frame/window/shortcut_bindings.h"

#include "ui/dialog_layout.h"
#include "win32/translation.h"

#include <win32ui/dialog/edit.h>
#include <win32ui/dialog/tooltip.h>

namespace regkit
{
using namespace window_detail;

namespace
{

bool ValueNameExists(const RegistryNode& node, const std::wstring& name)
{
    RegistryValue entry;
    return RegistryStore::QueryValue(node, name, &entry);
}

bool KeyNameExists(const RegistryNode& parent, const std::wstring& name)
{
    KeyInfo info = {};
    return RegistryStore::QueryKeyInfo(ChildNode(parent, name), &info);
}

void ReportNameTaken(HWND owner, const wchar_t* message, const wchar_t* title, const std::wstring& name)
{
    ui::PromptKeyChoice(owner, message, registry_path::DisplayName(name), title, L"", L"", util::Tr(L"OK"));
}

void FormatCellFileTime(const FILETIME& filetime, wchar_t* buffer, int capacity)
{
    lstrcpynW(buffer, FormatFileTime(filetime).c_str(), capacity);
}

std::wstring HistoryActionText(const HistoryEntry& entry)
{
    return entry.backup_mode ? entry.action + L" (" + util::Tr(L"Backup/restore mode") + L")" : entry.action;
}

}

LRESULT MainWindow::Impl::HandleNotification(LPARAM lparam)
{
    auto* header = reinterpret_cast<NMHDR*>(lparam);
    if (!header)
    {
        return 0;
    }
    // toolbar Delete follows focused pane
    if (header->code == NM_SETFOCUS)
    {
        UpdateToolbarStates();
    }
    if (LRESULT toolbar_result = 0; win32ui::ToolbarNotify(header, &toolbar_result))
    {
        return toolbar_result;
    }
    if (header->code == TTN_SHOW)
    {
        return HandleTooltipNotification(header);
    }
    if (header->hwndFrom == tab_)
    {
        return HandleTabNotification(header, lparam);
    }
    if (header->hwndFrom == browse_.tree().hwnd() || header->hwndFrom == regedit_compat_tree_.hwnd())
    {
        return HandleTreeNotification(header, lparam);
    }
    if (header->hwndFrom == browse_.values().hwnd())
    {
        return HandleValueNotification(header, lparam);
    }
    if (header->hwndFrom == history_list_)
    {
        return HandleHistoryNotification(header, lparam);
    }
    if (header->hwndFrom == search_results_list_)
    {
        return HandleSearchNotification(header, lparam);
    }
    return 0;
}

std::wstring MainWindow::Impl::SearchCellFieldText(const search::Result& result, int subitem) const
{
    switch (subitem)
    {
    case 0:
        return result.key_path;
    case 1:
        return registry_path::DisplayName(search::DisplayName(result));
    case 2:
        return search::TypeText(result);
    case 3:
        return result.data_text;
    case 4:
        return search::IsKeyRow(result) || result.kind == search::ResultKind::kTraceValue
                   ? std::wstring()
                   : std::to_wstring(result.data_size);
    case 5:
        return FormatFileTime(result.modified);
    default:
        return std::wstring();
    }
}

std::wstring MainWindow::Impl::ListCellFieldText(HWND list, int item, int subitem)
{
    if (list == browse_.values().hwnd())
    {
        const ListRow* row = browse_.values().RowAt(item);
        return row ? ValueRowFieldText(*row, subitem) : std::wstring();
    }
    if (list == search_results_list_)
    {
        const int tab_index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
        const SearchTab* tab = tab_index >= 0 && static_cast<size_t>(tab_index) < search_tabs_.size()
                                   ? &search_tabs_[static_cast<size_t>(tab_index)]
                                   : nullptr;
        if (tab && tab->is_compare)
        {
            if (item < 0 || static_cast<size_t>(item) >= tab->compare_rows.size())
            {
                return std::wstring();
            }
            const search::compare::Row& row = tab->compare_rows[static_cast<size_t>(item)];
            switch (subitem)
            {
            case 0:
                return row.key_path;
            case 1:
                return row.is_key ? std::wstring(util::Tr(L"(Key)"))
                                  : (row.value_name.empty() ? std::wstring(util::Tr(L"(Default)")) : row.value_name);
            case 2:
                return row.first_text;
            case 3:
                return row.second_text;
            case 4:
                return row.matches ? std::wstring(util::Tr(L"Same")) : std::wstring(util::Tr(L"Different"));
            default:
                return std::wstring();
            }
        }
        if (!tab || item < 0 || static_cast<size_t>(item) >= tab->results.size())
        {
            return std::wstring();
        }
        const search::Result& result = tab->results[static_cast<size_t>(item)];
        return subitem == 6 ? search::SourceText(result, tab->sources) : SearchCellFieldText(result, subitem);
    }
    if (list == history_list_)
    {
        const auto& entries = change_history_.entries();
        if (item < 0 || static_cast<size_t>(item) >= entries.size())
        {
            return std::wstring();
        }
        const HistoryEntry& entry = entries[static_cast<size_t>(item)];
        switch (subitem)
        {
        case 0:
            return entry.time_text;
        case 1:
            return HistoryActionText(entry);
        case 2:
            return entry.old_data;
        case 3:
            return entry.new_data;
        default:
            return std::wstring();
        }
    }
    return std::wstring();
}

LRESULT MainWindow::Impl::HandleTooltipNotification(NMHDR* header)
{
    if (header->code == TTN_SHOW && header->hwndFrom == value_tooltip_)
    {
        return win32ui::PlaceTipNearCursor(value_tooltip_, win32ui::WindowDpi(hwnd_)) ? TRUE : 0;
    }
    return 0;
}

bool MainWindow::Impl::OpenSearchResultRow(int item, bool new_tab)
{
    const int index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
    if (item < 0 || index < 0 || static_cast<size_t>(item) >= SearchRowCount(index))
    {
        return false;
    }
    const std::wstring path = SearchRowKeyPath(index, item);
    const search::Result* row = SearchResultAt(item);
    const bool value_row = row && !search::IsKeyRow(*row);
    const std::wstring value_name = value_row ? row->value_name : std::wstring();
    const SearchTab& search_tab = search_tabs_[static_cast<size_t>(index)];
    const search::Source source = row && row->source < search_tab.sources.size() ? search_tab.sources[row->source] : search::Source{};
    if (source.kind != search::Source::Kind::kLocal || source.view)
    {
        OpenSourceEntry(source, path, value_name, new_tab || FindSourceTab(source) < 0);
        return true;
    }
    if (new_tab)
    {
        OpenLocalRegistryTab();
    }
    else
    {
        ActivateLocalRegistryTab();
    }
    ApplyViewVisibility();
    UpdateStatus();
    SelectTreePath(path);
    if (value_row)
    {
        SelectValueWhenReady(value_name);
    }
    return true;
}

bool MainWindow::Impl::OpenSelectedSearchResult(bool new_tab)
{
    if (!search_results_list_ || GetFocus() != search_results_list_ || IsCompareTabSelected())
    {
        return false;
    }
    return OpenSearchResultRow(ListView_GetNextItem(search_results_list_, -1, LVNI_SELECTED), new_tab);
}

LRESULT MainWindow::Impl::HandleTabNotification(NMHDR* header, LPARAM lparam)
{
    (void)lparam;
    if (header->hwndFrom == tab_ && header->code == TCN_SELCHANGING)
    {
        if (!suppress_tab_change_ && tab_)
        {
            int current = TabCtrl_GetCurSel(tab_);
            CaptureRegistryTabState(current, false);
        }
        return 0;
    }
    if (header->hwndFrom == tab_ && header->code == TCN_SELCHANGE)
    {
        if (suppress_tab_change_)
        {
            ApplyViewVisibility();
            UpdateSearchResultsView();
            UpdateStatus();
            return 0;
        }
        int sel = TabCtrl_GetCurSel(tab_);
        ApplyTabSelection(sel);
        ApplyViewVisibility();
        UpdateSearchResultsView();
        UpdateStatus();
        return 0;
    }
    return 0;
}

LRESULT MainWindow::Impl::HandleTreeNotification(NMHDR* header, LPARAM lparam)
{
    if (header->hwndFrom == regedit_compat_tree_.hwnd())
    {
        if (header->code == TVN_ITEMEXPANDINGW)
        {
            regedit_compat_tree_.OnItemExpanding(reinterpret_cast<NMTREEVIEWW*>(lparam));
            return 0;
        }
        if (header->code == TVN_GETDISPINFOW)
        {
            regedit_compat_tree_.OnGetDispInfo(reinterpret_cast<NMTVDISPINFOW*>(lparam));
            return 0;
        }
        if (header->code == TVN_SELCHANGEDW)
        {
            RegistryNode* node = regedit_compat_tree_.OnSelectionChanged(reinterpret_cast<NMTREEVIEWW*>(lparam));
            if (node)
            {
                QueueCompatJump(*node);
            }
            return 0;
        }
        return 0;
    }
    if (header->hwndFrom == browse_.tree().hwnd())
    {
        if (header->code == TVN_ITEMEXPANDINGW)
        {
            // Explorer tree style animates rows below expanding keys
            SuspendTreeRedraw();
            browse_.tree().OnItemExpanding(reinterpret_cast<NMTREEVIEWW*>(lparam));
            return 0;
        }
        if (header->code == TVN_GETDISPINFOW)
        {
            browse_.tree().OnGetDispInfo(reinterpret_cast<NMTVDISPINFOW*>(lparam));
            return 0;
        }
        if (header->code == TVN_ITEMEXPANDEDW)
        {
            if (!jump_ui_batch_active_)
            {
                MarkTreeStateDirty();
            }
            return 0;
        }
        if (header->code == TVN_BEGINLABELEDITW)
        {
            if (settings_.read_only)
            {
                return TRUE;
            }
            auto* disp = reinterpret_cast<NMTVDISPINFOW*>(lparam);
            if (!disp)
            {
                return TRUE;
            }
            RegistryNode* node = browse_.tree().NodeFromItem(disp->item.hItem);
            if (!node || node->subkey.empty())
            {
                return TRUE;
            }
            HWND edit = TreeView_GetEditControl(browse_.tree().hwnd());
            if (edit)
            {
                win32ui::ApplyThemeToWindow(edit);
                win32ui::AddSelectAllShortcut(edit);
            }
            return FALSE;
        }
        if (header->code == TVN_ENDLABELEDITW)
        {
            if (settings_.read_only)
            {
                return FALSE;
            }
            auto* disp = reinterpret_cast<NMTVDISPINFOW*>(lparam);
            if (!disp || !disp->item.pszText)
            {
                return FALSE;
            }
            RegistryNode* node = browse_.tree().NodeFromItem(disp->item.hItem);
            if (!node || node->subkey.empty())
            {
                return FALSE;
            }
            std::wstring new_name = registry_path::RawName(disp->item.pszText);
            std::wstring old_name = LeafName(*node);
            if (new_name.empty() || new_name == old_name)
            {
                return FALSE;
            }
            RegistryNode rename_parent = *node;
            rename_parent.subkey = registry_path::Parent(node->subkey);
            if (!EqualsInsensitive(new_name, old_name) && KeyNameExists(rename_parent, new_name))
            {
                ReportNameTaken(hwnd_, util::Tr(L"A key with this name already exists:"), util::Tr(L"Rename Key"), new_name);
                return FALSE;
            }
            if (!RegistryStore::RenameKey(*node, new_name))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to rename key."));
                return FALSE;
            }
            UpdateLeafName(node, new_name);
            if (browse_.current_node() && SameNode(*browse_.current_node(), *node))
            {
                UpdateAddressBar(browse_.current_node());
            }
            AppendHistoryEntry(L"Rename key", registry_path::DisplayName(old_name), registry_path::DisplayName(new_name));
            MarkOfflineDirty();
            changes::UndoOperation op;
            op.type = changes::UndoOperation::Type::kRenameKey;
            op.node = rename_parent;
            op.name = old_name;
            op.new_name = new_name;
            PushUndo(std::move(op));
            RefreshTreeSelection();
            RefreshMatchingTreeNodes(TreeView_GetParent(browse_.tree().hwnd(), disp->item.hItem));
            UpdateValueListForNode(browse_.current_node());
            return TRUE;
        }
        if (header->code == TVN_SELCHANGEDW)
        {
            auto* info = reinterpret_cast<NMTREEVIEWW*>(lparam);
            RegistryNode* previous_node = browse_.current_node();
            RegistryNode* node = browse_.tree().OnSelectionChanged(info);
            if (startup_tree_restore_pending_ && !applying_startup_tree_restore_)
            {
                startup_tree_restore_pending_ = false;
            }
            browse_.set_current_node(node);
            if (previous_node && node && SameNode(*previous_node, *node))
            {
                return 0;
            }
            if (!jump_ui_batch_active_)
            {
                ApplyTreeSelectionEffects(node);
            }
            return 0;
        }
        if (header->code == NM_CUSTOMDRAW)
        {
            // hovered item in dark trees
            LRESULT result = CDRF_DODEFAULT;
            win32ui::ThemeCustomDraw(header, &result);
            return result;
        }
    }
    return 0;
}

LRESULT MainWindow::Impl::HandleValueNotification(NMHDR* header, LPARAM lparam)
{
    if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_ODCACHEHINT)
    {
        auto* hint = reinterpret_cast<NMLVCACHEHINT*>(lparam);
        if (hint)
        {
            QueueValuePreviews(hint->iFrom, hint->iTo);
        }
        return 0;
    }
    if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_GETDISPINFOW)
    {
        auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
        ListRow* mutable_row = browse_.values().MutableRowAt(disp->item.iItem);
        const ListRow* row = mutable_row;
        if (!row)
        {
            if (disp->item.mask & LVIF_TEXT)
            {
                if (disp->item.pszText && disp->item.cchTextMax > 0)
                {
                    disp->item.pszText[0] = L'\0';
                }
            }
            if (disp->item.mask & LVIF_IMAGE)
            {
                disp->item.iImage = 0;
            }
            return 0;
        }
        if (disp->item.mask & LVIF_TEXT)
        {
            const int subitem = disp->item.iSubItem;
            if (subitem == kValueColData && mutable_row && !mutable_row->data_ready && !value_preview_request_posted_)
            {
                value_preview_request_posted_ = true;
                if (!PostMessageW(hwnd_, frame::message_id::kValuePreviewRequest, static_cast<WPARAM>(disp->item.iItem), 0))
                {
                    value_preview_request_posted_ = false;
                }
            }
            const std::wstring& text = ValueRowFieldText(*row, subitem);
            if (text.size() > kCellTextDrawLimit && disp->item.pszText && disp->item.cchTextMax > 0)
            {
                lstrcpynW(disp->item.pszText, text.c_str(), disp->item.cchTextMax);
            }
            else
            {
                disp->item.pszText = const_cast<wchar_t*>(text.c_str());
            }
        }
        if (disp->item.mask & LVIF_IMAGE)
        {
            disp->item.iImage = row->image_index;
        }
        return 0;
    }
    if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_BEGINLABELEDITW)
    {
        if (settings_.read_only)
        {
            return TRUE;
        }
        auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
        if (!disp)
        {
            return TRUE;
        }
        const ListRow* row = browse_.values().RowAt(disp->item.iItem);
        if (!row || row->extra.empty() || (row->kind != rowkind::kValue && row->kind != rowkind::kKey))
        {
            return TRUE;
        }
        HWND edit = ListView_GetEditControl(browse_.values().hwnd());
        if (edit)
        {
            win32ui::ApplyThemeToWindow(edit);
            win32ui::AddSelectAllShortcut(edit);
        }
        return FALSE;
    }
    if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_ENDLABELEDITW)
    {
        if (settings_.read_only)
        {
            return FALSE;
        }
        auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
        if (!disp || !disp->item.pszText || !browse_.current_node())
        {
            return FALSE;
        }
        const ListRow* row = browse_.values().RowAt(disp->item.iItem);
        if (!row || row->extra.empty())
        {
            return FALSE;
        }
        std::wstring new_name = registry_path::RawName(disp->item.pszText);
        std::wstring old_name = row->extra;
        if (new_name.empty() || new_name == old_name)
        {
            return FALSE;
        }
        const bool case_only = EqualsInsensitive(new_name, old_name);
        if (row->kind == rowkind::kKey)
        {
            RegistryNode child = ChildNode(*browse_.current_node(), old_name);
            if (!case_only && KeyNameExists(*browse_.current_node(), new_name))
            {
                ReportNameTaken(hwnd_, util::Tr(L"A key with this name already exists:"), util::Tr(L"Rename Key"), new_name);
                return FALSE;
            }
            if (!RegistryStore::RenameKey(child, new_name))
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to rename key."));
                return FALSE;
            }
            AppendHistoryEntry(L"Rename key " + registry_path::DisplayName(old_name), registry_path::DisplayName(old_name), registry_path::DisplayName(new_name));
            MarkOfflineDirty();
            changes::UndoOperation op;
            op.type = changes::UndoOperation::Type::kRenameKey;
            op.node = *browse_.current_node();
            op.name = old_name;
            op.new_name = new_name;
            PushUndo(std::move(op));
            RefreshTreeSelection();
            UpdateValueListForNode(browse_.current_node());
            return TRUE;
        }
        if (!case_only && ValueNameExists(*browse_.current_node(), new_name))
        {
            ReportNameTaken(hwnd_, util::Tr(L"A value with this name already exists:"), util::Tr(L"Rename Value"), new_name);
            return FALSE;
        }
        bool both_names_left = false;
        if (!RegistryStore::RenameValue(*browse_.current_node(), old_name, new_name, &both_names_left))
        {
            if (both_names_left)
            {
                // value rename copies then deletes, either step can fail
                MarkOfflineDirty();
                UpdateValueListForNode(browse_.current_node());
                ui::ShowError(hwnd_, util::Tr(L"The value was copied to the new name but the old name "
                                              L"couldn't be removed. Both names now exist."));
            }
            else
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to rename value."));
            }
            return FALSE;
        }
        AppendValueHistoryEntry(L"Rename value " + old_name, old_name, new_name, *browse_.current_node(), new_name, HistoryEntry::RevertKind::kNone);
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kRenameValue;
        op.node = *browse_.current_node();
        op.name = old_name;
        op.new_name = new_name;
        PushUndo(std::move(op));
        if (EqualsInsensitive(old_name, appended_value_name_))
        {
            ListRow* updated = browse_.values().MutableRowAt(disp->item.iItem);
            if (updated)
            {
                updated->name = new_name;
                updated->extra = new_name;
                browse_.values().InvalidateFilterCache(updated);
                browse_.values().RefreshFilter();
                ListView_RedrawItems(browse_.values().hwnd(), disp->item.iItem, disp->item.iItem);
                browse_.SelectValue(new_name);
            }
            appended_value_name_.clear();
            return TRUE;
        }
        UpdateValueListForNode(browse_.current_node());
        retained_value_name_ = new_name;
        retained_value_key_path_ = registry_path::Build(*browse_.current_node());
        return TRUE;
    }
    if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_ITEMCHANGED)
    {
        auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
        // coalesce selection changes into one status update
        if (!updating_value_list_ && info && ((info->uOldState ^ info->uNewState) & LVIS_SELECTED) != 0 && !status_update_pending_ &&
            PostMessageW(hwnd_, frame::message_id::kStatusUpdate, 0, 0))
        {
            status_update_pending_ = true;
        }
        return 0;
    }
    if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_COLUMNCLICK)
    {
        auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
        if (info)
        {
            SortValueList(info->iSubItem, true);
        }
        return 0;
    }
    if (header->hwndFrom == browse_.values().hwnd() && (header->code == NM_DBLCLK || header->code == LVN_ITEMACTIVATE))
    {
        auto* activate = reinterpret_cast<NMITEMACTIVATE*>(lparam);
        if (activate && activate->iItem >= 0 && browse_.current_node())
        {
            const ListRow* row = browse_.values().RowAt(activate->iItem);
            bool fast_activate = false;
            if (header->code == LVN_ITEMACTIVATE)
            {
                if (!value_activate_from_key_)
                {
                    return 0;
                }
                value_activate_from_key_ = false;
                fast_activate = true;
            }
            if (header->code == NM_DBLCLK)
            {
                fast_activate = true;
            }
            if (row && row->kind == rowkind::kKey)
            {
                if (fast_activate)
                {
                    std::wstring path = registry_path::Build(*browse_.current_node());
                    if (!row->extra.empty())
                    {
                        path.append(L"\\");
                        path.append(row->extra);
                    }
                    SelectTreePath(path);
                }
                return 0;
            }
            if (row && row->kind == rowkind::kValue)
            {
                if (activate->iSubItem == kValueColComment)
                {
                    HandleMenuCommand(cmd::kEditModifyComment);
                }
                else
                {
                    HandleMenuCommand(cmd::kEditModify);
                }
                return 0;
            }
        }
        return 0;
    }

    if (header->code == NM_CUSTOMDRAW)
    {
        auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);
        if (!draw)
        {
            return CDRF_DODEFAULT;
        }
        LRESULT result = CDRF_DODEFAULT;
        switch (draw->nmcd.dwDrawStage)
        {
        case CDDS_PREPAINT:
            result = CDRF_NOTIFYITEMDRAW;
            break;
        case CDDS_ITEMPREPAINT:
            draw->nmcd.uItemState &= ~CDIS_FOCUS;
            break;
        default:
            break;
        }
        return win32ui::DrawListGrid(browse_.values().hwnd(), draw, result);
    }
    return 0;
}

search::Result* MainWindow::Impl::SearchResultAt(int item)
{
    const int tab_index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
    if (item < 0 || tab_index < 0 || static_cast<size_t>(tab_index) >= search_tabs_.size())
    {
        return nullptr;
    }
    SearchTab& tab = search_tabs_[static_cast<size_t>(tab_index)];
    if (tab.is_compare)
    {
        return nullptr;
    }
    return static_cast<size_t>(item) < tab.results.size() ? &tab.results[static_cast<size_t>(item)] : nullptr;
}

std::wstring MainWindow::Impl::SearchRowKeyPath(int tab_index, int item) const
{
    if (item < 0 || tab_index < 0 || static_cast<size_t>(tab_index) >= search_tabs_.size())
    {
        return std::wstring();
    }
    const SearchTab& tab = search_tabs_[static_cast<size_t>(tab_index)];
    const size_t row = static_cast<size_t>(item);
    if (tab.is_compare)
    {
        return row < tab.compare_rows.size() ? tab.compare_rows[row].key_path : std::wstring();
    }
    return row < tab.results.size() ? tab.results[row].key_path : std::wstring();
}

size_t MainWindow::Impl::SearchRowCount(int tab_index) const
{
    if (tab_index < 0 || static_cast<size_t>(tab_index) >= search_tabs_.size())
    {
        return 0;
    }
    const SearchTab& tab = search_tabs_[static_cast<size_t>(tab_index)];
    return tab.is_compare ? tab.compare_rows.size() : tab.results.size();
}

LRESULT MainWindow::Impl::HandleSearchListCustomDraw(NMLVCUSTOMDRAW* draw)
{
    if (!draw || !search_results_list_)
    {
        return CDRF_DODEFAULT;
    }
    const int item = static_cast<int>(draw->nmcd.dwItemSpec);
    switch (draw->nmcd.dwDrawStage)
    {
    case CDDS_PREPAINT:
        return win32ui::DrawListGrid(search_results_list_, draw, CDRF_NOTIFYITEMDRAW);
    case CDDS_ITEMPREPAINT:
        {
            draw->nmcd.uItemState &= ~CDIS_FOCUS;
            LRESULT stage = CDRF_DODEFAULT;
            const search::Result* result = SearchResultAt(item);
            if (result && result->match_length > 0)
            {
                stage |= CDRF_NOTIFYSUBITEMDRAW;
            }
            return win32ui::DrawListGrid(search_results_list_, draw, stage);
        }
    case CDDS_ITEMPOSTPAINT:
        return win32ui::DrawListGrid(search_results_list_, draw, CDRF_DODEFAULT);
    case CDDS_POSTPAINT:
        return win32ui::DrawListGrid(search_results_list_, draw, CDRF_DODEFAULT);
    case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
        {
            search::Result* result = SearchResultAt(item);
            if (!result || SearchMatchSubItem(*result) != draw->iSubItem)
            {
                return CDRF_DODEFAULT;
            }
            return CDRF_NOTIFYPOSTPAINT;
        }
    case CDDS_ITEMPOSTPAINT | CDDS_SUBITEM:
        {
            const search::Result* result = SearchResultAt(item);
            const int subitem = draw->iSubItem;
            if (!result || SearchMatchSubItem(*result) != subitem)
            {
                return CDRF_DODEFAULT;
            }
            RECT cell = {};
            if (draw->iSubItem == 0)
            {
                RECT row = {};
                if (!ListView_GetItemRect(search_results_list_, item, &cell, LVIR_LABEL) ||
                    !ListView_GetItemRect(search_results_list_, item, &row, LVIR_BOUNDS))
                {
                    return CDRF_DODEFAULT;
                }
                cell.right = row.left + ListView_GetColumnWidth(search_results_list_, 0);
                cell.left += kLabelTextInset;
            }
            else
            {
                cell.top = draw->iSubItem;
                cell.left = LVIR_BOUNDS;
                if (!SendMessageW(search_results_list_, LVM_GETSUBITEMRECT, item, reinterpret_cast<LPARAM>(&cell)))
                {
                    return CDRF_DODEFAULT;
                }
                cell.left += kCellTextPadding;
            }
            cell.right -= kCellTextPadding;
            if (cell.left >= cell.right)
            {
                return CDRF_DODEFAULT;
            }
            std::wstring_view cell_text;
            std::wstring name_text;
            switch (subitem)
            {
            case 0:
                cell_text = result->key_path;
                break;
            case 1:
                name_text = registry_path::DisplayName(search::DisplayName(*result));
                cell_text = name_text;
                break;
            case 3:
                cell_text = result->data_text;
                break;
            default:
                return CDRF_DODEFAULT;
            }
            HFONT font = reinterpret_cast<HFONT>(SendMessageW(search_results_list_, WM_GETFONT, 0, 0));
            HFONT old_font = font ? reinterpret_cast<HFONT>(SelectObject(draw->nmcd.hdc, font)) : nullptr;
            // draw matched range after the list view paints the full cell
            win32ui::DrawCellTextMatch(draw->nmcd.hdc, cell, cell_text, static_cast<int>(result->match_start), static_cast<int>(result->match_length), win32ui::CurrentThemeColors().focus);
            if (old_font)
            {
                SelectObject(draw->nmcd.hdc, old_font);
            }
            return CDRF_DODEFAULT;
        }
    default:
        return CDRF_DODEFAULT;
    }
}

LRESULT MainWindow::Impl::HandleHistoryNotification(NMHDR* header, LPARAM lparam)
{
    if (header->hwndFrom == history_list_ && header->code == LVN_GETDISPINFOW)
    {
        auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
        const auto& entries = change_history_.entries();
        if (!disp || disp->item.iItem < 0 || static_cast<size_t>(disp->item.iItem) >= entries.size())
        {
            if (disp && (disp->item.mask & LVIF_TEXT) && disp->item.pszText && disp->item.cchTextMax > 0)
            {
                disp->item.pszText[0] = L'\0';
            }
            return 0;
        }
        if (disp->item.mask & LVIF_TEXT)
        {
            const auto& entry = entries[static_cast<size_t>(disp->item.iItem)];
            const std::wstring* text = &entry.time_text;
            std::wstring marked;
            switch (disp->item.iSubItem)
            {
            case 1:
                if (entry.backup_mode)
                {
                    marked = HistoryActionText(entry);
                    text = &marked;
                }
                else
                {
                    text = &entry.action;
                }
                break;
            case 2:
                text = &entry.old_data;
                break;
            case 3:
                text = &entry.new_data;
                break;
            default:
                break;
            }
            const bool has_buffer = disp->item.pszText && disp->item.cchTextMax > 0;
            if (has_buffer && (text == &marked || text->size() > kCellTextDrawLimit))
            {
                lstrcpynW(disp->item.pszText, text->c_str(), disp->item.cchTextMax);
            }
            else if (text != &marked)
            {
                disp->item.pszText = const_cast<wchar_t*>(text->c_str());
            }
        }
        return 0;
    }
    if (header->hwndFrom == history_list_ && header->code == LVN_COLUMNCLICK)
    {
        auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
        if (info)
        {
            SortHistoryList(info->iSubItem, true);
        }
        return 0;
    }
    if (header->code == NM_CUSTOMDRAW)
    {
        auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);
        if (!draw)
        {
            return CDRF_DODEFAULT;
        }
        LRESULT result = CDRF_DODEFAULT;
        if (draw->nmcd.dwDrawStage == CDDS_PREPAINT)
        {
            result = CDRF_NOTIFYITEMDRAW;
        }
        if (draw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
        {
            draw->nmcd.uItemState &= ~CDIS_FOCUS;
        }
        return win32ui::DrawListGrid(history_list_, draw, result);
    }
    return 0;
}

LRESULT MainWindow::Impl::HandleSearchNotification(NMHDR* header, LPARAM lparam)
{
    if (header->hwndFrom == search_results_list_ && header->code == LVN_ODCACHEHINT)
    {
        auto* hint = reinterpret_cast<NMLVCACHEHINT*>(lparam);
        if (hint)
        {
            QueueSearchPreviews(hint->iFrom, hint->iTo);
        }
        return 0;
    }
    if (header->hwndFrom == search_results_list_ && header->code == LVN_GETDISPINFOW)
    {
        auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
        const int index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
        SearchTab* tab = index >= 0 && static_cast<size_t>(index) < search_tabs_.size()
                             ? &search_tabs_[static_cast<size_t>(index)]
                             : nullptr;
        auto clear_cell = [&]() {
            if ((disp->item.mask & LVIF_TEXT) && disp->item.pszText && disp->item.cchTextMax > 0)
            {
                disp->item.pszText[0] = L'\0';
            }
            if (disp->item.mask & LVIF_IMAGE)
            {
                disp->item.iImage = 0;
            }
        };
        if (!tab || disp->item.iItem < 0)
        {
            clear_cell();
            return 0;
        }
        const size_t row_index = static_cast<size_t>(disp->item.iItem);
        const int subitem = disp->item.iSubItem;

        if (tab->is_compare)
        {
            if (row_index >= tab->compare_rows.size())
            {
                clear_cell();
                return 0;
            }
            const search::compare::Row& row = tab->compare_rows[row_index];
            if (disp->item.mask & LVIF_TEXT)
            {
                if (disp->item.pszText && disp->item.cchTextMax > 0)
                {
                    disp->item.pszText[0] = L'\0';
                }
                auto set_text = [&](const std::wstring& text) {
                    if (text.size() > kCellTextDrawLimit && disp->item.pszText && disp->item.cchTextMax > 0)
                    {
                        lstrcpynW(disp->item.pszText, text.c_str(), disp->item.cchTextMax);
                    }
                    else
                    {
                        disp->item.pszText = const_cast<wchar_t*>(text.c_str());
                    }
                };
                switch (subitem)
                {
                case 0:
                    set_text(row.key_path);
                    break;
                case 1:
                    if (row.is_key)
                    {
                        disp->item.pszText = const_cast<wchar_t*>(util::Tr(L"(Key)"));
                    }
                    else if (row.value_name.empty())
                    {
                        disp->item.pszText = const_cast<wchar_t*>(util::Tr(L"(Default)"));
                    }
                    else
                    {
                        set_text(row.value_name);
                    }
                    break;
                case 2:
                    set_text(row.first_text);
                    break;
                case 3:
                    set_text(row.second_text);
                    break;
                case 4:
                    disp->item.pszText = const_cast<wchar_t*>(row.matches ? util::Tr(L"Same") : util::Tr(L"Different"));
                    break;
                default:
                    break;
                }
            }
            if (disp->item.mask & LVIF_IMAGE)
            {
                disp->item.iImage = row.is_key ? kFolderIconIndex : kValueIconIndex;
            }
            return 0;
        }

        if (row_index >= tab->results.size())
        {
            clear_cell();
            return 0;
        }
        search::Result& result = tab->results[row_index];

        if (subitem == 3 && result.data_state == search::DataState::kNotLoaded && !search_preview_request_posted_)
        {
            search_preview_request_posted_ = true;
            if (!PostMessageW(hwnd_, frame::message_id::kSearchPreviewRequest, 0, 0))
            {
                search_preview_request_posted_ = false;
            }
        }
        if (disp->item.mask & LVIF_TEXT)
        {
            wchar_t* buffer = disp->item.pszText;
            const int capacity = disp->item.cchTextMax;
            if (buffer && capacity > 0)
            {
                buffer[0] = L'\0';
            }
            auto set_text = [&](const std::wstring& text) {
                if (text.size() > kCellTextDrawLimit && buffer && capacity > 0)
                {
                    lstrcpynW(buffer, text.c_str(), capacity);
                }
                else
                {
                    disp->item.pszText = const_cast<wchar_t*>(text.c_str());
                }
            };
            switch (subitem)
            {
            case 0:
                set_text(result.key_path);
                break;
            case 1:
                if (search::IsKeyRow(result))
                {
                    disp->item.pszText = const_cast<wchar_t*>(L"");
                }
                else if (result.value_name.empty())
                {
                    disp->item.pszText = const_cast<wchar_t*>(util::Tr(L"(Default)"));
                }
                else
                {
                    disp->item.pszText = const_cast<wchar_t*>(result.value_name.c_str());
                }
                break;
            case 2:
                if (search::IsKeyRow(result))
                {
                    disp->item.pszText = const_cast<wchar_t*>(util::Tr(L"Key"));
                }
                else if (result.kind == search::ResultKind::kTraceValue)
                {
                    disp->item.pszText = const_cast<wchar_t*>(L"TRACE");
                }
                else if (buffer && capacity > 0)
                {
                    lstrcpynW(buffer, value_format::TypeName(result.type).c_str(), capacity);
                }
                break;
            case 3:
                set_text(result.data_text);
                break;
            case 4:
                if (buffer && capacity > 0 && !search::IsKeyRow(result) &&
                    result.kind != search::ResultKind::kTraceValue)
                {
                    swprintf_s(buffer, static_cast<size_t>(capacity), L"%lu", static_cast<unsigned long>(result.data_size));
                }
                break;
            case 5:
                if (buffer && capacity > 0)
                {
                    FormatCellFileTime(result.modified, buffer, capacity);
                }
                break;
            case 6:
                if (buffer && capacity > 0 && result.source < tab->sources.size())
                {
                    lstrcpynW(buffer, search::SourceText(result, tab->sources).c_str(), capacity);
                }
                break;
            default:
                break;
            }
        }
        if (disp->item.mask & LVIF_IMAGE)
        {
            if (search::IsKeyRow(result))
            {
                disp->item.iImage = kFolderIconIndex;
            }
            else if (result.kind == search::ResultKind::kTraceValue)
            {
                disp->item.iImage = kTraceIconIndex;
            }
            else if (UseBinaryValueIcon(result.type))
            {
                disp->item.iImage = kBinaryIconIndex;
            }
            else
            {
                disp->item.iImage = kValueIconIndex;
            }
        }
        return 0;
    }
    if (header->hwndFrom == search_results_list_ && header->code == LVN_COLUMNCLICK)
    {
        auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
        if (info)
        {
            SortSearchResults(info->iSubItem, true);
        }
        return 0;
    }
    if (header->hwndFrom == search_results_list_ && (header->code == NM_DBLCLK || header->code == LVN_ITEMACTIVATE))
    {
        if (IsCompareTabSelected())
        {
            return 0;
        }
        auto* activate = reinterpret_cast<NMITEMACTIVATE*>(lparam);
        if (activate && activate->iItem >= 0)
        {
            OpenSearchResultRow(activate->iItem, SearchResultOpensInNewTab());
        }
        return 0;
    }
    if (header->code == NM_CUSTOMDRAW)
    {
        return HandleSearchListCustomDraw(reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
    }
    return 0;
}

}
