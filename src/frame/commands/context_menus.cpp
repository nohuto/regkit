// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/dialog/message_dialog.h>
#include <win32ui/system/shell_open.h>

namespace regkit
{
using namespace command_detail;

namespace
{

std::vector<std::wstring> FindLinks(std::wstring_view text)
{
    std::vector<std::wstring> links;
    for (size_t start = 0; start < text.size() && links.size() <= cmd::kOpenLinkMax - cmd::kOpenLinkBase; ++start)
    {
        const size_t scheme = util::StartsWithInsensitive(text.substr(start), L"https://")  ? 8
                              : util::StartsWithInsensitive(text.substr(start), L"http://") ? 7
                                                                                            : 0;
        if (!scheme || (start > 0 && iswalnum(text[start - 1])))
        {
            continue;
        }
        size_t end = start + scheme;
        while (end < text.size() && text[end] > L' ' && !wcschr(L"\"'<>`", text[end]))
        {
            ++end;
        }
        std::wstring link(text.substr(start, end - start));
        while (link.size() > scheme && wcschr(L".,;:!?)]}", link.back()) &&
               (link.back() != L')' ||
                std::count(link.begin(), link.end(), L'(') < std::count(link.begin(), link.end(), L')')))
        {
            link.pop_back();
        }
        if (link.size() > scheme && std::find(links.begin(), links.end(), link) == links.end())
        {
            links.push_back(std::move(link));
        }
        start = end - 1;
    }
    return links;
}

struct KeyMenu
{
    bool has_node = true;
    bool simulated = false;
    bool can_modify = true;
    bool can_comment = true;
    bool can_rename = false;
    bool source_hive = false;
    bool hive_file = false;
    bool can_toggle = false;
    bool has_children = false;
    std::wstring expand_label;
};

void AppendKeyMenu(HMENU menu, const KeyMenu& key)
{
    const auto flags = [](bool enabled) { return MF_STRING | (enabled ? 0u : MF_GRAYED); };
    const bool writable = key.has_node && key.can_modify;
    frame::AppendCommand(menu, flags(key.has_node), cmd::kEditCopyKey, util::Tr(L"&Copy Key Name"));
    frame::AppendCommand(menu, flags(key.has_node), cmd::kEditCopyKeyPath, util::Tr(L"Copy Key &Path"));
    AppendMenuW(menu, MF_POPUP | (key.has_node ? 0u : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), util::Tr(L"Copy Key Path &As"));
    frame::AppendCommand(menu, flags(key.can_comment), cmd::kEditModifyComment, util::Tr(L"Modif&y Comment..."));
    if (key.source_hive)
    {
        frame::AppendCommand(menu, MF_STRING, cmd::kEditOpenSourceHive, util::Tr(L"&Open in Source Hive"));
    }
    if (key.hive_file)
    {
        frame::AppendCommand(menu, MF_STRING, cmd::kOptionsHiveFileDir, util::Tr(L"On-Disk &Hive File"));
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    frame::AppendCommand(menu, flags(key.can_toggle), cmd::kTreeToggleExpand, key.expand_label.c_str());
    frame::AppendCommand(menu, flags(key.has_children), cmd::kTreeExpandAll, util::Tr(L"Expand All &Subkeys"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (key.simulated)
    {
        frame::AppendCommand(menu, flags(writable), cmd::kCreateSimulatedKey, util::Tr(L"Crea&te Key"));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        frame::AppendCommand(menu, MF_STRING, cmd::kViewRefresh, util::Tr(L"&Refresh"));
        return;
    }
    AppendMenuW(menu, MF_POPUP | (writable ? 0u : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildNewMenu()), util::Tr(L"Ne&w"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    frame::AppendCommand(menu, flags(key.has_node), cmd::kFileExport, util::Tr(L"&Export..."));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    frame::AppendCommand(menu, MF_STRING, cmd::kViewRefresh, util::Tr(L"&Refresh"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    frame::AppendCommand(menu, flags(key.can_rename && key.can_modify), cmd::kEditRename, util::Tr(L"Re&name"));
    frame::AppendCommand(menu, flags(key.can_rename && key.can_modify), cmd::kEditDelete, util::Tr(L"&Delete"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    frame::AppendCommand(menu, flags(writable), cmd::kEditPermissions, util::Tr(L"Perm&issions..."));
    frame::AppendCommand(menu, MF_STRING, cmd::kEditKeyInfo, util::Tr(L"&Key Information..."));
}

}

void MainWindow::Impl::RecordNavigation(const std::wstring& path)
{
    if (browse_.RecordNavigation(path))
    {
        UpdateToolbarStates();
    }
}

void MainWindow::Impl::NavigateBack()
{
    if (const auto target = browse_.Back())
    {
        if (!SelectTreePath(*target))
        {
            browse_.UndoNavigation(1);
        }
        UpdateToolbarStates();
    }
}

void MainWindow::Impl::NavigateForward()
{
    if (const auto target = browse_.Forward())
    {
        if (!SelectTreePath(*target))
        {
            browse_.UndoNavigation(-1);
        }
        UpdateToolbarStates();
    }
}

void MainWindow::Impl::NavigateUp()
{
    if (const auto target = browse_.Up())
    {
        if (!SelectTreePath(*target))
        {
            browse_.UndoNavigation(0);
        }
        UpdateToolbarStates();
    }
}

void MainWindow::Impl::ShowTreeContextMenu(HTREEITEM item, POINT screen_pt)
{
    if (!browse_.tree().hwnd())
    {
        return;
    }
    if (item)
    {
        TreeView_SelectItem(browse_.tree().hwnd(), item);
    }
    SetFocus(browse_.tree().hwnd());
    HTREEITEM target = item ? item : TreeView_GetSelection(browse_.tree().hwnd());
    RegistryNode* node = browse_.tree().NodeFromItem(target);

    KeyMenu key;
    key.has_node = node != nullptr;
    key.simulated = key.has_node && node->simulated;
    key.can_modify = !settings_.read_only;
    key.can_comment = key.has_node;
    key.can_rename = key.has_node && !node->subkey.empty();
    key.source_hive = key.has_node && !key.simulated && node->root == HKEY_CLASSES_ROOT && !node->subkey.empty();
    key.hive_file = key.has_node && !key.simulated && CanOpenHiveFile(*node);
    bool expanded = false;
    if (target)
    {
        TVITEMW tvi = {};
        tvi.hItem = target;
        tvi.mask = TVIF_STATE | TVIF_CHILDREN;
        tvi.stateMask = TVIS_EXPANDED;
        if (TreeView_GetItem(browse_.tree().hwnd(), &tvi))
        {
            expanded = (tvi.state & TVIS_EXPANDED) != 0;
            key.has_children = TreeView_GetChild(browse_.tree().hwnd(), target) != nullptr || tvi.cChildren != 0;
            key.can_toggle = expanded || key.has_children;
        }
    }
    key.expand_label = expanded ? util::Tr(L"Co&llapse Key") : util::Tr(L"E&xpand Key");
    HMENU menu = CreatePopupMenu();
    AppendKeyMenu(menu, key);
    ApplySimpleMode(menu);

    const int command = win32ui::TrackContextMenu(menu, screen_pt, hwnd_);

    if (command != 0)
    {
        HandleMenuCommand(command);
    }
}

bool MainWindow::Impl::CanOpenHiveFile(const RegistryNode& node)
{
    bool is_root = false;
    const bool is_current_user_root = node.subkey.empty() && (node.root == HKEY_CURRENT_USER || util::EqualsInsensitive(node.root_name, L"HKEY_CURRENT_USER"));
    return !LookupHivePath(node, &is_root).empty() && is_root && !is_current_user_root;
}

void MainWindow::Impl::ShowValueContextMenu(int index, POINT screen_pt)
{
    if (!browse_.values().hwnd())
    {
        return;
    }
    const ListRow* row = nullptr;
    if (index >= 0)
    {
        row = browse_.values().RowAt(index);
    }
    SetFocus(browse_.values().hwnd());

    HMENU menu = CreatePopupMenu();
    std::vector<std::wstring> links;
    if (row && row->kind == rowkind::kKey)
    {
        KeyMenu key;
        key.simulated = row->simulated;
        key.can_modify = !settings_.read_only;
        key.can_comment = !row->extra.empty();
        key.can_rename = !row->extra.empty();
        key.expand_label = util::Tr(L"E&xpand Key");
        if (browse_.current_node() && !key.simulated)
        {
            key.source_hive = browse_.current_node()->root == HKEY_CLASSES_ROOT;
            key.hive_file = CanOpenHiveFile(row->extra.empty() ? *browse_.current_node() : ChildNode(*browse_.current_node(), row->extra));
        }
        AppendKeyMenu(menu, key);
    }
    else if (row && row->kind == rowkind::kValue)
    {
        std::vector<ListRow> selected_rows = SelectedListRows(browse_.values());
        const bool all_values = !selected_rows.empty() &&
                                std::all_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& selected) {
                                    return selected.kind == rowkind::kValue && !selected.simulated;
                                });
        const bool single_value = all_values && selected_rows.size() == 1;
        const bool has_data = std::any_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& selected) {
            return !selected.missing;
        });
        bool can_rename = !settings_.read_only && single_value && has_data && !selected_rows.front().extra.empty();
        bool can_delete = !settings_.read_only && all_values && has_data;
        bool can_comment = !CommentTargets().empty();
        bool can_export = !row->simulated && browse_.current_node() && !browse_.current_node()->simulated && has_data;
        UINT rename_flags = MF_STRING | (can_rename ? 0 : MF_GRAYED);
        UINT inspect_flags = MF_STRING | (single_value ? 0 : MF_GRAYED);
        UINT delete_flags = MF_STRING | (can_delete ? 0 : MF_GRAYED);
        UINT single_flags = MF_STRING | (single_value ? 0 : MF_GRAYED);
        UINT data_flags = MF_STRING | (single_value && has_data ? 0 : MF_GRAYED);
        UINT export_flags = MF_STRING | (can_export ? 0 : MF_GRAYED);
        UINT comment_flags = MF_STRING | (can_comment ? 0 : MF_GRAYED);
        win32ui::AppendMenuItem(menu, inspect_flags, cmd::kEditModify, util::Tr(L"&Modify..."), L"Enter");
        frame::AppendCommand(menu, inspect_flags, cmd::kEditModifyBinary, util::Tr(L"Modify &Binary Data..."));
        frame::AppendCommand(menu, data_flags, cmd::kEditChangeType, util::Tr(L"C&hange Data Type..."));
        frame::AppendCommand(menu, data_flags, cmd::kEditBits, util::Tr(L"Edit Bit&s..."));
        frame::AppendCommand(menu, data_flags, cmd::kEditDecodeValue, util::Tr(L"Decode &Value..."));
        frame::AppendCommand(menu, comment_flags, cmd::kEditModifyComment, util::Tr(L"Modif&y Comment..."));
        frame::AppendCommand(menu, MF_STRING, cmd::kEditResetDefault, util::Tr(L"Reset &to Default"));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        frame::AppendCommand(menu, single_flags, cmd::kEditCopyValueName, util::Tr(L"Co&py Value Name"));
        frame::AppendCommand(menu, data_flags, cmd::kEditCopyValueData, util::Tr(L"&Copy Value Data"));
        if (single_value && browse_.current_node())
        {
            std::wstring text = changes::ResolveComment(value_comments_, default_comments_, {CommentKeyPath(*browse_.current_node()), row->extra, row->value_type, row->value_data_size})
                                    .text;
            RegistryValue entry;
            if ((row->value_type == REG_SZ || row->value_type == REG_EXPAND_SZ || row->value_type == REG_MULTI_SZ) &&
                GetValueEntry(*browse_.current_node(), row->extra, &entry))
            {
                text = value_format::DisplayData(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size())) +
                       L'\n' + text;
            }
            links = FindLinks(text);
        }
        if (links.size() == 1)
        {
            AppendMenuW(menu, MF_STRING, cmd::kOpenLinkBase, util::Tr(L"&Open Link"));
        }
        else if (!links.empty())
        {
            HMENU link_menu = CreatePopupMenu();
            for (size_t link = 0; link < links.size(); ++link)
            {
                AppendMenuW(link_menu, MF_STRING, cmd::kOpenLinkBase + link, win32ui::MenuName(links[link], 96).c_str());
            }
            AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(link_menu), util::Tr(L"&Open Link"));
        }
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        frame::AppendCommand(menu, export_flags, cmd::kFileExport, util::Tr(L"&Export..."));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        frame::AppendCommand(menu, rename_flags, cmd::kEditRename, util::Tr(L"Re&name"));
        frame::AppendCommand(menu, delete_flags, cmd::kEditDelete, util::Tr(L"&Delete"));
    }
    else
    {
        bool is_simulated = browse_.current_node() && browse_.current_node()->simulated;
        bool can_modify = !settings_.read_only;
        UINT edit_flags = MF_STRING | (browse_.current_node() ? 0 : MF_GRAYED);
        UINT modify_flags = MF_STRING | ((browse_.current_node() && can_modify) ? 0 : MF_GRAYED);
        frame::AppendCommand(menu, edit_flags, cmd::kEditCopyKey, util::Tr(L"&Copy Key Name"));
        frame::AppendCommand(menu, edit_flags, cmd::kEditCopyKeyPath, util::Tr(L"Copy Key &Path"));
        AppendMenuW(menu, MF_POPUP | (browse_.current_node() ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), util::Tr(L"Copy Key Path &As"));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        if (is_simulated)
        {
            frame::AppendCommand(menu, modify_flags, cmd::kCreateSimulatedKey, util::Tr(L"Crea&te Key"));
        }
        else
        {
            AppendMenuW(menu, MF_POPUP | ((browse_.current_node() && can_modify) ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildNewMenu()), util::Tr(L"Ne&w"));
        }
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        if (!is_simulated)
        {
            frame::AppendCommand(menu, edit_flags, cmd::kFileExport, util::Tr(L"&Export..."));
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        }
        frame::AppendCommand(menu, MF_STRING, cmd::kViewRefresh, util::Tr(L"&Refresh"));
        if (!is_simulated)
        {
            AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
            frame::AppendCommand(menu, modify_flags, cmd::kEditPermissions, util::Tr(L"Perm&issions..."));
            frame::AppendCommand(menu, MF_STRING, cmd::kEditKeyInfo, util::Tr(L"&Key Information..."));
        }
    }
    ApplySimpleMode(menu);

    const int command = win32ui::TrackContextMenu(menu, screen_pt, hwnd_);

    if (command >= cmd::kOpenLinkBase && command < cmd::kOpenLinkBase + static_cast<int>(links.size()))
    {
        ui::ReportFileDialogResult(hwnd_, win32ui::ShellOpen(hwnd_, links[command - cmd::kOpenLinkBase].c_str()));
    }
    else if (command != 0)
    {
        HandleMenuCommand(command);
    }
}

void MainWindow::Impl::ShowHistoryContextMenu(int index, POINT screen_pt)
{
    if (!history_list_)
    {
        return;
    }
    // copy entries because the menu loop allows appends and sorting
    std::optional<HistoryEntry> entry;
    if (index >= 0 && static_cast<size_t>(index) < change_history_.entries().size())
    {
        entry = change_history_.entries()[static_cast<size_t>(index)];
    }

    HMENU menu = CreatePopupMenu();
    if (entry)
    {
        HistoryEntry prepared_revert;
        AppendMenuW(menu, MF_STRING | (entry->key_path.empty() ? MF_GRAYED : 0), cmd::kHistoryOpenTarget, util::Tr(L"&Open Entry"));
        AppendMenuW(menu, MF_STRING | (PrepareHistoryRevert(*entry, &prepared_revert) ? 0 : MF_GRAYED), cmd::kHistoryRevert, util::Tr(L"R&evert"));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, cmd::kEditCopyKey, util::Tr(L"&Copy"));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, cmd::kHistoryRemove, util::Tr(L"Remove &from History"));
    }
    else
    {
        AppendMenuW(menu, MF_STRING, cmd::kViewRefresh, util::Tr(L"&Refresh"));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING | (change_history_.entries().empty() ? MF_GRAYED : 0), cmd::kEditDelete, util::Tr(L"Clear &History"));

    const int command = win32ui::TrackContextMenu(menu, screen_pt, hwnd_);

    if (command == cmd::kHistoryOpenTarget && entry)
    {
        if (!OpenHistoryTarget(*entry))
        {
            ui::ShowError(hwnd_, util::Tr(L"Failed to open history target."));
        }
    }
    else if (command == cmd::kHistoryRevert && entry)
    {
        RevertHistoryEntry(*entry);
    }
    else if (command == cmd::kEditCopyKey && entry)
    {
        const std::wstring combined =
            entry->time_text + L" | " + entry->action + L" | " + entry->old_data + L" | " + entry->new_data;
        win32ui::CopyTextToClipboard(hwnd_, combined);
    }
    else if (command == cmd::kHistoryRemove)
    {
        RemoveSelectedHistoryItems();
    }
    else if (command == cmd::kViewRefresh)
    {
        RefreshHistory();
    }
    else if (command == cmd::kEditDelete)
    {
        ClearHistoryItems(true);
    }
}

void MainWindow::Impl::ShowSearchResultContextMenu(int index, POINT screen_pt)
{
    if (!search_results_list_ || index < 0)
    {
        return;
    }

    int sel_tab = TabCtrl_GetCurSel(tab_);
    int search_index = SearchIndexFromTab(sel_tab);
    if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size())
    {
        return;
    }
    if (static_cast<size_t>(index) >= SearchRowCount(search_index))
    {
        return;
    }

    std::wstring key_path = SearchRowKeyPath(search_index, index);
    std::wstring first_key_path;
    std::wstring second_key_path;
    std::wstring row_value_name;
    search::Source first_source;
    search::Source second_source;
    const bool compare_row = IsCompareTabSelected();
    if (compare_row)
    {
        const SearchTab& compare_tab = search_tabs_[static_cast<size_t>(search_index)];
        if (!compare_tab.sources.empty())
        {
            first_source = compare_tab.sources.front();
        }
        if (compare_tab.sources.size() > 1)
        {
            second_source = compare_tab.sources[1];
        }
        if (static_cast<size_t>(index) < compare_tab.compare_rows.size())
        {
            const auto& row = compare_tab.compare_rows[static_cast<size_t>(index)];
            first_key_path = row.first_key_path;
            second_key_path = row.second_key_path;
            if (!row.is_key)
            {
                row_value_name = row.value_name;
            }
        }
    }

    const search::Result* result = SearchResultAt(index);
    search::Source row_source;
    if (!compare_row && result)
    {
        const SearchTab& result_tab = search_tabs_[static_cast<size_t>(search_index)];
        if (result->source < result_tab.sources.size())
        {
            row_source = result_tab.sources[result->source];
        }
    }
    const bool is_key_row = !result || search::IsKeyRow(*result);
    if (!is_key_row && result)
    {
        row_value_name = result->value_name;
    }
    if (key_path.empty())
    {
        return;
    }

    RegistryNode node;
    bool node_ok = ResolvePathToNode(key_path, &node);
    KeyInfo info = {};
    bool key_exists = node_ok && RegistryStore::QueryKeyInfo(node, &info);
    bool can_modify = !settings_.read_only;
    bool can_rename = key_exists && can_modify && (is_key_row ? !node.subkey.empty() : !row_value_name.empty());
    bool can_delete = key_exists && can_modify && (!is_key_row || !node.subkey.empty());
    bool can_export = key_exists;
    bool can_permissions = key_exists && can_modify;
    bool can_open_hive = key_exists && CanOpenHiveFile(node);

    enum
    {
        kSearchOpenKey = 51000,
        kSearchOpenKeyNewTab = 51001,
        kSearchOpenSecondKey = 51019,
        kSearchOpenSecondKeyNewTab = 51020,
        kSearchModify = 51002,
        kSearchModifyBinary = 51003,
        kSearchModifyComment = 51004,
        kSearchDecodeValue = 51021,
        kSearchCopyKeyName = 51005,
        kSearchCopyKeyPath = 51006,
        kSearchPermissions = 51007,
        kSearchOpenHive = 51008,
        kSearchExport = 51009,
        kSearchRename = 51010,
        kSearchDelete = 51011,
    };

    HMENU menu = CreatePopupMenu();
    if (compare_row)
    {
        const UINT first_flags = MF_STRING | (first_key_path.empty() ? MF_GRAYED : 0);
        const UINT second_flags = MF_STRING | (second_key_path.empty() ? MF_GRAYED : 0);
        const UINT first_open_flags = first_flags | (FindSourceTab(first_source) < 0 ? MF_GRAYED : 0);
        const UINT second_open_flags = second_flags | (FindSourceTab(second_source) < 0 ? MF_GRAYED : 0);
        AppendMenuW(menu, first_open_flags, kSearchOpenKey, util::Tr(L"Open &First Entry"));
        AppendMenuW(menu, first_flags, kSearchOpenKeyNewTab, util::Tr(L"Open Fi&rst Entry in New Tab"));
        AppendMenuW(menu, second_open_flags, kSearchOpenSecondKey, util::Tr(L"&Open Second Entry"));
        AppendMenuW(menu, second_flags, kSearchOpenSecondKeyNewTab, util::Tr(L"Open &Second Entry in New Tab"));
    }
    else
    {
        const UINT open_flags = MF_STRING | (FindSourceTab(row_source) < 0 ? MF_GRAYED : 0);
        win32ui::AppendMenuItem(menu, open_flags, kSearchOpenKey, util::Tr(L"Open &Key"), SearchResultOpensInNewTab() ? nullptr : L"Enter");
        win32ui::AppendMenuItem(menu, MF_STRING, kSearchOpenKeyNewTab, util::Tr(L"Open Key in New &Tab"), L"Ctrl+Shift+G");
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (!is_key_row)
    {
        AppendMenuW(menu, MF_STRING, kSearchModify, util::Tr(L"&Modify..."));
        AppendMenuW(menu, MF_STRING, kSearchModifyBinary, util::Tr(L"Modify &Binary Data..."));
        const bool trace_row = result && result->kind == search::ResultKind::kTraceValue;
        AppendMenuW(menu, MF_STRING | (trace_row ? MF_GRAYED : 0), kSearchDecodeValue, util::Tr(L"Decode &Value..."));
        AppendMenuW(menu, MF_STRING, kSearchModifyComment, util::Tr(L"Modif&y Comment..."));
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING, kSearchCopyKeyName, util::Tr(L"&Copy Key Name"));
    AppendMenuW(menu, MF_STRING, kSearchCopyKeyPath, util::Tr(L"Copy Key &Path"));
    AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), util::Tr(L"Copy Key Path &As"));
    UINT open_hive_flags = MF_STRING | (can_open_hive ? 0 : MF_GRAYED);
    AppendMenuW(menu, open_hive_flags, kSearchOpenHive, util::Tr(L"On-Disk &Hive File"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    UINT export_flags = MF_STRING | (can_export ? 0 : MF_GRAYED);
    AppendMenuW(menu, export_flags, kSearchExport, util::Tr(L"&Export..."));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    UINT rename_flags = MF_STRING | (can_rename ? 0 : MF_GRAYED);
    UINT delete_flags = MF_STRING | (can_delete ? 0 : MF_GRAYED);
    AppendMenuW(menu, rename_flags, kSearchRename, util::Tr(L"Re&name"));
    AppendMenuW(menu, delete_flags, kSearchDelete, util::Tr(L"&Delete"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    UINT permissions_flags = MF_STRING | (can_permissions ? 0 : MF_GRAYED);
    AppendMenuW(menu, permissions_flags, kSearchPermissions, util::Tr(L"Perm&issions..."));

    const int command = win32ui::TrackContextMenu(menu, screen_pt, hwnd_);
    if (command == 0)
    {
        return;
    }

    auto open_key = [&](bool new_tab) {
        if (compare_row)
        {
            OpenSourceEntry(first_source, first_key_path, row_value_name, new_tab);
            return;
        }
        OpenSourceEntry(row_source, key_path, row_value_name, new_tab);
    };
    auto focus_key = [&]() {
        open_key(false);
        if (browse_.tree().hwnd())
        {
            SetFocus(browse_.tree().hwnd());
        }
    };
    auto run_on_value = [&](int command) {
        if (is_key_row)
        {
            return;
        }
        open_key(false);
        if (SelectValueByName(result->value_name))
        {
            if (browse_.values().hwnd())
            {
                SetFocus(browse_.values().hwnd());
            }
            HandleMenuCommand(command);
            return;
        }
        pending_external_value_key_path_ = key_path;
        pending_external_value_name_ = result->value_name;
        pending_value_command_ = command;
    };
    auto run_on_key = [&](int command) {
        focus_key();
        HandleMenuCommand(command);
    };
    auto run_on_target = [&](int command) {
        if (is_key_row)
        {
            run_on_key(command);
        }
        else
        {
            run_on_value(command);
        }
    };

    switch (command)
    {
    case kSearchOpenKey:
        open_key(SearchResultOpensInNewTab());
        return;
    case kSearchOpenKeyNewTab:
        open_key(true);
        return;
    case kSearchOpenSecondKey:
        OpenSourceEntry(second_source, second_key_path, row_value_name, SearchResultOpensInNewTab());
        return;
    case kSearchOpenSecondKeyNewTab:
        OpenSourceEntry(second_source, second_key_path, row_value_name, true);
        return;
    case kSearchModify:
        run_on_value(cmd::kEditModify);
        return;
    case kSearchModifyBinary:
        run_on_value(cmd::kEditModifyBinary);
        return;
    case kSearchModifyComment:
        run_on_value(cmd::kEditModifyComment);
        return;
    case kSearchDecodeValue:
        run_on_value(cmd::kEditDecodeValue);
        return;
    case kSearchCopyKeyName:
        {
            std::wstring name;
            if (node_ok)
            {
                name = LeafName(node);
            }
            else
            {
                size_t pos = key_path.find_last_of(L"\\/");
                name = (pos == std::wstring::npos) ? key_path : key_path.substr(pos + 1);
            }
            if (!name.empty())
            {
                win32ui::CopyTextToClipboard(hwnd_, name);
            }
            return;
        }
    case kSearchCopyKeyPath:
        if (!key_path.empty())
        {
            win32ui::CopyTextToClipboard(hwnd_, key_path);
        }
        return;
    case cmd::kEditCopyKeyPathAbbrev:
    case cmd::kEditCopyKeyPathRegEdit:
    case cmd::kEditCopyKeyPathRegFile:
    case cmd::kEditCopyKeyPathPowerShell:
    case cmd::kEditCopyKeyPathPowerShellProvider:
    case cmd::kEditCopyKeyPathEscaped:
    case cmd::kEditCopyKeyPathNative:
    case cmd::kEditCopyKeyPathNativeResolved:
        CopyKeyPathAs(command, node, key_path);
        return;
    case kSearchPermissions:
        if (node_ok && key_exists)
        {
            ShowPermissionsDialog(node);
        }
        return;
    case kSearchOpenHive:
        if (can_open_hive)
        {
            focus_key();
            HandleMenuCommand(cmd::kOptionsHiveFileDir);
        }
        return;
    case kSearchExport:
        if (can_export)
        {
            run_on_target(cmd::kFileExport);
        }
        return;
    case kSearchRename:
        if (can_rename)
        {
            run_on_target(cmd::kEditRename);
        }
        return;
    case kSearchDelete:
        if (can_delete)
        {
            run_on_target(cmd::kEditDelete);
        }
        return;
    default:
        return;
    }
}

}
