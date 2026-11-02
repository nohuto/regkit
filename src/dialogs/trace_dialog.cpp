// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/trace_dialog.h"

#include <win32ui/dialog/dialog_window.h>
#include <win32ui/tree_view/tree_view.h>

#include <algorithm>
#include <cwchar>
#include <memory>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include <commctrl.h>
#include <windowsx.h>

#include "registry/registry_path.h"
#include "resource.h"
#include "ui/feedback.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#ifndef NMTVITEMCHANGE
typedef struct tagNMTVITEMCHANGE
{
    NMHDR hdr;
    UINT uChanged;
    HTREEITEM hItem;
    UINT uStateNew;
    UINT uStateOld;
    LPARAM lParam;
} NMTVITEMCHANGE;
#endif

namespace regkit
{

namespace
{

constexpr UINT kDialogAddEntriesMessage = WM_APP + 1;
constexpr UINT kDialogDoneMessage = WM_APP + 2;
constexpr UINT kDialogProcessEntriesMessage = WM_APP + 3;

using util::ToLower;

struct TraceNodeData
{
    bool is_value = false;
    std::wstring key_path;
    std::wstring value_name;
};

struct TraceDialogState
{
    HWND dialog = nullptr;
    HWND status = nullptr;
    HWND tree = nullptr;
    std::wstring title;
    trace::Selection* out = nullptr;
    bool loading_done = false;
    bool show_values = true;
    TraceDialogReadyCallback on_ready = nullptr;
    void* on_ready_context = nullptr;
    bool processing_entries = false;
    bool updating_checks = false;
    size_t pending_index = 0;

    std::unordered_map<std::wstring, HTREEITEM> key_nodes;
    std::unordered_set<std::wstring> key_entries;
    std::unordered_map<std::wstring, std::unordered_map<std::wstring, std::wstring>> values_by_key;
    std::unordered_map<std::wstring, std::unordered_set<std::wstring>> values_loaded;
    std::vector<std::unique_ptr<TraceNodeData>> node_storage;
    std::vector<KeyValueDialogEntry> pending_entries;
    size_t key_count = 0;
    size_t value_count = 0;
};

TraceNodeData* StoreNodeData(TraceDialogState* state, bool is_value, const std::wstring& key_path, const std::wstring& value_name)
{
    auto node = std::make_unique<TraceNodeData>();
    node->is_value = is_value;
    node->key_path = key_path;
    node->value_name = value_name;
    TraceNodeData* raw = node.get();
    state->node_storage.push_back(std::move(node));
    return raw;
}

void UpdateStatus(TraceDialogState* state)
{
    if (!state || !state->status)
    {
        return;
    }
    std::wstring text = util::TrLabel(L"Keys", std::to_wstring(state->key_count));
    if (state->show_values)
    {
        text.append(L", ").append(util::TrLabel(L"Values", std::to_wstring(state->value_count)));
    }
    const bool ready = state->loading_done && !state->processing_entries;
    if (!ready)
    {
        text.append(L" (").append(util::Tr(L"loading...")).append(L")");
    }
    EnableWindow(GetDlgItem(state->dialog, IDOK), ready);
    SetWindowTextW(state->status, text.c_str());
}

HTREEITEM EnsureKeyNode(HWND tree, TraceDialogState* state, const std::wstring& key_path, const std::wstring& display_path)
{
    if (!state || !tree || key_path.empty())
    {
        return nullptr;
    }
    std::wstring tree_path = display_path.empty() ? key_path : display_path;
    std::vector<std::wstring> tree_parts = registry_path::Split(tree_path);
    if (tree_parts.empty())
    {
        return nullptr;
    }
    std::vector<std::wstring> key_parts = registry_path::Split(key_path);
    size_t match_suffix = 0;
    while (match_suffix < tree_parts.size() && match_suffix < key_parts.size())
    {
        size_t tree_index = tree_parts.size() - 1 - match_suffix;
        size_t key_index = key_parts.size() - 1 - match_suffix;
        if (!util::EqualsInsensitive(tree_parts[tree_index], key_parts[key_index]))
        {
            break;
        }
        ++match_suffix;
    }
    size_t align_start_tree = tree_parts.size() - match_suffix;
    size_t align_start_key = key_parts.size() - match_suffix;
    auto key_prefix_for_index = [&](size_t index) -> std::wstring {
        if (key_parts.empty())
        {
            return L"";
        }
        size_t part_count = 0;
        if (index < align_start_tree)
        {
            if (index == 0 && util::EqualsInsensitive(tree_parts[0], L"REGISTRY"))
            {
                return L"";
            }
            part_count = 1;
        }
        else
        {
            part_count = align_start_key + (index - align_start_tree) + 1;
        }
        return registry_path::JoinPrefix(key_parts, part_count);
    };

    std::wstring current_tree;
    HTREEITEM parent = TVI_ROOT;
    for (size_t i = 0; i < tree_parts.size(); ++i)
    {
        if (!current_tree.empty())
        {
            current_tree.append(L"\\");
        }
        current_tree.append(tree_parts[i]);
        std::wstring tree_lower = ToLower(current_tree);
        auto it = state->key_nodes.find(tree_lower);
        if (it != state->key_nodes.end())
        {
            parent = it->second;
            continue;
        }

        std::wstring node_key_path = (i + 1 == tree_parts.size()) ? key_path : key_prefix_for_index(i);
        TraceNodeData* data = StoreNodeData(state, false, node_key_path, L"");
        TVINSERTSTRUCTW insert = {};
        insert.hParent = parent;
        insert.hInsertAfter = TVI_LAST;
        insert.item.mask = TVIF_TEXT | TVIF_PARAM;
        insert.item.pszText = const_cast<wchar_t*>(tree_parts[i].c_str());
        insert.item.lParam = reinterpret_cast<LPARAM>(data);
        HTREEITEM item = TreeView_InsertItem(tree, &insert);
        if (parent != TVI_ROOT && TreeView_GetCheckState(tree, parent) != FALSE)
        {
            bool prior = state->updating_checks;
            state->updating_checks = true;
            TreeView_SetCheckState(tree, item, TRUE);
            state->updating_checks = prior;
        }
        state->key_nodes.emplace(std::move(tree_lower), item);
        parent = item;
    }
    return parent == TVI_ROOT ? nullptr : parent;
}

HTREEITEM InsertValueNode(HWND tree, TraceDialogState* state, HTREEITEM key_item, const std::wstring& key_path, const std::wstring& value_name)
{
    if (!state || !tree || !key_item)
    {
        return nullptr;
    }
    std::wstring display = value_name.empty() ? util::Tr(L"(Default)") : value_name;
    TraceNodeData* data = StoreNodeData(state, true, key_path, value_name);
    TVINSERTSTRUCTW insert = {};
    insert.hParent = key_item;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM;
    insert.item.pszText = const_cast<wchar_t*>(display.c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(data);
    HTREEITEM item = TreeView_InsertItem(tree, &insert);
    if (TreeView_GetCheckState(tree, key_item) != FALSE)
    {
        bool prior = state->updating_checks;
        state->updating_checks = true;
        TreeView_SetCheckState(tree, item, TRUE);
        state->updating_checks = prior;
    }
    return item;
}

void EnsureValueNodes(HWND tree, TraceDialogState* state, HTREEITEM key_item, const std::wstring& key_path)
{
    if (!state || !tree || !key_item || key_path.empty())
    {
        return;
    }
    std::wstring key_lower = ToLower(key_path);
    auto values_it = state->values_by_key.find(key_lower);
    if (values_it == state->values_by_key.end())
    {
        return;
    }
    auto& loaded = state->values_loaded[key_lower];
    for (const auto& [value_lower, value_name] : values_it->second)
    {
        if (loaded.insert(value_lower).second)
        {
            InsertValueNode(tree, state, key_item, key_path, value_name);
        }
    }
}

void AddEntry(HWND tree, TraceDialogState* state, const KeyValueDialogEntry& entry)
{
    if (!state || !tree || entry.key_path.empty())
    {
        return;
    }
    std::wstring display_path = entry.display_path.empty() ? entry.key_path : entry.display_path;
    HTREEITEM key_item = EnsureKeyNode(tree, state, entry.key_path, display_path);
    std::wstring key_lower = ToLower(entry.key_path);
    if (state->key_entries.insert(key_lower).second)
    {
        state->key_count++;
    }
    if (!entry.has_value || !state->show_values)
    {
        return;
    }
    std::wstring value_lower = ToLower(entry.value_name);
    auto& values = state->values_by_key[key_lower];
    if (values.emplace(value_lower, entry.value_name).second)
    {
        state->value_count++;
        if (key_item && (TreeView_GetItemState(tree, key_item, TVIS_EXPANDED) & TVIS_EXPANDED) && state->values_loaded[key_lower].insert(value_lower).second)
        {
            InsertValueNode(tree, state, key_item, entry.key_path, entry.value_name);
        }
    }
}

TraceNodeData* GetNodeData(HWND tree, HTREEITEM item)
{
    LPARAM data = 0;
    return win32ui::TreeItemData(tree, item, &data) ? reinterpret_cast<TraceNodeData*>(data) : nullptr;
}

void AppendCheckedNodes(HWND tree, HTREEITEM item, trace::Selection* selection, std::unordered_set<std::wstring>* seen_keys)
{
    while (item)
    {
        TraceNodeData* data = GetNodeData(tree, item);
        if (data && TreeView_GetCheckState(tree, item))
        {
            if (data->is_value)
            {
                std::wstring key_lower = ToLower(data->key_path);
                std::wstring value_lower = ToLower(data->value_name);
                selection->values_by_key[key_lower].insert(value_lower);
                if (seen_keys && seen_keys->insert(key_lower).second)
                {
                    selection->key_paths.push_back(data->key_path);
                }
            }
            else
            {
                std::wstring key_lower = ToLower(data->key_path);
                if (seen_keys && seen_keys->insert(key_lower).second)
                {
                    selection->key_paths.push_back(data->key_path);
                }
                for (HTREEITEM child = TreeView_GetChild(tree, item); child;
                     child = TreeView_GetNextSibling(tree, child))
                {
                    TraceNodeData* child_data = GetNodeData(tree, child);
                    if (child_data && child_data->is_value)
                    {
                        selection->values_by_key[key_lower];
                        break;
                    }
                }
            }
        }
        HTREEITEM child = TreeView_GetChild(tree, item);
        if (child)
        {
            AppendCheckedNodes(tree, child, selection, seen_keys);
        }
        item = TreeView_GetNextSibling(tree, item);
    }
}

void QueueEntries(HWND hwnd, TraceDialogState* state, std::vector<KeyValueDialogEntry>&& entries)
{
    if (!state || entries.empty())
    {
        return;
    }
    state->pending_entries.reserve(state->pending_entries.size() + entries.size());
    for (auto& entry : entries)
    {
        state->pending_entries.push_back(std::move(entry));
    }
    if (!state->processing_entries)
    {
        state->processing_entries = true;
        PostMessageW(hwnd, kDialogProcessEntriesMessage, 0, 0);
    }
}

void ProcessPendingEntries(HWND hwnd, TraceDialogState* state)
{
    if (!state || !state->processing_entries)
    {
        return;
    }
    constexpr size_t kBatchSize = 128;
    constexpr DWORD kBatchMs = 8;
    uint64_t start_tick = GetTickCount64();
    size_t processed = 0;
    while (state->pending_index < state->pending_entries.size())
    {
        AddEntry(state->tree, state, state->pending_entries[state->pending_index]);
        ++state->pending_index;
        ++processed;
        if (processed >= kBatchSize)
        {
            break;
        }
        if (GetTickCount64() - start_tick >= kBatchMs)
        {
            break;
        }
    }
    if (state->pending_index >= state->pending_entries.size())
    {
        state->pending_entries.clear();
        state->pending_index = 0;
        state->processing_entries = false;
    }
    else
    {
        PostMessageW(hwnd, kDialogProcessEntriesMessage, 0, 0);
    }
    UpdateStatus(state);
}

void AcceptSelection(HWND hwnd, TraceDialogState* state, bool select_all)
{
    if (!state || !state->out)
    {
        return;
    }
    bool recursive = IsDlgButtonChecked(hwnd, IDC_TRACE_RECURSIVE) == BST_CHECKED;
    if (select_all)
    {
        state->out->select_all = true;
        state->out->recursive = recursive;
        state->out->key_paths.clear();
        state->out->values_by_key.clear();
        win32ui::EndDialogWindow(hwnd, IDOK);
        return;
    }

    trace::Selection selection = {};
    selection.select_all = false;
    selection.recursive = recursive;
    std::unordered_set<std::wstring> seen_keys;
    HTREEITEM root = TreeView_GetRoot(state->tree);
    if (root)
    {
        AppendCheckedNodes(state->tree, root, &selection, &seen_keys);
    }
    if (selection.key_paths.empty() && selection.values_by_key.empty())
    {
        ui::ShowWarning(hwnd, util::Tr(L"Select at least one key or value."));
        return;
    }
    *state->out = std::move(selection);
    win32ui::EndDialogWindow(hwnd, IDOK);
}

INT_PTR CALLBACK TraceDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<TraceDialogState*>(GetWindowLongPtrW(hwnd, DWLP_USER));
    if (msg == WM_INITDIALOG)
    {
        state = reinterpret_cast<TraceDialogState*>(lparam);
        SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        state->dialog = hwnd;
        state->status = GetDlgItem(hwnd, IDC_TRACE_STATUS);
        state->tree = GetDlgItem(hwnd, IDC_TRACE_TREE);
        // set after creation so the tree builds check box state images before adding items
        SetWindowLongPtrW(state->tree, GWL_STYLE, GetWindowLongPtrW(state->tree, GWL_STYLE) | TVS_CHECKBOXES);
        SendMessageW(state->tree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);
        CheckDlgButton(hwnd, IDC_TRACE_RECURSIVE, BST_CHECKED);
        if (!state->title.empty())
        {
            SetWindowTextW(hwnd, state->title.c_str());
        }
        UpdateStatus(state);
        SetFocus(state->tree);
        if (state->on_ready)
        {
            state->on_ready(hwnd, state->on_ready_context);
        }
        return FALSE;
    }
    if (!state)
    {
        return FALSE;
    }
    switch (msg)
    {
    case WM_DESTROY:
        {
            MSG pending = {};
            while (PeekMessageW(&pending, hwnd, kDialogAddEntriesMessage, kDialogAddEntriesMessage, PM_REMOVE))
            {
                delete reinterpret_cast<std::vector<KeyValueDialogEntry>*>(pending.lParam);
            }
            if (HIMAGELIST checks = TreeView_SetImageList(state->tree, nullptr, TVSIL_STATE))
            {
                ImageList_Destroy(checks);
            }
            return TRUE;
        }
    case WM_COMMAND:
        switch (LOWORD(wparam))
        {
        case IDC_TRACE_SELECT_ALL:
        case IDOK:
            AcceptSelection(hwnd, state, LOWORD(wparam) == IDC_TRACE_SELECT_ALL);
            return TRUE;
        case IDCANCEL:
            win32ui::EndDialogWindow(hwnd, IDCANCEL);
            return TRUE;
        default:
            break;
        }
        break;
    case WM_NOTIFY:
        {
            auto* header = reinterpret_cast<NMHDR*>(lparam);
            if (header && header->code == TVN_ITEMEXPANDINGW)
            {
                auto* info = reinterpret_cast<NMTREEVIEWW*>(lparam);
                if (info->action == TVE_EXPAND)
                {
                    TraceNodeData* data = GetNodeData(state->tree, info->itemNew.hItem);
                    if (data && !data->is_value)
                    {
                        EnsureValueNodes(state->tree, state, info->itemNew.hItem, data->key_path);
                    }
                }
            }
            if (header && header->code == TVN_ITEMCHANGEDW)
            {
                auto* change = reinterpret_cast<NMTVITEMCHANGE*>(lparam);
                if (change && (change->uChanged & TVIF_STATE) &&
                    ((change->uStateNew ^ change->uStateOld) & TVIS_STATEIMAGEMASK))
                {
                    if (!state->updating_checks)
                    {
                        TraceNodeData* data = GetNodeData(state->tree, change->hItem);
                        if (data && !data->is_value)
                        {
                            bool checked = TreeView_GetCheckState(state->tree, change->hItem) != FALSE;
                            state->updating_checks = true;
                            win32ui::SetTreeChecksBelow(state->tree, change->hItem, checked);
                            state->updating_checks = false;
                        }
                    }
                }
            }
            break;
        }
    case kDialogAddEntriesMessage:
        {
            std::unique_ptr<std::vector<KeyValueDialogEntry>> owned(
                reinterpret_cast<std::vector<KeyValueDialogEntry>*>(lparam)
            );
            if (owned)
            {
                QueueEntries(hwnd, state, std::move(*owned));
            }
            return TRUE;
        }
    case kDialogDoneMessage:
        state->loading_done = wparam != 0;
        UpdateStatus(state);
        return TRUE;
    case kDialogProcessEntriesMessage:
        ProcessPendingEntries(hwnd, state);
        return TRUE;
    default:
        break;
    }
    return FALSE;
}

}

bool ShowTraceDialog(HWND owner, const TraceDialogOptions& options, trace::Selection* selection, TraceDialogReadyCallback on_ready, void* context)
{
    TraceDialogState state;
    state.out = selection;
    state.title = options.title;
    state.show_values = options.show_values;
    state.on_ready = on_ready;
    state.on_ready_context = context;
    return selection && win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_TRACE_SELECT), TraceDialogProc, reinterpret_cast<LPARAM>(&state)) == IDOK;
}
void TraceDialogPostEntries(HWND dialog, std::vector<KeyValueDialogEntry>* entries)
{
    if (!dialog || !entries)
    {
        delete entries;
        return;
    }
    if (!PostMessageW(dialog, kDialogAddEntriesMessage, 0, reinterpret_cast<LPARAM>(entries)))
    {
        delete entries;
    }
}

void TraceDialogPostDone(HWND dialog, bool done)
{
    if (!dialog)
    {
        return;
    }
    PostMessageW(dialog, kDialogDoneMessage, done ? 1 : 0, 0);
}

}
