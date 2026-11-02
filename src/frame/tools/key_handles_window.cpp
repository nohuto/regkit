// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/tools/key_handles_window.h"

#include "registry/registry_path.h"
#include "resource.h"
#include "ui/dialog_layout.h"
#include "ui/dialog_support.h"
#include "ui/feedback.h"
#include "win32/key_handle_snapshot.h"
#include "win32/process_rights.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/dialog/dialog_layout.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/message_dialog.h>
#include <win32ui/list_view/sort.h>
#include <win32ui/menu/popup_menu.h>
#include <win32ui/status_bar/status_bar.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/worker.h>
#include <win32ui/theme/theme.h>
#include <win32ui/window.h>
#include <win32ui/window_metrics.h>

#include <algorithm>
#include <array>
#include <memory>
#include <numeric>
#include <utility>
#include <vector>

namespace regkit
{

namespace
{

namespace support = editors::dialog_support;

constexpr UINT kSnapshotReady = WM_APP + 1;
constexpr size_t kNoRow = static_cast<size_t>(-1);

enum Column
{
    kProcess,
    kPid,
    kHandle,
    kObject,
    kKey,
    kAccess,
    kAttributes,
    kNative,
    kColumnCount,
};

enum MenuCommand
{
    kMenuOpen = 1,
    kMenuOpenNewTab,
    kMenuCopyKey,
    kMenuCopyNative,
    kMenuCopyProcess,
    kMenuCopyPid,
    kMenuCopyHandle,
    kMenuCopyAccess,
    kMenuRefresh,
    kMenuClose,
};

struct Row
{
    std::array<std::wstring, kColumnCount> text;
    win32::KeyHandle handle;
};

struct Payload
{
    uint64_t generation = 0;
    std::vector<Row> rows;
    size_t inaccessible = 0;
    DWORD error = ERROR_SUCCESS;
};

struct State
{
    KeyHandlesNavigate navigate;
    HWND list = nullptr;
    HWND status = nullptr;
    win32ui::Session session;
    std::vector<Row> rows;
    // virtual list indexes rows by sort order, then filtered order
    std::vector<size_t> order;
    std::vector<size_t> shown;
    std::wstring filter;
    size_t inaccessible = 0;
    DWORD error = ERROR_SUCCESS;
    bool scanning = false;
    int sort_column = kPid;
    bool ascending = true;
    util::SortKeys sort_earlier;
};

std::wstring DecodeAccess(ACCESS_MASK mask)
{
    static constexpr std::pair<ACCESS_MASK, const wchar_t*> kRights[] = {
        {KEY_ALL_ACCESS, L"KEY_ALL_ACCESS"},
        {KEY_READ, L"KEY_READ"},
        {KEY_WRITE, L"KEY_WRITE"},
        {KEY_QUERY_VALUE, L"KEY_QUERY_VALUE"},
        {KEY_SET_VALUE, L"KEY_SET_VALUE"},
        {KEY_CREATE_SUB_KEY, L"KEY_CREATE_SUB_KEY"},
        {KEY_ENUMERATE_SUB_KEYS, L"KEY_ENUMERATE_SUB_KEYS"},
        {KEY_NOTIFY, L"KEY_NOTIFY"},
        {KEY_CREATE_LINK, L"KEY_CREATE_LINK"},
        {DELETE, L"DELETE"},
        {READ_CONTROL, L"READ_CONTROL"},
        {WRITE_DAC, L"WRITE_DAC"},
        {WRITE_OWNER, L"WRITE_OWNER"},
    };
    std::wstring text;
    ACCESS_MASK rest = mask;
    for (const auto& [bits, name] : kRights)
    {
        if ((rest & bits) == bits)
        {
            text.append(text.empty() ? L"" : L", ").append(name);
            rest &= ~bits;
        }
    }
    wchar_t number[32] = {};
    if (rest != 0)
    {
        swprintf_s(number, L"0x%lX", rest);
        text.append(text.empty() ? L"" : L", ").append(number);
    }
    swprintf_s(number, L"%s(0x%08lX)", text.empty() ? L"" : L" ", mask);
    return text.append(number);
}

std::wstring DecodeAttributes(ULONG attributes)
{
    std::wstring text;
    for (const auto& [bit, name] :
         {std::pair<ULONG, const wchar_t*>{0x2, L"Inherit"}, {0x1, L"Protect"}, {0x4, L"Audit"}})
    {
        if (attributes & bit)
        {
            text.append(text.empty() ? L"" : L", ").append(name);
        }
    }
    return text;
}

std::vector<Row> BuildRows(std::vector<win32::KeyHandle>&& handles)
{
    const std::wstring sid = util::GetCurrentUserSidString();
    std::vector<Row> rows;
    rows.reserve(handles.size());
    for (win32::KeyHandle& handle : handles)
    {
        Row row;
        wchar_t number[32] = {};
        row.text[kProcess] = std::move(handle.process);
        swprintf_s(number, L"%Iu", handle.process_id);
        row.text[kPid] = number;
        swprintf_s(number, L"0x%IX", handle.handle);
        row.text[kHandle] = number;
        if (handle.object != 0)
        {
            swprintf_s(number, L"0x%IX", handle.object);
            row.text[kObject] = number;
        }
        if (!handle.name.empty())
        {
            row.text[kKey] = registry_path::Normalize(handle.name, sid);
        }
        row.text[kAccess] = DecodeAccess(handle.access);
        row.text[kAttributes] = DecodeAttributes(handle.attributes);
        row.text[kNative] = std::move(handle.name);
        row.handle = std::move(handle);
        rows.push_back(std::move(row));
    }
    return rows;
}

template <typename T>
int Compare(T left, T right)
{
    return left < right ? -1 : (right < left ? 1 : 0);
}

int CompareRows(const Row& left, const Row& right, int column)
{
    switch (column)
    {
    case kPid:
        return Compare(left.handle.process_id, right.handle.process_id);
    case kHandle:
        return Compare(left.handle.handle, right.handle.handle);
    case kObject:
        return Compare(left.handle.object, right.handle.object);
    case kAccess:
        return Compare(left.handle.access, right.handle.access);
    default:
        return util::CompareListText(left.text[column], right.text[column]);
    }
}

std::vector<const Row*> SelectedRows(const State* state)
{
    std::vector<const Row*> rows;
    for (int item = ListView_GetNextItem(state->list, -1, LVNI_SELECTED); item >= 0;
         item = ListView_GetNextItem(state->list, item, LVNI_SELECTED))
    {
        if (static_cast<size_t>(item) < state->shown.size())
        {
            rows.push_back(&state->rows[state->shown[static_cast<size_t>(item)]]);
        }
    }
    return rows;
}

void UpdateStatus(const State* state)
{
    std::wstring parts[3];
    if (state->scanning)
    {
        parts[0] = util::Tr(L"Scanning handles...");
    }
    else if (state->error != ERROR_SUCCESS)
    {
        parts[0] = util::Tr(L"Handles couldn't be enumerated.") + std::wstring(L" ") + win32ui::FormatWin32Error(state->error);
    }
    else
    {
        parts[0] = util::TrLabel(L"Handles", std::to_wstring(state->rows.size()));
        parts[1] = util::TrLabel(L"Shown", std::to_wstring(state->shown.size()));
        parts[2] = util::TrLabel(L"Inaccessible processes", std::to_wstring(state->inaccessible));
    }
    for (int part = 0; part < 3; ++part)
    {
        win32ui::SetStatusText(state->status, part, parts[part]);
    }
}

void Layout(HWND dialog, const State* state)
{
    SendMessageW(state->status, WM_SIZE, 0, 0);
    RECT client = {};
    RECT bar = {};
    RECT list = {};
    GetClientRect(dialog, &client);
    win32ui::ChildRect(dialog, state->status, &bar);
    win32ui::ChildRect(dialog, state->list, &list);
    SetWindowPos(state->list, nullptr, -1, list.top, client.right + 2, bar.top - list.top + 1, SWP_NOZORDER | SWP_NOACTIVATE);
}

bool Matches(const Row& row, const std::wstring& filter)
{
    return filter.empty() || std::any_of(row.text.begin(), row.text.end(), [&](const std::wstring& text) { return util::ContainsInsensitive(text, filter); });
}

// extended filters only remove rows, scan the visible subset
void FilterRows(State* state, std::wstring filter)
{
    const bool narrower = !state->filter.empty() && util::ContainsInsensitive(filter, state->filter);
    std::vector<size_t> shown;
    for (const size_t index : narrower ? state->shown : state->order)
    {
        if (Matches(state->rows[index], filter))
        {
            shown.push_back(index);
        }
    }
    state->shown = std::move(shown);
    state->filter = std::move(filter);
}

// sort clicks keep shown order for ties, new snapshots use earlier sort columns
void SortRows(State* state, int column, bool toggle)
{
    win32ui::UpdateListSortState(column, toggle, &state->sort_column, &state->ascending, &state->sort_earlier);
    const util::SortKeys none;
    const util::SortKeys& earlier = toggle ? none : state->sort_earlier;
    const auto before = [&](size_t left, size_t right) {
        return util::CompareSorted(state->sort_column, state->ascending, earlier, [&](int key) { return CompareRows(state->rows[left], state->rows[right], key); }) < 0;
    };
    std::stable_sort(state->order.begin(), state->order.end(), before);
    std::stable_sort(state->shown.begin(), state->shown.end(), before);
    win32ui::UpdateListViewSort(state->list, state->sort_column, state->ascending);
}

template <typename Change>
void ChangeView(State* state, Change change)
{
    const auto row_at = [state](int item) {
        return item >= 0 && static_cast<size_t>(item) < state->shown.size() ? state->shown[static_cast<size_t>(item)] : kNoRow;
    };
    const win32ui::ListSelection before = win32ui::CaptureListSelection(state->list);
    std::vector<size_t> selected;
    if (before.all)
    {
        selected = state->shown;
    }
    for (const int item : before.rows)
    {
        if (const size_t row = row_at(item); row != kNoRow)
        {
            selected.push_back(row);
        }
    }
    const size_t focused = row_at(before.focused);
    const size_t mark = row_at(before.mark);
    change();
    std::vector<int> position(state->rows.size(), -1);
    for (size_t item = 0; item < state->shown.size(); ++item)
    {
        position[state->shown[item]] = static_cast<int>(item);
    }
    const auto item_of = [&position](size_t row) { return row < position.size() ? position[row] : -1; };
    ListView_SetItemState(state->list, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetItemCountEx(state->list, static_cast<int>(state->shown.size()), 0);
    win32ui::ListSelection after;
    after.rows.reserve(selected.size());
    for (const size_t row : selected)
    {
        after.rows.push_back(item_of(row));
    }
    after.focused = item_of(focused);
    after.mark = item_of(mark);
    win32ui::ApplyListSelection(state->list, after);
}

void ShowRows(HWND dialog, State* state)
{
    state->order.resize(state->rows.size());
    std::iota(state->order.begin(), state->order.end(), size_t{0});
    state->shown.clear();
    state->filter.clear();
    FilterRows(state, support::ReadText(dialog, IDC_KH_FILTER));
    SortRows(state, state->sort_column, false);
    ListView_SetItemState(state->list, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetSelectionMark(state->list, -1);
    ListView_SetItemCountEx(state->list, static_cast<int>(state->shown.size()), 0);
    if (!state->shown.empty())
    {
        ListView_EnsureVisible(state->list, 0, FALSE);
    }
    UpdateStatus(state);
}

void ApplyFilter(HWND dialog, State* state)
{
    std::wstring filter = support::ReadText(dialog, IDC_KH_FILTER);
    ChangeView(state, [&] { FilterRows(state, std::move(filter)); });
    // keep focus visible if it passes the filter, else show the first match
    const int focused = ListView_GetNextItem(state->list, -1, LVNI_FOCUSED);
    if (!state->shown.empty())
    {
        ListView_EnsureVisible(state->list, focused >= 0 ? focused : 0, FALSE);
    }
    UpdateStatus(state);
}

// list stores no text, library uses this for copy and text search
bool CellText(void*, HWND list, int item, int subitem, std::wstring* text)
{
    const auto* state = reinterpret_cast<const State*>(GetWindowLongPtrW(GetParent(list), DWLP_USER));
    if (!state || item < 0 || static_cast<size_t>(item) >= state->shown.size() || subitem < 0 || subitem >= kColumnCount)
    {
        return false;
    }
    *text = state->rows[state->shown[static_cast<size_t>(item)]].text[static_cast<size_t>(subitem)];
    return true;
}

void Refresh(HWND dialog, State* state)
{
    state->scanning = true;
    UpdateStatus(state);
    const uint64_t started = state->session.Start(L"KeyHandleScanThread", [dialog](uint64_t generation, std::atomic_bool& cancel) {
        win32::KeyHandleSnapshot snapshot = win32::SnapshotKeyHandles(cancel);
        if (cancel.load())
        {
            return;
        }
        auto payload = std::make_unique<Payload>();
        payload->generation = generation;
        payload->rows = BuildRows(std::move(snapshot.handles));
        payload->inaccessible = snapshot.inaccessible_processes;
        payload->error = snapshot.error;
        if (PostMessageW(dialog, kSnapshotReady, 0, reinterpret_cast<LPARAM>(payload.get())))
        {
            payload.release();
        }
    });
    // no thread for the scan, F5 tries again
    if (!started)
    {
        state->scanning = false;
        state->error = ERROR_NOT_ENOUGH_MEMORY;
        UpdateStatus(state);
    }
}

void Copy(HWND dialog, const State* state, int column)
{
    std::wstring text;
    for (const Row* row : SelectedRows(state))
    {
        text.append(text.empty() ? L"" : L"\r\n").append(row->text[column]);
    }
    if (!text.empty())
    {
        win32ui::CopyTextToClipboard(dialog, text);
    }
}

void OpenKey(const State* state, bool new_tab)
{
    const std::vector<const Row*> rows = SelectedRows(state);
    if (rows.size() == 1 && !rows.front()->text[kKey].empty() && state->navigate)
    {
        state->navigate(rows.front()->text[kKey], new_tab);
    }
}

void CloseHandles(HWND dialog, State* state)
{
    std::vector<win32::KeyHandle> targets;
    std::wstring handles;
    for (const Row* row : SelectedRows(state))
    {
        targets.push_back(row->handle);
        handles.append(handles.empty() ? L"" : L"\r\n")
            .append(row->text[kProcess])
            .append(L" (")
            .append(row->text[kPid])
            .append(L") ")
            .append(row->text[kHandle])
            .append(L"  ")
            .append(row->text[kKey]);
    }
    if (targets.empty())
    {
        return;
    }
    const std::wstring message =
        util::Tr(L"Closing a handle owned by another process can make that process malfunction or crash.\n\n"
                 L"Each handle is checked against a fresh snapshot first, but it can still be closed and reused between that "
                 L"check and the close.");
    if (ui::PromptKeyChoice(dialog, message, handles, targets.size() == 1 ? util::Tr(L"Close Handle") : util::Tr(L"Close Handles"), util::Tr(L"Close"), L"", util::Tr(L"Cancel"), {70, 70, 70}) != IDYES)
    {
        return;
    }
    size_t closed = 0;
    const DWORD error = win32::CloseKeyHandles(targets, &closed);
    if (error != ERROR_SUCCESS)
    {
        ui::ShowError(dialog, util::TrDetail(L"Not all handles were closed.", win32ui::FormatWin32Error(error)));
    }
    Refresh(dialog, state);
}

bool ShowRowMenu(void*, HWND list, int, POINT screen)
{
    const HWND dialog = GetParent(list);
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (!state)
    {
        return false;
    }
    const std::vector<const Row*> rows = SelectedRows(state);
    const UINT any = rows.empty() ? MF_GRAYED : 0;
    const UINT one = rows.size() == 1 && !rows.front()->text[kKey].empty() ? 0 : MF_GRAYED;
    HMENU menu = CreatePopupMenu();
    if (!menu)
    {
        return true;
    }
    win32ui::AppendMenuItem(menu, MF_STRING | one, kMenuOpen, util::Tr(L"Open &Key"), L"Enter");
    AppendMenuW(menu, MF_STRING | one, kMenuOpenNewTab, util::Tr(L"Open Key in New &Tab"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING | any, kMenuCopyKey, util::Tr(L"Cop&y Key"));
    AppendMenuW(menu, MF_STRING | any, kMenuCopyNative, util::Tr(L"Copy &Native Name"));
    AppendMenuW(menu, MF_STRING | any, kMenuCopyProcess, util::Tr(L"Copy &Process"));
    AppendMenuW(menu, MF_STRING | any, kMenuCopyPid, util::Tr(L"Copy P&ID"));
    AppendMenuW(menu, MF_STRING | any, kMenuCopyHandle, util::Tr(L"Copy &Handle"));
    AppendMenuW(menu, MF_STRING | any, kMenuCopyAccess, util::Tr(L"C&opy Access"));
    win32ui::AppendListRowCommands(menu, list);
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    win32ui::AppendMenuItem(menu, MF_STRING, kMenuRefresh, util::Tr(L"&Refresh"), L"F5");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    win32ui::AppendMenuItem(menu, MF_STRING | any, kMenuClose, util::Tr(L"C&lose Handle..."), L"Del");
    const int chosen = win32ui::TrackContextMenu(menu, screen, dialog);
    // modeless window may be destroyed during the menu loop
    if (!IsWindow(list) || win32ui::RunListRowCommand(list, static_cast<UINT>(chosen)))
    {
        return true;
    }
    static constexpr int kCopyColumns[] = {kKey, kNative, kProcess, kPid, kHandle, kAccess};
    switch (chosen)
    {
    case kMenuOpen:
    case kMenuOpenNewTab:
        OpenKey(state, chosen == kMenuOpenNewTab);
        break;
    case kMenuRefresh:
        Refresh(dialog, state);
        break;
    case kMenuClose:
        CloseHandles(dialog, state);
        break;
    default:
        if (chosen >= kMenuCopyKey && chosen <= kMenuCopyAccess)
        {
            Copy(dialog, state, kCopyColumns[chosen - kMenuCopyKey]);
        }
        break;
    }
    return true;
}

INT_PTR HandleNotify(HWND dialog, State* state, NMHDR* header)
{
    switch (header->idFrom == IDC_KH_LIST ? header->code : 0)
    {
    case LVN_GETDISPINFOW:
        {
            auto* info = reinterpret_cast<NMLVDISPINFOW*>(header);
            if ((info->item.mask & LVIF_TEXT) && info->item.iItem >= 0 && static_cast<size_t>(info->item.iItem) < state->shown.size() && info->item.iSubItem >= 0 &&
                info->item.iSubItem < kColumnCount)
            {
                const Row& row = state->rows[state->shown[static_cast<size_t>(info->item.iItem)]];
                info->item.pszText = const_cast<wchar_t*>(row.text[static_cast<size_t>(info->item.iSubItem)].c_str());
            }
            return TRUE;
        }
    case LVN_COLUMNCLICK:
        ChangeView(state, [&] { SortRows(state, reinterpret_cast<NMLISTVIEW*>(header)->iSubItem, true); });
        return TRUE;
    case LVN_ITEMACTIVATE:
        OpenKey(state, false);
        return TRUE;
    case LVN_KEYDOWN:
        {
            const WORD key = reinterpret_cast<NMLVKEYDOWN*>(header)->wVKey;
            if (key == VK_F5)
            {
                Refresh(dialog, state);
            }
            else if (key == VK_DELETE)
            {
                CloseHandles(dialog, state);
            }
            return TRUE;
        }
    default:
        break;
    }
    return FALSE;
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    switch (message)
    {
    case WM_INITDIALOG:
        {
            state = reinterpret_cast<std::unique_ptr<State>*>(lparam)->release();
            SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
            state->list = GetDlgItem(dialog, IDC_KH_LIST);
            state->status = GetDlgItem(dialog, IDC_KH_STATUS);
            SendDlgItemMessageW(dialog, IDC_KH_FILTER, EM_SETCUEBANNER, TRUE, reinterpret_cast<LPARAM>(util::Tr(L"Filter")));
            static const win32ui::ListViewHooks hooks = [] {
                win32ui::ListViewHooks list_hooks = win32ui::DialogListHooks(ShowRowMenu);
                list_hooks.cell_text = CellText;
                return list_hooks;
            }();
            win32ui::SetupDialogList(state->list, &hooks, 0, {{util::Tr(L"Process"), 150}, {util::Tr(L"PID"), 60}, {util::Tr(L"Handle"), 80}, {util::Tr(L"Object Address"), 140}, {util::Tr(L"Key"), 400}, {util::Tr(L"Access"), 300}, {util::Tr(L"Attributes"), 90}, {util::Tr(L"Native Name"), 400}});
            SetWindowPos(state->status, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            win32ui::AnchorDialogControls(dialog, {
                                                      {IDC_KH_FILTER, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  });
            static const win32ui::StatusBarHooks status_hooks = appearance::StatusBarThemeHooks();
            win32ui::AttachStatusBar(state->status, util::Tr(L"Status Bar"), &status_hooks);
            win32ui::SetStatusParts(state->status, {{L"", 170}, {L"", 170}, {L"", 0, true}});
            Layout(dialog, state);
            ShowRows(dialog, state);
            Refresh(dialog, state);
            return TRUE;
        }
    case kSnapshotReady:
        {
            std::unique_ptr<Payload> payload(reinterpret_cast<Payload*>(lparam));
            if (state && state->session.IsCurrent(payload->generation))
            {
                state->rows = std::move(payload->rows);
                state->inaccessible = payload->inaccessible;
                state->error = payload->error;
                state->scanning = false;
                ShowRows(dialog, state);
            }
            return TRUE;
        }
    case WM_SIZE:
        if (state)
        {
            Layout(dialog, state);
        }
        return FALSE;
    case WM_NOTIFY:
        return state ? HandleNotify(dialog, state, reinterpret_cast<NMHDR*>(lparam)) : FALSE;
    case WM_COMMAND:
        if (!state)
        {
            return FALSE;
        }
        switch (LOWORD(wparam))
        {
        case IDC_KH_FILTER:
            if (HIWORD(wparam) == EN_CHANGE)
            {
                ApplyFilter(dialog, state);
            }
            return TRUE;
        case IDOK:
            OpenKey(state, false);
            return TRUE;
        case IDCANCEL:
            win32ui::EndDialogWindow(dialog, IDCANCEL);
            return TRUE;
        default:
            return FALSE;
        }
    case WM_DESTROY:
        if (state)
        {
            state->session.CancelAndJoin();
            MSG pending = {};
            while (PeekMessageW(&pending, dialog, kSnapshotReady, kSnapshotReady, PM_REMOVE))
            {
                delete reinterpret_cast<Payload*>(pending.lParam);
            }
        }
        return TRUE;
    case WM_NCDESTROY:
        delete state;
        return TRUE;
    default:
        break;
    }
    return FALSE;
}

}

HWND ShowKeyHandlesWindow(HWND owner, KeyHandlesNavigate navigate)
{
    auto state = std::make_unique<State>();
    state->navigate = std::move(navigate);
    HWND dialog = win32ui::OpenProcDialog(MAKEINTRESOURCEW(IDD_KEY_HANDLES), DialogProc, reinterpret_cast<LPARAM>(&state), owner);
    if (!dialog)
    {
        return nullptr;
    }
    return dialog;
}

}
