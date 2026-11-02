// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/bitfield_definition_editor.h"

#include "ui/dialog_layout.h"
#include "ui/dialog_support.h"
#include "ui/feedback.h"
#include "win32/file_dialog.h"
#include "win32/translation.h"

#include "resource.h"
#include "win32/text_transform.h"

#include <win32ui/dialog/combo_box.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/edit.h>
#include <win32ui/dialog/file_dialog.h>
#include <win32ui/dialog/message_dialog.h>
#include <win32ui/list_view/sort.h>
#include <win32ui/menu/popup_menu.h>
#include <win32ui/theme/theme.h>
#include <win32ui/window.h>
#include <win32ui/window_metrics.h>

#include <commctrl.h>

#include <algorithm>
#include <optional>
#include <utility>

namespace regkit::editors
{

namespace
{

using bitfield::Definition;
using bitfield::DefinitionFile;
using bitfield::Field;
using bitfield::State;

constexpr unsigned kWidths[] = {8, 16, 32, 64};

int SelectedItemIndex(HWND list)
{
    LPARAM data = -1;
    return win32ui::SelectedListViewData(list, &data) ? static_cast<int>(data) : -1;
}

struct FieldEditor
{
    Field field;
    const Definition* parent = nullptr;
    int editing = -1;
    bool accepted = false;
    int sort_column = 0;
    bool sort_ascending = true;
    util::SortKeys sort_earlier;
};

struct StateEditor
{
    State state;
    const std::vector<State>* states = nullptr;
    int editing = -1;
    unsigned bit_count = 0;
};

struct Editor
{
    DefinitionFile file;
    int selected = 0;
    bool single = false;
    bool lock_width = false;
    bool dirty = false;
    bool saved = false;
    bool accepted = false;
    bool loading = false;
    bool filtering = false;
    bool updating_combo = false;
    int field_sort_column = 0;
    bool field_sort_ascending = true;
    util::SortKeys field_sort_earlier;

    Definition& definition()
    {
        return file.definitions[static_cast<size_t>(selected)];
    }
};

std::wstring StatesText(const Field& field);

int CALLBACK CompareFieldBits(LPARAM left_data, LPARAM right_data, int column, void* context)
{
    auto* state = static_cast<FieldEditor*>(context);
    const unsigned left = static_cast<unsigned>(left_data);
    const unsigned right = static_cast<unsigned>(right_data);
    int result = 0;
    if (column == 0 || column == 1)
    {
        result = left < right ? -1 : left > right ? 1
                                                  : 0;
    }
    else if (column == 2 && state && state->parent)
    {
        const int left_owner = state->parent->FieldIndexForBit(left);
        const int right_owner = state->parent->FieldIndexForBit(right);
        const std::wstring left_name =
            left_owner >= 0 ? state->parent->fields[static_cast<size_t>(left_owner)].name : std::wstring();
        const std::wstring right_name =
            right_owner >= 0 ? state->parent->fields[static_cast<size_t>(right_owner)].name : std::wstring();
        result = util::CompareListText(left_name, right_name);
    }
    return result;
}

int CALLBACK CompareDefinitionFields(LPARAM left_data, LPARAM right_data, int column, void* context)
{
    auto* state = static_cast<Editor*>(context);
    if (!state || left_data < 0 || right_data < 0)
    {
        return 0;
    }
    const Definition& definition = state->definition();
    const size_t left_index = static_cast<size_t>(left_data);
    const size_t right_index = static_cast<size_t>(right_data);
    if (left_index >= definition.fields.size() || right_index >= definition.fields.size())
    {
        return 0;
    }
    const Field& left = definition.fields[left_index];
    const Field& right = definition.fields[right_index];
    int result = 0;
    switch (column)
    {
    case 0:
        result = util::CompareListText(left.name, right.name);
        break;
    case 1:
        result =
            std::lexicographical_compare(left.bits.begin(), left.bits.end(), right.bits.begin(), right.bits.end()) ? -1
            : std::lexicographical_compare(right.bits.begin(), right.bits.end(), left.bits.begin(), left.bits.end())
                ? 1
                : 0;
        break;
    case 2:
        result = util::CompareListText(StatesText(left), StatesText(right));
        break;
    case 3:
        result = util::CompareListText(left.meaning, right.meaning);
        break;
    default:
        break;
    }
    return result;
}

std::wstring BitsText(const Field& field)
{
    std::wstring text;
    for (size_t i = 0; i < field.bits.size(); ++i)
    {
        if (i > 0)
        {
            text.append(L", ");
        }
        text.append(std::to_wstring(field.bits[i]));
    }
    return text;
}

std::wstring StatesText(const Field& field)
{
    std::wstring text;
    for (size_t i = 0; i < field.states.size(); ++i)
    {
        if (i > 0)
        {
            text.append(L", ");
        }
        text.append(std::to_wstring(field.states[i].value)).append(L"=").append(field.states[i].name);
    }
    return text;
}

bool ShowEditListMenu(HWND list, POINT screen, std::initializer_list<std::pair<int, const wchar_t*>> commands)
{
    const HWND dialog = GetParent(list);
    const UINT row_flags = MF_STRING | (ListView_GetSelectedCount(list) > 0 ? 0 : MF_GRAYED);
    HMENU menu = CreatePopupMenu();
    if (!menu)
    {
        return true;
    }
    for (const auto& [command, label] : commands)
    {
        AppendMenuW(menu, command == commands.begin()->first ? MF_STRING : row_flags, command, label);
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    win32ui::AppendListRowCommands(menu, list);
    const int chosen = win32ui::TrackContextMenu(menu, screen, dialog);
    if (!win32ui::RunListRowCommand(list, static_cast<UINT>(chosen)) && chosen != 0)
    {
        SendMessageW(dialog, WM_COMMAND, MAKEWPARAM(chosen, BN_CLICKED), 0);
    }
    return true;
}

bool FieldRowMenu(void*, HWND list, int, POINT screen)
{
    return ShowEditListMenu(list, screen, {{IDC_DEF_ADD, util::Tr(L"Add &Field...")}, {IDC_DEF_EDIT, util::Tr(L"&Edit Field...")}, {IDC_DEF_REMOVE, util::Tr(L"&Remove Field")}});
}

bool StateRowMenu(void*, HWND list, int, POINT screen)
{
    return ShowEditListMenu(list, screen, {{IDC_FIELD_STATE_ADD, util::Tr(L"Add &State...")}, {IDC_FIELD_STATE_EDIT, util::Tr(L"&Edit State...")}, {IDC_FIELD_STATE_REMOVE, util::Tr(L"&Remove State")}});
}

void EnableRowButtons(HWND dialog, int list_id, int edit_id, int remove_id)
{
    const HWND list = GetDlgItem(dialog, list_id);
    const bool selected = ListView_GetSelectedCount(list) > 0;
    for (const int id : {edit_id, remove_id})
    {
        const HWND button = GetDlgItem(dialog, id);
        if (!selected && GetFocus() == button)
        {
            SendMessageW(dialog, WM_NEXTDLGCTL, reinterpret_cast<WPARAM>(list), TRUE);
        }
        EnableWindow(button, selected);
    }
}

std::vector<unsigned> CheckedBits(HWND dialog)
{
    const HWND list = GetDlgItem(dialog, IDC_FIELD_BITS);
    std::vector<unsigned> bits;
    for (int row = 0, count = ListView_GetItemCount(list); row < count; ++row)
    {
        LPARAM bit = 0;
        if (ListView_GetCheckState(list, row) && win32ui::ListViewItemData(list, row, &bit))
        {
            bits.push_back(static_cast<unsigned>(bit));
        }
    }
    std::sort(bits.begin(), bits.end());
    return bits;
}

void FillStates(HWND dialog, const std::vector<State>& states, std::optional<uint64_t> select)
{
    const HWND list = GetDlgItem(dialog, IDC_FIELD_STATES);
    ListView_DeleteAllItems(list);
    for (size_t i = 0; i < states.size(); ++i)
    {
        const std::wstring value = std::to_wstring(states[i].value);
        const std::wstring meaning = win32ui::SingleLineText(states[i].meaning);
        LVITEMW item = {};
        item.mask = LVIF_TEXT | LVIF_PARAM;
        item.iItem = static_cast<int>(i);
        item.pszText = const_cast<wchar_t*>(value.c_str());
        item.lParam = static_cast<LPARAM>(i);
        const int row = ListView_InsertItem(list, &item);
        ListView_SetItemText(list, row, 1, const_cast<wchar_t*>(states[i].name.c_str()));
        ListView_SetItemText(list, row, 2, const_cast<wchar_t*>(meaning.c_str()));
        if (select == states[i].value)
        {
            win32ui::SelectOnlyListRow(list, row);
        }
    }
    EnableRowButtons(dialog, IDC_FIELD_STATES, IDC_FIELD_STATE_EDIT, IDC_FIELD_STATE_REMOVE);
}

INT_PTR CALLBACK StateDialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<StateEditor*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<StateEditor*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        SetWindowTextW(dialog, state->editing < 0 ? util::Tr(L"Add State") : util::Tr(L"Edit State"));
        if (state->editing >= 0)
        {
            SetDlgItemTextW(dialog, IDC_STATE_VALUE, std::to_wstring(state->state.value).c_str());
        }
        SetDlgItemTextW(dialog, IDC_STATE_NAME, state->state.name.c_str());
        SetDlgItemTextW(dialog, IDC_STATE_MEANING, win32ui::EditLineBreaks(state->state.meaning).c_str());
        win32ui::AnchorDialogControls(dialog, {
                                                  {IDC_STATE_VALUE, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_STATE_NAME, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_STATE_MEANING, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                              });
        return TRUE;
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    if (LOWORD(wparam) == IDCANCEL)
    {
        win32ui::EndDialogWindow(dialog, IDCANCEL);
        return TRUE;
    }
    if (LOWORD(wparam) != IDOK)
    {
        return FALSE;
    }
    State result;
    const std::optional<uint64_t> parsed = win32ui::ParseUnsigned(dialog_support::ReadText(dialog, IDC_STATE_VALUE), 10, true);
    const unsigned long long value = parsed.value_or(0);
    result.name = dialog_support::ReadText(dialog, IDC_STATE_NAME);
    result.meaning = win32ui::PlainLineBreaks(dialog_support::ReadText(dialog, IDC_STATE_MEANING));
    const wchar_t* error = nullptr;
    if (!parsed)
    {
        error = util::Tr(L"Invalid number.");
    }
    else if (value > bitfield::WidthMask(state->bit_count))
    {
        error = util::Tr(L"A state value doesn't fit the bits of its field.");
    }
    else if (std::any_of(state->states->begin(), state->states->end(), [&](const State& other) {
                 return other.value == value && &other - state->states->data() != state->editing;
             }))
    {
        error = util::Tr(L"Two states of one field share the same value.");
    }
    else if (result.name.empty() || result.name.size() > bitfield::kMaxNameLength)
    {
        error = util::Tr(L"Every state needs a name of at most 256 characters.");
    }
    else if (result.meaning.size() > bitfield::kMaxMeaningLength)
    {
        error = util::Tr(L"A state meaning is longer than the format allows.");
    }
    if (error)
    {
        ui::ShowError(dialog, error);
        return TRUE;
    }
    result.value = value;
    state->state = std::move(result);
    win32ui::EndDialogWindow(dialog, IDOK);
    return TRUE;
}

void EditState(HWND dialog, FieldEditor* state, bool create)
{
    std::vector<State>& states = state->field.states;
    const int selected = create ? -1 : SelectedItemIndex(GetDlgItem(dialog, IDC_FIELD_STATES));
    if (!create && (selected < 0 || static_cast<size_t>(selected) >= states.size()))
    {
        return;
    }
    if (create && states.size() >= bitfield::kMaxStates)
    {
        ui::ShowError(dialog, util::Tr(L"A field lists more than 256 states."));
        return;
    }
    StateEditor editor;
    editor.states = &states;
    editor.editing = selected;
    editor.bit_count = static_cast<unsigned>(CheckedBits(dialog).size());
    if (editor.bit_count == 0)
    {
        ui::ShowError(dialog, util::Tr(L"Select at least one bit."));
        return;
    }
    if (!create)
    {
        editor.state = states[static_cast<size_t>(selected)];
    }
    if (win32ui::RunProcDialog(dialog, MAKEINTRESOURCEW(IDD_BITFIELD_STATE), StateDialogProc, reinterpret_cast<LPARAM>(&editor)) != IDOK)
    {
        return;
    }
    if (create)
    {
        states.push_back(editor.state);
    }
    else
    {
        states[static_cast<size_t>(selected)] = editor.state;
    }
    std::stable_sort(states.begin(), states.end(), [](const State& left, const State& right) { return left.value < right.value; });
    FillStates(dialog, states, editor.state.value);
}

void RemoveState(HWND dialog, FieldEditor* state)
{
    const int selected = SelectedItemIndex(GetDlgItem(dialog, IDC_FIELD_STATES));
    if (selected < 0 || static_cast<size_t>(selected) >= state->field.states.size())
    {
        return;
    }
    state->field.states.erase(state->field.states.begin() + selected);
    FillStates(dialog, state->field.states, std::nullopt);
}

constexpr int kBitsMenuCheckAll = 1;
constexpr int kBitsMenuUncheckAll = 2;
constexpr int kBitsMenuCopyMask = 3;

bool BitsRowMenu(void*, HWND list, int row, POINT screen)
{
    const HWND dialog = GetParent(list);
    const auto* state = reinterpret_cast<const FieldEditor*>(GetWindowLongPtrW(dialog, DWLP_USER));
    // empty space menus target the selected bit
    LPARAM bit = 0;
    const bool has_bit = state && win32ui::ListViewItemData(list, row >= 0 ? row : ListView_GetNextItem(list, -1, LVNI_SELECTED), &bit);
    HMENU menu = CreatePopupMenu();
    if (!menu)
    {
        return true;
    }
    AppendMenuW(menu, MF_STRING, kBitsMenuCheckAll, util::Tr(L"Check All &Bits"));
    AppendMenuW(menu, MF_STRING, kBitsMenuUncheckAll, util::Tr(L"&Uncheck All Bits"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING | (has_bit ? 0 : MF_GRAYED), kBitsMenuCopyMask, util::Tr(L"Copy &Mask"));
    win32ui::AppendListRowCommands(menu, list);
    const int chosen = win32ui::TrackContextMenu(menu, screen, dialog);
    if (win32ui::RunListRowCommand(list, static_cast<UINT>(chosen)))
    {
        return true;
    }
    if (chosen == kBitsMenuCheckAll || chosen == kBitsMenuUncheckAll)
    {
        for (int item = ListView_GetItemCount(list) - 1; item >= 0; --item)
        {
            ListView_SetCheckState(list, item, chosen == kBitsMenuCheckAll);
        }
    }
    else if (chosen == kBitsMenuCopyMask && has_bit)
    {
        win32ui::CopyTextToClipboard(dialog, win32ui::FormatHex(1ull << bit, static_cast<int>(state->parent->bit_width / 4)));
    }
    return true;
}

INT_PTR CALLBACK FieldDialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<FieldEditor*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<FieldEditor*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        SetWindowTextW(dialog, state->editing < 0 ? util::Tr(L"Add Field") : util::Tr(L"Edit Field"));
        SetDlgItemTextW(dialog, IDC_FIELD_NAME, state->field.name.c_str());
        SetDlgItemTextW(dialog, IDC_FIELD_MEANING, win32ui::EditLineBreaks(state->field.meaning).c_str());
        const HWND list = GetDlgItem(dialog, IDC_FIELD_BITS);
        static const win32ui::ListViewHooks bits_hooks = win32ui::DialogListHooks(BitsRowMenu);
        static const win32ui::ListViewHooks states_hooks = win32ui::DialogListHooks(StateRowMenu);
        win32ui::SetupDialogList(list, &bits_hooks, LVS_EX_CHECKBOXES, {{util::Tr(L"Bit"), 60}, {util::Tr(L"Mask"), 150}, {util::Tr(L"Field"), 160}});
        // hide bits owned by other fields to prevent overlap
        for (unsigned bit = 0; bit < state->parent->bit_width; ++bit)
        {
            const int owner = state->parent->FieldIndexForBit(bit);
            if (owner >= 0 && owner != state->editing)
            {
                continue;
            }
            const std::wstring text = std::to_wstring(bit);
            LVITEMW item = {};
            item.mask = LVIF_TEXT | LVIF_PARAM;
            item.iItem = ListView_GetItemCount(list);
            item.pszText = const_cast<wchar_t*>(text.c_str());
            item.lParam = static_cast<LPARAM>(bit);
            ListView_InsertItem(list, &item);
            std::wstring mask = win32ui::FormatHex(1ull << bit, static_cast<int>(state->parent->bit_width / 4));
            ListView_SetItemText(list, item.iItem, 1, mask.data());
            if (owner == state->editing && state->editing >= 0)
            {
                ListView_SetItemText(list, item.iItem, 2, const_cast<wchar_t*>(state->field.name.c_str()));
            }
            const bool checked =
                std::find(state->field.bits.begin(), state->field.bits.end(), bit) != state->field.bits.end();
            ListView_SetCheckState(list, item.iItem, checked);
        }
        win32ui::SortListViewItems(list, state->sort_column, false, &state->sort_column, &state->sort_ascending, &state->sort_earlier, CompareFieldBits, state);
        win32ui::SetupDialogList(GetDlgItem(dialog, IDC_FIELD_STATES), &states_hooks, 0, {{util::Tr(L"Value"), 60}, {util::Tr(L"Name"), 120}, {util::Tr(L"Meaning"), 200}});
        FillStates(dialog, state->field.states, std::nullopt);
        win32ui::AnchorDialogControls(dialog, {
                                                  {IDC_FIELD_NAME, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_FIELD_MEANING, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_FIELD_STATES, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_FIELD_STATE_ADD, win32ui::kAnchorLeft | win32ui::kAnchorTop},
                                                  {IDC_FIELD_STATE_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop},
                                                  {IDC_FIELD_STATE_REMOVE, win32ui::kAnchorLeft | win32ui::kAnchorTop},
                                                  {IDC_FIELD_BITS, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                              });
        return TRUE;
    }
    if (message == WM_NOTIFY && state)
    {
        auto* header = reinterpret_cast<NMHDR*>(lparam);
        if (header->idFrom == IDC_FIELD_BITS && header->code == LVN_COLUMNCLICK)
        {
            auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
            win32ui::SortListViewItems(GetDlgItem(dialog, IDC_FIELD_BITS), info->iSubItem, true, &state->sort_column, &state->sort_ascending, &state->sort_earlier, CompareFieldBits, state);
            return TRUE;
        }
        if (header->idFrom == IDC_FIELD_STATES && header->code == NM_DBLCLK)
        {
            EditState(dialog, state, false);
            return TRUE;
        }
        if (header->idFrom == IDC_FIELD_STATES && header->code == LVN_ITEMCHANGED)
        {
            EnableRowButtons(dialog, IDC_FIELD_STATES, IDC_FIELD_STATE_EDIT, IDC_FIELD_STATE_REMOVE);
        }
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    switch (LOWORD(wparam))
    {
    case IDOK:
        {
            Field result;
            result.name = dialog_support::ReadText(dialog, IDC_FIELD_NAME);
            result.meaning = win32ui::PlainLineBreaks(dialog_support::ReadText(dialog, IDC_FIELD_MEANING));
            if (result.name.empty())
            {
                ui::ShowError(dialog, util::Tr(L"Enter a field name."));
                return TRUE;
            }
            for (size_t i = 0; i < state->parent->fields.size(); ++i)
            {
                if (static_cast<int>(i) != state->editing &&
                    util::EqualsInsensitive(state->parent->fields[i].name, result.name))
                {
                    ui::ShowError(dialog, util::Tr(L"Another field already uses that name."));
                    return TRUE;
                }
            }
            result.bits = CheckedBits(dialog);
            if (result.bits.empty())
            {
                ui::ShowError(dialog, util::Tr(L"Select at least one bit."));
                return TRUE;
            }
            const uint64_t limit = bitfield::WidthMask(static_cast<unsigned>(result.bits.size()));
            if (std::any_of(state->field.states.begin(), state->field.states.end(), [limit](const State& entry) { return entry.value > limit; }))
            {
                ui::ShowError(dialog, util::Tr(L"A state value doesn't fit the bits of its field."));
                return TRUE;
            }
            result.states = std::move(state->field.states);
            state->field = std::move(result);
            state->accepted = true;
            win32ui::EndDialogWindow(dialog, IDOK);
            return TRUE;
        }
    case IDC_FIELD_STATE_ADD:
    case IDC_FIELD_STATE_EDIT:
        EditState(dialog, state, LOWORD(wparam) == IDC_FIELD_STATE_ADD);
        return TRUE;
    case IDC_FIELD_STATE_REMOVE:
        RemoveState(dialog, state);
        return TRUE;
    case IDCANCEL:
        win32ui::EndDialogWindow(dialog, IDCANCEL);
        return TRUE;
    default:
        return FALSE;
    }
}

void RefreshFieldList(HWND dialog, Editor* state)
{
    const Definition& definition = state->definition();
    const HWND list = GetDlgItem(dialog, IDC_DEF_LIST);
    const win32ui::RedrawScope redraw(list, RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
    ListView_DeleteAllItems(list);
    for (size_t i = 0; i < definition.fields.size(); ++i)
    {
        const Field& field = definition.fields[i];
        LVITEMW item = {};
        item.mask = LVIF_TEXT | LVIF_PARAM;
        item.iItem = static_cast<int>(i);
        item.pszText = const_cast<wchar_t*>(field.name.c_str());
        item.lParam = static_cast<LPARAM>(i);
        ListView_InsertItem(list, &item);
        const std::wstring bits = BitsText(field);
        ListView_SetItemText(list, item.iItem, 1, const_cast<wchar_t*>(bits.c_str()));
        const std::wstring states = win32ui::SingleLineText(StatesText(field));
        ListView_SetItemText(list, item.iItem, 2, const_cast<wchar_t*>(states.c_str()));
        const std::wstring meaning = win32ui::SingleLineText(field.meaning);
        ListView_SetItemText(list, item.iItem, 3, const_cast<wchar_t*>(meaning.c_str()));
    }
    win32ui::SortListViewItems(list, state->field_sort_column, false, &state->field_sort_column, &state->field_sort_ascending, &state->field_sort_earlier, CompareDefinitionFields, state);
    EnableRowButtons(dialog, IDC_DEF_LIST, IDC_DEF_EDIT, IDC_DEF_REMOVE);
}

unsigned SelectedWidth(HWND dialog)
{
    const int index = static_cast<int>(SendDlgItemMessageW(dialog, IDC_DEF_WIDTH, CB_GETCURSEL, 0, 0));
    if (index < 0 || index >= static_cast<int>(_countof(kWidths)))
    {
        return 32;
    }
    return kWidths[index];
}

void SelectWidth(HWND dialog, unsigned width)
{
    for (int i = 0; i < static_cast<int>(_countof(kWidths)); ++i)
    {
        if (kWidths[i] == width)
        {
            SendDlgItemMessageW(dialog, IDC_DEF_WIDTH, CB_SETCURSEL, static_cast<WPARAM>(i), 0);
            return;
        }
    }
}

bool ReadOffset(HWND dialog, unsigned* offset)
{
    const std::wstring text = dialog_support::ReadText(dialog, IDC_DEF_OFFSET);
    const std::optional<uint64_t> value = text.empty() ? std::optional<uint64_t>(0) : win32ui::ParseUnsigned(text);
    if (!value)
    {
        return false;
    }
    *offset = static_cast<unsigned>(std::min<uint64_t>(*value, bitfield::kMaxByteOffset));
    return true;
}

void ShowDefinition(HWND dialog, Editor* state)
{
    state->loading = true;
    const Definition& definition = state->definition();
    SetDlgItemTextW(dialog, IDC_DEF_VALUE_NAME, definition.value_name.c_str());
    SetDlgItemTextW(dialog, IDC_DEF_KEY_PATHS, util::JoinLines(definition.key_paths).c_str());
    SetDlgItemTextW(dialog, IDC_DEF_COMMENT, win32ui::EditLineBreaks(definition.comment).c_str());
    SetDlgItemInt(dialog, IDC_DEF_OFFSET, definition.byte_offset, FALSE);
    SelectWidth(dialog, definition.bit_width);
    RefreshFieldList(dialog, state);
    state->loading = false;
}

void RefreshDefinitionCombo(HWND dialog, Editor* state)
{
    const HWND combo = GetDlgItem(dialog, IDC_DEF_SELECT);
    const std::wstring filter = state->filtering ? dialog_support::ReadText(dialog, IDC_DEF_SELECT) : std::wstring();
    const DWORD selection = state->filtering ? static_cast<DWORD>(SendMessageW(combo, CB_GETEDITSEL, 0, 0)) : 0;
    // block edit notifications caused by rebuilding the combo
    state->updating_combo = true;
    const win32ui::RedrawScope redraw(combo, RDW_INVALIDATE | RDW_ERASE);
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    bool shown = false;
    for (size_t i = 0; i < state->file.definitions.size(); ++i)
    {
        const std::wstring label = bitfield::DisplayName(state->file.definitions[i]);
        if (static_cast<int>(i) != state->selected && !dialog_support::Matches(label, filter))
        {
            continue;
        }
        const int index =
            static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str())));
        SendMessageW(combo, CB_SETITEMDATA, static_cast<WPARAM>(index), static_cast<LPARAM>(i));
        if (static_cast<int>(i) == state->selected && !state->filtering)
        {
            SendMessageW(combo, CB_SETCURSEL, static_cast<WPARAM>(index), 0);
            shown = true;
        }
    }
    if (!shown && !state->filtering)
    {
        SendMessageW(combo, CB_SETCURSEL, static_cast<WPARAM>(-1), 0);
    }
    if (!state->filtering)
    {
        SendMessageW(combo, CB_SETEDITSEL, 0, MAKELPARAM(-1, 0));
    }
    if (state->filtering)
    {
        SetWindowTextW(combo, filter.c_str());
        SendMessageW(combo, CB_SETEDITSEL, 0, static_cast<LPARAM>(selection));
    }
    win32ui::FitDroppedWidth(combo);
    EnableWindow(GetDlgItem(dialog, IDC_DEF_REMOVE_DEF), !state->single && state->file.definitions.size() > 1);
    state->updating_combo = false;
}

bool CollectDefinition(HWND dialog, Editor* state, Definition* out)
{
    Definition draft = state->definition();
    draft.value_name = dialog_support::ReadText(dialog, IDC_DEF_VALUE_NAME);
    draft.key_paths = util::SplitLines(dialog_support::ReadText(dialog, IDC_DEF_KEY_PATHS));
    draft.comment = win32ui::PlainLineBreaks(dialog_support::ReadText(dialog, IDC_DEF_COMMENT));
    draft.bit_width = SelectedWidth(dialog);
    if (!ReadOffset(dialog, &draft.byte_offset))
    {
        ui::ShowError(dialog, util::Tr(L"Invalid number."));
        return false;
    }
    std::wstring error;
    if (!bitfield::Validate(&draft, &error))
    {
        ui::ShowError(dialog, error);
        return false;
    }
    *out = std::move(draft);
    return true;
}

bool CommitDefinition(HWND dialog, Editor* state)
{
    // validate a copy so rejected edits leave the definition unchanged
    Definition draft;
    if (!CollectDefinition(dialog, state, &draft))
    {
        return false;
    }
    state->definition() = std::move(draft);
    return true;
}

bool SaveToFile(HWND owner, DefinitionFile* file)
{
    std::wstring path;
    const std::wstring suggested = file->definitions.size() == 1
                                       ? bitfield::SuggestedFileName(file->definitions.front().value_name)
                                       : bitfield::SuggestedFileName(file->name);
    if (!ui::ReportFileDialogResult(owner, win32::ChooseFileToSave(owner, bitfield::FileFilter(), suggested.c_str(), &path)))
    {
        return false;
    }
    std::wstring error;
    if (!bitfield::Save(path, *file, &error))
    {
        ui::ShowError(owner, error);
        return false;
    }
    file->path = path;
    return true;
}

bool SaveDefinition(HWND dialog, Editor* state)
{
    if (!CommitDefinition(dialog, state))
    {
        return false;
    }
    if (!SaveToFile(dialog, &state->file))
    {
        return false;
    }
    state->dirty = false;
    state->saved = true;
    return true;
}

bool DiscardChanges(HWND dialog, Editor* state)
{
    if (!state->dirty)
    {
        return true;
    }
    const int choice = ui::PromptChoice(dialog, util::Tr(L"This definition file has unsaved changes."), util::Tr(L"Bit Definitions"), util::Tr(L"Save"), util::Tr(L"Discard"), util::Tr(L"Cancel"), {70, 80, 70});
    if (choice == IDCANCEL)
    {
        return false;
    }
    if (choice == IDYES)
    {
        return SaveDefinition(dialog, state);
    }
    return true;
}

void LoadIntoEditor(HWND dialog, Editor* state, const std::wstring& preset)
{
    if (!DiscardChanges(dialog, state))
    {
        return;
    }
    std::wstring path = preset;
    if (path.empty())
    {
        if (!ui::ReportFileDialogResult(dialog, win32ui::ChooseFileToOpen(dialog, {.filter = bitfield::FileFilter()}, &path)))
        {
            return;
        }
    }
    DefinitionFile loaded;
    std::wstring error;
    if (!bitfield::Load(path, &loaded, &error))
    {
        ui::ShowError(dialog, error);
        return;
    }
    if (state->lock_width && loaded.definitions.front().bit_width != state->definition().bit_width)
    {
        ui::ShowError(dialog, util::Tr(L"The definition was made for a different value width."));
        return;
    }
    if (state->single)
    {
        loaded.definitions.resize(1);
    }
    state->file = std::move(loaded);
    state->selected = 0;
    state->dirty = false;
    state->saved = true;
    RefreshDefinitionCombo(dialog, state);
    ShowDefinition(dialog, state);
}

void ChangeWidth(HWND dialog, Editor* state)
{
    const unsigned width = SelectedWidth(dialog);
    Definition& definition = state->definition();
    bool affected = false;
    for (const Field& field : definition.fields)
    {
        if (!field.bits.empty() && field.bits.back() >= width)
        {
            affected = true;
            break;
        }
    }
    if (affected)
    {
        const int choice = ui::PromptChoice(dialog, util::Tr(L"Some fields use bits outside the new width. Remove those bits and any states or fields that no longer fit?"), util::Tr(L"Bit Definitions"), util::Tr(L"Remove"), util::Tr(L"Cancel"), L"", {80, 70, 70});
        if (choice != IDYES)
        {
            SelectWidth(dialog, definition.bit_width);
            return;
        }
        // new widths can remove bits, oversized states and fields left empty
        for (Field& field : definition.fields)
        {
            std::erase_if(field.bits, [width](unsigned bit) { return bit >= width; });
            const uint64_t limit = bitfield::WidthMask(static_cast<unsigned>(field.bits.size()));
            std::erase_if(field.states, [limit](const bitfield::State& value) { return value.value > limit; });
        }
        definition.fields.erase(std::remove_if(definition.fields.begin(), definition.fields.end(), [](const Field& field) { return field.bits.empty(); }), definition.fields.end());
        RefreshFieldList(dialog, state);
    }
    definition.bit_width = width;
    bitfield::BuildLookup(&definition);
    state->dirty = true;
}

void AddOrEditField(HWND dialog, Editor* state, bool create)
{
    const HWND list = GetDlgItem(dialog, IDC_DEF_LIST);
    const int selected = create ? -1 : SelectedItemIndex(list);
    if (!create && selected < 0)
    {
        return;
    }
    Definition& definition = state->definition();
    Field field;
    if (!create)
    {
        field = definition.fields[static_cast<size_t>(selected)];
    }
    if (!EditBitfieldField(dialog, definition, selected, &field))
    {
        return;
    }
    Definition draft = definition;
    if (create)
    {
        draft.fields.push_back(std::move(field));
    }
    else
    {
        draft.fields[static_cast<size_t>(selected)] = std::move(field);
    }
    std::wstring error;
    if (!bitfield::Validate(&draft, &error))
    {
        ui::ShowError(dialog, error);
        return;
    }
    definition = std::move(draft);
    state->dirty = true;
    RefreshFieldList(dialog, state);
}

void RemoveField(HWND dialog, Editor* state)
{
    const HWND list = GetDlgItem(dialog, IDC_DEF_LIST);
    const int selected = SelectedItemIndex(list);
    if (selected < 0)
    {
        return;
    }
    Definition& definition = state->definition();
    definition.fields.erase(definition.fields.begin() + selected);
    bitfield::BuildLookup(&definition);
    state->dirty = true;
    RefreshFieldList(dialog, state);
}

void AddDefinition(HWND dialog, Editor* state)
{
    if (!CommitDefinition(dialog, state))
    {
        return;
    }
    Definition fresh;
    fresh.owner.fill(-1);
    fresh.bit_width = state->definition().bit_width;
    fresh.name = L"New definition";
    state->file.definitions.push_back(std::move(fresh));
    state->selected = static_cast<int>(state->file.definitions.size()) - 1;
    state->dirty = true;
    RefreshDefinitionCombo(dialog, state);
    ShowDefinition(dialog, state);
}

void RemoveDefinition(HWND dialog, Editor* state)
{
    if (state->file.definitions.size() <= 1)
    {
        return;
    }
    state->file.definitions.erase(state->file.definitions.begin() + state->selected);
    state->selected = std::min<int>(state->selected, static_cast<int>(state->file.definitions.size()) - 1);
    state->dirty = true;
    RefreshDefinitionCombo(dialog, state);
    ShowDefinition(dialog, state);
}

void RenameFromCombo(HWND dialog, Editor* state)
{
    // reuse filter text as a new name when editing ends
    std::wstring text = dialog_support::ReadText(dialog, IDC_DEF_SELECT);
    state->filtering = false;
    if (text.empty() || text == state->definition().name)
    {
        RefreshDefinitionCombo(dialog, state);
        return;
    }
    for (const Definition& other : state->file.definitions)
    {
        if (&other != &state->definition() && other.name == text)
        {
            RefreshDefinitionCombo(dialog, state);
            return;
        }
    }
    state->definition().name = std::move(text);
    state->dirty = true;
    RefreshDefinitionCombo(dialog, state);
}

void SelectDefinition(HWND dialog, Editor* state)
{
    const HWND combo = GetDlgItem(dialog, IDC_DEF_SELECT);
    const int entry = static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0));
    state->filtering = false;
    if (entry < 0)
    {
        return;
    }
    const int index = static_cast<int>(SendMessageW(combo, CB_GETITEMDATA, static_cast<WPARAM>(entry), 0));
    if (index == state->selected || index < 0 || static_cast<size_t>(index) >= state->file.definitions.size())
    {
        return;
    }
    if (!CommitDefinition(dialog, state))
    {
        RefreshDefinitionCombo(dialog, state);
        return;
    }
    state->selected = index;
    ShowDefinition(dialog, state);
}

INT_PTR CALLBACK DefinitionDialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<Editor*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<Editor*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        for (const unsigned width : kWidths)
        {
            SendDlgItemMessageW(dialog, IDC_DEF_WIDTH, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(std::to_wstring(width).c_str()));
        }
        EnableWindow(GetDlgItem(dialog, IDC_DEF_WIDTH), !state->lock_width);
        EnableWindow(GetDlgItem(dialog, IDC_DEF_ADD_DEF), !state->single);
        static const win32ui::ListViewHooks hooks = win32ui::DialogListHooks(FieldRowMenu);
        win32ui::SetupDialogList(GetDlgItem(dialog, IDC_DEF_LIST), &hooks, 0, {{util::Tr(L"Field"), 150}, {util::Tr(L"Bits"), 100}, {util::Tr(L"States"), 180}, {util::Tr(L"Meaning"), 320}});
        RefreshDefinitionCombo(dialog, state);
        ShowDefinition(dialog, state);
        win32ui::AllowLineBreaks(GetDlgItem(dialog, IDC_DEF_KEY_PATHS));
        win32ui::AnchorDialogControls(dialog, {
                                                  {IDC_DEF_SELECT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_DEF_ADD_DEF, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_DEF_REMOVE_DEF, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_DEF_VALUE_NAME, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_DEF_KEY_PATHS, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_DEF_COMMENT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_DEF_LIST, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDC_DEF_ADD, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                  {IDC_DEF_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                  {IDC_DEF_REMOVE, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                  {IDC_DEF_OPEN, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDC_DEF_SAVE_AS, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                              });
        state->dirty = false;
        return TRUE;
    }
    if (message == WM_NOTIFY && state)
    {
        auto* header = reinterpret_cast<NMHDR*>(lparam);
        if (header->idFrom == IDC_DEF_LIST && header->code == LVN_COLUMNCLICK)
        {
            auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
            win32ui::SortListViewItems(GetDlgItem(dialog, IDC_DEF_LIST), info->iSubItem, true, &state->field_sort_column, &state->field_sort_ascending, &state->field_sort_earlier, CompareDefinitionFields, state);
            return TRUE;
        }
        if (header->idFrom == IDC_DEF_LIST && header->code == NM_DBLCLK)
        {
            AddOrEditField(dialog, state, false);
            return TRUE;
        }
        if (header->idFrom == IDC_DEF_LIST && header->code == LVN_ITEMCHANGED)
        {
            EnableRowButtons(dialog, IDC_DEF_LIST, IDC_DEF_EDIT, IDC_DEF_REMOVE);
        }
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    const int id = LOWORD(wparam);
    const int code = HIWORD(wparam);
    if (id == IDC_DEF_SELECT)
    {
        if (code == CBN_EDITCHANGE)
        {
            if (state->updating_combo)
            {
                return TRUE;
            }
            state->filtering = true;
            RefreshDefinitionCombo(dialog, state);
            return TRUE;
        }
        if (code == CBN_SELCHANGE)
        {
            SelectDefinition(dialog, state);
            return TRUE;
        }
        if (code == CBN_KILLFOCUS && state->filtering)
        {
            RenameFromCombo(dialog, state);
            return TRUE;
        }
        return FALSE;
    }
    if (code == EN_CHANGE && !state->loading)
    {
        state->dirty = true;
        return TRUE;
    }
    if (code == CBN_SELCHANGE && id == IDC_DEF_WIDTH)
    {
        ChangeWidth(dialog, state);
        return TRUE;
    }
    switch (id)
    {
    case IDC_DEF_ADD:
        AddOrEditField(dialog, state, true);
        return TRUE;
    case IDC_DEF_EDIT:
        AddOrEditField(dialog, state, false);
        return TRUE;
    case IDC_DEF_REMOVE:
        RemoveField(dialog, state);
        return TRUE;
    case IDC_DEF_ADD_DEF:
        AddDefinition(dialog, state);
        return TRUE;
    case IDC_DEF_REMOVE_DEF:
        RemoveDefinition(dialog, state);
        return TRUE;
    case IDC_DEF_OPEN:
        LoadIntoEditor(dialog, state, std::wstring());
        return TRUE;
    case IDC_DEF_SAVE_AS:
        if (state->filtering)
        {
            RenameFromCombo(dialog, state);
        }
        SaveDefinition(dialog, state);
        return TRUE;
    case IDOK:
        if (state->filtering)
        {
            RenameFromCombo(dialog, state);
        }
        if (!CommitDefinition(dialog, state))
        {
            return TRUE;
        }
        state->accepted = true;
        win32ui::EndDialogWindow(dialog, IDOK);
        return TRUE;
    case IDCANCEL:
        if (DiscardChanges(dialog, state))
        {
            win32ui::EndDialogWindow(dialog, IDCANCEL);
        }
        return TRUE;
    default:
        return FALSE;
    }
}

INT_PTR RunDefinitionDialog(HWND owner, Editor* state)
{
    return win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_BITFIELD_DEFINITION), DefinitionDialogProc, reinterpret_cast<LPARAM>(state));
}

}

bool EditBitfieldField(HWND owner, const Definition& parent, int editing, Field* field)
{
    FieldEditor state;
    state.parent = &parent;
    state.editing = editing;
    state.field = *field;
    const INT_PTR outcome = win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_BITFIELD_FIELD), FieldDialogProc, reinterpret_cast<LPARAM>(&state));
    if (outcome != IDOK || !state.accepted)
    {
        return false;
    }
    *field = std::move(state.field);
    return true;
}

bool EditBitfieldDefinition(HWND owner, Definition* definition, bool lock_width)
{
    if (!definition)
    {
        return false;
    }
    Editor state;
    state.single = true;
    state.lock_width = lock_width;
    state.file.definitions.push_back(*definition);
    Definition& working = state.file.definitions.front();
    if (!bitfield::ValidWidth(working.bit_width))
    {
        working.bit_width = 32;
    }
    bitfield::BuildLookup(&working);
    if (RunDefinitionDialog(owner, &state) != IDOK || !state.accepted)
    {
        return false;
    }
    *definition = std::move(state.file.definitions.front());
    return true;
}

void ShowBitfieldDefinitionEditor(HWND owner, const std::wstring& path)
{
    Editor state;
    if (!path.empty())
    {
        std::wstring error;
        if (!bitfield::Load(path, &state.file, &error))
        {
            ui::ShowError(owner, error);
            return;
        }
        state.saved = true;
    }
    else
    {
        Definition fresh;
        fresh.owner.fill(-1);
        state.file.definitions.push_back(std::move(fresh));
    }
    if (RunDefinitionDialog(owner, &state) == IDOK && state.accepted && (!state.saved || state.dirty))
    {
        SaveToFile(owner, &state.file);
    }
}

}
