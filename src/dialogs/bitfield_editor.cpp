// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/bitfield_editor.h"

#include "registry/registry_path.h"
#include "registry/value_format.h"

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
#include <win32ui/dialog/file_dialog.h>
#include <win32ui/dialog/message_dialog.h>
#include <win32ui/list_view/sort.h>
#include <win32ui/menu/popup_menu.h>
#include <win32ui/theme/theme.h>
#include <win32ui/window.h>
#include <win32ui/window_metrics.h>

#include <commctrl.h>

#include <algorithm>
#include <cstdio>
#include <cstring>
#include <utility>

namespace regkit::editors
{

namespace
{

using bitfield::Definition;
using bitfield::DefinitionFile;
using bitfield::Field;

constexpr int kColumnMask = 1;
constexpr int kColumnState = 2;
constexpr int kColumnField = 3;
constexpr int kColumnValue = 4;
constexpr int kColumnMeaning = 5;

constexpr int kMenuEnable = 100;
constexpr int kMenuDisable = 101;
constexpr int kMenuEditField = 102;
constexpr int kMenuRemoveField = 103;
constexpr int kMenuCopyMask = 104;
constexpr int kMenuSetAllBits = 106;
constexpr int kMenuClearAllBits = 107;
constexpr int kMenuStateBase = 200;

struct Editor
{
    std::wstring value_name;
    std::wstring key_path;
    std::vector<BYTE> bytes;
    bool binary_mode = false;
    bool read_only = false;
    unsigned width = 32;
    unsigned offset = 0;
    uint64_t value = 0;
    bool updating = false;
    bool filtering = false;
    bool updating_combo = false;
    bool accepted = false;
    int sort_column = 0;
    bool sort_ascending = false;
    util::SortKeys sort_earlier;
    std::vector<Definition> choices;
    int choice = 0;

    Definition& definition()
    {
        return choices[static_cast<size_t>(choice)];
    }
    const Definition& definition() const
    {
        return choices[static_cast<size_t>(choice)];
    }
};

const std::wstring& RowMeaning(const Field& field, uint64_t value);

bool BitForRow(HWND list, int row, unsigned* bit)
{
    LPARAM data = 0;
    if (!bit || !win32ui::ListViewItemData(list, row, &data) || data < 0)
    {
        return false;
    }
    *bit = static_cast<unsigned>(data);
    return true;
}

int RowForBit(HWND list, unsigned bit)
{
    return win32ui::FindListViewItemByData(list, static_cast<LPARAM>(bit));
}

int CALLBACK CompareBitRows(LPARAM left_data, LPARAM right_data, int column, void* context)
{
    auto* editor = static_cast<Editor*>(context);
    const unsigned left = static_cast<unsigned>(left_data);
    const unsigned right = static_cast<unsigned>(right_data);
    if (!editor || left >= editor->width || right >= editor->width)
    {
        return 0;
    }
    int result = 0;
    if (column == 0 || column == kColumnMask)
    {
        result = left < right ? -1 : left > right ? 1
                                                  : 0;
    }
    else if (column == kColumnState)
    {
        const bool left_set = ((editor->value >> left) & 1ull) != 0;
        const bool right_set = ((editor->value >> right) & 1ull) != 0;
        result = left_set == right_set ? 0 : left_set ? 1
                                                      : -1;
    }
    else
    {
        const Field* left_field = editor->definition().FieldForBit(left);
        const Field* right_field = editor->definition().FieldForBit(right);
        if (column == kColumnField)
        {
            result = util::CompareListText(left_field ? left_field->name : std::wstring(), right_field ? right_field->name : std::wstring());
        }
        else if (column == kColumnValue)
        {
            const uint64_t left_value = left_field ? left_field->Extract(editor->value) : 0;
            const uint64_t right_value = right_field ? right_field->Extract(editor->value) : 0;
            result = left_value < right_value ? -1 : left_value > right_value ? 1
                                                                              : 0;
        }
        else if (column == kColumnMeaning)
        {
            result = util::CompareListText(left_field ? RowMeaning(*left_field, editor->value) : std::wstring(), right_field ? RowMeaning(*right_field, editor->value) : std::wstring());
        }
    }
    return result;
}

void SortBitRows(HWND list, Editor* editor, int column, bool toggle)
{
    win32ui::SortListViewItems(list, column, toggle, &editor->sort_column, &editor->sort_ascending, &editor->sort_earlier, CompareBitRows, editor);
}

uint64_t ReadWindow(const Editor& editor)
{
    const size_t count = editor.width / 8;
    return editor.offset + count > editor.bytes.size() ? 0 : value_format::ReadUnsigned(std::span(editor.bytes).subspan(editor.offset), count);
}

void WriteWindow(Editor* editor)
{
    const size_t count = editor->width / 8;
    if (editor->offset + count > editor->bytes.size())
    {
        return;
    }
    const std::vector<BYTE> bytes = value_format::UnsignedBytes(editor->value, count);
    std::copy(bytes.begin(), bytes.end(), editor->bytes.begin() + static_cast<std::ptrdiff_t>(editor->offset));
}

std::wstring MaskText(unsigned width, uint64_t mask)
{
    return win32ui::FormatHex(mask, static_cast<int>(width / 4));
}

std::wstring FieldValueText(const Field& field, uint64_t value)
{
    const uint64_t extracted = field.Extract(value);
    std::wstring text = std::to_wstring(extracted);
    if (const bitfield::State* state = field.StateFor(extracted))
    {
        text.append(L" (").append(state->name).append(L")");
    }
    return text;
}

const std::wstring& RowMeaning(const Field& field, uint64_t value)
{
    if (const bitfield::State* state = field.StateFor(field.Extract(value)))
    {
        if (!state->meaning.empty())
        {
            return state->meaning;
        }
    }
    return field.meaning;
}

bool FitsWindow(const Editor& editor, const Definition& definition)
{
    if (!editor.binary_mode)
    {
        return definition.bit_width == editor.width && definition.byte_offset == 0;
    }
    return static_cast<size_t>(definition.byte_offset) + definition.bit_width / 8 <= editor.bytes.size();
}

void UpdateValueText(HWND dialog, const Editor& editor)
{
    std::wstring text = MaskText(editor.width, editor.value);
    if (editor.binary_mode)
    {
        text.append(L"  @ ").append(MaskText(16, editor.offset));
    }
    SetDlgItemTextW(dialog, IDC_VALUE_BYTES, text.c_str());
}

void SetRowText(HWND list, const Editor& editor, int row)
{
    unsigned bit = 0;
    if (!BitForRow(list, row, &bit) || bit >= editor.width)
    {
        return;
    }
    const bool set = (editor.value >> bit) & 1ull;
    ListView_SetItemText(list, row, kColumnState, const_cast<wchar_t*>(set ? L"1" : L"0"));
    const Field* field = editor.definition().FieldForBit(bit);
    if (!field)
    {
        ListView_SetItemText(list, row, kColumnField, const_cast<wchar_t*>(L""));
        ListView_SetItemText(list, row, kColumnValue, const_cast<wchar_t*>(L""));
        ListView_SetItemText(list, row, kColumnMeaning, const_cast<wchar_t*>(L""));
        return;
    }
    const std::wstring value = FieldValueText(*field, editor.value);
    const std::wstring meaning = win32ui::SingleLineText(RowMeaning(*field, editor.value));
    ListView_SetItemText(list, row, kColumnField, const_cast<wchar_t*>(field->name.c_str()));
    ListView_SetItemText(list, row, kColumnValue, const_cast<wchar_t*>(value.c_str()));
    ListView_SetItemText(list, row, kColumnMeaning, const_cast<wchar_t*>(meaning.c_str()));
}

void RefreshFieldRows(HWND dialog, const Editor& editor, const Field& field)
{
    const HWND list = GetDlgItem(dialog, IDC_BITFIELD_LIST);
    for (const unsigned bit : field.bits)
    {
        if (bit < editor.width)
        {
            const int row = RowForBit(list, bit);
            if (row >= 0)
            {
                SetRowText(list, editor, row);
            }
        }
    }
}

void FillList(HWND dialog, Editor* editor)
{
    const HWND list = GetDlgItem(dialog, IDC_BITFIELD_LIST);
    editor->updating = true;
    const win32ui::RedrawScope redraw(list, RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
    ListView_DeleteAllItems(list);
    for (unsigned row = 0; row < editor->width; ++row)
    {
        const unsigned bit = editor->width - 1 - row;
        const std::wstring bit_text = std::to_wstring(bit);
        LVITEMW item = {};
        item.mask = LVIF_TEXT | LVIF_PARAM;
        item.iItem = static_cast<int>(row);
        item.pszText = const_cast<wchar_t*>(bit_text.c_str());
        item.lParam = static_cast<LPARAM>(bit);
        ListView_InsertItem(list, &item);
        const std::wstring mask = MaskText(editor->width, 1ull << bit);
        ListView_SetItemText(list, static_cast<int>(row), kColumnMask, const_cast<wchar_t*>(mask.c_str()));
        SetRowText(list, *editor, static_cast<int>(row));
        if (!editor->read_only)
        {
            ListView_SetCheckState(list, static_cast<int>(row), (editor->value >> bit) & 1ull);
        }
    }
    SortBitRows(list, editor, editor->sort_column, false);
    editor->updating = false;
    UpdateValueText(dialog, *editor);
}

void RefreshCombo(HWND dialog, Editor* editor)
{
    const HWND combo = GetDlgItem(dialog, IDC_BITFIELD_DEFINITION);
    const std::wstring filter =
        editor->filtering ? dialog_support::ReadText(dialog, IDC_BITFIELD_DEFINITION) : std::wstring();
    const DWORD selection = editor->filtering ? static_cast<DWORD>(SendMessageW(combo, CB_GETEDITSEL, 0, 0)) : 0;
    editor->updating_combo = true;
    const win32ui::RedrawScope redraw(combo, RDW_INVALIDATE | RDW_ERASE);
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    for (size_t i = 0; i < editor->choices.size(); ++i)
    {
        const std::wstring label = i == 0 ? std::wstring(util::Tr(L"(none)")) : bitfield::DisplayName(editor->choices[i]);
        if (static_cast<int>(i) != editor->choice && !dialog_support::Matches(label, filter))
        {
            continue;
        }
        const int index =
            static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str())));
        SendMessageW(combo, CB_SETITEMDATA, static_cast<WPARAM>(index), static_cast<LPARAM>(i));
        if (static_cast<int>(i) == editor->choice && !editor->filtering)
        {
            SendMessageW(combo, CB_SETCURSEL, static_cast<WPARAM>(index), 0);
        }
    }
    if (editor->filtering)
    {
        SetWindowTextW(combo, filter.c_str());
        SendMessageW(combo, CB_SETEDITSEL, 0, static_cast<LPARAM>(selection));
    }
    win32ui::FitDroppedWidth(combo);
    editor->updating_combo = false;
}

void SelectRow(HWND dialog, const Editor& editor, unsigned bit)
{
    if (bit >= editor.width)
    {
        return;
    }
    const HWND list = GetDlgItem(dialog, IDC_BITFIELD_LIST);
    const int row = RowForBit(list, bit);
    win32ui::SelectOnlyListRow(list, row);
}

void ApplyChoice(HWND dialog, Editor* editor, int choice)
{
    if (choice < 0 || static_cast<size_t>(choice) >= editor->choices.size())
    {
        return;
    }
    if (editor->binary_mode)
    {
        WriteWindow(editor);
    }
    editor->choice = choice;
    RefreshCombo(dialog, editor);
    const Definition& definition = editor->definition();
    if (editor->binary_mode && choice != 0)
    {
        editor->width = definition.bit_width;
        editor->offset = definition.byte_offset;
        editor->value = ReadWindow(*editor);
    }
    SetDlgItemTextW(dialog, IDC_BITFIELD_COMMENT, win32ui::EditLineBreaks(definition.comment).c_str());
    EnableWindow(GetDlgItem(dialog, IDC_BITFIELD_EDIT_DEF), choice != 0);
    EnableWindow(GetDlgItem(dialog, IDC_BITFIELD_SAVE_AS), choice != 0);
    FillList(dialog, editor);
}

void AddChoice(HWND dialog, Editor* editor, Definition definition)
{
    editor->choices.push_back(std::move(definition));
    ApplyChoice(dialog, editor, static_cast<int>(editor->choices.size()) - 1);
}

bool AcceptDefinition(HWND dialog, const Editor& editor, const Definition& definition)
{
    if (!FitsWindow(editor, definition))
    {
        ui::ShowError(dialog, editor.binary_mode ? util::Tr(L"The definition describes bytes outside this value.") : util::Tr(L"The definition was made for a different value width."));
        return false;
    }
    if (!definition.value_name.empty() && !util::EqualsInsensitive(definition.value_name, editor.value_name))
    {
        std::wstring message = util::Tr(L"This definition was made for a different value. Use it anyway?");
        message.append(L"\r\n\r\n")
            .append(util::TrLabel(L"Definition", definition.value_name))
            .append(L"\r\n")
            .append(util::TrLabel(L"This value", editor.value_name.empty() ? util::Tr(L"(Default)") : editor.value_name));
        if (ui::PromptChoice(dialog, message, util::Tr(L"Bit Definition"), util::Tr(L"Use"), util::Tr(L"Cancel"), L"") != IDYES)
        {
            return false;
        }
    }
    return true;
}

void LoadFromFile(HWND dialog, Editor* editor)
{
    std::wstring path;
    if (!ui::ReportFileDialogResult(dialog, win32ui::ChooseFileToOpen(dialog, {.filter = bitfield::FileFilter()}, &path)))
    {
        return;
    }
    DefinitionFile file;
    std::wstring error;
    if (!bitfield::Load(path, &file, &error))
    {
        ui::ShowError(dialog, error);
        return;
    }
    std::vector<Definition> named;
    std::vector<Definition> fitting;
    for (Definition& definition : file.definitions)
    {
        if (!FitsWindow(*editor, definition))
        {
            continue;
        }
        if (util::EqualsInsensitive(definition.value_name, editor->value_name))
        {
            named.push_back(std::move(definition));
        }
        else
        {
            fitting.push_back(std::move(definition));
        }
    }
    std::vector<Definition>& chosen = named.empty() ? fitting : named;
    if (chosen.empty())
    {
        ui::ShowError(dialog, editor->binary_mode ? util::Tr(L"No definition in that file fits this value.") : util::Tr(L"No definition in that file was made for this value width."));
        return;
    }
    if (named.empty() && !AcceptDefinition(dialog, *editor, chosen.front()))
    {
        return;
    }
    const int first = static_cast<int>(editor->choices.size());
    for (Definition& definition : chosen)
    {
        editor->choices.push_back(std::move(definition));
    }
    ApplyChoice(dialog, editor, first);
}

void SaveCurrent(HWND dialog, Editor* editor)
{
    const Definition& definition = editor->definition();
    std::wstring path;
    const DWORD chosen = win32::ChooseFileToSave(
        dialog,
        bitfield::FileFilter(),
        bitfield::SuggestedFileName(definition.value_name.empty() ? editor->value_name : definition.value_name).c_str(),
        &path
    );
    if (!ui::ReportFileDialogResult(dialog, chosen))
    {
        return;
    }
    DefinitionFile file;
    file.name = bitfield::DisplayName(definition);
    file.definitions.push_back(definition);
    std::wstring error;
    if (!bitfield::Save(path, file, &error))
    {
        ui::ShowError(dialog, error);
    }
}

void EditOrCreate(HWND dialog, Editor* editor, bool create)
{
    Definition definition;
    if (create)
    {
        definition.value_name = editor->value_name;
        definition.bit_width = editor->width;
        definition.byte_offset = editor->offset;
        if (!editor->key_path.empty())
        {
            definition.key_paths.push_back(editor->key_path);
        }
    }
    else
    {
        definition = editor->definition();
    }
    if (!EditBitfieldDefinition(dialog, &definition, !editor->binary_mode))
    {
        return;
    }
    if (!AcceptDefinition(dialog, *editor, definition))
    {
        return;
    }
    if (create)
    {
        AddChoice(dialog, editor, std::move(definition));
        return;
    }
    editor->choices[static_cast<size_t>(editor->choice)] = std::move(definition);
    ApplyChoice(dialog, editor, editor->choice);
}

void SetBit(HWND dialog, Editor* editor, unsigned bit, bool set)
{
    if (editor->read_only || bit >= editor->width)
    {
        return;
    }
    if (set)
    {
        editor->value |= 1ull << bit;
    }
    else
    {
        editor->value &= ~(1ull << bit);
    }
    const HWND list = GetDlgItem(dialog, IDC_BITFIELD_LIST);
    const int row = RowForBit(list, bit);
    if (row < 0)
    {
        return;
    }
    editor->updating = true;
    ListView_SetCheckState(list, row, set);
    editor->updating = false;
    if (const Field* field = editor->definition().FieldForBit(bit))
    {
        RefreshFieldRows(dialog, *editor, *field);
    }
    else
    {
        SetRowText(list, *editor, row);
    }
    UpdateValueText(dialog, *editor);
}

void ApplyState(HWND dialog, Editor* editor, unsigned bit, size_t state_index)
{
    const int index = editor->definition().FieldIndexForBit(bit);
    if (editor->read_only || index < 0)
    {
        return;
    }
    const Field& field = editor->definition().fields[static_cast<size_t>(index)];
    if (state_index >= field.states.size())
    {
        return;
    }
    editor->value = field.Apply(editor->value, field.states[state_index].value);
    const HWND list = GetDlgItem(dialog, IDC_BITFIELD_LIST);
    editor->updating = true;
    for (const unsigned owned : field.bits)
    {
        const int row = RowForBit(list, owned);
        if (row >= 0)
        {
            ListView_SetCheckState(list, row, (editor->value >> owned) & 1ull);
        }
    }
    editor->updating = false;
    RefreshFieldRows(dialog, *editor, field);
    UpdateValueText(dialog, *editor);
}

void EditRowField(HWND dialog, Editor* editor, unsigned bit)
{
    Definition& definition = editor->definition();
    if (definition.value_name.empty() && definition.bit_width != editor->width)
    {
        definition.bit_width = editor->width;
    }
    const int index = definition.FieldIndexForBit(bit);
    Field field;
    if (index >= 0)
    {
        field = definition.fields[static_cast<size_t>(index)];
    }
    else
    {
        field.bits.push_back(bit);
    }
    if (!EditBitfieldField(dialog, definition, index, &field))
    {
        return;
    }
    Definition draft = definition;
    if (index >= 0)
    {
        draft.fields[static_cast<size_t>(index)] = std::move(field);
    }
    else
    {
        draft.fields.push_back(std::move(field));
    }
    std::wstring error;
    if (!bitfield::Validate(&draft, &error))
    {
        ui::ShowError(dialog, error);
        return;
    }
    definition = std::move(draft);
    FillList(dialog, editor);
    SelectRow(dialog, *editor, bit);
}

void RemoveRowField(HWND dialog, Editor* editor, unsigned bit)
{
    Definition& definition = editor->definition();
    const int index = definition.FieldIndexForBit(bit);
    if (index < 0)
    {
        return;
    }
    definition.fields.erase(definition.fields.begin() + index);
    bitfield::BuildLookup(&definition);
    FillList(dialog, editor);
    SelectRow(dialog, *editor, bit);
}

void SetAllBits(HWND dialog, Editor* editor, bool set)
{
    if (editor->read_only)
    {
        return;
    }
    editor->value = set ? bitfield::WidthMask(editor->width) : 0;
    const HWND list = GetDlgItem(dialog, IDC_BITFIELD_LIST);
    editor->updating = true;
    const win32ui::RedrawScope redraw(list, RDW_INVALIDATE | RDW_NOERASE | RDW_ALLCHILDREN);
    for (unsigned row = 0; row < editor->width; ++row)
    {
        ListView_SetCheckState(list, static_cast<int>(row), set);
        SetRowText(list, *editor, static_cast<int>(row));
    }
    editor->updating = false;
    UpdateValueText(dialog, *editor);
}

bool BitRowMenu(void*, HWND list, int row, POINT screen)
{
    const HWND dialog = GetParent(list);
    auto* editor = reinterpret_cast<Editor*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (!editor)
    {
        return false;
    }
    // empty space menus target the selected bit
    unsigned bit = 0;
    const bool has_bit = BitForRow(list, row >= 0 ? row : ListView_GetNextItem(list, -1, LVNI_SELECTED), &bit) && bit < editor->width;
    const bool set = has_bit && ((editor->value >> bit) & 1ull);
    const Field* field = has_bit ? editor->definition().FieldForBit(bit) : nullptr;
    HMENU menu = CreatePopupMenu();
    if (!menu)
    {
        return true;
    }
    const UINT edit_flags = editor->read_only ? MF_GRAYED : 0;
    const UINT bit_flags = has_bit ? 0 : MF_GRAYED;
    AppendMenuW(menu, MF_STRING | edit_flags | bit_flags | (set ? MF_GRAYED : 0), kMenuEnable, util::Tr(L"Enable &Bit"));
    AppendMenuW(menu, MF_STRING | edit_flags | bit_flags | (set ? 0 : MF_GRAYED), kMenuDisable, util::Tr(L"&Disable Bit"));
    if (field && !field->states.empty())
    {
        HMENU states = CreatePopupMenu();
        const uint64_t current = field->Extract(editor->value);
        for (size_t i = 0; i < field->states.size(); ++i)
        {
            const bitfield::State& state = field->states[i];
            std::wstring label = win32ui::MenuName(state.name) + L"   (" + std::to_wstring(state.value) + L")";
            AppendMenuW(states, MF_STRING | edit_flags | (state.value == current ? MF_CHECKED : 0), kMenuStateBase + i, label.c_str());
        }
        AppendMenuW(menu, MF_POPUP | edit_flags, reinterpret_cast<UINT_PTR>(states), util::Tr(L"Set &Field To"));
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING | edit_flags, kMenuSetAllBits, util::Tr(L"&Set All Bits"));
    AppendMenuW(menu, MF_STRING | edit_flags, kMenuClearAllBits, util::Tr(L"C&lear All Bits"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING | bit_flags, kMenuEditField, field ? util::Tr(L"&Edit Field...") : util::Tr(L"Describe Bit..."));
    AppendMenuW(menu, MF_STRING | (field ? 0 : MF_GRAYED), kMenuRemoveField, util::Tr(L"&Remove Field"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING | bit_flags, kMenuCopyMask, util::Tr(L"Copy &Mask"));
    win32ui::AppendListRowCommands(menu, list);

    const int chosen = win32ui::TrackContextMenu(menu, screen, dialog);
    if (win32ui::RunListRowCommand(list, static_cast<UINT>(chosen)))
    {
        return true;
    }
    switch (chosen)
    {
    case kMenuSetAllBits:
    case kMenuClearAllBits:
        SetAllBits(dialog, editor, chosen == kMenuSetAllBits);
        return true;
    case kMenuEnable:
        SetBit(dialog, editor, bit, true);
        return true;
    case kMenuDisable:
        SetBit(dialog, editor, bit, false);
        return true;
    case kMenuEditField:
        EditRowField(dialog, editor, bit);
        return true;
    case kMenuRemoveField:
        RemoveRowField(dialog, editor, bit);
        return true;
    case kMenuCopyMask:
        win32ui::CopyTextToClipboard(dialog, MaskText(editor->width, 1ull << bit));
        return true;
    default:
        break;
    }
    if (chosen >= kMenuStateBase)
    {
        ApplyState(dialog, editor, bit, static_cast<size_t>(chosen - kMenuStateBase));
    }
    return true;
}

void HandleItemChanged(HWND dialog, Editor* editor, const NMLISTVIEW* info)
{
    if (editor->updating || !(info->uChanged & LVIF_STATE))
    {
        return;
    }
    const UINT old_image = info->uOldState & LVIS_STATEIMAGEMASK;
    const UINT new_image = info->uNewState & LVIS_STATEIMAGEMASK;
    if (old_image == 0 || new_image == 0 || old_image == new_image)
    {
        return;
    }
    unsigned bit = 0;
    const HWND list = GetDlgItem(dialog, IDC_BITFIELD_LIST);
    if (!BitForRow(list, info->iItem, &bit) || bit >= editor->width)
    {
        return;
    }
    const bool set = new_image == INDEXTOSTATEIMAGEMASK(2);
    if (set)
    {
        editor->value |= 1ull << bit;
    }
    else
    {
        editor->value &= ~(1ull << bit);
    }
    if (const Field* field = editor->definition().FieldForBit(bit))
    {
        RefreshFieldRows(dialog, *editor, *field);
    }
    else
    {
        SetRowText(list, *editor, info->iItem);
    }
    UpdateValueText(dialog, *editor);
}

void PopulateChoices(HWND dialog, Editor* editor)
{
    Definition none;
    none.owner.fill(-1);
    none.bit_width = editor->width;
    editor->choices.push_back(std::move(none));
    for (Definition& definition : bitfield::Matching(editor->key_path, editor->value_name))
    {
        if (FitsWindow(*editor, definition))
        {
            editor->choices.push_back(std::move(definition));
        }
    }
    RefreshCombo(dialog, editor);
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* editor = reinterpret_cast<Editor*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        editor = reinterpret_cast<Editor*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(editor));
        SetDlgItemTextW(dialog, IDC_VALUE_NAME, editor->value_name.empty() ? util::Tr(L"(Default)") : registry_path::DisplayName(editor->value_name).c_str());
        SendDlgItemMessageW(dialog, IDC_VALUE_NAME, EM_SETREADONLY, TRUE, 0);
        SendDlgItemMessageW(dialog, IDC_BITFIELD_COMMENT, EM_SETREADONLY, TRUE, 0);
        const HWND name = GetDlgItem(dialog, IDC_VALUE_NAME);
        SetWindowLongPtrW(name, GWL_STYLE, GetWindowLongPtrW(name, GWL_STYLE) & ~WS_TABSTOP);
        if (editor->read_only)
        {
            EnableWindow(GetDlgItem(dialog, IDC_BITFIELD_SET_ALL), FALSE);
            EnableWindow(GetDlgItem(dialog, IDC_BITFIELD_CLEAR_ALL), FALSE);
            EnableWindow(GetDlgItem(dialog, IDOK), FALSE);
        }
        static const win32ui::ListViewHooks hooks = win32ui::DialogListHooks(BitRowMenu);
        win32ui::SetupDialogList(
            GetDlgItem(dialog, IDC_BITFIELD_LIST),
            &hooks,
            editor->read_only ? 0u : static_cast<DWORD>(LVS_EX_CHECKBOXES),
            {{util::Tr(L"Bit"), 44}, {util::Tr(L"Mask"), 150}, {util::Tr(L"State"), 62}, {util::Tr(L"Field"), 160}, {util::Tr(L"Value"), 110}, {util::Tr(L"Meaning"), 420}}
        );
        PopulateChoices(dialog, editor);
        win32ui::AnchorDialogControls(dialog, {
                                                  {IDC_VALUE_NAME, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_VALUE_BYTES_LABEL, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_VALUE_BYTES, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_BITFIELD_DEFINITION, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_BITFIELD_LOAD, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_BITFIELD_NEW, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_BITFIELD_EDIT_DEF, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_BITFIELD_SAVE_AS, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_BITFIELD_COMMENT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_BITFIELD_LIST, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDC_BITFIELD_SET_ALL, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                  {IDC_BITFIELD_CLEAR_ALL, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                  {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                              });
        ApplyChoice(dialog, editor, editor->choices.size() > 1 ? 1 : 0);
        SetFocus(GetDlgItem(dialog, IDC_BITFIELD_LIST));
        return FALSE;
    }
    if (message == WM_NOTIFY && editor)
    {
        auto* header = reinterpret_cast<NMHDR*>(lparam);
        if (header->idFrom == IDC_BITFIELD_LIST && header->code == LVN_COLUMNCLICK)
        {
            auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
            SortBitRows(GetDlgItem(dialog, IDC_BITFIELD_LIST), editor, info->iSubItem, true);
            return TRUE;
        }
        if (header->idFrom == IDC_BITFIELD_LIST && header->code == LVN_ITEMCHANGED)
        {
            HandleItemChanged(dialog, editor, reinterpret_cast<NMLISTVIEW*>(lparam));
            return TRUE;
        }
    }
    if (message != WM_COMMAND || !editor)
    {
        return FALSE;
    }
    const int id = LOWORD(wparam);
    if (id == IDC_BITFIELD_DEFINITION)
    {
        const HWND combo = GetDlgItem(dialog, IDC_BITFIELD_DEFINITION);
        if (HIWORD(wparam) == CBN_EDITCHANGE)
        {
            if (editor->updating_combo)
            {
                return TRUE;
            }
            editor->filtering = true;
            RefreshCombo(dialog, editor);
            return TRUE;
        }
        if (HIWORD(wparam) == CBN_SELCHANGE)
        {
            const int index = static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0));
            editor->filtering = false;
            if (index >= 0)
            {
                ApplyChoice(dialog, editor, static_cast<int>(SendMessageW(combo, CB_GETITEMDATA, static_cast<WPARAM>(index), 0)));
            }
            return TRUE;
        }
        if (HIWORD(wparam) == CBN_KILLFOCUS && editor->filtering)
        {
            editor->filtering = false;
            RefreshCombo(dialog, editor);
            return TRUE;
        }
        return FALSE;
    }
    switch (id)
    {
    case IDC_BITFIELD_LOAD:
        LoadFromFile(dialog, editor);
        return TRUE;
    case IDC_BITFIELD_NEW:
        EditOrCreate(dialog, editor, true);
        return TRUE;
    case IDC_BITFIELD_EDIT_DEF:
        EditOrCreate(dialog, editor, false);
        return TRUE;
    case IDC_BITFIELD_SAVE_AS:
        SaveCurrent(dialog, editor);
        return TRUE;
    case IDC_BITFIELD_SET_ALL:
        SetAllBits(dialog, editor, true);
        return TRUE;
    case IDC_BITFIELD_CLEAR_ALL:
        SetAllBits(dialog, editor, false);
        return TRUE;
    case IDOK:
        if (!editor->read_only)
        {
            WriteWindow(editor);
            editor->accepted = true;
            win32ui::EndDialogWindow(dialog, IDOK);
        }
        return TRUE;
    case IDCANCEL:
        win32ui::EndDialogWindow(dialog, IDCANCEL);
        return TRUE;
    default:
        return FALSE;
    }
}

}

bool EditBitfield(HWND owner, const BitfieldRequest& request, BitfieldResult* result)
{
    if (!result)
    {
        return false;
    }
    Editor editor;
    editor.value_name = request.value_name;
    editor.key_path = request.key_path;
    editor.read_only = request.read_only;
    editor.binary_mode = !request.data.empty();
    if (editor.binary_mode)
    {
        editor.bytes.assign(request.data.begin(), request.data.end());
        const size_t usable = std::min<size_t>(editor.bytes.size(), 4);
        editor.width = usable >= 4 ? 32 : (usable >= 2 ? 16 : 8);
        editor.value = ReadWindow(editor);
    }
    else
    {
        if (!bitfield::ValidWidth(request.bit_count))
        {
            return false;
        }
        editor.width = request.bit_count;
        editor.value = request.value & bitfield::WidthMask(editor.width);
        editor.bytes.resize(editor.width / 8);
        WriteWindow(&editor);
    }

    const INT_PTR outcome = win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_BITFIELD), DialogProc, reinterpret_cast<LPARAM>(&editor));
    if (outcome != IDOK || !editor.accepted)
    {
        return false;
    }
    result->value = editor.value;
    result->data = std::move(editor.bytes);
    return true;
}

}
