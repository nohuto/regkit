// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/value_editor.h"
#include "win32/text_transform.h"

#include "dialogs/bitfield_editor.h"
#include "ui/dialog_layout.h"

#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/edit.h>
#include <win32ui/gdi/fonts.h>

#include <algorithm>
#include <cerrno>
#include <commctrl.h>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <initializer_list>
#include <limits>

#include "ui/feedback.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/translation.h"
#include "resource.h"

namespace regkit::editors
{

namespace
{

bool IsMultilineEdit(HWND dialog, int id)
{
    HWND edit = GetDlgItem(dialog, id);
    return edit && (GetWindowLongPtrW(edit, GWL_STYLE) & ES_MULTILINE) != 0;
}

struct BinaryGroupState
{
    int group_bytes = 1;
    bool unicode = false;
    bool updating = false;
};

struct TraceValueDialogState
{
    std::wstring value_name;
    std::wstring key_path;
    DWORD type = REG_SZ;
    std::vector<BYTE> data;
    std::wstring shown_text;
    int shown_dword_base = 16;
    int shown_qword_base = 16;
    bool accepted = false;
    bool read_only = false;
    int dword_base = 16;
    int qword_base = 16;
    BinaryGroupState binary;
    BinaryGroupState none;
    win32ui::UniqueFont mono_font;
};

struct ExtendedValueDialogState
{
    DWORD base_type = REG_SZ;
    std::wstring value_name;
    std::wstring key_path;
    std::wstring initial_text;
    std::vector<BYTE> initial_data;
    std::vector<BYTE> data;
    bool accepted = false;
    bool read_only = false;
    int number_base = 16;
    int initial_number_base = 16;
};

std::wstring RegDataToString(const std::vector<BYTE>& data);
int TraceEditorId(DWORD type);

void ConfigureReadOnlyNameField(HWND dlg, const std::wstring& name)
{
    HWND name_value = GetDlgItem(dlg, IDC_VALUE_NAME);
    if (!name_value)
    {
        return;
    }
    SetWindowTextW(name_value, name.empty() ? util::Tr(L"(Default)") : registry_path::DisplayName(name).c_str());
    ShowWindow(name_value, SW_SHOW);
    SendMessageW(name_value, EM_SETREADONLY, TRUE, 0);
    LONG_PTR style = GetWindowLongPtrW(name_value, GWL_STYLE);
    if (style & WS_TABSTOP)
    {
        style &= ~WS_TABSTOP;
        SetWindowLongPtrW(name_value, GWL_STYLE, style);
    }
}

void ConfigureReadOnlyControls(HWND dialog, std::initializer_list<int> edit_ids, std::initializer_list<int> control_ids)
{
    for (const int id : edit_ids)
    {
        SendDlgItemMessageW(dialog, id, EM_SETREADONLY, TRUE, 0);
    }
    for (const int id : control_ids)
    {
        EnableWindow(GetDlgItem(dialog, id), FALSE);
    }
}

struct BinaryGroupIds
{
    int edit_id = 0;
    int preview_id = 0;
    int format_byte_id = 0;
    int format_word_id = 0;
    int format_dword_id = 0;
    int format_qword_id = 0;
    int text_ansi_id = 0;
    int text_unicode_id = 0;
};

const BinaryGroupIds kBinaryIds = {IDC_REG_BINARY_EDIT, IDC_REG_BINARY_PREVIEW, IDC_REG_BINARY_FORMAT_BYTE, IDC_REG_BINARY_FORMAT_WORD, IDC_REG_BINARY_FORMAT_DWORD, IDC_REG_BINARY_FORMAT_QWORD, IDC_REG_BINARY_TEXT_ANSI, IDC_REG_BINARY_TEXT_UNICODE};
const BinaryGroupIds kNoneIds = {IDC_REG_NONE_EDIT, IDC_REG_NONE_PREVIEW, IDC_REG_NONE_FORMAT_BYTE, IDC_REG_NONE_FORMAT_WORD, IDC_REG_NONE_FORMAT_DWORD, IDC_REG_NONE_FORMAT_QWORD, IDC_REG_NONE_TEXT_ANSI, IDC_REG_NONE_TEXT_UNICODE};

int TypeToComboIndex(DWORD type)
{
    const auto types = value_format::TypeLabels();
    const auto found = std::find_if(types.begin(), types.end(), [&](const value_format::TypeLabel& label) { return label.type == type; });
    // no selection keeps an unlisted type until the user picks one
    return found == types.end() ? -1 : static_cast<int>(found - types.begin());
}

DWORD ComboIndexToType(int index)
{
    const auto types = value_format::TypeLabels();
    return index >= 0 && static_cast<size_t>(index) < types.size() ? types[static_cast<size_t>(index)].type : REG_SZ;
}

void PopulateTraceTypeCombo(HWND dlg)
{
    HWND combo = GetDlgItem(dlg, IDC_TYPE_COMBO);
    if (!combo)
    {
        return;
    }
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    for (const auto& entry : value_format::TypeLabels())
    {
        int idx = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(entry.name)));
        SendMessageW(combo, CB_SETITEMDATA, idx, static_cast<LPARAM>(entry.type));
    }
}

void SetGroupVisibility(HWND dlg, const int* ids, size_t count, bool visible)
{
    if (!dlg || !ids)
    {
        return;
    }
    int cmd = visible ? SW_SHOW : SW_HIDE;
    for (size_t i = 0; i < count; ++i)
    {
        HWND ctrl = GetDlgItem(dlg, ids[i]);
        if (ctrl)
        {
            ShowWindow(ctrl, cmd);
        }
    }
}

const int kRegSzGroupIds[] = {IDC_GROUP_REG_SZ, IDC_REG_SZ_EDIT};
const int kRegExpandGroupIds[] = {IDC_GROUP_REG_EXPAND, IDC_REG_EXPAND_EDIT};
const int kRegMultiGroupIds[] = {IDC_GROUP_REG_MULTI, IDC_REG_MULTI_EDIT};
const int kRegDwordGroupIds[] = {IDC_GROUP_REG_DWORD, IDC_REG_DWORD_EDIT, IDC_REG_DWORD_BITS, IDC_REG_DWORD_BASE_GROUP, IDC_REG_DWORD_HEX, IDC_REG_DWORD_DEC, IDC_REG_DWORD_BIN};
const int kRegQwordGroupIds[] = {IDC_GROUP_REG_QWORD, IDC_REG_QWORD_EDIT, IDC_REG_QWORD_BITS, IDC_REG_QWORD_BASE_GROUP, IDC_REG_QWORD_HEX, IDC_REG_QWORD_DEC, IDC_REG_QWORD_BIN};
const int kRegBinaryGroupIds[] = {IDC_GROUP_REG_BINARY, IDC_REG_BINARY_LABEL_HEX, IDC_REG_BINARY_EDIT, IDC_REG_BINARY_LABEL_PREVIEW, IDC_REG_BINARY_PREVIEW, IDC_REG_BINARY_FORMAT_GROUP, IDC_REG_BINARY_FORMAT_BYTE, IDC_REG_BINARY_FORMAT_WORD, IDC_REG_BINARY_FORMAT_DWORD, IDC_REG_BINARY_FORMAT_QWORD, IDC_REG_BINARY_TEXT_GROUP, IDC_REG_BINARY_TEXT_ANSI, IDC_REG_BINARY_TEXT_UNICODE};
const int kRegNoneGroupIds[] = {IDC_GROUP_REG_NONE, IDC_REG_NONE_LABEL_HEX, IDC_REG_NONE_EDIT, IDC_REG_NONE_LABEL_PREVIEW, IDC_REG_NONE_PREVIEW, IDC_REG_NONE_FORMAT_GROUP, IDC_REG_NONE_FORMAT_BYTE, IDC_REG_NONE_FORMAT_WORD, IDC_REG_NONE_FORMAT_DWORD, IDC_REG_NONE_FORMAT_QWORD, IDC_REG_NONE_TEXT_GROUP, IDC_REG_NONE_TEXT_ANSI, IDC_REG_NONE_TEXT_UNICODE};

struct TraceTypeGroup
{
    DWORD type = REG_SZ;
    const int* ids = nullptr;
    size_t count = 0;
};

const TraceTypeGroup kTraceTypeGroups[] = {
    {REG_SZ, kRegSzGroupIds, _countof(kRegSzGroupIds)},
    {REG_EXPAND_SZ, kRegExpandGroupIds, _countof(kRegExpandGroupIds)},
    {REG_MULTI_SZ, kRegMultiGroupIds, _countof(kRegMultiGroupIds)},
    {REG_LINK, kRegSzGroupIds, _countof(kRegSzGroupIds)},
    {REG_DWORD, kRegDwordGroupIds, _countof(kRegDwordGroupIds)},
    {REG_DWORD_BIG_ENDIAN, kRegDwordGroupIds, _countof(kRegDwordGroupIds)},
    {REG_QWORD, kRegQwordGroupIds, _countof(kRegQwordGroupIds)},
    {REG_BINARY, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_RESOURCE_LIST, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_FULL_RESOURCE_DESCRIPTOR, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_RESOURCE_REQUIREMENTS_LIST, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_NONE, kRegNoneGroupIds, _countof(kRegNoneGroupIds)},
};

const wchar_t* TraceTypeLabel(DWORD type)
{
    for (const auto& entry : value_format::TypeLabels())
    {
        if (entry.type == type)
        {
            return entry.name;
        }
    }
    return L"REG_BINARY";
}

bool IsBinaryGroupType(DWORD type)
{
    switch (type)
    {
    case REG_BINARY:
    case REG_RESOURCE_LIST:
    case REG_FULL_RESOURCE_DESCRIPTOR:
    case REG_RESOURCE_REQUIREMENTS_LIST:
        return true;
    default:
        return false;
    }
}

void ShowTraceTypeGroup(HWND dlg, DWORD type)
{
    for (const auto& group : kTraceTypeGroups)
    {
        SetGroupVisibility(dlg, group.ids, group.count, false);
    }
    if (IsBinaryGroupType(type))
    {
        SetGroupVisibility(dlg, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds), true);
        return;
    }
    for (const auto& group : kTraceTypeGroups)
    {
        if (group.type == type)
        {
            SetGroupVisibility(dlg, group.ids, group.count, true);
            break;
        }
    }
}

void UpdateTraceGroupLabels(HWND dlg, DWORD type)
{
    if (!dlg)
    {
        return;
    }
    const wchar_t* sz_label = (type == REG_LINK) ? L"REG_LINK" : L"REG_SZ";
    const wchar_t* dword_label = (type == REG_DWORD_BIG_ENDIAN) ? L"REG_DWORD_BIG_ENDIAN" : L"REG_DWORD";
    const wchar_t* raw_label = IsBinaryGroupType(type) ? TraceTypeLabel(type) : L"REG_BINARY";
    SetDlgItemTextW(dlg, IDC_GROUP_REG_SZ, sz_label);
    SetDlgItemTextW(dlg, IDC_GROUP_REG_DWORD, dword_label);
    SetDlgItemTextW(dlg, IDC_GROUP_REG_BINARY, raw_label);
}

void SelectTraceType(HWND dlg, TraceValueDialogState* state, DWORD type)
{
    if (!dlg || !state)
    {
        return;
    }
    state->type = type;
    HWND combo = GetDlgItem(dlg, IDC_TYPE_COMBO);
    if (combo)
    {
        int index = TypeToComboIndex(type);
        SendMessageW(combo, CB_SETCURSEL, index, 0);
    }
    ShowTraceTypeGroup(dlg, type);
    UpdateTraceGroupLabels(dlg, type);
}

DWORD ReadTraceType(HWND dlg, TraceValueDialogState* state)
{
    if (!dlg)
    {
        return state ? state->type : REG_SZ;
    }
    HWND combo = GetDlgItem(dlg, IDC_TYPE_COMBO);
    if (!combo)
    {
        return state ? state->type : REG_SZ;
    }
    int index = static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0));
    if (index == CB_ERR)
    {
        return state ? state->type : REG_SZ;
    }
    return ComboIndexToType(index);
}

void SetBinaryGroupSelection(HWND dlg, const BinaryGroupIds& ids, int control_id)
{
    if (!dlg)
    {
        return;
    }
    CheckDlgButton(dlg, ids.format_byte_id, control_id == ids.format_byte_id ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dlg, ids.format_word_id, control_id == ids.format_word_id ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dlg, ids.format_dword_id, control_id == ids.format_dword_id ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dlg, ids.format_qword_id, control_id == ids.format_qword_id ? BST_CHECKED : BST_UNCHECKED);
}

void SetBinaryTextSelection(HWND dlg, const BinaryGroupIds& ids, int control_id)
{
    if (!dlg)
    {
        return;
    }
    CheckDlgButton(dlg, ids.text_ansi_id, control_id == ids.text_ansi_id ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dlg, ids.text_unicode_id, control_id == ids.text_unicode_id ? BST_CHECKED : BST_UNCHECKED);
}

void UpdateBinaryPreviewEx(HWND dlg, BinaryGroupState* state, const BinaryGroupIds& ids)
{
    if (!dlg || !state || state->updating)
    {
        return;
    }
    std::wstring text = util::DialogItemText(dlg, ids.edit_id);
    std::vector<BYTE> parsed;
    if (!win32ui::ParseHexBytes(text, &parsed))
    {
        SetDlgItemTextW(dlg, ids.preview_id, util::Tr(L"Invalid hex input."));
        return;
    }
    std::wstring preview = win32ui::HexDump(parsed, state->group_bytes, state->unicode);
    SetDlgItemTextW(dlg, ids.preview_id, preview.c_str());
}

void SelectBinaryMode(HWND dlg, BinaryGroupState* state, const BinaryGroupIds& ids, int control_id)
{
    if (control_id == ids.text_ansi_id || control_id == ids.text_unicode_id)
    {
        const bool unicode = control_id == ids.text_unicode_id;
        if (state->unicode == unicode)
        {
            return;
        }
        state->unicode = unicode;
        SetBinaryTextSelection(dlg, ids, control_id);
    }
    else
    {
        const int group = control_id == ids.format_word_id ? 2 : control_id == ids.format_dword_id ? 4
                                                             : control_id == ids.format_qword_id   ? 8
                                                                                                   : 1;
        if (state->group_bytes == group)
        {
            return;
        }
        state->group_bytes = group;
        SetBinaryGroupSelection(dlg, ids, control_id);
    }
    UpdateBinaryPreviewEx(dlg, state, ids);
}

std::wstring FormatNumberValue(unsigned long long value, int base)
{
    wchar_t buffer[64] = {};
    if (base == 16)
    {
        swprintf_s(buffer, L"%llX", value);
        return buffer;
    }
    if (base == 10)
    {
        swprintf_s(buffer, L"%llu", value);
        return buffer;
    }
    if (base == 2)
    {
        if (value == 0)
        {
            return L"0";
        }
        std::wstring out;
        while (value > 0)
        {
            out.push_back((value & 1ULL) ? L'1' : L'0');
            value >>= 1;
        }
        std::reverse(out.begin(), out.end());
        return out;
    }
    swprintf_s(buffer, L"%llu", value);
    return buffer;
}

unsigned long long ReadNumberWithFallback(HWND dlg, int edit_id, int base, unsigned long long fallback)
{
    // keep base switching usable when current text is incomplete
    return win32ui::ParseUnsigned(util::DialogItemText(dlg, edit_id), base, true).value_or(fallback);
}

void RunBitfieldEditor(HWND dlg, const std::wstring& value_name, const std::wstring& key_path, int edit_id, int base, unsigned bit_count, bool read_only)
{
    const std::optional<uint64_t> value = win32ui::ParseUnsigned(util::DialogItemText(dlg, edit_id), base, true);
    if (!value)
    {
        ui::ShowError(dlg, util::Tr(L"Enter a valid number before editing its bits."));
        return;
    }
    BitfieldRequest request;
    request.value_name = value_name;
    request.key_path = key_path;
    request.value = *value;
    request.bit_count = bit_count;
    request.read_only = read_only;
    BitfieldResult result;
    if (!EditBitfield(dlg, request, &result))
    {
        return;
    }
    SetDlgItemTextW(dlg, edit_id, FormatNumberValue(result.value, base).c_str());
    SendDlgItemMessageW(dlg, edit_id, EM_SETSEL, 0, -1);
}

bool ConvertValueData(DWORD from, const std::vector<BYTE>& data, DWORD to, std::vector<BYTE>* out)
{
    if (!out || from == to)
    {
        return false;
    }
    const DWORD src = value_format::NormalizeType(from);
    const DWORD dst = value_format::NormalizeType(to);
    auto is_text = [](DWORD type) { return type == REG_SZ || type == REG_EXPAND_SZ || type == REG_LINK; };
    auto is_number = [](DWORD type) { return type == REG_DWORD || type == REG_DWORD_BIG_ENDIAN || type == REG_QWORD; };
    auto is_raw = [](DWORD type) { return type == REG_BINARY || type == REG_NONE; };
    auto is_readable = [](const std::wstring& text) {
        return !text.empty() && std::none_of(text.begin(), text.end(), [](wchar_t ch) {
            return iswcntrl(ch) && ch != L'\t' && ch != L'\r' && ch != L'\n';
        });
    };

    if (is_text(src) && is_text(dst))
    {
        *out = data;
        return true;
    }
    if (is_text(src) && dst == REG_MULTI_SZ)
    {
        *out = value_format::MultiStringData(RegDataToString(data));
        return true;
    }
    if (src == REG_MULTI_SZ && is_text(dst))
    {
        const std::vector<std::wstring> items = value_format::MultiStringItems(data);
        // converting multiple strings to one would discard item boundaries
        if (items.size() != 1)
        {
            return false;
        }
        *out = value_format::StringData(items.front());
        return true;
    }
    if (is_number(src) && is_number(dst))
    {
        const size_t src_size = src == REG_QWORD ? sizeof(unsigned long long) : sizeof(DWORD);
        // malformed sizes have no number to keep
        if (data.size() != src_size)
        {
            return false;
        }
        const unsigned long long value = src == REG_DWORD_BIG_ENDIAN ? value_format::ReadUnsigned(data, src_size, true)
                                                                     : value_format::ReadUnsigned(data, src_size);
        // reject narrowing conversions that would truncate the value
        if (dst != REG_QWORD && value > std::numeric_limits<DWORD>::max())
        {
            return false;
        }
        if (dst == REG_DWORD_BIG_ENDIAN)
        {
            *out = value_format::UnsignedBytes(value, sizeof(DWORD), true);
            return true;
        }
        const size_t dst_size = dst == REG_QWORD ? sizeof(unsigned long long) : sizeof(DWORD);
        out->assign(dst_size, 0);
        std::memcpy(out->data(), &value, dst_size);
        return true;
    }
    if (is_raw(dst))
    {
        *out = data;
        return true;
    }
    if (is_raw(src) && is_number(dst))
    {
        const size_t dst_size = dst == REG_QWORD ? sizeof(unsigned long long) : sizeof(DWORD);
        // raw bytes must already match the target integer width
        if (data.size() != dst_size)
        {
            return false;
        }
        *out = data;
        return true;
    }
    if (is_raw(src) && dst == REG_MULTI_SZ)
    {
        const std::vector<std::wstring> items = value_format::MultiStringItems(data);
        if (items.empty() || !std::all_of(items.begin(), items.end(), is_readable))
        {
            return false;
        }
        *out = data;
        return true;
    }
    if (is_raw(src) && is_text(dst))
    {
        std::wstring text;
        if (!value_format::DecodeString(data, &text) || !is_readable(text))
        {
            return false;
        }
        *out = data;
        return true;
    }
    return false;
}

void PopulateTraceValueEditors(HWND dlg, TraceValueDialogState* state)
{
    if (!dlg || !state)
    {
        return;
    }
    const std::vector<BYTE>& data = state->data;
    const bool filled = !data.empty();
    switch (value_format::NormalizeType(state->type))
    {
    case REG_SZ:
    case REG_LINK:
        SetDlgItemTextW(dlg, IDC_REG_SZ_EDIT, filled ? RegDataToString(data).c_str() : L"");
        break;
    case REG_EXPAND_SZ:
        SetDlgItemTextW(dlg, IDC_REG_EXPAND_EDIT, filled ? RegDataToString(data).c_str() : L"");
        break;
    case REG_MULTI_SZ:
        SetDlgItemTextW(dlg, IDC_REG_MULTI_EDIT, filled ? value_format::MultiStringText(data).c_str() : L"");
        break;
    case REG_DWORD:
        SetDlgItemTextW(
            dlg,
            IDC_REG_DWORD_EDIT,
            filled ? FormatNumberValue(value_format::ReadUnsigned(data, sizeof(DWORD)), state->dword_base).c_str() : L""
        );
        break;
    case REG_DWORD_BIG_ENDIAN:
        SetDlgItemTextW(
            dlg,
            IDC_REG_DWORD_EDIT,
            filled ? FormatNumberValue(value_format::ReadUnsigned(data, sizeof(DWORD), true), state->dword_base).c_str()
                   : L""
        );
        break;
    case REG_QWORD:
        SetDlgItemTextW(
            dlg,
            IDC_REG_QWORD_EDIT,
            filled
                ? FormatNumberValue(value_format::ReadUnsigned(data, sizeof(unsigned long long)), state->qword_base).c_str()
                : L""
        );
        break;
    case REG_NONE:
        SetDlgItemTextW(dlg, IDC_REG_NONE_EDIT, win32ui::HexByteLines(data).c_str());
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        break;
    default:
        SetDlgItemTextW(dlg, IDC_REG_BINARY_EDIT, win32ui::HexByteLines(data).c_str());
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        break;
    }
    state->shown_text = util::DialogItemText(dlg, TraceEditorId(state->type));
    state->shown_dword_base = state->dword_base;
    state->shown_qword_base = state->qword_base;
}

bool TraceEditorUntouched(HWND dlg, const TraceValueDialogState& state, DWORD type)
{
    return type == state.type && state.dword_base == state.shown_dword_base && state.qword_base == state.shown_qword_base &&
           util::DialogItemText(dlg, TraceEditorId(type)) == state.shown_text;
}

int TraceEditorId(DWORD type)
{
    switch (type)
    {
    case REG_EXPAND_SZ:
        return IDC_REG_EXPAND_EDIT;
    case REG_MULTI_SZ:
        return IDC_REG_MULTI_EDIT;
    case REG_DWORD:
    case REG_DWORD_BIG_ENDIAN:
        return IDC_REG_DWORD_EDIT;
    case REG_QWORD:
        return IDC_REG_QWORD_EDIT;
    case REG_BINARY:
    case REG_RESOURCE_LIST:
    case REG_FULL_RESOURCE_DESCRIPTOR:
    case REG_RESOURCE_REQUIREMENTS_LIST:
        return IDC_REG_BINARY_EDIT;
    case REG_NONE:
        return IDC_REG_NONE_EDIT;
    default:
        return IDC_REG_SZ_EDIT;
    }
}

bool SerializeTraceEditor(HWND dlg, TraceValueDialogState* state, DWORD type, std::vector<BYTE>* out)
{
    if (!state || !out)
    {
        return false;
    }
    std::vector<BYTE> data;
    bool ok = true;
    switch (type)
    {
    case REG_SZ:
        {
            std::wstring text = util::DialogItemText(dlg, IDC_REG_SZ_EDIT);
            data = value_format::StringData(text);
            break;
        }
    case REG_EXPAND_SZ:
        {
            std::wstring text = util::DialogItemText(dlg, IDC_REG_EXPAND_EDIT);
            data = value_format::StringData(text);
            break;
        }
    case REG_LINK:
        {
            std::wstring text = util::DialogItemText(dlg, IDC_REG_SZ_EDIT);
            data = value_format::StringData(text);
            break;
        }
    case REG_MULTI_SZ:
        {
            std::wstring text = util::DialogItemText(dlg, IDC_REG_MULTI_EDIT);
            data = value_format::MultiStringData(text);
            break;
        }
    case REG_DWORD:
        {
            const std::optional<uint64_t> value = win32ui::ParseUnsigned(util::DialogItemText(dlg, IDC_REG_DWORD_EDIT), state->dword_base, true);
            if (!value || *value > std::numeric_limits<DWORD>::max())
            {
                ok = false;
            }
            else
            {
                data.resize(sizeof(DWORD));
                DWORD v32 = static_cast<DWORD>(*value);
                memcpy(data.data(), &v32, sizeof(DWORD));
            }
            break;
        }
    case REG_DWORD_BIG_ENDIAN:
        {
            const std::optional<uint64_t> value = win32ui::ParseUnsigned(util::DialogItemText(dlg, IDC_REG_DWORD_EDIT), state->dword_base, true);
            if (!value || *value > std::numeric_limits<DWORD>::max())
            {
                ok = false;
            }
            else
            {
                data = value_format::UnsignedBytes(*value, sizeof(DWORD), true);
            }
            break;
        }
    case REG_QWORD:
        {
            const std::optional<uint64_t> value = win32ui::ParseUnsigned(util::DialogItemText(dlg, IDC_REG_QWORD_EDIT), state->qword_base, true);
            if (!value)
            {
                ok = false;
            }
            else
            {
                data.resize(sizeof(unsigned long long));
                memcpy(data.data(), &*value, sizeof(unsigned long long));
            }
            break;
        }
    case REG_BINARY:
        {
            std::wstring text = util::DialogItemText(dlg, IDC_REG_BINARY_EDIT);
            ok = win32ui::ParseHexBytes(text, &data);
            break;
        }
    case REG_RESOURCE_LIST:
    case REG_FULL_RESOURCE_DESCRIPTOR:
    case REG_RESOURCE_REQUIREMENTS_LIST:
        {
            std::wstring text = util::DialogItemText(dlg, IDC_REG_BINARY_EDIT);
            ok = win32ui::ParseHexBytes(text, &data);
            break;
        }
    case REG_NONE:
        {
            std::wstring text = util::DialogItemText(dlg, IDC_REG_NONE_EDIT);
            ok = win32ui::ParseHexBytes(text, &data);
            break;
        }
    default:
        ok = false;
        break;
    }
    if (!ok)
    {
        return false;
    }
    *out = std::move(data);
    return true;
}

INT_PTR CALLBACK CustomValueDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<TraceValueDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    switch (msg)
    {
    case WM_INITDIALOG:
        {
            state = reinterpret_cast<TraceValueDialogState*>(lparam);
            SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
            PopulateTraceTypeCombo(dlg);
            if (state)
            {
                ConfigureReadOnlyNameField(dlg, state->value_name);
                SelectTraceType(dlg, state, state->type);
                PopulateTraceValueEditors(dlg, state);
            }
            else
            {
                ConfigureReadOnlyNameField(dlg, L"");
                TraceValueDialogState temp;
                SelectTraceType(dlg, &temp, REG_SZ);
            }

            CheckDlgButton(dlg, IDC_REG_DWORD_HEX, BST_CHECKED);
            CheckDlgButton(dlg, IDC_REG_DWORD_DEC, BST_UNCHECKED);
            CheckDlgButton(dlg, IDC_REG_DWORD_BIN, BST_UNCHECKED);
            CheckDlgButton(dlg, IDC_REG_QWORD_HEX, BST_CHECKED);
            CheckDlgButton(dlg, IDC_REG_QWORD_DEC, BST_UNCHECKED);
            CheckDlgButton(dlg, IDC_REG_QWORD_BIN, BST_UNCHECKED);
            if (state)
            {
                state->dword_base = 16;
                state->qword_base = 16;
                state->binary.group_bytes = 1;
                state->binary.unicode = false;
                state->none.group_bytes = 1;
                state->none.unicode = false;
            }
            SetBinaryGroupSelection(dlg, kBinaryIds, IDC_REG_BINARY_FORMAT_BYTE);
            SetBinaryTextSelection(dlg, kBinaryIds, IDC_REG_BINARY_TEXT_ANSI);
            SetBinaryGroupSelection(dlg, kNoneIds, IDC_REG_NONE_FORMAT_BYTE);
            SetBinaryTextSelection(dlg, kNoneIds, IDC_REG_NONE_TEXT_ANSI);

            if (!state)
            {
                return FALSE;
            }
            if (state->read_only)
            {
                ConfigureReadOnlyControls(dlg, {IDC_REG_SZ_EDIT, IDC_REG_EXPAND_EDIT, IDC_REG_MULTI_EDIT, IDC_REG_DWORD_EDIT, IDC_REG_QWORD_EDIT, IDC_REG_BINARY_EDIT, IDC_REG_NONE_EDIT}, {IDC_TYPE_COMBO, IDC_REG_DWORD_HEX, IDC_REG_DWORD_DEC, IDC_REG_DWORD_BIN, IDC_REG_QWORD_HEX, IDC_REG_QWORD_DEC, IDC_REG_QWORD_BIN, IDC_REG_BINARY_FORMAT_BYTE, IDC_REG_BINARY_FORMAT_WORD, IDC_REG_BINARY_FORMAT_DWORD, IDC_REG_BINARY_FORMAT_QWORD, IDC_REG_BINARY_TEXT_ANSI, IDC_REG_BINARY_TEXT_UNICODE, IDC_REG_NONE_FORMAT_BYTE, IDC_REG_NONE_FORMAT_WORD, IDC_REG_NONE_FORMAT_DWORD, IDC_REG_NONE_FORMAT_QWORD, IDC_REG_NONE_TEXT_ANSI, IDC_REG_NONE_TEXT_UNICODE});
            }
            win32ui::AllowLineBreaks(GetDlgItem(dlg, IDC_REG_MULTI_EDIT));
            win32ui::AnchorDialogControls(dlg, {
                                                   {IDC_VALUE_NAME, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                   {IDC_GROUP_REG_SZ, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_SZ_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                   {IDC_GROUP_REG_EXPAND, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_EXPAND_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                   {IDC_GROUP_REG_MULTI, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_MULTI_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_GROUP_REG_DWORD, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_GROUP_REG_QWORD, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_GROUP_REG_BINARY, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_LABEL_HEX, win32ui::kAnchorLeft | win32ui::kAnchorTop},
                                                   {IDC_REG_BINARY_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_LABEL_PREVIEW, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_PREVIEW, win32ui::kAnchorLeft | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_FORMAT_GROUP, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_FORMAT_BYTE, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_FORMAT_WORD, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_FORMAT_DWORD, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_FORMAT_QWORD, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_TEXT_GROUP, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_TEXT_ANSI, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_BINARY_TEXT_UNICODE, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_GROUP_REG_NONE, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_LABEL_HEX, win32ui::kAnchorLeft | win32ui::kAnchorTop},
                                                   {IDC_REG_NONE_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_LABEL_PREVIEW, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_PREVIEW, win32ui::kAnchorLeft | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_FORMAT_GROUP, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_FORMAT_BYTE, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_FORMAT_WORD, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_FORMAT_DWORD, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_FORMAT_QWORD, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_TEXT_GROUP, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_TEXT_ANSI, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDC_REG_NONE_TEXT_UNICODE, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                               });
            if (state)
            {
                win32ui::SetMonoFont(dlg, {IDC_REG_BINARY_EDIT, IDC_REG_BINARY_PREVIEW, IDC_REG_NONE_EDIT, IDC_REG_NONE_PREVIEW}, nullptr, &state->mono_font);
            }
            UpdateBinaryPreviewEx(dlg, state ? &state->binary : nullptr, kBinaryIds);
            UpdateBinaryPreviewEx(dlg, state ? &state->none : nullptr, kNoneIds);
            return TRUE;
        }
    case WM_SETFONT:
        if (state)
        {
            if (state->mono_font)
            {
                win32ui::SetMonoFont(dlg, {IDC_REG_BINARY_EDIT, IDC_REG_BINARY_PREVIEW, IDC_REG_NONE_EDIT, IDC_REG_NONE_PREVIEW}, reinterpret_cast<HFONT>(wparam), &state->mono_font);
            }
        }
        return FALSE;
    case WM_COMMAND:
        {
            if (!state)
            {
                return TRUE;
            }
            int id = LOWORD(wparam);
            int code = HIWORD(wparam);
            if (code == CBN_SELCHANGE && id == IDC_TYPE_COMBO)
            {
                const DWORD previous = state->type;
                const DWORD type = ReadTraceType(dlg, state);
                if (type == previous)
                {
                    return TRUE;
                }
                const bool untouched = TraceEditorUntouched(dlg, *state, previous);
                const bool empty_editor = !untouched && util::DialogItemText(dlg, TraceEditorId(previous)).empty();
                std::vector<BYTE> current = state->data;
                if (!untouched && !empty_editor && !SerializeTraceEditor(dlg, state, previous, &current))
                {
                    ui::ShowError(dlg, util::Tr(L"The current data isn't valid, so the type can't be changed."));
                    SelectTraceType(dlg, state, previous);
                    return TRUE;
                }
                if (empty_editor || current.empty())
                {
                    state->type = type;
                    state->data.clear();
                    SelectTraceType(dlg, state, type);
                    PopulateTraceValueEditors(dlg, state);
                    return TRUE;
                }
                std::vector<BYTE> carried;
                // carry data only when the new type can represent it without loss
                if (ConvertValueData(previous, current, type, &carried))
                {
                    state->type = type;
                    state->data = std::move(carried);
                    SelectTraceType(dlg, state, type);
                    PopulateTraceValueEditors(dlg, state);
                    return TRUE;
                }
                // require confirmation before clearing data that cant be converted
                const int choice =
                    ui::PromptKeyChoice(dlg, util::Tr(L"The current data can't be represented in this type. Continue and start with an empty value?"), value_format::TypeName(type), util::Tr(L"Change Data Type"), util::Tr(L"Continue"), util::Tr(L"Cancel"), L"", {85, 70, 70});
                if (choice != IDYES)
                {
                    SelectTraceType(dlg, state, previous);
                    return TRUE;
                }
                state->type = type;
                state->data.clear();
                SelectTraceType(dlg, state, type);
                PopulateTraceValueEditors(dlg, state);
                return TRUE;
            }

            if (code == BN_CLICKED)
            {
                switch (id)
                {
                case IDC_REG_DWORD_BITS:
                    if (state)
                    {
                        RunBitfieldEditor(dlg, state->value_name, state->key_path, IDC_REG_DWORD_EDIT, state->dword_base, 32u, state->read_only);
                    }
                    return TRUE;
                case IDC_REG_QWORD_BITS:
                    if (state)
                    {
                        RunBitfieldEditor(dlg, state->value_name, state->key_path, IDC_REG_QWORD_EDIT, state->qword_base, 64u, state->read_only);
                    }
                    return TRUE;
                case IDC_REG_DWORD_HEX:
                case IDC_REG_DWORD_DEC:
                case IDC_REG_DWORD_BIN:
                    if (state)
                    {
                        unsigned long long fallback = state->type == REG_DWORD_BIG_ENDIAN
                                                          ? value_format::ReadUnsigned(state->data, sizeof(DWORD), true)
                                                          : value_format::ReadUnsigned(state->data, sizeof(DWORD));
                        unsigned long long value =
                            ReadNumberWithFallback(dlg, IDC_REG_DWORD_EDIT, state->dword_base, fallback);
                        if (id == IDC_REG_DWORD_HEX)
                        {
                            state->dword_base = 16;
                        }
                        else if (id == IDC_REG_DWORD_BIN)
                        {
                            state->dword_base = 2;
                        }
                        else
                        {
                            state->dword_base = 10;
                        }
                        CheckDlgButton(dlg, IDC_REG_DWORD_HEX, state->dword_base == 16 ? BST_CHECKED : BST_UNCHECKED);
                        CheckDlgButton(dlg, IDC_REG_DWORD_DEC, state->dword_base == 10 ? BST_CHECKED : BST_UNCHECKED);
                        CheckDlgButton(dlg, IDC_REG_DWORD_BIN, state->dword_base == 2 ? BST_CHECKED : BST_UNCHECKED);
                        std::wstring formatted = FormatNumberValue(value, state->dword_base);
                        SetDlgItemTextW(dlg, IDC_REG_DWORD_EDIT, formatted.c_str());
                        SendDlgItemMessageW(dlg, IDC_REG_DWORD_EDIT, EM_SETSEL, 0, -1);
                    }
                    return TRUE;
                case IDC_REG_QWORD_HEX:
                case IDC_REG_QWORD_DEC:
                case IDC_REG_QWORD_BIN:
                    if (state)
                    {
                        int old_base = state->qword_base;
                        unsigned long long fallback = value_format::ReadUnsigned(state->data, sizeof(unsigned long long));
                        unsigned long long value = ReadNumberWithFallback(dlg, IDC_REG_QWORD_EDIT, old_base, fallback);
                        if (id == IDC_REG_QWORD_HEX)
                        {
                            state->qword_base = 16;
                        }
                        else if (id == IDC_REG_QWORD_BIN)
                        {
                            state->qword_base = 2;
                        }
                        else
                        {
                            state->qword_base = 10;
                        }
                        CheckDlgButton(dlg, IDC_REG_QWORD_HEX, state->qword_base == 16 ? BST_CHECKED : BST_UNCHECKED);
                        CheckDlgButton(dlg, IDC_REG_QWORD_DEC, state->qword_base == 10 ? BST_CHECKED : BST_UNCHECKED);
                        CheckDlgButton(dlg, IDC_REG_QWORD_BIN, state->qword_base == 2 ? BST_CHECKED : BST_UNCHECKED);
                        std::wstring formatted = FormatNumberValue(value, state->qword_base);
                        SetDlgItemTextW(dlg, IDC_REG_QWORD_EDIT, formatted.c_str());
                        SendDlgItemMessageW(dlg, IDC_REG_QWORD_EDIT, EM_SETSEL, 0, -1);
                    }
                    return TRUE;
                case IDC_REG_BINARY_FORMAT_BYTE:
                case IDC_REG_BINARY_FORMAT_WORD:
                case IDC_REG_BINARY_FORMAT_DWORD:
                case IDC_REG_BINARY_FORMAT_QWORD:
                case IDC_REG_BINARY_TEXT_ANSI:
                case IDC_REG_BINARY_TEXT_UNICODE:
                    SelectBinaryMode(dlg, &state->binary, kBinaryIds, id);
                    return TRUE;
                case IDC_REG_NONE_FORMAT_BYTE:
                case IDC_REG_NONE_FORMAT_WORD:
                case IDC_REG_NONE_FORMAT_DWORD:
                case IDC_REG_NONE_FORMAT_QWORD:
                case IDC_REG_NONE_TEXT_ANSI:
                case IDC_REG_NONE_TEXT_UNICODE:
                    SelectBinaryMode(dlg, &state->none, kNoneIds, id);
                    return TRUE;
                default:
                    break;
                }
            }

            if (code == EN_CHANGE)
            {
                if (id == IDC_REG_BINARY_EDIT)
                {
                    UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
                    return TRUE;
                }
                if (id == IDC_REG_NONE_EDIT)
                {
                    UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
                    return TRUE;
                }
            }

            if (id == IDOK)
            {
                DWORD type = ReadTraceType(dlg, state);
                std::vector<BYTE> data = state->data;
                if (!TraceEditorUntouched(dlg, *state, type) && !SerializeTraceEditor(dlg, state, type, &data))
                {
                    ui::ShowError(dlg, util::Tr(L"Invalid value data."));
                    return TRUE;
                }
                state->type = type;
                state->data = std::move(data);
                state->accepted = true;
                win32ui::EndDialogWindow(dlg, IDOK);
                return TRUE;
            }
            if (id == IDCANCEL)
            {
                state->accepted = false;
                win32ui::EndDialogWindow(dlg, IDCANCEL);
                return TRUE;
            }
            break;
        }
    default:
        break;
    }
    return FALSE;
}

INT_PTR CALLBACK ExtendedValueDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<ExtendedValueDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    switch (msg)
    {
    case WM_INITDIALOG:
        {
            state = reinterpret_cast<ExtendedValueDialogState*>(lparam);
            SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
            if (state)
            {
                SetDlgItemTextW(dlg, IDC_EDIT, state->initial_text.c_str());
                state->initial_text = util::DialogItemText(dlg, IDC_EDIT);
                ConfigureReadOnlyNameField(dlg, state->value_name);
            }

            if (state && (state->base_type == REG_DWORD || state->base_type == REG_DWORD_BIG_ENDIAN ||
                          state->base_type == REG_QWORD))
            {
                CheckDlgButton(dlg, IDC_HEX, state->number_base == 16 ? BST_CHECKED : BST_UNCHECKED);
                CheckDlgButton(dlg, IDC_DEC, state->number_base == 10 ? BST_CHECKED : BST_UNCHECKED);
                CheckDlgButton(dlg, IDC_BIN, state->number_base == 2 ? BST_CHECKED : BST_UNCHECKED);
            }
            if (!state)
            {
                return FALSE;
            }
            if (state->read_only)
            {
                ConfigureReadOnlyControls(dlg, {IDC_EDIT}, {IDC_HEX, IDC_DEC, IDC_BIN});
            }
            if (state->base_type == REG_MULTI_SZ)
            {
                win32ui::AllowLineBreaks(GetDlgItem(dlg, IDC_EDIT));
            }
            if (IsMultilineEdit(dlg, IDC_EDIT))
            {
                win32ui::AnchorDialogControls(dlg, {
                                                       {IDC_VALUE_NAME, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                       {IDC_LABEL, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                       {IDC_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                       {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                       {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                   });
            }
            return TRUE;
        }
    case WM_COMMAND:
        {
            if (!state)
            {
                return TRUE;
            }
            int id = LOWORD(wparam);
            int code = HIWORD(wparam);

            if (code == BN_CLICKED)
            {
                switch (id)
                {
                case IDC_BITS:
                    if (state)
                    {
                        RunBitfieldEditor(dlg, state->value_name, state->key_path, IDC_EDIT, state->number_base, state->base_type == REG_QWORD ? 64u : 32u, state->read_only);
                    }
                    return TRUE;
                case IDC_HEX:
                case IDC_DEC:
                case IDC_BIN:
                    if (state)
                    {
                        unsigned long long fallback = 0;
                        if (state->base_type == REG_DWORD)
                        {
                            fallback = value_format::ReadUnsigned(state->initial_data, sizeof(DWORD));
                        }
                        else if (state->base_type == REG_DWORD_BIG_ENDIAN)
                        {
                            fallback = value_format::ReadUnsigned(state->initial_data, sizeof(DWORD), true);
                        }
                        else if (state->base_type == REG_QWORD)
                        {
                            fallback = value_format::ReadUnsigned(state->initial_data, sizeof(unsigned long long));
                        }
                        unsigned long long value = ReadNumberWithFallback(dlg, IDC_EDIT, state->number_base, fallback);
                        if (id == IDC_HEX)
                        {
                            state->number_base = 16;
                        }
                        else if (id == IDC_BIN)
                        {
                            state->number_base = 2;
                        }
                        else
                        {
                            state->number_base = 10;
                        }
                        CheckDlgButton(dlg, IDC_HEX, state->number_base == 16 ? BST_CHECKED : BST_UNCHECKED);
                        CheckDlgButton(dlg, IDC_DEC, state->number_base == 10 ? BST_CHECKED : BST_UNCHECKED);
                        CheckDlgButton(dlg, IDC_BIN, state->number_base == 2 ? BST_CHECKED : BST_UNCHECKED);
                        std::wstring formatted = FormatNumberValue(value, state->number_base);
                        SetDlgItemTextW(dlg, IDC_EDIT, formatted.c_str());
                        SendDlgItemMessageW(dlg, IDC_EDIT, EM_SETSEL, 0, -1);
                    }
                    return TRUE;
                default:
                    break;
                }
            }

            if (id == IDOK)
            {
                std::wstring base_text = util::DialogItemText(dlg, IDC_EDIT);
                const bool is_number = state->base_type == REG_DWORD || state->base_type == REG_DWORD_BIG_ENDIAN ||
                                       state->base_type == REG_QWORD;
                const bool unchanged =
                    base_text == state->initial_text && (!is_number || state->number_base == state->initial_number_base);
                if (unchanged)
                {
                    state->data = state->initial_data;
                }
                else
                {
                    std::vector<BYTE> base_data;
                    switch (state->base_type)
                    {
                    case REG_SZ:
                    case REG_EXPAND_SZ:
                    case REG_LINK:
                        base_data = value_format::StringData(base_text);
                        break;
                    case REG_MULTI_SZ:
                        base_data = value_format::MultiStringData(base_text);
                        break;
                    case REG_DWORD:
                    case REG_DWORD_BIG_ENDIAN:
                    case REG_QWORD:
                        {
                            const std::optional<uint64_t> parsed = win32ui::ParseUnsigned(base_text, state->number_base, true);
                            if (!parsed)
                            {
                                ui::ShowError(dlg, util::Tr(L"Invalid number."));
                                return TRUE;
                            }
                            const unsigned long long value = *parsed;
                            if ((state->base_type == REG_DWORD || state->base_type == REG_DWORD_BIG_ENDIAN) &&
                                value > std::numeric_limits<DWORD>::max())
                            {
                                ui::ShowError(dlg, util::Tr(L"Number is out of range."));
                                return TRUE;
                            }
                            if (state->base_type == REG_DWORD)
                            {
                                DWORD v32 = static_cast<DWORD>(value);
                                base_data.resize(sizeof(DWORD));
                                memcpy(base_data.data(), &v32, sizeof(DWORD));
                            }
                            else if (state->base_type == REG_DWORD_BIG_ENDIAN)
                            {
                                base_data = value_format::UnsignedBytes(value, sizeof(DWORD), true);
                            }
                            else
                            {
                                base_data.resize(sizeof(unsigned long long));
                                memcpy(base_data.data(), &value, sizeof(unsigned long long));
                            }
                            break;
                        }
                    default:
                        ui::ShowError(dlg, util::Tr(L"Invalid value data."));
                        return TRUE;
                    }
                    state->data = std::move(base_data);
                }
                state->accepted = true;
                win32ui::EndDialogWindow(dlg, IDOK);
                return TRUE;
            }
            if (id == IDCANCEL)
            {
                state->accepted = false;
                win32ui::EndDialogWindow(dlg, IDCANCEL);
                return TRUE;
            }
            break;
        }
    default:
        break;
    }
    return FALSE;
}

std::wstring RegDataToString(const std::vector<BYTE>& data)
{
    if (data.empty())
    {
        return L"";
    }
    size_t wchar_count = data.size() / sizeof(wchar_t);
    std::wstring text(reinterpret_cast<const wchar_t*>(data.data()), wchar_count);
    // hide registry string terminators from text control
    while (!text.empty() && text.back() == L'\0')
    {
        text.pop_back();
    }
    return text;
}

}

bool EditValueBits(HWND owner, const BitsRequest& request, std::vector<BYTE>* data)
{
    if (!data)
    {
        return false;
    }
    BitfieldRequest bits;
    bits.value_name = request.value_name;
    bits.key_path = request.key_path;
    bits.read_only = request.read_only;
    std::vector<BYTE> source(request.data.begin(), request.data.end());

    if (request.base_type == REG_DWORD && source.size() == sizeof(DWORD))
    {
        bits.bit_count = 32;
        bits.value = value_format::ReadUnsigned(source, sizeof(DWORD));
    }
    else if (request.base_type == REG_QWORD && source.size() == sizeof(unsigned long long))
    {
        bits.bit_count = 64;
        bits.value = value_format::ReadUnsigned(source, sizeof(unsigned long long));
    }
    else if (request.base_type == REG_DWORD_BIG_ENDIAN && source.size() == sizeof(DWORD))
    {
        bits.bit_count = 32;
        bits.value = value_format::ReadUnsigned(source, sizeof(DWORD), true);
    }
    else
    {
        if (source.empty())
        {
            ui::ShowError(owner, util::Tr(L"This value has no data to edit as bits."));
            return false;
        }
        bits.data = source;
    }

    BitfieldResult result;
    if (!EditBitfield(owner, bits, &result))
    {
        return false;
    }
    if (!bits.data.empty())
    {
        *data = std::move(result.data);
        return true;
    }
    if (request.base_type == REG_DWORD_BIG_ENDIAN)
    {
        // restore registry type big endian byte order after editing
        *data = value_format::UnsignedBytes(result.value, sizeof(DWORD), true);
        return true;
    }
    const size_t width = bits.bit_count / 8;
    data->assign(width, 0);
    memcpy(data->data(), &result.value, width);
    return true;
}

bool EditCustomValue(HWND owner, const CustomValueRequest& request, CustomValueResult* result)
{
    if (!result)
    {
        return false;
    }
    TraceValueDialogState state;
    state.value_name = request.value_name;
    state.key_path = request.key_path;
    state.type = request.type;
    state.data = request.data;
    state.read_only = request.read_only;
    const INT_PTR dialog_result = win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_CUSTOM_VALUE), CustomValueDialogProc, reinterpret_cast<LPARAM>(&state));
    if (dialog_result != IDOK || !state.accepted)
    {
        return false;
    }
    result->type = state.type;
    result->data = std::move(state.data);
    return true;
}

bool EditFlaggedValue(HWND owner, const FlaggedValueRequest& request, FlaggedValueResult* result)
{
    if (!result)
    {
        return false;
    }
    ExtendedValueDialogState state;
    state.base_type = request.base_type;
    state.value_name = request.value_name;
    state.key_path = request.key_path;
    state.initial_data.assign(request.data.begin(), request.data.end());
    state.read_only = request.read_only;
    state.number_base = 16;
    state.initial_number_base = state.number_base;

    switch (request.base_type)
    {
    case REG_MULTI_SZ:
        state.initial_text = value_format::MultiStringText(state.initial_data);
        break;
    case REG_DWORD:
        {
            DWORD value = 0;
            if (request.data.size() >= sizeof(DWORD))
            {
                memcpy(&value, request.data.data(), sizeof(value));
            }
            state.initial_text = FormatNumberValue(value, state.number_base);
            break;
        }
    case REG_DWORD_BIG_ENDIAN:
        {
            DWORD value = 0;
            if (request.data.size() >= sizeof(DWORD))
            {
                value = static_cast<DWORD>(value_format::ReadUnsigned(state.initial_data, sizeof(DWORD), true));
            }
            state.initial_text = FormatNumberValue(value, state.number_base);
            break;
        }
    case REG_QWORD:
        {
            unsigned long long value = 0;
            if (request.data.size() >= sizeof(unsigned long long))
            {
                memcpy(&value, request.data.data(), sizeof(value));
            }
            state.initial_text = FormatNumberValue(value, state.number_base);
            break;
        }
    default:
        state.initial_text = RegDataToString(state.initial_data);
        break;
    }

    int dialog_id = IDD_INPUT_BINARY;
    if (request.base_type == REG_MULTI_SZ)
    {
        dialog_id = IDD_MULTI_BINARY;
    }
    else if (request.base_type == REG_DWORD || request.base_type == REG_DWORD_BIG_ENDIAN ||
             request.base_type == REG_QWORD)
    {
        dialog_id = IDD_NUMBER_BINARY;
    }

    const INT_PTR dialog_result = win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(dialog_id), ExtendedValueDialogProc, reinterpret_cast<LPARAM>(&state));
    if (dialog_result != IDOK || !state.accepted)
    {
        return false;
    }
    result->data = std::move(state.data);
    return true;
}

}
