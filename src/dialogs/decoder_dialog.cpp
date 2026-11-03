// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/decoder_dialog.h"

#include "registry/registry_path.h"
#include "registry/value_decoder.h"
#include "registry/value_format.h"
#include "ui/dialog_layout.h"
#include "ui/feedback.h"

#include "resource.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/dialog/combo_box.h>
#include <win32ui/dialog/dialog_layout.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/gdi/fonts.h>
#include <win32ui/list_view/list_view.h>
#include <win32ui/theme/theme.h>
#include <win32ui/window.h>

#include <algorithm>
#include <span>
#include <utility>

namespace regkit::editors
{

namespace
{

using value_decoder::DecoderEntry;
using value_decoder::DecoderId;
using value_decoder::TransformEntry;
using value_decoder::TransformId;

constexpr size_t kRawBytePreviewLimit = 64 * 1024;

struct State
{
    const DecodeRequest* request = nullptr;
    std::vector<BYTE> transformed;
    std::vector<TransformEntry> transforms;
    std::vector<DecoderEntry> decoders;
    TransformId transform = TransformId::kNone;
    std::wstring output;
    win32ui::UniqueFont mono_font;
};

const std::vector<BYTE>& Bytes(const State& state)
{
    return state.transform == TransformId::kNone ? state.request->data : state.transformed;
}

template <typename Entry>
void FillCombo(HWND dialog, int control, const std::vector<Entry>& entries, int selected)
{
    std::vector<std::wstring> names;
    names.reserve(entries.size());
    for (const Entry& entry : entries)
    {
        names.emplace_back(entry.name);
    }
    const HWND combo = GetDlgItem(dialog, control);
    win32ui::SetComboItems(combo, names);
    win32ui::SetComboIndex(combo, selected);
}

int IndexOfDecoder(const std::vector<DecoderEntry>& entries, DecoderId id)
{
    for (size_t i = 0; i < entries.size(); ++i)
    {
        if (entries[i].id == id)
        {
            return static_cast<int>(i);
        }
    }
    return 0;
}

void ShowOutput(HWND dialog, State* state, std::wstring text, const std::vector<value_decoder::Field>* fields = nullptr)
{
    state->output = std::move(text);
    SetDlgItemTextW(dialog, IDC_EDIT, fields ? L"" : state->output.c_str());
    const HWND list = GetDlgItem(dialog, IDC_DECODE_FIELDS);
    {
        const win32ui::RedrawScope redraw(list, 0);
        win32ui::TextRows rows;
        for (size_t i = 0; fields && i < fields->size(); ++i)
        {
            rows.push_back({(*fields)[i].name, (*fields)[i].value});
        }
        win32ui::SetListRows(list, rows);
        // keep user hidden columns hidden
        if (ListView_GetColumnWidth(list, 0) > 0)
        {
            win32ui::FitListColumn(list, 0);
        }
        if (ListView_GetColumnWidth(list, 1) > 0)
        {
            ListView_SetColumnWidth(list, 1, LVSCW_AUTOSIZE_USEHEADER);
        }
    }
    ShowWindow(list, fields ? SW_SHOW : SW_HIDE);
    ShowWindow(GetDlgItem(dialog, IDC_EDIT), fields ? SW_HIDE : SW_SHOW);
}

void RunDecoder(HWND dialog, State* state)
{
    const int index = static_cast<int>(SendDlgItemMessageW(dialog, IDC_DECODE_FORMAT, CB_GETCURSEL, 0, 0));
    if (index < 0 || static_cast<size_t>(index) >= state->decoders.size())
    {
        return;
    }
    const DecoderId id = state->decoders[static_cast<size_t>(index)].id;
    const std::vector<BYTE>& bytes = Bytes(*state);
    const value_decoder::Decoded decoded = value_decoder::Decode(id, bytes.data(), bytes.size());
    if (!decoded.ok)
    {
        ShowOutput(dialog, state, decoded.error);
        return;
    }
    std::wstring text;
    for (const value_decoder::Field& field : decoded.fields)
    {
        text.append(field.name).append(L": ").append(field.value).append(L"\r\n");
    }
    if (id != DecoderId::kRawBytes)
    {
        ShowOutput(dialog, state, std::move(text), &decoded.fields);
        return;
    }
    // cap formatted preview so large registry values stay responsive
    const size_t shown = std::min(bytes.size(), kRawBytePreviewLimit);
    const std::span<const BYTE> span(bytes.data(), shown);
    text.append(L"\r\n").append(win32ui::HexBytes(span, L' ', true));
    text.append(L"\r\n\r\n").append(win32ui::HexDump(span, 1, false));
    if (shown != bytes.size())
    {
        text.append(L"\r\n\r\n").append(util::TrLabel(L"Bytes shown", std::to_wstring(shown)));
    }
    ShowOutput(dialog, state, std::move(text));
}

void RebuildDecoders(HWND dialog, State* state)
{
    const std::vector<BYTE>& bytes = Bytes(*state);
    state->decoders = value_decoder::AvailableDecoders(bytes.data(), bytes.size());
    DecoderId wanted = DecoderId::kRawBytes;
    if (state->transform == TransformId::kNone)
    {
        wanted = value_decoder::Suggest(state->request->key_path, state->request->value_name, bytes.size());
    }
    FillCombo(dialog, IDC_DECODE_FORMAT, state->decoders, IndexOfDecoder(state->decoders, wanted));
    RunDecoder(dialog, state);
}

void ApplyTransform(HWND dialog, State* state)
{
    const int index = static_cast<int>(SendDlgItemMessageW(dialog, IDC_DECODE_ENCODING, CB_GETCURSEL, 0, 0));
    if (index < 0 || static_cast<size_t>(index) >= state->transforms.size())
    {
        return;
    }
    const TransformId id = state->transforms[static_cast<size_t>(index)].id;
    std::vector<BYTE> bytes;
    std::wstring error;
    if (!value_decoder::Transform(id, state->request->type, state->request->data, &bytes, &error))
    {
        state->transformed.clear();
        state->decoders.clear();
        SendDlgItemMessageW(dialog, IDC_DECODE_FORMAT, CB_RESETCONTENT, 0, 0);
        ShowOutput(dialog, state, error);
        return;
    }
    state->transform = id;
    state->transformed = std::move(bytes);
    RebuildDecoders(dialog, state);
}

void ConfigureIdentity(HWND dialog, const DecodeRequest& request)
{
    const std::wstring name = request.value_name.empty() ? util::Tr(L"(Default)") : registry_path::DisplayName(request.value_name);
    SetDlgItemTextW(dialog, IDC_VALUE_NAME, name.c_str());
    SendDlgItemMessageW(dialog, IDC_VALUE_NAME, EM_SETREADONLY, TRUE, 0);
    const HWND name_control = GetDlgItem(dialog, IDC_VALUE_NAME);
    SetWindowLongPtrW(name_control, GWL_STYLE, GetWindowLongPtrW(name_control, GWL_STYLE) & ~WS_TABSTOP);
    std::wstring summary = value_format::TypeName(request.type);
    summary.append(L", ").append(value_format::ByteCount(request.data.size()));
    SetDlgItemTextW(dialog, IDC_VALUE_BYTES, summary.c_str());
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_SETFONT && state)
    {
        if (state->mono_font)
        {
            win32ui::SetMonoFont(dialog, {IDC_EDIT}, reinterpret_cast<HFONT>(wparam), &state->mono_font);
        }
        return FALSE;
    }
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<State*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        ConfigureIdentity(dialog, *state->request);
        SendDlgItemMessageW(dialog, IDC_EDIT, EM_SETREADONLY, TRUE, 0);
        state->transforms = value_decoder::AvailableTransforms(state->request->type, state->request->data);
        FillCombo(dialog, IDC_DECODE_ENCODING, state->transforms, 0);
        // type/size summary was set after fitting template text
        win32ui::FitDialogText(dialog);
        win32ui::SetMonoFont(dialog, {IDC_EDIT}, nullptr, &state->mono_font);
        win32ui::SetupDialogList(GetDlgItem(dialog, IDC_DECODE_FIELDS), nullptr, 0, {{util::Tr(L"Property"), 140}, {util::Tr(L"Value"), 200}});
        win32ui::AnchorDialogControls(dialog, {
                                                  {IDC_VALUE_NAME, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_VALUE_BYTES, win32ui::kAnchorTop | win32ui::kAnchorRight},
                                                  {IDC_DECODE_ENCODING, win32ui::kAnchorLeft | win32ui::kAnchorTop},
                                                  {IDC_DECODE_FORMAT, win32ui::kAnchorLeft | win32ui::kAnchorTop},
                                                  {IDC_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDC_DECODE_FIELDS, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                                  {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                                              });
        RebuildDecoders(dialog, state);
        return TRUE;
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    const int id = LOWORD(wparam);
    if (HIWORD(wparam) == CBN_SELCHANGE)
    {
        if (id == IDC_DECODE_ENCODING)
        {
            ApplyTransform(dialog, state);
            return TRUE;
        }
        if (id == IDC_DECODE_FORMAT)
        {
            RunDecoder(dialog, state);
            return TRUE;
        }
    }
    switch (id)
    {
    case IDOK:
        // keep decoder open after copying so other interpretations can be tried
        ui::CopyText(dialog, state->output);
        return TRUE;
    case IDCANCEL:
        win32ui::EndDialogWindow(dialog, IDCANCEL);
        return TRUE;
    default:
        return FALSE;
    }
}

}

void ShowValueDecoder(HWND owner, const DecodeRequest& request)
{
    State state;
    state.request = &request;
    win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_DECODE_VALUE), DialogProc, reinterpret_cast<LPARAM>(&state));
}

}
