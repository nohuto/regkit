// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/binary_editor.h"

#include "appearance/feedback.h"
#include "editors/binary_text.h"
#include "editors/dialog_support.h"
#include "registry/value_format.h"

#include "resource.h"

#include <utility>

namespace regkit::editors {

namespace {

struct State {
  const BinaryRequest* request = nullptr;
  BinaryResult value;
  std::wstring text;
  int group_bytes = 1;
  bool unicode = false;
  bool accepted = false;
  HFONT ui_font = nullptr;
  HFONT mono_font = nullptr;
};

void SelectGroup(HWND dialog, int selected) {
  constexpr int controls[] = {IDC_FORMAT_BYTE, IDC_FORMAT_WORD,
                              IDC_FORMAT_DWORD, IDC_FORMAT_QWORD};
  for (const int id : controls) {
    CheckDlgButton(dialog, id, id == selected ? BST_CHECKED : BST_UNCHECKED);
  }
}

void SelectTextMode(HWND dialog, int selected) {
  CheckDlgButton(dialog, IDC_TEXT_ANSI,
                 selected == IDC_TEXT_ANSI ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dialog, IDC_TEXT_UNICODE,
                 selected == IDC_TEXT_UNICODE ? BST_CHECKED : BST_UNCHECKED);
}

void UpdatePreview(HWND dialog, State* state) {
  if (!state) {
    return;
  }
  state->text = dialog_support::ReadText(dialog, IDC_EDIT);
  std::vector<BYTE> bytes;
  if (!value_format::ParseHex(state->text, &bytes)) {
    SetDlgItemTextW(dialog, IDC_BINARY_PREVIEW, L"Invalid hex input.");
    SetDlgItemTextW(dialog, IDC_VALUE_BYTES, L"Invalid");
    return;
  }
  const std::wstring preview =
      binary_text::Preview(bytes, state->group_bytes, state->unicode);
  SetDlgItemTextW(dialog, IDC_BINARY_PREVIEW, preview.c_str());
  wchar_t count[64] = {};
  swprintf_s(count, L"%llu byte%s",
             static_cast<unsigned long long>(bytes.size()),
             bytes.size() == 1 ? L"" : L"s");
  SetDlgItemTextW(dialog, IDC_VALUE_BYTES, count);
}

void ConfigureIdentity(HWND dialog, const BinaryRequest& request) {
  const std::wstring name =
      request.value_name.empty() ? L"(Default)" : request.value_name;
  SetDlgItemTextW(dialog, IDC_VALUE_NAME, name.c_str());
  SendDlgItemMessageW(dialog, IDC_VALUE_NAME, EM_SETREADONLY, TRUE, 0);
  const HWND name_control = GetDlgItem(dialog, IDC_VALUE_NAME);
  SetWindowLongPtrW(name_control, GWL_STYLE,
                    GetWindowLongPtrW(name_control, GWL_STYLE) & ~WS_TABSTOP);
  ShowWindow(GetDlgItem(dialog, IDC_VALUE_TYPE_LABEL), SW_HIDE);
  ShowWindow(GetDlgItem(dialog, IDC_VALUE_TYPE), SW_HIDE);
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam,
                            LPARAM lparam) {
  auto* state = reinterpret_cast<State*>(
      GetWindowLongPtrW(dialog, DWLP_USER));
  if (message == WM_INITDIALOG) {
    state = reinterpret_cast<State*>(lparam);
    state->text = binary_text::Hex(state->request->data);
    SetWindowLongPtrW(dialog, DWLP_USER,
                      reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dialog, L"Edit Value");
    SetDlgItemTextW(dialog, IDC_LABEL, L"Hex bytes:");
    SetDlgItemTextW(dialog, IDC_NOTE, L"Preview:");
    SetDlgItemTextW(dialog, IDC_EDIT, state->text.c_str());
    ConfigureIdentity(dialog, *state->request);
    SelectGroup(dialog, IDC_FORMAT_BYTE);
    SelectTextMode(dialog, IDC_TEXT_ANSI);
    dialog_support::Initialize(
        dialog, &state->ui_font,
        {IDC_VALUE_NAME, IDC_EDIT, IDC_BINARY_PREVIEW});
    state->mono_font = CreateFontW(
        -12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_MODERN,
        L"Consolas");
    if (state->mono_font) {
      SendDlgItemMessageW(dialog, IDC_EDIT, WM_SETFONT,
                          reinterpret_cast<WPARAM>(state->mono_font), TRUE);
      SendDlgItemMessageW(dialog, IDC_BINARY_PREVIEW, WM_SETFONT,
                          reinterpret_cast<WPARAM>(state->mono_font), TRUE);
    }
    UpdatePreview(dialog, state);
    return TRUE;
  }
  if (message == WM_DESTROY) {
    if (state) {
      dialog_support::ReleaseFont(&state->mono_font);
      dialog_support::ReleaseFont(&state->ui_font);
    }
    return TRUE;
  }
  INT_PTR themed = 0;
  if (dialog_support::HandleThemeMessage(
          dialog, message, wparam, lparam, &themed)) {
    return themed;
  }
  if (message != WM_COMMAND || !state) {
    return FALSE;
  }
  const int id = LOWORD(wparam);
  if (id == IDC_EDIT && HIWORD(wparam) == EN_CHANGE) {
    UpdatePreview(dialog, state);
    return TRUE;
  }
  switch (id) {
  case IDC_FORMAT_BYTE:
  case IDC_FORMAT_WORD:
  case IDC_FORMAT_DWORD:
  case IDC_FORMAT_QWORD:
    state->group_bytes = id == IDC_FORMAT_BYTE ? 1 : id == IDC_FORMAT_WORD ? 2
                                                 : id == IDC_FORMAT_DWORD  ? 4
                                                                           : 8;
    SelectGroup(dialog, id);
    UpdatePreview(dialog, state);
    return TRUE;
  case IDC_TEXT_ANSI:
  case IDC_TEXT_UNICODE:
    state->unicode = id == IDC_TEXT_UNICODE;
    SelectTextMode(dialog, id);
    UpdatePreview(dialog, state);
    return TRUE;
  case IDOK:
    state->text = dialog_support::ReadText(dialog, IDC_EDIT);
    state->accepted = true;
    EndDialog(dialog, IDOK);
    return TRUE;
  case IDCANCEL:
    EndDialog(dialog, IDCANCEL);
    return TRUE;
  default:
    return FALSE;
  }
}

} // namespace

bool EditBinary(HWND owner, const BinaryRequest& request,
                BinaryResult* result) {
  if (!result) {
    return false;
  }
  State state;
  state.request = &request;
  const INT_PTR dialog_result = DialogBoxParamW(
      GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_BINARY), owner,
      DialogProc, reinterpret_cast<LPARAM>(&state));
  if (dialog_result != IDOK || !state.accepted) {
    return false;
  }
  if (!value_format::ParseHex(state.text, &state.value.data)) {
    ui::ShowError(owner, L"Invalid hex input.");
    return false;
  }
  *result = std::move(state.value);
  return true;
}

} // namespace regkit::editors
