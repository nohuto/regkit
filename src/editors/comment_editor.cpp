// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/comment_editor.h"

#include "editors/dialog_support.h"

#include "resource.h"

#include <utility>

namespace regkit::editors {

namespace {

struct State {
  CommentResult value;
  HFONT font = nullptr;
  bool accepted = false;
};

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam,
                            LPARAM lparam) {
  auto* state = reinterpret_cast<State*>(
      GetWindowLongPtrW(dialog, DWLP_USER));
  if (message == WM_INITDIALOG) {
    state = reinterpret_cast<State*>(lparam);
    SetWindowLongPtrW(dialog, DWLP_USER,
                      reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dialog, L"Edit Comment");
    SetDlgItemTextW(dialog, IDC_EDIT, state->value.text.c_str());
    CheckDlgButton(dialog, IDC_COMMENT_ALL,
                   state->value.apply_to_same_name ? BST_CHECKED
                                                   : BST_UNCHECKED);
    dialog_support::Initialize(dialog, &state->font, {IDC_EDIT});
    return TRUE;
  }
  if (message == WM_DESTROY) {
    if (state) {
      dialog_support::ReleaseFont(&state->font);
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
  if (LOWORD(wparam) == IDOK) {
    state->value.text = dialog_support::ReadText(dialog, IDC_EDIT);
    state->value.apply_to_same_name =
        IsDlgButtonChecked(dialog, IDC_COMMENT_ALL) == BST_CHECKED;
    state->accepted = true;
    EndDialog(dialog, IDOK);
    return TRUE;
  }
  if (LOWORD(wparam) == IDCANCEL) {
    EndDialog(dialog, IDCANCEL);
    return TRUE;
  }
  return FALSE;
}

} // namespace

bool EditComment(HWND owner, const CommentRequest& request,
                 CommentResult* result) {
  if (!result) {
    return false;
  }
  State state;
  state.value.text = request.text;
  state.value.apply_to_same_name = request.apply_to_same_name;
  const INT_PTR dialog_result = DialogBoxParamW(
      GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_COMMENT), owner,
      DialogProc, reinterpret_cast<LPARAM>(&state));
  if (dialog_result != IDOK || !state.accepted) {
    return false;
  }
  *result = std::move(state.value);
  return true;
}

} // namespace regkit::editors
