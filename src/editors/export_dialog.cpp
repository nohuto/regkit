// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/export_dialog.h"

#include "appearance/feedback.h"
#include "editors/dialog_support.h"
#include "win32/file_dialog.h"

#include "resource.h"

#include <utility>

namespace regkit::editors {

namespace {

struct State {
  ExportResult value;
  HFONT font = nullptr;
  bool accepted = false;
};

bool ChoosePath(HWND owner, std::wstring* path) {
  const HRESULT hr = win32::ChooseFileToSave(
      owner, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0", L"reg",
      path && !path->empty() ? path->c_str() : nullptr, path);
  if (FAILED(hr) && !win32::DialogCancelled(hr)) {
    ui::ShowError(owner, win32::FormatDialogError(hr));
  }
  return SUCCEEDED(hr);
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam,
                            LPARAM lparam) {
  auto* state = reinterpret_cast<State*>(
      GetWindowLongPtrW(dialog, DWLP_USER));
  if (message == WM_INITDIALOG) {
    state = reinterpret_cast<State*>(lparam);
    SetWindowLongPtrW(dialog, DWLP_USER,
                      reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dialog, L"RegKit");
    SetDlgItemTextW(dialog, IDC_EXPORT_PATH, state->value.path.c_str());
    CheckDlgButton(dialog, IDC_EXPORT_RANGE_BRANCH,
                   state->value.include_subkeys ? BST_CHECKED
                                                : BST_UNCHECKED);
    CheckDlgButton(dialog, IDC_EXPORT_RANGE_KEY,
                   state->value.include_subkeys ? BST_UNCHECKED
                                                : BST_CHECKED);
    CheckDlgButton(dialog, IDC_EXPORT_OPEN_AFTER,
                   state->value.open_after ? BST_CHECKED : BST_UNCHECKED);
    dialog_support::Initialize(dialog, &state->font,
                               {IDC_EXPORT_PATH});
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
  const int id = LOWORD(wparam);
  if (id == IDC_EXPORT_BROWSE && HIWORD(wparam) == BN_CLICKED) {
    std::wstring path =
        dialog_support::ReadText(dialog, IDC_EXPORT_PATH);
    if (ChoosePath(dialog, &path)) {
      SetDlgItemTextW(dialog, IDC_EXPORT_PATH, path.c_str());
    }
    return TRUE;
  }
  if (id == IDOK) {
    state->value.path =
        dialog_support::ReadText(dialog, IDC_EXPORT_PATH);
    if (state->value.path.empty()) {
      ui::ShowError(dialog, L"Select a destination file.");
      return TRUE;
    }
    state->value.include_subkeys =
        IsDlgButtonChecked(dialog, IDC_EXPORT_RANGE_BRANCH) == BST_CHECKED;
    state->value.open_after =
        IsDlgButtonChecked(dialog, IDC_EXPORT_OPEN_AFTER) == BST_CHECKED;
    state->accepted = true;
    EndDialog(dialog, IDOK);
    return TRUE;
  }
  if (id == IDCANCEL) {
    EndDialog(dialog, IDCANCEL);
    return TRUE;
  }
  return FALSE;
}

} // namespace

bool ChooseExport(HWND owner, const ExportRequest& request,
                  ExportResult* result) {
  if (!result) {
    return false;
  }
  State state;
  state.value = request;
  const INT_PTR dialog_result = DialogBoxParamW(
      GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_EXPORT_OPTIONS), owner,
      DialogProc, reinterpret_cast<LPARAM>(&state));
  if (dialog_result != IDOK || !state.accepted) {
    return false;
  }
  *result = std::move(state.value);
  return true;
}

} // namespace regkit::editors
