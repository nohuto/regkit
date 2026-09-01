// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/hive_dialog.h"

#include "appearance/feedback.h"
#include "editors/dialog_support.h"
#include "win32/file_dialog.h"

#include "resource.h"

#include <utility>

namespace regkit::editors {

namespace {

struct State {
  LoadHiveResult value;
  HFONT font = nullptr;
  bool accepted = false;
};

bool ChooseFile(HWND owner, std::wstring* path) {
  const HRESULT hr = win32::ChooseFileToOpen(owner, L"Hive Files (*.*)\0*.*\0", path);
  if (FAILED(hr) && !win32::DialogCancelled(hr)) {
    ui::ShowError(owner, win32::FormatDialogError(hr));
  }
  return SUCCEEDED(hr);
}

std::wstring FileNameOf(const std::wstring& path) {
  const size_t slash = path.find_last_of(L"\\/");
  std::wstring name = slash == std::wstring::npos ? path : path.substr(slash + 1);
  const size_t dot = name.find_last_of(L'.');
  if (dot != std::wstring::npos && dot > 0) {
    name.erase(dot);
  }
  return name;
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam,
                            LPARAM lparam) {
  auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
  if (message == WM_INITDIALOG) {
    state = reinterpret_cast<State*>(lparam);
    SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetDlgItemTextW(dialog, IDC_LOAD_HIVE_PATH, state->value.file.c_str());
    SetDlgItemTextW(dialog, IDC_LOAD_HIVE_NAME, state->value.key_name.c_str());
    const bool users = state->value.root == HKEY_USERS;
    CheckDlgButton(dialog, IDC_LOAD_HIVE_HKLM, users ? BST_UNCHECKED : BST_CHECKED);
    CheckDlgButton(dialog, IDC_LOAD_HIVE_HKU, users ? BST_CHECKED : BST_UNCHECKED);
    dialog_support::Initialize(dialog, &state->font,
                               {IDC_LOAD_HIVE_PATH, IDC_LOAD_HIVE_NAME});
    return TRUE;
  }
  if (message == WM_DESTROY) {
    if (state) {
      dialog_support::ReleaseFont(&state->font);
    }
    return TRUE;
  }
  INT_PTR themed = 0;
  if (dialog_support::HandleThemeMessage(dialog, message, wparam, lparam, &themed)) {
    return themed;
  }
  if (message != WM_COMMAND || !state) {
    return FALSE;
  }
  const int id = LOWORD(wparam);
  if (id == IDC_LOAD_HIVE_BROWSE && HIWORD(wparam) == BN_CLICKED) {
    std::wstring path = dialog_support::ReadText(dialog, IDC_LOAD_HIVE_PATH);
    if (ChooseFile(dialog, &path)) {
      SetDlgItemTextW(dialog, IDC_LOAD_HIVE_PATH, path.c_str());
      if (dialog_support::ReadText(dialog, IDC_LOAD_HIVE_NAME).empty()) {
        SetDlgItemTextW(dialog, IDC_LOAD_HIVE_NAME, FileNameOf(path).c_str());
      }
    }
    return TRUE;
  }
  if (id == IDOK) {
    state->value.file = dialog_support::ReadText(dialog, IDC_LOAD_HIVE_PATH);
    if (state->value.file.empty()) {
      ui::ShowError(dialog, L"Select a hive file.");
      return TRUE;
    }
    state->value.key_name = dialog_support::ReadText(dialog, IDC_LOAD_HIVE_NAME);
    if (state->value.key_name.empty()) {
      ui::ShowError(dialog, L"Enter a key name for the loaded hive.");
      return TRUE;
    }
    if (state->value.key_name.find(L'\\') != std::wstring::npos) {
      ui::ShowError(dialog, L"The key name cannot contain a backslash.");
      return TRUE;
    }
    state->value.root =
        IsDlgButtonChecked(dialog, IDC_LOAD_HIVE_HKU) == BST_CHECKED
            ? HKEY_USERS
            : HKEY_LOCAL_MACHINE;
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

bool ChooseHiveToLoad(HWND owner, LoadHiveResult* result) {
  if (!result) {
    return false;
  }
  State state;
  state.value = *result;
  const INT_PTR dialog_result = DialogBoxParamW(
      GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_LOAD_HIVE), owner,
      DialogProc, reinterpret_cast<LPARAM>(&state));
  if (dialog_result != IDOK || !state.accepted) {
    return false;
  }
  *result = std::move(state.value);
  return true;
}

} // namespace regkit::editors
