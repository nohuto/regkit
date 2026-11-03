// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/comment_editor.h"

#include "ui/autocomplete.h"
#include "ui/dialog_layout.h"
#include "win32/translation.h"

#include "resource.h"

#include <win32ui/dialog/autocomplete.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/edit.h>
#include <win32ui/system/strings.h>
#include <win32ui/window.h>

#include <utility>

namespace regkit::editors
{

namespace
{

struct State
{
    const CommentRequest* request = nullptr;
    CommentResult value;
    bool accepted = false;
};

bool Checked(HWND dialog, int id)
{
    return IsDlgButtonChecked(dialog, id) == BST_CHECKED;
}

CommentScope ReadScope(HWND dialog)
{
    CommentScope scope;
    scope.rule = Checked(dialog, IDC_COMMENT_RULE);
    scope.same_type = Checked(dialog, IDC_COMMENT_TYPE);
    scope.same_size = Checked(dialog, IDC_COMMENT_SIZE);
    scope.in_key = Checked(dialog, IDC_COMMENT_KEY);
    scope.include_subkeys = scope.in_key && Checked(dialog, IDC_COMMENT_SUBKEYS);
    scope.key_path = win32ui::DialogItemText(dialog, IDC_COMMENT_KEY_PATH);
    return scope;
}

void UpdateControls(HWND dialog)
{
    const CommentScope scope = ReadScope(dialog);
    for (const int id : {IDC_COMMENT_TYPE, IDC_COMMENT_SIZE, IDC_COMMENT_KEY})
    {
        EnableWindow(GetDlgItem(dialog, id), scope.rule);
    }
    EnableWindow(GetDlgItem(dialog, IDC_COMMENT_KEY_PATH), scope.rule && scope.in_key);
    EnableWindow(GetDlgItem(dialog, IDC_COMMENT_SUBKEYS), scope.rule && scope.in_key);
}

void InitControls(HWND dialog, const State* state)
{
    const CommentRequest& request = *state->request;
    SetDlgItemTextW(dialog, IDC_EDIT, win32ui::EditLineBreaks(request.text).c_str());
    CheckRadioButton(dialog, IDC_COMMENT_VALUE, IDC_COMMENT_RULE, request.scope.rule ? IDC_COMMENT_RULE : IDC_COMMENT_VALUE);
    if (request.multiple)
    {
        SetDlgItemTextW(dialog, IDC_COMMENT_VALUE, util::Tr(L"Selected values only"));
    }
    SetDlgItemTextW(
        dialog,
        IDC_COMMENT_RULE,
        (request.multiple ? std::wstring(util::Tr(L"Values with the selected names")) : util::Tr(L"Values named") + std::wstring(L" ") + request.name).c_str()
    );
    CheckDlgButton(dialog, IDC_COMMENT_TYPE, request.scope.same_type ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dialog, IDC_COMMENT_SIZE, request.scope.same_size ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dialog, IDC_COMMENT_KEY, request.scope.in_key ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dialog, IDC_COMMENT_SUBKEYS, request.scope.include_subkeys ? BST_CHECKED : BST_UNCHECKED);
    SetDlgItemTextW(dialog, IDC_COMMENT_TYPE_TEXT, request.type.c_str());
    SetDlgItemTextW(dialog, IDC_COMMENT_SIZE_TEXT, request.size.c_str());
    SetDlgItemTextW(dialog, IDC_COMMENT_KEY_PATH, request.scope.key_path.c_str());
    ShowWindow(GetDlgItem(dialog, IDC_COMMENT_RESTORE), request.can_restore ? SW_SHOW : SW_HIDE);
    if (!request.key)
    {
        UpdateControls(dialog);
    }
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<State*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        InitControls(dialog, state);
        win32ui::FitDialogText(dialog);
        win32ui::AttachAutoComplete(GetDlgItem(dialog, IDC_COMMENT_KEY_PATH), appearance::SuggestKeys);
        win32ui::AllowLineBreaks(GetDlgItem(dialog, IDC_EDIT));
        win32ui::AnchorDialogControls(
            dialog,
            {
                {IDC_LABEL, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight},
                {IDC_EDIT, win32ui::kAnchorLeft | win32ui::kAnchorTop | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                {IDC_COMMENT_APPLY_GROUP, win32ui::kAnchorLeft | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                {IDC_COMMENT_VALUE, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                {IDC_COMMENT_RULE, win32ui::kAnchorLeft | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                {IDC_COMMENT_MATCH_GROUP, win32ui::kAnchorLeft | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                {IDC_COMMENT_TYPE, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                {IDC_COMMENT_TYPE_TEXT, win32ui::kAnchorLeft | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                {IDC_COMMENT_SIZE, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                {IDC_COMMENT_SIZE_TEXT, win32ui::kAnchorLeft | win32ui::kAnchorRight | win32ui::kAnchorBottom},
                {IDC_COMMENT_KEY, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                {IDC_COMMENT_KEY_PATH, win32ui::kAnchorLeft | win32ui::kAnchorRight | (state->request->key ? win32ui::kAnchorTop : win32ui::kAnchorBottom)},
                {IDC_COMMENT_SUBKEYS, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                {IDC_COMMENT_RESTORE, win32ui::kAnchorLeft | win32ui::kAnchorBottom},
                {IDOK, win32ui::kAnchorRight | win32ui::kAnchorBottom},
                {IDCANCEL, win32ui::kAnchorRight | win32ui::kAnchorBottom},
            }
        );
        return TRUE;
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    switch (LOWORD(wparam))
    {
    case IDOK:
    case IDC_COMMENT_RESTORE:
        state->value.text = win32ui::PlainLineBreaks(win32ui::DialogItemText(dialog, IDC_EDIT));
        state->value.scope = ReadScope(dialog);
        state->value.restore_default = LOWORD(wparam) == IDC_COMMENT_RESTORE;
        state->accepted = true;
        win32ui::EndDialogWindow(dialog, IDOK);
        return TRUE;
    case IDCANCEL:
        win32ui::EndDialogWindow(dialog, IDCANCEL);
        return TRUE;
    case IDC_COMMENT_VALUE:
    case IDC_COMMENT_RULE:
    case IDC_COMMENT_TYPE:
    case IDC_COMMENT_SIZE:
    case IDC_COMMENT_KEY:
        UpdateControls(dialog);
        return TRUE;
    default:
        return FALSE;
    }
}

}

bool EditComment(HWND owner, const CommentRequest& request, CommentResult* result)
{
    State state;
    state.request = &request;
    const INT_PTR dialog_result =
        win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(request.key ? IDD_KEY_COMMENT : IDD_COMMENT), DialogProc, reinterpret_cast<LPARAM>(&state));
    if (dialog_result != IDOK || !state.accepted)
    {
        return false;
    }
    *result = std::move(state.value);
    return true;
}

}
