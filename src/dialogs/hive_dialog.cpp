// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/hive_dialog.h"

#include "ui/autocomplete.h"
#include "ui/feedback.h"
#include "win32/translation.h"

#include "resource.h"

#include <win32ui/dialog/autocomplete.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/file_dialog.h>
#include <win32ui/system/paths.h>
#include <win32ui/window.h>

#include <utility>

namespace regkit::editors
{

namespace
{

struct State
{
    LoadHiveResult value;
    bool accepted = false;
};

bool ChooseFile(HWND owner, std::wstring* path)
{
    return ui::ReportFileDialogResult(owner, win32ui::ChooseFileToOpen(owner, {.filter = ui::kHiveFileFilter}, path));
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<State*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        SetDlgItemTextW(dialog, IDC_LOAD_HIVE_PATH, state->value.file.c_str());
        SetDlgItemTextW(dialog, IDC_LOAD_HIVE_NAME, state->value.key_name.c_str());
        const bool users = state->value.root == HKEY_USERS;
        CheckDlgButton(dialog, IDC_LOAD_HIVE_HKLM, users ? BST_UNCHECKED : BST_CHECKED);
        CheckDlgButton(dialog, IDC_LOAD_HIVE_HKU, users ? BST_CHECKED : BST_UNCHECKED);
        return TRUE;
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    const int id = LOWORD(wparam);
    if (id == IDC_LOAD_HIVE_BROWSE && HIWORD(wparam) == BN_CLICKED)
    {
        std::wstring path = win32ui::DialogItemText(dialog, IDC_LOAD_HIVE_PATH);
        if (ChooseFile(dialog, &path))
        {
            SetDlgItemTextW(dialog, IDC_LOAD_HIVE_PATH, path.c_str());
            if (win32ui::DialogItemText(dialog, IDC_LOAD_HIVE_NAME).empty())
            {
                SetDlgItemTextW(dialog, IDC_LOAD_HIVE_NAME, win32ui::FileBaseName(path).c_str());
            }
        }
        return TRUE;
    }
    if (id == IDOK)
    {
        state->value.file = win32ui::DialogItemText(dialog, IDC_LOAD_HIVE_PATH);
        if (state->value.file.empty())
        {
            ui::ShowError(dialog, util::Tr(L"Select a hive file."));
            return TRUE;
        }
        state->value.key_name = win32ui::DialogItemText(dialog, IDC_LOAD_HIVE_NAME);
        if (state->value.key_name.empty())
        {
            ui::ShowError(dialog, util::Tr(L"Enter a key name for the loaded hive."));
            return TRUE;
        }
        if (state->value.key_name.find(L'\\') != std::wstring::npos)
        {
            ui::ShowError(dialog, util::Tr(L"The key name can't contain a backslash."));
            return TRUE;
        }
        state->value.root =
            IsDlgButtonChecked(dialog, IDC_LOAD_HIVE_HKU) == BST_CHECKED ? HKEY_USERS : HKEY_LOCAL_MACHINE;
        state->accepted = true;
        win32ui::EndDialogWindow(dialog, IDOK);
        return TRUE;
    }
    if (id == IDCANCEL)
    {
        win32ui::EndDialogWindow(dialog, IDCANCEL);
        return TRUE;
    }
    return FALSE;
}

}

bool ChooseHiveToLoad(HWND owner, LoadHiveResult* result)
{
    if (!result)
    {
        return false;
    }
    State state;
    state.value = *result;
    const INT_PTR dialog_result = win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_LOAD_HIVE), DialogProc, reinterpret_cast<LPARAM>(&state));
    if (dialog_result != IDOK || !state.accepted)
    {
        return false;
    }
    *result = std::move(state.value);
    return true;
}

namespace
{

struct SymbolicLinkDialogState
{
    SymbolicLinkResult* result = nullptr;
    const BrowseKeyCallback* browse = nullptr;
};

INT_PTR CALLBACK SymbolicLinkDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    auto* dialog = reinterpret_cast<SymbolicLinkDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    SymbolicLinkResult* state = dialog ? dialog->result : nullptr;
    switch (msg)
    {
    case WM_INITDIALOG:
        {
            dialog = reinterpret_cast<SymbolicLinkDialogState*>(lparam);
            state = dialog ? dialog->result : nullptr;
            SetWindowLongPtrW(dlg, DWLP_USER, static_cast<LONG_PTR>(lparam));
            if (state)
            {
                SetDlgItemTextW(dlg, IDC_SYMLINK_NAME, state->name.c_str());
                SetDlgItemTextW(dlg, IDC_SYMLINK_TARGET, state->target.c_str());
            }
            win32ui::AttachAutoComplete(GetDlgItem(dlg, IDC_SYMLINK_TARGET), appearance::SuggestKeys);
            return TRUE;
        }
    case WM_COMMAND:
        switch (LOWORD(wparam))
        {
        case IDOK:
            {
                if (state)
                {
                    state->name = win32ui::DialogItemText(dlg, IDC_SYMLINK_NAME);
                    state->target = win32ui::DialogItemText(dlg, IDC_SYMLINK_TARGET);
                    if (state->name.empty() || state->target.empty())
                    {
                        ui::ShowWarning(dlg, util::Tr(L"Enter a link name and a target key."));
                        return TRUE;
                    }
                    if (state->name.find(L'\\') != std::wstring::npos)
                    {
                        ui::ShowWarning(dlg, util::Tr(L"The link name can't contain a backslash."));
                        return TRUE;
                    }
                }
                win32ui::EndDialogWindow(dlg, IDOK);
                return TRUE;
            }
        case IDC_SYMLINK_BROWSE:
            {
                std::wstring selected;
                if (dialog && dialog->browse && (*dialog->browse)(dlg, &selected) && !selected.empty())
                {
                    SetDlgItemTextW(dlg, IDC_SYMLINK_TARGET, selected.c_str());
                }
                return TRUE;
            }
        case IDCANCEL:
            win32ui::EndDialogWindow(dlg, IDCANCEL);
            return TRUE;
        default:
            break;
        }
        break;
    default:
        break;
    }
    return FALSE;
}

}

bool PromptSymbolicLink(HWND owner, const std::wstring& suggested_name, const BrowseKeyCallback& browse, SymbolicLinkResult* result)
{
    if (!result)
    {
        return false;
    }
    result->name = suggested_name;
    SymbolicLinkDialogState dialog;
    dialog.result = result;
    dialog.browse = &browse;
    return win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_NEW_SYMLINK), SymbolicLinkDialogProc, reinterpret_cast<LPARAM>(&dialog)) == IDOK;
}

}
