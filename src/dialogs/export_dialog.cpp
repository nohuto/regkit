// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/export_dialog.h"

#include "ui/dialog_support.h"
#include "ui/feedback.h"
#include "win32/file_dialog.h"
#include "win32/shell_paths.h"
#include "win32/translation.h"

#include "resource.h"

#include <win32ui/dialog/dialog_window.h>

#include <algorithm>
#include <utility>

namespace regkit::editors
{

namespace
{

void ApplyFormat(HWND dialog, bool hive)
{
    // hive exports always contain the full branch and have no associated viewer
    EnableWindow(GetDlgItem(dialog, IDC_EXPORT_RANGE_KEY), !hive);
    EnableWindow(GetDlgItem(dialog, IDC_EXPORT_OPEN_AFTER), !hive);
    if (hive)
    {
        CheckRadioButton(dialog, IDC_EXPORT_RANGE_BRANCH, IDC_EXPORT_RANGE_KEY, IDC_EXPORT_RANGE_BRANCH);
    }
    std::wstring path = dialog_support::ReadText(dialog, IDC_EXPORT_PATH);
    const std::wstring_view from = hive ? L".reg" : L".hiv";
    if (path.size() > from.size() && util::EqualsInsensitive(std::wstring_view(path).substr(path.size() - from.size()), from))
    {
        path.replace(path.size() - from.size(), from.size(), hive ? L".hiv" : L".reg");
        SetDlgItemTextW(dialog, IDC_EXPORT_PATH, path.c_str());
    }
}

bool HiveSelected(HWND dialog)
{
    return SendDlgItemMessageW(dialog, IDC_EXPORT_FORMAT, CB_GETCURSEL, 0, 0) == 1;
}

struct State
{
    ExportResult value;
    std::wstring confirmed_path;
    bool accepted = false;
};

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<State*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        SetDlgItemTextW(dialog, IDC_EXPORT_PATH, state->value.path.c_str());
        CheckDlgButton(dialog, IDC_EXPORT_RANGE_BRANCH, state->value.include_subkeys ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(dialog, IDC_EXPORT_RANGE_KEY, state->value.include_subkeys ? BST_UNCHECKED : BST_CHECKED);
        for (const wchar_t* item : {util::Tr(L"Don't open"), util::Tr(L"In text editor"), util::Tr(L"In RegKit")})
        {
            SendDlgItemMessageW(dialog, IDC_EXPORT_OPEN_AFTER, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
        }
        SendDlgItemMessageW(dialog, IDC_EXPORT_OPEN_AFTER, CB_SETCURSEL, static_cast<WPARAM>(state->value.open_after), 0);
        SendDlgItemMessageW(dialog, IDC_EXPORT_FORMAT, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(util::Tr(L"Registry File (.reg)")));
        if (state->value.allow_hive)
        {
            SendDlgItemMessageW(dialog, IDC_EXPORT_FORMAT, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(util::Tr(L"Registry Hive File")));
        }
        SendDlgItemMessageW(dialog, IDC_EXPORT_FORMAT, CB_SETCURSEL, 0, 0);
        EnableWindow(GetDlgItem(dialog, IDC_EXPORT_FORMAT), state->value.allow_hive);
        return TRUE;
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    const int id = LOWORD(wparam);
    if (id == IDC_EXPORT_FORMAT && HIWORD(wparam) == CBN_SELCHANGE)
    {
        ApplyFormat(dialog, HiveSelected(dialog));
        return TRUE;
    }
    if (id == IDC_EXPORT_BROWSE && HIWORD(wparam) == BN_CLICKED)
    {
        std::wstring path = dialog_support::ReadText(dialog, IDC_EXPORT_PATH);
        if (ui::ReportFileDialogResult(dialog, win32::ChooseFileToSave(dialog, HiveSelected(dialog) ? ui::kHiveFileFilter : ui::kRegFileFilter, path.empty() ? nullptr : path.c_str(), &path)))
        {
            state->confirmed_path = path;
            SetDlgItemTextW(dialog, IDC_EXPORT_PATH, path.c_str());
        }
        return TRUE;
    }
    if (id == IDOK)
    {
        state->value.hive = HiveSelected(dialog);
        state->value.path = dialog_support::ReadText(dialog, IDC_EXPORT_PATH);
        if (!state->value.hive)
        {
            state->value.path = util::EnsureFileExtension(state->value.path, L".reg");
        }
        if (state->value.path.empty())
        {
            ui::ShowError(dialog, util::Tr(L"Select a destination file."));
            return TRUE;
        }
        if (!ui::ConfirmOverwrite(dialog, state->value.path, state->confirmed_path))
        {
            return TRUE;
        }
        state->value.include_subkeys = IsDlgButtonChecked(dialog, IDC_EXPORT_RANGE_BRANCH) == BST_CHECKED;
        state->value.open_after = static_cast<win32::OpenAfter>(std::max<LRESULT>(0, SendDlgItemMessageW(dialog, IDC_EXPORT_OPEN_AFTER, CB_GETCURSEL, 0, 0)));
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

bool ChooseExport(HWND owner, const ExportRequest& request, ExportResult* result)
{
    if (!result)
    {
        return false;
    }
    State state;
    state.value = request;
    const INT_PTR dialog_result = win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_EXPORT_OPTIONS), DialogProc, reinterpret_cast<LPARAM>(&state));
    if (dialog_result != IDOK || !state.accepted)
    {
        return false;
    }
    *result = std::move(state.value);
    return true;
}

}
