// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/replace_dialog.h"

#include "dialogs/query_dialog.h"
#include "resource.h"
#include "ui/autocomplete.h"
#include "ui/dialog_support.h"
#include "ui/feedback.h"
#include "win32/translation.h"
#include "workspace/dialog_state.h"

#include <win32ui/dialog/autocomplete.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/tooltip.h>
#include <win32ui/theme/theme.h>

namespace regkit
{

namespace
{

namespace dialog_support = editors::dialog_support;

struct State
{
    ReplaceDialogResult* out = nullptr;
};

constexpr struct
{
    int id;
    const wchar_t* key;
    bool ReplaceDialogResult::* field;
} kChecks[] = {
    {IDC_REPLACE_RECURSIVE, L"replace_recursive", &ReplaceDialogResult::recursive},
    {IDC_REPLACE_CASE, L"replace_match_case", &ReplaceDialogResult::match_case},
    {IDC_REPLACE_WHOLE, L"replace_match_whole", &ReplaceDialogResult::match_whole},
    {IDC_REPLACE_REGEX, L"replace_regex", &ReplaceDialogResult::use_regex},
    {IDC_REPLACE_KEYS, L"replace_keys", &ReplaceDialogResult::replace_keys},
    {IDC_REPLACE_VALUES, L"replace_values", &ReplaceDialogResult::replace_values},
    {IDC_REPLACE_DATA, L"replace_data", &ReplaceDialogResult::replace_data},
    {IDC_REPLACE_DECIMAL, L"replace_decimal", &ReplaceDialogResult::number_decimal},
    {IDC_REPLACE_HEX, L"replace_hex", &ReplaceDialogResult::number_hex},
};

void UpdateValueDataOptions(HWND dialog)
{
    const bool enabled = IsDlgButtonChecked(dialog, IDC_REPLACE_DATA) == BST_CHECKED;
    for (const int id : {IDC_REPLACE_VALUE_GROUP, IDC_REPLACE_DECIMAL, IDC_REPLACE_HEX})
    {
        EnableWindow(GetDlgItem(dialog, id), enabled);
    }
}

bool Accept(HWND dialog, State* state)
{
    ReplaceDialogResult result;
    result.find_text = dialog_support::ReadText(dialog, IDC_REPLACE_FIND);
    if (result.find_text.empty())
    {
        ui::ShowWarning(dialog, util::Tr(L"Enter text to find."));
        return false;
    }
    result.replace_text = dialog_support::ReadText(dialog, IDC_REPLACE_WITH);
    result.start_key = dialog_support::ReadText(dialog, IDC_REPLACE_KEY);
    for (const auto& check : kChecks)
    {
        result.*check.field = IsDlgButtonChecked(dialog, check.id) == BST_CHECKED;
    }
    if (!result.replace_keys && !result.replace_values && !result.replace_data)
    {
        ui::ShowWarning(dialog, util::Tr(L"Select what should be replaced."));
        return false;
    }
    *state->out = std::move(result);
    return true;
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<State*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        const ReplaceDialogResult& initial = *state->out;
        SetDlgItemTextW(dialog, IDC_REPLACE_FIND, initial.find_text.c_str());
        SetDlgItemTextW(dialog, IDC_REPLACE_WITH, initial.replace_text.c_str());
        SetDlgItemTextW(dialog, IDC_REPLACE_KEY, initial.start_key.c_str());
        for (const auto& check : kChecks)
        {
            CheckDlgButton(dialog, check.id, initial.*check.field ? BST_CHECKED : BST_UNCHECKED);
        }
        UpdateValueDataOptions(dialog);
        win32ui::AttachAutoComplete(GetDlgItem(dialog, IDC_REPLACE_KEY), appearance::SuggestKeys);
        win32ui::ApplyThemeToWindow(win32ui::AddTooltip(
            dialog,
            GetDlgItem(dialog, IDC_REPLACE_REGEX),
            util::Tr(L"PCRE syntax: ^ $ anchors, character classes, greedy, lazy (*?) and possessive (*+) quantifiers,\n"
                     L"(?<name>...) groups, lookaround (?=...) (?<=...), backreferences \\1 and Unicode classes \\p{L}, \\w, "
                     L"\\X.\n"
                     L"Replace with: $1 or ${1} for a group, $<name> or ${name} for a named group, $& for the whole match, $$ "
                     L"for a dollar.")
        ));
        return TRUE;
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    switch (LOWORD(wparam))
    {
    case IDC_REPLACE_DATA:
        UpdateValueDataOptions(dialog);
        return TRUE;
    case IDC_REPLACE_BROWSE:
        {
            std::wstring selected;
            if (ShowBrowseKeyDialog(dialog, &selected) && !selected.empty())
            {
                SetDlgItemTextW(dialog, IDC_REPLACE_KEY, selected.c_str());
            }
            return TRUE;
        }
    case IDOK:
        if (Accept(dialog, state))
        {
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

bool ShowReplaceDialog(HWND owner, ReplaceDialogResult* result)
{
    State state;
    state.out = result;
    return result && win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_REPLACE), DialogProc, reinterpret_cast<LPARAM>(&state)) == IDOK;
}

void ReplaceDialogFields(workspace::DialogFields& fields, ReplaceDialogResult* result)
{
    for (const auto& check : kChecks)
    {
        fields.Field(check.key, &(result->*check.field));
    }
    fields.Field(L"replace_find", &result->find_text);
    fields.Field(L"replace_with", &result->replace_text);
    fields.Field(L"replace_start_key", &result->start_key);
}

}
