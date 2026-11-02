// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/tools/convert_dialog.h"

#include "ui/autocomplete.h"
#include "ui/dialog_support.h"
#include "ui/feedback.h"
#include "win32/file_dialog.h"
#include "win32/text_transform.h"
#include "win32/translation.h"
#include "workspace/dialog_state.h"

#include "resource.h"

#include <win32ui/dialog/autocomplete.h>
#include <win32ui/dialog/dialog_layout.h>
#include <win32ui/dialog/dialog_window.h>
#include <win32ui/dialog/file_dialog.h>
#include <win32ui/system/paths.h>
#include <win32ui/system/shell_open.h>

#include <algorithm>
#include <initializer_list>
#include <set>

namespace regkit
{
namespace
{

namespace dialog_support = editors::dialog_support;

constexpr const wchar_t* kFormatFilters[] = {
    L"Registry Files (*.reg)\0*.reg\0",
    L"Batch Files (*.bat;*.cmd)\0*.bat;*.cmd\0",
    L"PowerShell Scripts (*.ps1)\0*.ps1\0",
};

struct State
{
    ConvertSettings* settings = nullptr;
    std::wstring confirmed_output;
};

int Selection(HWND dialog, int id)
{
    return std::max(0, static_cast<int>(SendDlgItemMessageW(dialog, id, CB_GETCURSEL, 0, 0)));
}

ConvertSource SourceAt(HWND dialog)
{
    return static_cast<ConvertSource>(Selection(dialog, IDC_CONVERT_SOURCE));
}

regfile::Format FormatAt(HWND dialog)
{
    return static_cast<regfile::Format>(Selection(dialog, IDC_CONVERT_FORMAT));
}

regfile::Format SourceFormat(ConvertSource source)
{
    return static_cast<regfile::Format>(static_cast<int>(source) - 1);
}

std::wstring WithExtension(std::wstring path, regfile::Format format)
{
    regfile::Format current = regfile::Format::kReg;
    if (path.empty() || (regfile::FormatFromPath(path, &current) && current == format))
    {
        return path;
    }
    return win32ui::ReplaceFileExtension(std::move(path), regfile::FormatExtension(format));
}

void UpdateInputControls(HWND dialog)
{
    const bool file = SourceAt(dialog) != ConvertSource::kRegistry;
    EnableWindow(GetDlgItem(dialog, IDC_CONVERT_INPUT), file);
    EnableWindow(GetDlgItem(dialog, IDC_CONVERT_INPUT_BROWSE), file);
}

void PopulateKeys(HWND dialog)
{
    HWND combo = GetDlgItem(dialog, IDC_CONVERT_KEY);
    const std::wstring current = dialog_support::ReadText(dialog, IDC_CONVERT_KEY);
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    const std::wstring path = dialog_support::ReadText(dialog, IDC_CONVERT_INPUT);
    std::vector<regfile::Operation> operations;
    std::wstring error;
    if (SourceAt(dialog) != ConvertSource::kRegistry && !path.empty() &&
        regfile::ReadOperations(path, SourceFormat(SourceAt(dialog)), &operations, &error))
    {
        auto less = [](const std::wstring& left, const std::wstring& right) { return util::CompareInsensitive(left, right) < 0; };
        std::set<std::wstring, decltype(less)> keys(less);
        for (const regfile::Operation& operation : operations)
        {
            keys.insert(operation.path);
        }
        for (const std::wstring& key : keys)
        {
            SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(key.c_str()));
        }
    }
    SetWindowTextW(combo, current.c_str());
}

bool Convert(HWND dialog, State* state)
{
    ConvertSettings& settings = *state->settings;
    settings.source = SourceAt(dialog);
    settings.input_path = util::TrimWhitespace(dialog_support::ReadText(dialog, IDC_CONVERT_INPUT));
    settings.key_path = util::TrimWhitespace(dialog_support::ReadText(dialog, IDC_CONVERT_KEY));
    settings.recursive = IsDlgButtonChecked(dialog, IDC_CONVERT_RECURSIVE) == BST_CHECKED;
    settings.format = FormatAt(dialog);
    settings.output_path = WithExtension(util::TrimWhitespace(dialog_support::ReadText(dialog, IDC_CONVERT_OUTPUT)), settings.format);
    settings.admin_check = IsDlgButtonChecked(dialog, IDC_CONVERT_ADMIN) == BST_CHECKED;
    settings.open_in_editor = IsDlgButtonChecked(dialog, IDC_CONVERT_OPEN) == BST_CHECKED;
    SetDlgItemTextW(dialog, IDC_CONVERT_OUTPUT, settings.output_path.c_str());

    const bool registry = settings.source == ConvertSource::kRegistry;
    if (registry ? settings.key_path.empty() : settings.input_path.empty())
    {
        ui::ShowError(dialog, registry ? util::Tr(L"Registry key is required.") : util::Tr(L"Input file path is required."));
        return false;
    }
    if (settings.output_path.empty())
    {
        ui::ShowError(dialog, util::Tr(L"Output file path is required."));
        return false;
    }
    if (!registry && util::EqualsInsensitive(settings.input_path, settings.output_path))
    {
        ui::ShowError(dialog, util::Tr(L"The output file can't be the input file."));
        return false;
    }

    std::vector<regfile::Operation> operations;
    std::wstring error;
    std::vector<std::wstring> skipped;
    HCURSOR cursor = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    const bool read = registry ? regfile::ReadRegistry(settings.key_path, settings.recursive, &operations, &error, &skipped)
                               : regfile::ReadOperations(settings.input_path, SourceFormat(settings.source), &operations, &error) &&
                                     (settings.key_path.empty() || regfile::SelectKey(settings.key_path, settings.recursive, &operations, &error));
    const std::wstring text = read ? regfile::RenderOperations(settings.format, operations, settings.admin_check, &skipped) : std::wstring();
    SetCursor(cursor);
    if (!read)
    {
        ui::ShowError(dialog, error);
        return false;
    }
    if (!ui::ConfirmOverwrite(dialog, settings.output_path, state->confirmed_output))
    {
        return false;
    }
    if (!skipped.empty() && !ui::ConfirmConversionSkips(dialog, skipped))
    {
        return false;
    }
    if (!regfile::SaveRendered(settings.output_path, settings.format, text))
    {
        ui::ShowError(dialog, util::TrDetail(L"The output file couldn't be written.", settings.output_path));
        return false;
    }
    if (settings.open_in_editor)
    {
        ui::ReportFileDialogResult(dialog, win32ui::ShellOpen(GetWindow(dialog, GW_OWNER), settings.output_path.c_str(), L".txt"));
    }
    else
    {
        ui::ShowConversionSucceeded(dialog, settings.output_path);
    }
    return true;
}

void Browse(HWND dialog, State* state, bool input)
{
    std::wstring path = dialog_support::ReadText(dialog, input ? IDC_CONVERT_INPUT : IDC_CONVERT_OUTPUT);
    const DWORD error = input ? win32ui::ChooseFileToOpen(dialog, {.filter = regfile::kConvertFilter}, &path)
                              : win32::ChooseFileToSave(dialog, kFormatFilters[static_cast<int>(FormatAt(dialog))], path.empty() ? nullptr : path.c_str(), &path);
    if (!ui::ReportFileDialogResult(dialog, error))
    {
        return;
    }
    regfile::Format format = regfile::Format::kReg;
    const bool known = regfile::FormatFromPath(path, &format);
    SetDlgItemTextW(dialog, input ? IDC_CONVERT_INPUT : IDC_CONVERT_OUTPUT, path.c_str());
    if (!input)
    {
        state->confirmed_output = path;
        SendDlgItemMessageW(dialog, IDC_CONVERT_FORMAT, CB_SETCURSEL, known ? static_cast<int>(format) : Selection(dialog, IDC_CONVERT_FORMAT), 0);
        return;
    }
    if (known)
    {
        SendDlgItemMessageW(dialog, IDC_CONVERT_SOURCE, CB_SETCURSEL, static_cast<int>(format) + 1, 0);
    }
    SetDlgItemTextW(dialog, IDC_CONVERT_KEY, L"");
    PopulateKeys(dialog);
    if (FormatAt(dialog) != SourceFormat(SourceAt(dialog)))
    {
        SetDlgItemTextW(dialog, IDC_CONVERT_OUTPUT, WithExtension(path, FormatAt(dialog)).c_str());
    }
}

void Fill(HWND dialog, int id, std::initializer_list<const wchar_t*> items, int selected)
{
    for (const wchar_t* item : items)
    {
        SendDlgItemMessageW(dialog, id, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item));
    }
    SendDlgItemMessageW(dialog, id, CB_SETCURSEL, selected, 0);
}

INT_PTR CALLBACK DialogProc(HWND dialog, UINT message, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<State*>(GetWindowLongPtrW(dialog, DWLP_USER));
    if (message == WM_INITDIALOG)
    {
        state = reinterpret_cast<State*>(lparam);
        SetWindowLongPtrW(dialog, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
        const ConvertSettings& settings = *state->settings;
        Fill(dialog, IDC_CONVERT_SOURCE, {util::Tr(L"Registry"), util::Tr(L"Reg File"), util::Tr(L"Batch File"), util::Tr(L"PowerShell Script")}, static_cast<int>(settings.source));
        Fill(dialog, IDC_CONVERT_FORMAT, {util::Tr(L"Reg File"), util::Tr(L"Batch File"), util::Tr(L"PowerShell Script")}, static_cast<int>(settings.format));
        SetDlgItemTextW(dialog, IDC_CONVERT_INPUT, settings.input_path.c_str());
        SetDlgItemTextW(dialog, IDC_CONVERT_OUTPUT, settings.output_path.c_str());
        CheckDlgButton(dialog, IDC_CONVERT_RECURSIVE, settings.recursive ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(dialog, IDC_CONVERT_ADMIN, settings.admin_check ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(dialog, IDC_CONVERT_OPEN, settings.open_in_editor ? BST_CHECKED : BST_UNCHECKED);
        SetDlgItemTextW(dialog, IDC_CONVERT_KEY, settings.key_path.c_str());
        PopulateKeys(dialog);
        UpdateInputControls(dialog);
        HWND key = GetDlgItem(dialog, IDC_CONVERT_KEY);
        win32ui::AttachAutoComplete(key, [dialog, key](const std::wstring& text) {
            return SourceAt(dialog) == ConvertSource::kRegistry ? appearance::SuggestKeys(text) : win32ui::SuggestComboPaths(key, text);
        });
        win32ui::MatchComboHeights(GetDlgItem(dialog, IDC_CONVERT_INPUT), {GetDlgItem(dialog, IDC_CONVERT_SOURCE), GetDlgItem(dialog, IDC_CONVERT_KEY), GetDlgItem(dialog, IDC_CONVERT_FORMAT)});
        return TRUE;
    }
    if (message != WM_COMMAND || !state)
    {
        return FALSE;
    }
    const int id = LOWORD(wparam);
    const int code = HIWORD(wparam);
    if (code == EN_KILLFOCUS && id == IDC_CONVERT_INPUT)
    {
        PopulateKeys(dialog);
    }
    else if (code == CBN_SELCHANGE && id == IDC_CONVERT_SOURCE)
    {
        SetDlgItemTextW(dialog, IDC_CONVERT_KEY, L"");
        UpdateInputControls(dialog);
        PopulateKeys(dialog);
    }
    else if (code == CBN_SELCHANGE && id == IDC_CONVERT_FORMAT)
    {
        SetDlgItemTextW(dialog, IDC_CONVERT_OUTPUT, WithExtension(dialog_support::ReadText(dialog, IDC_CONVERT_OUTPUT), FormatAt(dialog)).c_str());
    }
    else if (code == BN_CLICKED && (id == IDC_CONVERT_INPUT_BROWSE || id == IDC_CONVERT_OUTPUT_BROWSE))
    {
        Browse(dialog, state, id == IDC_CONVERT_INPUT_BROWSE);
    }
    else if (id == IDOK)
    {
        if (Convert(dialog, state))
        {
            win32ui::EndDialogWindow(dialog, IDOK);
        }
    }
    else if (id == IDCANCEL)
    {
        win32ui::EndDialogWindow(dialog, IDCANCEL);
    }
    else
    {
        return FALSE;
    }
    return TRUE;
}

}

void ShowConvertDialog(HWND owner, ConvertSettings* settings)
{
    State state;
    state.settings = settings;
    win32ui::RunProcDialog(owner, MAKEINTRESOURCEW(IDD_CONVERT), DialogProc, reinterpret_cast<LPARAM>(&state));
}

void ConvertDialogFields(workspace::DialogFields& fields, ConvertSettings* settings)
{
    fields.Field(L"convert_source", &settings->source, ConvertSource::kPowerShell);
    fields.Field(L"convert_input", &settings->input_path);
    fields.Field(L"convert_key", &settings->key_path);
    fields.Field(L"convert_recursive", &settings->recursive);
    fields.Field(L"convert_format", &settings->format, regfile::Format::kPowerShell);
    fields.Field(L"convert_output", &settings->output_path);
    fields.Field(L"convert_admin_check", &settings->admin_check);
    fields.Field(L"convert_open_in_editor", &settings->open_in_editor);
}

}
