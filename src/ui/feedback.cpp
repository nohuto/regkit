// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "ui/feedback.h"

#include "win32/translation.h"

#include <win32ui/dialog/file_dialog.h>
#include <win32ui/dialog/message_dialog.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/process_launch.h>
#include <win32ui/system/shell_open.h>

#include <algorithm>
#include <cwchar>
#include <vector>

#include "win32/file_dialog.h"
#include "win32/process_rights.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"

namespace regkit::ui
{

namespace
{
constexpr wchar_t kAppTitle[] = L"RegKit";

// up to three buttons in IDYES, IDNO, IDCANCEL order, omit empty labels, first is default, width is outer width in DIPs
int ShowChoiceDialog(HWND owner, const std::wstring& title, const std::wstring& message, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, PCWSTR icon_id, int width, ChoiceButtonWidths button_widths = {}, const std::wstring& detail = std::wstring())
{
    win32ui::MessageSpec spec;
    spec.title = title.empty() ? kAppTitle : title;
    spec.message = message;
    spec.details = detail;
    spec.icon = icon_id;
    spec.copy_text = util::Tr(L"Copy");
    const std::pair<const std::wstring*, int> labels[] = {
        {&yes_label, IDYES},
        {&no_label, IDNO},
        {&cancel_label, IDCANCEL}
    };
    const int min_widths[] = {button_widths.yes, button_widths.no, button_widths.cancel};
    for (size_t index = 0; index < std::size(labels); ++index)
    {
        if (!labels[index].first->empty())
        {
            spec.buttons.push_back({labels[index].second, *labels[index].first, min_widths[index]});
        }
    }
    spec.default_id = spec.buttons.empty() ? IDCANCEL : spec.buttons.front().id;
    spec.outer_width_dip = width;
    return static_cast<int>(win32ui::ShowMessage(owner, spec));
}

// first line is message, remaining lines are details
void ShowErrorDialog(HWND owner, const std::wstring& title, const std::wstring& message)
{
    std::wstring text;
    std::wstring detail;
    win32ui::SplitMessageText(message, &text, &detail);
    ShowChoiceDialog(owner, title, text, util::Tr(L"OK"), L"", L"", nullptr, detail.empty() ? 320 : 520, {}, detail);
}

}

void ShowError(HWND owner, const std::wstring& message)
{
    ShowErrorDialog(owner, util::Tr(L"Error"), message);
}

void ShowStartError(HWND owner)
{
    ShowError(owner, util::TrDetail(L"The background task couldn't be started.", win32ui::FormatWin32Error(ERROR_NOT_ENOUGH_MEMORY)));
}

void ShowWarning(HWND owner, const std::wstring& message)
{
    ShowErrorDialog(owner, util::Tr(L"Warning"), message);
}

void ShowInfo(HWND owner, const std::wstring& message)
{
    ShowErrorDialog(owner, util::Tr(L"Information"), message);
}

void RevealFile(HWND owner, const std::wstring& path)
{
    // shell under another account fails only after a long wait
    if (!util::ShellUserMatchesProcess() || win32ui::RevealInExplorer(path.c_str()) != ERROR_SUCCESS)
    {
        ShowError(owner, util::TrDetail(L"The file couldn't be shown in Explorer.", path));
    }
}

void ShowAbout(HWND owner)
{
    win32ui::AboutSpec spec;
    spec.title = util::Tr(L"About RegKit");
    spec.rows = {
        L"\u00A9 nohuto 2026",
        util::TrLabel(L"Repository", L"<a href=\"https://github.com/nohuto/regkit\">https://github.com/nohuto/regkit</a>"),
        L"Discord: <a href=\"https://discord.noverse.dev\">https://discord.noverse.dev</a>",
        util::TrLabel(L"Website", L"<a href=\"https://www.noverse.dev/\">https://www.noverse.dev/</a>"),
        util::TrLabel(L"Email", L"<a href=\"mailto:contact@noverse.dev\">contact@noverse.dev</a>")
    };
    spec.ok_text = util::Tr(L"OK");
    win32ui::ShowAbout(owner, spec);
}

bool ConfirmRegFileMerge(HWND owner, const std::wstring& path)
{
    const std::wstring message = util::Tr(L"Adding information can unintentionally change or delete values and\n"
                                          L"cause components to stop working correctly. If you don't trust the\n"
                                          L"source of this information, don't add it to the registry.\n\n"
                                          L"Are you sure you want to continue?");
    return ShowChoiceDialog(owner, kAppTitle, message, util::Tr(L"Yes"), util::Tr(L"No"), L"", IDI_WARNING, 560, {}, path) == IDYES;
}

void ShowRegFileMergeSucceeded(HWND owner, const std::wstring& path)
{
    const std::wstring message = util::Tr(L"The keys and values it contains have been added to the registry.");
    ShowChoiceDialog(owner, kAppTitle, message, util::Tr(L"OK"), L"", L"", IDI_INFORMATION, 520, {}, path);
}

void ShowRegFileMergeFailed(HWND owner, const std::wstring& path, const std::wstring& detail)
{
    std::wstring message = util::Tr(L"The registry file couldn't be imported.");
    if (!detail.empty())
    {
        message += L"\n\n";
        message += detail;
    }
    ShowChoiceDialog(owner, kAppTitle, message, util::Tr(L"OK"), L"", L"", IDI_ERROR, 520, {}, path);
}

bool ConfirmDelete(HWND owner, const std::wstring& title, const std::vector<std::wstring>& names, const std::wstring& override_message)
{
    if (names.empty())
    {
        return false;
    }
    const bool many = names.size() > 1;
    std::wstring message = override_message;
    if (message.empty())
    {
        if (title == util::Tr(L"Delete Key"))
        {
            message = many ? util::Tr(L"Delete these keys and all of their subkeys?") : util::Tr(L"Delete this key and all of its subkeys?");
        }
        else if (title == util::Tr(L"Delete Value") || title == util::Tr(L"Delete Values"))
        {
            message = many ? util::Tr(L"Delete these values?") : util::Tr(L"Delete this value?");
        }
        else
        {
            message = many ? util::Tr(L"Delete these items?") : util::Tr(L"Delete this item?");
        }
    }

    std::wstring detail;
    for (const std::wstring& name : names)
    {
        detail.append(detail.empty() ? L"" : L"\r\n").append(name.empty() ? util::Tr(L"(Default)") : name);
    }

    return ShowChoiceDialog(owner, title, message, util::Tr(L"Delete"), L"", util::Tr(L"Cancel"), nullptr, 460, {}, detail) == IDYES;
}

bool ConfirmOverwrite(HWND owner, const std::wstring& path, const std::wstring& confirmed_path)
{
    return util::IsMissing(path.c_str()) || util::EqualsInsensitive(path, confirmed_path) ||
           PromptKeyChoice(owner, util::Tr(L"The file already exists. Replace it?"), path, kAppTitle, util::Tr(L"Replace"), L"", util::Tr(L"Cancel")) == IDYES;
}

bool ConfirmConversionSkips(HWND owner, const std::vector<std::wstring>& skipped, const wchar_t* action)
{
    std::wstring detail;
    for (const std::wstring& entry : skipped)
    {
        detail.append(detail.empty() ? L"" : L"\r\n").append(entry);
    }
    const std::wstring message = util::Tr(L"These entries can't be written in this format and will be left out.");
    return ShowChoiceDialog(owner, kAppTitle, message, action ? action : util::Tr(L"Convert"), L"", util::Tr(L"Cancel"), IDI_WARNING, 800, {}, detail) == IDYES;
}

void ShowConversionSucceeded(HWND owner, const std::wstring& path)
{
    ShowChoiceDialog(owner, kAppTitle, util::Tr(L"The file was converted."), util::Tr(L"OK"), L"", L"", IDI_INFORMATION, 520, {}, path);
}

bool ConfirmDelete(HWND owner, const std::wstring& title, const std::wstring& name, const std::wstring& override_message)
{
    return ConfirmDelete(owner, title, std::vector<std::wstring>{name}, override_message);
}

int PromptKeyChoice(HWND owner, const std::wstring& message, const std::wstring& key_path, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, ChoiceButtonWidths widths)
{
    return ShowChoiceDialog(owner, title, message, yes_label, no_label, cancel_label, nullptr, 560, widths, key_path);
}

int PromptChoice(HWND owner, const std::wstring& message, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, ChoiceButtonWidths widths, int width)
{
    return ShowChoiceDialog(owner, title, message, yes_label, no_label, cancel_label, nullptr, width, widths);
}

bool ConfirmPrivilegedRestart(HWND owner, const std::wstring& title)
{
    // relaunch executes whichever file occupies this path at that time
    const std::wstring exe_path = util::ModulePath();
    return !util::IsWritableByNonAdmins(exe_path) ||
           PromptKeyChoice(
               owner,
               util::Tr(L"RegKit is running from a location that non administrators can write to, so a program "
                        L"without administrator rights could replace it before it restarts with these rights. "
                        L"Install RegKit for all users first, or move it somewhere only administrators can write.\n\n"
                        L"Restart anyway to continue from this location."),
               exe_path,
               title,
               util::Tr(L"Restart Anyway"),
               L"",
               util::Tr(L"Cancel"),
               {110, 70, 70}
           ) == IDYES;
}

bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path)
{
    return ReportFileDialogResult(owner, win32ui::ChooseFileToOpen(owner, {.filter = filter}, path));
}

bool PromptSaveFile(HWND owner, const wchar_t* filter, std::wstring* path, win32::OpenAfter* open_after, bool regkit)
{
    return ReportFileDialogResult(owner, win32::ChooseFileToSave(owner, filter, nullptr, path, open_after, regkit));
}

bool ReportFileDialogResult(HWND owner, DWORD error)
{
    if (error != ERROR_SUCCESS && error != ERROR_CANCELLED)
    {
        ShowError(owner, win32::FormatDialogError(error));
    }
    return error == ERROR_SUCCESS;
}

DWORD LaunchNewInstance(HWND owner, const std::wstring& arguments)
{
    return win32ui::LaunchProcess(owner, util::ModulePath(), arguments);
}
}
