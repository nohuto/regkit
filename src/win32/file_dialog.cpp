// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/file_dialog.h"

#include "win32/translation.h"

#include <win32ui/dialog/file_dialog.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/strings.h>

#include <cwchar>
#include <span>

namespace regkit::win32
{

namespace
{

OpenAfter g_open_after = OpenAfter::kNone;

}

DWORD ChooseFileToSave(HWND owner, const wchar_t* filter, const wchar_t* suggested_name, std::wstring* path, OpenAfter* open_after, bool regkit)
{
    static const win32ui::FileDialogChoice kChoices[] = {
        {static_cast<DWORD>(OpenAfter::kNone), util::TrNoop(L"Don't open")},
        {static_cast<DWORD>(OpenAfter::kEditor), util::TrNoop(L"In text editor")},
        {static_cast<DWORD>(OpenAfter::kRegKit), util::TrNoop(L"In RegKit")},
    };
    DWORD choice = static_cast<DWORD>(!regkit && g_open_after == OpenAfter::kRegKit ? OpenAfter::kNone : g_open_after);
    win32ui::FileDialogSpec spec;
    spec.filter = filter;
    spec.initial_name = suggested_name ? suggested_name : L"";
    if (open_after)
    {
        spec.choice_label = util::TrNoop(L"Open after saving:");
        spec.choices = std::span(kChoices).first(regkit ? 3 : 2);
        spec.choice = &choice;
    }
    const DWORD error = win32ui::ChooseFileToSave(owner, spec, path);
    if (error == ERROR_SUCCESS && open_after)
    {
        g_open_after = *open_after = static_cast<OpenAfter>(choice);
    }
    return error;
}

std::wstring FormatDialogError(DWORD error)
{
    // Win32FromHResult keeps the bits of HRESULTs outside FACILITY_WIN32
    if ((error & 0x80000000u) == 0)
    {
        return win32ui::FormatWin32Error(error);
    }
    return util::TrDetail(L"The file dialog failed.", win32ui::FormatWin32Error(error) + L" (" + win32ui::FormatHex(error, 8) + L")");
}

}
