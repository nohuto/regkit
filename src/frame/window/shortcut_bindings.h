// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include "frame/commands/command_ids.h"

#include <win32ui/menu/popup_menu.h>
#include <win32ui/menu/shortcuts.h>

namespace regkit::frame
{

inline constexpr win32ui::Shortcut kShortcutBindings[] = {
    {cmd::kEditCopy, FCONTROL, 'C', L"Ctrl+C"},
    {cmd::kEditPaste, FCONTROL, 'V', L"Ctrl+V"},
    {cmd::kViewSelectAll, FCONTROL, 'A', L"Ctrl+A"},
    {cmd::kEditUndo, FCONTROL, 'Z', L"Ctrl+Z"},
    {cmd::kEditRedo, FCONTROL, 'Y', L"Ctrl+Y"},
    {cmd::kEditFind, FCONTROL, 'F', L"Ctrl+F"},
    {cmd::kEditGoTo, FCONTROL, 'G', L"Ctrl+G"},
    {cmd::kEditReplace, FCONTROL, 'H', L"Ctrl+H"},
    {cmd::kFileSave, FCONTROL, 'S', L"Ctrl+S"},
    {cmd::kFileExport, FCONTROL, 'E', L"Ctrl+E"},
    {cmd::kRegistryLocal, FCONTROL, 'N', L"Ctrl+N"},
    {cmd::kRegistryNetwork, FCONTROL, 'R', L"Ctrl+R"},
    {cmd::kRegistryOffline, FCONTROL, 'O', L"Ctrl+O"},
    {cmd::kViewFocusFilter, FCONTROL, 'K', L"Ctrl+K"},
    {cmd::kTabClose, FCONTROL, 'W', L"Ctrl+W"},
    {cmd::kTabNext, FCONTROL, VK_TAB, L"Ctrl+Tab"},
    {cmd::kTabPrevious, FCONTROL | FSHIFT, VK_TAB, L"Ctrl+Shift+Tab"},
    {cmd::kEditCopyKey, FCONTROL | FSHIFT, 'C', L"Ctrl+Shift+C"},
    {cmd::kWindowNew, FCONTROL | FSHIFT, 'N', L"Ctrl+Shift+N"},
    {cmd::kFileOpenRegFile, FCONTROL | FSHIFT, 'O', L"Ctrl+Shift+O"},
    {cmd::kViewHistory, FCONTROL | FSHIFT, 'H', L"Ctrl+Shift+H"},
    {cmd::kEditDelete, 0, VK_DELETE, L"Del"},
    {cmd::kHelpContents, 0, VK_F1, L"F1"},
    {cmd::kEditRename, 0, VK_F2, L"F2"},
    {cmd::kViewRefresh, 0, VK_F5, L"F5"},
    {cmd::kNewKey, 0, VK_F7, L"F7"},
    {cmd::kEditKeyInfo, FALT, VK_RETURN, L"Alt+Enter"},
    {cmd::kNavBack, FALT, VK_LEFT, L"Alt+Left"},
    {cmd::kNavForward, FALT, VK_RIGHT, L"Alt+Right"},
    {cmd::kNavUp, FALT, VK_UP, L"Alt+Up"},
    // Ctrl+1 through Ctrl+9 select tabs without menu labels
    {cmd::kTabSelectBase + 0, FCONTROL, '1', nullptr},
    {cmd::kTabSelectBase + 1, FCONTROL, '2', nullptr},
    {cmd::kTabSelectBase + 2, FCONTROL, '3', nullptr},
    {cmd::kTabSelectBase + 3, FCONTROL, '4', nullptr},
    {cmd::kTabSelectBase + 4, FCONTROL, '5', nullptr},
    {cmd::kTabSelectBase + 5, FCONTROL, '6', nullptr},
    {cmd::kTabSelectBase + 6, FCONTROL, '7', nullptr},
    {cmd::kTabSelectBase + 7, FCONTROL, '8', nullptr},
    {cmd::kTabSelectBase + 8, FCONTROL, '9', nullptr},
};

inline const wchar_t* ShortcutText(int command)
{
    return win32ui::ShortcutText(kShortcutBindings, command);
}

inline void AppendCommand(HMENU menu, UINT flags, int command, const wchar_t* text)
{
    win32ui::AppendMenuItem(menu, flags, static_cast<UINT_PTR>(command), text, ShortcutText(command));
}

inline void AppendRadioCommand(HMENU menu, int command, const wchar_t* text)
{
    win32ui::AppendRadioItem(menu, static_cast<UINT_PTR>(command), text, ShortcutText(command));
}

}
