// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "ui/dialog_layout.h"

#include "ui/theme.h"

namespace regkit::appearance
{

const win32ui::DialogTheme* DialogThemeHooks()
{
    // template edits are ES_MULTILINE for their vertical centring, so they are kept to one line
    return win32ui::ThemeDialogHooks(win32ui::kPrepareTranslate | win32ui::kPrepareNoTextLimit | win32ui::kPrepareSingleLine);
}

win32ui::StatusBarHooks StatusBarThemeHooks()
{
    win32ui::StatusBarHooks hooks;
    hooks.frame = DialogThemeHooks();
    return hooks;
}

win32ui::TabStripHooks TabStripThemeHooks()
{
    win32ui::TabStripHooks hooks;
    hooks.frame = DialogThemeHooks();
    return hooks;
}

}
