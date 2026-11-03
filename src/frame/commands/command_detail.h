// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/window_detail.h"
#include "frame/detail/draw_detail.h"
#include "frame/search/compare_dialog.h"
#include "frame/window_impl.h"

#include <win32ui/gdi/fonts.h>
#include <win32ui/gdi/gdi_cache.h>
#include <win32ui/window_metrics.h>

#include <algorithm>
#include <commctrl.h>
#include <cwctype>
#include <mutex>
#include <shellapi.h>
#include <unordered_map>
#include <unordered_set>

#include "dialogs/binary_editor.h"
#include "dialogs/hive_dialog.h"
#include "dialogs/transfer_dialogs.h"
#include "dialogs/value_editor.h"
#include "frame/commands/command_ids.h"
#include "frame/window/shortcut_bindings.h"
#include "regfile/reg_file.h"
#include "regfile/registry_transfer.h"
#include "registry/hive_files.h"
#include "registry/registry_backends.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/value_format.h"
#include "resource.h"
#include "search/compare.h"
#include "search/search.h"
#include "ui/dialog_layout.h"
#include "ui/feedback.h"
#include "ui/theme.h"
#include "win32/file_dialog.h"
#include "win32/process_rights.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"
#include "workspace/favorites.h"

namespace regkit::command_detail
{

using window_detail::ChildNode;
using window_detail::EqualsInsensitive;
using window_detail::FindChildByText;
using window_detail::kIconSetClassic;
using window_detail::kIconSetCustom;
using window_detail::kIconSetPhosphor;
using window_detail::LeafName;
using window_detail::MakeValueListRow;
using window_detail::StartsWithInsensitive;
using workspace::FavoritesStore;

constexpr wchar_t kHelpUrl[] = L"https://noverse.dev/docs/regkit/overview/";
constexpr wchar_t kDiscordUrl[] = L"https://discord.noverse.dev";

struct PaneToggle
{
    int command;
    bool workspace::Settings::* shown;
};

constexpr PaneToggle kPaneToggles[] = {
    {cmd::kViewToolbar, &workspace::Settings::show_toolbar},
    {cmd::kViewAddressBar, &workspace::Settings::show_address_bar},
    {cmd::kViewFilterBar, &workspace::Settings::show_filter_bar},
    {cmd::kViewTabControl, &workspace::Settings::show_tab_control},
    {cmd::kViewKeyTree, &workspace::Settings::show_tree},
    {cmd::kViewHistory, &workspace::Settings::show_history},
    {cmd::kViewStatusBar, &workspace::Settings::show_status_bar},
};

using util::ToLower;
using util::TrimWhitespace;
using win32ui::MenuName;

inline HMENU BuildCopyKeyPathMenu()
{
    HMENU menu = CreatePopupMenu();
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathAbbrev, util::Tr(L"&Abbreviated (HKLM)"));
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathRegEdit, util::Tr(L"RegEdit Address &Bar"));
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathRegFile, util::Tr(L".reg &File Header"));
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathPowerShell, util::Tr(L"PowerShell &Drive"));
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathPowerShellProvider, util::Tr(L"&PowerShell Provider"));
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathEscaped, util::Tr(L"&Escaped Backslashes"));
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathNative, util::Tr(L"&Native"));
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathNativeResolved, util::Tr(L"Native (&Resolved)"));
    return menu;
}
inline HMENU BuildNewMenu()
{
    HMENU menu = CreatePopupMenu();
    frame::AppendCommand(menu, MF_STRING, cmd::kNewKey, util::Tr(L"&Key"));
    AppendMenuW(menu, MF_STRING, cmd::kNewVolatileKey, util::Tr(L"&Volatile Key"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, cmd::kNewString, util::Tr(L"&String Value"));
    AppendMenuW(menu, MF_STRING, cmd::kNewBinary, util::Tr(L"&Binary Value"));
    AppendMenuW(menu, MF_STRING, cmd::kNewDword, util::Tr(L"&DWORD (32-bit) Value"));
    AppendMenuW(menu, MF_STRING, cmd::kNewQword, util::Tr(L"&QWORD (64-bit) Value"));
    AppendMenuW(menu, MF_STRING, cmd::kNewMultiString, util::Tr(L"&Multi-String Value"));
    AppendMenuW(menu, MF_STRING, cmd::kNewExpandString, util::Tr(L"&Expandable String Value"));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, cmd::kNewSymbolicLink, util::Tr(L"Symbolic &Link"));
    return menu;
}

using util::JoinLines;
using util::SplitLines;

inline const ListRow* SelectedValueRow(const ValueList& list, int* out_index)
{
    if (!list.hwnd())
    {
        return nullptr;
    }
    int index = ListView_GetNextItem(list.hwnd(), -1, LVNI_SELECTED);
    if (index < 0)
    {
        return nullptr;
    }
    if (out_index)
    {
        *out_index = index;
    }
    return list.RowAt(index);
}

// valid until the list rows change, a modal loop can reload them
inline std::vector<const ListRow*> SelectedListRows(const ValueList& list)
{
    const std::vector<int> indices = win32ui::SelectedListItems(list.hwnd());
    std::vector<const ListRow*> rows;
    rows.reserve(indices.size());
    for (const int index : indices)
    {
        if (const ListRow* row = list.RowAt(index))
        {
            rows.push_back(row);
        }
    }
    return rows;
}

inline bool GetValueEntry(const RegistryNode& node, const std::wstring& name, RegistryValue* out)
{
    if (RegistryStore::QueryValue(node, name, out))
    {
        return true;
    }
    if (out && name.empty())
    {
        out->name.clear();
        out->type = REG_SZ;
        out->data.clear();
        return true;
    }
    return false;
}

}
