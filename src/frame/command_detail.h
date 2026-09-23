// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/window_detail.h"
#include "frame/window_draw_detail.h"
#include "frame/compare_dialog.h"
#include "frame/window_impl.h"

#include <algorithm>
#include <commctrl.h>
#include <cwctype>
#include <mutex>
#include <shellapi.h>
#include <unordered_map>
#include <unordered_set>

#include "appearance/default_font.h"
#include "appearance/dialog_layout.h"
#include "appearance/feedback.h"
#include "appearance/font_picker.h"
#include "appearance/gdi_cache.h"
#include "appearance/theme.h"
#include "editors/binary_editor.h"
#include "editors/hive_dialog.h"
#include "editors/value_editor.h"
#include "frame/command_dispatch.h"
#include "frame/command_ids.h"
#include "regfile/reg_file.h"
#include "regfile/registry_transfer.h"
#include "registry/registry_backends.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/value_format.h"
#include "resource.h"
#include "search/compare.h"
#include "search/search.h"
#include "win32/file_dialog.h"
#include "win32/process_rights.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/window_metrics.h"
#include "workspace/favorites.h"

namespace regkit::command_detail
{

using window_detail::ChildNode;
using window_detail::EqualsInsensitive;
using window_detail::FetchListViewItemText;
using window_detail::FileBaseName;
using window_detail::FileNameOnly;
using window_detail::FindChildByText;
using window_detail::kIconSetClassic;
using window_detail::kIconSetCustom;
using window_detail::kIconSetPhosphor;
using window_detail::LeafName;
using window_detail::MakeValueListRow;
using window_detail::ShortDefaultLabel;
using window_detail::StartsWithInsensitive;
using workspace::FavoritesStore;

constexpr wchar_t kHelpUrl[] = L"https://discord.noverse.dev";

using util::ToLower;
using util::TrimWhitespace;

inline HMENU BuildCopyKeyPathMenu()
{
    HMENU menu = CreatePopupMenu();
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathAbbrev, L"Abbreviated (HKLM)");
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathRegEdit, L"RegEdit Address Bar");
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathRegFile, L".reg File Header");
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathPowerShell, L"PowerShell Drive");
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathPowerShellProvider, L"PowerShell Provider");
    AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathEscaped, L"Escaped Backslashes");
    return menu;
}
constexpr wchar_t kOneKeyPerLineText[] = L"Each line should include one key.";

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

inline std::vector<ListRow> SelectedListRows(const ValueList& list)
{
    std::vector<ListRow> rows;
    if (!list.hwnd())
    {
        return rows;
    }
    rows.reserve(static_cast<size_t>(ListView_GetSelectedCount(list.hwnd())));
    int index = -1;
    while ((index = ListView_GetNextItem(list.hwnd(), index, LVNI_SELECTED)) >= 0)
    {
        const ListRow* row = list.RowAt(index);
        if (row)
        {
            rows.push_back(*row);
        }
    }
    return rows;
}

inline bool GetValueEntry(const RegistryNode& node, const std::wstring& name, ValueEntry* out)
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

inline bool SelectValueByName(ValueList& list, const std::wstring& name)
{
    for (size_t i = 0; i < list.RowCount(); ++i)
    {
        const ListRow* row = list.RowAt(static_cast<int>(i));
        if (!row || row->kind != rowkind::kValue)
        {
            continue;
        }
        if (row->extra == name)
        {
            ListView_SetItemState(list.hwnd(), static_cast<int>(i), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_EnsureVisible(list.hwnd(), static_cast<int>(i), FALSE);
            return true;
        }
    }
    return false;
}

inline bool GetListViewColumnInfo(HWND list, int display_index, int* subitem, int* width)
{
    if (subitem)
    {
        *subitem = -1;
    }
    if (width)
    {
        *width = 0;
    }
    if (!list || display_index < 0)
    {
        return false;
    }
    LVCOLUMNW col = {};
    col.mask = LVCF_SUBITEM | LVCF_WIDTH;
    if (!ListView_GetColumn(list, display_index, &col))
    {
        return false;
    }
    if (subitem)
    {
        *subitem = col.iSubItem;
    }
    if (width)
    {
        *width = col.cx;
    }
    return true;
}

inline std::wstring BuildSelectedListViewText(HWND list)
{
    if (!list)
    {
        return L"";
    }
    HWND header = ListView_GetHeader(list);
    int columns = header ? Header_GetItemCount(header) : 0;
    std::vector<int> subitems;
    subitems.reserve(columns);
    for (int i = 0; i < columns; ++i)
    {
        int subitem = -1;
        int width = 0;
        if (!GetListViewColumnInfo(list, i, &subitem, &width))
        {
            continue;
        }
        if (width <= 0 || subitem < 0)
        {
            continue;
        }
        subitems.push_back(subitem);
    }
    if (subitems.empty())
    {
        return L"";
    }

    std::wstring output;
    std::wstring buffer(256, L'\0');
    int index = -1;
    bool first_row = true;
    while ((index = ListView_GetNextItem(list, index, LVNI_SELECTED)) >= 0)
    {
        if (!first_row)
        {
            output.append(L"\r\n");
        }
        first_row = false;
        for (size_t i = 0; i < subitems.size(); ++i)
        {
            if (i > 0)
            {
                output.append(L"\t");
            }
            buffer.assign(256, L'\0');
            int length = FetchListViewItemText(list, index, subitems[i], &buffer);
            if (length > 0)
            {
                output.append(buffer.c_str(), static_cast<size_t>(length));
            }
        }
    }
    return output;
}
} // namespace regkit::command_detail
