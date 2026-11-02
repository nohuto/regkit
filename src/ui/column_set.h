// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <commctrl.h>

#include <string>
#include <vector>

namespace regkit::ui
{

struct ColumnInfo
{
    std::wstring title;
    int width = 0;
    int fmt = LVCFMT_LEFT;
};

struct ColumnSet
{
    std::vector<ColumnInfo> items;
    std::vector<int> widths;
    std::vector<bool> visible;

    void Reset(std::vector<ColumnInfo> columns)
    {
        items = std::move(columns);
        widths.clear();
        for (const ColumnInfo& column : items)
        {
            widths.push_back(column.width);
        }
        visible.assign(items.size(), true);
    }
};

}
