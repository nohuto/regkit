// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <win32ui/list_view/list_view.h>

#include <utility>
#include <vector>

namespace regkit::ui
{

struct ColumnSet
{
    std::vector<win32ui::ListColumn> items;

    void Reset(std::vector<win32ui::ListColumn> columns)
    {
        for (win32ui::ListColumn& column : columns)
        {
            column.default_width_dip = column.width_dip;
        }
        items = std::move(columns);
    }
};

}
