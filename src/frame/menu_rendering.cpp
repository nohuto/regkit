// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"
#include "frame/window_impl.h"

namespace regkit
{
using namespace command_detail;

void MainWindow::Impl::PrepareMenusForOwnerDraw(HMENU menu)
{
    if (!menu)
    {
        return;
    }
    MENUINFO menu_info = {};
    menu_info.cbSize = sizeof(menu_info);
    menu_info.fMask = MIM_BACKGROUND;
    menu_info.hbrBack = Theme::Current().BackgroundBrush();
    SetMenuInfo(menu, &menu_info);

    HDC hdc = GetDC(hwnd_);
    HFONT old_font = nullptr;
    if (hdc && ui_font_)
    {
        old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
    }

    int count = GetMenuItemCount(menu);
    for (int i = 0; i < count; ++i)
    {
        MENUITEMINFOW info = {};
        info.cbSize = sizeof(info);
        info.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_SUBMENU | MIIM_ID;
        wchar_t text[256] = {};
        info.dwTypeData = text;
        info.cch = static_cast<UINT>(_countof(text));
        if (!GetMenuItemInfoW(menu, i, TRUE, &info))
        {
            continue;
        }

        auto data = std::make_unique<MenuItemData>();
        data->text = text;
        data->separator = (info.fType & MFT_SEPARATOR) != 0;
        if (data->separator)
        {
            data->width = 4;
            data->height = 8;
        }
        else
        {
            RECT measure = {};
            if (hdc)
            {
                DrawTextW(hdc, data->text.c_str(), -1, &measure, DT_SINGLELINE | DT_CALCRECT);
            }
            data->height = 18;
            data->width = static_cast<int>(measure.right - measure.left) + 8;
        }

        MenuItemData* raw = data.get();
        menu_items_.push_back(std::move(data));

        info.fMask = MIIM_FTYPE | MIIM_DATA;
        info.fType |= MFT_OWNERDRAW;
        info.dwItemData = reinterpret_cast<ULONG_PTR>(raw);
        SetMenuItemInfoW(menu, i, TRUE, &info);
    }

    if (hdc && old_font)
    {
        SelectObject(hdc, old_font);
    }
    if (hdc)
    {
        ReleaseDC(hwnd_, hdc);
    }
}

void MainWindow::Impl::OnMeasureMenuItem(MEASUREITEMSTRUCT* info)
{
    if (!info)
    {
        return;
    }
    auto* data = reinterpret_cast<MenuItemData*>(info->itemData);
    if (!data)
    {
        return;
    }
    if (data->width > 0 && data->height > 0)
    {
        info->itemWidth = static_cast<UINT>(data->width);
        info->itemHeight = static_cast<UINT>(data->height);
        return;
    }
    if (data->separator)
    {
        info->itemHeight = 8;
        info->itemWidth = 4;
        return;
    }

    SIZE size = {};
    HDC hdc = GetDC(hwnd_);
    if (hdc)
    {
        HFONT old = nullptr;
        if (ui_font_)
        {
            old = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
        }
        GetTextExtentPoint32W(hdc, data->text.c_str(), static_cast<int>(data->text.size()), &size);
        if (old)
        {
            SelectObject(hdc, old);
        }
        ReleaseDC(hwnd_, hdc);
    }
    info->itemHeight = 18;
    info->itemWidth = size.cx + 8;
}

void MainWindow::Impl::OnDrawMenuItem(const DRAWITEMSTRUCT* info)
{
    if (!info)
    {
        return;
    }
    auto* data = reinterpret_cast<MenuItemData*>(info->itemData);
    if (!data)
    {
        return;
    }
    const Theme& theme = Theme::Current();
    HDC hdc = info->hDC;
    RECT rect = info->rcItem;

    if (data->separator)
    {
        HPEN pen = appearance::CachedPen(theme.BorderColor());
        HPEN old = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
        int y = (rect.top + rect.bottom) / 2;
        MoveToEx(hdc, rect.left + 8, y, nullptr);
        LineTo(hdc, rect.right - 8, y);
        SelectObject(hdc, old);
        return;
    }

    const bool selected = (info->itemState & (ODS_SELECTED | ODS_HOTLIGHT)) != 0;
    const bool disabled = (info->itemState & ODS_DISABLED) != 0;
    COLORREF fg = disabled ? theme.MutedTextColor() : theme.TextColor();
    HBRUSH bg_brush = selected ? appearance::CachedBrush(theme.HoverColor()) : theme.BackgroundBrush();
    FillRect(hdc, &rect, bg_brush);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, fg);
    HFONT old_font = nullptr;
    if (ui_font_)
    {
        old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
    }
    UINT format = DT_SINGLELINE | DT_VCENTER | DT_CENTER | DT_END_ELLIPSIS;
    if ((info->itemState & ODS_NOACCEL) != 0)
    {
        format |= DT_HIDEPREFIX;
    }
    DrawTextW(hdc, data->text.c_str(), -1, &rect, format);
    if (old_font)
    {
        SelectObject(hdc, old_font);
    }
}

} // namespace regkit
