// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/query_prompts.h"

#include <win32ui/dialog/check_list.h>
#include <win32ui/dialog/dialog_layout.h>
#include <win32ui/dialog/dialog_metrics.h>
#include <win32ui/dialog/dialog_window.h>

#include <algorithm>
#include <vector>

#include <commctrl.h>
#include <windowsx.h>

#include "browse/key_tree.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/value_format.h"
#include "resource.h"
#include "ui/feedback.h"
#include "win32/translation.h"

namespace regkit::query_prompts
{

struct DataTypeItem
{
    DWORD type = 0;
    std::wstring label;
};

constexpr DWORD kExtendedTypeFlags[] = {0x20000, 0x40000};

std::vector<DataTypeItem> BuildDataTypeItems()
{
    std::vector<DataTypeItem> items;
    const auto base_types = value_format::TypeLabels();
    items.reserve(base_types.size() * (1 + _countof(kExtendedTypeFlags)));
    for (const auto& entry : base_types)
    {
        items.push_back({entry.type, entry.name});
    }
    for (DWORD flag : kExtendedTypeFlags)
    {
        for (const auto& entry : base_types)
        {
            DWORD type = flag | entry.type;
            items.push_back({type, value_format::TypeName(type)});
        }
    }
    return items;
}

bool RequireOneType(void*, HWND dialog, const std::vector<size_t>& checked)
{
    if (checked.empty())
    {
        ui::ShowWarning(dialog, util::Tr(L"Select at least one data type."));
        return false;
    }
    return true;
}

bool ShowDataTypes(HWND owner, std::vector<DWORD>* types)
{
    if (!types)
    {
        return false;
    }
    const std::vector<DataTypeItem> items = BuildDataTypeItems();
    std::vector<std::wstring> labels;
    std::vector<size_t> checked;
    for (size_t index = 0; index < items.size(); ++index)
    {
        labels.push_back(items[index].label);
        // no saved types means all
        if (types->empty() || std::find(types->begin(), types->end(), items[index].type) != types->end())
        {
            checked.push_back(index);
        }
    }
    if (win32ui::PickChecks(owner, {.title = util::Tr(L"Data Types"), .labels = labels, .accept = RequireOneType}, &checked) != 0)
    {
        return false;
    }
    types->clear();
    for (const size_t index : checked)
    {
        types->push_back(items[index].type);
    }
    return true;
}

// registry tree is created at runtime
class BrowseKeyDialog final : public win32ui::DialogWindow
{
  public:
    const std::wstring& selected_path() const noexcept
    {
        return selected_path_;
    }

  protected:
    void OnCreate() override
    {
        // keep outer size at 420 by 420 DIPs regardless of frame metrics, opening size is the resize minimum
        SetWindowPos(hwnd(), nullptr, 0, 0, Px(420), Px(420), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        ok_button_ = GetDlgItem(hwnd(), IDOK);
        cancel_button_ = GetDlgItem(hwnd(), IDCANCEL);
        tree_.Create(hwnd(), GetModuleHandleW(nullptr), IDC_BROWSE_KEY_TREE);
        tree_.PopulateRoots(RegistryStore::DefaultRoots());
        // tree precedes template buttons in tab order
        SetWindowLongPtrW(tree_.hwnd(), GWL_STYLE, GetWindowLongPtrW(tree_.hwnd(), GWL_STYLE) | WS_TABSTOP);
        SetWindowPos(tree_.hwnd(), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        ShowWindow(tree_.hwnd(), SW_SHOW);
    }

    HWND initial_focus() const override
    {
        return tree_.hwnd();
    }

    void OnLayout(int width, int height) override
    {
        const int margin = Px(win32ui::metrics::kDialogContentMargin);
        const int button_h = TextPx(win32ui::metrics::kButtonHeight);
        const int right_margin = Px(win32ui::metrics::kDialogButtonRightMargin);
        const int bottom_y = height - Px(win32ui::metrics::kDialogButtonBottomMargin) - button_h;
        win32ui::Place(tree_.hwnd(), margin, margin, width - margin * 2, std::max(0, bottom_y - Px(win32ui::metrics::kBlockGap) - margin));
        const int buttons_w = win32ui::PlaceButtonRow({ok_button_, cancel_button_}, width - right_margin, bottom_y, Px(win32ui::metrics::kButtonMinWidth), button_h, Px(win32ui::metrics::kButtonGap));
        win32ui::GrowDialogWidth(hwnd(), margin + buttons_w + right_margin);
    }

    LRESULT OnNotify(NMHDR* header, bool* handled) override
    {
        *handled = false;
        if (header->hwndFrom != tree_.hwnd())
        {
            return 0;
        }
        switch (header->code)
        {
        case TVN_ITEMEXPANDINGW:
            *handled = true;
            tree_.OnItemExpanding(reinterpret_cast<NMTREEVIEWW*>(header));
            return 0;
        case TVN_GETDISPINFOW:
            *handled = true;
            tree_.OnGetDispInfo(reinterpret_cast<NMTVDISPINFOW*>(header));
            return 0;
        case TVN_SELCHANGEDW:
            *handled = true;
            {
                // a group node has no key, OK then asks for one
                const RegistryNode* node = tree_.OnSelectionChanged(reinterpret_cast<NMTREEVIEWW*>(header));
                selected_path_ = node ? registry_path::Build(*node) : std::wstring();
            }
            return 0;
        default:
            return 0;
        }
    }

    bool OnCommand(int id, int, HWND) override
    {
        if (id != IDOK)
        {
            return false;
        }
        if (selected_path_.empty())
        {
            ui::ShowWarning(hwnd(), util::Tr(L"Select a key."));
            return true;
        }
        End(IDOK);
        return true;
    }

  private:
    HWND ok_button_ = nullptr;
    HWND cancel_button_ = nullptr;
    RegistryTree tree_;
    std::wstring selected_path_;
};

bool ShowRegistryKey(HWND owner, std::wstring* selected_path)
{
    if (!selected_path)
    {
        return false;
    }
    BrowseKeyDialog dialog;
    if (dialog.RunModal(owner, MAKEINTRESOURCEW(IDD_BROWSE_KEY)) != IDOK)
    {
        return false;
    }
    *selected_path = dialog.selected_path();
    return true;
}

}
