// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "browse/key_tree.h"
#include "browse/value_table.h"
#include "ui/column_set.h"
#include "win32/text_transform.h"

#include <windows.h>

#include <commctrl.h>

#include <algorithm>
#include <memory>
#include <optional>
#include <string>
#include <vector>

namespace regkit::browse
{

struct CreateRequest
{
    HWND parent = nullptr;
    HINSTANCE instance = nullptr;
    int address_id = 0;
    int go_id = 0;
    int filter_id = 0;
    int tree_id = 0;
    int values_id = 0;
    SUBCLASSPROC address_proc = nullptr;
    UINT_PTR address_subclass_id = 0;
    SUBCLASSPROC filter_proc = nullptr;
    UINT_PTR filter_subclass_id = 0;
    SUBCLASSPROC tree_proc = nullptr;
    UINT_PTR tree_subclass_id = 0;
    DWORD_PTR callback_context = 0;
};

struct ColumnState : ui::ColumnSet
{
    std::vector<int> saved_widths;
    std::vector<bool> saved_visible;
    bool saved = false;
    int sort_column = 0;
    bool sort_ascending = true;
    util::SortKeys sort_earlier;

    bool SortsBy(int column) const
    {
        return sort_column == column || std::any_of(sort_earlier.begin(), sort_earlier.end(), [column](const auto& key) { return key.first == column; });
    }
};

struct NavigationAvailability
{
    bool back = false;
    bool forward = false;
    bool up = false;
};

class Pane
{
  public:
    bool Create(const CreateRequest& request);

    HWND address() const noexcept;
    HWND go_button() const noexcept;
    HWND filter() const noexcept;
    RegistryTree& tree() noexcept;
    const RegistryTree& tree() const noexcept;
    RegistryTree* AddTree();
    void SetActiveTree(RegistryTree* tree) noexcept;
    void RemoveTree(RegistryTree* tree);
    const std::vector<std::unique_ptr<RegistryTree>>& trees() const noexcept;
    ValueList& values() noexcept;
    const ValueList& values() const noexcept;

    RegistryNode* current_node() const noexcept;
    void set_current_node(RegistryNode* node) noexcept;
    std::vector<RegistryRootEntry>& roots() noexcept;
    const std::vector<RegistryRootEntry>& roots() const noexcept;
    ColumnState& columns() noexcept;
    const ColumnState& columns() const noexcept;

    bool RecordNavigation(const std::wstring& path);
    std::optional<std::wstring> Back();
    std::optional<std::wstring> Forward();
    std::optional<std::wstring> Up();
    void UndoNavigation(int delta);
    NavigationAvailability navigation() const noexcept;
    void ResetNavigation();

    bool SelectValue(const std::wstring& name);

  private:
    HWND address_ = nullptr;
    HWND go_button_ = nullptr;
    HWND filter_ = nullptr;
    CreateRequest tree_request_;
    std::vector<std::unique_ptr<RegistryTree>> trees_;
    // placeholder until Create, layout runs before pane creation
    RegistryTree no_tree_;
    RegistryTree* tree_ = &no_tree_;
    ValueList values_;
    ColumnState columns_;
    std::vector<std::wstring> navigation_history_;
    int navigation_index_ = -1;
    bool programmatic_navigation_ = false;
};

}
