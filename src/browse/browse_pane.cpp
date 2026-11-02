// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "browse/browse_pane.h"

#include "registry/registry_path.h"
#include "win32/text_transform.h"

#include <win32ui/list_view/list_view.h>

#include <algorithm>

namespace regkit::browse
{

bool Pane::Create(const CreateRequest& request)
{
    if (!request.parent || !request.instance)
    {
        return false;
    }
    address_ = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_MULTILINE, 0, 0, 0, 0, request.parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(request.address_id)), request.instance, nullptr);
    go_button_ =
        CreateWindowExW(0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, 0, 0, 0, 0, request.parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(request.go_id)), request.instance, nullptr);
    filter_ = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_MULTILINE, 0, 0, 0, 0, request.parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(request.filter_id)), request.instance, nullptr);
    tree_request_ = request;
    tree_ = AddTree();
    values_.Create(request.parent, request.instance, request.values_id);
    if (!address_ || !go_button_ || !filter_ || !tree_ || !values_.hwnd())
    {
        return false;
    }

    if (request.address_proc &&
        !SetWindowSubclass(address_, request.address_proc, request.address_subclass_id, request.callback_context))
    {
        return false;
    }
    if (request.filter_proc &&
        !SetWindowSubclass(filter_, request.filter_proc, request.filter_subclass_id, request.callback_context))
    {
        return false;
    }
    return true;
}

HWND Pane::address() const noexcept
{
    return address_;
}
HWND Pane::go_button() const noexcept
{
    return go_button_;
}
HWND Pane::filter() const noexcept
{
    return filter_;
}
RegistryTree& Pane::tree() noexcept
{
    return *tree_;
}
const RegistryTree& Pane::tree() const noexcept
{
    return *tree_;
}

// each tab keeps its tree, switching only swaps windows
RegistryTree* Pane::AddTree()
{
    auto tree = std::make_unique<RegistryTree>();
    tree->Create(tree_request_.parent, tree_request_.instance, tree_request_.tree_id, false, true);
    if (!tree->hwnd() || (tree_request_.tree_proc &&
                          !SetWindowSubclass(tree->hwnd(), tree_request_.tree_proc, tree_request_.tree_subclass_id, tree_request_.callback_context)))
    {
        if (tree->hwnd())
        {
            DestroyWindow(tree->hwnd());
        }
        return nullptr;
    }
    trees_.push_back(std::move(tree));
    return trees_.back().get();
}

void Pane::SetActiveTree(RegistryTree* tree) noexcept
{
    tree_ = tree;
}

void Pane::RemoveTree(RegistryTree* tree)
{
    auto it = std::find_if(trees_.begin(), trees_.end(), [tree](const auto& owned) { return owned.get() == tree; });
    if (it == trees_.end() || tree == tree_)
    {
        return;
    }
    DestroyWindow(tree->hwnd());
    trees_.erase(it);
}

const std::vector<std::unique_ptr<RegistryTree>>& Pane::trees() const noexcept
{
    return trees_;
}
ValueList& Pane::values() noexcept
{
    return values_;
}
const ValueList& Pane::values() const noexcept
{
    return values_;
}
RegistryNode* Pane::current_node() const noexcept
{
    return tree_->current_node();
}
void Pane::set_current_node(RegistryNode* node) noexcept
{
    tree_->set_current_node(node);
}
std::vector<RegistryRootEntry>& Pane::roots() noexcept
{
    return tree_->roots();
}
const std::vector<RegistryRootEntry>& Pane::roots() const noexcept
{
    return tree_->roots();
}
ColumnState& Pane::columns() noexcept
{
    return columns_;
}
const ColumnState& Pane::columns() const noexcept
{
    return columns_;
}

bool Pane::RecordNavigation(const std::wstring& path)
{
    if (path.empty())
    {
        return false;
    }
    if (programmatic_navigation_)
    {
        programmatic_navigation_ = false;
        return false;
    }
    if (navigation_index_ >= 0 && navigation_index_ < static_cast<int>(navigation_history_.size()) &&
        navigation_history_[static_cast<size_t>(navigation_index_)] == path)
    {
        return false;
    }
    if (navigation_index_ + 1 < static_cast<int>(navigation_history_.size()))
    {
        navigation_history_.erase(navigation_history_.begin() + navigation_index_ + 1, navigation_history_.end());
    }
    navigation_history_.push_back(path);
    navigation_index_ = static_cast<int>(navigation_history_.size()) - 1;
    return true;
}

std::optional<std::wstring> Pane::Back()
{
    if (navigation_index_ <= 0)
    {
        return std::nullopt;
    }
    --navigation_index_;
    programmatic_navigation_ = true;
    return navigation_history_[static_cast<size_t>(navigation_index_)];
}

std::optional<std::wstring> Pane::Forward()
{
    if (navigation_index_ + 1 >= static_cast<int>(navigation_history_.size()))
    {
        return std::nullopt;
    }
    ++navigation_index_;
    programmatic_navigation_ = true;
    return navigation_history_[static_cast<size_t>(navigation_index_)];
}

std::optional<std::wstring> Pane::Up()
{
    const RegistryNode* node = current_node();
    if (!node || node->subkey.empty())
    {
        return std::nullopt;
    }
    programmatic_navigation_ = true;
    return registry_path::Parent(registry_path::Build(*node));
}

void Pane::UndoNavigation(int delta)
{
    navigation_index_ += delta;
    if (navigation_index_ < -1)
    {
        navigation_index_ = -1;
    }
    programmatic_navigation_ = false;
}

NavigationAvailability Pane::navigation() const noexcept
{
    NavigationAvailability available;
    available.back = navigation_index_ > 0;
    available.forward = navigation_index_ + 1 < static_cast<int>(navigation_history_.size());
    available.up = current_node() && !current_node()->subkey.empty();
    return available;
}

void Pane::ResetNavigation()
{
    navigation_history_.clear();
    navigation_index_ = -1;
    programmatic_navigation_ = false;
}

bool Pane::SelectValue(const std::wstring& name)
{
    if (!values_.hwnd())
    {
        return false;
    }
    for (size_t index = 0; index < values_.RowCount(); ++index)
    {
        const ListRow* row = values_.RowAt(static_cast<int>(index));
        if (!row || row->kind != rowkind::kValue || !util::EqualsInsensitive(row->extra, name))
        {
            continue;
        }
        return win32ui::SelectOnlyListRow(values_.hwnd(), static_cast<int>(index));
    }
    return false;
}

}
