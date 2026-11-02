// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "browse/key_tree.h"

#include <algorithm>

#include "registry/registry_path.h"
#include "win32/text_transform.h"

namespace regkit
{

namespace
{
constexpr int kFolderIconIndex = 0;
constexpr int kRootKeysIconIndex = 7;
constexpr int kRegistryIconIndex = 8;
constexpr wchar_t kRootKeysGroupLabel[] = L"Root Keys";
constexpr wchar_t kRealGroupLabel[] = L"REGISTRY";

using util::ToLower;

HTREEITEM InsertFolderItem(HWND tree, HTREEITEM parent, const wchar_t* label, int icon = kFolderIconIndex)
{
    TVINSERTSTRUCTW insert = {};
    insert.hParent = parent;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM;
    insert.item.pszText = const_cast<wchar_t*>(label);
    insert.item.iImage = icon;
    insert.item.iSelectedImage = icon;
    return TreeView_InsertItem(tree, &insert);
}

HTREEITEM InsertNodeItem(HWND tree, HTREEITEM parent, HTREEITEM after, const std::wstring& label, RegistryNode* node)
{
    TVINSERTSTRUCTW insert = {};
    insert.hParent = parent;
    insert.hInsertAfter = after;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_CHILDREN;
    insert.item.pszText = const_cast<wchar_t*>(label.c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(node);
    insert.item.iImage = I_IMAGECALLBACK;
    insert.item.iSelectedImage = I_IMAGECALLBACK;
    insert.item.cChildren = I_CHILDRENCALLBACK;
    return TreeView_InsertItem(tree, &insert);
}

void SetChildState(HWND tree, HTREEITEM item, RegistryNode* node, bool has_children)
{
    node->has_children = has_children ? 1 : 0;
    TVITEMW state = {};
    state.mask = TVIF_CHILDREN;
    state.hItem = item;
    state.cChildren = has_children ? 1 : 0;
    TreeView_SetItem(tree, &state);
}
}

void RegistryTree::Create(HWND parent, HINSTANCE instance, int control_id, bool show_border, bool allow_label_edit)
{
    // hide until positioned to avoid a white flash at 0,0
    DWORD style = WS_CHILD | WS_CLIPSIBLINGS | TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS;
    if (show_border)
    {
        style |= WS_BORDER;
    }
    if (allow_label_edit)
    {
        style |= TVS_EDITLABELS;
    }
    hwnd_ = CreateWindowExW(0, WC_TREEVIEWW, L"", style, 0, 0, 100, 100, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)), instance, nullptr);
    hooks_.load_children = [](void* context, HWND, HTREEITEM item) { static_cast<RegistryTree*>(context)->LoadChildren(item); };
    hooks_.is_group = [](void* context, HWND, HTREEITEM item) { return static_cast<RegistryTree*>(context)->IsGroupItem(item); };
    hooks_.expand_all = [](void* context, HWND, HTREEITEM) {
        const auto* self = static_cast<RegistryTree*>(context);
        if (!self->expand_all_window_)
        {
            return false;
        }
        SendMessageW(self->expand_all_window_, WM_COMMAND, self->expand_all_command_, 0);
        return true;
    };
    hooks_.context = this;
    win32ui::AttachTreeView(hwnd_, &hooks_);
}

HWND RegistryTree::hwnd() const
{
    return hwnd_;
}

void RegistryTree::SetImageList(HIMAGELIST image_list)
{
    if (!hwnd_)
    {
        return;
    }
    TreeView_SetImageList(hwnd_, image_list, TVSIL_NORMAL);
}

void RegistryTree::SetIconResolver(std::function<int(const RegistryNode&)> resolver)
{
    icon_resolver_ = std::move(resolver);
}

void RegistryTree::SetVirtualChildProvider(
    std::function<void(const RegistryNode&, const std::unordered_set<std::wstring>&, std::vector<std::wstring>*)>
        provider
)
{
    virtual_child_provider_ = std::move(provider);
}

void RegistryTree::SetRootLabel(const std::wstring& label, int icon)
{
    root_icon_ = icon;
    if (label.empty())
    {
        root_label_ = L"Computer";
        return;
    }
    root_label_ = label;
}

void RegistryTree::SetRegEditLayout(bool enabled)
{
    regedit_layout_ = enabled;
}

void RegistryTree::SuspendRedraw()
{
    if (redraw_suspended_++ == 0)
    {
        SendMessageW(hwnd_, WM_SETREDRAW, FALSE, 0);
    }
}

void RegistryTree::ResumeRedraw()
{
    if (redraw_suspended_ > 0 && --redraw_suspended_ == 0)
    {
        SendMessageW(hwnd_, WM_SETREDRAW, TRUE, 0);
        RedrawWindow(hwnd_, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
    }
}

void RegistryTree::Clear()
{
    TreeView_DeleteAllItems(hwnd_);
    nodes_.clear();
    roots_.clear();
    current_node_ = nullptr;
    root_item_ = nullptr;
    standard_group_item_ = nullptr;
    real_group_item_ = nullptr;
}

void RegistryTree::PopulateRoots(const std::vector<RegistryRootEntry>& roots)
{
    TreeView_DeleteAllItems(hwnd_);
    nodes_.clear();
    nodes_.reserve(roots.size() + 3);

    root_item_ = InsertFolderItem(hwnd_, TVI_ROOT, root_label_.c_str(), root_icon_);
    // RegEdit layout uses a flat tree instead of root groups
    const auto has_group = [&](bool real) {
        return !regedit_layout_ && std::any_of(roots.begin(), roots.end(), [&](const RegistryRootEntry& entry) {
            return (entry.group == RegistryRootGroup::kReal) == real;
        });
    };
    standard_group_item_ =
        has_group(false) ? InsertFolderItem(hwnd_, root_item_, kRootKeysGroupLabel, kRootKeysIconIndex) : nullptr;
    real_group_item_ =
        has_group(true) ? InsertFolderItem(hwnd_, root_item_, kRealGroupLabel, kRegistryIconIndex) : nullptr;

    for (const auto& root_entry : roots)
    {
        if (regedit_layout_ && root_entry.group == RegistryRootGroup::kReal)
        {
            continue;
        }

        auto node = std::make_unique<RegistryNode>();
        node->root = root_entry.root;
        node->view = root_entry.view;
        node->subkey = root_entry.subkey_prefix;
        node->root_name = root_entry.path_name;
        RegistryNode* stored = StoreNode(std::move(node));

        int icon_index = kFolderIconIndex;
        if (icon_resolver_)
        {
            icon_index = icon_resolver_(*stored);
        }

        if (!regedit_layout_ && root_entry.group == RegistryRootGroup::kReal && real_group_item_ &&
            util::EqualsInsensitive(root_entry.display_name, kRealGroupLabel) && root_entry.subkey_prefix.empty())
        {
            TVITEMW item = {};
            item.mask = TVIF_PARAM | TVIF_CHILDREN;
            item.hItem = real_group_item_;
            item.lParam = reinterpret_cast<LPARAM>(stored);
            item.cChildren = I_CHILDRENCALLBACK;
            TreeView_SetItem(hwnd_, &item);
            continue;
        }

        TVINSERTSTRUCTW root_item = {};
        if (regedit_layout_)
        {
            root_item.hParent = root_item_;
        }
        else if (root_entry.group == RegistryRootGroup::kReal)
        {
            root_item.hParent = real_group_item_ ? real_group_item_ : root_item_;
        }
        else
        {
            root_item.hParent = standard_group_item_ ? standard_group_item_ : root_item_;
        }
        root_item.hInsertAfter = TVI_LAST;
        root_item.item.mask = TVIF_TEXT | TVIF_PARAM | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_CHILDREN;
        root_item.item.pszText = const_cast<wchar_t*>(root_entry.display_name.c_str());
        root_item.item.lParam = reinterpret_cast<LPARAM>(stored);
        root_item.item.iImage = icon_index;
        root_item.item.iSelectedImage = icon_index;
        root_item.item.cChildren = I_CHILDRENCALLBACK;
        TreeView_InsertItem(hwnd_, &root_item);
    }

    TreeView_Expand(hwnd_, root_item_, TVE_EXPAND);
    if (standard_group_item_)
    {
        TreeView_Expand(hwnd_, standard_group_item_, TVE_EXPAND);
    }
    if (real_group_item_)
    {
        TreeView_Expand(hwnd_, real_group_item_, TVE_EXPAND);
    }
}

RegistryNode* RegistryTree::NodeFromItem(HTREEITEM item) const
{
    LPARAM data = 0;
    return win32ui::TreeItemData(hwnd_, item, &data) ? reinterpret_cast<RegistryNode*>(data) : nullptr;
}

void RegistryTree::LoadChildren(HTREEITEM item)
{
    RegistryNode* node = NodeFromItem(item);
    if (node && !node->children_loaded)
    {
        node->children_loaded = AddChildren(item, node);
    }
}

void RegistryTree::OnItemExpanding(const NMTREEVIEWW* info)
{
    if (info && info->action == TVE_EXPAND)
    {
        LoadChildren(info->itemNew.hItem);
    }
}

void RegistryTree::Resync(HTREEITEM item, bool load)
{
    if (RegistryNode* node = NodeFromItem(item))
    {
        node->icon = -1;
        if (!load && !node->children_loaded)
        {
            node->has_children = -1;
            if (TreeView_GetItemState(hwnd_, item, TVIS_EXPANDEDONCE))
            {
                TVITEMW state = {};
                state.mask = TVIF_CHILDREN | TVIF_STATE;
                state.hItem = item;
                state.stateMask = TVIS_EXPANDEDONCE;
                state.cChildren = I_CHILDRENCALLBACK;
                TreeView_SetItem(hwnd_, &state);
            }
            return;
        }
        node->children_loaded = AddChildren(item, node);
    }
    for (HTREEITEM child = TreeView_GetChild(hwnd_, item); child; child = TreeView_GetNextSibling(hwnd_, child))
    {
        Resync(child, false);
    }
}

RegistryNode* RegistryTree::OnSelectionChanged(const NMTREEVIEWW* info)
{
    if (!info)
    {
        return nullptr;
    }
    return NodeFromItem(info->itemNew.hItem);
}

RegistryNode* RegistryTree::StoreNode(std::unique_ptr<RegistryNode> node)
{
    // tree lParam holds node pointers, keep them valid
    RegistryNode* stored = node.get();
    nodes_.emplace(stored, std::move(node));
    return stored;
}

void RegistryTree::CollectSubtree(HTREEITEM item, std::vector<RegistryNode*>* nodes)
{
    for (HTREEITEM child = TreeView_GetChild(hwnd_, item); child; child = TreeView_GetNextSibling(hwnd_, child))
    {
        CollectSubtree(child, nodes);
    }
    if (RegistryNode* node = NodeFromItem(item))
    {
        nodes->push_back(node);
    }
}

HTREEITEM RegistryTree::InsertChild(HTREEITEM parent, const std::wstring& name)
{
    RegistryNode* parent_node = NodeFromItem(parent);
    if (!hwnd_ || !parent || !parent_node || !parent_node->children_loaded || name.empty())
    {
        return nullptr;
    }
    HTREEITEM after = TVI_FIRST;
    const std::wstring label = registry_path::DisplayName(name);
    wchar_t text[256] = {};
    for (HTREEITEM sibling = TreeView_GetChild(hwnd_, parent); sibling;
         sibling = TreeView_GetNextSibling(hwnd_, sibling))
    {
        TVITEMW item = {};
        item.mask = TVIF_TEXT;
        item.hItem = sibling;
        item.pszText = text;
        item.cchTextMax = static_cast<int>(_countof(text));
        if (!TreeView_GetItem(hwnd_, &item))
        {
            continue;
        }
        const int order = util::CompareInsensitive(text, label);
        if (order == 0)
        {
            return sibling;
        }
        if (order > 0)
        {
            break;
        }
        after = sibling;
    }

    auto child = std::make_unique<RegistryNode>(registry_path::ChildNode(*parent_node, name));
    child->simulated = false;
    RegistryNode* stored = StoreNode(std::move(child));

    HTREEITEM item = InsertNodeItem(hwnd_, parent, after, label, stored);
    if (item)
    {
        SetChildState(hwnd_, parent, parent_node, true);
        TreeView_Expand(hwnd_, parent, TVE_EXPAND);
    }
    return item;
}

bool RegistryTree::AddChildren(HTREEITEM parent, RegistryNode* node)
{
    if (!node)
    {
        return false;
    }
    std::vector<std::wstring> children;
    const bool enumerated = RegistryStore::EnumKeyStreaming(
        *node,
        false,
        false,
        true,
        nullptr,
        RegistryStore::ValueStreamCallback(),
        [&](const std::wstring& name) {
            children.push_back(name);
            return true;
        },
        MAXDWORD,
        nullptr,
        false
    );
    std::vector<std::wstring> virtual_children;
    // traces can add simulated keys missing from the current key
    if (virtual_child_provider_ && (enumerated || node->simulated))
    {
        std::unordered_set<std::wstring> existing_lower;
        existing_lower.reserve(children.size());
        for (const auto& name : children)
        {
            existing_lower.insert(ToLower(name));
        }
        virtual_child_provider_(*node, existing_lower, &virtual_children);
    }

    struct ChildEntry
    {
        std::wstring name;
        std::wstring label;
        bool simulated = false;
    };
    std::vector<ChildEntry> entries;
    entries.reserve(children.size() + virtual_children.size());
    for (const auto& name : children)
    {
        entries.push_back({name, registry_path::DisplayName(name), false});
    }
    for (const auto& name : virtual_children)
    {
        entries.push_back({name, registry_path::DisplayName(name), true});
    }
    std::sort(entries.begin(), entries.end(), [](const ChildEntry& left, const ChildEntry& right) {
        return util::CompareInsensitive(left.label, right.label) < 0;
    });

    if (!entries.empty())
    {
        nodes_.reserve(nodes_.size() + entries.size());
    }
    std::vector<RegistryNode*> released;
    HTREEITEM child = TreeView_GetChild(hwnd_, parent);
    // inserting after a handle walks sibling items
    const bool fresh = !child;
    for (auto it = entries.rbegin(); fresh && it != entries.rend(); ++it)
    {
        auto added = std::make_unique<RegistryNode>(registry_path::ChildNode(*node, it->name));
        added->simulated = it->simulated;
        InsertNodeItem(hwnd_, parent, TVI_FIRST, it->label, StoreNode(std::move(added)));
    }
    HTREEITEM after = TVI_FIRST;
    wchar_t text[512] = {};
    for (size_t index = fresh ? entries.size() : 0; child || index < entries.size();)
    {
        int order = 1;
        if (child)
        {
            TVITEMW item = {};
            item.mask = TVIF_TEXT;
            item.hItem = child;
            item.pszText = text;
            item.cchTextMax = static_cast<int>(_countof(text));
            order = index == entries.size() || !TreeView_GetItem(hwnd_, &item) ? -1 : util::CompareInsensitive(text, entries[index].label);
        }
        if (order < 0)
        {
            HTREEITEM next = TreeView_GetNextSibling(hwnd_, child);
            CollectSubtree(child, &released);
            TreeView_DeleteItem(hwnd_, child);
            child = next;
            continue;
        }
        if (order == 0)
        {
            if (RegistryNode* existing = NodeFromItem(child))
            {
                existing->simulated = entries[index].simulated;
            }
            after = child;
            child = TreeView_GetNextSibling(hwnd_, child);
        }
        else
        {
            auto added = std::make_unique<RegistryNode>(registry_path::ChildNode(*node, entries[index].name));
            added->simulated = entries[index].simulated;
            after = InsertNodeItem(hwnd_, parent, after, entries[index].label, StoreNode(std::move(added)));
        }
        ++index;
    }
    for (RegistryNode* stale : released)
    {
        nodes_.erase(stale);
    }
    SetChildState(hwnd_, parent, node, !entries.empty());
    return enumerated;
}

void RegistryTree::OnGetDispInfo(NMTVDISPINFOW* info)
{
    if (!info)
    {
        return;
    }
    auto* node = reinterpret_cast<RegistryNode*>(info->item.lParam);
    // defer child state and icon lookup until requested by the tree
    if (info->item.mask & TVIF_CHILDREN)
    {
        if (!node)
        {
            info->item.cChildren = 0;
        }
        else
        {
            if (node->has_children < 0)
            {
                node->has_children = HasChildren(*node) ? 1 : 0;
            }
            info->item.cChildren = node->has_children;
        }
    }
    if (info->item.mask & (TVIF_IMAGE | TVIF_SELECTEDIMAGE))
    {
        int icon = kFolderIconIndex;
        if (node)
        {
            if (node->icon < 0)
            {
                node->icon = icon_resolver_ ? icon_resolver_(*node) : kFolderIconIndex;
            }
            icon = node->icon;
        }
        info->item.iImage = icon;
        info->item.iSelectedImage = icon;
    }
}

bool RegistryTree::HasChildren(const RegistryNode& node)
{
    if (RegistryStore::HasSubKeys(node))
    {
        return true;
    }
    if (!virtual_child_provider_)
    {
        return false;
    }
    std::unordered_set<std::wstring> existing_lower;
    std::vector<std::wstring> virtual_children;
    virtual_child_provider_(node, existing_lower, &virtual_children);
    return !virtual_children.empty();
}

}
