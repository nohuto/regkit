// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#include "app/registry_tree.h"

#include <algorithm>
#include <cwctype>

namespace regkit {

namespace {
constexpr int kFolderIconIndex = 0;
constexpr wchar_t kStandardGroupLabel[] = L"Standart Hives";
constexpr wchar_t kRealGroupLabel[] = L"REGISTRY";
#ifndef TVS_EX_DOUBLEBUFFER
#define TVS_EX_DOUBLEBUFFER 0x0004
#endif

std::wstring ToLower(const std::wstring& text) {
  std::wstring out = text;
  for (auto& ch : out) {
    ch = static_cast<wchar_t>(towlower(ch));
  }
  return out;
}
} // namespace

void RegistryTree::Create(HWND parent, HINSTANCE instance, int control_id, bool show_border) {
  DWORD style = WS_CHILD | WS_VISIBLE | TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS | TVS_EDITLABELS;
  if (show_border) {
    style |= WS_BORDER;
  }
  hwnd_ = CreateWindowExW(0, WC_TREEVIEWW, L"", style, 0, 0, 100, 100, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)), instance, nullptr);
  if (hwnd_) {
    TreeView_SetExtendedStyle(hwnd_, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);
  }
}

HWND RegistryTree::hwnd() const {
  return hwnd_;
}

void RegistryTree::SetImageList(HIMAGELIST image_list) {
  if (!hwnd_) {
    return;
  }
  TreeView_SetImageList(hwnd_, image_list, TVSIL_NORMAL);
}

void RegistryTree::SetIconResolver(std::function<int(const RegistryNode&)> resolver) {
  icon_resolver_ = std::move(resolver);
}

void RegistryTree::SetVirtualChildProvider(std::function<void(const RegistryNode&, const std::unordered_set<std::wstring>&, std::vector<std::wstring>*)> provider) {
  virtual_child_provider_ = std::move(provider);
}

void RegistryTree::SetRootLabel(const std::wstring& label) {
  if (label.empty()) {
    root_label_ = L"Computer";
    return;
  }
  root_label_ = label;
}

void RegistryTree::PopulateRoots(const std::vector<RegistryRootEntry>& roots) {
  TreeView_DeleteAllItems(hwnd_);
  nodes_.clear();
  nodes_.reserve(roots.size() + 3);

  TVINSERTSTRUCTW insert = {};
  insert.hParent = TVI_ROOT;
  insert.hInsertAfter = TVI_LAST;
  insert.item.mask = TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM;
  insert.item.pszText = const_cast<wchar_t*>(root_label_.c_str());
  insert.item.iImage = kFolderIconIndex;
  insert.item.iSelectedImage = kFolderIconIndex;
  insert.item.lParam = 0;
  root_item_ = TreeView_InsertItem(hwnd_, &insert);
  standard_group_item_ = nullptr;
  real_group_item_ = nullptr;

  bool has_standard = false;
  bool has_real = false;
  for (const auto& root_entry : roots) {
    if (root_entry.group == RegistryRootGroup::kReal) {
      has_real = true;
    } else {
      has_standard = true;
    }
  }
  if (has_standard) {
    TVINSERTSTRUCTW standard_group = {};
    standard_group.hParent = root_item_;
    standard_group.hInsertAfter = TVI_LAST;
    standard_group.item.mask = TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM;
    standard_group.item.pszText = const_cast<wchar_t*>(kStandardGroupLabel);
    standard_group.item.iImage = kFolderIconIndex;
    standard_group.item.iSelectedImage = kFolderIconIndex;
    standard_group.item.lParam = 0;
    standard_group_item_ = TreeView_InsertItem(hwnd_, &standard_group);
  }
  if (has_real) {
    TVINSERTSTRUCTW real_root = {};
    real_root.hParent = root_item_;
    real_root.hInsertAfter = TVI_LAST;
    real_root.item.mask = TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM;
    real_root.item.pszText = const_cast<wchar_t*>(kRealGroupLabel);
    real_root.item.iImage = kFolderIconIndex;
    real_root.item.iSelectedImage = kFolderIconIndex;
    real_root.item.lParam = 0;
    real_group_item_ = TreeView_InsertItem(hwnd_, &real_root);
  }

  for (const auto& root_entry : roots) {
    auto node = std::make_unique<RegistryNode>();
    node->root = root_entry.root;
    node->subkey = root_entry.subkey_prefix;
    node->root_name = root_entry.path_name;
    RegistryNode* stored = StoreNode(std::move(node));

    int icon_index = kFolderIconIndex;
    if (icon_resolver_) {
      icon_index = icon_resolver_(*stored);
    }

    if (root_entry.group == RegistryRootGroup::kReal && real_group_item_ && _wcsicmp(root_entry.display_name.c_str(), kRealGroupLabel) == 0 && root_entry.subkey_prefix.empty()) {
      TVITEMW item = {};
      item.mask = TVIF_PARAM | TVIF_IMAGE | TVIF_SELECTEDIMAGE;
      item.hItem = real_group_item_;
      item.lParam = reinterpret_cast<LPARAM>(stored);
      item.iImage = icon_index;
      item.iSelectedImage = icon_index;
      TreeView_SetItem(hwnd_, &item);
      AddDummyChildIfNeeded(real_group_item_, stored);
      continue;
    }

    TVINSERTSTRUCTW root_item = {};
    if (root_entry.group == RegistryRootGroup::kReal) {
      root_item.hParent = real_group_item_ ? real_group_item_ : root_item_;
    } else {
      root_item.hParent = standard_group_item_ ? standard_group_item_ : root_item_;
    }
    root_item.hInsertAfter = TVI_LAST;
    root_item.item.mask = TVIF_TEXT | TVIF_PARAM | TVIF_IMAGE | TVIF_SELECTEDIMAGE;
    root_item.item.pszText = const_cast<wchar_t*>(root_entry.display_name.c_str());
    root_item.item.lParam = reinterpret_cast<LPARAM>(stored);
    root_item.item.iImage = icon_index;
    root_item.item.iSelectedImage = icon_index;
    HTREEITEM item = TreeView_InsertItem(hwnd_, &root_item);
    AddDummyChildIfNeeded(item, stored);
  }

  TreeView_Expand(hwnd_, root_item_, TVE_EXPAND);
  if (standard_group_item_) {
    TreeView_Expand(hwnd_, standard_group_item_, TVE_EXPAND);
  }
  if (real_group_item_) {
    TreeView_Expand(hwnd_, real_group_item_, TVE_EXPAND);
  }
}

RegistryNode* RegistryTree::NodeFromItem(HTREEITEM item) {
  if (!item) {
    return nullptr;
  }
  TVITEMW tvi = {};
  tvi.hItem = item;
  tvi.mask = TVIF_PARAM;
  if (!TreeView_GetItem(hwnd_, &tvi)) {
    return nullptr;
  }
  return reinterpret_cast<RegistryNode*>(tvi.lParam);
}

void RegistryTree::OnItemExpanding(const NMTREEVIEWW* info) {
  if (!info || info->action != TVE_EXPAND) {
    return;
  }
  RegistryNode* node = NodeFromItem(info->itemNew.hItem);
  if (!node || node->children_loaded) {
    return;
  }

  HTREEITEM child = TreeView_GetChild(hwnd_, info->itemNew.hItem);
  if (child) {
    TreeView_DeleteItem(hwnd_, child);
  }
  AddChildren(info->itemNew.hItem, node);
  node->children_loaded = true;
}

RegistryNode* RegistryTree::OnSelectionChanged(const NMTREEVIEWW* info) {
  if (!info) {
    return nullptr;
  }
  return NodeFromItem(info->itemNew.hItem);
}

RegistryNode* RegistryTree::StoreNode(std::unique_ptr<RegistryNode> node) {
  nodes_.push_back(std::move(node));
  return nodes_.back().get();
}

void RegistryTree::AddChildren(HTREEITEM parent, RegistryNode* node) {
  if (!node) {
    return;
  }
  auto children = RegistryProvider::EnumSubKeyNames(*node);
  std::unordered_set<std::wstring> existing_lower;
  existing_lower.reserve(children.size());
  for (const auto& name : children) {
    existing_lower.insert(ToLower(name));
  }
  std::vector<std::wstring> virtual_children;
  if (virtual_child_provider_) {
    virtual_child_provider_(*node, existing_lower, &virtual_children);
  }

  struct ChildEntry {
    std::wstring name;
    bool simulated = false;
  };
  std::vector<ChildEntry> entries;
  entries.reserve(children.size() + virtual_children.size());
  for (const auto& name : children) {
    entries.push_back({name, false});
  }
  for (const auto& name : virtual_children) {
    entries.push_back({name, true});
  }
  std::sort(entries.begin(), entries.end(), [](const ChildEntry& left, const ChildEntry& right) { return _wcsicmp(left.name.c_str(), right.name.c_str()) < 0; });

  if (!entries.empty()) {
    nodes_.reserve(nodes_.size() + entries.size());
  }
  for (const auto& entry : entries) {
    const std::wstring& name = entry.name;
    auto child = std::make_unique<RegistryNode>();
    child->root = node->root;
    child->root_name = node->root_name;
    if (node->subkey.empty()) {
      child->subkey = name;
    } else {
      child->subkey = node->subkey + L"\\" + name;
    }
    child->simulated = entry.simulated;
    RegistryNode* stored = StoreNode(std::move(child));

    int icon_index = kFolderIconIndex;
    if (icon_resolver_) {
      icon_index = icon_resolver_(*stored);
    }

    TVINSERTSTRUCTW insert = {};
    insert.hParent = parent;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM | TVIF_IMAGE | TVIF_SELECTEDIMAGE;
    insert.item.pszText = const_cast<wchar_t*>(name.c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(stored);
    insert.item.iImage = icon_index;
    insert.item.iSelectedImage = icon_index;
    HTREEITEM item = TreeView_InsertItem(hwnd_, &insert);
    AddDummyChildIfNeeded(item, stored);
  }
}

void RegistryTree::AddDummyChildIfNeeded(HTREEITEM parent, RegistryNode* node) {
  if (!node) {
    return;
  }
  bool has_children = RegistryProvider::HasSubKeys(*node);
  if (!has_children && virtual_child_provider_) {
    std::unordered_set<std::wstring> existing_lower;
    std::vector<std::wstring> virtual_children;
    virtual_child_provider_(*node, existing_lower, &virtual_children);
    has_children = !virtual_children.empty();
  }
  if (!has_children) {
    return;
  }
  TVINSERTSTRUCTW insert = {};
  insert.hParent = parent;
  insert.hInsertAfter = TVI_LAST;
  insert.item.mask = TVIF_TEXT | TVIF_PARAM;
  insert.item.pszText = const_cast<wchar_t*>(L"");
  insert.item.lParam = 0;
  TreeView_InsertItem(hwnd_, &insert);
}

} // namespace regkit
