// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::RefreshTreeSelection() {
  if (!browse_.tree().hwnd()) {
    return;
  }
  HTREEITEM item = TreeView_GetSelection(browse_.tree().hwnd());
  if (!item) {
    return;
  }
  RegistryNode* node = browse_.tree().NodeFromItem(item);
  if (!node) {
    return;
  }
  HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), item);
  while (child) {
    HTREEITEM next = TreeView_GetNextSibling(browse_.tree().hwnd(), child);
    TreeView_DeleteItem(browse_.tree().hwnd(), child);
    child = next;
  }
  node->children_loaded = false;
  NMTREEVIEWW info = {};
  info.action = TVE_EXPAND;
  info.itemNew.hItem = item;
  browse_.tree().OnItemExpanding(&info);
  TreeView_Expand(browse_.tree().hwnd(), item, TVE_EXPAND);
  MarkTreeStateDirty();
}

void MainWindow::Impl::UpdateSimulatedChain(HTREEITEM item) {
  if (!browse_.tree().hwnd() || !item) {
    return;
  }
  while (item) {
    RegistryNode* node = browse_.tree().NodeFromItem(item);
    if (node && node->simulated) {
      KeyInfo info = {};
      if (RegistryStore::QueryKeyInfo(*node, &info)) {
        node->simulated = false;
        int icon = KeyIconIndex(*node, nullptr, nullptr);
        TVITEMW tvi = {};
        tvi.mask = TVIF_IMAGE | TVIF_SELECTEDIMAGE;
        tvi.hItem = item;
        tvi.iImage = icon;
        tvi.iSelectedImage = icon;
        TreeView_SetItem(browse_.tree().hwnd(), &tvi);
      }
    }
    item = TreeView_GetParent(browse_.tree().hwnd(), item);
  }
}

void MainWindow::Impl::CaptureTreeState(std::wstring* selected_path, std::vector<std::wstring>* expanded_paths) const {
  if (selected_path) {
    selected_path->clear();
  }
  if (expanded_paths) {
    expanded_paths->clear();
  }
  if (!browse_.tree().hwnd()) {
    return;
  }
  if (selected_path) {
    RegistryNode* node = browse_.current_node();
    if (!node) {
      HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
      if (selected) {
        TVITEMW tvi = {};
        tvi.hItem = selected;
        tvi.mask = TVIF_PARAM;
        if (TreeView_GetItem(browse_.tree().hwnd(), &tvi)) {
          node = reinterpret_cast<RegistryNode*>(tvi.lParam);
        }
      }
    }
    if (node) {
      *selected_path = registry_path::Build(*node);
    }
  }
  if (!expanded_paths) {
    return;
  }
  HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
  if (!root) {
    return;
  }
  std::function<void(HTREEITEM, bool)> walk = [&](HTREEITEM item, bool ancestors_expanded) {
    while (item) {
      TVITEMW tvi = {};
      tvi.hItem = item;
      tvi.mask = TVIF_STATE | TVIF_PARAM;
      tvi.stateMask = TVIS_EXPANDED;
      if (TreeView_GetItem(browse_.tree().hwnd(), &tvi)) {
        bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
        if (ancestors_expanded && expanded) {
          RegistryNode* node = reinterpret_cast<RegistryNode*>(tvi.lParam);
          if (node) {
            expanded_paths->push_back(registry_path::Build(*node));
          }
        }
      }
      HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), item);
      if (child) {
        bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
        if (ancestors_expanded && expanded) {
          walk(child, true);
        }
      }
      item = TreeView_GetNextSibling(browse_.tree().hwnd(), item);
    }
  };
  walk(root, true);
}

void MainWindow::Impl::RestoreTreeState() {
  if (tree_state_restored_) {
    return;
  }
  if (!save_tree_state_) {
    return;
  }
  tree_state_restored_ = true;
  if (!browse_.tree().hwnd()) {
    return;
  }
  workspace::TreeState state = saved_tree_state_;
  state.Normalize();
  for (const auto& path : state.expanded_paths) {
    ExpandTreePath(path);
  }
  if (!state.selected_path.empty()) {
    SelectTreePath(state.selected_path);
  }
}

void MainWindow::Impl::ApplySavedWindowPlacement() {
  if (!window_placement_loaded_ || !hwnd_) {
    return;
  }
  const int min_width = 640;
  const int min_height = 480;
  int width = std::max(window_width_, min_width);
  int height = std::max(window_height_, min_height);
  if (window_width_ <= 0 || window_height_ <= 0) {
    return;
  }
  SetWindowPos(hwnd_, nullptr, window_x_, window_y_, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
}

bool MainWindow::Impl::ExpandTreePath(const std::wstring& path) {
  if (!browse_.tree().hwnd()) {
    return false;
  }
  std::vector<std::wstring> parts = BuildVisibleTreePathParts(path);
  if (parts.empty()) {
    return false;
  }
  HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
  HTREEITEM current = root;
  for (const auto& part : parts) {
    TreeView_Expand(browse_.tree().hwnd(), current, TVE_EXPAND);
    HTREEITEM child = FindChildByText(browse_.tree().hwnd(), current, part);
    if (!child) {
      return false;
    }
    current = child;
  }
  if (current) {
    TreeView_Expand(browse_.tree().hwnd(), current, TVE_EXPAND);
    return true;
  }
  return false;
}

} // namespace regkit
