// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::RefreshTreePath(const std::wstring& path) {
  RefreshTreeItem(FindTreeItem(path));
}

void MainWindow::Impl::RefreshMatchingTreeNodes() {
  HWND tree = browse_.tree().hwnd();
  HTREEITEM selected = tree ? TreeView_GetSelection(tree) : nullptr;
  if (!selected) {
    return;
  }
  auto label_of = [&](HTREEITEM item, wchar_t* buffer, int size) -> bool {
    TVITEMW info = {};
    info.hItem = item;
    info.mask = TVIF_TEXT;
    info.pszText = buffer;
    info.cchTextMax = size;
    return TreeView_GetItem(tree, &info) != FALSE && buffer[0] != 0;
  };
  wchar_t wanted[256] = {};
  if (!label_of(selected, wanted, static_cast<int>(_countof(wanted)))) {
    return;
  }

  std::vector<std::pair<int, HTREEITEM>> matches;
  std::vector<std::pair<int, HTREEITEM>> pending;
  if (HTREEITEM root = TreeView_GetRoot(tree)) {
    pending.emplace_back(0, root);
  }
  while (!pending.empty()) {
    const int depth = pending.back().first;
    HTREEITEM item = pending.back().second;
    pending.pop_back();
    for (HTREEITEM child = TreeView_GetChild(tree, item); child;
         child = TreeView_GetNextSibling(tree, child)) {
      pending.emplace_back(depth + 1, child);
    }
    if (item == selected || !browse_.tree().NodeFromItem(item)) {
      continue;
    }
    wchar_t text[256] = {};
    if (label_of(item, text, static_cast<int>(_countof(text))) &&
        _wcsicmp(text, wanted) == 0) {
      matches.emplace_back(depth, item);
    }
  }

  // Deepest first, so refreshing a shallower match cannot invalidate a
  // handle that is still queued.
  std::sort(matches.begin(), matches.end(),
            [](const std::pair<int, HTREEITEM>& left,
               const std::pair<int, HTREEITEM>& right) {
              return left.first > right.first;
            });
  for (const auto& match : matches) {
    const bool was_expanded =
        (TreeView_GetItemState(tree, match.second, TVIS_EXPANDED) & TVIS_EXPANDED) != 0;
    RefreshTreeItem(match.second);
    if (!was_expanded) {
      TreeView_Expand(tree, match.second, TVE_COLLAPSE);
    }
  }
}

void MainWindow::Impl::RefreshTreeSelection() {
  if (!browse_.tree().hwnd()) {
    return;
  }
  RefreshTreeItem(TreeView_GetSelection(browse_.tree().hwnd()));
}

void MainWindow::Impl::RefreshTreeItem(HTREEITEM item) {
  if (!browse_.tree().hwnd() || !item) {
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
  if (window_width_ <= 0 || window_height_ <= 0) {
    return;
  }
  const int min_width = 640;
  const int min_height = 480;
  const int width = std::max(window_width_, min_width);
  const int height = std::max(window_height_, min_height);

  RECT work = {};
  if (!SystemParametersInfoW(SPI_GETWORKAREA, 0, &work, 0)) {
    work = {};
  }
  RECT target = {window_x_ + work.left, window_y_ + work.top,
                 window_x_ + work.left + width, window_y_ + work.top + height};
  win32::ClampToWorkArea(&target);
  SetWindowPos(hwnd_, nullptr, target.left, target.top,
               target.right - target.left, target.bottom - target.top,
               SWP_NOZORDER | SWP_NOACTIVATE);
}

HTREEITEM MainWindow::Impl::FindTreeItem(const std::wstring& path) {
  if (!browse_.tree().hwnd()) {
    return nullptr;
  }
  std::vector<std::wstring> parts = BuildVisibleTreePathParts(path);
  if (parts.empty()) {
    return nullptr;
  }
  HTREEITEM current = TreeView_GetRoot(browse_.tree().hwnd());
  for (const auto& part : parts) {
    TreeView_Expand(browse_.tree().hwnd(), current, TVE_EXPAND);
    HTREEITEM child = FindChildByText(browse_.tree().hwnd(), current, part);
    if (!child) {
      return nullptr;
    }
    current = child;
  }
  return current;
}

bool MainWindow::Impl::ExpandTreePath(const std::wstring& path) {
  HTREEITEM item = FindTreeItem(path);
  if (!item) {
    return false;
  }
  TreeView_Expand(browse_.tree().hwnd(), item, TVE_EXPAND);
  return true;
}

} // namespace regkit
