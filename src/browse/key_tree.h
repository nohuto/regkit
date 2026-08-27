// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>
#include <commctrl.h>

#include <functional>
#include <memory>
#include <unordered_set>
#include <vector>

#include "registry/registry_store.h"

namespace regkit {

class RegistryTree {
public:
  void Create(HWND parent, HINSTANCE instance, int control_id, bool show_border = true);
  HWND hwnd() const;
  void SetImageList(HIMAGELIST image_list);
  void SetIconResolver(std::function<int(const RegistryNode&)> resolver);
  void SetVirtualChildProvider(std::function<void(const RegistryNode&, const std::unordered_set<std::wstring>&, std::vector<std::wstring>*)> provider);
  void SetRootLabel(const std::wstring& label);
  void SetRegeditLayout(bool enabled);

  void PopulateRoots(const std::vector<RegistryRootEntry>& roots);
  RegistryNode* NodeFromItem(HTREEITEM item);
  void OnItemExpanding(const NMTREEVIEWW* info);
  RegistryNode* OnSelectionChanged(const NMTREEVIEWW* info);

private:
  RegistryNode* StoreNode(std::unique_ptr<RegistryNode> node);
  void AddChildren(HTREEITEM parent, RegistryNode* node);
  void AddDummyChildIfNeeded(HTREEITEM parent, RegistryNode* node);

  HWND hwnd_ = nullptr;
  HTREEITEM root_item_ = nullptr;
  HTREEITEM standard_group_item_ = nullptr;
  HTREEITEM real_group_item_ = nullptr;
  std::vector<std::unique_ptr<RegistryNode>> nodes_;
  std::function<int(const RegistryNode&)> icon_resolver_;
  std::function<void(const RegistryNode&, const std::unordered_set<std::wstring>&, std::vector<std::wstring>*)> virtual_child_provider_;
  std::wstring root_label_ = L"Computer";
  bool regedit_layout_ = false;
};

} // namespace regkit
