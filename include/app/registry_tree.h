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

#pragma once

#include <windows.h>
#include <commctrl.h>

#include <functional>
#include <memory>
#include <unordered_set>
#include <vector>

#include "registry/registry_provider.h"

namespace regkit {

class RegistryTree {
public:
  void Create(HWND parent, HINSTANCE instance, int control_id, bool show_border = true);
  HWND hwnd() const;
  void SetImageList(HIMAGELIST image_list);
  void SetIconResolver(std::function<int(const RegistryNode&)> resolver);
  void SetVirtualChildProvider(std::function<void(const RegistryNode&, const std::unordered_set<std::wstring>&, std::vector<std::wstring>*)> provider);
  void SetRootLabel(const std::wstring& label);

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
};

} // namespace regkit
