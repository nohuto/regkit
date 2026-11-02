// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <commctrl.h>

#include <functional>
#include <memory>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "registry/registry_store.h"

#include <win32ui/tree_view/tree_view.h>

namespace regkit
{

class RegistryTree
{
  public:
    void Create(HWND parent, HINSTANCE instance, int control_id, bool show_border = true, bool allow_label_edit = false);
    HWND hwnd() const;
    void SetImageList(HIMAGELIST image_list);
    void SetIconResolver(std::function<int(const RegistryNode&)> resolver);
    void SetVirtualChildProvider(
        std::function<void(const RegistryNode&, const std::unordered_set<std::wstring>&, std::vector<std::wstring>*)>
            provider
    );
    void SetRootLabel(const std::wstring& label, int icon = 0);
    void SetRegEditLayout(bool enabled);
    void SuspendRedraw();
    void ResumeRedraw();
    bool IsRedrawSuspended() const noexcept
    {
        return redraw_suspended_ > 0;
    }

    void PopulateRoots(const std::vector<RegistryRootEntry>& roots);
    void Clear();
    RegistryNode* current_node() const noexcept
    {
        return current_node_;
    }
    void set_current_node(RegistryNode* node) noexcept
    {
        current_node_ = node;
    }
    std::vector<RegistryRootEntry>& roots() noexcept
    {
        return roots_;
    }
    uint64_t synced_revision = 0;
    RegistryNode* NodeFromItem(HTREEITEM item) const;
    HTREEITEM InsertChild(HTREEITEM parent, const std::wstring& name);
    void Resync(HTREEITEM item, bool load = true);
    void OnItemExpanding(const NMTREEVIEWW* info);
    void OnGetDispInfo(NMTVDISPINFOW* info);
    RegistryNode* OnSelectionChanged(const NMTREEVIEWW* info);
    bool IsGroupItem(HTREEITEM item) const noexcept
    {
        return item && (item == standard_group_item_ || item == real_group_item_);
    }
    void SetExpandAllCommand(HWND window, int command) noexcept
    {
        expand_all_window_ = window;
        expand_all_command_ = command;
    }

  private:
    void LoadChildren(HTREEITEM item);
    RegistryNode* StoreNode(std::unique_ptr<RegistryNode> node);
    bool AddChildren(HTREEITEM parent, RegistryNode* node);
    bool HasChildren(const RegistryNode& node);
    void CollectSubtree(HTREEITEM item, std::vector<RegistryNode*>* nodes);

    HWND hwnd_ = nullptr;
    HTREEITEM root_item_ = nullptr;
    HTREEITEM standard_group_item_ = nullptr;
    HTREEITEM real_group_item_ = nullptr;
    std::unordered_map<RegistryNode*, std::unique_ptr<RegistryNode>> nodes_;
    std::vector<RegistryRootEntry> roots_;
    RegistryNode* current_node_ = nullptr;
    std::function<int(const RegistryNode&)> icon_resolver_;
    std::function<void(const RegistryNode&, const std::unordered_set<std::wstring>&, std::vector<std::wstring>*)>
        virtual_child_provider_;
    std::wstring root_label_ = L"Computer";
    int root_icon_ = 0;
    int redraw_suspended_ = 0;
    bool regedit_layout_ = false;
    win32ui::TreeViewHooks hooks_;
    HWND expand_all_window_ = nullptr;
    int expand_all_command_ = 0;
};

}
