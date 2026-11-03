// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "win32/text_transform.h"

#include <win32ui/window_metrics.h>
#include <win32ui/window_placement.h>

namespace regkit
{
using namespace window_detail;

void MainWindow::Impl::RefreshTreePath(const std::wstring& path)
{
    RefreshTreeItem(FindTreeItem(path));
}

namespace
{

bool IsAncestorItem(HWND tree, HTREEITEM candidate, HTREEITEM item)
{
    for (HTREEITEM parent = TreeView_GetParent(tree, item); parent; parent = TreeView_GetParent(tree, parent))
    {
        if (parent == candidate)
        {
            return true;
        }
    }
    return false;
}

std::wstring KernelName(const RegistryNode& node)
{
    KeyDetails details;
    return RegistryStore::QueryKeyDetails(node, &details) ? details.native.native_name : std::wstring();
}

std::wstring NativePath(const RegistryNode& node)
{
    return node.view ? std::wstring() : registry_path::BuildNative(node);
}

}

void MainWindow::Impl::RefreshMatchingTreeNodes(HTREEITEM selected)
{
    HWND tree = browse_.tree().hwnd();
    if (tree && !selected)
    {
        selected = TreeView_GetSelection(tree);
    }
    const RegistryNode* target = selected ? browse_.tree().NodeFromItem(selected) : nullptr;
    if (!target)
    {
        return;
    }
    const std::wstring native = NativePath(*target);
    const std::wstring kernel = KernelName(*target);
    if (native.empty() && kernel.empty())
    {
        return;
    }
    const std::wstring wanted = win32ui::TreeItemText(tree, selected);
    if (wanted.empty())
    {
        return;
    }

    std::vector<std::pair<int, HTREEITEM>> matches;
    std::vector<std::pair<int, HTREEITEM>> pending;
    if (HTREEITEM root = TreeView_GetRoot(tree))
    {
        pending.emplace_back(0, root);
    }
    while (!pending.empty())
    {
        const int depth = pending.back().first;
        HTREEITEM item = pending.back().second;
        pending.pop_back();
        for (HTREEITEM child = TreeView_GetChild(tree, item); child; child = TreeView_GetNextSibling(tree, child))
        {
            pending.emplace_back(depth + 1, child);
        }
        const RegistryNode* node = browse_.tree().NodeFromItem(item);
        if (item == selected || !node || !node->children_loaded || IsAncestorItem(tree, item, selected))
        {
            continue;
        }
        const std::wstring path = NativePath(*node);
        if ((!path.empty() && (util::EqualsInsensitive(path, native) || util::EqualsInsensitive(path, kernel))) ||
            (!kernel.empty() && util::EqualsInsensitive(win32ui::TreeItemText(tree, item), wanted) && util::EqualsInsensitive(KernelName(*node), kernel)))
        {
            matches.emplace_back(depth, item);
        }
    }

    if (matches.empty())
    {
        return;
    }
    std::sort(matches.begin(), matches.end(), [](const std::pair<int, HTREEITEM>& left, const std::pair<int, HTREEITEM>& right) {
        return left.first > right.first;
    });
    HTREEITEM first_visible = TreeView_GetFirstVisible(tree);
    browse_.tree().SuspendRedraw();
    for (const auto& match : matches)
    {
        RefreshTreeItem(match.second);
    }
    if (first_visible)
    {
        TreeView_SelectSetFirstVisible(tree, first_visible);
    }
    browse_.tree().ResumeRedraw();
}

void MainWindow::Impl::RefreshTreeSelection()
{
    // other tab trees resync when shown
    RegistryStore::NoteKeyChange();
    if (!browse_.tree().hwnd())
    {
        return;
    }
    HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
    RefreshTreeItem(selected);
    RefreshMatchingTreeNodes(selected);
}

void MainWindow::Impl::RefreshTreeItem(HTREEITEM item)
{
    if (!browse_.tree().hwnd() || !item)
    {
        return;
    }
    browse_.tree().SuspendRedraw();
    browse_.tree().Resync(item);
    browse_.tree().ResumeRedraw();
    MarkTreeStateDirty();
}

void MainWindow::Impl::RefreshWholeTree()
{
    if (browse_.tree().hwnd())
    {
        // sample version before resync so changes during it still mark the tree stale
        const uint64_t revision = RegistryStore::KeyRevision();
        RefreshTreeItem(TreeView_GetRoot(browse_.tree().hwnd()));
        browse_.tree().synced_revision = revision;
    }
}

void MainWindow::Impl::UpdateSimulatedChain(HTREEITEM item)
{
    if (!browse_.tree().hwnd() || !item)
    {
        return;
    }
    while (item)
    {
        RegistryNode* node = browse_.tree().NodeFromItem(item);
        if (node && node->simulated)
        {
            KeyInfo info = {};
            if (RegistryStore::QueryKeyInfo(*node, &info))
            {
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

void MainWindow::Impl::CaptureTreeState(const RegistryTree& tree, std::wstring* selected_path, std::vector<std::wstring>* expanded_paths) const
{
    if (selected_path)
    {
        selected_path->clear();
    }
    if (expanded_paths)
    {
        expanded_paths->clear();
    }
    const HWND hwnd = tree.hwnd();
    if (!hwnd)
    {
        return;
    }
    if (selected_path)
    {
        RegistryNode* node = tree.current_node();
        if (!node)
        {
            node = tree.NodeFromItem(TreeView_GetSelection(hwnd));
        }
        if (node)
        {
            *selected_path = registry_path::Build(*node);
        }
    }
    if (!expanded_paths)
    {
        return;
    }
    HTREEITEM root = TreeView_GetRoot(hwnd);
    if (!root)
    {
        return;
    }
    std::function<void(HTREEITEM, bool)> walk = [&](HTREEITEM item, bool ancestors_expanded) {
        while (item)
        {
            TVITEMW tvi = {};
            tvi.hItem = item;
            tvi.mask = TVIF_STATE | TVIF_PARAM;
            tvi.stateMask = TVIS_EXPANDED;
            if (TreeView_GetItem(hwnd, &tvi))
            {
                bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
                if (ancestors_expanded && expanded)
                {
                    RegistryNode* node = reinterpret_cast<RegistryNode*>(tvi.lParam);
                    if (node)
                    {
                        expanded_paths->push_back(registry_path::Build(*node));
                    }
                }
            }
            HTREEITEM child = TreeView_GetChild(hwnd, item);
            if (child)
            {
                bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
                if (ancestors_expanded && expanded)
                {
                    walk(child, true);
                }
            }
            item = TreeView_GetNextSibling(hwnd, item);
        }
    };
    walk(root, true);
}

// global state belongs to the plain local tab opened without saved tabs
void MainWindow::Impl::RestoreTreeState(workspace::TreeState state)
{
    const int local = FindLocalRegistryTabIndex();
    if (local < 0 || !browse_.tree().hwnd())
    {
        return;
    }
    TabEntry& entry = tabs_[static_cast<size_t>(local)];
    if (!entry.selected_path.empty() || !entry.expanded_paths.empty())
    {
        return;
    }
    state.Normalize();
    if (local != TabCtrl_GetCurSel(tab_))
    {
        entry.selected_path = std::move(state.selected_path);
        entry.expanded_paths = std::move(state.expanded_paths);
        return;
    }
    ExpandTreePaths(state.expanded_paths);
    if (!state.selected_path.empty())
    {
        SelectTreePath(state.selected_path);
    }
}

void MainWindow::Impl::ApplySavedWindowPlacement()
{
    if (!hwnd_)
    {
        return;
    }
    const UINT dpi = win32ui::DpiForWindow(hwnd_);
    if (!window_placement_loaded_ || settings_.window_width <= 0 || settings_.window_height <= 0)
    {
        // first start keeps the default position
        SetWindowPos(hwnd_, nullptr, 0, 0, win32ui::Scale(1200, dpi), win32ui::Scale(800, dpi), SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        return;
    }
    // files without window_dpi were saved at the current DPI
    const UINT saved_dpi = settings_.window_dpi > 0 ? static_cast<UINT>(settings_.window_dpi) : dpi;
    const int width = std::max(settings_.window_width, win32ui::Scale(640, saved_dpi));
    const int height = std::max(settings_.window_height, win32ui::Scale(480, saved_dpi));
    // saved in workspace coordinates, as GetWindowPlacement reports them
    const RECT saved = {settings_.window_x, settings_.window_y, settings_.window_x + width, settings_.window_y + height};
    win32ui::RestoreNormalRect(hwnd_, &saved, saved_dpi, {});
}

HTREEITEM MainWindow::Impl::FindTreeItem(const std::wstring& path)
{
    if (!browse_.tree().hwnd())
    {
        return nullptr;
    }
    std::vector<std::wstring> parts = BuildVisibleTreePathParts(path);
    if (parts.empty())
    {
        return nullptr;
    }
    HTREEITEM current = TreeView_GetRoot(browse_.tree().hwnd());
    for (const auto& part : parts)
    {
        TreeView_Expand(browse_.tree().hwnd(), current, TVE_EXPAND);
        HTREEITEM child = FindChildByText(browse_.tree().hwnd(), current, part);
        if (!child)
        {
            return nullptr;
        }
        current = child;
    }
    return current;
}

void MainWindow::Impl::ExpandTreePaths(const std::vector<std::wstring>& paths)
{
    HWND tree = browse_.tree().hwnd();
    if (!tree || paths.empty())
    {
        return;
    }
    std::unordered_set<std::wstring> wanted;
    wanted.reserve(paths.size());
    for (const auto& path : paths)
    {
        std::wstring lower = ToLower(path);
        while (!lower.empty())
        {
            if (!wanted.insert(lower).second)
            {
                break;
            }
            const size_t separator = lower.rfind(L'\\');
            lower = separator == std::wstring::npos ? std::wstring() : lower.substr(0, separator);
        }
    }
    if (wanted.empty())
    {
        return;
    }
    std::vector<HTREEITEM> pending;
    if (HTREEITEM root = TreeView_GetRoot(tree))
    {
        pending.push_back(root);
    }
    while (!pending.empty())
    {
        HTREEITEM item = pending.back();
        pending.pop_back();
        RegistryNode* node = browse_.tree().NodeFromItem(item);
        if (node && wanted.find(ToLower(registry_path::Build(*node))) == wanted.end())
        {
            continue;
        }
        TreeView_Expand(tree, item, TVE_EXPAND);
        for (HTREEITEM child = TreeView_GetChild(tree, item); child; child = TreeView_GetNextSibling(tree, child))
        {
            pending.push_back(child);
        }
    }
}

void MainWindow::Impl::MarkTreeStateDirty()
{
    if (!settings_.save_tree_state || !hwnd_ || !browse_.tree().hwnd() || !IsWindow(browse_.tree().hwnd()))
    {
        return;
    }
    SetTimer(hwnd_, kTreeStateTimerId, 400, nullptr);
}

void MainWindow::Impl::CaptureTreeStateNow()
{
    workspace::TreeState state;
    if (CaptureLocalTreeState(&state, true))
    {
        tree_state_saver_.Submit(std::move(state));
    }
}

// shown_only skips hidden local tabs because their state cannot change
bool MainWindow::Impl::CaptureLocalTreeState(workspace::TreeState* state, bool shown_only) const
{
    const int local = FindLocalRegistryTabIndex();
    if (!settings_.save_tree_state || local < 0 || (shown_only && local != TabCtrl_GetCurSel(tab_)))
    {
        return false;
    }
    const TabEntry& entry = tabs_[static_cast<size_t>(local)];
    if (entry.tree && IsWindow(entry.tree->hwnd()))
    {
        CaptureTreeState(*entry.tree, &state->selected_path, &state->expanded_paths);
        return true;
    }
    state->selected_path = entry.selected_path;
    state->expanded_paths = entry.expanded_paths;
    return true;
}

}
