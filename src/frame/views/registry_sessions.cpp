// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

#include "win32/file_dialog.h"

#include <win32ui/dialog/file_dialog.h>
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/dialog/text_prompt.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/process_token.h>
#include <win32ui/window.h>

namespace regkit
{
using namespace window_detail;

namespace
{

bool SaveHiveAtomically(HKEY root, const std::wstring& path, std::wstring* error)
{
    std::wstring temp;
    for (int attempt = 0; attempt < 16 && temp.empty(); ++attempt)
    {
        const std::wstring candidate = path + util::RandomFileSuffix(L".part");
        if (GetFileAttributesW(candidate.c_str()) == INVALID_FILE_ATTRIBUTES && GetLastError() == ERROR_FILE_NOT_FOUND)
        {
            temp = candidate;
        }
    }
    if (temp.empty())
    {
        if (error)
        {
            *error = FormatWin32Error(ERROR_FILE_EXISTS);
        }
        return false;
    }
    if (!RegistryStore::SaveOfflineHive(root, temp, error))
    {
        DeleteFileW(temp.c_str());
        return false;
    }
    if (!MoveFileExW(temp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH))
    {
        if (error)
        {
            *error = FormatWin32Error(GetLastError());
        }
        DeleteFileW(temp.c_str());
        return false;
    }
    return true;
}

}

MainWindow::Impl::RegistrySession::~RegistrySession()
{
    for (HKEY key : {remote_hklm, remote_hku})
    {
        if (key)
        {
            RegCloseKey(key);
        }
    }
    for (HKEY root : offline_roots)
    {
        RegistryStore::RemoveOfflineRoot(root);
        RegistryStore::CloseOfflineHive(root, nullptr);
    }
}

MainWindow::Impl::RegistrySession* MainWindow::Impl::CurrentTabSession()
{
    const int index = CurrentRegistryTabIndex();
    return index >= 0 && static_cast<size_t>(index) < tabs_.size() ? tabs_[static_cast<size_t>(index)].session.get() : nullptr;
}

void MainWindow::Impl::ShowSession(const std::shared_ptr<RegistrySession>& session)
{
    const int index = CurrentRegistryTabIndex();
    if (index >= 0 && static_cast<size_t>(index) < tabs_.size())
    {
        tabs_[static_cast<size_t>(index)].session = session;
    }
    const bool shown = session_ == session && !browse_.roots().empty();
    session_ = session;
    if (session_->mode == RegistryMode::kLocal)
    {
        session_->roots = LocalRoots(session_->view);
    }
    if (!shown)
    {
        ApplyRegistryRoots(session_->roots);
    }
    RefreshRegistryTabLabels();
    UpdateToolbarStates();
}

void MainWindow::Impl::ApplyRegistryRoots(const std::vector<RegistryRootEntry>& roots)
{
    browse_.roots() = roots;
    ResetHiveListCache();
    browse_.set_current_node(nullptr);
    browse_.values().Clear();
    current_key_count_ = 0;
    current_value_count_ = 0;
    browse_.tree().SetRegEditLayout(false);
    browse_.tree().SetRootLabel(TreeRootLabel(), TreeRootIcon());
    browse_.tree().PopulateRoots(browse_.roots());
    ResetNavigationState();
    UpdateStatus();

    SelectDefaultTreeItem();
}

std::vector<std::wstring> MainWindow::Impl::BuildVisibleTreePathParts(const std::wstring& path) const
{
    std::vector<std::wstring> parts = registry_path::Split(path);
    if (parts.empty())
    {
        return parts;
    }

    std::wstring root_label = TreeRootLabel();
    if (!root_label.empty() && !parts.empty() && EqualsInsensitive(parts.front(), root_label))
    {
        parts.erase(parts.begin());
    }
    if (!parts.empty() && EqualsInsensitive(parts.front(), L"Computer"))
    {
        parts.erase(parts.begin());
    }

    if (session_->mode == RegistryMode::kRemote && !session_->remote_machine.empty())
    {
        std::wstring machine = StripMachinePrefix(session_->remote_machine);
        if (!machine.empty() && !parts.empty() && EqualsInsensitive(parts.front(), machine))
        {
            parts.erase(parts.begin());
        }
    }
    if (session_->mode == RegistryMode::kOffline && !session_->offline_root_labels.empty() && parts.size() >= 2)
    {
        std::wstring root_name = session_->offline_root_name;
        auto is_offline_label = [&](const std::wstring& name) {
            for (const auto& label : session_->offline_root_labels)
            {
                if (EqualsInsensitive(label, name))
                {
                    return true;
                }
            }
            return false;
        };
        if (!root_name.empty() && EqualsInsensitive(parts[0], root_name) && is_offline_label(parts[1]))
        {
            parts.erase(parts.begin());
        }
    }

    if (!parts.empty())
    {
        if (!EqualsInsensitive(parts.front(), kRootKeysGroupLabel) &&
            !EqualsInsensitive(parts.front(), kRealGroupLabel))
        {
            if (EqualsInsensitive(parts.front(), L"REGISTRY"))
            {
                parts.insert(parts.begin(), kRealGroupLabel);
            }
            else
            {
                parts.insert(parts.begin(), kRootKeysGroupLabel);
            }
        }
    }
    return parts;
}

std::wstring MainWindow::Impl::TreeRootLabel() const
{
    if (session_->mode == RegistryMode::kRemote && !session_->remote_machine.empty())
    {
        return StripMachinePrefix(session_->remote_machine);
    }
    const std::wstring& name = win32ui::ComputerName();
    return name.empty() ? std::wstring(L"Computer") : name;
}

int MainWindow::Impl::TreeRootIcon() const
{
    return kLocalRegistryIconIndex + static_cast<int>(session_->mode);
}

void MainWindow::Impl::SelectDefaultTreeItem()
{
    if (!browse_.tree().hwnd())
    {
        return;
    }
    HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
    if (!root)
    {
        return;
    }
    HTREEITEM group = TreeView_GetChild(browse_.tree().hwnd(), root);
    HTREEITEM standard_group = nullptr;
    while (group)
    {
        wchar_t text[128] = {};
        TVITEMW tvi = {};
        tvi.mask = TVIF_TEXT;
        tvi.hItem = group;
        tvi.pszText = text;
        tvi.cchTextMax = static_cast<int>(_countof(text));
        if (TreeView_GetItem(browse_.tree().hwnd(), &tvi))
        {
            if (util::EqualsInsensitive(text, kRootKeysGroupLabel))
            {
                standard_group = group;
                break;
            }
        }
        group = TreeView_GetNextSibling(browse_.tree().hwnd(), group);
    }
    if (standard_group)
    {
        TreeView_SelectItem(browse_.tree().hwnd(), standard_group);
        return;
    }
    group = TreeView_GetChild(browse_.tree().hwnd(), root);
    while (group)
    {
        RegistryNode* node = browse_.tree().NodeFromItem(group);
        if (node)
        {
            TreeView_SelectItem(browse_.tree().hwnd(), group);
            return;
        }
        HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), group);
        if (child)
        {
            TreeView_SelectItem(browse_.tree().hwnd(), child);
            return;
        }
        group = TreeView_GetNextSibling(browse_.tree().hwnd(), group);
    }
}

// hidden trees retain state, tree_state also captures expansion for saving
void MainWindow::Impl::CaptureRegistryTabState(int index, bool tree_state)
{
    if (index < 0 || static_cast<size_t>(index) >= tabs_.size())
    {
        return;
    }
    TabEntry& entry = tabs_[static_cast<size_t>(index)];
    if (entry.kind == TabEntry::Kind::kSearch || !entry.tree)
    {
        return;
    }
    if (drop_tree_state_)
    {
        entry.selected_path.clear();
        entry.expanded_paths.clear();
    }
    else
    {
        CaptureTreeState(*entry.tree, &entry.selected_path, tree_state ? &entry.expanded_paths : nullptr);
    }
    if (entry.tree != &browse_.tree())
    {
        return;
    }
    entry.selected_values.clear();
    entry.value_top_index = 0;
    HWND list = browse_.values().hwnd();
    if (!list)
    {
        return;
    }
    entry.value_top_index = ListView_GetTopIndex(list);
    for (const int item : win32ui::SelectedListItems(list))
    {
        const ListRow* row = browse_.values().RowAt(item);
        if (row && row->kind == rowkind::kValue)
        {
            entry.selected_values.push_back(row->extra);
        }
    }
}

void MainWindow::Impl::ResetRegistryTreeState()
{
    if (!browse_.tree().hwnd())
    {
        return;
    }
    HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
    if (!root)
    {
        return;
    }

    SuspendTreeRedraw();
    std::function<void(HTREEITEM)> collapse = [&](HTREEITEM item) {
        while (item)
        {
            HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), item);
            if (child)
            {
                collapse(child);
            }
            // PopulateRoots leaves groups expanded, tab state stores key paths only
            if (browse_.tree().NodeFromItem(item))
            {
                TreeView_Expand(browse_.tree().hwnd(), item, TVE_COLLAPSE);
            }
            item = TreeView_GetNextSibling(browse_.tree().hwnd(), item);
        }
    };
    HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), root);
    if (child)
    {
        collapse(child);
    }
    TreeView_SelectItem(browse_.tree().hwnd(), root);
}

// tree repaints once after this message
void MainWindow::Impl::FlushTreeRedraw()
{
    if (tree_redraw_pending_)
    {
        tree_redraw_pending_ = false;
        browse_.tree().ResumeRedraw();
    }
}

void MainWindow::Impl::SuspendTreeRedraw()
{
    if (!tree_redraw_pending_ && PostMessageW(hwnd_, frame::message_id::kTreeRedraw, 0, 0))
    {
        tree_redraw_pending_ = true;
        browse_.tree().SuspendRedraw();
    }
}

void MainWindow::Impl::RestoreRegistryTabState(int index)
{
    if (index < 0 || static_cast<size_t>(index) >= tabs_.size() || !browse_.tree().hwnd())
    {
        return;
    }
    const TabEntry& entry = tabs_[static_cast<size_t>(index)];
    if (entry.kind == TabEntry::Kind::kSearch)
    {
        return;
    }
    ResetRegistryTreeState();
    if (entry.expanded_paths.empty() && entry.selected_path.empty())
    {
        SelectDefaultTreeItem();
        return;
    }
    ExpandTreePaths(entry.expanded_paths);
    if (!entry.selected_path.empty() && SelectTreePath(entry.selected_path))
    {
        RestoreValueSelection(entry);
        return;
    }
    SelectDefaultTreeItem();
}

void MainWindow::Impl::RestoreValueSelection(const TabEntry& entry)
{
    pending_value_selection_ = entry.selected_values;
    pending_value_top_index_ = entry.value_top_index;
    pending_value_selection_key_ =
        pending_value_selection_.empty() && pending_value_top_index_ == 0 ? std::wstring() : entry.selected_path;
}

bool MainWindow::Impl::ActivateTabTree(int index)
{
    TabEntry& entry = tabs_[static_cast<size_t>(index)];
    const bool fresh = !entry.tree;
    if (fresh)
    {
        // startup tree and cleared closed tab trees belong to no tab
        RegistryTree* shown = &browse_.tree();
        const bool owned = std::any_of(tabs_.begin(), tabs_.end(), [shown](const TabEntry& tab) { return tab.tree == shown; });
        entry.tree = owned ? browse_.AddTree() : shown;
        if (!entry.tree)
        {
            entry.tree = shown;
        }
        else if (owned)
        {
            ConfigureTree(*entry.tree);
        }
        entry.tree->synced_revision = RegistryStore::KeyRevision();
    }
    RegistryTree& previous = browse_.tree();
    if (entry.tree == &previous)
    {
        return fresh;
    }
    FlushTreeRedraw();
    RECT rect = {};
    win32ui::ChildRect(hwnd_, previous.hwnd(), &rect);
    const bool focused = GetFocus() == previous.hwnd();
    browse_.SetActiveTree(entry.tree);
    SetWindowPos(entry.tree->hwnd(), nullptr, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, SWP_NOZORDER | SWP_NOACTIVATE);
    // use style bit, IsWindowVisible is false before the main window is shown
    ShowWindow(entry.tree->hwnd(), (GetWindowLongPtrW(previous.hwnd(), GWL_STYLE) & WS_VISIBLE) ? SW_SHOW : SW_HIDE);
    ShowWindow(previous.hwnd(), SW_HIDE);
    if (focused)
    {
        SetFocus(entry.tree->hwnd());
    }
    ResetNavigationState();
    return fresh;
}

// existing tab trees only resync keys changed elsewhere
void MainWindow::Impl::ResumeTabTree(int index)
{
    const TabEntry& entry = tabs_[static_cast<size_t>(index)];
    if (entry.session)
    {
        session_ = entry.session;
    }
    const auto same_root = [](const RegistryRootEntry& left, const RegistryRootEntry& right) {
        return left.root == right.root && left.view == right.view && left.display_name == right.display_name;
    };
    if (entry.session && !std::equal(browse_.roots().begin(), browse_.roots().end(), session_->roots.begin(), session_->roots.end(), same_root))
    {
        ReapplyRoots();
    }
    else if (browse_.tree().synced_revision != RegistryStore::KeyRevision())
    {
        RefreshWholeTree();
    }
    RestoreValueSelection(entry);
    ApplyTreeSelectionEffects(browse_.current_node());
    RefreshRegistryTabLabels();
    UpdateToolbarStates();
}

std::wstring MainWindow::Impl::LocalRegistryTabLabel(int index) const
{
    if (index < 0 || static_cast<size_t>(index) >= tabs_.size())
    {
        return util::Tr(L"Local Registry");
    }
    const REGSAM view = tabs_[static_cast<size_t>(index)].registry_view;
    const std::wstring base = view ? util::Tr(L"Local Registry (32-bit)") : util::Tr(L"Local Registry");
    int local_count = 0;
    int local_index = 0;
    for (size_t i = 0; i < tabs_.size(); ++i)
    {
        const TabEntry& entry = tabs_[i];
        if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kLocal || entry.registry_view != view)
        {
            continue;
        }
        ++local_count;
        if (static_cast<int>(i) == index)
        {
            local_index = local_count;
        }
    }
    if (local_count <= 1 || local_index <= 1)
    {
        return base;
    }
    return base + L" (" + std::to_wstring(local_index) + L")";
}

void MainWindow::Impl::RefreshRegistryTabLabels()
{
    if (!tab_)
    {
        return;
    }
    for (size_t i = 0; i < tabs_.size(); ++i)
    {
        const TabEntry& entry = tabs_[i];
        if (entry.kind != TabEntry::Kind::kRegistry)
        {
            continue;
        }
        std::wstring label;
        if (entry.registry_mode == RegistryMode::kLocal)
        {
            label = LocalRegistryTabLabel(static_cast<int>(i));
        }
        else
        {
            continue;
        }
        TCITEMW item = {};
        item.mask = TCIF_TEXT;
        item.pszText = const_cast<wchar_t*>(label.c_str());
        TabCtrl_SetItem(tab_, static_cast<int>(i), &item);
    }
    UpdateTabWidth();
    InvalidateRect(tab_, nullptr, FALSE);
}

void MainWindow::Impl::AppendRealRegistryRoot(std::vector<RegistryRootEntry>* roots)
{
    if (!roots)
    {
        return;
    }
    if (!registry_root_.get())
    {
        registry_root_ = util::OpenNativeRegistryRoot();
    }
    if (!registry_root_.get())
    {
        return;
    }
    RegistryRootEntry entry;
    entry.root = registry_root_.get();
    entry.display_name = L"REGISTRY";
    entry.path_name = L"REGISTRY";
    entry.subkey_prefix = L"";
    entry.group = RegistryRootGroup::kReal;
    roots->push_back(std::move(entry));
}

std::vector<RegistryRootEntry> MainWindow::Impl::LocalRoots(REGSAM view)
{
    std::vector<RegistryRootEntry> roots = RegistryStore::DefaultRoots(settings_.show_extra_hives);
    for (RegistryRootEntry& root : roots)
    {
        root.view = view;
    }
    // native paths have no 32 bit view
    if (!view && !settings_.simple_mode)
    {
        AppendRealRegistryRoot(&roots);
    }
    return roots;
}

void MainWindow::Impl::ReloadLocalRoots()
{
    local_session_->roots = LocalRoots(local_session_->view);
    for (const TabEntry& entry : tabs_)
    {
        if (entry.kind == TabEntry::Kind::kRegistry && entry.session && entry.session->mode == RegistryMode::kLocal)
        {
            entry.session->roots = LocalRoots(entry.session->view);
        }
    }
    RegistryStore::NoteKeyChange();
    const int index = tab_ ? TabCtrl_GetCurSel(tab_) : -1;
    if (session_->mode == RegistryMode::kLocal && (index < 0 || (static_cast<size_t>(index) < tabs_.size() && tabs_[static_cast<size_t>(index)].kind == TabEntry::Kind::kRegistry)))
    {
        ReapplyRoots();
    }
}

void MainWindow::Impl::ReapplyRoots()
{
    std::wstring selected;
    std::vector<std::wstring> expanded;
    CaptureTreeState(browse_.tree(), &selected, &expanded);
    ApplyRegistryRoots(session_->roots);
    ExpandTreePaths(expanded);
    if (!selected.empty())
    {
        SelectTreePath(selected);
    }
    browse_.tree().synced_revision = RegistryStore::KeyRevision();
}

std::shared_ptr<MainWindow::Impl::RegistrySession> MainWindow::Impl::LocalViewSession(REGSAM view) const
{
    auto session = std::make_shared<RegistrySession>();
    session->view = view;
    return session;
}

void MainWindow::Impl::GoToOtherView()
{
    const RegistryNode* node = browse_.current_node();
    if (!node || session_->mode != RegistryMode::kLocal || !win32::HasAlternateView())
    {
        return;
    }
    const std::wstring path = registry_path::Build(*node);
    const REGSAM view = session_->view ? 0 : win32::kAlternateRegistryView;
    const auto found = std::find_if(tabs_.begin(), tabs_.end(), [view](const TabEntry& entry) {
        return entry.kind == TabEntry::Kind::kRegistry && entry.registry_mode == RegistryMode::kLocal && entry.registry_view == view;
    });
    if (found == tabs_.end())
    {
        OpenLocalRegistryTab(view);
    }
    else
    {
        const int index = static_cast<int>(found - tabs_.begin());
        suppress_tab_change_ = true;
        SelectTabIndex(index);
        suppress_tab_change_ = false;
        ApplyTabSelection(index);
    }
    BeginJumpUiBatch();
    if (SelectTreePath(path))
    {
        ApplyTreeSelectionEffects(browse_.current_node());
    }
    EndJumpUiBatch();
}

std::wstring MainWindow::Impl::VirtualStoreTarget() const
{
    const RegistryNode* node = browse_.current_node();
    KeyDetails details;
    if (!node || session_->mode != RegistryMode::kLocal || !RegistryStore::QueryKeyDetails(*node, &details))
    {
        return {};
    }
    const std::wstring store = registry_path::VirtualStorePath(details.native.native_name);
    RegistryNode store_node;
    KeyInfo info;
    return !store.empty() && registry_path::ParseRoot(store, &store_node) && RegistryStore::QueryKeyInfo(store_node, &info) ? store : std::wstring();
}

std::wstring MainWindow::Impl::GlobalKeyTarget() const
{
    const RegistryNode* node = browse_.current_node();
    if (!node || session_->mode != RegistryMode::kLocal)
    {
        return {};
    }
    const std::wstring global = registry_path::GlobalKeyPath(registry_path::Build(*node));
    RegistryNode global_node;
    KeyInfo info;
    return !global.empty() && registry_path::ParseRoot(global, &global_node) && RegistryStore::QueryKeyInfo(global_node, &info) ? global : std::wstring();
}

bool MainWindow::Impl::SwitchToLocalRegistry()
{
    RegistrySession* current = CurrentTabSession();
    if (current && !ConfirmOfflineChanges(*current, util::Tr(L"The offline registry has unsaved changes.\n"
                                                             L"Save before switching?")))
    {
        return false;
    }
    if (current)
    {
        OfferRemoteServiceRestore(*current);
    }
    UpdateRegistryTabEntry(RegistryMode::kLocal, L"", L"");
    ShowSession(local_session_);
    return true;
}
bool MainWindow::Impl::SwitchToRemoteRegistry()
{
    std::wstring machine;
    const DWORD picked = win32ui::ChooseComputer(hwnd_, &machine);
    if (picked == ERROR_CANCELLED)
    {
        return false;
    }
    if (picked != ERROR_SUCCESS)
    {
        machine = session_->remote_machine;
        if (win32ui::PromptText(hwnd_, {.title = util::Tr(L"Connect to Remote Registry"), .label = util::Tr(L"Computer name (e.g. \\\\MACHINE):")}, &machine) != 0)
        {
            return false;
        }
    }
    return ConnectRemoteRegistry(machine, true);
}

bool MainWindow::Impl::ConnectRemoteRegistry(const std::wstring& name, bool open_new_tab)
{
    const std::wstring machine = NormalizeMachineName(name);
    if (machine.empty())
    {
        ui::ShowError(hwnd_, util::Tr(L"Computer name is required."));
        return false;
    }
    auto session = std::make_shared<RegistrySession>();
    session->mode = RegistryMode::kRemote;
    session->remote_machine = machine;
    LONG result = RegConnectRegistryW(machine.c_str(), HKEY_LOCAL_MACHINE, &session->remote_hklm);
    if (result != ERROR_SUCCESS && OfferRemoteServiceStart(machine, session.get()))
    {
        result = RegConnectRegistryW(machine.c_str(), HKEY_LOCAL_MACHINE, &session->remote_hklm);
    }
    if (result != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, FormatWin32Error(result));
        OfferRemoteServiceRestore(*session);
        return false;
    }
    const LONG hku_result = RegConnectRegistryW(machine.c_str(), HKEY_USERS, &session->remote_hku);
    RegistrySession* current = CurrentTabSession();
    if (!open_new_tab && current && !ConfirmOfflineChanges(*current, util::Tr(L"The offline registry has unsaved changes.\n"
                                                                              L"Save before switching?")))
    {
        OfferRemoteServiceRestore(*session);
        return false;
    }
    if (!open_new_tab && current)
    {
        OfferRemoteServiceRestore(*current);
    }
    const std::wstring prefix = machine + L"\\";
    session->roots.push_back({session->remote_hklm, L"HKEY_LOCAL_MACHINE", prefix + L"HKEY_LOCAL_MACHINE", L""});
    if (session->remote_hku)
    {
        session->roots.push_back({session->remote_hku, L"HKEY_USERS", prefix + L"HKEY_USERS", L""});
    }
    if (tab_ && open_new_tab)
    {
        AddRegistryTab(RegistryMode::kRemote, util::Tr(L"Remote Registry"));
    }
    UpdateRegistryTabEntry(RegistryMode::kRemote, L"", machine);
    UpdateTabText(util::Tr(L"Remote Registry") + std::wstring(L" (") + StripMachinePrefix(machine) + L")");
    ShowSession(session);
    if (hku_result != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, util::TrDetail(L"Connected to HKEY_LOCAL_MACHINE, but HKEY_USERS was unavailable.", FormatWin32Error(hku_result)));
    }
    return true;
}

bool MainWindow::Impl::OfferRemoteServiceStart(const std::wstring& machine, RegistrySession* session)
{
    util::ServiceState service;
    if (util::QueryRemoteRegistryService(machine, &service) != ERROR_SUCCESS || service.state != SERVICE_STOPPED)
    {
        return false;
    }
    const bool disabled = service.start_type == SERVICE_DISABLED;
    const wchar_t* message = disabled ? util::Tr(L"The Remote Registry service on this computer is disabled. Set it to start manually and start it?")
                                      : util::Tr(L"The Remote Registry service on this computer is stopped. Start it?");
    if (ui::PromptKeyChoice(hwnd_, message, machine, util::Tr(L"Remote Registry"), util::Tr(L"Start"), L"", util::Tr(L"Cancel")) != IDYES)
    {
        return false;
    }
    HCURSOR previous = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    const LONG started = util::StartRemoteRegistryService(machine, disabled);
    SetCursor(previous);
    if (started != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, util::TrDetail(L"The Remote Registry service couldn't be started.", FormatWin32Error(started)));
        return false;
    }
    session->remote_service_started = true;
    session->remote_service_start_type = service.start_type;
    return true;
}

void MainWindow::Impl::OfferRemoteServiceRestore(RegistrySession& session, bool hand_over)
{
    if (session.mode != RegistryMode::kRemote || !session.remote_service_started)
    {
        return;
    }
    session.remote_service_started = false;
    // another tab still needs the remote service and will offer restoration when closed
    for (TabEntry& entry : tabs_)
    {
        if (hand_over && entry.session && entry.session.get() != &session && entry.session->mode == RegistryMode::kRemote &&
            util::EqualsInsensitive(entry.session->remote_machine, session.remote_machine))
        {
            entry.session->remote_service_started = true;
            entry.session->remote_service_start_type = session.remote_service_start_type;
            return;
        }
    }
    const wchar_t* message = session.remote_service_start_type == SERVICE_DISABLED
                                 ? util::Tr(L"RegKit started the Remote Registry service on this computer. Stop it and disable it again?")
                                 : util::Tr(L"RegKit started the Remote Registry service on this computer. Stop it again?");
    if (ui::PromptKeyChoice(hwnd_, message, session.remote_machine, util::Tr(L"Remote Registry"), util::Tr(L"Stop"), L"", util::Tr(L"Keep Running")) != IDYES)
    {
        return;
    }
    const LONG stopped = util::StopRemoteRegistryService(session.remote_machine, session.remote_service_start_type);
    if (stopped != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, util::TrDetail(L"The Remote Registry service couldn't be stopped.", FormatWin32Error(stopped)));
    }
}
bool MainWindow::Impl::SwitchToOfflineRegistry()
{
    const int choice =
        ui::PromptChoice(hwnd_, util::Tr(L"Load the offline registry from a single hive file, or from a folder of hives?"), util::Tr(L"Offline Registry"), util::Tr(L"Hive File"), util::Tr(L"Folder"), util::Tr(L"Cancel"), {90, 70, 70}, 470);
    std::wstring hive_path;
    DWORD error = ERROR_SUCCESS;
    if (choice == IDYES)
    {
        error = win32ui::ChooseFileToOpen(hwnd_, {.filter = kOfflineHiveFilter}, &hive_path);
    }
    else if (choice == IDNO)
    {
        error = win32ui::ChooseFolder(hwnd_, &hive_path);
    }
    else
    {
        return false;
    }
    if (!ui::ReportFileDialogResult(hwnd_, error))
    {
        return false;
    }
    return LoadOfflineRegistryFromPath(hive_path, true);
}

bool MainWindow::Impl::LoadOfflineRegistryFromPath(const std::wstring& path, bool open_new_tab)
{
    RegistrySession* current = CurrentTabSession();
    if (!open_new_tab && current && !ConfirmOfflineChanges(*current, util::Tr(L"The offline registry has unsaved changes.\n"
                                                                              L"Save before switching?")))
    {
        return false;
    }
    if (!open_new_tab && current)
    {
        OfferRemoteServiceRestore(*current);
    }

    std::wstring selection_path = util::TrimTrailingSeparators(path);
    if (selection_path.empty())
    {
        return false;
    }

    bool is_dir = util::IsDirectory(selection_path.c_str());
    std::vector<OfflineHiveCandidate> candidates;
    if (is_dir)
    {
        CollectOfflineHivesInFolder(selection_path, &candidates);
        if (candidates.empty())
        {
            ui::ShowError(hwnd_, util::Tr(L"The selected folder doesn't contain a registry hive file."));
            return false;
        }
    }
    else
    {
        std::wstring mount_name = TrimWhitespace(util::FileBaseName(selection_path));
        if (mount_name.empty())
        {
            mount_name = L"OfflineHive";
        }
        candidates.push_back({selection_path, mount_name});
    }

    auto session = std::make_shared<RegistrySession>();
    session->mode = RegistryMode::kOffline;
    session->offline_root_name = ResolveOfflineRootName(selection_path, is_dir, browse_.current_node());

    std::wstring error;
    for (const auto& candidate : candidates)
    {
        HKEY hive = nullptr;
        if (!RegistryStore::OpenOfflineHive(candidate.path, &hive, &error))
        {
            if (!error.empty())
            {
                ui::ShowError(hwnd_, error);
            }
            return false;
        }
        RegistryStore::AddOfflineRoot(hive);
        std::wstring label = TrimWhitespace(candidate.label);
        if (label.empty())
        {
            label = TrimWhitespace(util::FileBaseName(candidate.path));
        }
        if (label.empty())
        {
            label = L"OfflineHive";
        }
        session->offline_roots.push_back(hive);
        session->offline_root_labels.push_back(label);
        session->offline_root_paths.push_back(candidate.path);
        session->roots.push_back({hive, label, session->offline_root_name + L"\\" + label, L""});
    }
    if (session->offline_roots.size() == 1)
    {
        session->offline_root = session->offline_roots.front();
        session->offline_mount = session->offline_root_labels.front();
    }

    if (tab_ && open_new_tab)
    {
        AddRegistryTab(RegistryMode::kOffline, util::Tr(L"Offline Registry"));
    }

    std::wstring tab_text = util::Tr(L"Offline Registry");
    tab_text += L" (" + session->offline_root_name + (session->offline_mount.empty() ? L"" : L"\\" + session->offline_mount) + L")";
    UpdateTabText(tab_text);
    UpdateRegistryTabEntry(RegistryMode::kOffline, selection_path, L"");
    ShowSession(session);
    HistoryEntry history;
    history.action = L"Load offline registry";
    history.new_data = selection_path;
    AppendHistoryEntry(std::move(history));
    return true;
}

bool MainWindow::Impl::SaveOfflineRegistry(RegistrySession& session, bool choose_path)
{
    if (session.mode != RegistryMode::kOffline || session.offline_roots.empty())
    {
        ui::ShowError(hwnd_, util::Tr(L"No offline registry is loaded."));
        return false;
    }
    if (session.offline_roots.size() > 1)
    {
        if (session.offline_root_paths.size() != session.offline_roots.size())
        {
            ui::ShowError(hwnd_, util::Tr(L"Failed to resolve offline hive paths for saving."));
            return false;
        }
        for (size_t i = 0; i < session.offline_roots.size(); ++i)
        {
            const std::wstring& path = session.offline_root_paths[i];
            if (path.empty())
            {
                ui::ShowError(hwnd_, util::Tr(L"Failed to resolve offline hive path for saving."));
                return false;
            }
            std::wstring error;
            if (!SaveHiveAtomically(session.offline_roots[i], path, &error))
            {
                ui::ShowError(hwnd_, error.empty() ? util::Tr(L"Failed to save offline hive.") : error);
                return false;
            }
        }
        session.offline_dirty = false;
        HistoryEntry history;
        history.action = L"Save offline registry";
        history.new_data = std::to_wstring(session.offline_roots.size()) + L" hives";
        AppendHistoryEntry(std::move(history));
        return true;
    }
    if (!session.offline_root)
    {
        ui::ShowError(hwnd_, util::Tr(L"No offline registry is loaded."));
        return false;
    }

    std::wstring path = session.offline_root_paths.empty() ? std::wstring() : session.offline_root_paths.front();
    if ((choose_path || path.empty()) &&
        !ui::ReportFileDialogResult(hwnd_, win32::ChooseFileToSave(hwnd_, ui::kHiveFileFilter, path.empty() ? nullptr : path.c_str(), &path)))
    {
        return false;
    }

    std::wstring error;
    if (!SaveHiveAtomically(session.offline_root, path, &error))
    {
        ui::ShowError(hwnd_, error.empty() ? util::Tr(L"Failed to save offline hive.") : error);
        return false;
    }
    if (!session.offline_root_paths.empty())
    {
        session.offline_root_paths.front() = path;
    }
    // tab reopens the saved destination file
    for (TabEntry& tab : tabs_)
    {
        if (tab.session.get() == &session)
        {
            tab.offline_path = path;
        }
    }
    session.offline_dirty = false;
    HistoryEntry history;
    history.action = L"Save offline registry";
    history.new_data = path;
    AppendHistoryEntry(std::move(history));
    return true;
}

void MainWindow::Impl::NavigateToAddress()
{
    std::wstring path;
    std::wstring value_name;
    bool value_missing = false;
    if (ResolveJumpTarget(util::WindowText(browse_.address()), &path, &value_name, &value_missing))
    {
        if (!SelectTreePath(path))
        {
            return;
        }
        UpdateAddressBar(browse_.current_node());
        if (value_name.empty())
        {
            return;
        }
        if (value_missing)
        {
            PromptMissingValue(value_name);
        }
        else
        {
            SelectValueWhenReady(value_name);
        }
        return;
    }
    if (!path.empty())
    {
        PromptMissingPath(path);
    }
}

void MainWindow::Impl::PromptMissingValue(const std::wstring& value_name)
{
    ui::PromptKeyChoice(hwnd_, util::Tr(L"The key was opened, but it doesn't contain this value:"), registry_path::DisplayName(value_name), util::Tr(L"Value Not Found"), util::Tr(L"OK"), L"", L"");
}

void MainWindow::Impl::PromptMissingPath(const std::wstring& path)
{
    std::wstring nearest;
    if (!FindNearestExistingPath(path, &nearest) || nearest.empty())
    {
        ui::ShowWarning(hwnd_, util::TrDetail(L"Registry path not found.", path));
        return;
    }
    std::wstring message = util::Tr(L"The registry key doesn't exist:");
    if (settings_.read_only)
    {
        message += L"\nRead only mode is enabled.";
        int result = ui::PromptKeyChoice(hwnd_, message, path, util::Tr(L"Registry Path Not Found"), util::Tr(L"Go to Nearest Key"), L"", util::Tr(L"Cancel"), {150, 70, 70});
        if (result == IDYES)
        {
            SelectTreePath(nearest);
        }
        return;
    }
    int result = ui::PromptKeyChoice(hwnd_, message, path, util::Tr(L"Registry Path Not Found"), util::Tr(L"Go to Nearest Key"), util::Tr(L"Create Key"), util::Tr(L"Cancel"), {150, 100, 70});
    if (result == IDYES)
    {
        SelectTreePath(nearest);
        return;
    }
    if (result == IDNO)
    {
        if (!CreateRegistryPath(path))
        {
            ui::ShowError(hwnd_, util::Tr(L"Failed to create registry key."));
            return;
        }
        RefreshTreePath(nearest);
        SelectTreePath(path);
    }
}

void MainWindow::Impl::ApplyQueuedExternalJump()
{
    if (queued_external_jump_target_.empty())
    {
        return;
    }
    std::wstring target = std::move(queued_external_jump_target_);
    queued_external_jump_target_.clear();
    if (!NavigateToExternalJump(target))
    {
        ui::ShowWarning(hwnd_, util::TrDetail(L"Registry path not found.", target));
    }
}

bool MainWindow::Impl::ResolveJumpTarget(const std::wstring& target, std::wstring* key_path, std::wstring* value_name, bool* value_missing) const
{
    return registry_path::ResolveJumpTarget(
        target,
        [this](const std::wstring& path) { return NormalizeRegistryPath(path); },
        [this](const std::wstring& path, RegistryNode* node) { return KeyPathExists(path, node); },
        key_path,
        value_name,
        value_missing
    );
}

bool MainWindow::Impl::ActivateLocalRegistryTab()
{
    if (!IsLocalRegistryTabIndex(TabCtrl_GetCurSel(tab_)))
    {
        const int local_tab = FindLocalRegistryTabIndex();
        if (local_tab < 0)
        {
            OpenLocalRegistryTab();
        }
        else
        {
            suppress_tab_change_ = true;
            SelectTabIndex(local_tab);
            suppress_tab_change_ = false;
            ApplyTabSelection(local_tab);
        }
    }
    return session_ == local_session_ || SwitchToLocalRegistry();
}

void MainWindow::Impl::QueueCompatJump(const RegistryNode& node)
{
    pending_compat_jump_ = registry_path::Build(node);
    SetTimer(hwnd_, kCompatJumpTimerId, kCompatJumpDelayMs, nullptr);
}

void MainWindow::Impl::FlushExternalNavigation()
{
    if (flushing_external_navigation_ || updating_value_list_)
    {
        return;
    }
    flushing_external_navigation_ = true;
    if (!pending_compat_jump_.empty())
    {
        KillTimer(hwnd_, kCompatJumpTimerId);
        const std::wstring target = std::move(pending_compat_jump_);
        pending_compat_jump_.clear();
        NavigateToExternalJump(target);
    }
    const ULONGLONG deadline = GetTickCount64() + 2000;
    MSG message = {};
    while (value_list_loading_ && GetTickCount64() < deadline)
    {
        if (PeekMessageW(&message, hwnd_, frame::message_id::kValueListReady, frame::message_id::kValueListReady, PM_REMOVE))
        {
            HandleValueWorkerMessage(message.message, message.wParam, message.lParam);
        }
        else
        {
            MsgWaitForMultipleObjects(0, nullptr, FALSE, 10, QS_POSTMESSAGE);
        }
    }
    flushing_external_navigation_ = false;
}

bool MainWindow::Impl::NavigateToExternalJump(const std::wstring& target)
{
    if (!ActivateLocalRegistryTab())
    {
        return false;
    }
    std::wstring key_path;
    std::wstring value_name;
    bool value_missing = false;
    if (!ResolveJumpTarget(target, &key_path, &value_name, &value_missing))
    {
        if (!key_path.empty())
        {
            PromptMissingPath(key_path);
        }
        return !key_path.empty();
    }
    if (!NavigateToResolvedExternalJump(key_path, value_missing ? std::wstring() : value_name))
    {
        return false;
    }
    if (value_missing)
    {
        PromptMissingValue(value_name);
    }
    return true;
}

bool MainWindow::Impl::SearchResultOpensInNewTab() const
{
    if (!tab_)
    {
        return false;
    }
    const int index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
    return index >= 0 && static_cast<size_t>(index) < search_tabs_.size() &&
           search_tabs_[static_cast<size_t>(index)].open_in_new_tab;
}

bool MainWindow::Impl::NavigateToResolvedExternalJump(const std::wstring& key_path, const std::wstring& value_name)
{
    if (key_path.empty() || !ActivateLocalRegistryTab())
    {
        return false;
    }

    ApplyViewVisibility();
    UpdateStatus();

    BeginJumpUiBatch();
    if (!SelectTreePath(key_path))
    {
        EndJumpUiBatch();
        return false;
    }
    ApplyTreeSelectionEffects(browse_.current_node());
    EndJumpUiBatch();

    pending_external_value_generation_ = 0;
    pending_external_value_name_.clear();
    pending_value_command_ = 0;
    if (!value_name.empty())
    {
        pending_external_value_name_ = value_name;
        if (!SelectValueByName(value_name) && browse_.current_node() && !value_list_loading_)
        {
            UpdateValueListForNode(browse_.current_node());
        }
        pending_external_value_generation_ = value_list_generation_.load();
    }
    return true;
}

}
