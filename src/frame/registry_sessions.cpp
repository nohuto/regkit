// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

#include "win32/file_dialog.h"
#include "win32/text_transform.h"

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

} // namespace

void MainWindow::Impl::ReleaseRemoteRegistry()
{
    if (remote_hklm_)
    {
        RegCloseKey(remote_hklm_);
        remote_hklm_ = nullptr;
    }
    if (remote_hku_)
    {
        RegCloseKey(remote_hku_);
        remote_hku_ = nullptr;
    }
    remote_machine_.clear();
}

bool MainWindow::Impl::UnloadOfflineRegistry(std::wstring* error)
{
    if (error)
    {
        error->clear();
    }
    if (offline_roots_.empty())
    {
        return true;
    }
    std::vector<HKEY> remaining_roots;
    std::vector<std::wstring> remaining_labels;
    std::vector<std::wstring> remaining_paths;
    for (size_t i = 0; i < offline_roots_.size(); ++i)
    {
        std::wstring close_error;
        if (RegistryStore::CloseOfflineHive(offline_roots_[i], &close_error))
        {
            continue;
        }
        remaining_roots.push_back(offline_roots_[i]);
        if (i < offline_root_labels_.size())
        {
            remaining_labels.push_back(offline_root_labels_[i]);
        }
        if (i < offline_root_paths_.size())
        {
            remaining_paths.push_back(offline_root_paths_[i]);
        }
        if (error && error->empty())
        {
            *error = close_error;
        }
    }
    if (!remaining_roots.empty())
    {
        offline_roots_ = std::move(remaining_roots);
        offline_root_labels_ = std::move(remaining_labels);
        offline_root_paths_ = std::move(remaining_paths);
        offline_root_ = offline_roots_.size() == 1 ? offline_roots_.front() : nullptr;
        offline_mount_ =
            offline_roots_.size() == 1 && !offline_root_labels_.empty() ? offline_root_labels_.front() : std::wstring();
        RegistryStore::SetOfflineRoots(offline_roots_);
        std::vector<RegistryRootEntry> roots;
        roots.reserve(offline_roots_.size());
        for (size_t i = 0; i < offline_roots_.size(); ++i)
        {
            const std::wstring label =
                i < offline_root_labels_.size() ? offline_root_labels_[i] : std::wstring(L"OfflineHive");
            roots.push_back({offline_roots_[i], label, offline_root_name_ + L"\\" + label, L""});
        }
        ApplyRegistryRoots(roots);
        return false;
    }
    ClearOfflineDirty();
    RegistryStore::SetOfflineRoots({});
    offline_roots_.clear();
    offline_root_labels_.clear();
    offline_root_paths_.clear();
    offline_root_ = nullptr;
    offline_mount_.clear();
    offline_root_name_.clear();
    return true;
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

    auto is_standard_root = [](const std::wstring& name) -> bool {
        if (StartsWithInsensitive(name, L"HKEY_"))
        {
            return true;
        }
        return EqualsInsensitive(name, L"HKLM") || EqualsInsensitive(name, L"HKCU") ||
               EqualsInsensitive(name, L"HKCR") || EqualsInsensitive(name, L"HKU") || EqualsInsensitive(name, L"HKCC");
    };
    if (!parts.empty() && EqualsInsensitive(parts.front(), L"Registry"))
    {
        parts.front() = (parts.size() > 1 && is_standard_root(parts[1])) ? kRootKeysGroupLabel : kRealGroupLabel;
    }
    else if (!parts.empty() && EqualsInsensitive(parts.front(), L"Real Registry"))
    {
        parts.front() = kRealGroupLabel;
        if (parts.size() > 1 && EqualsInsensitive(parts[1], kRealGroupLabel))
        {
            parts.erase(parts.begin() + 1);
        }
    }
    else if (!parts.empty() && EqualsInsensitive(parts.front(), L"Standard Hives"))
    {
        parts.front() = kRootKeysGroupLabel;
    }

    if (registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty())
    {
        std::wstring machine = StripMachinePrefix(remote_machine_);
        if (!machine.empty() && !parts.empty() && EqualsInsensitive(parts.front(), machine))
        {
            parts.erase(parts.begin());
        }
    }
    if (registry_mode_ == RegistryMode::kOffline && !offline_root_labels_.empty() && parts.size() >= 2)
    {
        std::wstring root_name = offline_root_name_;
        auto is_offline_label = [&](const std::wstring& name) {
            for (const auto& label : offline_root_labels_)
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
    if (registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty())
    {
        return StripMachinePrefix(remote_machine_);
    }
    wchar_t buffer[MAX_COMPUTERNAME_LENGTH + 1] = {};
    DWORD size = static_cast<DWORD>(_countof(buffer));
    if (GetComputerNameW(buffer, &size) && size > 0)
    {
        return std::wstring(buffer, size);
    }
    return L"Computer";
}

int MainWindow::Impl::TreeRootIcon() const
{
    return kLocalRegistryIconIndex + static_cast<int>(registry_mode_);
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

void MainWindow::Impl::CaptureRegistryTabState(int index)
{
    if (!browse_.tree().hwnd() || index < 0 || static_cast<size_t>(index) >= tabs_.size())
    {
        return;
    }
    TabEntry& entry = tabs_[static_cast<size_t>(index)];
    if (entry.kind == TabEntry::Kind::kSearch)
    {
        return;
    }
    CaptureTreeState(&entry.selected_path, &entry.expanded_paths);
    entry.selected_value.clear();
    entry.selected_values.clear();
    entry.value_top_index = 0;
    HWND list = browse_.values().hwnd();
    if (!list)
    {
        return;
    }
    entry.value_top_index = ListView_GetTopIndex(list);
    int item = ListView_GetNextItem(list, -1, LVNI_SELECTED);
    while (item >= 0)
    {
        const ListRow* row = browse_.values().RowAt(item);
        if (row && row->kind == rowkind::kValue)
        {
            if (entry.selected_value.empty())
            {
                entry.selected_value = row->extra;
            }
            entry.selected_values.push_back(row->extra);
        }
        item = ListView_GetNextItem(list, item, LVNI_SELECTED);
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

    SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, FALSE, 0);
    std::function<void(HTREEITEM)> collapse = [&](HTREEITEM item) {
        while (item)
        {
            HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), item);
            if (child)
            {
                collapse(child);
            }
            TreeView_Expand(browse_.tree().hwnd(), item, TVE_COLLAPSE);
            item = TreeView_GetNextSibling(browse_.tree().hwnd(), item);
        }
    };
    HTREEITEM child = TreeView_GetChild(browse_.tree().hwnd(), root);
    if (child)
    {
        collapse(child);
    }
    TreeView_SelectItem(browse_.tree().hwnd(), root);
    SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, TRUE, 0);
    RedrawWindow(browse_.tree().hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
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
        pending_value_selection_ = entry.selected_values;
        if (pending_value_selection_.empty() && !entry.selected_value.empty())
        {
            pending_value_selection_.push_back(entry.selected_value);
        }
        pending_value_top_index_ = entry.value_top_index;
        pending_value_selection_key_ =
            pending_value_selection_.empty() && pending_value_top_index_ == 0 ? std::wstring() : entry.selected_path;
        return;
    }
    SelectDefaultTreeItem();
}

std::wstring MainWindow::Impl::LocalRegistryTabLabel(int index) const
{
    if (index < 0 || static_cast<size_t>(index) >= tabs_.size())
    {
        return L"Local Registry";
    }
    int local_count = 0;
    int local_index = 0;
    for (size_t i = 0; i < tabs_.size(); ++i)
    {
        const TabEntry& entry = tabs_[i];
        if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kLocal)
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
        return L"Local Registry";
    }
    return L"Local Registry (" + std::to_wstring(local_index) + L")";
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

bool MainWindow::Impl::SwitchToLocalRegistry()
{
    bool needs_reload = registry_mode_ != RegistryMode::kLocal;
    if (!needs_reload)
    {
        if (browse_.roots().empty())
        {
            needs_reload = true;
        }
        else if (RegistryStore::IsVirtualRoot(browse_.roots().front().root))
        {
            needs_reload = true;
        }
        else
        {
            auto has_root = [&](HKEY root) -> bool {
                for (const auto& entry_root : browse_.roots())
                {
                    if (entry_root.root == root)
                    {
                        return true;
                    }
                }
                return false;
            };
            if (!has_root(HKEY_CLASSES_ROOT) || !has_root(HKEY_CURRENT_USER) || !has_root(HKEY_LOCAL_MACHINE) ||
                !has_root(HKEY_USERS) || !has_root(HKEY_CURRENT_CONFIG))
            {
                needs_reload = true;
            }
        }
    }
    if (!needs_reload)
    {
        return true;
    }
    if (!ConfirmOfflineChanges(L"The offline registry has unsaved changes.\n"
                               L"Save before switching?"))
    {
        return false;
    }
    if (registry_mode_ == RegistryMode::kOffline)
    {
        std::wstring error;
        if (!UnloadOfflineRegistry(&error))
        {
            if (!error.empty())
            {
                ui::ShowError(hwnd_, error);
            }
            return false;
        }
    }
    ReleaseRemoteRegistry();
    registry_mode_ = RegistryMode::kLocal;
    UpdateRegistryTabEntry(RegistryMode::kLocal, L"", L"");
    std::vector<RegistryRootEntry> roots = RegistryStore::DefaultRoots(show_extra_hives_);
    AppendRealRegistryRoot(&roots);
    ApplyRegistryRoots(roots);
    RefreshRegistryTabLabels();
    return true;
}

bool MainWindow::Impl::SwitchToRemoteRegistry()
{
    std::wstring machine;
    const HRESULT picked = win32::ChooseComputer(hwnd_, &machine);
    if (win32::DialogCancelled(picked))
    {
        return false;
    }
    if (FAILED(picked))
    {
        editors::TextRequest request;
        request.title = L"Connect to Remote Registry";
        request.label = L"Computer name (e.g. \\\\MACHINE):";
        request.text = remote_machine_;
        editors::TextResult text_result;
        if (!editors::EditText(hwnd_, request, &text_result))
        {
            return false;
        }
        machine = std::move(text_result.text);
    }
    return ConnectRemoteRegistry(machine, true);
}

bool MainWindow::Impl::ConnectRemoteRegistry(const std::wstring& name, bool open_new_tab)
{
    std::wstring machine = name;
    machine = NormalizeMachineName(machine);
    if (machine.empty())
    {
        ui::ShowError(hwnd_, L"Computer name is required.");
        return false;
    }

    HKEY hklm = nullptr;
    LONG result = RegConnectRegistryW(machine.c_str(), HKEY_LOCAL_MACHINE, &hklm);
    if (result != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, FormatWin32Error(result));
        return false;
    }

    HKEY hku = nullptr;
    LONG hku_result = RegConnectRegistryW(machine.c_str(), HKEY_USERS, &hku);

    if (registry_mode_ == RegistryMode::kOffline)
    {
        if (!ConfirmOfflineChanges(L"The offline registry has unsaved changes.\n"
                                   L"Save before switching?"))
        {
            if (hku)
            {
                RegCloseKey(hku);
            }
            RegCloseKey(hklm);
            return false;
        }
        std::wstring error;
        if (!UnloadOfflineRegistry(&error))
        {
            if (!error.empty())
            {
                ui::ShowError(hwnd_, error);
            }
            if (hku)
            {
                RegCloseKey(hku);
            }
            RegCloseKey(hklm);
            return false;
        }
    }

    if (tab_ && open_new_tab)
    {
        AddRegistryTab(RegistryMode::kRemote, L"Remote Registry");
    }
    ReleaseRemoteRegistry();
    registry_mode_ = RegistryMode::kRemote;
    remote_machine_ = machine;
    remote_hklm_ = hklm;
    remote_hku_ = hku;
    UpdateRegistryTabEntry(RegistryMode::kRemote, L"", remote_machine_);

    std::wstring prefix = machine + L"\\";
    std::vector<RegistryRootEntry> roots;
    roots.push_back({remote_hklm_, L"HKEY_LOCAL_MACHINE", prefix + L"HKEY_LOCAL_MACHINE", L""});
    if (remote_hku_)
    {
        roots.push_back({remote_hku_, L"HKEY_USERS", prefix + L"HKEY_USERS", L""});
    }

    UpdateTabText(L"Remote Registry (" + StripMachinePrefix(machine) + L")");
    ApplyRegistryRoots(roots);
    RefreshRegistryTabLabels();

    if (hku_result != ERROR_SUCCESS)
    {
        std::wstring message = L"Connected to HKEY_LOCAL_MACHINE, but HKEY_USERS was unavailable.\n";
        message += FormatWin32Error(hku_result);
        ui::ShowError(hwnd_, message);
    }
    return true;
}

bool MainWindow::Impl::SwitchToOfflineRegistry()
{
    const int choice =
        ui::PromptChoice(hwnd_, L"Load the offline registry from a single hive file, or from a folder of hives?", L"Offline Registry", L"Hive File", L"Folder", L"Cancel", {90, 70, 70}, 470);
    std::wstring hive_path;
    HRESULT hr = S_OK;
    if (choice == IDYES)
    {
        hr = win32::ChooseFileToOpen(hwnd_, kOfflineHiveFilter, &hive_path);
    }
    else if (choice == IDNO)
    {
        hr = win32::ChooseFolder(hwnd_, &hive_path);
    }
    else
    {
        return false;
    }
    if (!ui::ReportFileDialogResult(hwnd_, hr))
    {
        return false;
    }
    return LoadOfflineRegistryFromPath(hive_path, true);
}

bool MainWindow::Impl::LoadOfflineRegistryFromPath(const std::wstring& path, bool open_new_tab)
{
    if (registry_mode_ == RegistryMode::kOffline && !offline_roots_.empty())
    {
        if (!ConfirmOfflineChanges(L"The offline registry has unsaved changes.\n"
                                   L"Save before switching?"))
        {
            return false;
        }
        std::wstring error;
        if (!UnloadOfflineRegistry(&error))
        {
            if (!error.empty())
            {
                ui::ShowError(hwnd_, error);
            }
            return false;
        }
    }

    std::wstring selection_path = TrimTrailingSeparators(path);
    if (selection_path.empty())
    {
        return false;
    }

    bool is_dir = IsDirectoryPath(selection_path);
    std::vector<OfflineHiveCandidate> candidates;
    if (is_dir)
    {
        CollectOfflineHivesInFolder(selection_path, &candidates);
        if (candidates.empty())
        {
            ui::ShowError(hwnd_, L"The selected folder doesn't contain a registry hive file.");
            return false;
        }
    }
    else
    {
        std::wstring mount_name = TrimWhitespace(FileBaseName(selection_path));
        if (mount_name.empty())
        {
            mount_name = L"OfflineHive";
        }
        candidates.push_back({selection_path, mount_name});
    }

    offline_root_name_ = ResolveOfflineRootName(selection_path, is_dir, browse_.current_node());
    if (offline_root_name_.empty())
    {
        offline_root_name_ = L"HKEY_LOCAL_MACHINE";
    }

    std::wstring error;
    std::vector<HKEY> handles;
    std::vector<std::wstring> labels;
    std::vector<std::wstring> paths;
    std::vector<RegistryRootEntry> roots;
    handles.reserve(candidates.size());
    labels.reserve(candidates.size());
    paths.reserve(candidates.size());
    roots.reserve(candidates.size());
    auto close_handles = [&](std::vector<HKEY>* to_close) {
        if (!to_close)
        {
            return;
        }
        for (HKEY root : *to_close)
        {
            RegistryStore::CloseOfflineHive(root, nullptr);
        }
    };
    for (const auto& candidate : candidates)
    {
        HKEY hive_handle = nullptr;
        if (!RegistryStore::OpenOfflineHive(candidate.path, &hive_handle, &error))
        {
            close_handles(&handles);
            if (!error.empty())
            {
                ui::ShowError(hwnd_, error);
            }
            return false;
        }
        std::wstring label = TrimWhitespace(candidate.label);
        if (label.empty())
        {
            label = TrimWhitespace(FileBaseName(candidate.path));
            if (label.empty())
            {
                label = L"OfflineHive";
            }
        }
        std::wstring path_name = offline_root_name_ + L"\\" + label;
        roots.push_back({hive_handle, label, path_name, L""});
        handles.push_back(hive_handle);
        labels.push_back(label);
        paths.push_back(candidate.path);
    }

    if (tab_ && open_new_tab)
    {
        AddRegistryTab(RegistryMode::kOffline, L"Offline Registry");
    }

    ReleaseRemoteRegistry();
    registry_mode_ = RegistryMode::kOffline;
    offline_roots_ = std::move(handles);
    offline_root_labels_ = std::move(labels);
    offline_root_paths_ = std::move(paths);
    if (offline_roots_.size() == 1)
    {
        offline_root_ = offline_roots_.front();
        offline_mount_ = offline_root_labels_.front();
    }
    else
    {
        offline_root_ = nullptr;
        offline_mount_.clear();
    }
    RegistryStore::SetOfflineRoots(offline_roots_);

    std::wstring tab_text = L"Offline Registry";
    if (offline_roots_.size() == 1 && !offline_root_name_.empty() && !offline_mount_.empty())
    {
        tab_text = L"Offline Registry (" + offline_root_name_ + L"\\" + offline_mount_ + L")";
    }
    else if (!offline_root_name_.empty())
    {
        tab_text = L"Offline Registry (" + offline_root_name_ + L")";
    }
    UpdateTabText(tab_text);
    UpdateRegistryTabEntry(RegistryMode::kOffline, selection_path, L"");
    ApplyRegistryRoots(roots);
    RefreshRegistryTabLabels();
    HistoryEntry history;
    history.action = L"Load offline registry";
    history.new_data = selection_path;
    AppendHistoryEntry(std::move(history));
    return true;
}

bool MainWindow::Impl::SaveOfflineRegistry()
{
    if (registry_mode_ != RegistryMode::kOffline || offline_roots_.empty())
    {
        ui::ShowError(hwnd_, L"No offline registry is loaded.");
        return false;
    }
    if (offline_roots_.size() > 1)
    {
        if (offline_root_paths_.size() != offline_roots_.size())
        {
            ui::ShowError(hwnd_, L"Failed to resolve offline hive paths for saving.");
            return false;
        }
        for (size_t i = 0; i < offline_roots_.size(); ++i)
        {
            const std::wstring& path = offline_root_paths_[i];
            if (path.empty())
            {
                ui::ShowError(hwnd_, L"Failed to resolve offline hive path for saving.");
                return false;
            }
            std::wstring error;
            if (!SaveHiveAtomically(offline_roots_[i], path, &error))
            {
                ui::ShowError(hwnd_, error.empty() ? L"Failed to save offline hive." : error);
                return false;
            }
        }
        ClearOfflineDirty();
        HistoryEntry history;
        history.action = L"Save offline registry";
        history.new_data = std::to_wstring(offline_roots_.size()) + L" hives";
        AppendHistoryEntry(std::move(history));
        return true;
    }
    if (!offline_root_)
    {
        ui::ShowError(hwnd_, L"No offline registry is loaded.");
        return false;
    }

    std::wstring path;
    if (!ui::PromptSaveFile(hwnd_, ui::kHiveFileFilter, &path))
    {
        return false;
    }

    std::wstring error;
    if (!SaveHiveAtomically(offline_root_, path, &error))
    {
        ui::ShowError(hwnd_, error.empty() ? L"Failed to save offline hive." : error);
        return false;
    }
    ClearOfflineDirty();
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
            ui::PromptKeyChoice(hwnd_, L"The key was opened, but it doesn't contain this value:", registry_path::DisplayName(value_name), L"Value not found", L"OK", L"", L"");
        }
        else
        {
            SelectValueWhenReady(value_name);
        }
        return;
    }
    if (path.empty())
    {
        return;
    }
    std::wstring nearest;
    if (!FindNearestExistingPath(path, &nearest) || nearest.empty())
    {
        ui::ShowWarning(hwnd_, L"Registry path not found.");
        return;
    }
    std::wstring message = L"The registry key doesn't exist:";
    if (read_only_)
    {
        message += L"\nRead only mode is enabled.";
        int result = ui::PromptKeyChoice(hwnd_, message, path, L"Registry path not found", L"Go to nearest key", L"", L"Cancel", {150, 70, 70});
        if (result == IDYES)
        {
            SelectTreePath(nearest);
        }
        return;
    }
    int result = ui::PromptKeyChoice(hwnd_, message, path, L"Registry path not found", L"Go to nearest key", L"Create key", L"Cancel", {150, 100, 70});
    if (result == IDYES)
    {
        SelectTreePath(nearest);
        return;
    }
    if (result == IDNO)
    {
        if (!CreateRegistryPath(path))
        {
            ui::ShowError(hwnd_, L"Failed to create registry key.");
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
        ui::ShowWarning(hwnd_, L"Registry path not found:\n" + target);
    }
}

bool MainWindow::Impl::ResolveJumpTarget(const std::wstring& target, std::wstring* key_path, std::wstring* value_name, bool* value_missing) const
{
    const auto unwrap = [](std::wstring text, std::wstring_view pairs) {
        text = util::TrimWhitespace(text);
        for (size_t pair = 0; pair + 1 < pairs.size(); pair += 2)
        {
            if (text.size() >= 2 && text.front() == pairs[pair] && text.back() == pairs[pair + 1])
            {
                return util::TrimWhitespace(std::wstring_view(text).substr(1, text.size() - 2));
            }
        }
        return text;
    };
    RegistryNode node;
    KeyInfo info = {};
    ValueEntry value;
    const auto key_exists = [&](const std::wstring& path) {
        return !path.empty() && ResolvePathToNode(path, &node) && RegistryStore::QueryKeyInfo(node, &info);
    };
    value_name->clear();
    *value_missing = false;
    const std::wstring text = unwrap(target, L"\"\"''[]");
    *key_path = NormalizeRegistryPath(text);
    if (key_path->empty() || key_exists(*key_path))
    {
        return !key_path->empty();
    }

    std::wstring missing_key;
    std::wstring existing_key;
    std::wstring existing_name;
    for (size_t split = 1; split < text.size(); ++split)
    {
        const bool colon =
            text[split] == L':' && (iswspace(text[split - 1]) || split + 1 == text.size() || iswspace(text[split + 1]));
        if (text[split] != L'!' && !colon)
        {
            continue;
        }
        const std::wstring key = NormalizeRegistryPath(text.substr(0, split));
        const std::wstring name =
            registry_path::RawName(colon ? unwrap(text.substr(split + 1), L"\"\"") : text.substr(split + 1));
        if (!key_exists(key))
        {
            missing_key = missing_key.empty() ? key : missing_key;
            continue;
        }
        if (name.empty() || RegistryStore::QueryValue(node, name, &value))
        {
            *key_path = key;
            *value_name = name;
            return true;
        }
        existing_key = key;
        existing_name = name;
    }
    if (!existing_key.empty())
    {
        *key_path = existing_key;
        *value_name = existing_name;
        *value_missing = true;
        return true;
    }
    const std::wstring normalized = *key_path;
    for (size_t slash = normalized.rfind(L'\\'); slash != std::wstring::npos && slash > 0;
         slash = normalized.rfind(L'\\', slash - 1))
    {
        if (key_exists(normalized.substr(0, slash)) &&
            RegistryStore::QueryValue(node, normalized.substr(slash + 1), &value))
        {
            *key_path = normalized.substr(0, slash);
            *value_name = normalized.substr(slash + 1);
            return true;
        }
    }
    if (!missing_key.empty())
    {
        *key_path = missing_key;
    }
    return false;
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
    return registry_mode_ == RegistryMode::kLocal || SwitchToLocalRegistry();
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
    if (!ResolveJumpTarget(target, &key_path, &value_name, &value_missing) &&
        !FindNearestExistingPath(std::wstring(key_path), &key_path))
    {
        return false;
    }
    return NavigateToResolvedExternalJump(key_path, value_missing ? std::wstring() : value_name);
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

    pending_external_value_key_path_.clear();
    pending_external_value_name_.clear();
    if (!value_name.empty())
    {
        pending_external_value_key_path_ = key_path;
        pending_external_value_name_ = value_name;
        if (!SelectValueByName(value_name) && browse_.current_node() && !value_list_loading_)
        {
            UpdateValueListForNode(browse_.current_node());
        }
    }
    return true;
}

} // namespace regkit
