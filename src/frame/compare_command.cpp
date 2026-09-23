// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_detail.h"
#include "frame/window_impl.h"
#include "win32/system_error.h"

namespace regkit
{
using namespace command_detail;

namespace
{

bool ResolveRemoteNode(const std::wstring& machine, HKEY hklm, HKEY hku, const std::wstring& base, RegistryNode* node)
{
    if (!node)
    {
        return false;
    }
    std::wstring rest = base;
    const std::wstring prefix = machine + L"\\";
    if (StartsWithInsensitive(rest, prefix))
    {
        rest = rest.substr(prefix.size());
    }
    const size_t slash = rest.find(L'\\');
    const std::wstring root_name = slash == std::wstring::npos ? rest : rest.substr(0, slash);
    if (EqualsInsensitive(root_name, L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(root_name, L"HKLM"))
    {
        node->root = hklm;
    }
    else if (EqualsInsensitive(root_name, L"HKEY_USERS") || EqualsInsensitive(root_name, L"HKU"))
    {
        node->root = hku;
    }
    if (!node->root)
    {
        return false;
    }
    node->root_name = machine + L"\\" + root_name;
    node->subkey = slash == std::wstring::npos ? std::wstring() : rest.substr(slash + 1);
    KeyInfo info = {};
    return RegistryStore::QueryKeyInfo(*node, &info);
}

} // namespace

void MainWindow::Impl::StartCompareRegistries()
{
    CompareDialogDefaults defaults;
    CompareDialogSelection left;
    CompareDialogSelection right;
    left.type = CompareSourceType::kRegistry;
    right.type = CompareSourceType::kRegistry;
    left.recursive = true;
    right.recursive = true;
    if (browse_.current_node())
    {
        left.key_path = registry_path::Build(*browse_.current_node());
    }
    else
    {
        left.key_path = L"HKEY_LOCAL_MACHINE";
    }
    right.key_path = left.key_path;
    defaults.left = left;
    defaults.right = right;

    CompareDialogResult selection;
    if (!ShowCompareDialog(hwnd_, defaults, &selection))
    {
        return;
    }

    auto normalize_base = [&](const CompareDialogSelection& sel, std::wstring* out_base) -> bool {
        if (!out_base)
        {
            return false;
        }
        std::wstring base = NormalizeRegistryPath(sel.key_path);
        if (sel.type == CompareSourceType::kOfflineHive)
        {
            *out_base = std::move(base);
            return true;
        }
        if (base.empty())
        {
            return false;
        }
        *out_base = base;
        return true;
    };

    auto build_snapshot = [&](const CompareDialogSelection& source, search::compare::Snapshot* snapshot, std::wstring* error) -> bool {
        std::wstring base;
        if (!normalize_base(source, &base))
        {
            if (error)
            {
                *error = L"Invalid registry path.";
            }
            return false;
        }
        if (source.type == CompareSourceType::kRegFile)
        {
            return search::compare::LoadRegFile(
                source.file_path,
                base,
                source.recursive,
                [this](const std::wstring& path) { return NormalizeRegistryPath(path); },
                snapshot,
                error
            );
        }
        if (source.type == CompareSourceType::kOfflineHive)
        {
            HKEY hive = nullptr;
            if (!RegistryStore::OpenOfflineHive(source.file_path, &hive, error))
            {
                return false;
            }
            RegistryStore::AddOfflineRoot(hive);
            RegistryNode hive_node;
            hive_node.root = hive;
            hive_node.root_name = FileNameOnly(source.file_path);
            hive_node.subkey = base;
            const bool ok = search::compare::CaptureRegistry(base.empty() ? hive_node.root_name : base, hive_node, source.recursive, snapshot, error);
            RegistryStore::RemoveOfflineRoot(hive);
            RegistryStore::CloseOfflineHive(hive, nullptr);
            if (!ok && error && error->empty())
            {
                *error = L"Failed to read the hive file.\n" + source.file_path;
            }
            return ok;
        }

        RegistryNode node;
        if (source.type == CompareSourceType::kNetwork)
        {
            const std::wstring machine = TrimWhitespace(source.file_path);
            if (machine.empty())
            {
                if (error)
                {
                    *error = L"Select a computer to compare against.";
                }
                return false;
            }
            HKEY hklm = nullptr;
            const LONG connected = RegConnectRegistryW(machine.c_str(), HKEY_LOCAL_MACHINE, &hklm);
            if (connected != ERROR_SUCCESS)
            {
                if (error)
                {
                    *error = util::FormatWin32Error(connected);
                }
                return false;
            }
            HKEY hku = nullptr;
            RegConnectRegistryW(machine.c_str(), HKEY_USERS, &hku);
            bool ok = ResolveRemoteNode(machine, hklm, hku, base, &node);
            if (!ok && error)
            {
                *error = L"Network registry path not found.\n" + base;
            }
            if (ok)
            {
                ok = search::compare::CaptureRegistry(base, node, source.recursive, snapshot, error);
            }
            if (hku)
            {
                RegCloseKey(hku);
            }
            RegCloseKey(hklm);
            return ok;
        }
        KeyInfo info = {};
        if (!ResolvePathToNode(base, &node) || !RegistryStore::QueryKeyInfo(node, &info))
        {
            if (error)
            {
                *error = L"Registry path not found.\n" + base;
            }
            return false;
        }
        return search::compare::CaptureRegistry(base, node, source.recursive, snapshot, error);
    };

    search::compare::Snapshot left_snapshot;
    search::compare::Snapshot right_snapshot;
    std::wstring error;
    if (!build_snapshot(selection.left, &left_snapshot, &error))
    {
        if (!error.empty())
        {
            ui::ShowError(hwnd_, error);
        }
        return;
    }
    error.clear();
    if (!build_snapshot(selection.right, &right_snapshot, &error))
    {
        if (!error.empty())
        {
            ui::ShowError(hwnd_, error);
        }
        return;
    }

    std::vector<search::compare::Row> rows =
        search::compare::BuildRows(left_snapshot, right_snapshot, selection.filter);
    std::wstring tab_label = L"Registry Comparison";

    auto source_ref = [this](const CompareDialogSelection& sel) {
        switch (sel.type)
        {
        case CompareSourceType::kRegFile:
            return search::Source{search::Source::Kind::kRegFile, sel.file_path};
        case CompareSourceType::kOfflineHive:
            return search::Source{search::Source::Kind::kOffline, sel.file_path};
        case CompareSourceType::kNetwork:
            return search::Source{search::Source::Kind::kRemote, sel.file_path};
        default:
            break;
        }
        return search::Source{};
    };

    SearchTab tab;
    tab.label = std::move(tab_label);
    tab.compare_rows = std::move(rows);
    tab.is_compare = true;
    tab.compare_filter = selection.filter;
    tab.sources = {source_ref(selection.left), source_ref(selection.right)};
    search_tabs_.push_back(std::move(tab));
    int search_index = static_cast<int>(search_tabs_.size() - 1);
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(search_tabs_.back().label.c_str());
    int tab_index = TabCtrl_GetItemCount(tab_);
    TabCtrl_InsertItem(tab_, tab_index, &item);
    tabs_.push_back({TabEntry::Kind::kSearch, search_index});

    UpdateTabWidth();
    SelectTabIndex(tab_index);
    active_search_tab_index_ = tab_index;
    UpdateSearchResultsView();
    ApplyViewVisibility();
    UpdateStatus();
}

search::Source MainWindow::Impl::TabSource(int index) const
{
    if (index < 0 || static_cast<size_t>(index) >= tabs_.size())
    {
        return {};
    }
    const TabEntry& entry = tabs_[static_cast<size_t>(index)];
    if (entry.kind == TabEntry::Kind::kRegFile)
    {
        return {search::Source::Kind::kRegFile, entry.reg_file_path};
    }
    if (entry.kind == TabEntry::Kind::kRegistry)
    {
        if (entry.registry_mode == RegistryMode::kOffline)
        {
            return {search::Source::Kind::kOffline, entry.offline_path};
        }
        if (entry.registry_mode == RegistryMode::kRemote)
        {
            return {search::Source::Kind::kRemote, entry.remote_machine};
        }
    }
    return {};
}

search::Source MainWindow::Impl::CurrentTabSource() const
{
    return TabSource(tab_ ? TabCtrl_GetCurSel(tab_) : -1);
}

int MainWindow::Impl::FindSourceTab(const search::Source& source) const
{
    for (size_t i = 0; i < tabs_.size(); ++i)
    {
        const TabEntry& entry = tabs_[i];
        if (entry.kind == TabEntry::Kind::kSearch)
        {
            continue;
        }
        const search::Source candidate = TabSource(static_cast<int>(i));
        if (candidate.kind != source.kind)
        {
            continue;
        }
        if (source.kind == search::Source::Kind::kLocal || source.name.empty() ||
            EqualsInsensitive(candidate.name, source.name))
        {
            return static_cast<int>(i);
        }
    }
    return -1;
}

void MainWindow::Impl::OpenSourceEntry(const search::Source& source, const std::wstring& path, const std::wstring& value_name, bool new_tab)
{
    if (!tab_ || path.empty())
    {
        return;
    }
    const int existing = FindSourceTab(source);
    if (!new_tab && existing < 0)
    {
        return;
    }
    switch (source.kind)
    {
    case search::Source::Kind::kRegFile:
        pending_compare_key_path_ = path;
        pending_compare_value_name_ = value_name;
        if (new_tab)
        {
            if (!OpenRegFileTab(source.name, true))
            {
                pending_compare_key_path_.clear();
                pending_compare_value_name_.clear();
                return;
            }
        }
        else if (existing == TabCtrl_GetCurSel(tab_))
        {
            SyncRegFileTabSelection();
        }
        else
        {
            ActivateTabIndex(existing);
        }
        break;
    case search::Source::Kind::kOffline:
        {
            if (new_tab)
            {
                OpenLocalRegistryTab();
                if (!LoadOfflineRegistryFromPath(source.name, false))
                {
                    return;
                }
            }
            else
            {
                ActivateTabIndex(existing);
            }
            std::wstring target = path;
            if (!offline_mount_.empty())
            {
                for (const std::wstring& prefix : {offline_mount_, FileNameOnly(source.name)})
                {
                    if (prefix.empty())
                    {
                        continue;
                    }
                    if (EqualsInsensitive(target, prefix))
                    {
                        target.clear();
                        break;
                    }
                    if (StartsWithInsensitive(target, prefix + L"\\"))
                    {
                        target = target.substr(prefix.size() + 1);
                        break;
                    }
                }
                const std::wstring mount = offline_root_name_ + L"\\" + offline_mount_;
                target = target.empty() ? mount : mount + L"\\" + target;
            }
            SelectTreePath(target);
            break;
        }
    case search::Source::Kind::kRemote:
        if (new_tab)
        {
            OpenLocalRegistryTab();
            if (!ConnectRemoteRegistry(source.name))
            {
                return;
            }
        }
        else
        {
            ActivateTabIndex(existing);
        }
        SelectTreePath(path);
        break;
    case search::Source::Kind::kLocal:
    default:
        if (new_tab)
        {
            OpenLocalRegistryTab();
        }
        else
        {
            ActivateTabIndex(existing);
        }
        SelectTreePath(path);
        break;
    }
    ApplyViewVisibility();
    UpdateStatus();
    if (!value_name.empty() && pending_compare_key_path_.empty())
    {
        SelectValueWhenReady(value_name);
    }
}

} // namespace regkit
