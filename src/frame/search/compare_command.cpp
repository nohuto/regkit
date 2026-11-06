// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"
#include "win32/shell_paths.h"
#include "win32/translation.h"

#include <win32ui/system/error_text.h>

#include <array>

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

}

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
    workspace::DialogState state = LoadDialogState();
    workspace::DialogFields last(&state, false);
    CompareDialogFields(last, &defaults);
    for (CompareDialogSelection* side : {&defaults.left, &defaults.right})
    {
        if (side->type == CompareSourceType::kRegistry)
        {
            side->key_path = left.key_path;
        }
    }

    CompareDialogResult selection;
    if (!ShowCompareDialog(hwnd_, defaults, &selection))
    {
        return;
    }
    CompareDialogDefaults used{selection.left, selection.right, selection.filter};
    workspace::DialogFields store(&state, true);
    CompareDialogFields(store, &used);
    ui::ReportSaveFailure(hwnd_, workspace::DialogStatePath(), workspace::SaveDialogState(state));

    // resolve paths and visible tab keys here, worker only reads
    struct Side
    {
        CompareDialogSelection source;
        std::wstring base;
        RegistryNode node;
    };
    std::array<Side, 2> sides = {Side{selection.left}, Side{selection.right}};
    for (Side& side : sides)
    {
        side.base = NormalizeRegistryPath(side.source.key_path);
        if (side.base.empty() && side.source.type != CompareSourceType::kOfflineHive)
        {
            ui::ShowError(hwnd_, util::Tr(L"Invalid registry path."));
            return;
        }
        KeyInfo info = {};
        if (side.source.type == CompareSourceType::kRegistry &&
            (!ResolvePathToNode(side.base, &side.node) || !RegistryStore::QueryKeyInfo(side.node, &info)))
        {
            ui::ShowError(hwnd_, util::TrDetail(L"Registry path not found.", side.base));
            return;
        }
    }
    // the task releases its hold even if the worker throws or never starts
    auto hold = std::make_shared<SourceHold>();
    hold->sessions.push_back(session_);
    for (Side& side : sides)
    {
        if (side.source.type == CompareSourceType::kRegistry)
        {
            side.node.root = hold->Hold(side.node.root);
        }
    }

    auto source_ref = [shown = CurrentTabSource()](const CompareDialogSelection& sel) {
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
        return shown;
    };
    std::vector<search::Source> sources = {source_ref(selection.left), source_ref(selection.right)};

    const uint64_t started = compare_session_.Start(
        L"CompareThread",
        [hwnd = hwnd_, sides, hold, filter = selection.filter, sources, sid = util::GetCurrentUserSidString(), contexts = RegistryPathContexts()](
            uint64_t generation,
            std::atomic_bool& cancel
        ) {
            auto payload = std::make_unique<ComparePayload>();
            payload->generation = generation;
            payload->filter = filter;
            payload->sources = sources;
            auto build_snapshot = [&](const Side& side, search::compare::Snapshot* snapshot, std::wstring* error) -> bool {
                const CompareDialogSelection& source = side.source;
                if (source.type == CompareSourceType::kRegFile)
                {
                    if (!source.document)
                    {
                        *error = util::Tr(L"Failed to read registry file.");
                        return false;
                    }
                    return search::compare::LoadRegFile(
                        source.file_path,
                        *source.document,
                        side.base,
                        source.recursive,
                        [&](const std::wstring& path) { return NormalizeRegistryPath(path, sid, contexts); },
                        snapshot,
                        error,
                        &cancel
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
                    hive_node.root_name = util::FileName(source.file_path);
                    hive_node.subkey = side.base;
                    const bool ok = search::compare::CaptureRegistry(side.base.empty() ? hive_node.root_name : side.base, hive_node, source.recursive, snapshot, error, &cancel);
                    RegistryStore::RemoveOfflineRoot(hive);
                    RegistryStore::CloseOfflineHive(hive, nullptr);
                    if (!ok && error->empty())
                    {
                        *error = util::TrDetail(L"Failed to read the hive file.", source.file_path);
                    }
                    return ok;
                }
                if (source.type == CompareSourceType::kNetwork)
                {
                    const std::wstring machine = TrimWhitespace(source.file_path);
                    if (machine.empty())
                    {
                        *error = util::Tr(L"Select a computer to compare against.");
                        return false;
                    }
                    HKEY hklm = nullptr;
                    const LONG connected = RegConnectRegistryW(machine.c_str(), HKEY_LOCAL_MACHINE, &hklm);
                    if (connected != ERROR_SUCCESS)
                    {
                        *error = win32ui::FormatWin32Error(connected);
                        return false;
                    }
                    HKEY hku = nullptr;
                    RegConnectRegistryW(machine.c_str(), HKEY_USERS, &hku);
                    RegistryNode node;
                    bool ok = ResolveRemoteNode(machine, hklm, hku, side.base, &node);
                    if (!ok)
                    {
                        *error = util::TrDetail(L"Network registry path not found.", side.base);
                    }
                    if (ok)
                    {
                        ok = search::compare::CaptureRegistry(side.base, node, source.recursive, snapshot, error, &cancel);
                    }
                    if (hku)
                    {
                        RegCloseKey(hku);
                    }
                    RegCloseKey(hklm);
                    return ok;
                }
                return search::compare::CaptureRegistry(side.base, side.node, source.recursive, snapshot, error, &cancel);
            };

            search::compare::Snapshot left_snapshot;
            search::compare::Snapshot right_snapshot;
            std::wstring left_error;
            std::wstring right_error;
            bool right_ok = false;
            const auto read_right = [&]() noexcept {
                try
                {
                    right_ok = build_snapshot(sides[1], &right_snapshot, &right_error);
                }
                catch (...)
                {
                }
            };
            // read both sides concurrently, registry reads scale with reader count
            std::jthread right_reader;
            try
            {
                right_reader = std::jthread(read_right);
            }
            catch (const std::exception&)
            {
                // no second thread, the right side is read after the left one
            }
            const bool left_ok = build_snapshot(sides[0], &left_snapshot, &left_error);
            if (right_reader.joinable())
            {
                right_reader.join();
            }
            else
            {
                read_right();
            }
            if (cancel.load())
            {
                return;
            }
            if (!left_ok || !right_ok)
            {
                payload->error = left_ok ? right_error : left_error;
            }
            else
            {
                payload->rows = search::compare::BuildRows(left_snapshot, right_snapshot, filter, &cancel);
                for (const search::compare::Snapshot* side : {&left_snapshot, &right_snapshot})
                {
                    if (!side->unreadable.empty())
                    {
                        payload->unreadable = util::TrDetail(L"Couldn't read the registry key.", side->base_path + L"\\" + side->unreadable.front() + L"\n" + util::TrLabel(L"Skipped", std::to_wstring(side->unreadable.size())));
                        break;
                    }
                }
            }
            // free before posting, UI thread joins this worker when the result arrives
            left_snapshot = {};
            right_snapshot = {};
            if (!cancel.load())
            {
                win32ui::PostPayload(hwnd, frame::message_id::kCompareReady, static_cast<WPARAM>(generation), payload);
            }
        }
    );
    if (!started)
    {
        ui::ShowStartError(hwnd_);
        return;
    }
    // keep message until compare result arrives
    SetStatusMessage(util::Tr(L"Compare Registries..."), 0);
}

void MainWindow::Impl::ApplyComparePayload(std::unique_ptr<ComparePayload> payload)
{
    if (!payload || !compare_session_.IsCurrent(payload->generation))
    {
        return;
    }
    compare_session_.Join();
    SetStatusMessage(L"");
    if (!payload->error.empty())
    {
        ui::ShowError(hwnd_, payload->error);
        return;
    }

    SearchTab tab;
    tab.label = util::Tr(L"Registry Comparison");
    tab.compare_rows = std::move(payload->rows);
    tab.is_compare = true;
    tab.compare_filter = payload->filter;
    tab.sources = std::move(payload->sources);
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
    if (!payload->unreadable.empty())
    {
        ui::ShowWarning(hwnd_, payload->unreadable);
    }
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
        return {search::Source::Kind::kLocal, {}, entry.registry_view};
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
        if (source.kind == search::Source::Kind::kLocal ? entry.registry_mode == RegistryMode::kLocal && candidate.view == source.view
                                                        : source.name.empty() || EqualsInsensitive(candidate.name, source.name))
        {
            return static_cast<int>(i);
        }
    }
    return -1;
}

bool MainWindow::Impl::OpenSourceEntry(const search::Source& source, const std::wstring& path, const std::wstring& value_name, bool new_tab)
{
    if (!tab_ || path.empty())
    {
        return false;
    }
    const int existing = FindSourceTab(source);
    if (!new_tab && existing < 0)
    {
        return false;
    }
    bool selected = false;
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
                return false;
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
        // a tab still parsing selects the key once its roots exist
        selected = pending_compare_key_path_.empty() && SelectTreePath(path);
        break;
    case search::Source::Kind::kOffline:
        {
            if (new_tab)
            {
                OpenLocalRegistryTab();
                if (!LoadOfflineRegistryFromPath(source.name, false))
                {
                    return false;
                }
            }
            else
            {
                ActivateTabIndex(existing);
            }
            std::wstring target = path;
            if (!session_->offline_mount.empty())
            {
                // paths start with the full mount, the bare mount or the hive file name
                const std::wstring mount = session_->offline_root_name + L"\\" + session_->offline_mount;
                for (const std::wstring& prefix : {mount, session_->offline_mount, util::FileName(source.name)})
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
                target = target.empty() ? mount : mount + L"\\" + target;
            }
            selected = SelectTreePath(target);
            break;
        }
    case search::Source::Kind::kRemote:
        if (new_tab)
        {
            OpenLocalRegistryTab();
            if (!ConnectRemoteRegistry(source.name))
            {
                return false;
            }
        }
        else
        {
            ActivateTabIndex(existing);
        }
        selected = SelectTreePath(path);
        break;
    case search::Source::Kind::kLocal:
    default:
        if (new_tab)
        {
            OpenLocalRegistryTab(source.view);
        }
        else
        {
            ActivateTabIndex(existing);
        }
        selected = SelectTreePath(path);
        break;
    }
    ApplyViewVisibility();
    UpdateStatus();
    if (!value_name.empty() && pending_compare_key_path_.empty())
    {
        SelectValueWhenReady(value_name);
    }
    return selected;
}

}
