// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "registry/hive_files.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/system/shell_open.h>

namespace regkit
{
using namespace window_detail;

bool MainWindow::Impl::SaveRegFileTab(int tab_index)
{
    if (!IsRegFileTabIndex(tab_index) || static_cast<size_t>(tab_index) >= tabs_.size())
    {
        return false;
    }
    TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
    if (entry.reg_file_path.empty())
    {
        return false;
    }
    std::wstring content;
    if (!BuildRegFileContent(entry, &content))
    {
        return false;
    }
    if (win32ui::WriteTextFileAtomic(entry.reg_file_path.c_str(), content, true) != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, util::Tr(L"Failed to save registry file."));
        return false;
    }
    if (entry.reg_file_dirty)
    {
        entry.reg_file_dirty = false;
    }
    AppendHistoryEntry(L"Save .reg file " + util::FileName(entry.reg_file_path), L"", entry.reg_file_path);
    return true;
}

bool MainWindow::Impl::ExportRegFileTab(int tab_index, const std::wstring& path)
{
    if (!IsRegFileTabIndex(tab_index) || static_cast<size_t>(tab_index) >= tabs_.size())
    {
        return false;
    }
    if (path.empty())
    {
        return false;
    }
    std::wstring content;
    if (!BuildRegFileContent(tabs_[static_cast<size_t>(tab_index)], &content))
    {
        return false;
    }
    std::wstring target = util::EnsureFileExtension(path, L".reg");
    if (win32ui::WriteTextFileAtomic(target.c_str(), content, true) != ERROR_SUCCESS)
    {
        ui::ShowError(hwnd_, util::Tr(L"Failed to export registry file."));
        return false;
    }
    return true;
}

bool MainWindow::Impl::BuildRegFileContent(const TabEntry& entry, std::wstring* out) const
{
    if (!out)
    {
        return false;
    }
    out->clear();
    if (entry.kind != TabEntry::Kind::kRegFile)
    {
        return false;
    }

    regfile::Writer writer;
    std::function<void(const VirtualRegistryKey&, const std::wstring&)> append_key;
    append_key = [&](const VirtualRegistryKey& key, const std::wstring& full_path) {
        if (key.removed)
        {
            writer.AppendRemovedKey(full_path);
        }
        if (!key.removed || key.recreated || !key.values.empty())
        {
            std::vector<const RegistryValue*> values;
            values.reserve(key.values.size());
            for (const auto& source : key.values)
            {
                values.push_back(&source.second);
            }
            writer.AppendKey(full_path, std::move(values));
            for (const std::wstring& name : key.removed_values)
            {
                if (!key.values.contains(ToLower(name)))
                {
                    writer.AppendRemovedValue(name);
                }
            }
        }
        std::vector<const VirtualRegistryKey*> children;
        children.reserve(key.children.size());
        for (const auto& child : key.children)
        {
            if (child.second)
            {
                children.push_back(child.second.get());
            }
        }
        std::sort(children.begin(), children.end(), [](const VirtualRegistryKey* left, const VirtualRegistryKey* right) {
            if (!left || !right)
            {
                return left != nullptr;
            }
            return util::CompareInsensitive(left->name, right->name) < 0;
        });
        for (const auto* child : children)
        {
            if (!child)
            {
                continue;
            }
            append_key(*child, full_path + L"\\" + child->name);
        }
    };

    for (const auto& root : entry.reg_file_roots)
    {
        if (!root.data || !root.data->root)
        {
            continue;
        }
        std::wstring root_name = root.name;
        if (root_name.empty())
        {
            root_name = root.data->root_name;
        }
        if (root_name.empty())
        {
            continue;
        }
        append_key(*root.data->root, root_name);
    }
    *out = std::move(writer).Finish();
    return true;
}

void MainWindow::Impl::ReleaseRegFileRoots(TabEntry* entry)
{
    if (!entry)
    {
        return;
    }
    for (auto& root : entry->reg_file_roots)
    {
        if (root.root)
        {
            RegistryStore::UnregisterVirtualRoot(root.root);
            root.root = nullptr;
        }
        root.data.reset();
    }
    entry->reg_file_roots.clear();
    // clear undo and hidden tree references, released roots would reach the live registry
    entry->session.reset();
    if (entry->tree && entry->tree != &browse_.tree())
    {
        entry->tree->Clear();
    }
}

void MainWindow::Impl::StartRegFileParse(const std::wstring& path, const std::wstring& session_key)
{
    if (path.empty() || session_key.empty() ||
        reg_file_parse_sessions_.find(session_key) != reg_file_parse_sessions_.end())
    {
        return;
    }
    auto session = std::make_unique<RegFileParseSession>();
    session->source_path = path;
    session->source_lower = session_key;
    HWND hwnd = hwnd_;
    RegFileParseSession* session_ptr = session.get();
    // runs here without a thread, the result is posted like that of a thread and finds the session registered below
    session->work.StartOrRun(L"RegFileParseThread", [this, session_ptr, hwnd](uint64_t generation, std::atomic_bool& cancel) {
        auto payload = std::make_unique<RegFileParsePayload>();
        payload->generation = generation;
        payload->source_path = session_ptr->source_path;
        payload->source_lower = session_ptr->source_lower;
        std::wstring parse_error;
        std::vector<regfile::VirtualRoot> parsed_roots;
        bool cancelled = false;
        if (!regfile::LoadVirtualRoots(payload->source_path, &parsed_roots, &parse_error, &cancel, &cancelled))
        {
            if (!cancelled && parse_error.empty())
            {
                parse_error = util::Tr(L"Failed to read registry file.");
            }
        }
        payload->roots = std::move(parsed_roots);
        payload->error = std::move(parse_error);
        payload->cancelled = cancelled;
        if (!hwnd || !IsWindow(hwnd) ||
            !win32ui::PostPayload(hwnd, frame::message_id::kRegFileLoadReady, 0, payload))
        {
            return;
        }
    });
    reg_file_parse_sessions_.emplace(session_key, std::move(session));
}

void MainWindow::Impl::OpenSavedFile(const std::wstring& path, win32::OpenAfter open_after)
{
    if (open_after == win32::OpenAfter::kRegKit)
    {
        OpenRegFileTab(path, true);
    }
    else if (open_after == win32::OpenAfter::kEditor)
    {
        ui::ReportFileDialogResult(hwnd_, win32ui::ShellOpen(hwnd_, path.c_str(), L".txt"));
    }
}

bool MainWindow::Impl::OpenFile(const std::wstring& path)
{
    return IsHiveFile(path) ? LoadOfflineRegistryFromPath(path, true) : OpenRegFileTab(path);
}

bool MainWindow::Impl::OpenRegFileTab(const std::wstring& path, bool force_new_tab)
{
    if (!tab_ || path.empty())
    {
        return false;
    }
    if (!util::IsFile(path.c_str()))
    {
        ui::ShowError(hwnd_, util::Tr(L"Registry file not found."));
        return false;
    }
    std::wstring label = util::FileName(path);
    if (label.empty())
    {
        label = util::Tr(L"Registry File");
    }
    for (size_t i = 0; !force_new_tab && i < tabs_.size(); ++i)
    {
        TabEntry& entry = tabs_[i];
        if (entry.kind != TabEntry::Kind::kRegFile)
        {
            continue;
        }
        if (EqualsInsensitive(entry.reg_file_path, path))
        {
            entry.reg_file_path = path;
            entry.reg_file_label = label;
            entry.reg_file_loading = true;
            if (entry.reg_file_session_key.empty())
            {
                entry.reg_file_session_key = ToLower(path) + L"|" + std::to_wstring(++reg_file_session_serial_);
            }
            TCITEMW item = {};
            item.mask = TCIF_TEXT;
            item.pszText = const_cast<wchar_t*>(label.c_str());
            TabCtrl_SetItem(tab_, static_cast<int>(i), &item);
            SelectTabIndex(static_cast<int>(i));
            SyncRegFileTabSelection();
            ApplyViewVisibility();
            UpdateStatus();
            StartRegFileParse(path, entry.reg_file_session_key);
            AppendHistoryEntry(L"Open .reg file " + util::FileName(path), L"", path);
            return true;
        }
    }

    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(label.c_str());
    int index = TabCtrl_GetItemCount(tab_);
    TabCtrl_InsertItem(tab_, index, &item);
    TabEntry entry;
    entry.kind = TabEntry::Kind::kRegFile;
    entry.reg_file_path = path;
    entry.reg_file_label = label;
    entry.reg_file_session_key = ToLower(path) + L"|" + std::to_wstring(++reg_file_session_serial_);
    entry.reg_file_dirty = false;
    entry.reg_file_loading = true;
    const std::wstring session_key = entry.reg_file_session_key;
    tabs_.push_back(std::move(entry));
    UpdateTabWidth();
    SelectTabIndex(index);
    SyncRegFileTabSelection();
    ApplyViewVisibility();
    UpdateStatus();
    StartRegFileParse(path, session_key);
    AppendHistoryEntry(L"Open .reg file " + util::FileName(path), L"", path);
    return true;
}

}
