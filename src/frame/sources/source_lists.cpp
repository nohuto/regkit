// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "records/escaped_fields.h"
#include "win32/text_transform.h"
#include "trace/trace_paths.h"

namespace regkit
{
using namespace window_detail;

namespace
{

template <typename Source, typename Resolve>
std::wstring SerializeActiveList(const std::vector<Source>& sources, const Resolve& resolve_bundled)
{
    std::wstring content;
    for (const auto& source : sources)
    {
        if (source.source_path.empty())
        {
            continue;
        }
        const bool bundled = !source.label.empty() && EqualsInsensitive(resolve_bundled(source.label), source.source_path);
        content.append(bundled ? source.label : source.source_path).push_back(L'\n');
    }
    return content;
}

}

void MainWindow::Impl::StartDefaultLoadWorker()
{
    if (default_load_session_.running())
    {
        return;
    }
    const std::wstring active_path = ActiveDefaultsPath();
    std::vector<std::wstring> entries;
    std::wstring error;
    if (!ReadActiveEntries(active_path, &entries, &error))
    {
        if (!error.empty())
        {
            ReportFileError(active_path, error);
        }
        return;
    }
    const HWND hwnd = hwnd_;
    // runs here without a thread, saving an unrestored list would drop its entries
    default_load_session_.StartOrRun(L"DefaultsLoadThread", [this, entries = std::move(entries), hwnd](uint64_t generation, const std::atomic_bool& cancel) {
        auto payload = std::make_unique<DefaultLoadPayload>();
        payload->generation = generation;
        std::unordered_set<std::wstring> loaded;
        for (const auto& entry : entries)
        {
            if (cancel.load())
            {
                return;
            }
            std::wstring source;
            std::wstring use_label;
            if (!ResolveActiveSource(entry, [this](const std::wstring& label) { return ResolveBundledDefaultPath(label); }, L"Default", &source, &use_label))
            {
                continue;
            }
            std::wstring source_lower = ToLower(source);
            if (!loaded.insert(source_lower).second)
            {
                continue;
            }
            defaults::Data data;
            if (!defaults::Load(
                    source,
                    [](const std::wstring& path) { return trace::NormalizeKeyPathBasic(path); },
                    &data,
                    nullptr,
                    nullptr,
                    &cancel
                ))
            {
                continue;
            }
            std::shared_ptr<const defaults::Data> default_data = std::make_shared<defaults::Data>(std::move(data));
            trace::Selection selection = {};
            selection.select_all = true;
            selection.recursive = true;
            payload->defaults.push_back(
                {use_label, source, default_data, std::make_shared<trace::Selection>(selection)}
            );
        }

        if (cancel.load())
        {
            return;
        }
        if (hwnd && IsWindow(hwnd))
        {
            win32ui::PostPayload(hwnd, frame::message_id::kDefaultLoadReady, 0, payload);
        }
    });
}

void MainWindow::Impl::StopDefaultLoadWorker()
{
    default_load_session_.CancelAndJoin();
}

void MainWindow::Impl::StopDefaultParseSessions()
{
    for (auto& entry : default_parse_sessions_)
    {
        if (!entry.second)
        {
            continue;
        }
        entry.second->work.CancelAndJoin();
    }
    default_parse_sessions_.clear();
}

void MainWindow::Impl::StopRegFileParseSessions()
{
    for (auto& entry : reg_file_parse_sessions_)
    {
        if (!entry.second)
        {
            continue;
        }
        entry.second->work.CancelAndJoin();
    }
    reg_file_parse_sessions_.clear();
}

std::wstring MainWindow::Impl::ActiveTracesPath() const
{
    std::wstring folder = util::GetAppDataFolder();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"active_traces.ini");
}

std::wstring MainWindow::Impl::ActiveDefaultsPath() const
{
    std::wstring folder = util::GetAppDataFolder();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"active_defaults.ini");
}

std::wstring MainWindow::Impl::TraceSettingsPath() const
{
    std::wstring folder = util::GetAppDataFolder();
    if (folder.empty())
    {
        return L"";
    }
    return util::JoinPath(folder, L"trace_settings.ini");
}

void MainWindow::Impl::LoadTraceSettings()
{
    trace_selection_cache_.clear();
    std::wstring path = TraceSettingsPath();
    if (path.empty())
    {
        return;
    }
    std::wstring content;
    std::wstring error;
    if (!util::ReadStateFile(path, &content, &error))
    {
        ReportFileError(path, error);
        return;
    }
    size_t line_number = 0;
    const auto unreadable = [&] {
        trace_selection_cache_.clear();
        ReportFileError(path, util::BadLineReason(line_number));
    };

    trace::Selection selection = {};
    selection.select_all = true;
    selection.recursive = true;
    std::wstring current_path;
    std::wstring current_label;
    bool has_entry = false;

    auto normalize_selection = [&]() {
        std::vector<std::wstring> cleaned;
        cleaned.reserve(selection.key_paths.size());
        std::unordered_set<std::wstring> seen;
        for (auto& path : selection.key_paths)
        {
            std::wstring trimmed = TrimWhitespace(path);
            if (trimmed.empty())
            {
                continue;
            }
            std::wstring lower = ToLower(trimmed);
            if (seen.insert(lower).second)
            {
                cleaned.push_back(std::move(trimmed));
            }
        }
        for (const auto& entry : selection.values_by_key)
        {
            if (entry.first.empty())
            {
                continue;
            }
            if (seen.insert(entry.first).second)
            {
                cleaned.push_back(entry.first);
            }
        }
        selection.key_paths.swap(cleaned);
        if (selection.key_paths.empty() && selection.values_by_key.empty())
        {
            selection.select_all = true;
        }
    };

    auto flush_entry = [&]() {
        if (!has_entry)
        {
            return;
        }
        normalize_selection();
        std::wstring key = current_path;
        if (key.empty() && !current_label.empty())
        {
            std::wstring resolved = ResolveBundledTracePath(current_label);
            key = resolved.empty() ? current_label : resolved;
        }
        key = TrimWhitespace(key);
        if (!key.empty())
        {
            trace_selection_cache_[ToLower(key)] = selection;
        }
        selection = {};
        selection.select_all = true;
        selection.recursive = true;
        current_path.clear();
        current_label.clear();
        has_entry = false;
    };

    for (const std::wstring_view raw : win32ui::LineViews(content))
    {
        ++line_number;
        if (raw.starts_with(L"pair="))
        {
            const std::vector<std::wstring> fields = record_fields::DecodeRecord(raw.substr(5));
            if (fields.empty() || fields[0].empty() || fields.size() > 2)
            {
                unreadable();
                return;
            }
            auto& values = selection.values_by_key[fields[0]];
            if (fields.size() == 2)
            {
                values.insert(fields[1]);
            }
            has_entry = true;
            continue;
        }
        const std::wstring line = TrimWhitespace(raw);
        if (line.empty())
        {
            flush_entry();
            continue;
        }
        if (line.front() == L'#')
        {
            continue;
        }
        if (line.front() == L'[')
        {
            flush_entry();
            continue;
        }
        size_t sep = line.find(L'=');
        std::wstring key = sep == std::wstring::npos ? std::wstring() : TrimWhitespace(line.substr(0, sep));
        if (key.empty())
        {
            unreadable();
            return;
        }
        std::wstring value = line.substr(sep + 1);
        has_entry = true;
        if (EqualsInsensitive(key, L"path"))
        {
            current_path = value;
        }
        else if (EqualsInsensitive(key, L"label"))
        {
            current_label = value;
        }
        else if (EqualsInsensitive(key, L"select_all"))
        {
            selection.select_all = util::ParseBool(value);
        }
        else if (EqualsInsensitive(key, L"recursive"))
        {
            selection.recursive = util::ParseBool(value);
        }
        else if (EqualsInsensitive(key, L"key"))
        {
            selection.key_paths.push_back(value);
        }
        else
        {
            unreadable();
            return;
        }
    }
    flush_entry();
}

void MainWindow::Impl::SaveTraceSettings() const
{
    std::wstring path = TraceSettingsPath();
    if (path.empty())
    {
        return;
    }
    std::wstring content;
    for (const auto& trace : active_traces_)
    {
        if (!trace.data || !trace.selection)
        {
            continue;
        }
        content.append(L"[trace]\n");
        if (!trace.label.empty())
        {
            content.append(L"label=");
            content.append(trace.label);
            content.push_back(L'\n');
        }
        if (!trace.source_path.empty())
        {
            content.append(L"path=");
            content.append(trace.source_path);
            content.push_back(L'\n');
        }
        content.append(L"select_all=");
        content.append(trace.selection->select_all ? L"1\n" : L"0\n");
        content.append(L"recursive=");
        content.append(trace.selection->recursive ? L"1\n" : L"0\n");
        for (const auto& key_path : trace.selection->key_paths)
        {
            if (key_path.empty())
            {
                continue;
            }
            content.append(L"key=");
            content.append(key_path);
            content.push_back(L'\n');
        }
        for (const auto& entry : trace.selection->values_by_key)
        {
            if (entry.first.empty())
            {
                continue;
            }
            if (entry.second.empty())
            {
                content.append(L"pair=");
                record_fields::AppendRecord(&content, {entry.first});
                continue;
            }
            for (const auto& value_name : entry.second)
            {
                content.append(L"pair=");
                record_fields::AppendRecord(&content, {entry.first, value_name});
            }
        }
        content.push_back(L'\n');
    }
    win32ui::WriteTextFileAtomic(path.c_str(), content);
}

void MainWindow::Impl::SaveActiveTraces() const
{
    const std::wstring path = ActiveTracesPath();
    if (!path.empty())
    {
        win32ui::WriteTextFileAtomic(path.c_str(), SerializeActiveList(active_traces_, [this](const std::wstring& label) { return ResolveBundledTracePath(label); }));
    }
}

void MainWindow::Impl::SaveActiveDefaults() const
{
    const std::wstring path = ActiveDefaultsPath();
    if (!path.empty())
    {
        win32ui::WriteTextFileAtomic(path.c_str(), SerializeActiveList(active_defaults_, [this](const std::wstring& label) { return ResolveBundledDefaultPath(label); }));
    }
}

bool MainWindow::Impl::HasActiveTraces() const
{
    return !active_traces_.empty();
}

bool MainWindow::Impl::RemoveTraceByPath(const std::wstring& path)
{
    if (path.empty())
    {
        return false;
    }
    std::wstring target = TrimWhitespace(path);
    if (target.empty())
    {
        return false;
    }
    std::wstring target_lower = ToLower(target);
    auto session_it = trace_parse_sessions_.find(target_lower);
    if (session_it != trace_parse_sessions_.end())
    {
        if (session_it->second)
        {
            session_it->second->work.CancelAndJoin();
        }
        trace_parse_sessions_.erase(session_it);
    }
    size_t removed = 0;
    active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(), [&](const ActiveTrace& trace) {
                             if (!EqualsInsensitive(trace.source_path, target))
                             {
                                 return false;
                             }
                             ++removed;
                             return true;
                         }),
                         active_traces_.end());
    if (removed == 0)
    {
        return false;
    }
    trace_selection_cache_.erase(target_lower);
    SaveActiveTraces();
    SaveTraceSettings();
    RefreshTreeSelection();
    UpdateValueListForNode(browse_.current_node());
    SaveSettings();
    return true;
}

bool MainWindow::Impl::RemoveTraceByLabel(const std::wstring& label)
{
    if (label.empty())
    {
        return false;
    }
    for (auto it = trace_parse_sessions_.begin(); it != trace_parse_sessions_.end();)
    {
        if (it->second && util::EqualsInsensitive(it->second->label, label))
        {
            it->second->work.CancelAndJoin();
            it = trace_parse_sessions_.erase(it);
            continue;
        }
        ++it;
    }
    size_t removed = 0;
    active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(), [&](const ActiveTrace& trace) {
                             if (!util::EqualsInsensitive(trace.label, label))
                             {
                                 return false;
                             }
                             ++removed;
                             return true;
                         }),
                         active_traces_.end());
    if (removed == 0)
    {
        return false;
    }
    trace_selection_cache_.clear();
    for (const auto& trace : active_traces_)
    {
        if (!trace.source_path.empty())
        {
            if (trace.selection)
            {
                trace_selection_cache_[ToLower(trace.source_path)] = *trace.selection;
            }
        }
    }
    SaveActiveTraces();
    SaveTraceSettings();
    RefreshTreeSelection();
    UpdateValueListForNode(browse_.current_node());
    SaveSettings();
    return true;
}

bool MainWindow::Impl::RemoveDefaultByPath(const std::wstring& path)
{
    if (path.empty())
    {
        return false;
    }
    std::wstring target = TrimWhitespace(path);
    if (target.empty())
    {
        return false;
    }
    std::wstring target_lower = ToLower(target);
    auto session_it = default_parse_sessions_.find(target_lower);
    if (session_it != default_parse_sessions_.end())
    {
        if (session_it->second)
        {
            session_it->second->work.CancelAndJoin();
        }
        default_parse_sessions_.erase(session_it);
    }
    size_t removed = 0;
    active_defaults_.erase(std::remove_if(active_defaults_.begin(), active_defaults_.end(), [&](const ActiveDefault& defaults) {
                               if (!EqualsInsensitive(defaults.source_path, target))
                               {
                                   return false;
                               }
                               ++removed;
                               return true;
                           }),
                           active_defaults_.end());
    if (removed == 0)
    {
        return false;
    }
    SaveActiveDefaults();
    UpdateValueListForNode(browse_.current_node());
    SaveSettings();
    return true;
}

}
