// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "trace/trace_paths.h"

namespace regkit
{
using namespace window_detail;

void MainWindow::Impl::StartTraceLoadWorker()
{
    if (trace_load_session_.running())
    {
        return;
    }
    LoadTraceSettings();
    std::unordered_map<std::wstring, trace::Selection> selection_cache = trace_selection_cache_;
    const std::wstring active_path = ActiveTracesPath();
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
    trace_load_session_.StartOrRun(L"TraceLoadThread", [this, selection_cache = std::move(selection_cache), entries = std::move(entries), hwnd](uint64_t generation, const std::atomic_bool& cancel) mutable {
        auto payload = std::make_unique<TraceLoadPayload>();
        payload->generation = generation;
        payload->selection_cache = std::move(selection_cache);
        std::unordered_set<std::wstring> loaded;
        for (const auto& entry : entries)
        {
            if (cancel.load())
            {
                return;
            }
            std::wstring source;
            std::wstring use_label;
            if (!ResolveActiveSource(entry, [this](const std::wstring& label) { return ResolveBundledTracePath(label); }, L"Trace", &source, &use_label))
            {
                continue;
            }
            std::wstring source_lower = ToLower(source);
            if (!loaded.insert(source_lower).second)
            {
                continue;
            }
            trace::Data data;
            if (!trace::Load(use_label, source, trace::PathNormalizers(), &data, nullptr, &cancel))
            {
                continue;
            }
            std::shared_ptr<const trace::Data> trace_data = std::make_shared<trace::Data>(std::move(data));
            trace::Selection selection = {};
            selection.select_all = true;
            selection.recursive = true;
            auto it = payload->selection_cache.find(source_lower);
            if (it != payload->selection_cache.end())
            {
                selection = it->second;
            }
            trace::NormalizeSelection(*trace_data, &selection);
            payload->selection_cache[source_lower] = selection;
            payload->traces.push_back(
                {trace_data->label, source, trace_data, std::make_shared<trace::Selection>(selection)}
            );
        }

        if (cancel.load())
        {
            return;
        }
        if (hwnd && IsWindow(hwnd))
        {
            win32ui::PostPayload(hwnd, frame::message_id::kTraceLoadReady, 0, payload);
        }
    });
}

void MainWindow::Impl::StopTraceLoadWorker()
{
    trace_load_session_.CancelAndJoin();
}

void MainWindow::Impl::StopTraceParseSessions()
{
    for (auto& entry : trace_parse_sessions_)
    {
        if (!entry.second)
        {
            continue;
        }
        entry.second->work.CancelAndJoin();
    }
    trace_parse_sessions_.clear();
}

}
