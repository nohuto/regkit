// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::StartTraceLoadWorker() {
  if (trace_load_session_.running()) {
    return;
  }
  LoadTraceSettings();
  std::unordered_map<std::wstring, trace::Selection> selection_cache =
      trace_selection_cache_;
  std::wstring active_path = ActiveTracesPath();
  const HWND hwnd = hwnd_;
  trace_load_session_.StartIfIdle(
      [this, selection_cache = std::move(selection_cache), active_path, hwnd](
          uint64_t generation, const std::atomic_bool& cancel) mutable {
        auto payload = std::make_unique<TraceLoadPayload>();
        payload->generation = generation;
        payload->selection_cache = std::move(selection_cache);
        std::wstring content;
        if (!util::ReadTextFile(
                active_path, &content, nullptr,
                static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
          return;
        }

        std::vector<std::wstring> entries;
        size_t start = 0;
        while (start < content.size()) {
          size_t end = content.find(L'\n', start);
          if (end == std::wstring::npos) {
            end = content.size();
          }
          std::wstring line = content.substr(start, end - start);
          if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
          }
          start = end + 1;
          line = TrimWhitespace(line);
          if (line.empty() || line.front() == L'#') {
            continue;
          }
          if (StartsWithInsensitive(line, L"trace=")) {
            line.erase(0, wcslen(L"trace="));
          }
          line = TrimWhitespace(line);
          if (!line.empty()) {
            entries.push_back(std::move(line));
          }
        }

        std::unordered_set<std::wstring> loaded;
        for (const auto& entry : entries) {
          if (cancel.load()) {
            return;
          }
          std::wstring source = entry;
          std::wstring use_label;
          if (!FileExists(source)) {
            std::wstring bundled = ResolveBundledTracePath(source);
            if (!bundled.empty() && FileExists(bundled)) {
              source = bundled;
              use_label = entry;
            } else {
              continue;
            }
          }
          if (use_label.empty()) {
            use_label = FileBaseName(source);
          }
          if (use_label.empty()) {
            use_label = L"Trace";
          }
          std::wstring source_lower = ToLower(source);
          if (!loaded.insert(source_lower).second) {
            continue;
          }
          trace::Data data;
          if (!trace::Load(use_label, source, TraceNormalizers(), &data,
                           nullptr, &cancel)) {
            continue;
          }
          std::shared_ptr<const trace::Data> trace_data =
              std::make_shared<trace::Data>(std::move(data));
          trace::Selection selection = {};
          selection.select_all = true;
          selection.recursive = true;
          auto it = payload->selection_cache.find(source_lower);
          if (it != payload->selection_cache.end()) {
            selection = it->second;
          }
          trace::NormalizeSelection(*trace_data, &selection);
          payload->selection_cache[source_lower] = selection;
          payload->traces.push_back(
              {trace_data->label, source, trace_data,
               std::make_shared<trace::Selection>(selection)});
        }

        if (cancel.load()) {
          return;
        }
        if (hwnd && IsWindow(hwnd) &&
            PostMessageW(hwnd, frame::message_id::kTraceLoadReady, 0,
                         reinterpret_cast<LPARAM>(payload.get()))) {
          ReleasePostedPayload(payload);
        }
      });
}

void MainWindow::Impl::StopTraceLoadWorker() {
  trace_load_session_.CancelAndJoin();
}

void MainWindow::Impl::StopTraceParseSessions() {
  for (auto& entry : trace_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->work.CancelAndJoin();
  }
  trace_parse_sessions_.clear();
}

} // namespace regkit
