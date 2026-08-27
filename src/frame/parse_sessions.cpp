// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::MergeTraceEntries(
    TraceParseSession* session,
    const std::vector<KeyValueDialogEntry>& entries,
    std::unordered_set<std::wstring>* affected_keys) {
  if (!session || !session->data || entries.empty()) {
    return;
  }
  std::vector<trace::Entry> parsed;
  parsed.reserve(entries.size());
  for (const auto& entry : entries) {
    trace::Entry value;
    value.key_path = entry.key_path;
    value.display_path = entry.display_path;
    value.has_value = entry.has_value;
    value.value_name = entry.value_name;
    parsed.push_back(std::move(value));
  }
  trace::Merge(session->data.get(), parsed, affected_keys);
}

void MainWindow::Impl::MergeDefaultEntries(
    DefaultParseSession* session,
    const std::vector<KeyValueDialogEntry>& entries,
    std::unordered_set<std::wstring>* affected_keys) {
  if (!session || !session->data || entries.empty()) {
    return;
  }
  std::vector<defaults::Entry> parsed;
  parsed.reserve(entries.size());
  for (const auto& entry : entries) {
    defaults::Entry value;
    value.key_path = entry.key_path;
    value.has_value = entry.has_value;
    value.value_name = entry.value_name;
    value.type = entry.value_type;
    value.data = entry.value_data;
    parsed.push_back(std::move(value));
  }
  defaults::Merge(session->data.get(), parsed, [](const std::wstring& path) { return MapControlSetToCurrent(path); }, affected_keys);
}

void MainWindow::Impl::StartTraceParseThread(TraceParseSession* session) {
  if (!session || session->work.running()) {
    return;
  }
  HWND hwnd = hwnd_;
  std::wstring source = session->source_path;
  std::wstring source_lower = session->source_lower;
  session->work.Start(
      [this, session, hwnd, source, source_lower](
          uint64_t generation, std::atomic_bool& cancel) {
        constexpr size_t kBatchSize = 256;
        constexpr DWORD kBatchMs = 50;
        auto post_batch = [&](std::vector<KeyValueDialogEntry>* entries, bool done, const std::wstring& error, bool cancelled) {
          auto payload = std::make_unique<TraceParseBatch>();
          payload->generation = generation;
          payload->source_lower = source_lower;
          if (entries) {
            MergeTraceEntries(session, *entries, &payload->affected_keys);
            payload->entries = std::move(*entries);
          }
          payload->done = done;
          payload->error = error;
          payload->cancelled = cancelled;
          if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, frame::message_id::kTraceParseBatch, 0, reinterpret_cast<LPARAM>(payload.get()))) {
            return;
          }
          ReleasePostedPayload(payload);
        };

        std::vector<KeyValueDialogEntry> entries;
        entries.reserve(kBatchSize);
        uint64_t last_post = GetTickCount64();
        std::wstring parse_error;
        const bool parsed = trace::LoadEntries(
            source, TraceNormalizers(),
            [&](trace::Entry&& parsed_entry) {
              KeyValueDialogEntry entry;
              entry.key_path = std::move(parsed_entry.key_path);
              entry.display_path = std::move(parsed_entry.display_path);
              entry.has_value = true;
              entry.value_name = std::move(parsed_entry.value_name);
              entries.push_back(std::move(entry));
              const uint64_t now = GetTickCount64();
              if (entries.size() >= kBatchSize ||
                  now - last_post >= kBatchMs) {
                post_batch(&entries, false, L"", false);
                entries.clear();
                last_post = now;
              }
              return !cancel.load();
            },
            &parse_error, &cancel);
        if (!parsed) {
          post_batch(nullptr, true, cancel.load() ? L"" : parse_error,
                     cancel.load());
          return;
        }
        if (!entries.empty()) {
          post_batch(&entries, false, L"", false);
          entries.clear();
        }
        trace::Sort(session->data.get());
        post_batch(nullptr, true, L"", false);
      });
}

void MainWindow::Impl::StartDefaultParseThread(DefaultParseSession* session) {
  if (!session || session->work.running()) {
    return;
  }
  HWND hwnd = hwnd_;
  std::wstring source = session->source_path;
  std::wstring source_lower = session->source_lower;
  session->work.Start(
      [this, session, hwnd, source, source_lower](
          uint64_t generation, std::atomic_bool& cancel) {
        constexpr size_t kBatchSize = 256;
        constexpr DWORD kBatchMs = 50;
        auto post_batch = [&](std::vector<KeyValueDialogEntry>* entries, bool done, const std::wstring& error, bool cancelled) {
          auto payload = std::make_unique<DefaultParseBatch>();
          payload->generation = generation;
          payload->source_lower = source_lower;
          if (entries) {
            MergeDefaultEntries(session, *entries, &payload->affected_keys);
            payload->entries = std::move(*entries);
          }
          payload->done = done;
          payload->error = error;
          payload->cancelled = cancelled;
          if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, frame::message_id::kDefaultParseBatch, 0, reinterpret_cast<LPARAM>(payload.get()))) {
            return;
          }
          ReleasePostedPayload(payload);
        };

        std::vector<defaults::Entry> parsed_entries;
        std::wstring parse_error;
        if (!defaults::Load(
                source,
                [](const std::wstring& path) {
                  return NormalizeTraceKeyPathBasic(path);
                },
                nullptr, &parsed_entries, &parse_error, &cancel)) {
          post_batch(nullptr, true, cancel.load() ? L"" : parse_error,
                     cancel.load());
          return;
        }

        std::vector<KeyValueDialogEntry> entries;
        entries.reserve(kBatchSize);
        uint64_t last_post = GetTickCount64();
        auto flush_if_needed = [&] {
          const uint64_t now = GetTickCount64();
          if (entries.size() >= kBatchSize || now - last_post >= kBatchMs) {
            post_batch(&entries, false, L"", false);
            entries.clear();
            last_post = now;
          }
        };

        for (auto& parsed : parsed_entries) {
          if (cancel.load()) {
            post_batch(nullptr, true, L"", true);
            return;
          }
          KeyValueDialogEntry entry;
          entry.key_path = std::move(parsed.key_path);
          entry.display_path =
              NormalizeTraceSelectionPath(parsed.source_path);
          if (entry.display_path.empty()) {
            entry.display_path = entry.key_path;
          }
          entry.has_value = parsed.has_value;
          entry.value_name = std::move(parsed.value_name);
          entry.value_type = parsed.type;
          entry.value_data = std::move(parsed.data);
          entries.push_back(std::move(entry));
          flush_if_needed();
        }

        if (!entries.empty()) {
          post_batch(&entries, false, L"", false);
        }
        post_batch(nullptr, true, L"", false);
      });
}

} // namespace regkit
