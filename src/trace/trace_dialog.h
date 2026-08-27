// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"
#include "trace/trace_data.h"

#include <windows.h>

#include <string>
#include <vector>

namespace regkit {

struct KeyValueDialogEntry {
  std::wstring key_path;
  std::wstring display_path;
  bool has_value = false;
  std::wstring value_name;
  DWORD value_type = 0;
  std::wstring value_data;
};

struct TraceDialogOptions {
  std::wstring title;
  std::wstring prompt;
  bool show_values = true;
};

using TraceDialogReadyCallback = void (*)(HWND hwnd, void* context);

bool ShowTraceDialog(HWND owner, const TraceDialogOptions& options,
                     trace::Selection* selection,
                     TraceDialogReadyCallback on_ready = nullptr,
                     void* context = nullptr);

void TraceDialogPostEntries(HWND dialog, std::vector<KeyValueDialogEntry>* entries);
void TraceDialogPostDone(HWND dialog, bool done);

} // namespace regkit
