// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#pragma once

#include "../win32/windows_config.h"
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
