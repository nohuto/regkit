// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include <string>

namespace regkit::editors {

struct ExportRequest {
  std::wstring path;
  bool include_subkeys = true;
  bool open_after = false;
};

using ExportResult = ExportRequest;

bool ChooseExport(HWND owner, const ExportRequest& request,
                  ExportResult* result);

} // namespace regkit::editors
