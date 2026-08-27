// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <windows.h>

#include <string>

namespace regkit::editors {

struct CommentRequest {
  std::wstring text;
  bool apply_to_same_name = false;
};

struct CommentResult {
  std::wstring text;
  bool apply_to_same_name = false;
};

bool EditComment(HWND owner, const CommentRequest& request,
                 CommentResult* result);

} // namespace regkit::editors
