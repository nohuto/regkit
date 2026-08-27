// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

namespace regkit::frame {

enum class MessageArea {
  kUnknown,
  kLifecycle,
  kLayoutInput,
  kWorker,
  kExternal,
  kAppearance,
  kBrowse,
};

MessageArea ClassifyMessage(UINT message) noexcept;

} // namespace regkit::frame
