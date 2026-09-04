// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

namespace regkit::frame::message_id {

constexpr UINT kAddressEnter = WM_APP + 10;
constexpr UINT kFocusAddressBar = WM_APP + 11;
constexpr UINT kSearchResults = WM_APP + 20;
constexpr UINT kSearchFailed = WM_APP + 22;
constexpr UINT kSearchProgress = WM_APP + 23;
constexpr UINT kLoadTraces = WM_APP + 24;
constexpr UINT kTraceLoadReady = WM_APP + 25;
constexpr UINT kLoadDefaults = WM_APP + 26;
constexpr UINT kDefaultLoadReady = WM_APP + 27;
constexpr UINT kValueListReady = WM_APP + 30;
constexpr UINT kTraceParseBatch = WM_APP + 31;
constexpr UINT kDefaultParseBatch = WM_APP + 32;
constexpr UINT kRegFileLoadReady = WM_APP + 33;
constexpr UINT kDeferredStartup = WM_APP + 34;
constexpr UINT kStartupCacheReady = WM_APP + 35;
constexpr UINT kReplaceReady = WM_APP + 36;
constexpr UINT kValuePreviewReady = WM_APP + 37;
constexpr UINT kValuePreviewRequest = WM_APP + 38;
constexpr UINT kSearchPreviewRequest = WM_APP + 39;
constexpr UINT kSearchPreviewReady = WM_APP + 40;
constexpr UINT kSearchSortReady = WM_APP + 41;
constexpr UINT kSearchTabLoadReady = WM_APP + 42;

} // namespace regkit::frame::message_id
