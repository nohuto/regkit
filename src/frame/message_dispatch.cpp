// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/message_dispatch.h"

#include "frame/message_ids.h"

namespace regkit::frame {

MessageArea ClassifyMessage(UINT message) noexcept {
  switch (message) {
  case message_id::kSearchResults:
  case message_id::kSearchFailed:
  case message_id::kSearchProgress:
  case message_id::kSearchPreviewRequest:
  case message_id::kSearchPreviewReady:
  case message_id::kSearchSortReady:
  case message_id::kSearchTabLoadReady:
  case message_id::kLoadTraces:
  case message_id::kTraceLoadReady:
  case message_id::kLoadDefaults:
  case message_id::kDefaultLoadReady:
  case message_id::kValueListReady:
  case message_id::kValuePreviewReady:
  case message_id::kValuePreviewRequest:
  case message_id::kTraceParseBatch:
  case message_id::kDefaultParseBatch:
  case message_id::kRegFileLoadReady:
  case message_id::kDeferredStartup:
  case message_id::kStartupCacheReady:
  case message_id::kReplaceReady:
    return MessageArea::kWorker;
  case message_id::kAddressEnter:
  case message_id::kFocusAddressBar:
  case WM_DROPFILES:
  case WM_TIMER:
  case WM_COPYDATA:
  case WM_SETFOCUS:
    return MessageArea::kExternal;
  case WM_CREATE:
  case WM_DESTROY:
  case WM_NCPAINT:
  case WM_NCACTIVATE:
  case WM_GETMINMAXINFO:
  case WM_SIZE:
  case WM_DPICHANGED:
  case WM_DPICHANGED_AFTERPARENT:
  case WM_ACTIVATE:
  case WM_CLOSE:
    return MessageArea::kLifecycle;
  case WM_LBUTTONDOWN:
  case WM_LBUTTONUP:
  case WM_MOUSEMOVE:
  case WM_CAPTURECHANGED:
  case WM_SETCURSOR:
    return MessageArea::kLayoutInput;
  case WM_ERASEBKGND:
  case WM_PAINT:
  case WM_SETTINGCHANGE:
  case WM_THEMECHANGED:
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLOREDIT:
  case WM_INITMENUPOPUP:
  case WM_DRAWITEM:
  case WM_MEASUREITEM:
    return MessageArea::kAppearance;
  case WM_COMMAND:
  case WM_CONTEXTMENU:
  case WM_NOTIFY:
    return MessageArea::kBrowse;
  default:
    return MessageArea::kUnknown;
  }
}

} // namespace regkit::frame
