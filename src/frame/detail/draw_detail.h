// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/gdi/fonts.h>
#include <win32ui/gdi/gdi_cache.h>
#include <win32ui/gdi/icon_loader.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/worker.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include <algorithm>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <exception>
#include <functional>
#include <limits>
#include <string_view>

#include <pathcch.h>
#include <richedit.h>
#include <shellapi.h>
#include <shldisp.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <vssym32.h>
#include <windowsx.h>
#include <winternl.h>

#include "ui/feedback.h"
#include "defaults/default_loader.h"
#include "dialogs/comment_editor.h"
#include "dialogs/value_editor.h"
#include "frame/commands/command_ids.h"
#include "browse/value_table.h"
#include "dialogs/trace_dialog.h"
#include "frame/window/message_dispatch.h"
#include "frame/window/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "dialogs/security_dialog.h"
#include "registry/value_format.h"
#include "resource.h"
#include "search/result_file.h"
#include "trace/trace_loader.h"
#include "trace/trace_parser.h"
#include "win32/file_dialog.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/restart.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "workspace/settings.h"
#include "workspace/tab_state.h"

namespace regkit::window_detail
{

constexpr int kToolbarId = 100;
constexpr int kAddressEditId = 0;
constexpr int kTreeId = 1;
constexpr int kValueListId = 2;
constexpr int kTabId = 103;
constexpr int kHistoryListId = 106;
constexpr int kHistoryLabelId = 107;
constexpr int kTreeHeaderId = 108;
constexpr int kAddressGoId = 109;
constexpr int kTreeHeaderCloseId = 110;
constexpr int kStatusBarId = 111;
constexpr int kSearchResultsListId = 112;
constexpr int kHistoryHeaderCloseId = 114;
constexpr int kFilterEditId = 115;
constexpr int kFilterClearId = 118;
constexpr int kRegEditCompatTreeId = 119;
constexpr int kTreeSplitterId = 120;
constexpr int kHistorySplitterId = 121;
constexpr UINT_PTR kTreeStateTimerId = 42;
constexpr UINT_PTR kCompatJumpTimerId = 43;
constexpr UINT_PTR kAutoRefreshTimerId = 44;
constexpr UINT kCompatJumpDelayMs = 40;
using win32::kRestartAdminArg;
using win32::kRestartSystemArg;
using win32::kRestartTiArg;
using win32::kRestartUserArg;
template <typename T>
inline T ClampValue(T value, T low, T high)
{
    return value < low ? low : (high < value ? high : value);
}

constexpr wchar_t kRootKeysGroupLabel[] = L"Root Keys";
constexpr wchar_t kRealGroupLabel[] = L"REGISTRY";

constexpr DWORD kSearchResultsMaxMs = 15;
constexpr DWORD kSearchResultsRefreshMs = 1000;
constexpr DWORD kSearchProgressUiMs = 500;
constexpr size_t kSearchQueueBatch = 128;
constexpr size_t kSearchPendingRowLimit = 8192;
using frame::message_id::kEditRegFileCopyDataId;
using frame::message_id::kExternalJumpCopyDataId;
using frame::message_id::kExternalMessageMaxBytes;
constexpr UINT_PTR kAddressSubclassId = 1;
constexpr UINT_PTR kListViewSubclassId = 4;
constexpr UINT_PTR kTreeViewSubclassId = 5;
constexpr UINT_PTR kFilterSubclassId = 8;
constexpr UINT_PTR kCloseButtonSubclassId = 9;
constexpr wchar_t kMainWindowClassName[] = L"RegEdit_RegEdit";
using frame::message_id::kRegKitWindowProperty;
using util::ToLower;
using util::TrimWhitespace;
using win32ui::FormatWin32Error;

constexpr int kFolderIconIndex = 0;
constexpr int kSymlinkIconIndex = 1;
constexpr int kDatabaseIconIndex = 2;
constexpr int kFolderSimIconIndex = 3;
constexpr int kFolderDeniedIconIndex = 4;
constexpr int kDatabaseDeniedIconIndex = 5;
constexpr int kFolderVolatileIconIndex = 6;
constexpr int kValueIconIndex = 7;
constexpr int kBinaryIconIndex = 8;
constexpr int kTraceIconIndex = 9;
constexpr int kLocalRegistryIconIndex = 9;
constexpr int kHeaderTextPadding = 6;
constexpr int kSplitterWidth = 6;
constexpr int kHistorySplitterHeight = 4;
constexpr int kMinTreeWidth = 160;
constexpr int kMinValueListWidth = 240;
constexpr int kMinHistoryHeight = 80;
constexpr int kHistoryMaxPadding = 140;
constexpr int kHistoryGap = 2;
constexpr int kMainVerticalGap = 6;
constexpr int kValueColName = 0;
constexpr int kValueColType = 1;
constexpr int kValueColData = 2;
constexpr int kValueColDefault = 3;
constexpr int kValueColReadOnBoot = 4;
constexpr int kValueColSize = 5;
constexpr int kValueColDate = 6;
constexpr int kValueColDetails = 7;
constexpr int kValueColComment = 8;

struct ParseBatch : win32ui::PostedPayload
{
    uint64_t generation = 0;
    std::wstring source_lower;
    std::vector<KeyValueDialogEntry> entries;
    std::unordered_set<std::wstring> affected_keys;
    std::wstring error;
    bool done = false;
};

struct ValueListPayload : win32ui::PostedPayload
{
    uint64_t generation = 0;
    std::vector<ListRow> rows;
    int key_count = 0;
    int value_count = 0;
    LONG error = ERROR_SUCCESS; // reason key values could not be listed
};

constexpr int kCellTextPadding = 6;
constexpr int kLabelTextInset = 2;
constexpr int kPanelHeaderHeight = 22;
constexpr int kPanelCloseSize = 16;
constexpr int kPanelCloseInset = 2;
constexpr int kPanelBorderOverlap = 1;

int SearchMatchSubItem(const search::Result& result);

}
