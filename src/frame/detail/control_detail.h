// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/gdi/fonts.h>
#include <win32ui/gdi/gdi_cache.h>
#include <win32ui/gdi/icon_loader.h>
#include <win32ui/list_view/list_view.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include "frame/detail/file_detail.h"

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

#include <commdlg.h>
#include <pathcch.h>
#include <richedit.h>
#include <shellapi.h>
#include <shldisp.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <windowsx.h>
#include <winternl.h>

#include "defaults/default_loader.h"
#include "dialogs/comment_editor.h"
#include "dialogs/security_dialog.h"
#include "dialogs/value_editor.h"
#include "frame/commands/command_ids.h"
#include "frame/window/message_dispatch.h"
#include "frame/window/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/value_format.h"
#include "resource.h"
#include "search/result_file.h"
#include "trace/trace_loader.h"
#include "trace/trace_parser.h"
#include "ui/feedback.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "workspace/settings.h"
#include "workspace/tab_state.h"
#include "registry/hive_files.h"

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

#include <commdlg.h>
#include <pathcch.h>
#include <richedit.h>
#include <shellapi.h>
#include <shldisp.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <windowsx.h>
#include <winternl.h>

#include "ui/feedback.h"
#include "frame/commands/command_ids.h"
#include "browse/browse_pane.h"
#include "browse/value_table.h"
#include "registry/registry_store.h"
#include "dialogs/security_dialog.h"

#include "defaults/default_loader.h"
#include "dialogs/comment_editor.h"
#include "dialogs/value_editor.h"
#include "frame/window/message_dispatch.h"
#include "frame/window/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "resource.h"
#include "search/result_file.h"
#include "trace/trace_loader.h"
#include "trace/trace_parser.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "workspace/settings.h"
#include "workspace/tab_state.h"

namespace regkit::window_detail
{

using registry_path::ChildNode;

std::wstring LeafName(const RegistryNode& node);

bool UseBinaryValueIcon(DWORD type);

ListRow MakeValueListRow(const std::wstring& name, DWORD type, const BYTE* data, DWORD data_size, bool preview = false);

void UpdateLeafName(RegistryNode* node, const std::wstring& new_name);

std::wstring FormatFileTime(const FILETIME& filetime);

std::wstring FormatCommentDisplay(const std::wstring& text);

constexpr size_t kCellTextDrawLimit = 512;
constexpr size_t kValuePreviewLimit = 4096;
constexpr DWORD kValuePreviewBytes = 4096;

const std::wstring& ValueRowFieldText(const ListRow& row, int subitem);

void SortValueRows(std::vector<ListRow>* rows, int column, bool ascending, const util::SortKeys& earlier);

HFONT CreateUIFont();

void ApplyFont(HWND hwnd, HFONT font);

HTREEITEM FindChildByText(HWND tree, HTREEITEM parent, const std::wstring& text);

}
