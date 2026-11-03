// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/gdi/gdi_cache.h>
#include <win32ui/gdi/icon_loader.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include "frame/detail/draw_detail.h"

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

namespace regkit::window_detail
{

constexpr wchar_t kIconSetPhosphor[] = L"phosphor";
constexpr wchar_t kIconSetClassic[] = L"classic";
constexpr wchar_t kIconSetCustom[] = L"custom";

bool IsKnownIconSetName(const std::wstring& value);

std::wstring AssetsIconsRoot();

constexpr wchar_t kOfflineHiveFilter[] =
    L"Registry Hive Files\0*.dat;*.hiv;*.hive;*.sav;SYSTEM;SOFTWARE;SAM;SECURITY;DEFAULT;NTUSER.DAT;USRCLASS.DAT\0All "
    L"Files (*.*)\0*.*\0";

std::wstring NormalizeMachineName(const std::wstring& text);

std::wstring StripMachinePrefix(const std::wstring& machine);

bool ReadActiveEntries(const std::wstring& path, std::vector<std::wstring>* entries, std::wstring* error);
bool ResolveActiveSource(const std::wstring& entry, const std::function<std::wstring(const std::wstring&)>& resolve_bundled, const wchar_t* fallback_label, std::wstring* source, std::wstring* label);

using util::EqualsInsensitive;

using util::StartsWithInsensitive;

struct RegFileParsePayload : win32ui::PostedPayload
{
    uint64_t generation = 0;
    std::wstring source_path;
    std::wstring source_lower;
    std::vector<regfile::VirtualRoot> roots;
    std::wstring error;
    bool cancelled = false;
};

}
