// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include "frame/window_draw_detail.h"

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

#include "appearance/feedback.h"
#include "appearance/gdi_cache.h"
#include "appearance/icon_loader.h"
#include "defaults/default_loader.h"
#include "editors/comment_editor.h"
#include "editors/value_editor.h"
#include "frame/command_ids.h"
#include "frame/message_dispatch.h"
#include "frame/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/security_dialog.h"
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

std::wstring TrimTrailingSeparators(const std::wstring& path);

bool IsDirectoryPath(const std::wstring& path);

constexpr wchar_t kIconSetPhosphor[] = L"phosphor";
constexpr wchar_t kIconSetClassic[] = L"classic";
constexpr wchar_t kIconSetCustom[] = L"custom";

bool IsIconSetName(const std::wstring& value, const wchar_t* name);

bool IsKnownIconSetName(const std::wstring& value);

std::wstring FindAssetsIconsRoot();

std::wstring AssetsIconsRoot();

constexpr wchar_t kOfflineHiveFilter[] =
    L"Registry Hive Files\0*.dat;*.hiv;*.hive;*.sav;SYSTEM;SOFTWARE;SAM;SECURITY;DEFAULT;NTUSER.DAT;USRCLASS.DAT\0All "
    L"Files (*.*)\0*.*\0";

std::wstring NormalizeMachineName(const std::wstring& text);

std::wstring StripMachinePrefix(const std::wstring& machine);

bool FileExists(const std::wstring& path);

using util::EqualsInsensitive;

using util::StartsWithInsensitive;

bool WindowClassEquals(HWND hwnd, const wchar_t* class_name);

struct ParsedRegFileRoot
{
    std::wstring name;
    std::shared_ptr<VirtualRegistryData> data;
};

struct RegFileParsePayload : work::MoveOnly
{
    uint64_t generation = 0;
    std::wstring source_path;
    std::wstring source_lower;
    std::vector<ParsedRegFileRoot> roots;
    std::wstring error;
    bool cancelled = false;
};

VirtualRegistryKey* EnsureVirtualKey(VirtualRegistryKey* root, const std::wstring& subkey);

bool ParseRegFileToVirtualRoots(const std::wstring& path, std::vector<ParsedRegFileRoot>* roots, std::wstring* error, const std::atomic_bool* cancel, bool* cancelled);

} // namespace
