// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include "frame/window_file_detail.h"

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

std::wstring ResolveDevicePath(const std::wstring& path);

std::wstring NormalizeHiveFilePath(const std::wstring& raw_path);

std::wstring CurrentControlSetSegment();

std::wstring ReplaceControlSetSegment(const std::wstring& path, const std::wstring& from, const std::wstring& to);

std::wstring NormalizeCurrentControlSet(const std::wstring& path);

bool IsControlSetSegment(const std::wstring& text);

std::wstring MapControlSetToCurrent(const std::wstring& path);

std::wstring CleanTraceKeyText(const std::wstring& text, const std::wstring& sid);

std::wstring NormalizeTraceKeyPathBasic(const std::wstring& text);

std::wstring NormalizeTraceKeyPath(const std::wstring& text);

std::wstring NormalizeTraceSelectionPath(const std::wstring& text);

trace::Normalizers TraceNormalizers();

struct LinkTargetCache
{
    std::mutex mutex;
    std::unordered_map<std::wstring, std::wstring> targets;
    std::unordered_set<std::wstring> misses;
};

LinkTargetCache& GetLinkTargetCache();

bool ParseRegistryRoot(const std::wstring& input, RegistryNode* node, std::wstring* root_label);

bool QueryLinkTargetCached(const std::wstring& path, const RegistryNode& node, std::wstring* target);

std::wstring ResolveRegistryLinkPath(const std::wstring& path);

std::wstring FileNameOnly(const std::wstring& path);

std::vector<std::wstring> SplitLabelWords(const std::wstring& text);

bool IsReleaseWord(const std::wstring& word);

std::wstring ShortWindowsName(const std::wstring& folder);

std::wstring ShortDefaultLabel(const std::wstring& label, const std::wstring& source_path);

std::wstring FileBaseName(const std::wstring& path);

struct OfflineHiveCandidate
{
    std::wstring path;
    std::wstring label;
};

bool IsFilePath(const std::wstring& path);

void AddOfflineHiveCandidate(std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen, const std::wstring& path, const std::wstring& label);

std::wstring TopLevelFolderLabel(const std::wstring& base, const std::wstring& folder);

void CollectUserHiveCandidates(const std::wstring& folder, const std::wstring& base, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen);

void CollectUserHivesRecursive(const std::wstring& folder, const std::wstring& base, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen);

bool ShouldIncludeOfflineHiveFile(const std::wstring& name);

void CollectLooseHivesInFolder(const std::wstring& folder, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen);

void CollectOfflineHivesInFolder(const std::wstring& folder, std::vector<OfflineHiveCandidate>* out);

std::wstring ResolveOfflineRootName(const std::wstring& path, bool is_dir, const RegistryNode* current_node);

} // namespace
