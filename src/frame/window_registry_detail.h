// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/window_file_detail.h"

#include "frame/window_impl.h"

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
#include <regex>

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

#include "frame/command_ids.h"
#include "registry/security_dialog.h"
#include "appearance/feedback.h"
#include "appearance/gdi_cache.h"
#include "registry/registry_store.h"
#include "appearance/icon_loader.h"
#include "win32/text_transform.h"
#include "defaults/default_loader.h"
#include "editors/comment_editor.h"
#include "editors/value_editor.h"
#include "frame/message_dispatch.h"
#include "frame/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "search/result_file.h"
#include "trace/trace_loader.h"
#include "trace/trace_parser.h"
#include "workspace/settings.h"
#include "workspace/tab_state.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/shell_paths.h"
#include "resource.h"

namespace regkit::window_detail {
inline std::wstring ResolveDevicePath(const std::wstring& path) {
  if (!StartsWithInsensitive(path, L"\\Device\\")) {
    return path;
  }
  wchar_t drives[512] = {};
  DWORD drive_len = GetLogicalDriveStringsW(static_cast<DWORD>(_countof(drives) - 1), drives);
  if (drive_len == 0 || drive_len >= _countof(drives)) {
    return path;
  }
  for (const wchar_t* drive = drives; *drive; drive += wcslen(drive) + 1) {
    wchar_t device[MAX_PATH] = {};
    wchar_t drive_root[4] = {};
    wcsncpy_s(drive_root, drive, _TRUNCATE);
    size_t root_len = wcslen(drive_root);
    if (root_len >= 2 && drive_root[1] == L':') {
      drive_root[2] = L'\0';
    }
    if (!QueryDosDeviceW(drive_root, device, static_cast<DWORD>(_countof(device)))) {
      continue;
    }
    size_t device_len = wcslen(device);
    if (_wcsnicmp(path.c_str(), device, device_len) != 0) {
      continue;
    }
    std::wstring rest = path.substr(device_len);
    if (!rest.empty() && rest.front() != L'\\') {
      rest.insert(rest.begin(), L'\\');
    }
    std::wstring mapped = drive_root;
    mapped += rest;
    return mapped;
  }
  return path;
}

inline std::wstring NormalizeHiveFilePath(const std::wstring& raw_path) {
  if (raw_path.empty()) {
    return raw_path;
  }
  std::wstring path = raw_path;
  if (StartsWithInsensitive(path, L"\\??\\")) {
    path.erase(0, 4);
  } else if (StartsWithInsensitive(path, L"\\\\?\\")) {
    path.erase(0, 4);
  } else if (StartsWithInsensitive(path, L"\\DosDevices\\")) {
    path.erase(0, wcslen(L"\\DosDevices\\"));
  }
  if (StartsWithInsensitive(path, L"\\SystemRoot")) {
    wchar_t windows_dir[MAX_PATH] = {};
    UINT len = GetWindowsDirectoryW(windows_dir, _countof(windows_dir));
    if (len > 0 && len < _countof(windows_dir)) {
      std::wstring suffix = path.substr(wcslen(L"\\SystemRoot"));
      path = std::wstring(windows_dir) + suffix;
    }
  }
  std::wstring expanded = util::ExpandEnvironmentStringsDynamic(path);
  if (!expanded.empty()) {
    path = std::move(expanded);
  }
  path = ResolveDevicePath(path);
  return path;
}

inline std::wstring CurrentControlSetSegment() {
  static std::wstring cached;
  static bool loaded = false;
  if (loaded) {
    return cached;
  }
  loaded = true;
  RegistryNode node;
  node.root = HKEY_LOCAL_MACHINE;
  node.subkey = L"SYSTEM\\Select";
  ValueEntry entry;
  if (RegistryStore::QueryValue(node, L"Current", &entry) && entry.type == REG_DWORD && entry.data.size() >= sizeof(DWORD)) {
    DWORD current = 0;
    std::memcpy(&current, entry.data.data(), sizeof(DWORD));
    wchar_t buffer[32] = {};
    swprintf_s(buffer, L"ControlSet%03u", current);
    cached = buffer;
  }
  return cached;
}

inline std::wstring ReplaceControlSetSegment(const std::wstring& path, const std::wstring& from, const std::wstring& to) {
  if (path.empty() || from.empty() || to.empty()) {
    return L"";
  }
  std::vector<std::wstring> parts = registry_path::Split(path);
  if (parts.size() < 3) {
    return L"";
  }
  bool is_hklm = EqualsInsensitive(parts[0], L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(parts[0], L"HKLM");
  if (!is_hklm && parts.size() > 1) {
    if (EqualsInsensitive(parts[0], L"REGISTRY") && EqualsInsensitive(parts[1], L"MACHINE")) {
      is_hklm = true;
    }
  }
  if (!is_hklm) {
    return L"";
  }
  for (size_t i = 0; i + 1 < parts.size(); ++i) {
    if (EqualsInsensitive(parts[i], L"SYSTEM") && EqualsInsensitive(parts[i + 1], from)) {
      parts[i + 1] = to;
      return registry_path::Join(parts);
    }
  }
  return L"";
}

inline std::wstring NormalizeCurrentControlSet(const std::wstring& path) {
  std::wstring current = CurrentControlSetSegment();
  if (current.empty()) {
    return path;
  }
  std::wstring replaced = ReplaceControlSetSegment(path, L"CurrentControlSet", current);
  return replaced.empty() ? path : replaced;
}

inline bool IsControlSetSegment(const std::wstring& text) {
  constexpr wchar_t kPrefix[] = L"ControlSet";
  size_t prefix_len = wcslen(kPrefix);
  if (text.size() <= prefix_len || !StartsWithInsensitive(text, kPrefix)) {
    return false;
  }
  for (size_t i = prefix_len; i < text.size(); ++i) {
    if (!iswdigit(text[i])) {
      return false;
    }
  }
  return true;
}

inline std::wstring MapControlSetToCurrent(const std::wstring& path) {
  std::wstring current = CurrentControlSetSegment();
  if (current.empty()) {
    return L"";
  }
  std::vector<std::wstring> parts = registry_path::Split(path);
  if (parts.size() < 3) {
    return L"";
  }
  bool is_hklm = EqualsInsensitive(parts[0], L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(parts[0], L"HKLM");
  if (!is_hklm && parts.size() > 1) {
    if (EqualsInsensitive(parts[0], L"REGISTRY") && EqualsInsensitive(parts[1], L"MACHINE")) {
      is_hklm = true;
    }
  }
  if (!is_hklm) {
    return L"";
  }
  for (size_t i = 0; i + 1 < parts.size(); ++i) {
    if (EqualsInsensitive(parts[i], L"SYSTEM") && IsControlSetSegment(parts[i + 1])) {
      if (EqualsInsensitive(parts[i + 1], current)) {
        return L"";
      }
      parts[i + 1] = current;
      return registry_path::Join(parts);
    }
  }
  return L"";
}

inline std::wstring CleanTraceKeyText(const std::wstring& text, const std::wstring& sid) {
  std::wstring path = text;
  if (!sid.empty()) {
    const std::wstring marker = L"<CURRENT_USER_SID>";
    size_t pos = path.find(marker);
    while (pos != std::wstring::npos) {
      path.replace(pos, marker.size(), sid);
      pos = path.find(marker, pos + sid.size());
    }
  }
  path = registry_path::Clean(path);
  if (path.empty()) {
    return {};
  }
  wchar_t machine[MAX_COMPUTERNAME_LENGTH + 1] = {};
  DWORD machine_len = static_cast<DWORD>(_countof(machine));
  if (GetComputerNameW(machine, &machine_len) && machine_len > 0) {
    const std::wstring prefix = std::wstring(machine, machine_len) + L"\\";
    if (registry_path::StartsWith(path, prefix)) {
      path.erase(0, prefix.size());
    }
  }
  return path;
}

inline std::wstring NormalizeTraceKeyPathBasic(const std::wstring& text) {
  const std::wstring sid = util::GetCurrentUserSidString();
  std::wstring path = CleanTraceKeyText(text, sid);
  if (path.empty()) {
    return L"";
  }
  path = registry_path::Normalize(path, sid);
  const std::wstring current_user = L"HKEY_USERS\\" + sid;
  if (!sid.empty() &&
      (registry_path::Equals(path, current_user) ||
       registry_path::StartsWith(path, current_user + L"\\"))) {
    path.replace(0, current_user.size(), L"HKEY_CURRENT_USER");
  }
  RegistryNode node;
  if (registry_path::ParseRoot(path, &node)) {
    return NormalizeCurrentControlSet(path);
  }
  return L"";
}

inline std::wstring NormalizeTraceKeyPath(const std::wstring& text) {
  std::wstring path = NormalizeTraceKeyPathBasic(text);
  if (path.empty()) {
    return path;
  }
  return ResolveRegistryLinkPath(path);
}

inline std::wstring NormalizeTraceSelectionPath(const std::wstring& text) {
  std::wstring sid = util::GetCurrentUserSidString();
  std::wstring path = CleanTraceKeyText(text, sid);
  if (path.empty()) {
    return L"";
  }
  if (StartsWithInsensitive(path, L"REGISTRY")) {
    std::wstring rest = path.substr(wcslen(L"REGISTRY"));
    while (!rest.empty() && rest.front() == L'\\') {
      rest.erase(rest.begin());
    }
    return rest.empty() ? L"REGISTRY" : L"REGISTRY\\" + rest;
  }
  return path;
}

inline trace::Normalizers TraceNormalizers() {
  trace::Normalizers normalizers;
  normalizers.key = [](const std::wstring& path) {
    return NormalizeTraceKeyPath(path);
  };
  normalizers.display = [](const std::wstring& path) {
    return NormalizeTraceSelectionPath(path);
  };
  return normalizers;
}

struct LinkTargetCache {
  std::mutex mutex;
  std::unordered_map<std::wstring, std::wstring> targets;
  std::unordered_set<std::wstring> misses;
};

inline LinkTargetCache& GetLinkTargetCache() {
  static LinkTargetCache cache;
  return cache;
}

inline bool ParseRegistryRoot(const std::wstring& input, RegistryNode* node, std::wstring* root_label) {
  if (!node || !root_label) {
    return false;
  }
  if (!registry_path::ParseRoot(input, node)) {
    return false;
  }
  *root_label = node->root_name;
  return true;
}

inline bool QueryLinkTargetCached(const std::wstring& path, const RegistryNode& node, std::wstring* target) {
  if (!target) {
    return false;
  }
  *target = L"";
  std::wstring key = ToLower(path);
  LinkTargetCache& cache = GetLinkTargetCache();
  {
    std::lock_guard<std::mutex> lock(cache.mutex);
    auto it = cache.targets.find(key);
    if (it != cache.targets.end()) {
      *target = it->second;
      return true;
    }
    if (cache.misses.find(key) != cache.misses.end()) {
      return false;
    }
  }
  std::wstring resolved;
  if (RegistryStore::QuerySymbolicLinkTarget(node, &resolved)) {
    std::lock_guard<std::mutex> lock(cache.mutex);
    cache.targets.emplace(std::move(key), resolved);
    *target = resolved;
    return true;
  }
  {
    std::lock_guard<std::mutex> lock(cache.mutex);
    cache.misses.insert(std::move(key));
  }
  return false;
}

inline std::wstring ResolveRegistryLinkPath(const std::wstring& path) {
  if (path.empty()) {
    return path;
  }
  std::wstring current = path;
  std::unordered_set<std::wstring> visited;
  for (int depth = 0; depth < 8; ++depth) {
    std::wstring current_lower = ToLower(current);
    if (!visited.insert(current_lower).second) {
      break;
    }
    RegistryNode root_node;
    std::wstring root_label;
    if (!ParseRegistryRoot(current, &root_node, &root_label)) {
      break;
    }
    std::vector<std::wstring> parts = registry_path::Split(root_node.subkey);
    if (parts.empty()) {
      break;
    }
    std::wstring prefix;
    bool resolved = false;
    for (size_t i = 0; i < parts.size(); ++i) {
      if (!prefix.empty()) {
        prefix.append(L"\\");
      }
      prefix.append(parts[i]);
      RegistryNode node = root_node;
      node.subkey = prefix;
      std::wstring prefix_path = root_label;
      if (!prefix.empty()) {
        prefix_path.append(L"\\");
        prefix_path.append(prefix);
      }
      std::wstring target;
      if (!QueryLinkTargetCached(prefix_path, node, &target)) {
        continue;
      }
      std::wstring mapped_target = NormalizeTraceKeyPathBasic(target);
      if (mapped_target.empty()) {
        continue;
      }
      std::wstring remaining = registry_path::Join(parts, i + 1);
      std::wstring next = mapped_target;
      if (!remaining.empty()) {
        next.append(L"\\");
        next.append(remaining);
      }
      current = next;
      resolved = true;
      break;
    }
    if (!resolved) {
      break;
    }
  }
  return current;
}

inline std::wstring FileNameOnly(const std::wstring& path) {
  size_t pos = path.find_last_of(L"\\/");
  return (pos == std::wstring::npos) ? path : path.substr(pos + 1);
}

inline std::wstring FileBaseName(const std::wstring& path) {
  size_t pos = path.find_last_of(L"\\/");
  std::wstring name = (pos == std::wstring::npos) ? path : path.substr(pos + 1);
  size_t dot = name.find_last_of(L'.');
  if (dot != std::wstring::npos) {
    name = name.substr(0, dot);
  }
  return name;
}

struct OfflineHiveCandidate {
  std::wstring path;
  std::wstring label;
};

inline bool IsFilePath(const std::wstring& path) {
  DWORD attrs = GetFileAttributesW(path.c_str());
  return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

inline void AddOfflineHiveCandidate(std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen, const std::wstring& path, const std::wstring& label) {
  if (!out || !seen || !IsFilePath(path)) {
    return;
  }
  std::wstring key = ToLower(path);
  if (!seen->insert(key).second) {
    return;
  }
  std::wstring use_label = TrimWhitespace(label);
  if (use_label.empty()) {
    use_label = TrimWhitespace(FileBaseName(path));
    if (use_label.empty()) {
      use_label = L"OfflineHive";
    }
  }
  out->push_back({path, use_label});
}

inline std::wstring TopLevelFolderLabel(const std::wstring& base, const std::wstring& folder) {
  std::wstring prefix = base;
  if (!prefix.empty() && prefix.back() != L'\\' && prefix.back() != L'/') {
    prefix.push_back(L'\\');
  }
  if (StartsWithInsensitive(folder, prefix)) {
    std::wstring relative = folder.substr(prefix.size());
    size_t sep = relative.find_first_of(L"\\/");
    if (sep != std::wstring::npos) {
      return relative.substr(0, sep);
    }
    if (!relative.empty()) {
      return relative;
    }
  }
  return FileBaseName(folder);
}

inline void CollectUserHiveCandidates(const std::wstring& folder, const std::wstring& base, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen) {
  std::wstring label = TopLevelFolderLabel(base, folder);
  std::wstring ntuser = util::JoinPath(folder, L"NTUSER.DAT");
  AddOfflineHiveCandidate(out, seen, ntuser, label);
  std::wstring usrclass = util::JoinPath(folder, L"USRCLASS.DAT");
  std::wstring class_label = label.empty() ? L"" : (label + L"_Classes");
  AddOfflineHiveCandidate(out, seen, usrclass, class_label);
}

inline void CollectUserHivesRecursive(const std::wstring& folder, const std::wstring& base, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen) {
  WIN32_FIND_DATAW data = {};
  std::wstring search = util::JoinPath(folder, L"*");
  HANDLE find = FindFirstFileW(search.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
      continue;
    }
    if (wcscmp(data.cFileName, L".") == 0 || wcscmp(data.cFileName, L"..") == 0) {
      continue;
    }
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
      continue;
    }
    std::wstring subdir = util::JoinPath(folder, data.cFileName);
    CollectUserHiveCandidates(subdir, base, out, seen);
    CollectUserHivesRecursive(subdir, base, out, seen);
  } while (FindNextFileW(find, &data));
  FindClose(find);
}

inline bool ShouldIncludeOfflineHiveFile(const std::wstring& name) {
  size_t dot = name.find_last_of(L'.');
  if (dot == std::wstring::npos) {
    return true;
  }
  std::wstring ext = name.substr(dot);
  return _wcsicmp(ext.c_str(), L".dat") == 0;
}

inline void CollectLooseHivesInFolder(const std::wstring& folder, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen) {
  WIN32_FIND_DATAW data = {};
  std::wstring search = util::JoinPath(folder, L"*");
  HANDLE find = FindFirstFileW(search.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
      continue;
    }
    if (!ShouldIncludeOfflineHiveFile(data.cFileName)) {
      continue;
    }
    std::wstring candidate = util::JoinPath(folder, data.cFileName);
    std::wstring label = FileBaseName(data.cFileName);
    AddOfflineHiveCandidate(out, seen, candidate, label);
  } while (FindNextFileW(find, &data));
  FindClose(find);
}

inline void CollectOfflineHivesInFolder(const std::wstring& folder, std::vector<OfflineHiveCandidate>* out) {
  if (!out) {
    return;
  }
  out->clear();
  std::unordered_set<std::wstring> seen;
  static const wchar_t* kMachineHives[] = {
      L"SYSTEM",
      L"SOFTWARE",
      L"SAM",
      L"SECURITY",
      L"DEFAULT",
  };
  for (const auto* name : kMachineHives) {
    std::wstring candidate = util::JoinPath(folder, name);
    AddOfflineHiveCandidate(out, &seen, candidate, name);
  }
  CollectUserHiveCandidates(folder, folder, out, &seen);
  CollectLooseHivesInFolder(folder, out, &seen);
  CollectUserHivesRecursive(folder, folder, out, &seen);
}

inline std::wstring ResolveOfflineRootName(const std::wstring& path, bool is_dir, const RegistryNode* current_node) {
  std::wstring base = FileBaseName(path);
  if (is_dir) {
    if (EqualsInsensitive(base, L"HKEY_USERS") || EqualsInsensitive(base, L"HKU")) {
      return L"HKEY_USERS";
    }
    if (EqualsInsensitive(base, L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(base, L"HKLM")) {
      return L"HKEY_LOCAL_MACHINE";
    }
  } else {
    if (EqualsInsensitive(base, L"NTUSER") || EqualsInsensitive(base, L"USRCLASS")) {
      return L"HKEY_USERS";
    }
    if (EqualsInsensitive(base, L"SYSTEM") || EqualsInsensitive(base, L"SOFTWARE") || EqualsInsensitive(base, L"SAM") || EqualsInsensitive(base, L"SECURITY") || EqualsInsensitive(base, L"DEFAULT") || EqualsInsensitive(base, L"COMPONENTS") || EqualsInsensitive(base, L"BCD")) {
      return L"HKEY_LOCAL_MACHINE";
    }
  }
  if (current_node && (current_node->root == HKEY_LOCAL_MACHINE || current_node->root == HKEY_USERS)) {
    std::wstring root_name = registry_path::RootName(current_node->root);
    if (!root_name.empty()) {
      return root_name;
    }
  }
  return L"HKEY_LOCAL_MACHINE";
}

} // namespace regkit::window_detail
