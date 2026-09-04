// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/window_draw_detail.h"

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
inline bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  return ui::ReportFileDialogResult(owner, win32::ChooseFileToOpen(owner, filter, path));
}

inline bool PromptSaveFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  return ui::ReportFileDialogResult(owner, win32::ChooseFileToSave(owner, filter, nullptr, nullptr, path));
}

inline std::wstring TrimTrailingSeparators(const std::wstring& path) {
  std::wstring result = path;
  while (!result.empty() && (result.back() == L'\\' || result.back() == L'/')) {
    result.pop_back();
  }
  return result;
}

inline bool IsDirectoryPath(const std::wstring& path) {
  DWORD attrs = GetFileAttributesW(path.c_str());
  return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

constexpr wchar_t kIconSetDefault[] = L"phosphor-regedit";
constexpr wchar_t kIconSetPhosphor[] = L"phosphor";
constexpr wchar_t kIconSetLegacyDefault[] = L"default";
constexpr wchar_t kIconSetLucide[] = L"lucide";
constexpr wchar_t kIconSetMaterialSymbols[] = L"materialsymbols";
constexpr wchar_t kIconSetCustom[] = L"custom";

inline bool IsIconSetName(const std::wstring& value, const wchar_t* name) {
  return _wcsicmp(value.c_str(), name) == 0;
}

inline bool IsKnownIconSetName(const std::wstring& value) {
  return IsIconSetName(value, kIconSetDefault) ||
         IsIconSetName(value, kIconSetPhosphor) ||
         IsIconSetName(value, kIconSetLucide) ||
         IsIconSetName(value, kIconSetMaterialSymbols) ||
         IsIconSetName(value, kIconSetCustom);
}

inline std::wstring FindAssetsIconsRoot() {
  std::wstring base = util::GetModuleDirectory();
  for (int i = 0; i < 6; ++i) {
    if (base.empty()) {
      break;
    }
    std::wstring candidate = util::JoinPath(base, L"assets\\icons");
    if (IsDirectoryPath(candidate)) {
      return candidate;
    }
    base = registry_path::Parent(base);
  }
  DWORD len = GetCurrentDirectoryW(0, nullptr);
  if (len > 0) {
    std::wstring cwd(len, L'\0');
    DWORD written = GetCurrentDirectoryW(len, MutableData(cwd));
    if (written != 0) {
      if (written < cwd.size() && cwd[written] == L'\0') {
        cwd.resize(written);
      }
      base = cwd;
      for (int i = 0; i < 3; ++i) {
        if (base.empty()) {
          break;
        }
        std::wstring candidate = util::JoinPath(base, L"assets\\icons");
        if (IsDirectoryPath(candidate)) {
          return candidate;
        }
        base = registry_path::Parent(base);
      }
    }
  }
  return L"";
}

inline std::wstring AssetsIconsRoot() {
  static std::wstring cached;
  static bool cached_set = false;
  if (!cached_set) {
    cached = FindAssetsIconsRoot();
    cached_set = true;
  }
  return cached;
}

constexpr wchar_t kOfflineHiveFilter[] =
    L"Registry Hive Files\0*.dat;*.hiv;*.hive;*.sav;SYSTEM;SOFTWARE;SAM;SECURITY;DEFAULT;NTUSER.DAT;USRCLASS.DAT\0All Files (*.*)\0*.*\0";

inline bool HasRegExtension(const std::wstring& path) {
  size_t dot = path.find_last_of(L'.');
  if (dot == std::wstring::npos) {
    return false;
  }
  std::wstring ext = path.substr(dot);
  return _wcsicmp(ext.c_str(), L".reg") == 0;
}

inline std::wstring EnsureRegExtension(std::wstring path) {
  if (path.empty() || HasRegExtension(path)) {
    return path;
  }
  path.append(L".reg");
  return path;
}

inline bool IsWhitespaceOnly(const std::wstring& text) {
  for (wchar_t ch : text) {
    if (!iswspace(static_cast<wint_t>(ch))) {
      return false;
    }
  }
  return true;
}

inline std::wstring NormalizeMachineName(const std::wstring& text) {
  std::wstring trimmed = TrimWhitespace(text);
  while (!trimmed.empty() && (trimmed.back() == L'\\' || trimmed.back() == L'/')) {
    trimmed.pop_back();
  }
  if (trimmed.empty()) {
    return trimmed;
  }
  if (trimmed.rfind(L"\\\\", 0) == 0) {
    return trimmed;
  }
  return L"\\\\" + trimmed;
}

inline std::wstring StripMachinePrefix(const std::wstring& machine) {
  if (machine.rfind(L"\\\\", 0) == 0) {
    return machine.substr(2);
  }
  return machine;
}

inline bool FileExists(const std::wstring& path) {
  DWORD attrs = GetFileAttributesW(path.c_str());
  return attrs != INVALID_FILE_ATTRIBUTES && !(attrs & FILE_ATTRIBUTE_DIRECTORY);
}

inline bool EqualsInsensitive(const std::wstring& left, const std::wstring& right) {
  return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

inline bool StartsWithInsensitive(const std::wstring& text, const std::wstring& prefix) {
  if (prefix.empty()) {
    return true;
  }
  if (text.size() < prefix.size()) {
    return false;
  }
  return CompareStringOrdinal(text.c_str(), static_cast<int>(prefix.size()), prefix.c_str(), static_cast<int>(prefix.size()), TRUE) == CSTR_EQUAL;
}

inline bool WindowClassEquals(HWND hwnd, const wchar_t* class_name) {
  if (!hwnd || !class_name) {
    return false;
  }
  wchar_t buffer[64] = {};
  if (!GetClassNameW(hwnd, buffer, static_cast<int>(_countof(buffer)))) {
    return false;
  }
  return _wcsicmp(buffer, class_name) == 0;
}

inline std::wstring GetDefaultRegeditPath() {
  DWORD needed = GetWindowsDirectoryW(nullptr, 0);
  if (needed == 0) {
    return L"";
  }
  std::wstring windows_dir(needed, L'\0');
  DWORD written = GetWindowsDirectoryW(windows_dir.data(), needed);
  if (written == 0 || written >= needed) {
    return L"";
  }
  windows_dir.resize(written);
  BOOL wow64 = FALSE;
  if (IsWow64Process(GetCurrentProcess(), &wow64) && wow64) {
    return util::JoinPath(util::JoinPath(windows_dir, L"Sysnative"), L"regedt32.exe");
  }
  return util::JoinPath(windows_dir, L"regedit.exe");
}

struct ParsedRegFileRoot {
  std::wstring name;
  std::shared_ptr<VirtualRegistryData> data;
};

struct RegFileParsePayload : work::MoveOnly {
  uint64_t generation = 0;
  std::wstring source_path;
  std::wstring source_lower;
  std::vector<ParsedRegFileRoot> roots;
  std::wstring error;
  bool cancelled = false;
};

inline VirtualRegistryKey* EnsureVirtualKey(VirtualRegistryKey* root, const std::wstring& subkey) {
  if (!root) {
    return nullptr;
  }
  if (subkey.empty()) {
    return root;
  }
  auto parts = registry_path::Split(subkey);
  VirtualRegistryKey* current = root;
  for (const auto& part : parts) {
    std::wstring lower = ToLower(part);
    auto it = current->children.find(lower);
    if (it == current->children.end()) {
      auto child = std::make_unique<VirtualRegistryKey>();
      child->name = part;
      it = current->children.emplace(lower, std::move(child)).first;
    }
    current = it->second.get();
  }
  return current;
}

inline bool ParseRegFileToVirtualRoots(const std::wstring& path, std::vector<ParsedRegFileRoot>* roots, std::wstring* error, const std::atomic_bool* cancel, bool* cancelled) {
  if (!roots) {
    return false;
  }
  roots->clear();

  regfile::Document document;
  if (!regfile::Load(path, &document, error, cancel, cancelled)) {
    return false;
  }

  std::unordered_map<std::wstring, size_t> root_lookup;
  auto ensure_root = [&](const std::wstring& root_name) {
    const std::wstring lower = ToLower(root_name);
    auto existing = root_lookup.find(lower);
    if (existing != root_lookup.end()) {
      return roots->at(existing->second).data.get();
    }
    ParsedRegFileRoot root;
    root.name = root_name;
    root.data = std::make_shared<VirtualRegistryData>();
    root.data->root_name = root_name;
    root.data->root = std::make_unique<VirtualRegistryKey>();
    root.data->root->name = root_name;
    roots->push_back(std::move(root));
    root_lookup.emplace(lower, roots->size() - 1);
    return roots->back().data.get();
  };

  for (const auto& source_path : document.key_order) {
    if (cancel && cancel->load()) {
      if (cancelled) {
        *cancelled = true;
      }
      return false;
    }
    const std::wstring normalized = NormalizeTraceKeyPathBasic(source_path);
    const std::wstring key_path = normalized.empty() ? source_path : normalized;
    const size_t slash = key_path.find(L'\\');
    const std::wstring root_name = key_path.substr(0, slash);
    const std::wstring subkey = slash == std::wstring::npos
                                    ? L""
                                    : key_path.substr(slash + 1);
    if (root_name.empty()) {
      continue;
    }
    auto source = document.keys.find(ToLower(source_path));
    if (source == document.keys.end()) {
      continue;
    }
    VirtualRegistryData* root = ensure_root(root_name);
    VirtualRegistryKey* target =
        root ? EnsureVirtualKey(root->root.get(), subkey) : nullptr;
    if (!target) {
      continue;
    }
    target->values.reserve(source->second.values.size());
    for (const auto& entry : source->second.values) {
      VirtualRegistryValue value;
      value.name = entry.second.name;
      value.type = entry.second.type;
      value.data = entry.second.data;
      target->values.emplace(entry.first, std::move(value));
    }
  }
  return true;
}

} // namespace regkit::window_detail
