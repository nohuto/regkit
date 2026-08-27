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
inline bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  std::wstring buffer(32768, L'\0');
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer.data();
  ofn.nMaxFile = static_cast<DWORD>(buffer.size());
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
  if (!GetOpenFileNameW(&ofn)) {
    return false;
  }
  *path = buffer.c_str();
  return true;
}

inline bool PromptSaveFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  std::wstring buffer(32768, L'\0');
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer.data();
  ofn.nMaxFile = static_cast<DWORD>(buffer.size());
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
  if (!GetSaveFileNameW(&ofn)) {
    return false;
  }
  *path = buffer.c_str();
  return true;
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

constexpr DWORD kOfflinePickFolderButtonId = 0x2001;

inline std::wstring ShellItemPath(IShellItem* item) {
  if (!item) {
    return L"";
  }
  PWSTR raw = nullptr;
  if (FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &raw)) || !raw) {
    return L"";
  }
  std::wstring result(raw);
  CoTaskMemFree(raw);
  return result;
}

class OfflinePickerEvents final : public IFileDialogEvents, public IFileDialogControlEvents {
public:
  explicit OfflinePickerEvents(IFileDialog* dialog) : dialog_(dialog) {
    if (dialog_) {
      dialog_->AddRef();
    }
  }

  std::wstring picked_path() const { return picked_path_; }

  IFACEMETHODIMP QueryInterface(REFIID riid, void** result) override {
    if (!result) {
      return E_POINTER;
    }
    *result = nullptr;
    if (riid == IID_IUnknown || riid == IID_IFileDialogEvents) {
      *result = static_cast<IFileDialogEvents*>(this);
    } else if (riid == IID_IFileDialogControlEvents) {
      *result = static_cast<IFileDialogControlEvents*>(this);
    } else {
      return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
  }

  IFACEMETHODIMP_(ULONG)
  AddRef() override { return static_cast<ULONG>(InterlockedIncrement(&ref_)); }

  IFACEMETHODIMP_(ULONG)
  Release() override {
    ULONG ref = static_cast<ULONG>(InterlockedDecrement(&ref_));
    if (ref == 0) {
      delete this;
    }
    return ref;
  }

  IFACEMETHODIMP OnFileOk(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnFolderChanging(IFileDialog*, IShellItem*) override { return S_OK; }
  IFACEMETHODIMP OnFolderChange(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnSelectionChange(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnShareViolation(IFileDialog*, IShellItem*, FDE_SHAREVIOLATION_RESPONSE* response) override {
    if (response) {
      *response = FDESVR_DEFAULT;
    }
    return S_OK;
  }
  IFACEMETHODIMP OnTypeChange(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnOverwrite(IFileDialog*, IShellItem*, FDE_OVERWRITE_RESPONSE* response) override {
    if (response) {
      *response = FDEOR_DEFAULT;
    }
    return S_OK;
  }

  IFACEMETHODIMP OnItemSelected(IFileDialogCustomize*, DWORD, DWORD) override { return S_OK; }
  IFACEMETHODIMP OnButtonClicked(IFileDialogCustomize*, DWORD id) override {
    if (id != kOfflinePickFolderButtonId || !dialog_) {
      return S_OK;
    }
    picked_path_.clear();
    IShellItem* selection = nullptr;
    if (SUCCEEDED(dialog_->GetCurrentSelection(&selection)) && selection) {
      SFGAOF attrs = 0;
      if (SUCCEEDED(selection->GetAttributes(SFGAO_FOLDER, &attrs)) && (attrs & SFGAO_FOLDER)) {
        picked_path_ = ShellItemPath(selection);
      }
      selection->Release();
    }
    if (picked_path_.empty()) {
      IShellItem* folder = nullptr;
      if (SUCCEEDED(dialog_->GetFolder(&folder)) && folder) {
        picked_path_ = ShellItemPath(folder);
        folder->Release();
      }
    }
    if (!picked_path_.empty()) {
      dialog_->Close(S_OK);
    }
    return S_OK;
  }
  IFACEMETHODIMP OnCheckButtonToggled(IFileDialogCustomize*, DWORD, BOOL) override { return S_OK; }
  IFACEMETHODIMP OnControlActivating(IFileDialogCustomize*, DWORD) override { return S_OK; }

private:
  ~OfflinePickerEvents() {
    if (dialog_) {
      dialog_->Release();
    }
  }

  LONG ref_ = 1;
  IFileDialog* dialog_ = nullptr;
  std::wstring picked_path_;
};

inline bool PromptOpenFolderOrFile(HWND owner, const wchar_t* title, std::wstring* path) {
  if (!path) {
    return false;
  }
  HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
  bool uninit = false;
  if (SUCCEEDED(hr)) {
    uninit = true;
  } else if (hr != RPC_E_CHANGED_MODE) {
    return false;
  }

  IFileOpenDialog* dialog = nullptr;
  hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog));
  if (FAILED(hr) || !dialog) {
    if (uninit) {
      CoUninitialize();
    }
    return false;
  }

  DWORD options = 0;
  dialog->GetOptions(&options);
  options |= FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST | FOS_PATHMUSTEXIST;
  dialog->SetOptions(options);
  if (title) {
    dialog->SetTitle(title);
  }
  const COMDLG_FILTERSPEC filters[] = {
      {L"Registry Hive Files (*.dat;*.hiv;*.hive;*.sav;SYSTEM;SOFTWARE;SAM;SECURITY;DEFAULT;NTUSER.DAT;USRCLASS.DAT)", L"*.dat;*.hiv;*.hive;*.sav;SYSTEM;SOFTWARE;SAM;SECURITY;DEFAULT;NTUSER.DAT;USRCLASS.DAT"},
      {L"All Files (*.*)", L"*.*"},
  };
  dialog->SetFileTypes(static_cast<UINT>(_countof(filters)), filters);
  dialog->SetFileTypeIndex(1);

  IFileDialogCustomize* customize = nullptr;
  if (SUCCEEDED(dialog->QueryInterface(IID_PPV_ARGS(&customize))) && customize) {
    customize->AddPushButton(kOfflinePickFolderButtonId, L"Select Folder");
    customize->Release();
  }

  OfflinePickerEvents* events = new OfflinePickerEvents(dialog);
  DWORD cookie = 0;
  if (events) {
    dialog->Advise(events, &cookie);
  }

  hr = dialog->Show(owner);
  if (cookie) {
    dialog->Unadvise(cookie);
  }

  std::wstring selected;
  if (events) {
    selected = events->picked_path();
    events->Release();
  }

  if (selected.empty() && SUCCEEDED(hr)) {
    IShellItem* item = nullptr;
    if (SUCCEEDED(dialog->GetResult(&item)) && item) {
      selected = ShellItemPath(item);
      item->Release();
    }
  }

  dialog->Release();
  if (uninit) {
    CoUninitialize();
  }

  if (selected.empty() || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) {
    return false;
  }
  *path = selected;
  return true;
}

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

inline std::wstring ReadFontSubstitute(const wchar_t* value_name) {
  if (!value_name || !*value_name) {
    return L"";
  }
  const wchar_t* subkey = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\FontSubstitutes";
  auto query = [&](REGSAM sam) -> std::wstring {
    HKEY key = nullptr;
    LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, subkey, 0, sam, &key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    DWORD type = 0;
    DWORD bytes = 0;
    result = RegQueryValueExW(key, value_name, nullptr, &type, nullptr, &bytes);
    if (result != ERROR_SUCCESS || bytes == 0 || (type != REG_SZ && type != REG_EXPAND_SZ)) {
      RegCloseKey(key);
      return L"";
    }
    std::vector<wchar_t> buffer(bytes / sizeof(wchar_t) + 1, L'\0');
    result = RegQueryValueExW(key, value_name, nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &bytes);
    RegCloseKey(key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    std::wstring value(buffer.data());
    while (!value.empty() && value.back() == L'\0') {
      value.pop_back();
    }
    if (value.empty()) {
      return L"";
    }
    if (type == REG_EXPAND_SZ) {
      std::wstring expanded = util::ExpandEnvironmentStringsDynamic(value);
      if (!expanded.empty()) {
        value = std::move(expanded);
      }
    }
    return value;
  };

  std::wstring value = query(KEY_READ | KEY_WOW64_64KEY);
  if (!value.empty()) {
    return value;
  }
  return query(KEY_READ);
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
  return util::JoinPath(windows_dir, L"regedit.exe");
}

inline bool LaunchDefaultRegeditProcess(const std::wstring& regedit_path, DWORD* error_code) {
  if (error_code) {
    *error_code = ERROR_SUCCESS;
  }
  if (regedit_path.empty()) {
    if (error_code) {
      *error_code = ERROR_FILE_NOT_FOUND;
    }
    return false;
  }
  std::wstring command_line = L"\"" + regedit_path + L"\" /m";
  STARTUPINFOW startup = {};
  startup.cb = sizeof(startup);
  PROCESS_INFORMATION process = {};
  if (!CreateProcessW(nullptr, command_line.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &startup, &process)) {
    if (error_code) {
      *error_code = GetLastError();
    }
    return false;
  }
  CloseHandle(process.hThread);
  CloseHandle(process.hProcess);
  return true;
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
struct TextMatch {
  bool matched = false;
  size_t start = std::wstring::npos;
  size_t length = 0;
};

class TextMatcher {
public:
  TextMatcher(const std::wstring& query, bool use_regex, bool match_case, bool match_whole, bool* ok) : query_(query), use_regex_(use_regex), match_case_(match_case), match_whole_(match_whole) {
    if (use_regex_) {
      try {
        auto flags = std::regex_constants::ECMAScript;
        if (!match_case_) {
          flags |= std::regex_constants::icase;
        }
        regex_ = std::wregex(query_, flags);
      } catch (const std::regex_error&) {
        if (ok) {
          *ok = false;
        }
      }
    }
  }

  TextMatch Match(const std::wstring& text) const {
    TextMatch match;
    if (text.empty()) {
      return match;
    }
    if (use_regex_) {
      std::wsmatch regex_match;
      if (match_whole_) {
        if (std::regex_match(text, regex_match, regex_)) {
          match.matched = true;
          match.start = 0;
          match.length = regex_match.length();
        }
      } else if (std::regex_search(text, regex_match, regex_)) {
        match.matched = true;
        match.start = regex_match.position();
        match.length = regex_match.length();
      }
      return match;
    }

    if (match_whole_) {
      if (match_case_) {
        if (text == query_) {
          match.matched = true;
          match.start = 0;
          match.length = text.size();
        }
      } else if (CompareStringOrdinal(text.c_str(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE) == CSTR_EQUAL) {
        match.matched = true;
        match.start = 0;
        match.length = text.size();
      }
      return match;
    }

    if (match_case_) {
      size_t pos = text.find(query_);
      if (pos != std::wstring::npos) {
        match.matched = true;
        match.start = pos;
        match.length = query_.size();
      }
    } else {
      int pos = FindStringOrdinal(FIND_FROMSTART, text.c_str(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE);
      if (pos >= 0) {
        match.matched = true;
        match.start = static_cast<size_t>(pos);
        match.length = query_.size();
      }
    }
    return match;
  }

private:
  std::wstring query_;
  bool use_regex_ = false;
  bool match_case_ = false;
  bool match_whole_ = false;
  std::wregex regex_;
};

} // namespace regkit::window_detail
