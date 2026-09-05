// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/shell_paths.h"

#include "win32/windows_config.h"

#include <windows.h>

#include <algorithm>
#include <pathcch.h>
#include <shlobj.h>

namespace util {

std::wstring GetModulePath() {
  DWORD capacity = MAX_PATH;
  for (;;) {
    std::wstring path(capacity, L'\0');
    const DWORD length =
        GetModuleFileNameW(nullptr, path.data(), capacity);
    if (length == 0) {
      return {};
    }
    if (length < capacity) {
      path.resize(length);
      return path;
    }
    if (capacity >= 32768) {
      return {};
    }
    capacity = std::min<DWORD>(capacity * 2, 32768);
  }
}

namespace {

using PathCchRemoveFileSpecFn = HRESULT(WINAPI*)(PWSTR, size_t);
using PathCchCombineFn = HRESULT(WINAPI*)(PWSTR, size_t, PCWSTR, PCWSTR);

template <typename Fn>
Fn LoadPathFunction(const char* name) {
  HMODULE module = GetModuleHandleW(L"kernelbase.dll");
  if (!module) {
    module = LoadLibraryExW(L"kernelbase.dll", nullptr,
                            LOAD_LIBRARY_SEARCH_SYSTEM32);
  }
  return module ? reinterpret_cast<Fn>(GetProcAddress(module, name)) : nullptr;
}

} // namespace

std::wstring GetModuleDirectory() {
  std::wstring path = GetModulePath();
  if (path.empty()) {
    return {};
  }
  static const auto remove_file_spec =
      LoadPathFunction<PathCchRemoveFileSpecFn>("PathCchRemoveFileSpec");
  if (remove_file_spec) {
    path.push_back(L'\0');
    if (FAILED(remove_file_spec(path.data(), path.size()))) {
      return {};
    }
    path.resize(wcsnlen_s(path.data(), path.size()));
    return path;
  }
  const size_t separator = path.find_last_of(L"\\/");
  if (separator == std::wstring::npos) {
    return {};
  }
  path.resize(separator);
  return path;
}

std::wstring JoinPath(const std::wstring& left, const std::wstring& right) {
  if (left.empty()) {
    return right;
  }
  const size_t capacity =
      std::min<size_t>(std::max<size_t>(MAX_PATH,
                                        left.size() + right.size() + 2),
                       32768);
  static const auto combine =
      LoadPathFunction<PathCchCombineFn>("PathCchCombine");
  if (combine) {
    std::wstring output(capacity, L'\0');
    if (SUCCEEDED(combine(output.data(), output.size(), left.c_str(),
                          right.c_str()))) {
      return output.c_str();
    }
  }
  return left.back() == L'\\' ? left + right : left + L"\\" + right;
}

std::wstring GetAppDataFolder() {
  const DWORD override_size =
      GetEnvironmentVariableW(L"REGKIT_DATA_DIR", nullptr, 0);
  if (override_size > 1) {
    std::wstring path(override_size, L'\0');
    const DWORD written = GetEnvironmentVariableW(
        L"REGKIT_DATA_DIR", path.data(), override_size);
    if (written > 0 && written < override_size) {
      path.resize(written);
      SHCreateDirectoryExW(nullptr, path.c_str(), nullptr);
      return path;
    }
  }

  PWSTR base = nullptr;
  if (FAILED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &base)) ||
      !base) {
    return {};
  }
  std::wstring folder = JoinPath(base, L"Noverse\\RegKit");
  CoTaskMemFree(base);
  SHCreateDirectoryExW(nullptr, folder.c_str(), nullptr);
  return folder;
}

} // namespace util
