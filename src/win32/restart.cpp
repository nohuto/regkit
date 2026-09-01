// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/restart.h"

#include <shellapi.h>

namespace regkit::win32 {

std::wstring RestartArguments(const wchar_t* target_arg, DWORD parent_pid) {
  std::wstring arguments;
  if (target_arg && *target_arg) {
    arguments = target_arg;
  }
  if (parent_pid != 0) {
    if (!arguments.empty()) {
      arguments.push_back(L' ');
    }
    arguments += kRestartParentArg;
    arguments.push_back(L' ');
    arguments += std::to_wstring(parent_pid);
  }
  return arguments;
}

HRESULT LaunchElevated(HWND owner, const std::wstring& exe, const std::wstring& arguments) {
  if (exe.empty()) {
    return E_INVALIDARG;
  }
  SHELLEXECUTEINFOW info = {};
  info.cbSize = sizeof(info);
  info.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC;
  info.hwnd = owner;
  info.lpVerb = L"runas";
  info.lpFile = exe.c_str();
  info.lpParameters = arguments.empty() ? nullptr : arguments.c_str();
  info.nShow = SW_SHOWNORMAL;
  if (!ShellExecuteExW(&info)) {
    return HRESULT_FROM_WIN32(GetLastError());
  }
  if (info.hProcess) {
    CloseHandle(info.hProcess);
  }
  return S_OK;
}

DWORD RestartParentPid(const std::vector<std::wstring>& args) {
  for (size_t i = 0; i + 1 < args.size(); ++i) {
    if (_wcsicmp(args[i].c_str(), kRestartParentArg) != 0) {
      continue;
    }
    wchar_t* end = nullptr;
    const unsigned long value = wcstoul(args[i + 1].c_str(), &end, 10);
    if (end && *end == L'\0') {
      return static_cast<DWORD>(value);
    }
  }
  return 0;
}

void WaitForParentExit(DWORD parent_pid) {
  if (parent_pid == 0 || parent_pid == GetCurrentProcessId()) {
    return;
  }
  HANDLE parent = OpenProcess(SYNCHRONIZE, FALSE, parent_pid);
  if (!parent) {
    return;
  }
  WaitForSingleObject(parent, INFINITE);
  CloseHandle(parent);
}

} // namespace regkit::win32
