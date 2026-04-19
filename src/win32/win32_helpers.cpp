// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#include "win32/win32_helpers.h"

#include <array>
#include <cwchar>
#include <vector>

#include <pathcch.h>
#include <sddl.h>
#include <shlobj.h>
#include <userenv.h>
#include <winsvc.h>
#include <wtsapi32.h>

namespace {

class ScopedHandle {
public:
  ScopedHandle() noexcept = default;
  explicit ScopedHandle(HANDLE handle) noexcept : handle_(handle) {}
  ~ScopedHandle() { reset(); }
  ScopedHandle(const ScopedHandle&) = delete;
  ScopedHandle& operator=(const ScopedHandle&) = delete;
  ScopedHandle(ScopedHandle&& other) noexcept : handle_(other.handle_) { other.handle_ = nullptr; }
  ScopedHandle& operator=(ScopedHandle&& other) noexcept {
    if (this != &other) {
      reset();
      handle_ = other.handle_;
      other.handle_ = nullptr;
    }
    return *this;
  }

  HANDLE get() const noexcept { return handle_; }
  HANDLE* put() noexcept {
    reset();
    return &handle_;
  }
  HANDLE release() noexcept {
    HANDLE temp = handle_;
    handle_ = nullptr;
    return temp;
  }
  void reset(HANDLE handle = nullptr) noexcept {
    if (handle_ && handle_ != INVALID_HANDLE_VALUE) {
      CloseHandle(handle_);
    }
    handle_ = handle;
  }
  explicit operator bool() const noexcept { return handle_ && handle_ != INVALID_HANDLE_VALUE; }

private:
  HANDLE handle_ = nullptr;
};

class ScopedEnvBlock {
public:
  ScopedEnvBlock() noexcept = default;
  ~ScopedEnvBlock() { reset(); }
  ScopedEnvBlock(const ScopedEnvBlock&) = delete;
  ScopedEnvBlock& operator=(const ScopedEnvBlock&) = delete;

  LPVOID get() const noexcept { return block_; }
  LPVOID* put() noexcept {
    reset();
    return &block_;
  }
  void reset(LPVOID block = nullptr) noexcept {
    if (block_) {
      DestroyEnvironmentBlock(block_);
    }
    block_ = block;
  }

private:
  LPVOID block_ = nullptr;
};

DWORD GetActiveSessionId() {
  DWORD count = 0;
  PWTS_SESSION_INFOW sessions = nullptr;
  if (!WTSEnumerateSessionsW(WTS_CURRENT_SERVER_HANDLE, 0, 1, &sessions, &count)) {
    return static_cast<DWORD>(-1);
  }
  DWORD active_session = static_cast<DWORD>(-1);
  for (DWORD i = 0; i < count; ++i) {
    if (sessions[i].State == WTS_CONNECTSTATE_CLASS::WTSActive) {
      active_session = sessions[i].SessionId;
      break;
    }
  }
  WTSFreeMemory(sessions);
  return active_session;
}

bool CreateSystemToken(DWORD desired_access, HANDLE* token_handle) {
  if (!token_handle) {
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }
  *token_handle = nullptr;

  DWORD lsass_pid = 0;
  DWORD winlogon_pid = 0;
  DWORD process_count = 0;
  PWTS_PROCESS_INFOW processes = nullptr;
  DWORD session_id = GetActiveSessionId();

  if (WTSEnumerateProcessesW(WTS_CURRENT_SERVER_HANDLE, 0, 1, &processes, &process_count)) {
    for (DWORD i = 0; i < process_count; ++i) {
      const auto& process = processes[i];
      if (!process.pProcessName || !process.pUserSid || !IsWellKnownSid(process.pUserSid, WELL_KNOWN_SID_TYPE::WinLocalSystemSid)) {
        continue;
      }
      if (lsass_pid == 0 && process.SessionId == 0 && _wcsicmp(process.pProcessName, L"lsass.exe") == 0) {
        lsass_pid = process.ProcessId;
        continue;
      }
      if (winlogon_pid == 0 && process.SessionId == session_id && _wcsicmp(process.pProcessName, L"winlogon.exe") == 0) {
        winlogon_pid = process.ProcessId;
        continue;
      }
    }
    WTSFreeMemory(processes);
  }

  ScopedHandle system_process;
  if (lsass_pid != 0) {
    system_process.reset(OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, lsass_pid));
  }
  if (!system_process && winlogon_pid != 0) {
    system_process.reset(OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, winlogon_pid));
  }
  if (!system_process) {
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }

  ScopedHandle system_token;
  if (!OpenProcessToken(system_process.get(), TOKEN_DUPLICATE, system_token.put())) {
    return false;
  }

  if (!DuplicateTokenEx(system_token.get(), desired_access, nullptr, SecurityIdentification, TokenPrimary, token_handle)) {
    return false;
  }

  return true;
}

bool EnablePrivilege(HANDLE token, const wchar_t* privilege) {
  if (!token || !privilege) {
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }
  LUID luid = {};
  if (!LookupPrivilegeValueW(nullptr, privilege, &luid)) {
    return false;
  }
  TOKEN_PRIVILEGES tp = {};
  tp.PrivilegeCount = 1;
  tp.Privileges[0].Luid = luid;
  tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
  AdjustTokenPrivileges(token, FALSE, &tp, sizeof(tp), nullptr, nullptr);
  return GetLastError() == ERROR_SUCCESS;
}

bool AdjustTokenAllPrivileges(HANDLE token, DWORD attributes) {
  DWORD length = 0;
  GetTokenInformation(token, TokenPrivileges, nullptr, 0, &length);
  if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || length == 0) {
    return false;
  }
  std::vector<BYTE> buffer(length);
  if (!GetTokenInformation(token, TokenPrivileges, buffer.data(), length, &length)) {
    return false;
  }
  auto* privileges = reinterpret_cast<TOKEN_PRIVILEGES*>(buffer.data());
  for (DWORD i = 0; i < privileges->PrivilegeCount; ++i) {
    privileges->Privileges[i].Attributes = attributes;
  }
  AdjustTokenPrivileges(token, FALSE, privileges, length, nullptr, nullptr);
  return GetLastError() == ERROR_SUCCESS;
}

bool QueryServiceProcess(SC_HANDLE service, SERVICE_STATUS_PROCESS* status) {
  if (!status) {
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }
  DWORD bytes = 0;
  return QueryServiceStatusEx(service, SC_STATUS_PROCESS_INFO, reinterpret_cast<LPBYTE>(status), sizeof(SERVICE_STATUS_PROCESS), &bytes) != FALSE;
}

bool StartServiceAndGetProcessId(const wchar_t* service_name, DWORD* process_id) {
  if (!service_name || !process_id) {
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }
  *process_id = 0;

  SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
  if (!scm) {
    return false;
  }
  SC_HANDLE service = OpenServiceW(scm, service_name, SERVICE_QUERY_STATUS | SERVICE_START);
  if (!service) {
    CloseServiceHandle(scm);
    return false;
  }

  SERVICE_STATUS_PROCESS status = {};
  if (!QueryServiceProcess(service, &status)) {
    CloseServiceHandle(service);
    CloseServiceHandle(scm);
    return false;
  }

  if (status.dwCurrentState == SERVICE_STOPPED || status.dwCurrentState == SERVICE_STOP_PENDING) {
    StartServiceW(service, 0, nullptr);
  }

  for (int attempt = 0; attempt < 50; ++attempt) {
    if (!QueryServiceProcess(service, &status)) {
      CloseServiceHandle(service);
      CloseServiceHandle(scm);
      return false;
    }
    if (status.dwCurrentState == SERVICE_RUNNING && status.dwProcessId != 0) {
      *process_id = status.dwProcessId;
      CloseServiceHandle(service);
      CloseServiceHandle(scm);
      return true;
    }
    Sleep(100);
  }

  CloseServiceHandle(service);
  CloseServiceHandle(scm);
  SetLastError(ERROR_SERVICE_REQUEST_TIMEOUT);
  return false;
}

bool OpenServiceProcessToken(const wchar_t* service_name, DWORD desired_access, HANDLE* token_handle) {
  if (!token_handle) {
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }
  *token_handle = nullptr;

  DWORD process_id = 0;
  if (!StartServiceAndGetProcessId(service_name, &process_id)) {
    return false;
  }
  ScopedHandle process(OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, process_id));
  if (!process) {
    return false;
  }
  if (!OpenProcessToken(process.get(), desired_access, token_handle)) {
    return false;
  }
  return true;
}

} // namespace

namespace util {

ComInit::ComInit(DWORD flags) noexcept : hr_(CoInitializeEx(nullptr, flags)) {
}

ComInit::~ComInit() {
  if (SUCCEEDED(hr_)) {
    CoUninitialize();
  }
}

bool ComInit::ok() const noexcept {
  return SUCCEEDED(hr_);
}

UniqueHKey::UniqueHKey(HKEY key) noexcept : key_(key) {
}

UniqueHKey::~UniqueHKey() {
  reset();
}

UniqueHKey::UniqueHKey(UniqueHKey&& other) noexcept : key_(other.key_) {
  other.key_ = nullptr;
}

UniqueHKey& UniqueHKey::operator=(UniqueHKey&& other) noexcept {
  if (this == &other) {
    return *this;
  }
  reset();
  key_ = other.key_;
  other.key_ = nullptr;
  return *this;
}

HKEY UniqueHKey::get() const noexcept {
  return key_;
}

HKEY* UniqueHKey::put() noexcept {
  reset();
  return &key_;
}

HKEY UniqueHKey::release() noexcept {
  HKEY temp = key_;
  key_ = nullptr;
  return temp;
}

void UniqueHKey::reset(HKEY key) noexcept {
  if (key_) {
    RegCloseKey(key_);
  }
  key_ = key;
}

std::wstring GetModuleDirectory() {
  wchar_t buffer[MAX_PATH] = {};
  DWORD len = GetModuleFileNameW(nullptr, buffer, static_cast<DWORD>(std::size(buffer)));
  if (len == 0) {
    return L"";
  }
  std::wstring path(buffer, len);
  if (FAILED(PathCchRemoveFileSpec(path.data(), path.size() + 1))) {
    return L"";
  }
  size_t new_len = wcsnlen_s(path.data(), path.size());
  path.resize(new_len);
  return path;
}

std::wstring JoinPath(const std::wstring& left, const std::wstring& right) {
  if (left.empty()) {
    return right;
  }
  wchar_t buffer[MAX_PATH] = {};
  if (SUCCEEDED(PathCchCombine(buffer, std::size(buffer), left.c_str(), right.c_str()))) {
    return buffer;
  }
  if (left.back() == L'\\') {
    return left + right;
  }
  return left + L"\\" + right;
}

std::string WideToUtf8(const std::wstring& text) {
  if (text.empty()) {
    return {};
  }
  int bytes = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
  if (bytes <= 0) {
    return {};
  }
  std::string out(static_cast<size_t>(bytes), '\0');
  WideCharToMultiByte(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()), out.data(), bytes, nullptr, nullptr);
  return out;
}

std::wstring Utf8ToWide(const std::string& text) {
  if (text.empty()) {
    return L"";
  }
  int chars = MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
  if (chars <= 0) {
    return L"";
  }
  std::wstring out(static_cast<size_t>(chars), L'\0');
  MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), out.data(), chars);
  return out;
}

std::wstring ToHex(const BYTE* data, size_t size, size_t max_bytes) {
  if (!data || size == 0) {
    return L"";
  }
  size_t count = (max_bytes == 0 || max_bytes > size) ? size : max_bytes;
  std::wstring out;
  out.reserve(count * 3 + (count < size ? 4 : 0));
  static constexpr wchar_t kHexDigits[] = L"0123456789abcdef";
  for (size_t i = 0; i < count; ++i) {
    BYTE value = data[i];
    out.push_back(kHexDigits[(value >> 4) & 0xF]);
    out.push_back(kHexDigits[value & 0xF]);
    if (i + 1 < count) {
      out.push_back(L' ');
    }
  }
  if (count < size) {
    out.append(L" ...");
  }
  return out;
}

std::wstring GetCurrentUserSidString() {
  static const std::wstring cached = []() -> std::wstring {
    std::wstring sid_string;
    HANDLE token = nullptr;
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token)) {
      return sid_string;
    }
    DWORD size = 0;
    GetTokenInformation(token, TokenUser, nullptr, 0, &size);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || size == 0) {
      CloseHandle(token);
      return sid_string;
    }
    std::vector<BYTE> buffer(size);
    if (!GetTokenInformation(token, TokenUser, buffer.data(), size, &size)) {
      CloseHandle(token);
      return sid_string;
    }
    auto* user = reinterpret_cast<TOKEN_USER*>(buffer.data());
    LPWSTR sid = nullptr;
    if (ConvertSidToStringSidW(user->User.Sid, &sid) && sid) {
      sid_string.assign(sid);
      LocalFree(sid);
    }
    CloseHandle(token);
    return sid_string;
  }();
  return cached;
}

std::wstring GetAppDataFolder() {
  PWSTR base = nullptr;
  if (FAILED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &base)) || !base) {
    return L"";
  }
  std::wstring folder = util::JoinPath(base, L"Noverse\\RegKit");
  CoTaskMemFree(base);
  if (!folder.empty()) {
    SHCreateDirectoryExW(nullptr, folder.c_str(), nullptr);
  }
  return folder;
}

bool IsProcessElevated() {
  ScopedHandle token;
  if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, token.put())) {
    return false;
  }
  TOKEN_ELEVATION elevation = {};
  DWORD size = 0;
  bool elevated = false;
  if (GetTokenInformation(token.get(), TokenElevation, &elevation, sizeof(elevation), &size)) {
    elevated = (elevation.TokenIsElevated != 0);
  }
  return elevated;
}

bool IsProcessSystem() {
  ScopedHandle token;
  if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, token.put())) {
    return false;
  }
  DWORD size = 0;
  GetTokenInformation(token.get(), TokenUser, nullptr, 0, &size);
  if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || size == 0) {
    return false;
  }
  std::vector<BYTE> buffer(size);
  if (!GetTokenInformation(token.get(), TokenUser, buffer.data(), size, &size)) {
    return false;
  }
  auto* user = reinterpret_cast<TOKEN_USER*>(buffer.data());
  return IsWellKnownSid(user->User.Sid, WELL_KNOWN_SID_TYPE::WinLocalSystemSid) != FALSE;
}

bool IsProcessTrustedInstaller() {
  static const std::vector<BYTE> ti_sid = []() -> std::vector<BYTE> {
    const wchar_t* account = L"NT SERVICE\\TrustedInstaller";
    DWORD sid_size = 0;
    DWORD domain_size = 0;
    SID_NAME_USE use = SidTypeUnknown;
    LookupAccountNameW(nullptr, account, nullptr, &sid_size, nullptr, &domain_size, &use);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || sid_size == 0) {
      return {};
    }
    std::vector<BYTE> sid_buffer(sid_size);
    std::wstring domain(domain_size, L'\0');
    if (!LookupAccountNameW(nullptr, account, sid_buffer.data(), &sid_size, domain.data(), &domain_size, &use)) {
      return {};
    }
    sid_buffer.resize(sid_size);
    return sid_buffer;
  }();

  if (ti_sid.empty()) {
    return false;
  }

  ScopedHandle token;
  if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, token.put())) {
    return false;
  }
  DWORD size = 0;
  GetTokenInformation(token.get(), TokenUser, nullptr, 0, &size);
  if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || size == 0) {
    return false;
  }
  std::vector<BYTE> buffer(size);
  if (!GetTokenInformation(token.get(), TokenUser, buffer.data(), size, &size)) {
    return false;
  }
  auto* user = reinterpret_cast<TOKEN_USER*>(buffer.data());
  PSID ti_sid_ptr = const_cast<PSID>(static_cast<const void*>(ti_sid.data()));
  if (EqualSid(user->User.Sid, ti_sid_ptr)) {
    return true;
  }

  DWORD group_size = 0;
  GetTokenInformation(token.get(), TokenGroups, nullptr, 0, &group_size);
  if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || group_size == 0) {
    return false;
  }
  std::vector<BYTE> group_buffer(group_size);
  if (!GetTokenInformation(token.get(), TokenGroups, group_buffer.data(), group_size, &group_size)) {
    return false;
  }
  auto* groups = reinterpret_cast<TOKEN_GROUPS*>(group_buffer.data());
  for (DWORD i = 0; i < groups->GroupCount; ++i) {
    if (EqualSid(groups->Groups[i].Sid, ti_sid_ptr)) {
      return true;
    }
  }
  return false;
}

bool LaunchProcessAsSystem(const std::wstring& command_line, const std::wstring& work_dir, DWORD* error_code) {
  if (error_code) {
    *error_code = ERROR_SUCCESS;
  }
  if (command_line.empty()) {
    if (error_code) {
      *error_code = ERROR_INVALID_PARAMETER;
    }
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }

  bool result = false;
  DWORD error = ERROR_SUCCESS;

  // Impersonate to enable SeDebug, duplicate a SYSTEM token, then launch in the
  // active session.
  ScopedHandle current_token;
  ScopedHandle current_impersonation;
  ScopedHandle system_token;
  ScopedHandle system_impersonation;
  ScopedHandle target_token;
  ScopedEnvBlock env;
  DWORD session_id = static_cast<DWORD>(-1);
  STARTUPINFOW startup = {};
  PROCESS_INFORMATION process = {};
  std::wstring mutable_command;

  if (!OpenProcessToken(GetCurrentProcess(), MAXIMUM_ALLOWED, current_token.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!DuplicateTokenEx(current_token.get(), MAXIMUM_ALLOWED, nullptr, SecurityImpersonation, TokenImpersonation, current_impersonation.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!EnablePrivilege(current_impersonation.get(), SE_DEBUG_NAME)) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!SetThreadToken(nullptr, current_impersonation.get())) {
    error = GetLastError();
    goto Cleanup;
  }
  session_id = GetActiveSessionId();
  if (session_id == static_cast<DWORD>(-1)) {
    error = ERROR_NO_TOKEN;
    goto Cleanup;
  }
  if (!CreateSystemToken(MAXIMUM_ALLOWED, system_token.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!DuplicateTokenEx(system_token.get(), MAXIMUM_ALLOWED, nullptr, SecurityImpersonation, TokenImpersonation, system_impersonation.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!AdjustTokenAllPrivileges(system_impersonation.get(), SE_PRIVILEGE_ENABLED)) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!SetThreadToken(nullptr, system_impersonation.get())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!DuplicateTokenEx(system_token.get(), MAXIMUM_ALLOWED, nullptr, SecurityIdentification, TokenPrimary, target_token.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!SetTokenInformation(target_token.get(), TokenSessionId, &session_id, sizeof(session_id))) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!AdjustTokenAllPrivileges(target_token.get(), SE_PRIVILEGE_ENABLED)) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!CreateEnvironmentBlock(env.put(), current_token.get(), TRUE)) {
    error = GetLastError();
    goto Cleanup;
  }

  startup.cb = sizeof(startup);
  mutable_command = command_line;
  result = CreateProcessAsUserW(target_token.get(), nullptr, mutable_command.data(), nullptr, nullptr, FALSE, CREATE_UNICODE_ENVIRONMENT, env.get(), work_dir.empty() ? nullptr : work_dir.c_str(), &startup, &process);
  if (!result) {
    error = GetLastError();
    goto Cleanup;
  }
  CloseHandle(process.hThread);
  CloseHandle(process.hProcess);

Cleanup:
  SetThreadToken(nullptr, nullptr);
  if (!result) {
    if (error_code) {
      *error_code = error;
    }
    SetLastError(error);
  }
  return result;
}

bool LaunchProcessAsTrustedInstaller(const std::wstring& command_line, const std::wstring& work_dir, DWORD* error_code) {
  if (error_code) {
    *error_code = ERROR_SUCCESS;
  }
  if (command_line.empty()) {
    if (error_code) {
      *error_code = ERROR_INVALID_PARAMETER;
    }
    SetLastError(ERROR_INVALID_PARAMETER);
    return false;
  }

  bool result = false;
  DWORD error = ERROR_SUCCESS;

  ScopedHandle current_token;
  ScopedHandle current_impersonation;
  ScopedHandle system_token;
  ScopedHandle system_impersonation;
  ScopedHandle ti_token;
  ScopedHandle target_token;
  ScopedEnvBlock env;
  DWORD session_id = static_cast<DWORD>(-1);
  STARTUPINFOW startup = {};
  PROCESS_INFORMATION process = {};
  std::wstring mutable_command;

  if (!OpenProcessToken(GetCurrentProcess(), MAXIMUM_ALLOWED, current_token.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!DuplicateTokenEx(current_token.get(), MAXIMUM_ALLOWED, nullptr, SecurityImpersonation, TokenImpersonation, current_impersonation.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!EnablePrivilege(current_impersonation.get(), SE_DEBUG_NAME)) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!SetThreadToken(nullptr, current_impersonation.get())) {
    error = GetLastError();
    goto Cleanup;
  }
  session_id = GetActiveSessionId();
  if (session_id == static_cast<DWORD>(-1)) {
    error = ERROR_NO_TOKEN;
    goto Cleanup;
  }
  if (!CreateSystemToken(MAXIMUM_ALLOWED, system_token.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!DuplicateTokenEx(system_token.get(), MAXIMUM_ALLOWED, nullptr, SecurityImpersonation, TokenImpersonation, system_impersonation.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!AdjustTokenAllPrivileges(system_impersonation.get(), SE_PRIVILEGE_ENABLED)) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!SetThreadToken(nullptr, system_impersonation.get())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!OpenServiceProcessToken(L"TrustedInstaller", MAXIMUM_ALLOWED, ti_token.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!DuplicateTokenEx(ti_token.get(), MAXIMUM_ALLOWED, nullptr, SecurityIdentification, TokenPrimary, target_token.put())) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!SetTokenInformation(target_token.get(), TokenSessionId, &session_id, sizeof(session_id))) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!AdjustTokenAllPrivileges(target_token.get(), SE_PRIVILEGE_ENABLED)) {
    error = GetLastError();
    goto Cleanup;
  }
  if (!CreateEnvironmentBlock(env.put(), current_token.get(), TRUE)) {
    error = GetLastError();
    goto Cleanup;
  }

  startup.cb = sizeof(startup);
  mutable_command = command_line;
  result = CreateProcessAsUserW(target_token.get(), nullptr, mutable_command.data(), nullptr, nullptr, FALSE, CREATE_UNICODE_ENVIRONMENT, env.get(), work_dir.empty() ? nullptr : work_dir.c_str(), &startup, &process);
  if (!result) {
    error = GetLastError();
    goto Cleanup;
  }
  CloseHandle(process.hThread);
  CloseHandle(process.hProcess);

Cleanup:
  SetThreadToken(nullptr, nullptr);
  if (!result) {
    if (error_code) {
      *error_code = error;
    }
    SetLastError(error);
  }
  return result;
}

} // namespace util
