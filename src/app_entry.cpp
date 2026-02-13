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

#include <windows.h>

#include <commctrl.h>
#include <limits>
#include <shellapi.h>
#include <string>
#include <vector>

#include "app/app_window.h"
#include "app/theme.h"
#include "app/ui_helpers.h"
#include "win32/win32_helpers.h"

namespace {

constexpr wchar_t kRestartSystemArg[] = L"--restart-system";
constexpr wchar_t kRestartTiArg[] = L"--restart-ti";

bool ParseBool(const std::wstring& value) {
  return (_wcsicmp(value.c_str(), L"1") == 0 || _wcsicmp(value.c_str(), L"true") == 0 || _wcsicmp(value.c_str(), L"yes") == 0);
}

std::wstring FormatWin32Error(DWORD code) {
  if (code == 0) {
    return L"";
  }
  wchar_t buffer[512] = {};
  DWORD len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, code, 0, buffer, static_cast<DWORD>(_countof(buffer)), nullptr);
  if (len == 0) {
    return L"Unknown error.";
  }
  return buffer;
}

std::vector<std::wstring> GetCommandLineArgs() {
  int argc = 0;
  LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
  std::vector<std::wstring> args;
  if (!argv) {
    return args;
  }
  for (int i = 1; i < argc; ++i) {
    args.emplace_back(argv[i]);
  }
  LocalFree(argv);
  return args;
}

bool HasCommandLineArg(const std::vector<std::wstring>& args, const wchar_t* arg) {
  if (!arg || !*arg) {
    return false;
  }
  for (const auto& entry : args) {
    if (_wcsicmp(entry.c_str(), arg) == 0) {
      return true;
    }
  }
  return false;
}

bool HasRegExtension(const std::wstring& path) {
  size_t dot = path.find_last_of(L'.');
  if (dot == std::wstring::npos) {
    return false;
  }
  std::wstring ext = path.substr(dot);
  return _wcsicmp(ext.c_str(), L".reg") == 0;
}

bool LoadSingleInstanceSetting() {
  const bool default_value = true;
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return default_value;
  }
  std::wstring path = util::JoinPath(folder, L"settings.ini");
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return default_value;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return default_value;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return default_value;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  int chars = MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), nullptr, 0);
  if (chars <= 0) {
    return default_value;
  }
  std::wstring content(static_cast<size_t>(chars), L'\0');
  MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), content.data(), chars);

  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    if (line.empty()) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = line.substr(0, sep);
    std::wstring value = line.substr(sep + 1);
    if (_wcsicmp(key.c_str(), L"single_instance") == 0) {
      return ParseBool(value);
    }
  }
  return default_value;
}

bool LoadAlwaysRunAsAdminSetting() {
  const bool default_value = false;
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return default_value;
  }
  std::wstring path = util::JoinPath(folder, L"settings.ini");
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return default_value;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return default_value;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return default_value;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  int chars = MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), nullptr, 0);
  if (chars <= 0) {
    return default_value;
  }
  std::wstring content(static_cast<size_t>(chars), L'\0');
  MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), content.data(), chars);

  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    if (line.empty()) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = line.substr(0, sep);
    std::wstring value = line.substr(sep + 1);
    if (_wcsicmp(key.c_str(), L"always_run_as_admin") == 0) {
      return ParseBool(value);
    }
  }
  return default_value;
}

bool LoadAlwaysRunAsSystemSetting() {
  const bool default_value = false;
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return default_value;
  }
  std::wstring path = util::JoinPath(folder, L"settings.ini");
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return default_value;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return default_value;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return default_value;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  int chars = MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), nullptr, 0);
  if (chars <= 0) {
    return default_value;
  }
  std::wstring content(static_cast<size_t>(chars), L'\0');
  MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), content.data(), chars);

  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    if (line.empty()) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = line.substr(0, sep);
    std::wstring value = line.substr(sep + 1);
    if (_wcsicmp(key.c_str(), L"always_run_as_system") == 0) {
      return ParseBool(value);
    }
  }
  return default_value;
}

bool LoadAlwaysRunAsTrustedInstallerSetting() {
  const bool default_value = false;
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return default_value;
  }
  std::wstring path = util::JoinPath(folder, L"settings.ini");
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return default_value;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return default_value;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return default_value;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  int chars = MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), nullptr, 0);
  if (chars <= 0) {
    return default_value;
  }
  std::wstring content(static_cast<size_t>(chars), L'\0');
  MultiByteToWideChar(CP_UTF8, 0, buffer.data(), static_cast<int>(buffer.size()), content.data(), chars);

  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    if (line.empty()) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = line.substr(0, sep);
    std::wstring value = line.substr(sep + 1);
    if (_wcsicmp(key.c_str(), L"always_run_as_trustedinstaller") == 0) {
      return ParseBool(value);
    }
  }
  return default_value;
}

bool IsProcessElevated() {
  return util::IsProcessElevated();
}

bool IsProcessSystem() {
  return util::IsProcessSystem();
}

bool IsProcessTrustedInstaller() {
  return util::IsProcessTrustedInstaller();
}

bool RelaunchAsAdmin() {
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    return false;
  }
  exe_path.resize(len);
  HINSTANCE result = ShellExecuteW(nullptr, L"runas", exe_path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  return reinterpret_cast<INT_PTR>(result) > 32;
}

bool RestartAsSystem(std::wstring* error_message, bool* launched) {
  if (error_message) {
    error_message->clear();
  }
  if (launched) {
    *launched = false;
  }
  if (IsProcessSystem()) {
    return true;
  }
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    if (error_message) {
      *error_message = L"Failed to locate the executable path.";
    }
    return false;
  }
  exe_path.resize(len);
  if (!IsProcessElevated()) {
    HINSTANCE result = ShellExecuteW(nullptr, L"runas", exe_path.c_str(), kRestartSystemArg, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
      if (error_message) {
        *error_message = L"Failed to request SYSTEM restart.";
      }
      return false;
    }
    if (launched) {
      *launched = true;
    }
    return true;
  }

  std::wstring command_line = L"\"";
  command_line += exe_path;
  command_line += L"\" ";
  command_line += kRestartSystemArg;
  DWORD error = 0;
  if (!util::LaunchProcessAsSystem(command_line, L"", &error)) {
    if (error_message) {
      std::wstring message = L"Failed to restart with SYSTEM rights.";
      std::wstring detail = FormatWin32Error(error);
      if (!detail.empty()) {
        message += L"\n";
        message += detail;
      }
      *error_message = message;
    }
    return false;
  }
  if (launched) {
    *launched = true;
  }
  return true;
}

bool RestartAsTrustedInstaller(std::wstring* error_message, bool* launched) {
  if (error_message) {
    error_message->clear();
  }
  if (launched) {
    *launched = false;
  }
  if (IsProcessTrustedInstaller()) {
    return true;
  }
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    if (error_message) {
      *error_message = L"Failed to locate the executable path.";
    }
    return false;
  }
  exe_path.resize(len);
  if (!IsProcessElevated()) {
    HINSTANCE result = ShellExecuteW(nullptr, L"runas", exe_path.c_str(), kRestartTiArg, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
      if (error_message) {
        *error_message = L"Failed to request TrustedInstaller restart.";
      }
      return false;
    }
    if (launched) {
      *launched = true;
    }
    return true;
  }

  std::wstring command_line = L"\"";
  command_line += exe_path;
  command_line += L"\" ";
  command_line += kRestartTiArg;
  DWORD error = 0;
  if (!util::LaunchProcessAsTrustedInstaller(command_line, L"", &error)) {
    if (error_message) {
      std::wstring message = L"Failed to restart with TrustedInstaller rights.";
      std::wstring detail = FormatWin32Error(error);
      if (!detail.empty()) {
        message += L"\n";
        message += detail;
      }
      *error_message = message;
    }
    return false;
  }
  if (launched) {
    *launched = true;
  }
  return true;
}

} // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int cmd_show) {
  regkit::Theme::InitializeDarkModeSupport();
  util::ComInit com;
  if (!com.ok()) {
    regkit::ui::ShowError(nullptr, L"COM initialization failed.");
    return 1;
  }

  INITCOMMONCONTROLSEX icc = {};
  icc.dwSize = sizeof(icc);
  icc.dwICC = ICC_WIN95_CLASSES | ICC_STANDARD_CLASSES | ICC_BAR_CLASSES | ICC_TAB_CLASSES | ICC_DATE_CLASSES | ICC_COOL_CLASSES | ICC_PROGRESS_CLASS;
  InitCommonControlsEx(&icc);

  const auto args = GetCommandLineArgs();
  bool restart_system = HasCommandLineArg(args, kRestartSystemArg);
  bool restart_ti = HasCommandLineArg(args, kRestartTiArg);
  if (restart_ti) {
    std::wstring error;
    bool launched = false;
    if (RestartAsTrustedInstaller(&error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (restart_system) {
    std::wstring error;
    bool launched = false;
    if (RestartAsSystem(&error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (LoadAlwaysRunAsTrustedInstallerSetting() && !IsProcessTrustedInstaller()) {
    std::wstring error;
    bool launched = false;
    if (RestartAsTrustedInstaller(&error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (LoadAlwaysRunAsSystemSetting() && !IsProcessSystem()) {
    std::wstring error;
    bool launched = false;
    if (RestartAsSystem(&error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (LoadAlwaysRunAsAdminSetting() && !IsProcessElevated()) {
    if (RelaunchAsAdmin()) {
      return 0;
    }
    regkit::ui::ShowError(nullptr, L"Administrator restart was cancelled.");
  }

  HANDLE instance_mutex = nullptr;
  if (!restart_system && !restart_ti && LoadSingleInstanceSetting()) {
    instance_mutex = CreateMutexW(nullptr, TRUE, L"RegKit.SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
      HWND existing = FindWindowW(L"RegKitMainWindow", nullptr);
      if (existing) {
        ShowWindow(existing, SW_RESTORE);
        SetForegroundWindow(existing);
      }
      if (instance_mutex) {
        CloseHandle(instance_mutex);
      }
      return 0;
    }
  }

  regkit::MainWindow window;
  if (!window.Create(instance)) {
    regkit::ui::ShowError(nullptr, L"Failed to create the main window.");
    if (instance_mutex) {
      CloseHandle(instance_mutex);
    }
    return 1;
  }
  window.Show(cmd_show);

  for (const auto& arg : args) {
    if (!arg.empty() && arg[0] == L'-') {
      continue;
    }
    if (HasRegExtension(arg)) {
      window.OpenRegFileTab(arg);
    }
  }

  MSG msg = {};
  while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
    if (window.TranslateAccelerator(msg)) {
      continue;
    }
    TranslateMessage(&msg);
    DispatchMessageW(&msg);
  }
  if (instance_mutex) {
    CloseHandle(instance_mutex);
  }
  return static_cast<int>(msg.wParam);
}
