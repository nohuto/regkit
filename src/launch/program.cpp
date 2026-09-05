// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/windows_config.h"

#include <windows.h>

#include <algorithm>
#include <commctrl.h>
#include <uxtheme.h>
#include <cwctype>
#include <limits>
#include <shellapi.h>
#include <string>
#include <vector>

#include "cli/reg_command.h"
#include "frame/main_window.h"
#include "regfile/registry_transfer.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "appearance/theme.h"
#include "appearance/presets.h"
#include "appearance/feedback.h"
#include "win32/handle_owner.h"
#include "win32/system_error.h"
#include "win32/text_transform.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/restart.h"
#include "win32/shell_paths.h"

namespace {

using regkit::win32::kRestartSystemArg;
using regkit::win32::kRestartTiArg;
constexpr ULONG_PTR kExternalJumpCopyDataId = 0x52474A54;
constexpr wchar_t kRegKitWindowProperty[] = L"RegKitMainWindow";

using util::FormatWin32Error;
using util::TrimWhitespace;

bool ParseBool(const std::wstring& value) {
  return (_wcsicmp(value.c_str(), L"1") == 0 || _wcsicmp(value.c_str(), L"true") == 0 || _wcsicmp(value.c_str(), L"yes") == 0);
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

std::wstring BaseName(const std::wstring& path) {
  size_t slash = path.find_last_of(L"\\/");
  if (slash == std::wstring::npos) {
    return path;
  }
  return path.substr(slash + 1);
}

bool IsRegeditLaunchArg(const std::wstring& arg) {
  if (arg.empty()) {
    return false;
  }
  std::wstring name = BaseName(arg);
  return _wcsicmp(name.c_str(), L"regedit.exe") == 0 ||
         _wcsicmp(name.c_str(), L"regedit") == 0 ||
         _wcsicmp(name.c_str(), L"regedt32.exe") == 0 ||
         _wcsicmp(name.c_str(), L"regedt32") == 0;
}

bool IsInterceptedRegeditLaunch(const std::vector<std::wstring>& args) {
  for (const auto& arg : args) {
    if (IsRegeditLaunchArg(arg)) {
      return true;
    }
  }
  return false;
}

std::vector<std::wstring> StripRegeditLaunchArg(const std::vector<std::wstring>& args) {
  std::vector<std::wstring> stripped;
  stripped.reserve(args.size());
  for (const auto& arg : args) {
    if (!IsRegeditLaunchArg(arg)) {
      stripped.push_back(arg);
    }
  }
  return stripped;
}

std::vector<std::wstring> RegFilesFromArgs(const std::vector<std::wstring>& args) {
  std::vector<std::wstring> files;
  for (const auto& arg : args) {
    if (arg.empty() || arg[0] == L'-' || arg[0] == L'/' || IsRegeditLaunchArg(arg)) {
      continue;
    }
    if (HasRegExtension(arg)) {
      files.push_back(arg);
    }
  }
  return files;
}

bool LooksLikeRegistryPath(const std::wstring& arg) {
  regkit::RegistryNode node;
  return regkit::registry_path::ParseRoot(arg, &node);
}

bool ReadRegeditLastKey(std::wstring* out) {
  if (!out) {
    return false;
  }
  out->clear();

  HKEY key = nullptr;
  LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit", 0, KEY_QUERY_VALUE, &key);
  if (result != ERROR_SUCCESS) {
    return false;
  }

  DWORD type = 0;
  DWORD size = 0;
  result = RegQueryValueExW(key, L"LastKey", nullptr, &type, nullptr, &size);
  if (result != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || size < sizeof(wchar_t)) {
    RegCloseKey(key);
    return false;
  }

  std::wstring value;
  value.resize(size / sizeof(wchar_t));
  result = RegQueryValueExW(key, L"LastKey", nullptr, &type, reinterpret_cast<LPBYTE>(value.data()), &size);
  RegCloseKey(key);
  if (result != ERROR_SUCCESS) {
    return false;
  }

  while (!value.empty() && value.back() == L'\0') {
    value.pop_back();
  }
  if (value.empty()) {
    return false;
  }

  if (type == REG_EXPAND_SZ) {
    DWORD needed = ExpandEnvironmentStringsW(value.c_str(), nullptr, 0);
    if (needed > 1) {
      std::wstring expanded;
      expanded.resize(needed);
      DWORD written = ExpandEnvironmentStringsW(value.c_str(), expanded.data(), needed);
      if (written > 0 && written <= needed) {
        while (!expanded.empty() && expanded.back() == L'\0') {
          expanded.pop_back();
        }
        if (!expanded.empty()) {
          value = std::move(expanded);
        }
      }
    }
  }

  *out = value;
  return true;
}

bool ResolveExternalJumpTarget(const std::vector<std::wstring>& args, std::wstring* out) {
  if (!out) {
    return false;
  }
  out->clear();
  bool intercepted_regedit = false;
  std::wstring explicit_key_path;
  for (size_t index = 0; index < args.size(); ++index) {
    const std::wstring& arg = args[index];
    if (_wcsicmp(arg.c_str(), L"--goto") == 0 || _wcsicmp(arg.c_str(), L"/goto") == 0) {
      if (index + 1 < args.size()) {
        explicit_key_path = args[++index];
      }
      continue;
    }
    if (IsRegeditLaunchArg(arg)) {
      intercepted_regedit = true;
      continue;
    }
    if (arg.empty()) {
      continue;
    }
    if (arg[0] == L'-' || arg[0] == L'/') {
      continue;
    }
    if (LooksLikeRegistryPath(arg)) {
      explicit_key_path = arg;
      continue;
    }
    if (!explicit_key_path.empty()) {
      *out = explicit_key_path + L"\\" + arg;
      return true;
    }
  }
  if (!explicit_key_path.empty()) {
    *out = explicit_key_path;
    return true;
  }
  if (!intercepted_regedit) {
    return false;
  }
  return ReadRegeditLastKey(out);
}

BOOL CALLBACK FindRegKitWindowProc(HWND hwnd, LPARAM lparam) {
  if (!GetPropW(hwnd, kRegKitWindowProperty)) {
    return TRUE;
  }
  auto* found = reinterpret_cast<HWND*>(lparam);
  *found = hwnd;
  return FALSE;
}

HWND FindRunningRegKitWindow() {
  HWND found = nullptr;
  EnumWindows(FindRegKitWindowProc, reinterpret_cast<LPARAM>(&found));
  return found;
}

struct StartupSettings {
  bool single_instance = true;
  bool always_run_as_admin = false;
  bool always_run_as_system = false;
  bool always_run_as_trustedinstaller = false;
  regkit::ThemeMode theme_mode = regkit::ThemeMode::kSystem;
  std::wstring theme_preset;
};

bool ReadSettingsFileContent(std::wstring* content) {
  if (!content) {
    return false;
  }
  content->clear();
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return false;
  }
  std::wstring path = util::JoinPath(folder, L"settings.ini");
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return false;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return false;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  *content = util::Utf8ToWide(buffer);
  return !content->empty();
}

StartupSettings LoadStartupSettings() {
  StartupSettings settings;
  std::wstring content;
  if (!ReadSettingsFileContent(&content)) {
    return settings;
  }

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
    std::wstring key = TrimWhitespace(line.substr(0, sep));
    std::wstring value = TrimWhitespace(line.substr(sep + 1));
    if (_wcsicmp(key.c_str(), L"single_instance") == 0) {
      settings.single_instance = ParseBool(value);
    } else if (_wcsicmp(key.c_str(), L"always_run_as_admin") == 0) {
      settings.always_run_as_admin = ParseBool(value);
    } else if (_wcsicmp(key.c_str(), L"always_run_as_system") == 0) {
      settings.always_run_as_system = ParseBool(value);
    } else if (_wcsicmp(key.c_str(), L"always_run_as_trustedinstaller") == 0) {
      settings.always_run_as_trustedinstaller = ParseBool(value);
    } else if (_wcsicmp(key.c_str(), L"theme_mode") == 0) {
      if (_wcsicmp(value.c_str(), L"dark") == 0) {
        settings.theme_mode = regkit::ThemeMode::kDark;
      } else if (_wcsicmp(value.c_str(), L"light") == 0) {
        settings.theme_mode = regkit::ThemeMode::kLight;
      } else if (_wcsicmp(value.c_str(), L"custom") == 0) {
        settings.theme_mode = regkit::ThemeMode::kCustom;
      } else {
        settings.theme_mode = regkit::ThemeMode::kSystem;
      }
    } else if (_wcsicmp(key.c_str(), L"theme_preset") == 0) {
      settings.theme_preset = value;
    }
  }
  return settings;
}

void ApplyStartupTheme(const StartupSettings& settings) {
  if (settings.theme_mode == regkit::ThemeMode::kCustom) {
    std::vector<regkit::ThemePreset> presets;
    if (!regkit::ThemePresetStore::Load(&presets)) {
      presets = regkit::ThemePresetStore::BuiltInPresets();
    }
    if (!presets.empty()) {
      auto it = std::find_if(presets.begin(), presets.end(), [&](const regkit::ThemePreset& preset) {
        return _wcsicmp(preset.name.c_str(), settings.theme_preset.c_str()) == 0;
      });
      if (it == presets.end()) {
        it = presets.begin();
      }
      regkit::Theme::SetCustomColors(it->colors, it->is_dark);
      regkit::Theme::SetMode(regkit::ThemeMode::kCustom);
      return;
    }
  }
  regkit::Theme::SetMode(settings.theme_mode);
}

bool RelaunchAsAdmin(DWORD parent_pid) {
  const std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    return false;
  }
  return SUCCEEDED(regkit::win32::LaunchElevated(
      nullptr, exe_path, regkit::win32::RestartArguments(nullptr, parent_pid)));
}

bool RestartAsSystem(DWORD parent_pid, std::wstring* error_message, bool* launched) {
  if (error_message) {
    error_message->clear();
  }
  if (launched) {
    *launched = false;
  }
  if (util::IsProcessSystem()) {
    return true;
  }
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    if (error_message) {
      *error_message = L"Failed to locate the executable path.";
    }
    return false;
  }
  if (!util::IsProcessElevated()) {
    const HRESULT hr = regkit::win32::LaunchElevated(
        nullptr, exe_path, regkit::win32::RestartArguments(kRestartSystemArg, parent_pid));
    if (FAILED(hr)) {
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
  command_line += regkit::win32::RestartArguments(kRestartSystemArg, parent_pid);
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

bool RestartAsTrustedInstaller(DWORD parent_pid, std::wstring* error_message, bool* launched) {
  if (error_message) {
    error_message->clear();
  }
  if (launched) {
    *launched = false;
  }
  if (util::IsProcessTrustedInstaller()) {
    return true;
  }
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    if (error_message) {
      *error_message = L"Failed to locate the executable path.";
    }
    return false;
  }
  if (!util::IsProcessElevated()) {
    const HRESULT hr = regkit::win32::LaunchElevated(
        nullptr, exe_path, regkit::win32::RestartArguments(kRestartTiArg, parent_pid));
    if (FAILED(hr)) {
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
  command_line += regkit::win32::RestartArguments(kRestartTiArg, parent_pid);
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
  BufferedPaintInit();

  const auto args = GetCommandLineArgs();
  const bool regedit_compat_requested = IsInterceptedRegeditLaunch(args);
  int cli_exit = 0;
  if (regkit::cli::Execute(
          regedit_compat_requested ? StripRegeditLaunchArg(args) : args,
          &cli_exit)) {
    return cli_exit;
  }
  const StartupSettings startup_settings = LoadStartupSettings();
  ApplyStartupTheme(startup_settings);
  std::wstring startup_jump_target;
  const bool external_jump_requested = ResolveExternalJumpTarget(args, &startup_jump_target);
  const std::vector<std::wstring> regedit_merge_files = regedit_compat_requested ? RegFilesFromArgs(args) : std::vector<std::wstring>();
  const bool restart_system = HasCommandLineArg(args, kRestartSystemArg);
  const bool restart_ti = HasCommandLineArg(args, kRestartTiArg);
  const DWORD restart_parent_pid = regkit::win32::RestartParentPid(args);
  const DWORD handoff_pid =
      restart_parent_pid != 0 ? restart_parent_pid : GetCurrentProcessId();
  if (restart_ti) {
    std::wstring error;
    bool launched = false;
    if (RestartAsTrustedInstaller(handoff_pid, &error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (restart_system) {
    std::wstring error;
    bool launched = false;
    if (RestartAsSystem(handoff_pid, &error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (startup_settings.always_run_as_trustedinstaller &&
             !util::IsProcessTrustedInstaller()) {
    std::wstring error;
    bool launched = false;
    if (RestartAsTrustedInstaller(handoff_pid, &error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (startup_settings.always_run_as_system &&
             !util::IsProcessSystem()) {
    std::wstring error;
    bool launched = false;
    if (RestartAsSystem(handoff_pid, &error, &launched)) {
      if (launched) {
        return 0;
      }
    } else if (!error.empty()) {
      regkit::ui::ShowError(nullptr, error);
    }
  } else if (startup_settings.always_run_as_admin &&
             !util::IsProcessElevated()) {
    if (RelaunchAsAdmin(handoff_pid)) {
      return 0;
    }
    regkit::ui::ShowError(nullptr, L"Administrator restart was cancelled.");
  }

  if (regedit_compat_requested && !regedit_merge_files.empty()) {
    for (const auto& path : regedit_merge_files) {
      if (!regkit::ui::ConfirmRegFileMerge(nullptr, path)) {
        return 0;
      }
      std::wstring error;
      if (!regkit::ImportRegFileFromPath(path, &error)) {
        regkit::ui::ShowRegFileMergeFailed(nullptr, path, error);
        return 1;
      }
      regkit::ui::ShowRegFileMergeSucceeded(nullptr, path);
    }
    return 0;
  }

  regkit::win32::WaitForParentExit(restart_parent_pid);

  HANDLE instance_mutex = nullptr;
  if (startup_settings.single_instance) {
    instance_mutex = CreateMutexW(nullptr, TRUE, L"RegKit.SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
      HWND existing = FindRunningRegKitWindow();
      if (existing) {
        if (external_jump_requested && !startup_jump_target.empty()) {
          COPYDATASTRUCT data = {};
          data.dwData = kExternalJumpCopyDataId;
          data.cbData = static_cast<DWORD>((startup_jump_target.size() + 1) * sizeof(wchar_t));
          data.lpData = const_cast<wchar_t*>(startup_jump_target.c_str());
          DWORD_PTR ignored = 0;
          SendMessageTimeoutW(existing, WM_COPYDATA, 0, reinterpret_cast<LPARAM>(&data), SMTO_ABORTIFHUNG, 1500, &ignored);
        }
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
  if (external_jump_requested && !startup_jump_target.empty()) {
    window.QueueExternalJump(startup_jump_target);
  }
  window.Show(cmd_show);

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
  BufferedPaintUnInit();
  return static_cast<int>(msg.wParam);
}
