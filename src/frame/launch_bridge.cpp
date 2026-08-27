// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::ShowPermissionsDialog(const RegistryNode& node) {
  ShowRegistryPermissions(hwnd_, node);
}

bool MainWindow::Impl::RestartAsAdmin() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }
  HINSTANCE result = ShellExecuteW(hwnd_, L"runas", exe_path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  if (reinterpret_cast<INT_PTR>(result) <= 32) {
    ui::ShowError(hwnd_, L"Failed to restart with administrator rights.");
    return false;
  }
  PostMessageW(hwnd_, WM_CLOSE, 0, 0);
  return true;
}

bool MainWindow::Impl::RestartAsSystem() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }

  if (!util::IsProcessElevated()) {
    HINSTANCE result = ShellExecuteW(hwnd_, L"runas", exe_path.c_str(), kRestartSystemArg, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
      ui::ShowError(hwnd_, L"Failed to request SYSTEM restart.");
      return false;
    }
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  }

  std::wstring command_line = L"\"";
  command_line += exe_path;
  command_line += L"\" ";
  command_line += kRestartSystemArg;
  DWORD error = 0;
  if (!util::LaunchProcessAsSystem(command_line, L"", &error)) {
    std::wstring message = L"Failed to restart with SYSTEM rights.";
    std::wstring detail = FormatWin32Error(error);
    if (!detail.empty()) {
      message += L"\n";
      message += detail;
    }
    ui::ShowError(hwnd_, message);
    return false;
  }

  PostMessageW(hwnd_, WM_CLOSE, 0, 0);
  return true;
}

bool MainWindow::Impl::RestartAsTrustedInstaller() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }

  if (!util::IsProcessElevated()) {
    HINSTANCE result = ShellExecuteW(hwnd_, L"runas", exe_path.c_str(), kRestartTiArg, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
      ui::ShowError(hwnd_, L"Failed to request TrustedInstaller restart.");
      return false;
    }
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  }

  std::wstring command_line = L"\"";
  command_line += exe_path;
  command_line += L"\" ";
  command_line += kRestartTiArg;
  DWORD error = 0;
  if (!util::LaunchProcessAsTrustedInstaller(command_line, L"", &error)) {
    std::wstring message = L"Failed to restart with TrustedInstaller rights.";
    std::wstring detail = FormatWin32Error(error);
    if (!detail.empty()) {
      message += L"\n";
      message += detail;
    }
    ui::ShowError(hwnd_, message);
    return false;
  }

  PostMessageW(hwnd_, WM_CLOSE, 0, 0);
  return true;
}

void MainWindow::Impl::SyncReplaceRegeditState() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    return;
  }

  HKEY base = nullptr;
  LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\regedit.exe", 0, KEY_QUERY_VALUE, &base);
  if (result != ERROR_SUCCESS) {
    replace_regedit_ = false;
    return;
  }

  DWORD type = 0;
  DWORD size = 0;
  result = RegQueryValueExW(base, L"Debugger", nullptr, &type, nullptr, &size);
  if (result != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || size == 0) {
    RegCloseKey(base);
    replace_regedit_ = false;
    return;
  }

  std::wstring debugger;
  debugger.resize(size / sizeof(wchar_t));
  result = RegQueryValueExW(base, L"Debugger", nullptr, &type, reinterpret_cast<LPBYTE>(MutableData(debugger)), &size);
  RegCloseKey(base);
  if (result != ERROR_SUCCESS) {
    replace_regedit_ = false;
    return;
  }

  while (!debugger.empty() && debugger.back() == L'\0') {
    debugger.pop_back();
  }
  if (debugger.empty()) {
    replace_regedit_ = false;
    return;
  }

  std::wstring expanded = debugger;
  if (type == REG_EXPAND_SZ) {
    std::wstring resolved = util::ExpandEnvironmentStringsDynamic(debugger);
    if (!resolved.empty()) {
      expanded = std::move(resolved);
    }
  }

  const wchar_t* start = expanded.c_str();
  while (*start && iswspace(*start)) {
    ++start;
  }
  std::wstring path;
  if (*start == L'\"') {
    ++start;
    const wchar_t* end = wcschr(start, L'\"');
    if (end) {
      path.assign(start, static_cast<size_t>(end - start));
    } else {
      path.assign(start);
    }
  } else {
    const wchar_t* end = start;
    while (*end && !iswspace(*end)) {
      ++end;
    }
    path.assign(start, static_cast<size_t>(end - start));
  }

  if (path.empty()) {
    replace_regedit_ = false;
    return;
  }

  replace_regedit_ = (_wcsicmp(path.c_str(), exe_path.c_str()) == 0);
}

void MainWindow::Impl::ReplaceRegedit(bool enable) {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return;
  }

  HKEY base = nullptr;
  DWORD base_disp = 0;
  LONG result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options", 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, &base, &base_disp);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return;
  }

  std::wstring subkey = L"regedit.exe";
  if (enable) {
    HKEY app_key = nullptr;
    DWORD disposition = 0;
    result = RegCreateKeyExW(base, subkey.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, &app_key, &disposition);
    if (result != ERROR_SUCCESS) {
      RegCloseKey(base);
      ui::ShowError(hwnd_, FormatWin32Error(result));
      return;
    }
    std::wstring debugger = L"\"" + exe_path + L"\"";
    result = RegSetValueExW(app_key, L"Debugger", 0, REG_SZ, reinterpret_cast<const BYTE*>(debugger.c_str()), static_cast<DWORD>((debugger.size() + 1) * sizeof(wchar_t)));
    RegCloseKey(app_key);
    RegCloseKey(base);
    if (result != ERROR_SUCCESS) {
      ui::ShowError(hwnd_, FormatWin32Error(result));
      return;
    }
    replace_regedit_ = true;
  } else {
    HKEY app_key = nullptr;
    result = RegOpenKeyExW(base, subkey.c_str(), 0, KEY_READ | KEY_WRITE, &app_key);
    if (result == ERROR_SUCCESS) {
      RegDeleteValueW(app_key, L"Debugger");
      DWORD subkeys = 0;
      DWORD values = 0;
      if (RegQueryInfoKeyW(app_key, nullptr, nullptr, nullptr, &subkeys, nullptr, nullptr, &values, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS && subkeys == 0 && values == 0) {
        RegCloseKey(app_key);
        RegDeleteKeyW(base, subkey.c_str());
      } else {
        RegCloseKey(app_key);
      }
    }
    RegCloseKey(base);
    replace_regedit_ = false;
  }

  SaveSettings();
  BuildMenus();
}

bool MainWindow::Impl::OpenDefaultRegedit() {
  if (!util::IsProcessElevated() && !util::IsProcessSystem() &&
      !util::IsProcessTrustedInstaller()) {
    ui::ShowError(hwnd_, L"Administrator rights are required to open the default Regedit.");
    return false;
  }

  std::wstring regedit_path = GetDefaultRegeditPath();
  if (regedit_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the default Regedit executable.");
    return false;
  }
  auto launch_regedit = [&]() -> bool {
    DWORD launch_error = ERROR_SUCCESS;
    if (LaunchDefaultRegeditProcess(regedit_path, &launch_error)) {
      return true;
    }
    ui::ShowError(hwnd_, FormatWin32Error(launch_error));
    return false;
  };

  util::UniqueHKey key;
  LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, kRegeditIfeoPath, 0, KEY_READ | KEY_WRITE | KEY_WOW64_64KEY, key.put());
  if (result == ERROR_FILE_NOT_FOUND) {
    return launch_regedit();
  }
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }

  DWORD type = 0;
  DWORD size = 0;
  result = RegQueryValueExW(key.get(), L"Debugger", nullptr, &type, nullptr, &size);
  if (result == ERROR_FILE_NOT_FOUND) {
    return launch_regedit();
  }
  if (result != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || size == 0) {
    ui::ShowError(hwnd_, L"Failed to read the Regedit debugger value.");
    return false;
  }

  std::vector<BYTE> data(size);
  result = RegQueryValueExW(key.get(), L"Debugger", nullptr, &type, data.data(), &size);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }
  data.resize(size);

  std::wstring temp_name = L"Debugger_RegKitTemp";
  DWORD temp_type = 0;
  DWORD temp_size = 0;
  int suffix = 0;
  while (RegQueryValueExW(key.get(), temp_name.c_str(), nullptr, &temp_type, nullptr, &temp_size) == ERROR_SUCCESS) {
    ++suffix;
    temp_name = L"Debugger_RegKitTemp_" + std::to_wstring(suffix);
    if (suffix > 100) {
      ui::ShowError(hwnd_, L"Failed to prepare a temporary Regedit debugger value.");
      return false;
    }
  }

  result = RegSetValueExW(key.get(), temp_name.c_str(), 0, type, data.data(), size);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }
  result = RegDeleteValueW(key.get(), L"Debugger");
  if (result != ERROR_SUCCESS) {
    RegDeleteValueW(key.get(), temp_name.c_str());
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }

  DWORD launch_error = ERROR_SUCCESS;
  bool launched = LaunchDefaultRegeditProcess(regedit_path, &launch_error);

  LONG restore = RegSetValueExW(key.get(), L"Debugger", 0, type, data.data(), size);
  RegDeleteValueW(key.get(), temp_name.c_str());
  if (restore != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(restore));
    return false;
  }
  if (!launched) {
    ui::ShowError(hwnd_, FormatWin32Error(launch_error));
    return false;
  }
  return true;
}

std::wstring MainWindow::Impl::ResolveSelectedHiveFilePath() {
  if (registry_mode_ == RegistryMode::kRemote) {
    return L"";
  }
  RegistryNode* node = browse_.current_node();
  if (!node && browse_.tree().hwnd()) {
    HTREEITEM selected = TreeView_GetSelection(browse_.tree().hwnd());
    if (selected) {
      node = browse_.tree().NodeFromItem(selected);
    }
  }
  if (!node) {
    return L"";
  }
  RegistryNode target = *node;
  int index = browse_.values().hwnd()
                  ? ListView_GetNextItem(browse_.values().hwnd(), -1,
                                         LVNI_SELECTED)
                  : -1;
  if (index >= 0) {
    const ListRow* row = browse_.values().RowAt(index);
    if (row && row->kind == rowkind::kKey && !row->extra.empty()) {
      target = MakeChildNode(*node, row->extra);
    }
  }
  return LookupHivePath(target, nullptr);
}

void MainWindow::Impl::OpenHiveFileDir() {
  if (registry_mode_ == RegistryMode::kRemote) {
    ui::ShowError(hwnd_, L"Hive files are not available for remote registries.");
    return;
  }
  std::wstring hive_path = ResolveSelectedHiveFilePath();
  if (hive_path.empty()) {
    ui::ShowError(hwnd_, L"No hive file was found for this key.");
    return;
  }
  std::wstring args = L"/select,\"" + hive_path + L"\"";
  std::wstring folder = hive_path;
  folder.push_back(L'\0');
  if (SUCCEEDED(PathCchRemoveFileSpec(folder.data(), folder.size()))) {
    folder.resize(wcsnlen_s(folder.data(), folder.size()));
    ShellExecuteW(hwnd_, L"open", L"explorer.exe", args.c_str(), folder.c_str(), SW_SHOWNORMAL);
  } else {
    ShellExecuteW(hwnd_, L"open", L"explorer.exe", args.c_str(), nullptr, SW_SHOWNORMAL);
  }
}

LOGFONTW MainWindow::Impl::DefaultLogFont() const {
  LOGFONTW lf = {};
  std::wstring face = ReadFontSubstitute(L"Segoe UI");
  if (face.empty()) {
    face = L"Segoe UI";
  }
  lf.lfHeight = appearance::FontHeight(9);
  lf.lfWeight = FW_NORMAL;
  lf.lfCharSet = DEFAULT_CHARSET;
  wcsncpy_s(lf.lfFaceName, face.c_str(), _TRUNCATE);
  return lf;
}

} // namespace regkit
