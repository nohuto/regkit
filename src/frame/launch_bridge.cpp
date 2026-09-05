// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::ShowPermissionsDialog(const RegistryNode& node) {
  ShowRegistryPermissions(hwnd_, node);
}

namespace {

constexpr wchar_t kRegeditImageOptionsKey[] =
    L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution "
    L"Options\\regedit.exe";

bool BeginRestart(HWND owner, const wchar_t* target_arg, const wchar_t* failure) {
  const std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(owner, L"Failed to locate the executable path.");
    return false;
  }
  const std::wstring arguments =
      win32::RestartArguments(target_arg, GetCurrentProcessId());
  const HRESULT hr = win32::LaunchElevated(owner, exe_path, arguments);
  if (FAILED(hr)) {
    if (!win32::DialogCancelled(hr)) {
      ui::ShowError(owner, std::wstring(failure) + L"\n" + win32::FormatDialogError(hr));
    }
    return false;
  }
  PostMessageW(owner, WM_CLOSE, 0, 0);
  return true;
}

bool BrokerRestart(HWND owner, const wchar_t* target_arg, const wchar_t* failure,
                   bool (*launch)(const std::wstring&, const std::wstring&, DWORD*)) {
  const std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(owner, L"Failed to locate the executable path.");
    return false;
  }
  std::wstring command_line = L"\"";
  command_line += exe_path;
  command_line += L"\" ";
  command_line += win32::RestartArguments(target_arg, GetCurrentProcessId());
  DWORD error = 0;
  if (!launch(command_line, L"", &error)) {
    std::wstring message = failure;
    const std::wstring detail = FormatWin32Error(error);
    if (!detail.empty()) {
      message += L"\n";
      message += detail;
    }
    ui::ShowError(owner, message);
    return false;
  }
  PostMessageW(owner, WM_CLOSE, 0, 0);
  return true;
}

} // namespace

bool MainWindow::Impl::RestartAsAdmin() {
  return BeginRestart(hwnd_, nullptr, L"Failed to restart with administrator rights.");
}

bool MainWindow::Impl::RestartAsSystem() {
  if (!util::IsProcessElevated()) {
    return BeginRestart(hwnd_, kRestartSystemArg, L"Failed to request SYSTEM restart.");
  }
  return BrokerRestart(hwnd_, kRestartSystemArg, L"Failed to restart with SYSTEM rights.",
                       util::LaunchProcessAsSystem);
}

bool MainWindow::Impl::RestartAsTrustedInstaller() {
  if (!util::IsProcessElevated()) {
    return BeginRestart(hwnd_, kRestartTiArg, L"Failed to request TrustedInstaller restart.");
  }
  return BrokerRestart(hwnd_, kRestartTiArg, L"Failed to restart with TrustedInstaller rights.",
                       util::LaunchProcessAsTrustedInstaller);
}

void MainWindow::Impl::SyncReplaceRegeditState() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    return;
  }

  HKEY base = nullptr;
  LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, kRegeditImageOptionsKey, 0, KEY_QUERY_VALUE, &base);
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
  result = RegQueryValueExW(base, L"Debugger", nullptr, &type, reinterpret_cast<LPBYTE>(debugger.data()), &size);
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

  BuildMenus();
}

bool MainWindow::Impl::OpenDefaultRegedit() {
  const std::wstring regedit_path = GetDefaultRegeditPath();
  if (regedit_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the default Regedit executable.");
    return false;
  }
  HKEY hijack = nullptr;
  std::wstring debugger;
  DWORD debugger_type = REG_SZ;
  if (replace_regedit_ &&
      RegOpenKeyExW(HKEY_LOCAL_MACHINE, kRegeditImageOptionsKey, 0,
                    KEY_QUERY_VALUE | KEY_SET_VALUE, &hijack) == ERROR_SUCCESS) {
    DWORD size = 0;
    if (RegQueryValueExW(hijack, L"Debugger", nullptr, &debugger_type, nullptr,
                         &size) == ERROR_SUCCESS &&
        size > 0) {
      debugger.resize(size / sizeof(wchar_t));
      if (RegQueryValueExW(hijack, L"Debugger", nullptr, &debugger_type,
                           reinterpret_cast<LPBYTE>(debugger.data()),
                           &size) == ERROR_SUCCESS) {
        while (!debugger.empty() && debugger.back() == L'\0') {
          debugger.pop_back();
        }
        RegDeleteValueW(hijack, L"Debugger");
      } else {
        debugger.clear();
      }
    }
  }
  const HRESULT hr = util::IsProcessElevated()
                         ? win32::ShellOpen(hwnd_, regedit_path.c_str())
                         : win32::LaunchElevated(hwnd_, regedit_path, L"");
  if (hijack) {
    if (!debugger.empty()) {
      RegSetValueExW(hijack, L"Debugger", 0, debugger_type,
                     reinterpret_cast<const BYTE*>(debugger.c_str()),
                     static_cast<DWORD>((debugger.size() + 1) * sizeof(wchar_t)));
    }
    RegCloseKey(hijack);
  }
  if (FAILED(hr)) {
    if (!win32::DialogCancelled(hr)) {
      ui::ShowError(hwnd_, win32::FormatDialogError(hr));
    }
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
    ui::ShowError(hwnd_, L"Hive files aren't available for remote registries.");
    return;
  }
  std::wstring hive_path = ResolveSelectedHiveFilePath();
  if (hive_path.empty()) {
    ui::ShowError(hwnd_, L"No hive file was found for this key.");
    return;
  }
  const HRESULT hr = win32::RevealInExplorer(hive_path);
  if (FAILED(hr)) {
    ui::ShowError(hwnd_, win32::FormatDialogError(hr));
  }
}

LOGFONTW MainWindow::Impl::DefaultLogFont() const {
  return ui::DefaultUIFontLogFont();
}

} // namespace regkit
