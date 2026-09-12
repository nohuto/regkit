// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/shell_integration.h"

#include "win32/registry_native.h"

#include <shlobj.h>

namespace regkit::win32 {
namespace {

constexpr wchar_t kEditMenuKey[] =
    L"Software\\Classes\\SystemFileAssociations\\.reg\\shell\\RegKit.Edit";
constexpr wchar_t kEditMenuCommandKey[] =
    L"Software\\Classes\\SystemFileAssociations\\.reg\\shell\\RegKit.Edit\\command";

std::wstring EditMenuCommand(
    const std::wstring& exe_path
) {
  return L"\"" + exe_path + L"\" --edit-reg \"%1\"";
}

std::wstring EditMenuIcon(
    const std::wstring& exe_path
) {
  return exe_path + L",0";
}

LONG DeleteEditMenu() {
  const LONG result = RegDeleteTreeW(HKEY_CURRENT_USER, kEditMenuKey);
  return result == ERROR_FILE_NOT_FOUND || result == ERROR_PATH_NOT_FOUND
             ? ERROR_SUCCESS
             : result;
}

bool IsEditMenuCommandOwned(
    const std::wstring& exe_path
) {
  std::wstring command;
  return !exe_path.empty() &&
         util::ReadRegistryString(
             HKEY_CURRENT_USER,
             kEditMenuCommandKey,
             nullptr,
             &command
         ) &&
         _wcsicmp(command.c_str(), EditMenuCommand(exe_path).c_str()) == 0;
}

} // namespace

bool IsRegFileEditMenuRegistered(
    const std::wstring& exe_path
) {
  std::wstring label;
  std::wstring icon;
  return IsEditMenuCommandOwned(exe_path) &&
         util::ReadRegistryString(
             HKEY_CURRENT_USER,
             kEditMenuKey,
             nullptr,
             &label
         ) &&
         util::ReadRegistryString(
             HKEY_CURRENT_USER,
             kEditMenuKey,
             L"Icon",
             &icon
         ) &&
         label == L"Edit with RegKit" &&
         _wcsicmp(icon.c_str(), EditMenuIcon(exe_path).c_str()) == 0;
}

LONG SetRegFileEditMenu(
    const std::wstring& exe_path,
    bool enable,
    LONG* cleanup_error
) {
  if (cleanup_error) {
    *cleanup_error = ERROR_SUCCESS;
  }
  if (exe_path.empty()) {
    return ERROR_INVALID_PARAMETER;
  }

  LONG result = ERROR_SUCCESS;
  if (enable) {
    util::UniqueHKey verb_key;
    result = RegCreateKeyExW(
        HKEY_CURRENT_USER,
        kEditMenuKey,
        0,
        nullptr,
        REG_OPTION_NON_VOLATILE,
        KEY_SET_VALUE,
        nullptr,
        verb_key.put(),
        nullptr
    );
    if (result == ERROR_SUCCESS) {
      result = util::WriteRegistryString(
          verb_key.get(),
          nullptr,
          L"Edit with RegKit"
      );
    }
    if (result == ERROR_SUCCESS) {
      result = util::WriteRegistryString(
          verb_key.get(),
          L"Icon",
          EditMenuIcon(exe_path)
      );
    }

    util::UniqueHKey command_key;
    if (result == ERROR_SUCCESS) {
      result = RegCreateKeyExW(
          HKEY_CURRENT_USER,
          kEditMenuCommandKey,
          0,
          nullptr,
          REG_OPTION_NON_VOLATILE,
          KEY_SET_VALUE,
          nullptr,
          command_key.put(),
          nullptr
      );
    }
    if (result == ERROR_SUCCESS) {
      result = util::WriteRegistryString(
          command_key.get(),
          nullptr,
          EditMenuCommand(exe_path)
      );
    }
    if (result != ERROR_SUCCESS) {
      command_key.reset();
      verb_key.reset();
      const LONG cleanup = DeleteEditMenu();
      if (cleanup_error) {
        *cleanup_error = cleanup;
      }
    }
  } else {
    result = DeleteEditMenu();
  }

  SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
  return result;
}

LONG RemoveRegFileEditMenuIfOwned(
    const std::wstring& exe_path
) {
  if (!IsEditMenuCommandOwned(exe_path)) {
    return ERROR_SUCCESS;
  }
  return SetRegFileEditMenu(exe_path, false);
}

} // namespace regkit::win32
