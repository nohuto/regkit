// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/shell_integration.h"
#include "win32/text_transform.h"

#include "win32/process_rights.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/registry_value.h>

#include <shlobj.h>

#include <cwctype>

namespace regkit::win32
{
namespace
{

struct EditMenu
{
    const wchar_t* key;
    const wchar_t* applies_to;
};

// use regfile ProgID instead of SystemFileAssociations for W7 support
constexpr EditMenu kEditMenus[] = {
    {L"Software\\Classes\\regfile\\shell\\RegKit.Edit", nullptr},
    {L"Software\\Classes\\SystemFileAssociations\\.hiv\\shell\\RegKit.Edit", nullptr},
    {L"Software\\Classes\\SystemFileAssociations\\.hve\\shell\\RegKit.Edit", nullptr},
    {L"Software\\Classes\\*\\shell\\RegKit.Edit",
     L"System.FileName:=\"NTUSER.DAT\" OR System.FileName:=\"UsrClass.dat\" OR System.FileName:=\"SYSTEM\" OR System.FileName:=\"SOFTWARE\" OR "
     L"System.FileName:=\"SAM\" OR System.FileName:=\"SECURITY\" OR System.FileName:=\"DEFAULT\" OR System.FileName:=\"COMPONENTS\" OR "
     L"System.FileName:=\"DRIVERS\" OR System.FileName:=\"BCD\""},
};
constexpr wchar_t kEditMenuLabel[] = L"Edit with RegKit";
constexpr wchar_t kRegEditImageOptionsKey[] =
    L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\regedit.exe";

bool Missing(LONG result)
{
    return result == ERROR_FILE_NOT_FOUND || result == ERROR_PATH_NOT_FOUND;
}

std::wstring EditMenuCommand(const std::wstring& exe_path)
{
    return L"\"" + exe_path + L"\" --edit-reg \"%1\"";
}

std::wstring CommandKey(const EditMenu& menu)
{
    return std::wstring(menu.key) + L"\\command";
}

bool RegistryStringEquals(const wchar_t* subkey, const wchar_t* value_name, const std::wstring& expected)
{
    std::wstring value;
    return win32ui::ReadRegistryString(HKEY_CURRENT_USER, subkey, value_name, &value, RRF_NOEXPAND) == ERROR_SUCCESS &&
           util::EqualsInsensitive(value, expected);
}

LONG DeleteEditMenu()
{
    LONG first_error = ERROR_SUCCESS;
    for (const EditMenu& menu : kEditMenus)
    {
        LONG result = RegDeleteTreeW(HKEY_CURRENT_USER, menu.key);
        if (result == ERROR_SUCCESS)
        {
            result = RegDeleteKeyW(HKEY_CURRENT_USER, menu.key);
        }
        if (result != ERROR_SUCCESS && !Missing(result) && first_error == ERROR_SUCCESS)
        {
            first_error = result;
        }
    }
    return first_error;
}

bool IsEditMenuCommandOwned(const std::wstring& exe_path)
{
    return !exe_path.empty() && RegistryStringEquals(CommandKey(kEditMenus[0]).c_str(), nullptr, EditMenuCommand(exe_path));
}

bool OwnsRegEditDebugger(const std::wstring& debugger, const std::wstring& exe_path)
{
    // compare only the debugger executable and allow arguments after it
    const wchar_t* start = debugger.c_str();
    while (*start && iswspace(*start))
    {
        ++start;
    }
    const bool quoted = *start == L'\"';
    start += quoted ? 1 : 0;
    const wchar_t* end = start;
    while (*end && (quoted ? *end != L'\"' : !iswspace(*end)))
    {
        ++end;
    }
    return end != start && exe_path.size() == static_cast<size_t>(end - start) &&
           util::StartsWithInsensitive(start, exe_path);
}

LONG ReadRegEditDebugger(std::wstring* debugger)
{
    return win32ui::ReadRegistryString(HKEY_LOCAL_MACHINE, kRegEditImageOptionsKey, L"Debugger", debugger, RRF_NOEXPAND);
}

LONG WriteRegEditDebugger(const std::wstring& exe_path)
{
    return win32ui::WriteRegistryString(HKEY_LOCAL_MACHINE, kRegEditImageOptionsKey, L"Debugger", (L"\"" + exe_path + L"\"").c_str());
}

LONG DeleteOwnedRegEditDebugger(const std::wstring& exe_path)
{
    std::wstring debugger;
    LONG result = ReadRegEditDebugger(&debugger);
    if (result != ERROR_SUCCESS)
    {
        return Missing(result) ? ERROR_SUCCESS : result;
    }
    // leave debugger entries owned by other programs unchanged
    if (!OwnsRegEditDebugger(debugger, exe_path))
    {
        return ERROR_SUCCESS;
    }
    result = RegDeleteKeyValueW(HKEY_LOCAL_MACHINE, kRegEditImageOptionsKey, L"Debugger");
    if (result != ERROR_SUCCESS && !Missing(result))
    {
        return result;
    }
    win32ui::UniqueHKey key;
    DWORD subkeys = 0;
    DWORD values = 0;
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, kRegEditImageOptionsKey, 0, KEY_QUERY_VALUE, key.put()) == ERROR_SUCCESS &&
        RegQueryInfoKeyW(key.get(), nullptr, nullptr, nullptr, &subkeys, nullptr, nullptr, &values, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS &&
        subkeys == 0 && values == 0)
    {
        RegDeleteKeyW(HKEY_LOCAL_MACHINE, kRegEditImageOptionsKey);
    }
    return ERROR_SUCCESS;
}

}

bool IsEditMenuRegistered(const std::wstring& exe_path)
{
    for (const EditMenu& menu : kEditMenus)
    {
        std::wstring label;
        if (exe_path.empty() || !RegistryStringEquals(CommandKey(menu).c_str(), nullptr, EditMenuCommand(exe_path)) ||
            win32ui::ReadRegistryString(HKEY_CURRENT_USER, menu.key, nullptr, &label, RRF_NOEXPAND) != ERROR_SUCCESS || label != kEditMenuLabel ||
            !RegistryStringEquals(menu.key, L"Icon", exe_path + L",0"))
        {
            return false;
        }
    }
    return true;
}

LONG SetEditMenu(const std::wstring& exe_path, bool enable, LONG* cleanup_error)
{
    if (cleanup_error)
    {
        *cleanup_error = ERROR_SUCCESS;
    }
    if (exe_path.empty())
    {
        return ERROR_INVALID_PARAMETER;
    }
    LONG result = ERROR_SUCCESS;
    if (enable)
    {
        for (const EditMenu& menu : kEditMenus)
        {
            result = win32ui::WriteRegistryString(HKEY_CURRENT_USER, menu.key, nullptr, kEditMenuLabel);
            if (result == ERROR_SUCCESS)
            {
                result = win32ui::WriteRegistryString(HKEY_CURRENT_USER, menu.key, L"Icon", (exe_path + L",0").c_str());
            }
            if (result == ERROR_SUCCESS && menu.applies_to)
            {
                result = win32ui::WriteRegistryString(HKEY_CURRENT_USER, menu.key, L"AppliesTo", menu.applies_to);
            }
            if (result == ERROR_SUCCESS)
            {
                result = win32ui::WriteRegistryString(HKEY_CURRENT_USER, CommandKey(menu).c_str(), nullptr, EditMenuCommand(exe_path).c_str());
            }
            if (result != ERROR_SUCCESS)
            {
                break;
            }
        }
        if (result != ERROR_SUCCESS)
        {
            // remove partial registration when any write fails
            const LONG cleanup = DeleteEditMenu();
            if (cleanup_error)
            {
                *cleanup_error = cleanup;
            }
        }
    }
    else
    {
        result = DeleteEditMenu();
    }
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
    return result;
}

LONG RemoveEditMenuIfOwned(const std::wstring& exe_path)
{
    // uninstall only the command that still points to this executable
    return IsEditMenuCommandOwned(exe_path) ? SetEditMenu(exe_path, false) : ERROR_SUCCESS;
}

bool IsRegEditReplacementRegistered(const std::wstring& exe_path)
{
    std::wstring debugger;
    return !exe_path.empty() && ReadRegEditDebugger(&debugger) == ERROR_SUCCESS &&
           OwnsRegEditDebugger(debugger, exe_path);
}

LONG SetRegEditReplacement(const std::wstring& exe_path, bool enable, bool* conflict, bool overwrite_existing, bool allow_writable_location)
{
    if (conflict)
    {
        *conflict = false;
    }
    if (exe_path.empty())
    {
        return ERROR_INVALID_PARAMETER;
    }
    if (!enable)
    {
        return DeleteOwnedRegEditDebugger(exe_path);
    }
    if (!allow_writable_location && util::IsWritableByNonAdmins(exe_path))
    {
        return ERROR_ACCESS_DENIED;
    }
    std::wstring debugger;
    const LONG result = overwrite_existing ? ERROR_FILE_NOT_FOUND : ReadRegEditDebugger(&debugger);
    if (Missing(result))
    {
        return WriteRegEditDebugger(exe_path);
    }
    if (result == ERROR_SUCCESS && OwnsRegEditDebugger(debugger, exe_path))
    {
        return ERROR_SUCCESS;
    }
    if (result != ERROR_SUCCESS && result != ERROR_UNSUPPORTED_TYPE && result != ERROR_INVALID_DATA)
    {
        return result;
    }
    if (conflict)
    {
        *conflict = true;
    }
    return ERROR_ALREADY_EXISTS;
}

}
