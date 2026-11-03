// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/transfer_dialogs.h"

#include "dialogs/export_dialog.h"
#include "dialogs/hive_dialog.h"
#include "dialogs/value_editor.h"
#include "regfile/reg_file.h"
#include "regfile/registry_transfer.h"
#include "registry/key_algorithms.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "ui/feedback.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_view.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/dialog/text_prompt.h>
#include <win32ui/system/error_text.h>

#include <algorithm>
#include <vector>

namespace regkit
{

namespace
{

using win32ui::FormatWin32Error;

constexpr wchar_t kRegFileFilter[] = L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0";

bool ResolveExportKey(const std::wstring& key_path, REGSAM view, RegistryNode* node, std::wstring* display, std::wstring* error)
{
    *display = registry_path::Normalize(key_path, util::GetCurrentUserSidString());
    if (!registry_path::ParseRoot(*display, node) || !node->root)
    {
        if (error)
        {
            *error = util::Tr(L"Export supports the standard root keys only.");
        }
        return false;
    }
    node->view = view;
    return true;
}

bool WriteRegFile(const std::wstring& path, regfile::Writer&& writer, std::wstring* error)
{
    if (win32ui::WriteTextFileAtomic(path.c_str(), std::move(writer).Finish(), win32ui::TextEncoding::utf16) == ERROR_SUCCESS)
    {
        return true;
    }
    if (error)
    {
        *error = util::TrDetail(L"Failed to write the exported registry file.", path);
    }
    return false;
}

bool ReportUnreadableKey(LONG status, const std::wstring& display, std::wstring* error)
{
    if (error)
    {
        *error = win32ui::FormatWin32Error(static_cast<DWORD>(status)) + L"\n" + display;
    }
    return false;
}

}

bool ExportRegFile(HWND owner, const std::wstring& key_path, REGSAM view, bool allow_hive, std::wstring* error, std::wstring* saved_path, win32::OpenAfter* open_after)
{
    static win32::OpenAfter last_open_after = win32::OpenAfter::kNone;
    editors::ExportRequest request;
    request.path = DefaultExportPath(key_path, L".reg");
    request.open_after = last_open_after;
    request.allow_hive = allow_hive;
    editors::ExportResult options;
    if (!editors::ChooseExport(owner, request, &options))
    {
        return false;
    }
    RegistryNode node;
    std::wstring display;
    if (!ResolveExportKey(key_path, view, &node, &display, error))
    {
        return false;
    }
    if (options.hive)
    {
        const LONG saved = SaveKeyToHive(node.root, node.subkey, ViewOf(node), options.path);
        if (saved != ERROR_SUCCESS)
        {
            *error = HiveTransferError(saved, options.path);
            return false;
        }
        *saved_path = options.path;
        *open_after = win32::OpenAfter::kNone;
        return true;
    }
    last_open_after = options.open_after;
    regfile::Writer writer;
    std::vector<std::wstring> skipped;
    const LONG status = regfile::AppendRegistryTree(&writer, node.root, node.subkey, display, ViewOf(node), options.include_subkeys, &skipped, &skipped);
    if (status != ERROR_SUCCESS)
    {
        return ReportUnreadableKey(status, display, error);
    }
    if ((!skipped.empty() && !ui::ConfirmConversionSkips(owner, skipped, util::Tr(L"Export"))) || !WriteRegFile(options.path, std::move(writer), error))
    {
        return false;
    }
    *saved_path = options.path;
    *open_after = options.open_after;
    return true;
}

bool ExportRegFileSelection(HWND owner, const std::wstring& base_key_path, REGSAM view, const std::vector<std::wstring>& value_names, const std::vector<std::wstring>& subkey_names, std::wstring* error, std::wstring* saved_path, win32::OpenAfter* open_after)
{
    if (value_names.empty() && subkey_names.empty())
    {
        if (error)
        {
            *error = util::Tr(L"No data to export.");
        }
        return false;
    }
    const std::wstring first_name =
        !value_names.empty() ? (value_names.front().empty() ? L"Default" : value_names.front()) : subkey_names.front();
    std::wstring path;
    if (!ui::ReportFileDialogResult(
            owner,
            win32::ChooseFileToSave(owner, kRegFileFilter, ExportFileName(first_name, L".reg").c_str(), &path, open_after, true)
        ))
    {
        return false;
    }
    path = util::EnsureFileExtension(path, L".reg");

    regfile::Writer writer;
    std::vector<std::wstring> skipped;
    if (!value_names.empty())
    {
        RegistryNode base;
        std::wstring display;
        if (!ResolveExportKey(base_key_path, view, &base, &display, error))
        {
            return false;
        }
        registry_backend::KeyContents contents;
        const LONG status = registry_backend::ReadKeyContents(base.root, base.subkey, ViewOf(base), true, &contents);
        if (status != ERROR_SUCCESS)
        {
            return ReportUnreadableKey(status, display, error);
        }
        std::erase_if(contents.values, [&](const RegistryValue& value) {
            return std::none_of(value_names.begin(), value_names.end(), [&](const std::wstring& name) { return util::EqualsInsensitive(name, value.name); });
        });
        const bool found = !contents.values.empty();
        regfile::SkipNullNames(display, &contents.values, nullptr, &skipped);
        std::vector<const regfile::Value*> selected;
        for (const RegistryValue& value : contents.values)
        {
            selected.push_back(&value);
        }
        if (!found)
        {
            if (error)
            {
                *error = util::Tr(L"No selected values were found in the export.");
            }
            return false;
        }
        writer.AppendKey(display, std::move(selected), false);
    }
    for (const auto& subkey : subkey_names)
    {
        if (subkey.empty())
        {
            continue;
        }
        RegistryNode node;
        std::wstring display;
        if (!ResolveExportKey(base_key_path.empty() ? subkey : base_key_path + L"\\" + subkey, view, &node, &display, error))
        {
            return false;
        }
        const LONG status = regfile::AppendRegistryTree(&writer, node.root, node.subkey, display, ViewOf(node), true, &skipped, &skipped);
        if (status != ERROR_SUCCESS)
        {
            return ReportUnreadableKey(status, display, error);
        }
    }
    if (!skipped.empty() && !ui::ConfirmConversionSkips(owner, skipped, util::Tr(L"Export")))
    {
        return false;
    }
    *saved_path = path;
    return WriteRegFile(path, std::move(writer), error);
}

bool LoadHive(HWND owner, HKEY* root, std::wstring* error)
{
    if (!root)
    {
        return false;
    }
    editors::LoadHiveResult choice;
    choice.root = *root;
    if (!editors::ChooseHiveToLoad(owner, &choice))
    {
        return false;
    }
    *root = choice.root;
    const util::PrivilegeScope privileges({SE_RESTORE_NAME, SE_BACKUP_NAME});
    if (!privileges.held())
    {
        if (error)
        {
            *error = util::Tr(L"Loading a hive needs the backup and restore privileges. Run RegKit elevated.");
        }
        return false;
    }
    const LONG result = RegLoadKeyW(*root, choice.key_name.c_str(), choice.file.c_str());
    if (result != ERROR_SUCCESS)
    {
        if (error)
        {
            *error = FormatWin32Error(result);
        }
        return false;
    }
    return true;
}

bool UnloadHive(HWND owner, HKEY root, const std::wstring& subkey, std::wstring* error)
{
    std::wstring target = subkey;
    if (target.empty())
    {
        if (win32ui::PromptText(owner, {.title = util::Tr(L"Unload Hive"), .label = util::Tr(L"Key name:")}, &target) != 0)
        {
            return false;
        }
    }
    if (target.empty())
    {
        if (error)
        {
            *error = util::Tr(L"Key name is required.");
        }
        return false;
    }
    const util::PrivilegeScope privileges({SE_RESTORE_NAME, SE_BACKUP_NAME});
    if (!privileges.held())
    {
        if (error)
        {
            *error = util::Tr(L"Unloading a hive needs the backup and restore privileges. Run RegKit elevated.");
        }
        return false;
    }
    const LONG result = RegUnLoadKeyW(root, target.c_str());
    if (result != ERROR_SUCCESS)
    {
        if (error)
        {
            *error = FormatWin32Error(result);
        }
        return false;
    }
    return true;
}

}
