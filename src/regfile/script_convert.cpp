// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "regfile/script_convert.h"

#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/file_text.h"
#include "win32/registry_view.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/system/error_text.h>

#include <algorithm>

namespace regkit::regfile
{

bool FormatFromPath(std::wstring_view path, Format* format)
{
    if (util::EndsWithInsensitive(path, L".reg"))
    {
        *format = Format::kReg;
    }
    else if (util::EndsWithInsensitive(path, L".bat") || util::EndsWithInsensitive(path, L".cmd"))
    {
        *format = Format::kBatch;
    }
    else if (util::EndsWithInsensitive(path, L".ps1"))
    {
        *format = Format::kPowerShell;
    }
    else
    {
        return false;
    }
    return true;
}

const wchar_t* FormatExtension(Format format)
{
    return format == Format::kBatch ? L".bat" : format == Format::kPowerShell ? L".ps1"
                                                                              : L".reg";
}

bool ReadOperations(const std::wstring& path, std::vector<Operation>* operations, std::wstring* error)
{
    Format format = Format::kReg;
    if (!FormatFromPath(path, &format))
    {
        *error = util::TrDetail(L"Unsupported file type, use .reg, .bat, .cmd or .ps1.", path);
        return false;
    }
    return ReadOperations(path, format, operations, error);
}

bool ReadOperations(const std::wstring& path, Format format, std::vector<Operation>* operations, std::wstring* error)
{
    std::wstring content;
    if (!ReadText(path, &content, error))
    {
        return false;
    }
    const bool parsed = format == Format::kReg     ? ParseOperations(content, operations, nullptr, nullptr, error)
                        : format == Format::kBatch ? ParseBatch(content, operations, error)
                                                   : ParsePowerShell(content, operations, error);
    for (size_t index = 0; parsed && index < operations->size(); ++index)
    {
        std::wstring normalized;
        if (!NormalizeKeyPath((*operations)[index].path, &normalized))
        {
            *error = util::TrDetail(L"Unsupported root key.", (*operations)[index].path);
            return false;
        }
        (*operations)[index].path = std::move(normalized);
    }
    if (parsed && operations->empty())
    {
        *error = util::TrDetail(L"The file has no registry changes.", path);
        return false;
    }
    return parsed;
}

bool ReadRegistry(const std::wstring& key_path, bool recursive, std::vector<Operation>* operations, std::wstring* error, std::vector<std::wstring>* skipped)
{
    std::wstring path;
    if (!NormalizeKeyPath(registry_path::Normalize(key_path), &path))
    {
        *error = util::Tr(L"Enter a key under HKLM, HKCU, HKCR, HKU or HKCC.");
        return false;
    }
    const size_t split = path.find(L'\\');
    const std::wstring subkey = split == std::wstring::npos ? std::wstring() : path.substr(split + 1);
    const LONG status = ReadRegistryOperations(registry_path::RootFromName(path.substr(0, split)), subkey, path, win32::kDefaultRegistryView, recursive, operations, skipped, skipped);
    if (status != ERROR_SUCCESS)
    {
        *error = win32ui::FormatWin32Error(static_cast<DWORD>(status)) + L"\n" + path;
        return false;
    }
    return true;
}

bool SelectKey(const std::wstring& key_path, bool recursive, std::vector<Operation>* operations, std::wstring* error)
{
    std::wstring key;
    if (!NormalizeKeyPath(key_path, &key))
    {
        *error = util::TrDetail(L"Invalid key name.", key_path);
        return false;
    }
    std::erase_if(*operations, [&](const Operation& operation) {
        return recursive ? !IsUnderKey(operation.path, key) : !util::EqualsInsensitive(operation.path, key);
    });
    if (operations->empty())
    {
        *error = util::TrDetail(L"The file has no changes for this key.", key);
        return false;
    }
    return true;
}

std::wstring RenderOperations(Format format, const std::vector<Operation>& operations, bool admin_check, std::vector<std::wstring>* skipped)
{
    switch (format)
    {
    case Format::kBatch:
        return RenderBatch(operations, admin_check, skipped);
    case Format::kPowerShell:
        return RenderPowerShell(operations, admin_check, skipped);
    case Format::kReg:
        break;
    }
    return RenderReg(operations);
}

bool SaveRendered(const std::wstring& path, Format format, const std::wstring& text)
{
    const bool ascii = std::all_of(text.begin(), text.end(), [](wchar_t character) { return character < 0x80; });
    return win32ui::WriteTextFileAtomic(path.c_str(), text, format == Format::kReg || (format == Format::kPowerShell && !ascii)) == ERROR_SUCCESS;
}

bool NormalizeKeyPath(std::wstring_view text, std::wstring* path)
{
    while (!text.empty() && text.back() == L'\\')
    {
        text.remove_suffix(1);
    }
    const size_t split = text.find(L'\\');
    const HKEY root = registry_path::RootFromName(text.substr(0, split));
    if (root != HKEY_CLASSES_ROOT && root != HKEY_CURRENT_USER && root != HKEY_LOCAL_MACHINE && root != HKEY_USERS &&
        root != HKEY_CURRENT_CONFIG)
    {
        return false;
    }
    *path = registry_path::RootName(root);
    if (split != std::wstring_view::npos)
    {
        path->append(text.substr(split));
    }
    return true;
}

bool IsUnderKey(std::wstring_view path, std::wstring_view parent)
{
    return util::StartsWithInsensitive(path, parent) && (path.size() == parent.size() || path[parent.size()] == L'\\');
}

// regedit imports run elevated
bool NeedsAdmin(const std::vector<Operation>& operations)
{
    return std::any_of(operations.begin(), operations.end(), [](const Operation& operation) {
        return !IsUnderKey(operation.path, L"HKEY_CURRENT_USER");
    });
}

std::vector<std::wstring> MultiStringItems(const std::vector<BYTE>& data)
{
    const std::vector<std::wstring> empty_item{std::wstring()};
    return data == value_format::MultiStringData(empty_item) ? empty_item : value_format::MultiStringItems(data);
}

std::wstring Describe(const Operation& operation, std::wstring_view reason)
{
    std::wstring text = registry_path::Format(operation.path, registry_path::Style::kAbbreviated);
    if (operation.kind == Operation::Kind::kValue || operation.kind == Operation::Kind::kRemoveValue)
    {
        text += L" : " + (operation.value.name.empty() ? std::wstring(util::Tr(L"(Default)")) : operation.value.name);
    }
    return registry_path::DisplayName(text) + L" (" + std::wstring(reason) + L")";
}

}
