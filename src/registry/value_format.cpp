// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/value_format.h"

#include "win32/process_rights.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/system/paths.h>
#include <win32ui/system/registry_value.h>
#include <win32ui/system/strings.h>

#include <shlobj.h>
#include <shlwapi.h>

#include <algorithm>
#include <cstring>
#include <cwctype>

namespace regkit::value_format
{

namespace
{

constexpr TypeLabel kTypeLabels[] = {
    {REG_SZ, L"REG_SZ"},
    {REG_EXPAND_SZ, L"REG_EXPAND_SZ"},
    {REG_MULTI_SZ, L"REG_MULTI_SZ"},
    {REG_LINK, L"REG_LINK"},
    {REG_DWORD, L"REG_DWORD"},
    {REG_DWORD_BIG_ENDIAN, L"REG_DWORD_BIG_ENDIAN"},
    {REG_QWORD, L"REG_QWORD"},
    {REG_BINARY, L"REG_BINARY"},
    {REG_RESOURCE_LIST, L"REG_RESOURCE_LIST"},
    {REG_FULL_RESOURCE_DESCRIPTOR, L"REG_FULL_RESOURCE_DESCRIPTOR"},
    {REG_RESOURCE_REQUIREMENTS_LIST, L"REG_RESOURCE_REQUIREMENTS_LIST"},
    {REG_NONE, L"REG_NONE"},
};

const TypeLabel* FindTypeLabel(DWORD type)
{
    for (const TypeLabel& label : kTypeLabels)
    {
        if (label.type == (type & 0xFFFF))
        {
            return &label;
        }
    }
    return nullptr;
}

}

std::span<const TypeLabel> TypeLabels()
{
    return kTypeLabels;
}

DWORD NormalizeType(DWORD type)
{
    const TypeLabel* label = FindTypeLabel(type);
    return label ? label->type : type;
}

std::wstring TypeName(DWORD type)
{
    const TypeLabel* label = FindTypeLabel(type);
    if (label && label->type == type)
    {
        return label->name;
    }
    return std::wstring(label ? label->name : L"REG_UNKNOWN") + L" (" + win32ui::FormatHex(type) + L")";
}

std::wstring ByteCount(size_t size)
{
    return std::to_wstring(size) + L" " + (size == 1 ? util::Tr(L"byte") : util::Tr(L"bytes"));
}

std::wstring Data(DWORD type, const BYTE* data, DWORD size)
{
    if (!data || size == 0)
    {
        return {};
    }
    switch (NormalizeType(type))
    {
    case REG_SZ:
    case REG_EXPAND_SZ:
    case REG_LINK:
        {
            std::wstring text(reinterpret_cast<const wchar_t*>(data), size / sizeof(wchar_t));
            while (!text.empty() && text.back() == L'\0')
            {
                text.pop_back();
            }
            return text;
        }
    case REG_MULTI_SZ:
        {
            std::wstring joined;
            for (const std::wstring& item : MultiStringItems({data, size - size % sizeof(wchar_t)}))
            {
                if (!joined.empty())
                {
                    joined += L' ';
                }
                joined += item;
            }
            return joined;
        }
    case REG_DWORD:
    case REG_DWORD_BIG_ENDIAN:
        if (size >= sizeof(DWORD))
        {
            DWORD value = 0;
            std::memcpy(&value, data, sizeof(value));
            if (NormalizeType(type) == REG_DWORD_BIG_ENDIAN)
            {
                value = _byteswap_ulong(value);
            }
            wchar_t buffer[32] = {};
            swprintf_s(buffer, L"0x%08X (%u)", value, value);
            return buffer;
        }
        break;
    case REG_QWORD:
        if (size >= sizeof(unsigned long long))
        {
            unsigned long long value = 0;
            std::memcpy(&value, data, sizeof(value));
            wchar_t buffer[48] = {};
            swprintf_s(buffer, L"0x%016llX (%llu)", value, value);
            return buffer;
        }
        break;
    default:
        break;
    }
    return win32ui::HexBytes({data, size});
}

namespace
{

bool IsTrustedResourcePath(const wchar_t* full)
{
    static const std::vector<std::wstring> roots = [] {
        std::vector<std::wstring> list;
        for (const KNOWNFOLDERID& folder : {FOLDERID_Windows, FOLDERID_ProgramFiles, FOLDERID_ProgramFilesX86})
        {
            if (std::wstring root = win32ui::KnownFolderPath(folder); !root.empty())
            {
                list.push_back(std::move(root));
            }
        }
        if (std::wstring native; win32ui::ReadRegistryString(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion", L"ProgramW6432Dir", &native, RRF_SUBKEY_WOW6464KEY) == ERROR_SUCCESS)
        {
            list.push_back(std::move(native));
        }
        return list;
    }();
    return std::any_of(roots.begin(), roots.end(), [&](const std::wstring& root) {
        return util::StartsWithInsensitive(full, root) && full[root.size()] == L'\\';
    });
}

std::wstring TrustedIndirectString(const std::wstring& value)
{
    const bool package = value.size() > 1 && value[1] == L'{';
    const size_t start = package ? 2 : 1;
    const size_t end = package ? value.find(L'?') : value.find_last_of(L',');
    std::wstring source = win32ui::ExpandEnvironment(value.substr(start, end == std::wstring::npos ? std::wstring::npos : end - start).c_str());
    if (package && (source.size() < 2 || source[1] != L':'))
    {
        return value;
    }
    // only module name = System32
    if (source.find_first_of(L"\\/:") == std::wstring::npos)
    {
        wchar_t system[MAX_PATH] = {};
        const UINT length = GetSystemDirectoryW(system, MAX_PATH);
        source.insert(0, std::wstring(system, length < MAX_PATH ? length : 0) + L'\\');
    }
    if (!win32ui::IsDriveAbsolute(source))
    {
        return {};
    }
    wchar_t full[MAX_PATH] = {};
    const DWORD length = GetFullPathNameW(source.c_str(), MAX_PATH, full, nullptr);
    if (length == 0 || length >= MAX_PATH)
    {
        return {};
    }
    if (!win32ui::OnFixedDrive(full) || (util::IsProcessPrivileged() && !IsTrustedResourcePath(full)))
    {
        return {};
    }
    return value.substr(0, start) + full + (end == std::wstring::npos ? std::wstring() : value.substr(end));
}

}

std::wstring DisplayData(DWORD type, const BYTE* data, DWORD size, bool resolve_indirect)
{
    const DWORD base_type = NormalizeType(type);
    std::wstring value = Data(type, data, size);
    if (value.empty())
    {
        return value;
    }
    const std::wstring indirect =
        resolve_indirect && (base_type == REG_SZ || base_type == REG_EXPAND_SZ) && value.front() == L'@'
            ? TrustedIndirectString(value)
            : std::wstring();
    if (!indirect.empty())
    {
        std::wstring resolved(1024, L'\0');
        HRESULT result =
            SHLoadIndirectString(indirect.c_str(), resolved.data(), static_cast<UINT>(resolved.size()), nullptr);
        if (result == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER))
        {
            resolved.assign(4096, L'\0');
            result =
                SHLoadIndirectString(indirect.c_str(), resolved.data(), static_cast<UINT>(resolved.size()), nullptr);
        }
        if (SUCCEEDED(result))
        {
            while (!resolved.empty() && resolved.back() == L'\0')
            {
                resolved.pop_back();
            }
            if (!resolved.empty())
            {
                return resolved;
            }
        }
    }
    if (base_type == REG_EXPAND_SZ)
    {
        std::wstring expanded = win32ui::ExpandEnvironment(value.c_str());
        if (!expanded.empty() && expanded != value)
        {
            return expanded;
        }
    }
    return value;
}

uint64_t ReadUnsigned(std::span<const BYTE> data, size_t width, bool big_endian)
{
    if (width == 0 || width > sizeof(uint64_t) || data.size() < width)
    {
        return 0;
    }
    uint64_t value = 0;
    for (size_t index = 0; index < width; ++index)
    {
        value = (value << 8) | data[big_endian ? index : width - 1 - index];
    }
    return value;
}

std::vector<BYTE> UnsignedBytes(uint64_t value, size_t width, bool big_endian)
{
    std::vector<BYTE> bytes(width, 0);
    for (size_t index = 0; index < width && index < sizeof(uint64_t); ++index)
    {
        bytes[big_endian ? width - 1 - index : index] = static_cast<BYTE>(value >> (8 * index));
    }
    return bytes;
}

std::vector<BYTE> StringData(std::wstring_view text)
{
    std::vector<BYTE> data((text.size() + 1) * sizeof(wchar_t));
    std::memcpy(data.data(), text.data(), text.size() * sizeof(wchar_t));
    return data;
}

bool DecodeString(std::span<const BYTE> data, std::wstring* output)
{
    if (!output || data.size() % sizeof(wchar_t) != 0)
    {
        return false;
    }
    output->assign(reinterpret_cast<const wchar_t*>(data.data()), data.size() / sizeof(wchar_t));
    while (!output->empty() && output->back() == L'\0')
    {
        output->pop_back();
    }
    return output->find(L'\0') == std::wstring::npos;
}

std::vector<std::wstring> MultiStringItems(std::span<const BYTE> data)
{
    std::vector<std::wstring> items;
    if (data.size() % sizeof(wchar_t) != 0)
    {
        return items;
    }
    const wchar_t* current = reinterpret_cast<const wchar_t*>(data.data());
    size_t remaining = data.size() / sizeof(wchar_t);
    while (remaining > 0 && *current)
    {
        const size_t length = wcsnlen_s(current, remaining);
        items.emplace_back(current, length);
        if (length == remaining)
        {
            break;
        }
        current += length + 1;
        remaining -= length + 1;
    }
    return items;
}

std::vector<BYTE> MultiStringData(const std::vector<std::wstring>& items)
{
    size_t characters = 1;
    for (const auto& item : items)
    {
        characters += item.size() + 1;
    }
    std::vector<BYTE> data(characters * sizeof(wchar_t), 0);
    wchar_t* output = reinterpret_cast<wchar_t*>(data.data());
    for (const auto& item : items)
    {
        std::memcpy(output, item.data(), item.size() * sizeof(wchar_t));
        output += item.size() + 1;
    }
    return data;
}

std::wstring MultiStringText(const std::vector<BYTE>& data)
{
    std::wstring text;
    for (const auto& item : MultiStringItems(data))
    {
        text += item;
        text += L"\r\n";
    }
    return text;
}

std::vector<BYTE> MultiStringData(std::wstring_view lines)
{
    std::vector<std::wstring> items;
    for (const std::wstring_view line : win32ui::LineViews(lines))
    {
        if (!line.empty())
        {
            items.emplace_back(line);
        }
    }
    return MultiStringData(items);
}

}
