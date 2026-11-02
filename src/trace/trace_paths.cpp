// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "trace/trace_paths.h"

#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "win32/process_rights.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <cstring>
#include <cwctype>
#include <mutex>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace regkit::trace
{

namespace
{

using util::EqualsInsensitive;

const std::wstring& CurrentControlSet()
{
    static const std::wstring current = [] {
        RegistryNode node;
        node.root = HKEY_LOCAL_MACHINE;
        node.subkey = L"SYSTEM\\Select";
        RegistryValue entry;
        DWORD number = 0;
        if (!RegistryStore::QueryValue(node, L"Current", &entry) || entry.type != REG_DWORD || entry.data.size() < sizeof(number))
        {
            return std::wstring();
        }
        std::memcpy(&number, entry.data.data(), sizeof(number));
        wchar_t buffer[32] = {};
        swprintf_s(buffer, L"ControlSet%03u", number);
        return std::wstring(buffer);
    }();
    return current;
}

bool IsNumberedControlSet(const std::wstring& text)
{
    constexpr std::wstring_view kPrefix = L"ControlSet";
    return text.size() > kPrefix.size() && util::StartsWithInsensitive(text, kPrefix) &&
           std::all_of(text.begin() + kPrefix.size(), text.end(), [](wchar_t ch) { return iswdigit(ch) != 0; });
}

std::wstring WithCurrentControlSet(const std::wstring& path, bool numbered)
{
    const std::wstring& current = CurrentControlSet();
    std::vector<std::wstring> parts = registry_path::Split(path);
    if (current.empty() || parts.size() < 3 ||
        !(EqualsInsensitive(parts[0], L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(parts[0], L"HKLM") ||
          (EqualsInsensitive(parts[0], L"REGISTRY") && EqualsInsensitive(parts[1], L"MACHINE"))))
    {
        return {};
    }
    const size_t system = EqualsInsensitive(parts[0], L"REGISTRY") ? 2 : 1;
    if (parts.size() < system + 2 || !EqualsInsensitive(parts[system], L"SYSTEM"))
    {
        return {};
    }
    std::wstring& segment = parts[system + 1];
    if (!(numbered ? IsNumberedControlSet(segment) : EqualsInsensitive(segment, L"CurrentControlSet")) || EqualsInsensitive(segment, current))
    {
        return {};
    }
    segment = current;
    return registry_path::Join(parts);
}

std::wstring CleanKeyText(const std::wstring& text, const std::wstring& sid)
{
    std::wstring path = text;
    if (!sid.empty())
    {
        const std::wstring marker = L"<CURRENT_USER_SID>";
        for (size_t pos = path.find(marker); pos != std::wstring::npos; pos = path.find(marker, pos + sid.size()))
        {
            path.replace(pos, marker.size(), sid);
        }
    }
    path = registry_path::Clean(path);
    wchar_t machine[MAX_COMPUTERNAME_LENGTH + 1] = {};
    DWORD machine_length = static_cast<DWORD>(_countof(machine));
    if (!path.empty() && GetComputerNameW(machine, &machine_length) && machine_length > 0)
    {
        const std::wstring prefix = std::wstring(machine, machine_length) + L"\\";
        if (util::StartsWithInsensitive(path, prefix))
        {
            path.erase(0, prefix.size());
        }
    }
    return path;
}

bool LinkTarget(const std::wstring& path, const RegistryNode& node, std::wstring* target)
{
    static std::mutex mutex;
    static std::unordered_map<std::wstring, std::wstring> targets;
    static std::unordered_set<std::wstring> misses;
    static uint64_t revision = 0;
    std::wstring key = util::ToLower(path);
    {
        std::lock_guard<std::mutex> lock(mutex);
        // key changes and refreshes may add, remove or retarget links
        if (const uint64_t current = RegistryStore::KeyRevision(); current != revision)
        {
            targets.clear();
            misses.clear();
            revision = current;
        }
        if (const auto found = targets.find(key); found != targets.end())
        {
            *target = found->second;
            return true;
        }
        if (misses.contains(key))
        {
            return false;
        }
    }
    const bool resolved = RegistryStore::QuerySymbolicLinkTarget(node, target);
    std::lock_guard<std::mutex> lock(mutex);
    if (resolved)
    {
        targets.emplace(std::move(key), *target);
    }
    else
    {
        misses.insert(std::move(key));
    }
    return resolved;
}

std::wstring ResolveLinks(const std::wstring& path)
{
    std::wstring current = path;
    std::unordered_set<std::wstring> visited;
    for (int depth = 0; depth < 8 && visited.insert(util::ToLower(current)).second; ++depth)
    {
        RegistryNode root_node;
        if (!registry_path::ParseRoot(current, &root_node))
        {
            break;
        }
        const std::vector<std::wstring> parts = registry_path::Split(root_node.subkey);
        bool resolved = false;
        RegistryNode node = root_node;
        node.subkey.clear();
        for (size_t i = 0; i < parts.size() && !resolved; ++i)
        {
            node.subkey = registry_path::JoinSubkey(node.subkey, parts[i]);
            std::wstring target;
            if (!LinkTarget(root_node.root_name + L"\\" + node.subkey, node, &target))
            {
                continue;
            }
            const std::wstring mapped = NormalizeKeyPathBasic(target);
            if (mapped.empty())
            {
                continue;
            }
            const std::wstring remaining = registry_path::Join(parts, i + 1);
            current = remaining.empty() ? mapped : mapped + L"\\" + remaining;
            resolved = true;
        }
        if (!resolved)
        {
            break;
        }
    }
    return current;
}

}

std::wstring MapControlSetToCurrent(const std::wstring& path)
{
    return WithCurrentControlSet(path, true);
}

std::wstring NormalizeKeyPathBasic(const std::wstring& text)
{
    const std::wstring sid = util::GetCurrentUserSidString();
    std::wstring path = CleanKeyText(text, sid);
    if (path.empty())
    {
        return {};
    }
    // traces keep showing native machine class reads under the merged HKCR view
    constexpr std::wstring_view kMachineClasses = L"REGISTRY\\MACHINE\\SOFTWARE\\Classes";
    if (registry_path::HasComponentPrefix(path, kMachineClasses))
    {
        path.replace(0, kMachineClasses.size(), L"HKEY_CLASSES_ROOT");
    }
    path = registry_path::Normalize(path, sid);
    const std::wstring current_user = L"HKEY_USERS\\" + sid;
    if (!sid.empty() && registry_path::HasComponentPrefix(path, current_user))
    {
        path.replace(0, current_user.size(), L"HKEY_CURRENT_USER");
    }
    RegistryNode node;
    if (!registry_path::ParseRoot(path, &node))
    {
        return {};
    }
    const std::wstring replaced = WithCurrentControlSet(path, false);
    return replaced.empty() ? path : replaced;
}

std::wstring NormalizeKeyPath(const std::wstring& text)
{
    const std::wstring path = NormalizeKeyPathBasic(text);
    return path.empty() ? path : ResolveLinks(path);
}

std::wstring NormalizeSelectionPath(const std::wstring& text)
{
    const std::wstring path = CleanKeyText(text, util::GetCurrentUserSidString());
    if (!util::StartsWithInsensitive(path, L"REGISTRY"))
    {
        return path;
    }
    const size_t rest = path.find_first_not_of(L'\\', 8);
    return rest == std::wstring::npos ? L"REGISTRY" : L"REGISTRY\\" + path.substr(rest);
}

Normalizers PathNormalizers()
{
    Normalizers normalizers;
    normalizers.key = NormalizeKeyPath;
    normalizers.display = NormalizeSelectionPath;
    return normalizers;
}

}
