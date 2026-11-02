// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_path.h"

#include "registry/registry_store.h"
#include "win32/process_rights.h"
#include "win32/text_transform.h"

#include <algorithm>
#include <array>
#include <cwctype>

namespace regkit::registry_path
{
namespace
{

struct RootNames
{
    HKEY root;
    const wchar_t* name;
    const wchar_t* abbreviation;
};

// keep normal path parsing limited to five standard roots
constexpr size_t kBrowsableRoots = 5;

const std::array<RootNames, 8>& Roots()
{
    static const std::array<RootNames, 8> roots = {{
        {HKEY_CLASSES_ROOT, L"HKEY_CLASSES_ROOT", L"HKCR"},
        {HKEY_CURRENT_USER, L"HKEY_CURRENT_USER", L"HKCU"},
        {HKEY_LOCAL_MACHINE, L"HKEY_LOCAL_MACHINE", L"HKLM"},
        {HKEY_USERS, L"HKEY_USERS", L"HKU"},
        {HKEY_CURRENT_CONFIG, L"HKEY_CURRENT_CONFIG", L"HKCC"},
        {HKEY_PERFORMANCE_DATA, L"HKEY_PERFORMANCE_DATA", L"HKPD"},
        {HKEY_PERFORMANCE_TEXT, L"HKEY_PERFORMANCE_TEXT", L""},
        {HKEY_PERFORMANCE_NLSTEXT, L"HKEY_PERFORMANCE_NLSTEXT", L""},
    }};
    return roots;
}

const RootNames* FindRoot(std::wstring_view name, size_t count)
{
    for (size_t index = 0; index < count && !name.empty(); ++index)
    {
        const RootNames& entry = Roots()[index];
        if (util::EqualsInsensitive(name, entry.name) || util::EqualsInsensitive(name, entry.abbreviation))
        {
            return &entry;
        }
    }
    return nullptr;
}

std::wstring CanonicalRoot(std::wstring_view root)
{
    if (util::EqualsInsensitive(root, L"MACHINE"))
    {
        return L"HKEY_LOCAL_MACHINE";
    }
    if (util::EqualsInsensitive(root, L"USER") || util::EqualsInsensitive(root, L"USERS"))
    {
        return L"HKEY_USERS";
    }
    if (util::EqualsInsensitive(root, L"REGISTRY"))
    {
        return L"REGISTRY";
    }
    const RootNames* entry = FindRoot(root, kBrowsableRoots);
    return entry ? entry->name : L"";
}

std::wstring AbbreviatedRoot(std::wstring_view root)
{
    const RootNames* entry = FindRoot(root, kBrowsableRoots);
    return entry ? entry->abbreviation : std::wstring(root);
}
std::wstring Join(std::wstring_view root, std::wstring_view rest)
{
    if (rest.empty())
    {
        return std::wstring(root);
    }
    std::wstring result(root);
    result.push_back(L'\\');
    result.append(rest);
    return result;
}

std::wstring_view TrimLeadingSeparators(std::wstring_view text)
{
    while (!text.empty() && text.front() == L'\\')
    {
        text.remove_prefix(1);
    }
    return text;
}

std::wstring JoinRange(const std::vector<std::wstring>& parts, size_t first, size_t last)
{
    size_t characters = 0;
    for (size_t index = first; index < last; ++index)
    {
        characters += parts[index].size() + 1;
    }
    std::wstring result;
    result.reserve(characters);
    for (size_t index = first; index < last; ++index)
    {
        if (parts[index].empty())
        {
            continue;
        }
        if (!result.empty())
        {
            result.push_back(L'\\');
        }
        result += parts[index];
    }
    return result;
}

}

bool HasComponentPrefix(std::wstring_view path, std::wstring_view prefix)
{
    return util::StartsWithInsensitive(path, prefix) && (path.size() == prefix.size() || path[prefix.size()] == L'\\');
}

bool LinkLeavesUserHives(std::wstring_view source_native, std::wstring_view target_native)
{
    const auto user = [](std::wstring_view path) {
        constexpr std::wstring_view kUsers = L"\\REGISTRY\\USER\\";
        constexpr std::wstring_view kClasses = L"_Classes";
        if (!util::StartsWithInsensitive(path, kUsers))
        {
            return std::wstring_view();
        }
        std::wstring_view hive = path.substr(kUsers.size());
        hive = hive.substr(0, hive.find(L'\\'));
        if (hive.size() > kClasses.size() && util::EndsWithInsensitive(hive, kClasses))
        {
            hive.remove_suffix(kClasses.size());
        }
        return hive;
    };
    const std::wstring_view source_user = user(source_native);
    return !source_user.empty() && !util::EqualsInsensitive(source_user, user(target_native));
}

std::wstring DisplayName(std::wstring_view name)
{
    // show embedded nulls without truncating text in the UI
    std::wstring text(name);
    std::replace(text.begin(), text.end(), wchar_t(0), kNullSymbol);
    return text;
}

std::wstring RawName(std::wstring_view text)
{
    // restore embedded nulls before native registry operations
    std::wstring name(text);
    std::replace(name.begin(), name.end(), kNullSymbol, wchar_t(0));
    return name;
}

HKEY RootFromName(std::wstring_view name)
{
    const RootNames* entry = FindRoot(name, Roots().size());
    return entry ? entry->root : nullptr;
}

std::wstring RootName(HKEY root)
{
    std::wstring virtual_name;
    if (RegistryStore::GetVirtualRootName(root, &virtual_name))
    {
        return virtual_name;
    }
    for (const RootNames& entry : Roots())
    {
        if (entry.root == root)
        {
            return entry.name;
        }
    }
    return {};
}
std::wstring Build(const RegistryNode& node)
{
    const std::wstring root = node.root_name.empty() ? RootName(node.root) : node.root_name;
    return DisplayName(Join(root, node.subkey));
}

bool InVolatileHive(std::wstring_view native_path)
{
    // HARDWARE is the only volatile hive Windows mounts, its keys carry no volatile flag
    return HasComponentPrefix(native_path, L"\\REGISTRY\\MACHINE\\HARDWARE");
}

ClassSource ClassesSource(std::wstring_view native_path)
{
    if (HasComponentPrefix(native_path, L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes"))
    {
        return ClassSource::kMachine;
    }
    return util::ContainsInsensitive(native_path, L"_Classes") ? ClassSource::kUser : ClassSource::kNone;
}

namespace
{

constexpr std::wstring_view kVirtualStore = L"HKEY_CURRENT_USER\\Software\\Classes\\VirtualStore\\MACHINE\\";

}

std::wstring VirtualStorePath(std::wstring_view native_path)
{
    constexpr std::wstring_view kMachine = L"\\REGISTRY\\MACHINE\\";
    if (!HasComponentPrefix(native_path, L"\\REGISTRY\\MACHINE\\SOFTWARE"))
    {
        return {};
    }
    // Windows never virtualizes these subtrees
    for (const std::wstring_view software : {L"\\REGISTRY\\MACHINE\\SOFTWARE\\", L"\\REGISTRY\\MACHINE\\SOFTWARE\\WOW6432Node\\"})
    {
        for (const std::wstring_view excluded : {L"Classes", L"Microsoft\\Windows", L"Microsoft\\Windows NT"})
        {
            if (HasComponentPrefix(native_path, std::wstring(software) + std::wstring(excluded)))
            {
                return {};
            }
        }
    }
    return std::wstring(kVirtualStore) + std::wstring(native_path.substr(kMachine.size()));
}

std::wstring GlobalKeyPath(std::wstring_view path)
{
    if (path.size() <= kVirtualStore.size() || !util::StartsWithInsensitive(path, kVirtualStore))
    {
        return {};
    }
    return L"HKEY_LOCAL_MACHINE\\" + std::wstring(path.substr(kVirtualStore.size()));
}

std::wstring ClassesSourcePath(std::wstring_view native_path, std::wstring_view current_user_sid)
{
    constexpr std::wstring_view kMachine = L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes";
    const std::wstring user = L"\\REGISTRY\\USER\\" + std::wstring(current_user_sid) + L"_Classes";
    if (HasComponentPrefix(native_path, kMachine))
    {
        return Join(L"HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes", TrimLeadingSeparators(native_path.substr(kMachine.size())));
    }
    if (!current_user_sid.empty() && HasComponentPrefix(native_path, user))
    {
        return Join(L"HKEY_CURRENT_USER\\Software\\Classes", TrimLeadingSeparators(native_path.substr(user.size())));
    }
    return {};
}

std::wstring BuildNative(const RegistryNode& node)
{
    if (RegistryStore::IsVirtualRoot(node.root))
    {
        return {};
    }
    if (util::EqualsInsensitive(node.root_name, L"REGISTRY"))
    {
        return Join(L"\\REGISTRY", node.subkey);
    }

    std::wstring root;
    if (node.root == HKEY_LOCAL_MACHINE)
    {
        root = L"\\REGISTRY\\MACHINE";
    }
    else if (node.root == HKEY_USERS)
    {
        root = L"\\REGISTRY\\USER";
    }
    else if (node.root == HKEY_CURRENT_USER)
    {
        // native HKCU lives under the current user SID
        const std::wstring sid = util::GetCurrentUserSidString();
        if (sid.empty())
        {
            return {};
        }
        root = L"\\REGISTRY\\USER\\" + sid;
    }
    else if (node.root == HKEY_CURRENT_CONFIG)
    {
        // HKCC points to the active hardware profile
        root = L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Hardware Profiles\\Current";
    }
    else
    {
        return {};
    }
    return Join(root, node.subkey);
}

std::vector<std::wstring> Split(std::wstring_view path)
{
    std::vector<std::wstring> parts;
    size_t start = 0;
    while (start < path.size())
    {
        while (start < path.size() && path[start] == L'\\')
        {
            ++start;
        }
        if (start == path.size())
        {
            break;
        }
        size_t end = path.find(L'\\', start);
        if (end == std::wstring_view::npos)
        {
            end = path.size();
        }
        parts.emplace_back(path.substr(start, end - start));
        start = end + 1;
    }
    return parts;
}

std::wstring Join(const std::vector<std::wstring>& parts, size_t first_part)
{
    return JoinRange(parts, std::min(first_part, parts.size()), parts.size());
}

std::wstring JoinPrefix(const std::vector<std::wstring>& parts, size_t part_count)
{
    return JoinRange(parts, 0, std::min(part_count, parts.size()));
}

std::wstring JoinSubkey(std::wstring_view parent, std::wstring_view name)
{
    return parent.empty() ? std::wstring(name) : Join(parent, name);
}

RegistryNode ChildNode(const RegistryNode& parent, std::wstring_view name)
{
    RegistryNode child;
    child.root = parent.root;
    child.root_name = parent.root_name;
    child.view = parent.view;
    child.subkey = JoinSubkey(parent.subkey, name);
    child.simulated = parent.simulated;
    return child;
}

std::wstring Parent(std::wstring_view path)
{
    while (!path.empty() && path.back() == L'\\')
    {
        path.remove_suffix(1);
    }
    const size_t split = path.rfind(L'\\');
    return split == std::wstring_view::npos ? std::wstring{} : std::wstring(path.substr(0, split));
}

std::wstring Leaf(std::wstring_view path)
{
    while (!path.empty() && path.back() == L'\\')
    {
        path.remove_suffix(1);
    }
    const size_t split = path.rfind(L'\\');
    return std::wstring(path.substr(split == std::wstring_view::npos ? 0 : split + 1));
}

std::wstring Clean(std::wstring_view input)
{
    std::wstring path = util::TrimWhitespace(input);
    if (path.size() >= 2 &&
        ((path.front() == L'[' && path.back() == L']') || (path.front() == L'"' && path.back() == L'"') ||
         (path.front() == L'\'' && path.back() == L'\'')))
    {
        path = util::TrimWhitespace(std::wstring_view(path).substr(1, path.size() - 2));
    }
    if (!path.empty() && path.front() == L'-')
    {
        path = util::TrimWhitespace(std::wstring_view(path).substr(1));
    }
    if (path.find(L'\\') == std::wstring::npos)
    {
        std::replace(path.begin(), path.end(), L'/', L'\\');
    }

    std::wstring collapsed;
    collapsed.reserve(path.size());
    for (wchar_t character : path)
    {
        if (character != L'\\' || collapsed.empty() || collapsed.back() != L'\\')
        {
            collapsed.push_back(character);
        }
    }
    path = std::move(collapsed);

    if (util::StartsWithInsensitive(path, L"reg:"))
    {
        path = util::TrimWhitespace(std::wstring_view(path).substr(4));
    }
    if (util::StartsWithInsensitive(path, L"Registry::"))
    {
        path.erase(0, 10);
    }
    while (!path.empty() && path.front() == L'\\')
    {
        path.erase(path.begin());
    }
    if (util::StartsWithInsensitive(path, L"My Computer\\"))
    {
        path.erase(0, 12);
    }
    else if (util::StartsWithInsensitive(path, L"Computer\\"))
    {
        path.erase(0, 9);
    }
    return path;
}

std::wstring Normalize(std::wstring_view input, std::wstring_view current_user_sid)
{
    std::wstring path = Clean(input);

    if (util::StartsWithInsensitive(path, L"REGISTRY\\"))
    {
        const std::wstring native = path.substr(9);
        const auto native_rest = [&](std::wstring_view prefix) { return TrimLeadingSeparators(std::wstring_view(native).substr(prefix.size())); };
        constexpr std::wstring_view current_config = L"MACHINE\\SYSTEM\\CurrentControlSet\\Hardware Profiles\\Current";
        if (HasComponentPrefix(native, current_config))
        {
            return Join(L"HKEY_CURRENT_CONFIG", native_rest(current_config));
        }
        if (HasComponentPrefix(native, L"MACHINE"))
        {
            return Join(L"HKEY_LOCAL_MACHINE", native_rest(L"MACHINE"));
        }
        if (HasComponentPrefix(native, L"USER"))
        {
            const std::wstring_view rest = native_rest(L"USER");
            if (!current_user_sid.empty() && HasComponentPrefix(rest, current_user_sid))
            {
                return Join(L"HKEY_CURRENT_USER", TrimLeadingSeparators(rest.substr(current_user_sid.size())));
            }
            return Join(L"HKEY_USERS", rest);
        }
        const size_t native_split = native.find(L'\\');
        const std::wstring_view native_root = native_split == std::wstring::npos
                                                  ? std::wstring_view(native)
                                                  : std::wstring_view(native).substr(0, native_split);
        const std::wstring canonical = CanonicalRoot(native_root);
        if (!canonical.empty() && !util::EqualsInsensitive(canonical, L"REGISTRY"))
        {
            const std::wstring_view rest = native_split == std::wstring::npos
                                               ? std::wstring_view{}
                                               : std::wstring_view(native).substr(native_split + 1);
            return Join(canonical, rest);
        }
    }

    const size_t split = path.find_first_of(L":\\");
    const std::wstring_view root(path.data(), split == std::wstring::npos ? path.size() : split);
    const std::wstring_view rest = split == std::wstring::npos ? std::wstring_view{} : TrimLeadingSeparators(std::wstring_view(path).substr(split + (path[split] == L':' ? 1 : 0)));
    const std::wstring canonical = CanonicalRoot(root);
    return canonical.empty() ? path : Join(canonical, rest);
}

std::wstring Format(std::wstring_view path, Style style, std::wstring_view tree_root)
{
    const size_t split = path.find(L'\\');
    const std::wstring_view root = split == std::wstring_view::npos ? path : path.substr(0, split);
    const std::wstring_view rest = split == std::wstring_view::npos ? std::wstring_view{} : path.substr(split + 1);
    switch (style)
    {
    case Style::kAbbreviated:
        return Join(AbbreviatedRoot(root), rest);
    case Style::kRegEditAddress:
        return Join(tree_root.empty() ? L"Computer" : tree_root, path);
    case Style::kRegFileHeader:
        return L"[" + std::wstring(path) + L"]";
    case Style::kPowerShellDrive:
        {
            std::wstring result = AbbreviatedRoot(root) + L":";
            if (!rest.empty())
            {
                result += L"\\" + std::wstring(rest);
            }
            return result;
        }
    case Style::kPowerShellProvider:
        return L"Registry::" + std::wstring(path);
    case Style::kEscaped:
        {
            std::wstring result;
            result.reserve(path.size() * 2);
            for (wchar_t character : path)
            {
                if (character == L'\\')
                {
                    result.push_back(L'\\');
                }
                result.push_back(character);
            }
            return result;
        }
    case Style::kFull:
        return std::wstring(path);
    }
    return std::wstring(path);
}

bool ParseRoot(std::wstring_view input, RegistryNode* node)
{
    if (!node)
    {
        return false;
    }
    const std::wstring normalized = Normalize(input);
    const size_t split = normalized.find(L'\\');
    const std::wstring root = normalized.substr(0, split);
    const std::wstring rest = split == std::wstring::npos ? L"" : normalized.substr(split + 1);
    node->subkey = RawName(rest);
    node->root_name = root;
    const RootNames* entry = FindRoot(root, kBrowsableRoots);
    node->root = entry ? entry->root : nullptr;
    // native REGISTRY paths intentionally have no Win32 root handle
    return entry || util::EqualsInsensitive(root, L"REGISTRY");
}

bool ResolveJumpTarget(std::wstring_view target, const std::function<std::wstring(const std::wstring&)>& normalize, const std::function<bool(const std::wstring&, RegistryNode*)>& key_exists, std::wstring* key_path, std::wstring* value_name, bool* value_missing)
{
    const auto unwrap = [](std::wstring text, std::wstring_view pairs) {
        text = util::TrimWhitespace(text);
        for (size_t pair = 0; pair + 1 < pairs.size(); pair += 2)
        {
            if (text.size() >= 2 && text.front() == pairs[pair] && text.back() == pairs[pair + 1])
            {
                return util::TrimWhitespace(std::wstring_view(text).substr(1, text.size() - 2));
            }
        }
        return text;
    };
    RegistryNode node;
    RegistryValue value;
    value_name->clear();
    *value_missing = false;
    const std::wstring text = unwrap(std::wstring(target), L"\"\"''[]");
    *key_path = normalize(text);
    if (key_path->empty() || key_exists(*key_path, &node))
    {
        return !key_path->empty();
    }

    std::wstring missing_key;
    std::wstring existing_key;
    std::wstring existing_name;
    for (size_t split = 1; split < text.size(); ++split)
    {
        const bool colon =
            text[split] == L':' && (iswspace(text[split - 1]) || split + 1 == text.size() || iswspace(text[split + 1]));
        if (text[split] != L'!' && !colon)
        {
            continue;
        }
        const std::wstring key = normalize(text.substr(0, split));
        const std::wstring name =
            registry_path::RawName(colon ? unwrap(text.substr(split + 1), L"\"\"") : text.substr(split + 1));
        if (!key_exists(key, &node))
        {
            missing_key = missing_key.empty() ? key : missing_key;
            continue;
        }
        if (name.empty() || RegistryStore::QueryValue(node, name, &value))
        {
            *key_path = key;
            *value_name = name;
            return true;
        }
        existing_key = key;
        existing_name = name;
    }
    if (!existing_key.empty())
    {
        *key_path = existing_key;
        *value_name = existing_name;
        *value_missing = true;
        return true;
    }
    const std::wstring normalized = *key_path;
    for (size_t slash = normalized.rfind(L'\\'); slash != std::wstring::npos && slash > 0;
         slash = normalized.rfind(L'\\', slash - 1))
    {
        if (key_exists(normalized.substr(0, slash), &node) &&
            RegistryStore::QueryValue(node, normalized.substr(slash + 1), &value))
        {
            *key_path = normalized.substr(0, slash);
            *value_name = normalized.substr(slash + 1);
            return true;
        }
    }
    if (!missing_key.empty())
    {
        *key_path = missing_key;
    }
    return false;
}

}
