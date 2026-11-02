// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <cstddef>
#include <functional>
#include <string>
#include <string_view>
#include <vector>

namespace regkit
{

struct RegistryNode;
enum class ClassSource : unsigned char;

namespace registry_path
{

enum class Style
{
    kFull,
    kAbbreviated,
    kRegEditAddress,
    kRegFileHeader,
    kPowerShellDrive,
    kPowerShellProvider,
    kEscaped,
};

inline constexpr wchar_t kNullSymbol = 0x2400;

std::wstring DisplayName(std::wstring_view name);
std::wstring RawName(std::wstring_view text);
HKEY RootFromName(std::wstring_view name);
std::wstring RootName(HKEY root);
std::wstring Build(const RegistryNode& node);
std::wstring BuildNative(const RegistryNode& node);
bool InVolatileHive(std::wstring_view native_path);
ClassSource ClassesSource(std::wstring_view native_path);
std::wstring ClassesSourcePath(std::wstring_view native_path, std::wstring_view current_user_sid);
// UAC virtual store holds per user HKLM\SOFTWARE copies
std::wstring VirtualStorePath(std::wstring_view native_path);
std::wstring GlobalKeyPath(std::wstring_view path);

std::wstring Clean(std::wstring_view path);
bool HasComponentPrefix(std::wstring_view path, std::wstring_view prefix);
// Windows follows user hive links only into that user hive or its _Classes hive
bool LinkLeavesUserHives(std::wstring_view source_native, std::wstring_view target_native);
std::wstring Normalize(std::wstring_view path, std::wstring_view current_user_sid = {});
std::wstring Format(std::wstring_view normalized_path, Style style, std::wstring_view tree_root = L"Computer");
bool ParseRoot(std::wstring_view path, RegistryNode* node);
bool ResolveJumpTarget(std::wstring_view target, const std::function<std::wstring(const std::wstring&)>& normalize, const std::function<bool(const std::wstring&, RegistryNode*)>& key_exists, std::wstring* key_path, std::wstring* value_name, bool* value_missing);
std::vector<std::wstring> Split(std::wstring_view path);
std::wstring Join(const std::vector<std::wstring>& parts, size_t first_part = 0);
std::wstring JoinPrefix(const std::vector<std::wstring>& parts, size_t part_count);
std::wstring JoinSubkey(std::wstring_view parent, std::wstring_view name);
RegistryNode ChildNode(const RegistryNode& parent, std::wstring_view name);
std::wstring Parent(std::wstring_view path);
std::wstring Leaf(std::wstring_view path);

}
}
