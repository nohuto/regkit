// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

namespace regkit {

struct RegistryNode;

namespace registry_path {

enum class Style {
  kFull,
  kAbbreviated,
  kRegeditAddress,
  kRegFileHeader,
  kPowerShellDrive,
  kPowerShellProvider,
  kEscaped,
};

std::wstring RootName(HKEY root);
std::wstring Build(const RegistryNode& node);
std::wstring BuildNative(const RegistryNode& node);

std::wstring Clean(std::wstring_view path);
std::wstring Normalize(std::wstring_view path,
                       std::wstring_view current_user_sid = {});
std::wstring Format(std::wstring_view normalized_path, Style style,
                    std::wstring_view tree_root = L"Computer");
bool ParseRoot(std::wstring_view path, RegistryNode* node);
std::vector<std::wstring> Split(std::wstring_view path);
std::wstring Join(const std::vector<std::wstring>& parts,
                  size_t first_part = 0);
std::wstring JoinPrefix(const std::vector<std::wstring>& parts,
                        size_t part_count);
std::wstring Parent(std::wstring_view path);
std::wstring Leaf(std::wstring_view path);
bool Equals(std::wstring_view left, std::wstring_view right);
bool StartsWith(std::wstring_view text, std::wstring_view prefix);

} // namespace registry_path
} // namespace regkit
