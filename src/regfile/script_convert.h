// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "regfile/reg_file.h"

#include <string>
#include <string_view>
#include <vector>

namespace regkit::regfile
{

enum class Format
{
    kReg,
    kBatch,
    kPowerShell,
};

// indexed by Format
inline constexpr const wchar_t* kFormatFilters[] = {
    L"Registry Files (*.reg)\0*.reg\0",
    L"Batch Files (*.bat;*.cmd)\0*.bat;*.cmd\0",
    L"PowerShell Scripts (*.ps1)\0*.ps1\0",
};
inline constexpr wchar_t kConvertFilter[] =
    L"Registry Scripts (*.reg;*.bat;*.cmd;*.ps1)\0*.reg;*.bat;*.cmd;*.ps1\0Registry Files (*.reg)\0*.reg\0Batch Files (*.bat;*.cmd)\0*.bat;*.cmd\0PowerShell Scripts (*.ps1)\0*.ps1\0All Files (*.*)\0*.*\0";

bool FormatFromPath(std::wstring_view path, Format* format);
const wchar_t* FormatExtension(Format format);
bool ReadOperations(const std::wstring& path, std::vector<Operation>* operations, std::wstring* error);
bool ReadOperations(const std::wstring& path, Format format, std::vector<Operation>* operations, std::wstring* error);
bool ParseText(std::wstring_view content, Format format, std::vector<Operation>* operations, std::wstring* error);
bool ReadRegistry(const std::wstring& key_path, bool recursive, std::vector<Operation>* operations, std::wstring* error, std::vector<std::wstring>* skipped = nullptr);
bool SelectKey(const std::wstring& key_path, bool recursive, std::vector<Operation>* operations, std::wstring* error);
std::wstring RenderOperations(Format format, const std::vector<Operation>& operations, bool admin_check, std::vector<std::wstring>* skipped);
bool SaveRendered(const std::wstring& path, Format format, const std::wstring& text);

bool NormalizeKeyPath(std::wstring_view text, std::wstring* path);
bool IsUnderKey(std::wstring_view path, std::wstring_view parent);
bool NeedsAdmin(const std::vector<Operation>& operations);
std::vector<std::wstring> MultiStringItems(const std::vector<BYTE>& data);
std::wstring Describe(const Operation& operation, std::wstring_view reason);
bool ParseBatch(std::wstring_view content, std::vector<Operation>* output, std::wstring* error);
std::wstring RenderBatch(const std::vector<Operation>& operations, bool admin_check, std::vector<std::wstring>* skipped);
bool ParsePowerShell(std::wstring_view content, std::vector<Operation>* output, std::wstring* error);
std::wstring RenderPowerShell(const std::vector<Operation>& operations, bool admin_check, std::vector<std::wstring>* skipped);

}
