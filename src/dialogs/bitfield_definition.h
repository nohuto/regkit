// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <array>
#include <cstdint>
#include <string>
#include <vector>

namespace regkit::editors::bitfield
{

constexpr size_t kMaxNameLength = 256;
constexpr size_t kMaxStates = 256;
constexpr size_t kMaxMeaningLength = 4096;
constexpr size_t kMaxCommentLength = 16384;
constexpr size_t kMaxPathLength = 512;
constexpr size_t kMaxPaths = 16;
constexpr uint64_t kMaxFileBytes = 8ull * 1024ull * 1024ull;
constexpr unsigned kMaxByteOffset = 65536;

struct State
{
    uint64_t value = 0;
    std::wstring name;
    std::wstring meaning;
};

struct Field
{
    std::wstring name;
    std::vector<unsigned> bits;
    std::wstring meaning;
    std::vector<State> states;
    uint64_t mask = 0;

    uint64_t Extract(uint64_t value) const;
    uint64_t Apply(uint64_t value, uint64_t field_value) const;
    const State* StateFor(uint64_t field_value) const;
};

struct Definition
{
    std::wstring name;
    std::wstring value_name;
    std::vector<std::wstring> key_paths;
    unsigned bit_width = 32;
    unsigned byte_offset = 0;
    std::wstring comment;
    std::vector<Field> fields;
    std::array<signed char, 64> owner = {};

    const Field* FieldForBit(unsigned bit) const;
    int FieldIndexForBit(unsigned bit) const;
    bool MatchesPath(const std::wstring& key_path) const;
    unsigned byte_count() const
    {
        return bit_width / 8;
    }
};

struct DefinitionFile
{
    std::wstring name;
    std::wstring comment;
    std::wstring path;
    std::vector<Definition> definitions;
};

std::wstring DisplayName(const Definition& definition);
bool ValidWidth(unsigned bit_width);
uint64_t WidthMask(unsigned bit_width);
void BuildLookup(Definition* definition);

bool Validate(Definition* definition, std::wstring* error);
bool Validate(DefinitionFile* file, std::wstring* error);
bool Parse(const std::vector<BYTE>& utf8, DefinitionFile* file, std::wstring* error);
bool Load(const std::wstring& path, DefinitionFile* file, std::wstring* error);
std::wstring Serialize(const DefinitionFile& file);
bool Save(const std::wstring& path, const DefinitionFile& file, std::wstring* error);

const std::vector<DefinitionFile>& BundledFiles();
std::vector<Definition> Matching(const std::wstring& key_path, const std::wstring& value_name);
std::wstring SuggestedFileName(const std::wstring& value_name);
const wchar_t* FileFilter();

}
