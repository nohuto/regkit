// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "regfile/reg_file.h"
#include "registry/registry_store.h"
#include "search/search.h"

#include <atomic>
#include <functional>
#include <string>
#include <unordered_map>
#include <vector>

namespace regkit::search::compare
{

struct Row
{
    std::wstring key_path;
    std::wstring first_key_path;
    std::wstring second_key_path;
    std::wstring value_name;
    std::wstring first_text;
    std::wstring second_text;
    bool is_key = false;
    bool matches = false;
};

using Value = RegistryValue;

struct Key
{
    std::wstring relative_path;
    std::unordered_map<std::wstring, Value> values;
};

struct Snapshot
{
    std::wstring label;
    std::wstring base_path;
    std::unordered_map<std::wstring, Key> keys;
    std::vector<std::wstring> unreadable;
};

enum class RowFilter
{
    kDifferences,
    kMatches,
    kAll,
};

using NormalizePath = std::function<std::wstring(const std::wstring& path)>;

bool CaptureRegistry(const std::wstring& base_path, const RegistryNode& base_node, bool recursive, Snapshot* snapshot, std::wstring* error = nullptr, std::atomic_bool* cancel = nullptr);

bool LoadRegFile(const std::wstring& file_path, const regfile::Document& document, const std::wstring& base_path, bool recursive, const NormalizePath& normalize, Snapshot* snapshot, std::wstring* error, std::atomic_bool* cancel = nullptr);

void SortRows(std::vector<Row>* rows, int column, bool ascending, const util::SortKeys& earlier);

std::vector<Row> BuildRows(const Snapshot& first, const Snapshot& second, RowFilter filter = RowFilter::kDifferences, std::atomic_bool* cancel = nullptr);

std::wstring SerializeRows(const std::vector<Row>& rows);
bool ParseRows(const std::wstring& content, std::vector<Row>* rows, std::wstring* error = nullptr);
bool SaveRows(const std::wstring& path, const std::vector<Row>& rows);
bool LoadRows(const std::wstring& path, std::vector<Row>* rows, std::wstring* error = nullptr);

}
