// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_value.h"
#include "registry/virtual_registry.h"

#include <win32ui/windows_config.h>

#include <windows.h>

#include <atomic>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace regkit::regfile
{

using Value = RegistryValue;

struct Key
{
    std::wstring path;
    std::unordered_map<std::wstring, Value> values;
    std::vector<std::wstring> removed_values;
    bool removed = false;
    bool recreated = false;
};

struct Operation
{
    enum class Kind
    {
        kKey,
        kRemoveKey,
        kValue,
        kRemoveValue,
    };
    Kind kind = Kind::kKey;
    std::wstring path;
    Value value;
};

struct Document
{
    std::vector<std::wstring> key_order;
    std::unordered_map<std::wstring, Key> keys;
};

class Writer
{
  public:
    Writer();

    void AppendKey(std::wstring_view path, std::vector<const Value*> values, bool sorted = true);
    void AppendRemovedKey(std::wstring_view path);
    void AppendValue(const Value& value);
    void AppendRemovedValue(std::wstring_view name);
    std::wstring Finish() &&;

  private:
    void AppendKeyHeader(std::wstring_view prefix, std::wstring_view path);

    std::wstring output_;
};

void SkipNullNames(const std::wstring& display_path, std::vector<Value>* values, std::vector<std::wstring>* subkeys, std::vector<std::wstring>* skipped);
LONG AppendRegistryTree(Writer* writer, HKEY root, const std::wstring& subkey, const std::wstring& display_path, REGSAM view, bool recurse, std::vector<std::wstring>* skipped = nullptr, std::vector<std::wstring>* lost_volatility = nullptr);
LONG ReadRegistryOperations(HKEY root, const std::wstring& subkey, const std::wstring& display_path, REGSAM view, bool recurse, std::vector<Operation>* output, std::vector<std::wstring>* skipped = nullptr, std::vector<std::wstring>* lost_volatility = nullptr);
bool ParseOperations(std::wstring_view content, std::vector<Operation>* output, const std::atomic_bool* cancel = nullptr, bool* cancelled = nullptr, std::wstring* error = nullptr);
bool Parse(std::wstring_view content, Document* output, const std::atomic_bool* cancel = nullptr, bool* cancelled = nullptr, std::wstring* error = nullptr);
bool ReadText(const std::wstring& path, std::wstring* content, std::wstring* error);
bool Load(const std::wstring& path, Document* output, std::wstring* error, const std::atomic_bool* cancel = nullptr, bool* cancelled = nullptr);
std::wstring RenderReg(const std::vector<Operation>& operations);

struct VirtualRoot
{
    std::wstring name;
    std::shared_ptr<VirtualRegistryData> data;
};

bool LoadVirtualRoots(const std::wstring& path, std::vector<VirtualRoot>* roots, std::wstring* error, const std::atomic_bool* cancel, bool* cancelled);

}
