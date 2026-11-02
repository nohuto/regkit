// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/compare.h"
#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"

#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/translation.h"

#include <algorithm>
#include <unordered_set>
#include <utility>

namespace regkit::search::compare
{

namespace
{

bool Cancelled(const std::atomic_bool* cancel)
{
    return cancel && cancel->load();
}

bool IsWithin(const std::wstring& path, const std::wstring& base, bool recursive)
{
    return util::EqualsInsensitive(path, base) ||
           (recursive && path.size() > base.size() && path[base.size()] == L'\\' &&
            util::StartsWithInsensitive(path, base));
}

std::wstring Combine(const std::wstring& base, const std::wstring& relative)
{
    return registry_path::JoinSubkey(base, registry_path::DisplayName(relative));
}

std::wstring DataText(const Value& value)
{
    if (value.data.empty())
    {
        return L"";
    }
    return value_format::DisplayData(value.type, value.data.data(), static_cast<DWORD>(value.data.size()), false);
}

std::wstring EntryText(const Value* value)
{
    if (!value)
    {
        return util::Tr(L"(Missing)");
    }
    const std::wstring type = value_format::TypeName(value->type);
    const std::wstring data = DataText(*value);
    return data.empty() ? type : type + L": " + data;
}

template <typename Map>
void AppendKeys(const Map& source, std::unordered_set<std::wstring>* seen, std::vector<std::wstring>* target)
{
    for (const auto& pair : source)
    {
        if (seen->insert(pair.first).second)
        {
            target->push_back(pair.first);
        }
    }
}

}

bool CaptureRegistry(const std::wstring& base_path, const RegistryNode& base_node, bool recursive, Snapshot* snapshot, std::wstring* error, std::atomic_bool* cancel)
{
    if (!snapshot)
    {
        return false;
    }
    snapshot->label = base_path;
    snapshot->base_path = base_path;
    snapshot->keys.clear();
    snapshot->unreadable.clear();

    std::vector<std::pair<RegistryNode, std::wstring>> stack;
    stack.emplace_back(base_node, L"");
    while (!stack.empty())
    {
        if (Cancelled(cancel))
        {
            return false;
        }
        RegistryNode node = std::move(stack.back().first);
        std::wstring relative = std::move(stack.back().second);
        stack.pop_back();

        Key key;
        key.relative_path = relative;
        RegistryStore::KeyEnumResult enumeration;
        bool reserved = false;
        std::vector<std::wstring> children;
        const bool enumerated = RegistryStore::EnumKeyStreaming(
            node,
            true,
            true,
            recursive,
            &enumeration,
            [&](const ValueInfo& value, const BYTE* data, DWORD size) {
                if (Cancelled(cancel))
                {
                    return false;
                }
                if (!reserved)
                {
                    if (enumeration.info_valid)
                    {
                        key.values.reserve(enumeration.info.value_count);
                    }
                    reserved = true;
                }
                Value captured;
                captured.name = value.name;
                captured.type = value.type;
                if (data && size > 0)
                {
                    captured.data.assign(data, data + size);
                }
                key.values[util::ToLower(captured.name)] = std::move(captured);
                return true;
            },
            recursive ? RegistryStore::SubkeyStreamCallback([&](const std::wstring& name) {
                children.push_back(name);
                return true;
            })
                      : RegistryStore::SubkeyStreamCallback()
        );
        if (Cancelled(cancel))
        {
            return false;
        }
        if (!enumerated && !relative.empty())
        {
            snapshot->unreadable.push_back(std::move(relative));
            continue;
        }
        if (!enumerated)
        {
            if (error)
            {
                *error = util::TrDetail(L"Couldn't read the registry key.", base_path);
            }
            return false;
        }
        snapshot->keys[util::ToLower(relative)] = std::move(key);

        for (const auto& name : children)
        {
            stack.emplace_back(registry_path::ChildNode(node, name), registry_path::JoinSubkey(relative, name));
        }
    }
    return true;
}

bool LoadRegFile(const std::wstring& file_path, const regfile::Document& document, const std::wstring& base_path, bool recursive, const NormalizePath& normalize, Snapshot* snapshot, std::wstring* error, std::atomic_bool* cancel)
{
    if (!snapshot || !normalize)
    {
        return false;
    }
    if (document.keys.empty())
    {
        if (error)
        {
            *error = util::Tr(L"No registry keys were found in the .reg file.");
        }
        return false;
    }

    snapshot->base_path = base_path;
    snapshot->label = util::FileName(file_path);
    if (!base_path.empty())
    {
        snapshot->label += L": " + base_path;
    }
    snapshot->keys.clear();

    bool matched = false;
    for (const auto& original_path : document.key_order)
    {
        if (Cancelled(cancel))
        {
            return false;
        }
        const std::wstring normalized = normalize(original_path);
        if (normalized.empty())
        {
            continue;
        }
        if (!IsWithin(normalized, base_path, recursive))
        {
            if (!recursive && IsWithin(normalized, base_path, true))
            {
                snapshot->keys.try_emplace(L"", Key{L"", {}});
                matched = true;
            }
            continue;
        }
        matched = true;
        std::wstring relative;
        if (normalized.size() > base_path.size())
        {
            relative = normalized.substr(base_path.size() + 1);
        }
        auto source = document.keys.find(util::ToLower(original_path));
        if (source == document.keys.end())
        {
            source = document.keys.find(util::ToLower(normalized));
        }

        Key key;
        key.relative_path = relative;
        if (source != document.keys.end())
        {
            key.values = source->second.values;
        }
        snapshot->keys[util::ToLower(relative)] = std::move(key);
        // importing the file creates the keys above each listed one
        for (std::wstring parent = relative; !parent.empty();)
        {
            parent = registry_path::Parent(parent);
            if (!snapshot->keys.try_emplace(util::ToLower(parent), Key{parent, {}}).second)
            {
                break;
            }
        }
    }

    if (!matched)
    {
        if (error)
        {
            *error = util::Tr(L"No matching keys were found for the selected path.");
        }
        return false;
    }
    return true;
}

void SortRows(std::vector<Row>* rows, int column, bool ascending, const util::SortKeys& earlier)
{
    if (!rows || rows->size() < 2)
    {
        return;
    }
    auto field = [](const Row& row, int key) -> const std::wstring& {
        switch (key)
        {
        case 1:
            return row.value_name;
        case 2:
            return row.first_text;
        case 3:
            return row.second_text;
        default:
            return row.key_path;
        }
    };
    std::stable_sort(rows->begin(), rows->end(), [&](const Row& left, const Row& right) {
        return util::CompareSorted(column, ascending, earlier, [&](int key) {
                   return key == 4 ? static_cast<int>(left.matches) - static_cast<int>(right.matches) : util::CompareListText(field(left, key), field(right, key));
               }) < 0;
    });
}

std::vector<Row> BuildRows(const Snapshot& first, const Snapshot& second, RowFilter filter, std::atomic_bool* cancel)
{
    const bool include_differences = filter != RowFilter::kMatches;
    const bool include_matches = filter != RowFilter::kDifferences;
    // build key union with display paths once so sorting needs no map lookups
    std::vector<std::pair<const std::wstring*, const std::wstring*>> keys;
    keys.reserve(first.keys.size() + second.keys.size());
    for (const auto& [name, key] : first.keys)
    {
        keys.emplace_back(&name, &key.relative_path);
    }
    for (const auto& [name, key] : second.keys)
    {
        if (!first.keys.contains(name))
        {
            keys.emplace_back(&name, &key.relative_path);
        }
    }
    std::sort(keys.begin(), keys.end(), [](const auto& left, const auto& right) {
        return util::CompareInsensitive(*left.second, *right.second) < 0;
    });

    auto unread = [&](const std::wstring& key_name) {
        for (const Snapshot* side : {&first, &second})
        {
            for (const std::wstring& path : side->unreadable)
            {
                if (IsWithin(key_name, path, true))
                {
                    return true;
                }
            }
        }
        return false;
    };
    std::vector<Row> results;
    for (const auto& [name, display] : keys)
    {
        const std::wstring& key_name = *name;
        if (Cancelled(cancel))
        {
            break;
        }
        if (unread(key_name))
        {
            continue;
        }
        const auto first_it = first.keys.find(key_name);
        const auto second_it = second.keys.find(key_name);
        const Key* first_key = first_it == first.keys.end() ? nullptr : &first_it->second;
        const Key* second_key = second_it == second.keys.end() ? nullptr : &second_it->second;
        const std::wstring& relative = *display;
        const std::wstring first_path = Combine(first.base_path, relative);
        const std::wstring second_path = Combine(second.base_path, relative);

        if (!first_key || !second_key)
        {
            if (!include_differences)
            {
                continue;
            }
            Row result;
            result.is_key = true;
            result.key_path = first_key ? first_path : second_path;
            result.first_key_path = first_key ? first_path : std::wstring();
            result.second_key_path = second_key ? second_path : std::wstring();
            result.first_text = first_key ? util::Tr(L"Present") : util::Tr(L"(Missing)");
            result.second_text = second_key ? util::Tr(L"Present") : util::Tr(L"(Missing)");
            results.push_back(std::move(result));
            continue;
        }

        if (include_matches)
        {
            Row result;
            result.is_key = true;
            result.matches = true;
            result.key_path = first_path;
            result.first_key_path = first_path;
            result.second_key_path = second_path;
            result.first_text = util::Tr(L"Present");
            result.second_text = util::Tr(L"Present");
            results.push_back(std::move(result));
        }

        std::vector<std::wstring> values;
        values.reserve(first_key->values.size() + second_key->values.size());
        std::unordered_set<std::wstring> seen_values;
        seen_values.reserve(values.capacity());
        AppendKeys(first_key->values, &seen_values, &values);
        AppendKeys(second_key->values, &seen_values, &values);
        std::sort(values.begin(), values.end(), [](const std::wstring& left, const std::wstring& right) {
            return util::CompareInsensitive(left, right) < 0;
        });

        for (const auto& value_name : values)
        {
            if (Cancelled(cancel))
            {
                return results;
            }
            const auto first_value = first_key->values.find(value_name);
            const auto second_value = second_key->values.find(value_name);
            const Value* left = first_value == first_key->values.end() ? nullptr : &first_value->second;
            const Value* right = second_value == second_key->values.end() ? nullptr : &second_value->second;
            const bool matches = left && right && left->type == right->type && left->data == right->data;
            if ((matches && !include_matches) || (!matches && !include_differences))
            {
                continue;
            }

            Row result;
            result.matches = matches;
            result.key_path = first_path;
            result.first_key_path = first_path;
            result.second_key_path = second_path;
            result.value_name = left ? left->name : right->name;
            result.first_text = EntryText(left);
            result.second_text = EntryText(right);
            results.push_back(std::move(result));
        }
    }
    return results;
}

std::wstring SerializeRows(const std::vector<Row>& rows)
{
    std::wstring content;
    for (const Row& row : rows)
    {
        record_fields::AppendRecord(&content, {row.key_path, row.first_key_path, row.second_key_path, row.value_name, row.first_text, row.second_text, row.is_key ? L"1" : L"0", row.matches ? L"1" : L"0"});
    }
    return content;
}

bool ParseRows(const std::wstring& content, std::vector<Row>* rows, std::wstring* error)
{
    if (!rows)
    {
        return false;
    }
    std::vector<Row> parsed;
    size_t line_number = 0;
    for (const std::wstring_view line : win32ui::LineViews(content))
    {
        ++line_number;
        if (line.empty())
        {
            continue;
        }
        auto fields = record_fields::DecodeRecord(line);
        if (fields.size() != 8)
        {
            if (error)
            {
                *error = util::BadLineReason(line_number);
            }
            return false;
        }
        Row row;
        row.key_path = std::move(fields[0]);
        row.first_key_path = std::move(fields[1]);
        row.second_key_path = std::move(fields[2]);
        row.value_name = std::move(fields[3]);
        row.first_text = std::move(fields[4]);
        row.second_text = std::move(fields[5]);
        row.is_key = fields[6] == L"1";
        row.matches = fields[7] == L"1";
        parsed.push_back(std::move(row));
    }
    *rows = std::move(parsed);
    return true;
}

bool SaveRows(const std::wstring& path, const std::vector<Row>& rows)
{
    return !path.empty() && win32ui::WriteTextFileAtomic(path.c_str(), SerializeRows(rows)) == ERROR_SUCCESS;
}

bool LoadRows(const std::wstring& path, std::vector<Row>* rows, std::wstring* error)
{
    std::wstring content;
    return rows && !path.empty() && util::ReadStateFile(path, &content, error, 64ull * 1024ull * 1024ull) && !content.empty() && ParseRows(content, rows, error);
}

}
