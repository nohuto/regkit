// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "defaults/default_loader.h"

#include <cwctype>
#include "win32/text_transform.h"

#include "regfile/reg_file.h"
#include "registry/value_format.h"
#include "win32/translation.h"

#include <utility>

namespace regkit::defaults
{

bool Load(const std::wstring& path, const NormalizePath& normalize, Data* data, std::vector<Entry>* entries, std::wstring* error, const std::atomic_bool* cancel)
{
    if ((!data && !entries) || !normalize)
    {
        return false;
    }
    regfile::Document document;
    bool cancelled = false;
    if (!regfile::Load(path, &document, error, cancel, &cancelled))
    {
        return false;
    }

    Data loaded;
    bool saw_key = false;
    std::vector<Entry> parsed_entries;
    if (entries)
    {
        parsed_entries.reserve(document.key_order.size());
    }
    for (const auto& source_path : document.key_order)
    {
        if (cancel && cancel->load())
        {
            return false;
        }
        std::wstring key_path = normalize(source_path);
        if (key_path.empty())
        {
            key_path = source_path;
        }
        if (key_path.empty())
        {
            continue;
        }
        const auto source = document.keys.find(util::ToLower(source_path));
        if (source == document.keys.end())
        {
            continue;
        }

        saw_key = true;
        Key* target = nullptr;
        if (data)
        {
            target = &loaded.values_by_key[util::ToLower(key_path)];
            target->values.reserve(source->second.values.size());
        }
        if (entries)
        {
            // include empty keys so they remain selectable
            Entry key_entry;
            key_entry.source_path = source_path;
            key_entry.key_path = key_path;
            parsed_entries.push_back(std::move(key_entry));
        }
        for (const auto& pair : source->second.values)
        {
            if (cancel && cancel->load())
            {
                return false;
            }
            Value value;
            value.type = pair.second.type;
            value.raw = pair.second.data;
            value.data = value_format::DisplayData(pair.second.type, pair.second.data.empty() ? nullptr : pair.second.data.data(), static_cast<DWORD>(pair.second.data.size()));
            if (target)
            {
                auto& target_value = target->values[util::ToLower(pair.second.name)];
                if (entries)
                {
                    target_value = value;
                }
                else
                {
                    target_value = std::move(value);
                }
            }
            if (entries)
            {
                Entry entry;
                entry.source_path = source_path;
                entry.key_path = key_path;
                entry.has_value = true;
                entry.value_name = pair.second.name;
                entry.type = value.type;
                entry.raw = value.raw;
                entry.data = std::move(value.data);
                parsed_entries.push_back(std::move(entry));
            }
        }
    }

    if (!saw_key)
    {
        if (error)
        {
            *error = util::Tr(L"Default file contains no usable entries.");
        }
        return false;
    }
    if (data)
    {
        *data = std::move(loaded);
    }
    if (entries)
    {
        *entries = std::move(parsed_entries);
    }
    return true;
}

namespace
{

std::vector<std::wstring> SplitLabelWords(const std::wstring& text)
{
    std::vector<std::wstring> words;
    std::wstring word;
    for (wchar_t character : text)
    {
        if (character == L'-' || character == L'_' || character == L' ')
        {
            if (!word.empty())
            {
                words.push_back(std::move(word));
                word.clear();
            }
            continue;
        }
        word.push_back(character);
    }
    if (!word.empty())
    {
        words.push_back(std::move(word));
    }
    return words;
}

bool IsReleaseWord(const std::wstring& word)
{
    return word.size() == 4 && iswdigit(word[0]) && iswdigit(word[1]) && (word[2] == L'H' || word[2] == L'h') &&
           iswdigit(word[3]);
}

std::wstring ShortWindowsName(const std::wstring& folder)
{
    const std::vector<std::wstring> words = SplitLabelWords(folder);
    if (words.empty() || words[0].size() < 2 || (words[0][0] != L'W' && words[0][0] != L'w'))
    {
        return {};
    }
    std::wstring text = words[0];
    if (words.size() > 1 && IsReleaseWord(words[1]))
    {
        text.push_back(L' ');
        text.append(words[1]);
    }
    return text;
}

}

std::wstring ShortLabel(const std::wstring& label, const std::wstring& source_path)
{
    const size_t leaf = source_path.find_last_of(L"\\/");
    if (leaf != std::wstring::npos && leaf > 0)
    {
        const size_t parent = source_path.find_last_of(L"\\/", leaf - 1);
        const size_t start = parent == std::wstring::npos ? 0 : parent + 1;
        std::wstring folder = ShortWindowsName(source_path.substr(start, leaf - start));
        if (!folder.empty())
        {
            return folder;
        }
    }
    static const wchar_t* const kHiveWords[] = {L"HKLM", L"HKCU", L"HKU", L"HKCR", L"HKCC", L"HKEY", L"LOCAL", L"MACHINE", L"USER", L"USERS", L"CURRENT", L"SYSTEM", L"SOFTWARE", L"DEFAULT", L"CLASSES"};
    const std::vector<std::wstring> words = SplitLabelWords(label);
    size_t first = 0;
    while (first < words.size())
    {
        bool hive_word = false;
        for (const wchar_t* hive : kHiveWords)
        {
            if (util::EqualsInsensitive(words[first], hive))
            {
                hive_word = true;
                break;
            }
        }
        if (!hive_word)
        {
            break;
        }
        ++first;
    }
    if (first >= words.size())
    {
        return label;
    }
    std::wstring text = words[first];
    for (size_t i = first + 1; i < words.size(); ++i)
    {
        text.push_back(L' ');
        text.append(words[i]);
    }
    return text;
}

}
