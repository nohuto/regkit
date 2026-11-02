// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "trace/trace_parser.h"

#include "registry/registry_path.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <algorithm>
#include <cwctype>
#include <unordered_map>
#include <utility>

namespace regkit::trace
{

namespace
{

bool Cancelled(const std::atomic_bool* cancel)
{
    return cancel && cancel->load();
}

bool Decode(std::string_view buffer, std::wstring* content, std::wstring* error)
{
    if (!content)
    {
        return false;
    }
    if (buffer.empty())
    {
        if (error)
        {
            *error = util::Tr(L"Trace file is empty or too large to load.");
        }
        return false;
    }
    if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF &&
        static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF)
    {
        buffer.remove_prefix(3);
    }
    *content = win32ui::Utf8ToWide(buffer).value_or(std::wstring());
    if (content->empty())
    {
        if (error)
        {
            *error = util::Tr(L"Trace file has no readable entries.");
        }
        return false;
    }
    return true;
}

}

bool ParseEntries(std::string_view buffer, const Normalizers& normalizers, const EntryCallback& callback, std::wstring* error, const std::atomic_bool* cancel)
{
    if (!normalizers.key || !normalizers.display || !callback)
    {
        return false;
    }
    std::wstring content;
    if (!Decode(buffer, &content, error))
    {
        return false;
    }

    bool saw_entry = false;
    size_t start = 0;
    while (start < content.size())
    {
        if (Cancelled(cancel))
        {
            return false;
        }
        size_t end = content.find(L'\n', start);
        if (end == std::wstring::npos)
        {
            end = content.size();
        }
        std::wstring line = content.substr(start, end - start);
        start = end + 1;
        if (!line.empty() && line.back() == L'\r')
        {
            line.pop_back();
        }
        // trim keys only, preserve value name spaces such as bundled trace name "DssCounter "
        if (util::IsBlank(line))
        {
            continue;
        }

        size_t separator = line.rfind(L" : ");
        size_t separator_size = 3;
        if (separator == std::wstring::npos)
        {
            separator = line.rfind(L':');
            separator_size = 1;
        }
        if (separator == std::wstring::npos)
        {
            continue;
        }

        const std::wstring source_key = util::TrimWhitespace(line.substr(0, separator));
        if (source_key.empty())
        {
            continue;
        }
        Entry entry;
        entry.display_path = normalizers.display(source_key);
        if (entry.display_path.empty())
        {
            continue;
        }
        entry.key_path = normalizers.key(source_key);
        if (entry.key_path.empty())
        {
            entry.key_path = entry.display_path;
        }
        entry.value_name = line.substr(separator + separator_size);
        entry.has_value = true;
        if (util::EqualsInsensitive(entry.value_name, L"(Default)"))
        {
            entry.value_name.clear();
        }
        saw_entry = true;
        if (!callback(std::move(entry)))
        {
            return false;
        }
    }

    if (!saw_entry)
    {
        if (error)
        {
            *error = util::Tr(L"Trace file contains no usable entries.");
        }
        return false;
    }
    return true;
}

bool Parse(const std::wstring& label, const std::wstring& source, std::string_view buffer, const Normalizers& normalizers, Data* data, std::wstring* error, const std::atomic_bool* cancel)
{
    if (!data)
    {
        return false;
    }
    Data parsed;
    parsed.label = label;
    parsed.source_path = source;
    const bool ok = ParseEntries(
        buffer,
        normalizers,
        [&](Entry&& entry) {
            AddEntry(&parsed, entry);
            return !Cancelled(cancel);
        },
        error,
        cancel
    );
    if (!ok)
    {
        return false;
    }
    Sort(&parsed);
    *data = std::move(parsed);
    return true;
}
}
