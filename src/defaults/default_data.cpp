// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "defaults/default_data.h"
#include "win32/text_transform.h"

#include <mutex>
#include <utility>

namespace regkit::defaults
{

void Merge(Data* data, const std::vector<Entry>& entries, const AliasPath& alias, std::unordered_set<std::wstring>* affected_keys)
{
    if (!data || entries.empty())
    {
        return;
    }
    std::vector<std::wstring> aliases(entries.size());
    for (size_t index = 0; alias && index < entries.size(); ++index)
    {
        if (entries[index].has_value && !entries[index].key_path.empty())
        {
            aliases[index] = util::ToLower(alias(entries[index].key_path));
        }
    }
    std::unique_lock<std::shared_mutex> lock(*data->mutex);
    for (size_t index = 0; index < entries.size(); ++index)
    {
        const Entry& entry = entries[index];
        if (entry.key_path.empty())
        {
            continue;
        }
        const std::wstring key = util::ToLower(entry.key_path);
        if (affected_keys)
        {
            affected_keys->insert(key);
        }
        if (!entry.has_value)
        {
            continue;
        }
        const std::wstring name = util::ToLower(entry.value_name);
        Value& value = data->values_by_key[key].values[name];
        value.type = entry.type;
        value.data = entry.data;
        value.raw = entry.raw;
        if (!aliases[index].empty())
        {
            data->values_by_key[aliases[index]].values[name] = value;
            if (affected_keys)
            {
                affected_keys->insert(aliases[index]);
            }
        }
    }
}

}
