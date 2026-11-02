// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/system/date_time.h>

namespace regkit
{
using namespace window_detail;

void MainWindow::Impl::StartValueListWorker()
{
    if (value_loader_.running())
    {
        return;
    }
    value_loader_.Start(L"ValueListThread", [this](std::unique_ptr<ValueListTask> task, const std::atomic_bool& stopping) {
        if (!task || stopping.load() || task->generation != value_list_generation_.load())
        {
            return;
        }

        auto payload = std::make_unique<ValueListPayload>();
        payload->generation = task->generation;
        const std::wstring parent_native = ToLower(registry_path::BuildNative(task->snapshot));
        const bool children_may_be_hives = task->hive_roots && !parent_native.empty() &&
                                           std::any_of(task->hive_roots->begin(), task->hive_roots->end(), [&](const std::wstring& root) {
                                               return registry_path::Parent(root) == parent_native;
                                           });
        auto key_icon = [&](const RegistryNode& node, const KeyInspection& inspection) -> int {
            if (inspection.link)
            {
                return kSymlinkIconIndex;
            }
            if (children_may_be_hives && task->hive_roots->contains(ToLower(registry_path::BuildNative(node))))
            {
                return inspection.denied ? kDatabaseDeniedIconIndex : kDatabaseIconIndex;
            }
            return inspection.denied ? kFolderDeniedIconIndex : inspection.is_volatile ? kFolderVolatileIconIndex
                                                                                       : kFolderIconIndex;
        };

        std::vector<std::wstring> subkeys;
        const bool subkeys_readable = RegistryStore::EnumKeyStreaming(
            task->snapshot,
            false,
            false,
            true,
            nullptr,
            RegistryStore::ValueStreamCallback(),
            [&](const std::wstring& name) {
                subkeys.push_back(name);
                return true;
            },
            MAXDWORD,
            nullptr,
            false
        );
        std::unordered_set<std::wstring> existing_keys;
        existing_keys.reserve(subkeys.size());
        for (const auto& name : subkeys)
        {
            existing_keys.insert(ToLower(name));
        }
        std::vector<std::wstring> simulated_subkeys;
        auto append_trace_children = [&](const RegistryNode& node,
                                         const std::unordered_set<std::wstring>& existing_lower,
                                         std::vector<std::wstring>* out) {
            if (!out)
            {
                return;
            }
            out->clear();
            if (!task->show_simulated_keys)
            {
                return;
            }
            if (!subkeys_readable && !node.simulated)
            {
                return;
            }
            if (task->trace_data_list.empty())
            {
                return;
            }
            if (!node.root_name.empty() && EqualsInsensitive(node.root_name, L"REGISTRY"))
            {
                return;
            }
            CollectTraceChildren(task->trace_data_list, TracePathLowerForNode(node), existing_lower, out);
        };
        append_trace_children(task->snapshot, existing_keys, &simulated_subkeys);
        payload->key_count = static_cast<int>(subkeys.size() + simulated_subkeys.size());
        payload->rows.reserve((task->show_keys_in_list ? subkeys.size() + simulated_subkeys.size() : 0) + 16);

        struct TraceMatch
        {
            std::wstring label;
            trace::KeyValues values;
            const trace::Selection* selection = nullptr;
        };
        std::vector<TraceMatch> trace_matches;
        if (!task->trace_data_list.empty())
        {
            for (const auto& trace : task->trace_data_list)
            {
                if (!trace.data)
                {
                    continue;
                }
                std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
                if (!trace.selection || !trace::IncludesKey(*trace.selection, task->trace_path_lower))
                {
                    continue;
                }
                auto it = trace.data->values_by_key.find(task->trace_path_lower);
                if (it == trace.data->values_by_key.end())
                {
                    continue;
                }
                TraceMatch match;
                match.label = trace.label.empty() ? util::Tr(L"Trace") : trace.label;
                match.values = it->second;
                match.selection = trace.selection.get();
                trace_matches.push_back(std::move(match));
            }
        }

        struct DefaultMatch
        {
            std::wstring label;
            defaults::Key values;
            const trace::Selection* selection = nullptr;
        };
        std::vector<DefaultMatch> default_keys;
        if (!task->default_data_list.empty() && !task->default_path_lower.empty())
        {
            default_keys.reserve(task->default_data_list.size());
            for (const auto& defaults : task->default_data_list)
            {
                if (!defaults.data)
                {
                    continue;
                }
                std::shared_lock<std::shared_mutex> defaults_lock(*defaults.data->mutex);
                if (!defaults.selection || !trace::IncludesKey(*defaults.selection, task->default_path_lower))
                {
                    continue;
                }
                auto it = defaults.data->values_by_key.find(task->default_path_lower);
                if (it == defaults.data->values_by_key.end())
                {
                    continue;
                }
                default_keys.push_back(
                    {defaults::ShortLabel(defaults.label, defaults.source_path), it->second, defaults.selection.get()}
                );
            }
        }
        auto resolve_default_data = [&](const std::wstring& value_name) -> std::wstring {
            if (default_keys.empty())
            {
                return {};
            }
            std::wstring value_lower = ToLower(value_name);
            struct DefaultGroup
            {
                std::wstring text;
                std::wstring labels;
            };
            std::vector<DefaultGroup> groups;
            for (const auto& match : default_keys)
            {
                if (match.selection && !trace::IncludesValue(*match.selection, task->default_path_lower, value_lower))
                {
                    continue;
                }
                auto it = match.values.values.find(value_lower);
                std::wstring text = it != match.values.values.end() ? it->second.data : std::wstring(util::Tr(L"(Missing)"));
                auto group = std::find_if(groups.begin(), groups.end(), [&](const DefaultGroup& entry) { return entry.text == text; });
                if (group == groups.end())
                {
                    groups.push_back({std::move(text), match.label});
                    continue;
                }
                if (!match.label.empty())
                {
                    if (!group->labels.empty())
                    {
                        group->labels.append(L", ");
                    }
                    group->labels.append(match.label);
                }
            }
            if (groups.empty())
            {
                return {};
            }
            if (groups.size() == 1)
            {
                return std::move(groups.front().text);
            }
            std::wstring text;
            for (size_t i = 0; i < groups.size(); ++i)
            {
                if (i > 0)
                {
                    text.append(L" - ");
                }
                text.append(groups[i].text);
                if (!groups[i].labels.empty())
                {
                    text.append(L" (").append(groups[i].labels).append(L")");
                }
            }
            return text;
        };

        if (task->show_keys_in_list)
        {
            for (const auto& name : subkeys)
            {
                ListRow row;
                row.name = registry_path::DisplayName(name);
                const RegistryNode child = registry_path::ChildNode(task->snapshot, name);
                const KeyInspection inspection = RegistryStore::InspectKey(child, task->include_dates || task->include_details, task->include_details);
                row.image_index = key_icon(child, inspection);
                row.type = inspection.link ? util::Tr(L"Link") : inspection.is_volatile ? util::Tr(L"Volatile Key")
                                                                                        : util::Tr(L"Key");
                row.extra = name;
                row.kind = rowkind::kKey;
                const KeyInfo& info = inspection.info;
                if (inspection.info_valid && task->include_dates)
                {
                    row.date = FormatFileTime(info.last_write);
                    row.date_value = win32ui::FileTimeTicks(info.last_write);
                    row.has_date = (row.date_value != 0);
                }
                if (inspection.info_valid && task->include_details)
                {
                    row.detail_key_count = info.subkey_count;
                    row.detail_value_count = info.value_count;
                    row.has_details = true;
                    row.details = util::TrLabel(L"Keys", std::to_wstring(info.subkey_count)) + L", " + util::TrLabel(L"Values", std::to_wstring(info.value_count));
                    if (inspection.class_source != ClassSource::kNone)
                    {
                        row.details.insert(0, std::wstring(inspection.class_source == ClassSource::kMachine ? util::Tr(L"Machine") : util::Tr(L"User")) + L", ");
                    }
                }
                payload->rows.emplace_back(std::move(row));
            }
            for (const auto& name : simulated_subkeys)
            {
                if (name.empty())
                {
                    continue;
                }
                ListRow row;
                row.name = registry_path::DisplayName(name);
                row.image_index = kFolderSimIconIndex;
                row.simulated = true;
                row.type = util::Tr(L"Key");
                row.extra = name;
                row.kind = rowkind::kKey;
                payload->rows.emplace_back(std::move(row));
            }
        }

        std::wstring link_target;
        bool has_link = RegistryStore::QuerySymbolicLinkTarget(task->snapshot, &link_target);
        bool track_existing = (!trace_matches.empty()) || has_link;
        std::unordered_set<std::wstring> existing_values;
        if (track_existing)
        {
            existing_values.reserve(64);
        }

        auto gather_labels = [&](const std::wstring& value_lower) -> std::vector<std::wstring> {
            std::vector<std::wstring> labels;
            for (const auto& match : trace_matches)
            {
                if (match.values.values_lower.find(value_lower) != match.values.values_lower.end())
                {
                    if (match.selection && !trace::IncludesValue(*match.selection, task->trace_path_lower, value_lower))
                    {
                        continue;
                    }
                    labels.push_back(match.label);
                }
            }
            if (labels.size() < 2)
            {
                return labels;
            }
            std::vector<std::wstring> unique;
            unique.reserve(labels.size());
            std::unordered_set<std::wstring> seen;
            for (const auto& label : labels)
            {
                std::wstring key = ToLower(label);
                if (seen.insert(key).second)
                {
                    unique.push_back(label);
                }
            }
            return unique;
        };
        auto format_read_on_boot = [&](const std::vector<std::wstring>& labels) -> std::wstring {
            if (labels.empty())
            {
                return util::Tr(L"No");
            }
            std::wstring out = util::Tr(L"Yes") + std::wstring(L" (");
            for (size_t i = 0; i < labels.size(); ++i)
            {
                if (i > 0)
                {
                    out.append(L", ");
                }
                out.append(labels[i]);
            }
            out.push_back(L')');
            return out;
        };
        bool have_traces = !trace_matches.empty();

        bool has_default = false;
        bool has_symbolic_value = false;
        constexpr DWORD kEagerValueDataLimit = 64u * 1024u;
        const DWORD max_data_size = task->include_all_value_data ? MAXDWORD : kEagerValueDataLimit;
        auto append_value = [&](const ValueInfo& value, const BYTE* data, DWORD data_size) -> bool {
            if (value.name.empty())
            {
                has_default = true;
            }
            if (EqualsInsensitive(value.name, L"SymbolicLinkValue"))
            {
                has_symbolic_value = true;
            }
            ListRow row = MakeValueListRow(value.name, value.type, data, data_size, !task->include_all_value_data);
            row.default_data = resolve_default_data(value.name);
            if (!have_traces)
            {
                row.read_on_boot.clear();
            }
            else
            {
                std::wstring lower = ToLower(value.name);
                row.read_on_boot = format_read_on_boot(gather_labels(lower));
                if (track_existing)
                {
                    existing_values.insert(lower);
                }
            }
            payload->rows.emplace_back(std::move(row));
            ++payload->value_count;
            return !stopping.load() && task->generation == value_list_generation_.load();
        };
        RegistryStore::KeyEnumResult listed;
        RegistryStore::EnumKeyStreaming(task->snapshot, true, true, false, &listed, append_value, RegistryStore::SubkeyStreamCallback(), max_data_size);
        // trace only simulated keys are expected to be absent
        payload->error = task->snapshot.simulated ? ERROR_SUCCESS : listed.error;

        if (!has_symbolic_value && has_link && !link_target.empty())
        {
            ListRow row;
            row.name = L"SymbolicLinkValue";
            row.type = L"REG_LINK";
            row.data = link_target;
            row.data_ready = true;
            row.image_index = UseBinaryValueIcon(REG_LINK) ? kBinaryIconIndex : kValueIconIndex;
            row.kind = rowkind::kValue;
            row.extra = L"SymbolicLinkValue";
            row.default_data = resolve_default_data(row.extra);
            DWORD link_bytes = static_cast<DWORD>((link_target.size() + 1) * sizeof(wchar_t));
            row.size = std::to_wstring(link_bytes);
            row.size_value = link_bytes;
            row.value_data_size = link_bytes;
            row.has_size = true;
            row.value_type = REG_LINK;
            row.read_on_boot = have_traces ? util::Tr(L"No") : L"";
            row.simulated = true;
            payload->rows.emplace_back(std::move(row));
        }

        if (!has_default)
        {
            ListRow row;
            row.name = util::Tr(L"(Default)");
            row.type = L"REG_SZ";
            row.data = util::Tr(L"(value not set)");
            row.data_ready = true;
            row.image_index = kValueIconIndex;
            row.kind = rowkind::kValue;
            row.extra = L"";
            row.default_data = resolve_default_data(row.extra);
            row.size = L"0";
            row.size_value = 0;
            row.has_size = true;
            row.value_type = REG_SZ;
            row.missing = true;
            if (!have_traces)
            {
                row.read_on_boot.clear();
            }
            else
            {
                row.read_on_boot = format_read_on_boot(gather_labels(L""));
                if (track_existing)
                {
                    existing_values.insert(L"");
                }
            }
            payload->rows.emplace_back(std::move(row));
            payload->value_count += 1;
        }

        size_t trace_added = 0;
        if (!trace_matches.empty())
        {
            for (const auto& match : trace_matches)
            {
                payload->rows.reserve(payload->rows.size() + match.values.values_display.size());
                for (const auto& value_name : match.values.values_display)
                {
                    std::wstring value_lower = ToLower(value_name);
                    if (match.selection && !trace::IncludesValue(*match.selection, task->trace_path_lower, value_lower))
                    {
                        continue;
                    }
                    if (existing_values.find(value_lower) != existing_values.end())
                    {
                        continue;
                    }
                    ListRow row;
                    row.name = value_name.empty() ? util::Tr(L"(Default)") : value_name;
                    row.type = L"TRACE";
                    row.missing = true;
                    row.data = util::Tr(L"(value not set)");
                    row.read_on_boot = format_read_on_boot(gather_labels(value_lower));
                    row.image_index = kTraceIconIndex;
                    row.kind = rowkind::kValue;
                    row.extra = value_name;
                    row.data_ready = true;
                    row.default_data = resolve_default_data(value_name);
                    payload->rows.emplace_back(std::move(row));
                    ++trace_added;
                    existing_values.insert(value_lower);
                }
            }
            payload->value_count += static_cast<int>(trace_added);
        }

        if (task->sort_column != kValueColComment)
        {
            SortValueRows(&payload->rows, task->sort_column, task->sort_ascending, task->sort_earlier);
        }
        if (stopping.load() || task->generation != value_list_generation_.load())
        {
            return;
        }
        win32ui::PostPayload(task->hwnd, frame::message_id::kValueListReady, static_cast<WPARAM>(task->generation), payload);
    });
}

void MainWindow::Impl::StartValuePreviewWorker()
{
    if (value_preview_loader_.running())
    {
        return;
    }
    value_preview_loader_.Start(L"ValuePreviewThread", [this](std::unique_ptr<ValuePreviewTask> task, const std::atomic_bool& stopping) {
        if (!task || task->generation != value_list_generation_.load())
        {
            return;
        }
        auto payload = std::make_unique<ValuePreviewPayload>();
        payload->generation = task->generation;
        payload->items.reserve(task->indices.size());
        for (size_t i = 0; i < task->indices.size(); ++i)
        {
            if (stopping.load() || task->generation != value_list_generation_.load())
            {
                return;
            }
            RegistryValue entry;
            ValuePreviewItem item;
            item.index = task->indices[i];
            item.name = task->names[i];
            const bool ok = RegistryStore::QueryValue(task->snapshot, item.name, &entry);
            if (ok)
            {
                item.type = entry.type;
                item.size = static_cast<DWORD>(entry.data.size());
                item.preview = value_format::DisplayData(entry.type, entry.data.data(), std::min<DWORD>(item.size, kValuePreviewBytes));
                if (item.preview.size() > kValuePreviewLimit)
                {
                    item.preview.resize(kValuePreviewLimit);
                }
            }
            payload->items.push_back(std::move(item));
        }
        if (stopping.load() || task->generation != value_list_generation_.load())
        {
            return;
        }
        win32ui::PostPayload(task->hwnd, frame::message_id::kValuePreviewReady, static_cast<WPARAM>(task->generation), payload);
    });
}

void MainWindow::Impl::QueueValuePreviews(int first, int last)
{
    if (!browse_.current_node() || first < 0 || last < first)
    {
        return;
    }
    auto task = std::make_unique<ValuePreviewTask>();
    bool emptied = false;
    for (int i = first; i <= last; ++i)
    {
        ListRow* row = browse_.values().MutableRowAt(i);
        if (!row || row->kind != rowkind::kValue || row->data_ready)
        {
            continue;
        }
        if (row->value_data_size == 0)
        {
            row->data.clear();
            row->data_ready = true;
            browse_.values().InvalidateFilterCache(row);
            emptied = true;
            continue;
        }
        task->indices.push_back(i);
        task->names.push_back(row->extra);
    }
    if (emptied)
    {
        browse_.values().RefreshFilter();
    }
    if (task->indices.empty())
    {
        return;
    }
    StartValuePreviewWorker();
    task->generation = value_list_generation_.load();
    task->snapshot = *browse_.current_node();
    task->hwnd = hwnd_;
    value_preview_loader_.Submit(std::move(task));
}

void MainWindow::Impl::StartSearchPreviewWorker()
{
    if (search_preview_loader_.running())
    {
        return;
    }
    search_preview_loader_.Start(L"SearchPreviewThread", [this](std::unique_ptr<SearchPreviewTask> task, const std::atomic_bool& stopping) {
        if (!task)
        {
            return;
        }
        auto payload = std::make_unique<SearchPreviewPayload>();
        payload->generation = task->generation;
        payload->tab_index = task->tab_index;
        payload->items.reserve(task->requests.size());
        for (const auto& request : task->requests)
        {
            if (stopping.load())
            {
                return;
            }
            RegistryValue entry;
            SearchPreviewItem item;
            item.index = request.index;
            item.row_id = request.row_id;
            if (RegistryStore::QueryValue(request.node, request.value_name, &entry))
            {
                item.type = entry.type;
                item.data_size = static_cast<DWORD>(entry.data.size());
                item.preview = value_format::DisplayData(entry.type, entry.data.data(), std::min<DWORD>(item.data_size, kValuePreviewBytes));
                if (item.preview.size() > kValuePreviewLimit)
                {
                    item.preview.resize(kValuePreviewLimit);
                }
            }
            payload->items.push_back(std::move(item));
        }
        if (stopping.load())
        {
            return;
        }
        win32ui::PostPayload(task->hwnd, frame::message_id::kSearchPreviewReady, static_cast<WPARAM>(task->generation), payload);
    });
}

std::vector<std::vector<RegistryRootEntry>> MainWindow::Impl::SearchSourceRoots(const SearchTab& tab)
{
    std::vector<std::vector<RegistryRootEntry>> roots(tab.sources.size());
    for (size_t i = 0; i < tab.sources.size(); ++i)
    {
        const search::Source& source = tab.sources[i];
        if (source.kind == search::Source::Kind::kLocal)
        {
            roots[i] = LocalRoots(source.view);
            continue;
        }
        // closed/disconnected sources have no roots, show no row data
        const int index = FindSourceTab(source);
        if (index < 0)
        {
            continue;
        }
        const TabEntry& entry = tabs_[static_cast<size_t>(index)];
        if (entry.kind == TabEntry::Kind::kRegFile)
        {
            for (const auto& root : entry.reg_file_roots)
            {
                if (root.root)
                {
                    roots[i].push_back({root.root, root.name, root.name});
                }
            }
        }
        else if (entry.session)
        {
            roots[i] = entry.session->roots;
        }
    }
    return roots;
}

void MainWindow::Impl::QueueSearchPreviews(int first, int last)
{
    const int index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
    if (index < 0 || static_cast<size_t>(index) >= search_tabs_.size())
    {
        return;
    }
    SearchTab& tab = search_tabs_[static_cast<size_t>(index)];
    if (tab.is_compare || first < 0 || last < first)
    {
        return;
    }
    auto task = std::make_unique<SearchPreviewTask>();
    const size_t count = tab.results.size();
    const auto source_roots = SearchSourceRoots(tab);
    for (int i = first; i <= last && static_cast<size_t>(i) < count; ++i)
    {
        search::Result& row = tab.results[static_cast<size_t>(i)];
        if (row.data_state != search::DataState::kNotLoaded || search::IsKeyRow(row))
        {
            continue;
        }
        SearchPreviewRequest request;
        if (!ResolvePathInRoots(row.source < source_roots.size() ? source_roots[row.source] : browse_.roots(), row.key_path, &request.node))
        {
            row.data_text.clear();
            row.data_state = search::DataState::kLoaded;
            continue;
        }
        request.index = i;
        request.row_id = row.row_id;
        request.value_name = row.value_name;
        task->requests.push_back(std::move(request));
    }
    if (task->requests.empty())
    {
        return;
    }
    StartSearchPreviewWorker();
    task->generation = tab.generation;
    task->tab_index = index;
    task->hwnd = hwnd_;
    search_preview_loader_.Submit(std::move(task));
}

void MainWindow::Impl::StartSearchSortWorker()
{
    if (search_sort_loader_.running())
    {
        return;
    }
    search_sort_loader_.Start(L"SearchSortThread", [this](std::unique_ptr<SearchSortTask> task, const std::atomic_bool& stopping) {
        if (!task)
        {
            return;
        }
        for (size_t i = 0; i < task->rows.size(); ++i)
        {
            if (stopping.load())
            {
                return;
            }
            search::Result& row = task->rows[i];
            if (row.data_state != search::DataState::kNotLoaded)
            {
                continue;
            }
            row.data_state = search::DataState::kLoaded;
            RegistryValue entry;
            if (!RegistryStore::QueryValue(task->nodes[i], row.value_name, &entry))
            {
                row.data_text.clear();
                continue;
            }
            row.type = entry.type;
            row.data_size = static_cast<DWORD>(entry.data.size());
            std::wstring preview = value_format::DisplayData(entry.type, entry.data.data(), std::min<DWORD>(row.data_size, kValuePreviewBytes));
            if (preview.size() > kValuePreviewLimit)
            {
                preview.resize(kValuePreviewLimit);
            }
            row.data_text = std::move(preview);
        }
        if (stopping.load())
        {
            return;
        }
        search::SortResults(&task->rows, task->column, task->ascending, task->earlier, task->sources);
        auto payload = std::make_unique<SearchSortPayload>();
        payload->generation = task->generation;
        payload->tab_index = task->tab_index;
        payload->sort_revision = task->sort_revision;
        payload->rows = std::move(task->rows);
        win32ui::PostPayload(task->hwnd, frame::message_id::kSearchSortReady, static_cast<WPARAM>(task->generation), payload);
    });
}

void MainWindow::Impl::QueueSearchSort(SearchTab* tab, const util::SortKeys& earlier)
{
    if (!tab)
    {
        return;
    }
    int tab_index = -1;
    for (size_t i = 0; i < search_tabs_.size(); ++i)
    {
        if (&search_tabs_[i] == tab)
        {
            tab_index = static_cast<int>(i);
            break;
        }
    }
    if (tab_index < 0)
    {
        return;
    }
    auto task = std::make_unique<SearchSortTask>();
    task->rows = tab->results;
    task->nodes.resize(task->rows.size());
    const auto source_roots = SearchSourceRoots(*tab);
    for (size_t i = 0; i < task->rows.size(); ++i)
    {
        if (task->rows[i].data_state == search::DataState::kNotLoaded)
        {
            RegistryNode node;
            const uint16_t source = task->rows[i].source;
            if (ResolvePathInRoots(source < source_roots.size() ? source_roots[source] : browse_.roots(), task->rows[i].key_path, &node))
            {
                task->nodes[i] = std::move(node);
            }
            else
            {
                task->rows[i].data_state = search::DataState::kLoaded;
            }
        }
    }
    StartSearchSortWorker();
    task->generation = tab->generation;
    task->tab_index = tab_index;
    task->column = tab->sort_column;
    task->ascending = tab->sort_ascending;
    task->earlier = earlier;
    task->sort_revision = tab->sort_revision;
    task->sources = tab->sources;
    task->hwnd = hwnd_;
    search_sort_loader_.Submit(std::move(task));
}

void MainWindow::Impl::StartSearchTabLoadWorker()
{
    if (search_tab_loader_.running())
    {
        return;
    }
    // newer task is the load of another tab, this one still posts, or its tab would keep load_pending and never load
    search_tab_loader_.Start(L"SearchTabLoadThread", [](std::unique_ptr<SearchTabLoadTask> task, const std::atomic_bool&) {
        if (!task)
        {
            return;
        }
        auto payload = std::make_unique<SearchTabLoadPayload>();
        payload->generation = task->generation;
        payload->tab_index = task->tab_index;
        payload->sort_revision = task->sort_revision;

        payload->path = task->path;
        payload->ok = search::LoadResults(task->path, &payload->rows, &payload->error);
        if (payload->ok && task->sort_column >= 0)
        {
            search::SortResults(&payload->rows, task->sort_column, task->sort_ascending, task->sort_earlier, task->sources);
        }
        win32ui::PostPayload(task->hwnd, frame::message_id::kSearchTabLoadReady, static_cast<WPARAM>(task->generation), payload);
    });
}

void MainWindow::Impl::StopValueListWorker()
{
    value_loader_.Stop();
    value_preview_loader_.Stop();
    search_preview_loader_.Stop();
    search_sort_loader_.Stop();
    search_tab_loader_.Stop();
}

}
