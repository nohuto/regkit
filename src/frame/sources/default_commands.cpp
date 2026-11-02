// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"
#include "win32/translation.h"
#include "trace/trace_paths.h"

namespace regkit
{
using namespace window_detail;

bool MainWindow::Impl::AddDefaultFromFile(const std::wstring& label, const std::wstring& path, bool show_error, bool prompt_for_selection, bool update_ui)
{
    if (path.empty())
    {
        return false;
    }
    std::wstring source = path;
    std::wstring use_label = label;
    if (!util::IsFile(source.c_str()))
    {
        std::wstring bundled = ResolveBundledDefaultPath(path);
        if (!bundled.empty() && util::IsFile(bundled.c_str()))
        {
            source = bundled;
            if (use_label.empty())
            {
                use_label = path;
            }
        }
        else
        {
            if (show_error)
            {
                ui::ShowError(hwnd_, util::Tr(L"Default file not found."));
            }
            return false;
        }
    }
    if (use_label.empty())
    {
        use_label = util::FileBaseName(source);
    }
    if (use_label.empty())
    {
        use_label = L"Default";
    }
    for (const auto& defaults : active_defaults_)
    {
        if (EqualsInsensitive(defaults.source_path, source))
        {
            return false;
        }
    }
    std::wstring source_lower = ToLower(source);
    if (default_parse_sessions_.find(source_lower) != default_parse_sessions_.end())
    {
        return false;
    }

    trace::Selection selection = {};
    selection.select_all = true;
    selection.recursive = true;

    auto session = std::make_unique<DefaultParseSession>();
    session->label = use_label;
    session->source_path = source;
    session->source_lower = source_lower;
    session->data = std::make_shared<defaults::Data>();
    session->selection = selection;
    session->show_errors = show_error;

    DefaultParseSession* session_ptr = session.get();
    default_parse_sessions_.emplace(source_lower, std::move(session));

    if (prompt_for_selection)
    {
        trace::Selection dialog_selection = selection;
        TraceDialogOptions options;
        options.title = util::Tr(L"Default entries") + (use_label.empty() ? std::wstring() : L" - " + use_label);
        options.show_values = true;
        DefaultDialogStartContext context;
        context.window = this;
        context.session = session_ptr;
        if (!ShowTraceDialog(hwnd_, options, &dialog_selection, StartDefaultDialogLoad, &context))
        {
            session_ptr->work.CancelAndJoin();
            default_parse_sessions_.erase(source_lower);
            return false;
        }
        session_ptr->dialog = nullptr;
        session_ptr->selection = std::move(dialog_selection);
    }
    else
    {
        StartDefaultParseThread(session_ptr);
    }

    if (!session_ptr->selection.select_all && session_ptr->selection.key_paths.empty() &&
        session_ptr->selection.values_by_key.empty())
    {
        session_ptr->selection.select_all = true;
    }

    session_ptr->added_to_active = true;
    active_defaults_.push_back(
        {use_label, source, session_ptr->data, std::make_shared<trace::Selection>(session_ptr->selection)}
    );
    if (update_ui)
    {
        SaveActiveDefaults();
        UpdateValueListForNode(browse_.current_node());
        SaveSettings();
    }
    if (session_ptr->parsing_done)
    {
        session_ptr->work.Join();
    }
    if (session_ptr->parsing_done && !session_ptr->dialog)
    {
        default_parse_sessions_.erase(source_lower);
    }
    return true;
}

bool MainWindow::Impl::LoadDefaultFromFile(const std::wstring& label, const std::wstring& path)
{
    return AddDefaultFromFile(label, path, true, true, true);
}

bool MainWindow::Impl::LoadDefaultFromPrompt()
{
    std::wstring path;
    if (!ui::PromptOpenFile(hwnd_, ui::kRegFileFilter, &path))
    {
        return false;
    }
    std::wstring label = util::FileBaseName(path);
    if (label.empty())
    {
        label = L"Custom";
    }
    if (!LoadDefaultFromFile(label, path))
    {
        return false;
    }
    AddRecentDefaultPath(path);
    SaveSettings();
    return true;
}

void MainWindow::Impl::ClearDefaults()
{
    default_load_session_.Cancel();
    StopDefaultParseSessions();
    active_defaults_.clear();
    SaveActiveDefaults();
    UpdateValueListForNode(browse_.current_node());
    SaveSettings();
}

void MainWindow::Impl::AddRecentTracePath(const std::wstring& path)
{
    recent_trace_paths_.Add(util::TrimWhitespace(path));
}

void MainWindow::Impl::AddRecentDefaultPath(const std::wstring& path)
{
    recent_default_paths_.Add(util::TrimWhitespace(path));
}

std::vector<MainWindow::Impl::DefaultValueChoice> MainWindow::Impl::CollectDefaultChoices(
    const std::wstring& value_name
) const
{
    std::vector<DefaultValueChoice> choices;
    const RegistryNode* node = browse_.current_node();
    if (!node || active_defaults_.empty())
    {
        return choices;
    }
    std::wstring path = SourceLookupPath(*node);
    std::wstring default_path = trace::NormalizeKeyPathBasic(path);
    if (default_path.empty())
    {
        default_path = path;
    }
    const std::wstring key_lower = ToLower(default_path);
    const std::wstring value_lower = ToLower(value_name);
    for (const auto& defaults : active_defaults_)
    {
        if (!defaults.data || !defaults.selection)
        {
            continue;
        }
        if (!trace::IncludesKey(*defaults.selection, key_lower) ||
            !trace::IncludesValue(*defaults.selection, key_lower, value_lower))
        {
            continue;
        }
        std::shared_lock<std::shared_mutex> lock(*defaults.data->mutex);
        auto key = defaults.data->values_by_key.find(key_lower);
        if (key == defaults.data->values_by_key.end())
        {
            continue;
        }
        DefaultValueChoice choice;
        choice.label = defaults::ShortLabel(defaults.label, defaults.source_path);
        auto value = key->second.values.find(value_lower);
        if (value != key->second.values.end())
        {
            choice.present = true;
            choice.type = value->second.type;
            choice.data = value->second.raw;
        }
        choices.push_back(std::move(choice));
    }
    return choices;
}

}
