// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::StartTraceDialogLoad(HWND hwnd, void* context) {
  auto* ctx = reinterpret_cast<TraceDialogStartContext*>(context);
  if (!ctx || !ctx->window || !ctx->session) {
    return;
  }
  ctx->session->dialog = hwnd;
  ctx->window->StartTraceParseThread(ctx->session);
}

void MainWindow::Impl::StartDefaultDialogLoad(HWND hwnd, void* context) {
  auto* ctx = reinterpret_cast<DefaultDialogStartContext*>(context);
  if (!ctx || !ctx->window || !ctx->session) {
    return;
  }
  ctx->session->dialog = hwnd;
  ctx->window->StartDefaultParseThread(ctx->session);
}

bool MainWindow::Impl::AllowTraceSimulation(const RegistryNode& node) const {
  if (active_traces_.empty()) {
    return false;
  }
  if (!show_simulated_keys_) {
    return false;
  }
  if (!node.root_name.empty() && EqualsInsensitive(node.root_name, L"REGISTRY")) {
    return false;
  }
  return true;
}

std::wstring MainWindow::Impl::TracePathLowerForNode(const RegistryNode& node) const {
  std::wstring path = registry_path::Build(node);
  std::wstring trace_path = NormalizeTraceKeyPath(path);
  if (trace_path.empty()) {
    trace_path = path;
  }
  return ToLower(trace_path);
}

void MainWindow::Impl::AppendTraceChildren(const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) const {
  if (!out) {
    return;
  }
  out->clear();
  if (IsRegFileTabSelected()) {
    return;
  }
  if (!AllowTraceSimulation(node)) {
    return;
  }
  if (active_traces_.empty()) {
    return;
  }
  std::wstring key_lower = TracePathLowerForNode(node);
  if (key_lower.empty()) {
    return;
  }
  std::unordered_set<std::wstring> seen;
  for (const auto& trace : active_traces_) {
    if (!trace.data) {
      continue;
    }
    std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
    if (!trace.selection ||
        !trace::IncludesKey(*trace.selection, key_lower)) {
      continue;
    }
    auto it = trace.data->children_by_key.find(key_lower);
    if (it == trace.data->children_by_key.end()) {
      continue;
    }
    for (const auto& name : it->second) {
      if (name.empty()) {
        continue;
      }
      std::wstring name_lower = ToLower(name);
      if (existing_lower.find(name_lower) != existing_lower.end()) {
        continue;
      }
      if (!seen.insert(name_lower).second) {
        continue;
      }
      out->push_back(name);
    }
  }
  std::sort(out->begin(), out->end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
}

std::wstring MainWindow::Impl::ResolveBundledTracePath(const std::wstring& label) const {
  std::wstring file = TrimWhitespace(label);
  if (file.empty()) {
    return L"";
  }
  if (file.size() < 4 || _wcsicmp(file.c_str() + file.size() - 4, L".txt") != 0) {
    file.append(L".txt");
  }

  std::wstring module_dir = util::GetModuleDirectory();
  if (module_dir.empty()) {
    return L"";
  }
  std::wstring assets = util::JoinPath(module_dir, L"assets");
  return util::JoinPath(util::JoinPath(assets, L"records"), file);
}

bool MainWindow::Impl::LoadBundledTrace(
    const std::wstring& label,
    const trace::Selection* selection_override) {
  std::wstring path = ResolveBundledTracePath(label);
  if (path.empty()) {
    return false;
  }
  return LoadTraceFromFile(label, path, selection_override);
}

std::wstring MainWindow::Impl::ResolveBundledDefaultPath(const std::wstring& label) const {
  std::wstring file = TrimWhitespace(label);
  if (file.empty()) {
    return L"";
  }
  if (!HasRegExtension(file)) {
    file.append(L".reg");
  }

  std::wstring module_dir = util::GetModuleDirectory();
  if (module_dir.empty()) {
    return L"";
  }
  std::wstring assets = util::JoinPath(module_dir, L"assets");
  std::wstring defaults = util::JoinPath(assets, L"defaults");
  std::wstring direct = util::JoinPath(defaults, file);
  if (FileExists(direct)) {
    return direct;
  }
  std::wstring requested = FileBaseName(file);
  for (const auto& entry : bundled_defaults_) {
    if (EqualsInsensitive(entry.label, FileBaseName(label)) ||
        EqualsInsensitive(FileBaseName(entry.path), requested)) {
      return entry.path;
    }
  }
  WIN32_FIND_DATAW data = {};
  HANDLE find =
      FindFirstFileW(util::JoinPath(defaults, L"*").c_str(), &data);
  if (find != INVALID_HANDLE_VALUE) {
    do {
      if (!(data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ||
          wcscmp(data.cFileName, L".") == 0 ||
          wcscmp(data.cFileName, L"..") == 0) {
        continue;
      }
      std::wstring directory = util::JoinPath(defaults, data.cFileName);
      std::wstring candidate = util::JoinPath(directory, file);
      if (FileExists(candidate)) {
        FindClose(find);
        return candidate;
      }
    } while (FindNextFileW(find, &data));
    FindClose(find);
  }
  return direct;
}

bool MainWindow::Impl::AddTraceFromFile(
    const std::wstring& label, const std::wstring& path,
    const trace::Selection* selection_override,
    bool prompt_for_selection, bool update_ui) {
  std::wstring source = TrimWhitespace(path);
  if (source.empty()) {
    return false;
  }
  std::wstring use_label = label;
  if (!FileExists(source)) {
    std::wstring candidate_label = use_label.empty() ? source : use_label;
    std::wstring bundled = ResolveBundledTracePath(candidate_label);
    if (!bundled.empty() && FileExists(bundled)) {
      source = bundled;
      if (use_label.empty()) {
        use_label = candidate_label;
      }
    } else {
      if (update_ui) {
        ui::ShowError(hwnd_, L"Trace file not found.");
      }
      return false;
    }
  }
  if (use_label.empty()) {
    use_label = FileBaseName(source);
  }
  if (use_label.empty()) {
    use_label = L"Trace";
  }
  for (const auto& trace : active_traces_) {
    if (EqualsInsensitive(trace.source_path, source)) {
      return true;
    }
  }
  std::wstring source_lower = ToLower(source);
  if (trace_parse_sessions_.find(source_lower) != trace_parse_sessions_.end()) {
    return true;
  }

  trace::Selection selection = {};
  selection.select_all = true;
  selection.recursive = true;
  if (selection_override) {
    selection = *selection_override;
  } else if (!prompt_for_selection) {
    auto it = trace_selection_cache_.find(source_lower);
    if (it != trace_selection_cache_.end()) {
      selection = it->second;
    }
  }

  auto session = std::make_unique<TraceParseSession>();
  session->label = use_label;
  session->source_path = source;
  session->source_lower = source_lower;
  session->data = std::make_shared<trace::Data>();
  session->data->label = use_label;
  session->data->source_path = source;
  session->selection = selection;

  TraceParseSession* session_ptr = session.get();
  trace_parse_sessions_.emplace(source_lower, std::move(session));

  if (prompt_for_selection) {
    trace::Selection dialog_selection = selection;
    TraceDialogOptions options;
    options.title = use_label.empty() ? L"Trace entries" : L"Trace entries - " + use_label;
    options.prompt = L"";
    options.show_values = true;
    TraceDialogStartContext context;
    context.window = this;
    context.session = session_ptr;
    if (!ShowTraceDialog(hwnd_, options, &dialog_selection, StartTraceDialogLoad, &context)) {
      session_ptr->work.CancelAndJoin();
      trace_parse_sessions_.erase(source_lower);
      return false;
    }
    session_ptr->dialog = nullptr;
    session_ptr->selection = std::move(dialog_selection);
  } else {
    StartTraceParseThread(session_ptr);
  }

  if (!session_ptr->selection.select_all && session_ptr->selection.key_paths.empty() && session_ptr->selection.values_by_key.empty()) {
    session_ptr->selection.select_all = true;
  }

  session_ptr->added_to_active = true;
  active_traces_.push_back(
      {use_label, source, session_ptr->data,
       std::make_shared<trace::Selection>(session_ptr->selection)});
  trace_selection_cache_[source_lower] = session_ptr->selection;

  if (update_ui) {
    SaveActiveTraces();
    SaveTraceSettings();
    BuildMenus();
    RefreshTreeSelection();
    UpdateValueListForNode(browse_.current_node());
    SaveSettings();
  }
  if (session_ptr->parsing_done) {
    session_ptr->work.Join();
  }
  if (session_ptr->parsing_done && !session_ptr->dialog) {
    trace_parse_sessions_.erase(source_lower);
  }
  return true;
}

bool MainWindow::Impl::LoadTraceFromFile(
    const std::wstring& label, const std::wstring& path,
    const trace::Selection* selection_override) {
  return AddTraceFromFile(label, path, selection_override, true, true);
}

bool MainWindow::Impl::LoadTraceFromPrompt() {
  std::wstring path;
  if (!PromptOpenFile(hwnd_, L"Trace Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0", &path)) {
    return false;
  }
  std::wstring label = FileBaseName(path);
  if (label.empty()) {
    label = L"Custom";
  }
  if (!LoadTraceFromFile(label, path)) {
    return false;
  }
  AddRecentTracePath(path);
  BuildMenus();
  SaveSettings();
  return true;
}

void MainWindow::Impl::ClearTrace() {
  StopTraceParseSessions();
  active_traces_.clear();
  trace_selection_cache_.clear();
  SaveActiveTraces();
  SaveTraceSettings();
  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(browse_.current_node());
  SaveSettings();
}

} // namespace regkit
