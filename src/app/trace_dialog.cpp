// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#include "app/trace_dialog.h"

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <memory>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include <commctrl.h>

#include "app/theme.h"
#include "app/ui_helpers.h"

// Older SDKs omit the NMTVITEMCHANGE struct even though TVN_ITEMCHANGED is
// defined.
#ifndef NMTVITEMCHANGE
typedef struct tagNMTVITEMCHANGE {
  NMHDR hdr;
  UINT uChanged;
  HTREEITEM hItem;
  UINT uStateNew;
  UINT uStateOld;
  LPARAM lParam;
} NMTVITEMCHANGE;
#endif

namespace regkit {

namespace {

constexpr wchar_t kDialogClass[] = L"RegKitTraceDialog";
constexpr wchar_t kAppTitle[] = L"RegKit";
constexpr UINT kDialogAddEntriesMessage = WM_APP + 1;
constexpr UINT kDialogDoneMessage = WM_APP + 2;
constexpr UINT kDialogProcessEntriesMessage = WM_APP + 3;

enum ControlId {
  kTraceLabel = 100,
  kTraceStatus = 101,
  kTraceTree = 102,
  kRecursiveCheck = 103,
  kSelectAllButton = 104,
  kOkButton = IDOK,
  kCancelButton = IDCANCEL,
};

struct TraceNodeData {
  bool is_value = false;
  std::wstring key_path;
  std::wstring value_name;
};

struct TraceDialogState {
  HWND hwnd = nullptr;
  HWND label = nullptr;
  HWND status = nullptr;
  HWND tree = nullptr;
  HWND recursive = nullptr;
  HWND select_all = nullptr;
  HWND ok_button = nullptr;
  HWND cancel_button = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  TraceSelection* out = nullptr;
  bool accepted = false;
  bool owner_restored = false;
  bool loading_done = false;
  bool show_values = true;
  TraceDialogReadyCallback on_ready = nullptr;
  void* on_ready_context = nullptr;
  std::wstring prompt;
  bool processing_entries = false;
  bool updating_checks = false;
  size_t pending_index = 0;

  std::unordered_map<std::wstring, HTREEITEM> key_nodes;
  std::unordered_set<std::wstring> key_entries;
  std::unordered_map<std::wstring, std::unordered_map<std::wstring, std::wstring>> values_by_key;
  std::unordered_map<std::wstring, std::unordered_set<std::wstring>> values_loaded;
  std::vector<std::unique_ptr<TraceNodeData>> node_storage;
  std::vector<KeyValueDialogEntry> pending_entries;
  size_t key_count = 0;
  size_t value_count = 0;
};

HFONT CreateDialogFont() {
  return ui::DefaultUIFont();
}

void ApplyFont(HWND hwnd, HFONT font) {
  if (hwnd && font) {
    SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
  }
}

void RestoreOwnerWindow(HWND owner, bool* restored) {
  if (!owner || !restored || *restored) {
    return;
  }
  EnableWindow(owner, TRUE);
  SetActiveWindow(owner);
  SetForegroundWindow(owner);
  *restored = true;
}

std::wstring ToLower(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (wchar_t ch : text) {
    out.push_back(static_cast<wchar_t>(towlower(ch)));
  }
  return out;
}

bool EqualsInsensitive(const std::wstring& left, const std::wstring& right) {
  if (left.size() != right.size()) {
    return false;
  }
  for (size_t i = 0; i < left.size(); ++i) {
    if (towlower(left[i]) != towlower(right[i])) {
      return false;
    }
  }
  return true;
}

std::vector<std::wstring> SplitPath(const std::wstring& path) {
  std::vector<std::wstring> parts;
  std::wstring current;
  for (wchar_t ch : path) {
    if (ch == L'\\' || ch == L'/') {
      if (!current.empty()) {
        parts.push_back(current);
        current.clear();
      }
    } else {
      current.push_back(ch);
    }
  }
  if (!current.empty()) {
    parts.push_back(current);
  }
  return parts;
}

std::wstring JoinPathParts(const std::vector<std::wstring>& parts, size_t count) {
  if (count == 0 || parts.empty()) {
    return L"";
  }
  if (count > parts.size()) {
    count = parts.size();
  }
  std::wstring out = parts[0];
  for (size_t i = 1; i < count; ++i) {
    out.append(L"\\");
    out.append(parts[i]);
  }
  return out;
}

TraceNodeData* StoreNodeData(TraceDialogState* state, bool is_value, const std::wstring& key_path, const std::wstring& value_name) {
  auto node = std::make_unique<TraceNodeData>();
  node->is_value = is_value;
  node->key_path = key_path;
  node->value_name = value_name;
  TraceNodeData* raw = node.get();
  state->node_storage.push_back(std::move(node));
  return raw;
}

void UpdateStatus(HWND hwnd, TraceDialogState* state) {
  if (!state || !state->status) {
    return;
  }
  std::wstring text = L"Loaded ";
  text.append(std::to_wstring(state->key_count));
  text.append(L" keys");
  if (state->show_values) {
    text.append(L", ");
    text.append(std::to_wstring(state->value_count));
    text.append(L" values");
  }
  if (!state->loading_done) {
    text.append(L" (loading...)");
  }
  SetWindowTextW(state->status, text.c_str());
}

HTREEITEM EnsureKeyNode(HWND tree, TraceDialogState* state, const std::wstring& key_path, const std::wstring& display_path) {
  if (!state || !tree || key_path.empty()) {
    return nullptr;
  }
  std::wstring tree_path = display_path.empty() ? key_path : display_path;
  std::vector<std::wstring> tree_parts = SplitPath(tree_path);
  if (tree_parts.empty()) {
    return nullptr;
  }
  std::vector<std::wstring> key_parts = SplitPath(key_path);
  size_t match_suffix = 0;
  while (match_suffix < tree_parts.size() && match_suffix < key_parts.size()) {
    size_t tree_index = tree_parts.size() - 1 - match_suffix;
    size_t key_index = key_parts.size() - 1 - match_suffix;
    if (!EqualsInsensitive(tree_parts[tree_index], key_parts[key_index])) {
      break;
    }
    ++match_suffix;
  }
  size_t align_start_tree = tree_parts.size() - match_suffix;
  size_t align_start_key = key_parts.size() - match_suffix;
  auto key_prefix_for_index = [&](size_t index) -> std::wstring {
    if (key_parts.empty()) {
      return L"";
    }
    size_t part_count = 0;
    if (index < align_start_tree) {
      if (index == 0 && EqualsInsensitive(tree_parts[0], L"REGISTRY")) {
        return L"";
      }
      part_count = 1;
    } else {
      part_count = align_start_key + (index - align_start_tree) + 1;
    }
    return JoinPathParts(key_parts, part_count);
  };

  std::wstring current_tree;
  HTREEITEM parent = TVI_ROOT;
  for (size_t i = 0; i < tree_parts.size(); ++i) {
    if (!current_tree.empty()) {
      current_tree.append(L"\\");
    }
    current_tree.append(tree_parts[i]);
    std::wstring tree_lower = ToLower(current_tree);
    auto it = state->key_nodes.find(tree_lower);
    if (it != state->key_nodes.end()) {
      parent = it->second;
      continue;
    }

    std::wstring node_key_path = (i + 1 == tree_parts.size()) ? key_path : key_prefix_for_index(i);
    TraceNodeData* data = StoreNodeData(state, false, node_key_path, L"");
    TVINSERTSTRUCTW insert = {};
    insert.hParent = parent;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM;
    insert.item.pszText = const_cast<wchar_t*>(tree_parts[i].c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(data);
    HTREEITEM item = TreeView_InsertItem(tree, &insert);
    if (parent != TVI_ROOT && TreeView_GetCheckState(tree, parent) != FALSE) {
      bool prior = state->updating_checks;
      state->updating_checks = true;
      TreeView_SetCheckState(tree, item, TRUE);
      state->updating_checks = prior;
    }
    state->key_nodes.emplace(std::move(tree_lower), item);
    parent = item;
  }
  return parent == TVI_ROOT ? nullptr : parent;
}

HTREEITEM InsertValueNode(HWND tree, TraceDialogState* state, HTREEITEM key_item, const std::wstring& key_path, const std::wstring& value_name, const std::wstring& value_lower) {
  if (!state || !tree || !key_item) {
    return nullptr;
  }
  std::wstring display = value_name.empty() ? L"(Default)" : value_name;
  TraceNodeData* data = StoreNodeData(state, true, key_path, value_name);
  TVINSERTSTRUCTW insert = {};
  insert.hParent = key_item;
  insert.hInsertAfter = TVI_LAST;
  insert.item.mask = TVIF_TEXT | TVIF_PARAM;
  insert.item.pszText = const_cast<wchar_t*>(display.c_str());
  insert.item.lParam = reinterpret_cast<LPARAM>(data);
  HTREEITEM item = TreeView_InsertItem(tree, &insert);
  if (TreeView_GetCheckState(tree, key_item) != FALSE) {
    bool prior = state->updating_checks;
    state->updating_checks = true;
    TreeView_SetCheckState(tree, item, TRUE);
    state->updating_checks = prior;
  }
  return item;
}

void EnsureValueNodes(HWND tree, TraceDialogState* state, HTREEITEM key_item, const std::wstring& key_path) {
  if (!state || !tree || !key_item || key_path.empty()) {
    return;
  }
  std::wstring key_lower = ToLower(key_path);
  auto values_it = state->values_by_key.find(key_lower);
  if (values_it == state->values_by_key.end()) {
    return;
  }
  auto& loaded = state->values_loaded[key_lower];
  for (const auto& entry : values_it->second) {
    const std::wstring& value_lower = entry.first;
    const std::wstring& display = entry.second;
    if (!loaded.insert(value_lower).second) {
      continue;
    }
    std::wstring value_name = display == L"(Default)" ? L"" : display;
    InsertValueNode(tree, state, key_item, key_path, value_name, value_lower);
  }
}

void AddEntry(HWND tree, TraceDialogState* state, const KeyValueDialogEntry& entry) {
  if (!state || !tree || entry.key_path.empty()) {
    return;
  }
  std::wstring display_path = entry.display_path.empty() ? entry.key_path : entry.display_path;
  HTREEITEM key_item = EnsureKeyNode(tree, state, entry.key_path, display_path);
  std::wstring key_lower = ToLower(entry.key_path);
  if (state->key_entries.insert(key_lower).second) {
    state->key_count++;
  }
  if (!entry.has_value || !state->show_values) {
    return;
  }
  std::wstring value_name = entry.value_name;
  std::wstring display_name = value_name.empty() ? L"(Default)" : value_name;
  std::wstring value_lower = ToLower(value_name);
  auto& values = state->values_by_key[key_lower];
  if (values.emplace(value_lower, display_name).second) {
    state->value_count++;
    if (key_item) {
      UINT state_mask = TreeView_GetItemState(tree, key_item, TVIS_EXPANDED);
      if (state_mask & TVIS_EXPANDED) {
        InsertValueNode(tree, state, key_item, entry.key_path, value_name, value_lower);
      }
    }
  }
}

TraceNodeData* GetNodeData(HWND tree, HTREEITEM item) {
  if (!tree || !item) {
    return nullptr;
  }
  TVITEMW info = {};
  info.mask = TVIF_PARAM;
  info.hItem = item;
  if (!TreeView_GetItem(tree, &info)) {
    return nullptr;
  }
  return reinterpret_cast<TraceNodeData*>(info.lParam);
}

void AppendCheckedNodes(HWND tree, HTREEITEM item, TraceSelection* selection, std::unordered_set<std::wstring>* seen_keys) {
  while (item) {
    TraceNodeData* data = GetNodeData(tree, item);
    if (data && TreeView_GetCheckState(tree, item)) {
      if (data->is_value) {
        std::wstring key_lower = ToLower(data->key_path);
        std::wstring value_lower = ToLower(data->value_name);
        selection->values_by_key[key_lower].insert(value_lower);
        if (seen_keys && seen_keys->insert(key_lower).second) {
          selection->key_paths.push_back(data->key_path);
        }
      } else {
        std::wstring key_lower = ToLower(data->key_path);
        if (seen_keys && seen_keys->insert(key_lower).second) {
          selection->key_paths.push_back(data->key_path);
        }
      }
    }
    HTREEITEM child = TreeView_GetChild(tree, item);
    if (child) {
      AppendCheckedNodes(tree, child, selection, seen_keys);
    }
    item = TreeView_GetNextSibling(tree, item);
  }
}

void ApplyCheckStateToChildren(HWND tree, HTREEITEM parent, bool checked) {
  if (!tree || !parent) {
    return;
  }
  HTREEITEM child = TreeView_GetChild(tree, parent);
  while (child) {
    TreeView_SetCheckState(tree, child, checked);
    ApplyCheckStateToChildren(tree, child, checked);
    child = TreeView_GetNextSibling(tree, child);
  }
}

void QueueEntries(HWND hwnd, TraceDialogState* state, std::vector<KeyValueDialogEntry>&& entries) {
  if (!state || entries.empty()) {
    return;
  }
  state->pending_entries.reserve(state->pending_entries.size() + entries.size());
  for (auto& entry : entries) {
    state->pending_entries.push_back(std::move(entry));
  }
  if (!state->processing_entries) {
    state->processing_entries = true;
    PostMessageW(hwnd, kDialogProcessEntriesMessage, 0, 0);
  }
}

void ProcessPendingEntries(HWND hwnd, TraceDialogState* state) {
  if (!state || !state->processing_entries) {
    return;
  }
  constexpr size_t kBatchSize = 128;
  constexpr DWORD kBatchMs = 8;
  uint64_t start_tick = GetTickCount64();
  size_t processed = 0;
  while (state->pending_index < state->pending_entries.size()) {
    AddEntry(state->tree, state, state->pending_entries[state->pending_index]);
    ++state->pending_index;
    ++processed;
    if (processed >= kBatchSize) {
      break;
    }
    if (GetTickCount64() - start_tick >= kBatchMs) {
      break;
    }
  }
  UpdateStatus(hwnd, state);
  if (state->pending_index >= state->pending_entries.size()) {
    state->pending_entries.clear();
    state->pending_index = 0;
    state->processing_entries = false;
  } else {
    PostMessageW(hwnd, kDialogProcessEntriesMessage, 0, 0);
  }
}

void AcceptSelection(HWND hwnd, TraceDialogState* state, bool select_all) {
  if (!state || !state->out) {
    return;
  }
  bool recursive = SendMessageW(state->recursive, BM_GETCHECK, 0, 0) == BST_CHECKED;
  if (select_all) {
    state->out->select_all = true;
    state->out->recursive = recursive;
    state->out->key_paths.clear();
    state->out->values_by_key.clear();
    state->accepted = true;
    RestoreOwnerWindow(state->owner, &state->owner_restored);
    DestroyWindow(hwnd);
    return;
  }

  TraceSelection selection = {};
  selection.select_all = false;
  selection.recursive = recursive;
  std::unordered_set<std::wstring> seen_keys;
  HTREEITEM root = TreeView_GetRoot(state->tree);
  if (root) {
    AppendCheckedNodes(state->tree, root, &selection, &seen_keys);
  }
  if (selection.key_paths.empty() && selection.values_by_key.empty()) {
    ui::ShowWarning(hwnd, L"Select at least one key or value.");
    return;
  }
  *state->out = std::move(selection);
  state->accepted = true;
  RestoreOwnerWindow(state->owner, &state->owner_restored);
  DestroyWindow(hwnd);
}

void CenterWindowToOwner(HWND hwnd, HWND owner, int width, int height) {
  RECT rect = {};
  if (owner && GetWindowRect(owner, &rect)) {
    int owner_w = rect.right - rect.left;
    int owner_h = rect.bottom - rect.top;
    int x = rect.left + std::max(0, (owner_w - width) / 2);
    int y = rect.top + std::max(0, (owner_h - height) / 2);
    SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    return;
  }
  SetWindowPos(hwnd, nullptr, CW_USEDEFAULT, CW_USEDEFAULT, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
}

void LayoutDialog(HWND hwnd, TraceDialogState* state, HFONT font) {
  if (!state) {
    return;
  }
  RECT rect = {};
  GetClientRect(hwnd, &rect);
  int width = rect.right - rect.left;
  int height = rect.bottom - rect.top;
  int padding = 12;
  int gap = 6;
  int button_h = 22;
  int button_w = 90;
  int check_h = 22;
  int label_h = 18;

  int y = padding;
  if (state->label) {
    SetWindowPos(state->label, nullptr, padding, y, width - padding * 2, label_h, SWP_NOZORDER | SWP_NOACTIVATE);
    y += label_h + gap;
  }
  if (state->status) {
    SetWindowPos(state->status, nullptr, padding, y, width - padding * 2, label_h, SWP_NOZORDER | SWP_NOACTIVATE);
    y += label_h + gap;
  }

  int buttons_y = height - padding - button_h + 4;
  int check_y = buttons_y - check_h - gap;
  int tree_height = check_y - y - gap;
  if (tree_height < 80) {
    tree_height = 80;
  }
  if (state->tree) {
    SetWindowPos(state->tree, nullptr, padding, y, width - padding * 2, tree_height, SWP_NOZORDER | SWP_NOACTIVATE);
  }

  int select_all_w = 140;
  int recursive_w = 160;
  int select_x = padding;
  int recursive_x = select_x + select_all_w + gap;
  if (state->select_all) {
    SetWindowPos(state->select_all, nullptr, select_x, check_y, select_all_w, check_h, SWP_NOZORDER | SWP_NOACTIVATE);
  }
  if (state->recursive) {
    SetWindowPos(state->recursive, nullptr, recursive_x, check_y, recursive_w, check_h, SWP_NOZORDER | SWP_NOACTIVATE);
  }

  int button_right_margin = 8;
  int cancel_x = width - button_right_margin - button_w;
  int ok_x = cancel_x - button_w - gap;
  if (state->ok_button) {
    SetWindowPos(state->ok_button, nullptr, ok_x, buttons_y, button_w, button_h, SWP_NOZORDER | SWP_NOACTIVATE);
  }
  if (state->cancel_button) {
    SetWindowPos(state->cancel_button, nullptr, cancel_x, buttons_y, button_w, button_h, SWP_NOZORDER | SWP_NOACTIVATE);
  }

  if (font) {
    ApplyFont(hwnd, font);
  }
}

LRESULT CALLBACK TraceDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<TraceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<TraceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    state->font = CreateDialogFont();
    HFONT font = state->font;

    if (!state->prompt.empty()) {
      state->label = CreateWindowExW(0, L"STATIC", state->prompt.c_str(), WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kTraceLabel), nullptr, nullptr);
    }
    state->status = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kTraceStatus), nullptr, nullptr);
    state->tree = CreateWindowExW(0, WC_TREEVIEWW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS | TVS_CHECKBOXES, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kTraceTree), nullptr, nullptr);
    state->recursive = CreateWindowExW(0, L"BUTTON", L"Recursive", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kRecursiveCheck), nullptr, nullptr);
    state->select_all = CreateWindowExW(0, L"BUTTON", L"Select All Keys", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kSelectAllButton), nullptr, nullptr);
    state->ok_button = CreateWindowExW(0, L"BUTTON", L"Select", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOkButton), nullptr, nullptr);
    state->cancel_button = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kCancelButton), nullptr, nullptr);

    ApplyFont(hwnd, font);
    EnumChildWindows(
        hwnd,
        [](HWND child, LPARAM param) -> BOOL {
          HFONT font_handle = reinterpret_cast<HFONT>(param);
          ApplyFont(child, font_handle);
          return TRUE;
        },
        reinterpret_cast<LPARAM>(font));

    SendMessageW(state->recursive, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(state->tree, TVM_SETEXTENDEDSTYLE, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);

    Theme::Current().ApplyToTreeView(state->tree);
    Theme::Current().ApplyToChildren(hwnd);

    UpdateStatus(hwnd, state);
    LayoutDialog(hwnd, state, font);

    if (state->on_ready) {
      state->on_ready(hwnd, state->on_ready_context);
    }
    return 0;
  }
  case WM_DESTROY:
    if (state && state->font) {
      DeleteObject(state->font);
      state->font = nullptr;
    }
    return 0;
  case WM_SIZE: {
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    LayoutDialog(hwnd, state, font);
    return 0;
  }
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(hwnd);
      if (state && state->tree) {
        Theme::Current().ApplyToTreeView(state->tree);
      }
      Theme::Current().ApplyToChildren(hwnd);
      InvalidateRect(hwnd, nullptr, TRUE);
    }
    return 0;
  }
  case WM_CTLCOLORSTATIC: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_STATIC));
  }
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_BTN));
  }
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_EDIT));
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return 1;
  }
  case WM_COMMAND: {
    if (!state) {
      return 0;
    }
    if (HIWORD(wparam) == BN_CLICKED) {
      switch (LOWORD(wparam)) {
      case kSelectAllButton:
        AcceptSelection(hwnd, state, true);
        return 0;
      case kOkButton:
        AcceptSelection(hwnd, state, false);
        return 0;
      case kCancelButton:
        RestoreOwnerWindow(state->owner, &state->owner_restored);
        DestroyWindow(hwnd);
        return 0;
      default:
        break;
      }
    }
    break;
  }
  case WM_NOTIFY: {
    if (!state) {
      return 0;
    }
    auto* header = reinterpret_cast<NMHDR*>(lparam);
    if (header && header->code == TVN_ITEMEXPANDINGW) {
      auto* info = reinterpret_cast<NMTREEVIEWW*>(lparam);
      if (info->action == TVE_EXPAND) {
        TraceNodeData* data = GetNodeData(state->tree, info->itemNew.hItem);
        if (data && !data->is_value) {
          EnsureValueNodes(state->tree, state, info->itemNew.hItem, data->key_path);
        }
      }
    }
    if (header && header->code == TVN_ITEMCHANGEDW) {
      auto* change = reinterpret_cast<NMTVITEMCHANGE*>(lparam);
      if (change && (change->uChanged & TVIF_STATE) && ((change->uStateNew ^ change->uStateOld) & TVIS_STATEIMAGEMASK)) {
        if (!state->updating_checks) {
          TraceNodeData* data = GetNodeData(state->tree, change->hItem);
          if (data && !data->is_value) {
            bool checked = TreeView_GetCheckState(state->tree, change->hItem) != FALSE;
            state->updating_checks = true;
            ApplyCheckStateToChildren(state->tree, change->hItem, checked);
            state->updating_checks = false;
          }
        }
      }
    }
    break;
  }
  case kDialogAddEntriesMessage: {
    if (!state) {
      return 0;
    }
    auto* entries = reinterpret_cast<std::vector<KeyValueDialogEntry>*>(lparam);
    std::unique_ptr<std::vector<KeyValueDialogEntry>> owned(entries);
    if (owned) {
      QueueEntries(hwnd, state, std::move(*owned));
    }
    return 0;
  }
  case kDialogDoneMessage:
    if (state) {
      state->loading_done = (wparam != 0);
      UpdateStatus(hwnd, state);
    }
    return 0;
  case kDialogProcessEntriesMessage:
    if (state) {
      ProcessPendingEntries(hwnd, state);
    }
    return 0;
  case WM_CLOSE:
    if (state) {
      RestoreOwnerWindow(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

HWND CreateTraceDialogWindow(HINSTANCE instance, const std::wstring& title, HWND owner, TraceDialogState* state) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = TraceDialogProc;
  wc.hInstance = instance;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kDialogClass;
  RegisterClassW(&wc);

  const wchar_t* window_title = title.empty() ? kAppTitle : title.c_str();
  return CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kDialogClass, window_title, WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 560, 552, owner, nullptr, instance, state);
}

} // namespace

bool ShowTraceDialog(HWND owner, const TraceDialogOptions& options, TraceSelection* selection, TraceDialogReadyCallback on_ready, void* context) {
  if (!selection) {
    return false;
  }
  HINSTANCE instance = GetModuleHandleW(nullptr);
  TraceDialogState state;
  state.out = selection;
  state.owner = owner;
  state.show_values = options.show_values;
  state.on_ready = on_ready;
  state.on_ready_context = context;
  state.prompt = options.prompt;

  HWND hwnd = CreateTraceDialogWindow(instance, options.title, owner, &state);
  if (!hwnd) {
    return false;
  }

  Theme::Current().ApplyToWindow(hwnd);
  RECT rect = {};
  if (GetWindowRect(hwnd, &rect)) {
    CenterWindowToOwner(hwnd, owner, rect.right - rect.left, rect.bottom - rect.top);
  }

  EnableWindow(owner, FALSE);
  ShowWindow(hwnd, SW_SHOW);
  UpdateWindow(hwnd);

  MSG msg = {};
  while (IsWindow(hwnd) && GetMessageW(&msg, nullptr, 0, 0)) {
    if (!IsDialogMessageW(hwnd, &msg)) {
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  }

  RestoreOwnerWindow(owner, &state.owner_restored);
  return state.accepted;
}

void TraceDialogPostEntries(HWND dialog, std::vector<KeyValueDialogEntry>* entries) {
  if (!dialog || !entries) {
    delete entries;
    return;
  }
  if (!PostMessageW(dialog, kDialogAddEntriesMessage, 0, reinterpret_cast<LPARAM>(entries))) {
    delete entries;
  }
}

void TraceDialogPostDone(HWND dialog, bool done) {
  if (!dialog) {
    return;
  }
  PostMessageW(dialog, kDialogDoneMessage, done ? 1 : 0, 0);
}

} // namespace regkit
