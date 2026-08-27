// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "browse/browse_pane.h"

#include "registry/registry_path.h"

#include <algorithm>

namespace regkit::browse {

namespace {

constexpr DWORD kTypeSelectTimeoutMs = 1000;

bool EqualsInsensitive(const std::wstring& left,
                       const std::wstring& right) {
  return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

bool StartsWithInsensitive(const std::wstring& text,
                           const std::wstring& prefix) {
  return prefix.size() <= text.size() &&
         _wcsnicmp(text.c_str(), prefix.c_str(), prefix.size()) == 0;
}

int CompareInsensitive(const std::wstring& left,
                       const std::wstring& right) {
  return _wcsicmp(left.c_str(), right.c_str());
}

} // namespace

bool Pane::Create(const CreateRequest& request) {
  if (!request.parent || !request.instance) {
    return false;
  }
  address_ = CreateWindowExW(
      0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_MULTILINE,
      0, 0, 0, 0, request.parent,
      reinterpret_cast<HMENU>(static_cast<INT_PTR>(request.address_id)),
      request.instance, nullptr);
  go_button_ = CreateWindowExW(
      0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, 0, 0, 0,
      0, request.parent,
      reinterpret_cast<HMENU>(static_cast<INT_PTR>(request.go_id)),
      request.instance, nullptr);
  filter_ = CreateWindowExW(
      0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_MULTILINE,
      0, 0, 0, 0, request.parent,
      reinterpret_cast<HMENU>(static_cast<INT_PTR>(request.filter_id)),
      request.instance, nullptr);
  tree_.Create(request.parent, request.instance, request.tree_id, false);
  values_.Create(request.parent, request.instance, request.values_id);
  if (!address_ || !go_button_ || !filter_ || !tree_.hwnd() ||
      !values_.hwnd()) {
    return false;
  }

  if (request.address_proc) {
    SetWindowSubclass(address_, request.address_proc,
                      request.address_subclass_id,
                      request.callback_context);
  }
  if (request.filter_proc) {
    SetWindowSubclass(filter_, request.filter_proc,
                      request.filter_subclass_id,
                      request.callback_context);
  }
  if (request.tree_proc) {
    SetWindowSubclass(tree_.hwnd(), request.tree_proc,
                      request.tree_subclass_id,
                      request.callback_context);
  }
  if (request.values_proc) {
    SetWindowSubclass(values_.hwnd(), request.values_proc,
                      request.values_subclass_id,
                      request.callback_context);
  }
  SendMessageW(address_, EM_SETCUEBANNER, TRUE,
               reinterpret_cast<LPARAM>(L"Registry path"));
  return true;
}

HWND Pane::address() const noexcept { return address_; }
HWND Pane::go_button() const noexcept { return go_button_; }
HWND Pane::filter() const noexcept { return filter_; }
RegistryTree& Pane::tree() noexcept { return tree_; }
const RegistryTree& Pane::tree() const noexcept { return tree_; }
ValueList& Pane::values() noexcept { return values_; }
const ValueList& Pane::values() const noexcept { return values_; }
RegistryNode* Pane::current_node() const noexcept { return current_node_; }
void Pane::set_current_node(RegistryNode* node) noexcept {
  current_node_ = node;
}
std::vector<RegistryRootEntry>& Pane::roots() noexcept { return roots_; }
const std::vector<RegistryRootEntry>& Pane::roots() const noexcept {
  return roots_;
}
ColumnState& Pane::columns() noexcept { return columns_; }
const ColumnState& Pane::columns() const noexcept { return columns_; }

void Pane::AddAddressHistory(const std::wstring& path) {
  if (path.empty()) {
    return;
  }
  const auto existing =
      std::find(address_history_.begin(), address_history_.end(), path);
  if (existing != address_history_.end()) {
    address_history_.erase(existing);
  }
  address_history_.insert(address_history_.begin(), path);
  constexpr size_t kMaximumAddresses = 20;
  if (address_history_.size() > kMaximumAddresses) {
    address_history_.resize(kMaximumAddresses);
  }
}

const std::vector<std::wstring>& Pane::address_history() const noexcept {
  return address_history_;
}

bool Pane::RecordNavigation(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  if (programmatic_navigation_) {
    programmatic_navigation_ = false;
    return false;
  }
  if (navigation_index_ >= 0 &&
      navigation_index_ < static_cast<int>(navigation_history_.size()) &&
      navigation_history_[static_cast<size_t>(navigation_index_)] == path) {
    return false;
  }
  if (navigation_index_ + 1 <
      static_cast<int>(navigation_history_.size())) {
    navigation_history_.erase(
        navigation_history_.begin() + navigation_index_ + 1,
        navigation_history_.end());
  }
  navigation_history_.push_back(path);
  navigation_index_ = static_cast<int>(navigation_history_.size()) - 1;
  return true;
}

std::optional<std::wstring> Pane::Back() {
  if (navigation_index_ <= 0) {
    return std::nullopt;
  }
  --navigation_index_;
  programmatic_navigation_ = true;
  return navigation_history_[static_cast<size_t>(navigation_index_)];
}

std::optional<std::wstring> Pane::Forward() {
  if (navigation_index_ + 1 >=
      static_cast<int>(navigation_history_.size())) {
    return std::nullopt;
  }
  ++navigation_index_;
  programmatic_navigation_ = true;
  return navigation_history_[static_cast<size_t>(navigation_index_)];
}

std::optional<std::wstring> Pane::Up() {
  if (!current_node_ || current_node_->subkey.empty()) {
    return std::nullopt;
  }
  std::wstring path = registry_path::Build(*current_node_);
  const size_t separator = path.rfind(L'\\');
  if (separator == std::wstring::npos) {
    return std::nullopt;
  }
  programmatic_navigation_ = true;
  return path.substr(0, separator);
}

NavigationAvailability Pane::navigation() const noexcept {
  NavigationAvailability available;
  available.back = navigation_index_ > 0;
  available.forward = navigation_index_ + 1 <
                      static_cast<int>(navigation_history_.size());
  available.up = current_node_ && !current_node_->subkey.empty();
  return available;
}

void Pane::ResetNavigation() {
  navigation_history_.clear();
  navigation_index_ = -1;
  programmatic_navigation_ = false;
}

bool Pane::SelectValue(const std::wstring& name) {
  if (!values_.hwnd()) {
    return false;
  }
  for (size_t index = 0; index < values_.RowCount(); ++index) {
    const ListRow* row = values_.RowAt(static_cast<int>(index));
    if (!row || row->kind != rowkind::kValue || row->extra != name) {
      continue;
    }
    ListView_SetItemState(values_.hwnd(), -1, 0,
                          LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetItemState(values_.hwnd(), static_cast<int>(index),
                          LVIS_SELECTED | LVIS_FOCUSED,
                          LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(values_.hwnd(), static_cast<int>(index), FALSE);
    return true;
  }
  return false;
}

void Pane::UpdateTypeBuffer(wchar_t ch, DWORD now, std::wstring* buffer,
                            DWORD* tick) {
  if (!buffer || !tick) {
    return;
  }
  if (now - *tick > kTypeSelectTimeoutMs) {
    buffer->clear();
  }
  *tick = now;
  if (ch == L'\b') {
    if (!buffer->empty()) {
      buffer->pop_back();
    }
  } else {
    buffer->push_back(ch);
  }
}

void Pane::TypeSelectValues(wchar_t ch, DWORD now) {
  if (!values_.hwnd()) {
    return;
  }
  UpdateTypeBuffer(ch, now, &value_type_buffer_, &value_type_tick_);
  if (value_type_buffer_.empty() || values_.RowCount() == 0) {
    return;
  }

  int match = -1;
  for (size_t index = 0; index < values_.RowCount(); ++index) {
    const ListRow* row = values_.RowAt(static_cast<int>(index));
    if (row && StartsWithInsensitive(row->name, value_type_buffer_)) {
      match = static_cast<int>(index);
      break;
    }
  }
  if (match < 0) {
    std::wstring nearest;
    for (size_t index = 0; index < values_.RowCount(); ++index) {
      const ListRow* row = values_.RowAt(static_cast<int>(index));
      if (row && CompareInsensitive(row->name, value_type_buffer_) >= 0 &&
          (match < 0 || CompareInsensitive(row->name, nearest) < 0)) {
        match = static_cast<int>(index);
        nearest = row->name;
      }
    }
  }
  if (match < 0) {
    match = static_cast<int>(values_.RowCount() - 1);
  }
  ListView_SetItemState(values_.hwnd(), -1, 0,
                        LVIS_SELECTED | LVIS_FOCUSED);
  ListView_SetItemState(values_.hwnd(), match,
                        LVIS_SELECTED | LVIS_FOCUSED,
                        LVIS_SELECTED | LVIS_FOCUSED);
  ListView_EnsureVisible(values_.hwnd(), match, FALSE);
}

void Pane::TypeSelectTree(wchar_t ch, DWORD now) {
  if (!tree_.hwnd()) {
    return;
  }
  UpdateTypeBuffer(ch, now, &tree_type_buffer_, &tree_type_tick_);
  if (tree_type_buffer_.empty()) {
    return;
  }
  const HTREEITEM selected = TreeView_GetSelection(tree_.hwnd());
  if (!selected) {
    return;
  }

  auto load_children = [&](HTREEITEM item) {
    RegistryNode* node = tree_.NodeFromItem(item);
    if (!node || node->children_loaded) {
      return;
    }
    NMTREEVIEWW info = {};
    info.action = TVE_EXPAND;
    info.itemNew.hItem = item;
    tree_.OnItemExpanding(&info);
  };
  auto children = [&](HTREEITEM parent) {
    std::vector<HTREEITEM> items;
    for (HTREEITEM item = TreeView_GetChild(tree_.hwnd(), parent); item;
         item = TreeView_GetNextSibling(tree_.hwnd(), item)) {
      items.push_back(item);
    }
    return items;
  };
  auto text = [&](HTREEITEM item) {
    wchar_t buffer[256] = {};
    TVITEMW value = {};
    value.hItem = item;
    value.mask = TVIF_TEXT;
    value.pszText = buffer;
    value.cchTextMax = static_cast<int>(_countof(buffer));
    return TreeView_GetItem(tree_.hwnd(), &value)
               ? std::wstring(buffer)
               : std::wstring();
  };
  auto find = [&](const std::vector<HTREEITEM>& items) -> HTREEITEM {
    if (items.empty()) {
      return nullptr;
    }
    size_t start = 0;
    const auto selected_item = std::find(items.begin(), items.end(), selected);
    if (selected_item != items.end()) {
      start = static_cast<size_t>(selected_item - items.begin());
    }
    for (int exact = 1; exact >= 0; --exact) {
      for (size_t offset = 0; offset < items.size(); ++offset) {
        const HTREEITEM item = items[(start + offset) % items.size()];
        const std::wstring item_text = text(item);
        const bool matched = exact
                                 ? EqualsInsensitive(item_text,
                                                     tree_type_buffer_)
                                 : StartsWithInsensitive(item_text,
                                                         tree_type_buffer_);
        if (matched) {
          return item;
        }
      }
    }
    return nullptr;
  };

  HTREEITEM target = nullptr;
  if (tree_type_select_descend_) {
    load_children(selected);
    target = find(children(selected));
  }
  const HTREEITEM parent = TreeView_GetParent(tree_.hwnd(), selected);
  if (!target && parent) {
    load_children(parent);
    target = find(children(parent));
  }
  if (!target && !tree_type_select_descend_) {
    load_children(selected);
    target = find(children(selected));
  }
  if (target) {
    tree_type_select_descend_ = false;
    TreeView_SelectItem(tree_.hwnd(), target);
    TreeView_EnsureVisible(tree_.hwnd(), target);
  }
}

void Pane::set_tree_type_select_descend(bool descend) noexcept {
  tree_type_buffer_.clear();
  tree_type_select_descend_ = descend;
}

SourceKind Pane::source_kind() const noexcept { return source_kind_; }
void Pane::set_source_kind(SourceKind kind) noexcept { source_kind_ = kind; }

} // namespace regkit::browse
