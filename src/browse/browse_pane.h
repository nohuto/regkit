// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/command_ids.h"
#include "browse/key_tree.h"
#include "browse/value_table.h"

#include <commctrl.h>
#include <windows.h>

#include <optional>
#include <string>
#include <vector>

namespace regkit::browse {

enum class SourceKind {
  kLocal,
  kRemote,
  kOffline,
  kVirtual,
};

struct CreateRequest {
  HWND parent = nullptr;
  HINSTANCE instance = nullptr;
  int address_id = 0;
  int go_id = 0;
  int filter_id = 0;
  int tree_id = 0;
  int values_id = 0;
  SUBCLASSPROC address_proc = nullptr;
  UINT_PTR address_subclass_id = 0;
  SUBCLASSPROC filter_proc = nullptr;
  UINT_PTR filter_subclass_id = 0;
  SUBCLASSPROC tree_proc = nullptr;
  UINT_PTR tree_subclass_id = 0;
  SUBCLASSPROC values_proc = nullptr;
  UINT_PTR values_subclass_id = 0;
  DWORD_PTR callback_context = 0;
};

struct ColumnState {
  std::vector<ColumnInfo> items;
  std::vector<int> widths;
  std::vector<bool> visible;
  std::vector<int> saved_widths;
  std::vector<bool> saved_visible;
  bool saved = false;
  int sort_column = 0;
  bool sort_ascending = true;
};

struct NavigationAvailability {
  bool back = false;
  bool forward = false;
  bool up = false;
};

class Pane {
public:
  bool Create(const CreateRequest& request);

  HWND address() const noexcept;
  HWND go_button() const noexcept;
  HWND filter() const noexcept;
  RegistryTree& tree() noexcept;
  const RegistryTree& tree() const noexcept;
  ValueList& values() noexcept;
  const ValueList& values() const noexcept;

  RegistryNode* current_node() const noexcept;
  void set_current_node(RegistryNode* node) noexcept;
  std::vector<RegistryRootEntry>& roots() noexcept;
  const std::vector<RegistryRootEntry>& roots() const noexcept;
  ColumnState& columns() noexcept;
  const ColumnState& columns() const noexcept;

  void AddAddressHistory(const std::wstring& path);
  const std::vector<std::wstring>& address_history() const noexcept;

  bool RecordNavigation(const std::wstring& path);
  std::optional<std::wstring> Back();
  std::optional<std::wstring> Forward();
  std::optional<std::wstring> Up();
  NavigationAvailability navigation() const noexcept;
  void ResetNavigation();

  bool SelectValue(const std::wstring& name);
  void TypeSelectValues(wchar_t ch, DWORD now);
  void TypeSelectTree(wchar_t ch, DWORD now);
  void set_tree_type_select_descend(bool descend) noexcept;

  SourceKind source_kind() const noexcept;
  void set_source_kind(SourceKind kind) noexcept;

private:
  static void UpdateTypeBuffer(wchar_t ch, DWORD now,
                               std::wstring* buffer, DWORD* tick);

  HWND address_ = nullptr;
  HWND go_button_ = nullptr;
  HWND filter_ = nullptr;
  RegistryTree tree_;
  ValueList values_;
  RegistryNode* current_node_ = nullptr;
  std::vector<RegistryRootEntry> roots_;
  ColumnState columns_;
  std::vector<std::wstring> address_history_;
  std::vector<std::wstring> navigation_history_;
  int navigation_index_ = -1;
  bool programmatic_navigation_ = false;
  std::wstring tree_type_buffer_;
  std::wstring value_type_buffer_;
  DWORD tree_type_tick_ = 0;
  DWORD value_type_tick_ = 0;
  bool tree_type_select_descend_ = false;
  SourceKind source_kind_ = SourceKind::kLocal;
};

} // namespace regkit::browse
