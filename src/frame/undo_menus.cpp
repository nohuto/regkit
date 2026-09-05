// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::PushUndo(changes::UndoOperation operation) {
  if (is_replaying_) {
    return;
  }
  undo_stack_.Push(std::move(operation));
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                 undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, 0);
  }
}

void MainWindow::Impl::ClearRedo() {
  undo_stack_.ClearRedo();
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                 undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, 0);
  }
}

bool MainWindow::Impl::ApplyUndoOperation(const changes::UndoOperation& operation,
                                          bool redo) {
  if (!browse_.current_node()) {
    return false;
  }
  bool ok = false;
  is_replaying_ = true;
  switch (operation.type) {
  case changes::UndoOperation::Type::kCreateKey: {
    if (redo) {
      if (!operation.key_snapshot.name.empty()) {
        ok = changes::RestoreKey(operation.node, operation.key_snapshot);
      } else {
        ok = RegistryStore::CreateKey(operation.node, operation.name);
      }
      if (ok) {
        RefreshTreeSelection();
      }
    } else {
      RegistryNode child = MakeChildNode(operation.node, operation.name);
      ok = RegistryStore::DeleteKey(child);
      if (ok) {
        RefreshTreeSelection();
      }
    }
    break;
  }
  case changes::UndoOperation::Type::kDeleteKey: {
    if (redo) {
      RegistryNode child = MakeChildNode(operation.node, operation.name);
      ok = RegistryStore::DeleteKey(child);
      if (ok) {
        RefreshTreeSelection();
      }
    } else {
      ok = changes::RestoreKey(operation.node, operation.key_snapshot);
      if (ok) {
        RefreshTreeSelection();
      }
    }
    break;
  }
  case changes::UndoOperation::Type::kRenameKey: {
    std::wstring from = redo ? operation.name : operation.new_name;
    std::wstring to = redo ? operation.new_name : operation.name;
    RegistryNode child = MakeChildNode(operation.node, from);
    ok = RegistryStore::RenameKey(child, to);
    if (ok) {
      RefreshTreeSelection();
      std::wstring path = registry_path::Build(operation.node);
      if (!path.empty()) {
        path.append(L"\\");
        path.append(to);
        SelectTreePath(path);
      }
    }
    break;
  }
  case changes::UndoOperation::Type::kCreateValue: {
    if (redo) {
      ok = RegistryStore::SetValue(operation.node, operation.new_value.name, operation.new_value.type, operation.new_value.data);
    } else {
      ok = RegistryStore::DeleteValue(operation.node, operation.name);
    }
    break;
  }
  case changes::UndoOperation::Type::kDeleteValue: {
    if (redo) {
      ok = RegistryStore::DeleteValue(operation.node, operation.old_value.name);
    } else {
      ok = RegistryStore::SetValue(operation.node, operation.old_value.name, operation.old_value.type, operation.old_value.data);
    }
    break;
  }
  case changes::UndoOperation::Type::kModifyValue: {
    const ValueEntry& value = redo ? operation.new_value : operation.old_value;
    ok = RegistryStore::SetValue(operation.node, value.name, value.type, value.data);
    break;
  }
  case changes::UndoOperation::Type::kRenameValue: {
    std::wstring from = redo ? operation.name : operation.new_name;
    std::wstring to = redo ? operation.new_name : operation.name;
    ok = RegistryStore::RenameValue(operation.node, from, to);
    break;
  }
  default:
    break;
  }
  is_replaying_ = false;

  if (ok) {
    MarkOfflineDirty();
  }
  if (ok && browse_.current_node()) {
    UpdateValueListForNode(browse_.current_node());
  }
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                 undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo,
                 undo_stack_.CanRedo() ? TBSTATE_ENABLED : 0);
  }
  return ok;
}

bool MainWindow::Impl::SameNode(const RegistryNode& left, const RegistryNode& right) const {
  if (left.root != right.root) {
    return false;
  }
  if (!EqualsInsensitive(left.subkey, right.subkey)) {
    return false;
  }
  return EqualsInsensitive(left.root_name, right.root_name);
}

std::wstring MainWindow::Impl::MakeUniqueValueName(const RegistryNode& node, const std::wstring& base) const {
  std::unordered_set<std::wstring> value_names;
  RegistryStore::KeyEnumResult enum_result;
  bool names_reserved = false;
  RegistryStore::EnumKeyStreaming(
      node, true, false, false, &enum_result,
      [&](const ValueInfo& value, const BYTE*, DWORD) {
        if (!names_reserved) {
          if (enum_result.info_valid) {
            value_names.reserve(enum_result.info.value_count);
          }
          names_reserved = true;
        }
        value_names.insert(ToLower(value.name));
        return true;
      },
      {});
  auto exists = [&](const std::wstring& candidate) -> bool {
    return value_names.contains(ToLower(candidate));
  };

  std::wstring base_name = base;
  if (base_name.empty()) {
    if (!exists(base_name)) {
      return base_name;
    }
    base_name = L"Default";
  }
  if (!exists(base_name)) {
    return base_name;
  }
  for (int i = 2; i < 10000; ++i) {
    std::wstring next = base_name + L" (" + std::to_wstring(i) + L")";
    if (!exists(next)) {
      return next;
    }
  }
  return base_name;
}

std::wstring MainWindow::Impl::MakeUniqueKeyName(const RegistryNode& node, const std::wstring& base) const {
  auto keys = RegistryStore::EnumSubKeyNames(node, false);
  auto exists = [&](const std::wstring& candidate) -> bool {
    for (const auto& key : keys) {
      if (EqualsInsensitive(key, candidate)) {
        return true;
      }
    }
    return false;
  };

  std::wstring base_name = base;
  if (base_name.empty()) {
    base_name = L"New Key";
  }
  if (!exists(base_name)) {
    return base_name;
  }
  for (int i = 2; i < 10000; ++i) {
    std::wstring next = base_name + L" (" + std::to_wstring(i) + L")";
    if (!exists(next)) {
      return next;
    }
  }
  return base_name;
}

bool MainWindow::Impl::ResolvePathToNode(const std::wstring& path, RegistryNode* node) const {
  if (!node || path.empty()) {
    return false;
  }
  for (const auto& root_entry : browse_.roots()) {
    if (!StartsWithInsensitive(path, root_entry.path_name)) {
      continue;
    }
    std::wstring rest = path.substr(root_entry.path_name.size());
    if (!rest.empty() && (rest.front() == L'\\' || rest.front() == L'/')) {
      rest.erase(rest.begin());
    }
    if (root_entry.subkey_prefix.empty()) {
      node->root = root_entry.root;
      node->root_name = root_entry.path_name;
      node->subkey = rest;
      return true;
    }
    std::wstring prefix = root_entry.subkey_prefix;
    if (!rest.empty()) {
      if (!StartsWithInsensitive(rest, prefix)) {
        rest = prefix + L"\\" + rest;
      }
    } else {
      rest = prefix;
    }
    node->root = root_entry.root;
    node->root_name = root_entry.path_name;
    node->subkey = rest;
    return true;
  }
  return false;
}

void MainWindow::Impl::ShowValueHeaderMenu(POINT screen_pt) {
  HWND header_hwnd = ListView_GetHeader(browse_.values().hwnd());
  if (!header_hwnd) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(header_hwnd, &client_pt);
  HDHITTESTINFO hit = {};
  hit.pt = client_pt;
  int column_hit = static_cast<int>(SendMessageW(header_hwnd, HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hit)));
  last_header_column_ = (column_hit >= 0) ? column_hit : -1;

  HMENU menu = CreatePopupMenu();
  UINT fit_flags = MF_STRING | ((last_header_column_ >= 0) ? 0 : MF_GRAYED);
  AppendMenuW(menu, fit_flags, cmd::kHeaderSizeToFit, L"Size column to fit");
  AppendMenuW(menu, MF_STRING, cmd::kHeaderSizeAll, L"Size all columns to fit");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

  for (size_t i = 0; i < browse_.columns().items.size(); ++i) {
    UINT state = browse_.columns().visible[i] ? MF_CHECKED : MF_UNCHECKED;
    AppendMenuW(menu, MF_STRING | state, cmd::kHeaderToggleBase + static_cast<int>(i), browse_.columns().items[i].title.c_str());
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command == cmd::kHeaderSizeToFit && last_header_column_ >= 0) {
    int subitem = GetListViewColumnSubItem(browse_.values().hwnd(), last_header_column_);
    ListView_SetColumnWidth(browse_.values().hwnd(), last_header_column_, LVSCW_AUTOSIZE_USEHEADER);
    int width = ListView_GetColumnWidth(browse_.values().hwnd(), last_header_column_);
    if (subitem >= 0 && static_cast<size_t>(subitem) < browse_.columns().widths.size()) {
      browse_.columns().widths[static_cast<size_t>(subitem)] = width;
    }
    SaveSettings();
    return;
  }
  if (command == cmd::kHeaderSizeAll) {
    int last_visible = FindLastVisibleColumn(browse_.columns().visible);
    for (size_t i = 0; i < browse_.columns().items.size(); ++i) {
      if (i < browse_.columns().visible.size() && !browse_.columns().visible[i]) {
        continue;
      }
      int display_index = FindListViewColumnBySubItem(browse_.values().hwnd(), static_cast<int>(i));
      if (display_index < 0) {
        continue;
      }
      int width = 0;
      if (static_cast<int>(i) == last_visible) {
        width = CalcListViewColumnFitWidth(browse_.values().hwnd(), static_cast<int>(i), browse_.columns().items[i].width);
        ListView_SetColumnWidth(browse_.values().hwnd(), display_index, width);
      } else {
        ListView_SetColumnWidth(browse_.values().hwnd(), display_index, LVSCW_AUTOSIZE_USEHEADER);
        width = ListView_GetColumnWidth(browse_.values().hwnd(), display_index);
      }
      browse_.columns().widths[i] = width;
    }
    SaveSettings();
    return;
  }
  if (command >= cmd::kHeaderToggleBase) {
    int index = command - cmd::kHeaderToggleBase;
    if (index >= 0 && static_cast<size_t>(index) < browse_.columns().items.size()) {
      ToggleValueColumn(index, !browse_.columns().visible[static_cast<size_t>(index)]);
      SaveSettings();
    }
  }
}

void MainWindow::Impl::ShowHistoryHeaderMenu(POINT screen_pt) {
  HWND header_hwnd = ListView_GetHeader(history_list_);
  if (!header_hwnd) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(header_hwnd, &client_pt);
  HDHITTESTINFO hit = {};
  hit.pt = client_pt;
  int column_hit = static_cast<int>(SendMessageW(header_hwnd, HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hit)));

  HMENU menu = CreatePopupMenu();
  UINT fit_flags = MF_STRING | ((column_hit >= 0) ? 0 : MF_GRAYED);
  AppendMenuW(menu, fit_flags, cmd::kHeaderSizeToFit, L"Size column to fit");
  AppendMenuW(menu, MF_STRING, cmd::kHeaderSizeAll, L"Size all columns to fit");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

  for (size_t i = 0; i < history_columns_.size(); ++i) {
    UINT state = history_column_visible_[i] ? MF_CHECKED : MF_UNCHECKED;
    AppendMenuW(menu, MF_STRING | state, cmd::kHeaderToggleBase + static_cast<int>(i), history_columns_[i].title.c_str());
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command == cmd::kHeaderSizeToFit && column_hit >= 0) {
    int subitem = GetListViewColumnSubItem(history_list_, column_hit);
    ListView_SetColumnWidth(history_list_, column_hit, LVSCW_AUTOSIZE_USEHEADER);
    if (subitem >= 0 && static_cast<size_t>(subitem) < history_column_widths_.size()) {
      history_column_widths_[static_cast<size_t>(subitem)] = ListView_GetColumnWidth(history_list_, column_hit);
    }
    return;
  }
  if (command == cmd::kHeaderSizeAll) {
    int last_visible = FindLastVisibleColumn(history_column_visible_);
    for (size_t i = 0; i < history_columns_.size(); ++i) {
      if (i < history_column_visible_.size() && !history_column_visible_[i]) {
        continue;
      }
      int display_index = FindListViewColumnBySubItem(history_list_, static_cast<int>(i));
      if (display_index < 0) {
        continue;
      }
      int width = 0;
      if (static_cast<int>(i) == last_visible) {
        width = CalcListViewColumnFitWidth(history_list_, static_cast<int>(i), history_columns_[i].width);
        ListView_SetColumnWidth(history_list_, display_index, width);
      } else {
        ListView_SetColumnWidth(history_list_, display_index, LVSCW_AUTOSIZE_USEHEADER);
        width = ListView_GetColumnWidth(history_list_, display_index);
      }
      history_column_widths_[i] = width;
    }
    return;
  }
  if (command >= cmd::kHeaderToggleBase) {
    int index = command - cmd::kHeaderToggleBase;
    if (index >= 0 && static_cast<size_t>(index) < history_columns_.size()) {
      ToggleHistoryColumn(index, !history_column_visible_[static_cast<size_t>(index)]);
    }
  }
}

void MainWindow::Impl::ShowSearchHeaderMenu(POINT screen_pt) {
  HWND header_hwnd = ListView_GetHeader(search_results_list_);
  if (!header_hwnd) {
    return;
  }
  bool compare = IsCompareTabSelected();
  auto& columns = compare ? compare_columns_ : search_columns_;
  auto& widths = compare ? compare_column_widths_ : search_column_widths_;
  auto& visible = compare ? compare_column_visible_ : search_column_visible_;
  POINT client_pt = screen_pt;
  ScreenToClient(header_hwnd, &client_pt);
  HDHITTESTINFO hit = {};
  hit.pt = client_pt;
  int column_hit = static_cast<int>(SendMessageW(header_hwnd, HDM_HITTEST, 0, reinterpret_cast<LPARAM>(&hit)));

  HMENU menu = CreatePopupMenu();
  UINT fit_flags = MF_STRING | ((column_hit >= 0) ? 0 : MF_GRAYED);
  AppendMenuW(menu, fit_flags, cmd::kHeaderSizeToFit, L"Size column to fit");
  AppendMenuW(menu, MF_STRING, cmd::kHeaderSizeAll, L"Size all columns to fit");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);

  for (size_t i = 0; i < columns.size(); ++i) {
    UINT state = (i < visible.size() && visible[i]) ? MF_CHECKED : MF_UNCHECKED;
    AppendMenuW(menu, MF_STRING | state, cmd::kHeaderToggleBase + static_cast<int>(i), columns[i].title.c_str());
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command == cmd::kHeaderSizeToFit && column_hit >= 0) {
    int subitem = GetListViewColumnSubItem(search_results_list_, column_hit);
    ListView_SetColumnWidth(search_results_list_, column_hit, LVSCW_AUTOSIZE_USEHEADER);
    if (subitem >= 0 && static_cast<size_t>(subitem) < widths.size()) {
      widths[static_cast<size_t>(subitem)] = ListView_GetColumnWidth(search_results_list_, column_hit);
    }
    return;
  }
  if (command == cmd::kHeaderSizeAll) {
    int last_visible = FindLastVisibleColumn(visible);
    for (size_t i = 0; i < columns.size(); ++i) {
      if (i < visible.size() && !visible[i]) {
        continue;
      }
      int display_index = FindListViewColumnBySubItem(search_results_list_, static_cast<int>(i));
      if (display_index < 0) {
        continue;
      }
      int width = 0;
      if (static_cast<int>(i) == last_visible) {
        width = CalcListViewColumnFitWidth(search_results_list_, static_cast<int>(i), columns[i].width);
        ListView_SetColumnWidth(search_results_list_, display_index, width);
      } else {
        ListView_SetColumnWidth(search_results_list_, display_index, LVSCW_AUTOSIZE_USEHEADER);
        width = ListView_GetColumnWidth(search_results_list_, display_index);
      }
      widths[i] = width;
    }
    return;
  }
  if (command >= cmd::kHeaderToggleBase) {
    int index = command - cmd::kHeaderToggleBase;
    if (index >= 0 && static_cast<size_t>(index) < columns.size()) {
      bool show = !(index < static_cast<int>(visible.size()) && visible[static_cast<size_t>(index)]);
      ToggleSearchColumn(index, show);
    }
  }
}

void MainWindow::Impl::ToggleValueColumn(int column, bool visible) {
  if (column < 0 || static_cast<size_t>(column) >= browse_.columns().visible.size()) {
    return;
  }
  if (visible == browse_.columns().visible[static_cast<size_t>(column)]) {
    return;
  }

  if (visible) {
    int width = browse_.columns().widths[static_cast<size_t>(column)];
    if (width <= 0) {
      width = browse_.columns().items[static_cast<size_t>(column)].width;
    }
    browse_.columns().visible[static_cast<size_t>(column)] = true;
    browse_.columns().widths[static_cast<size_t>(column)] = width;
  } else {
    int display_index = FindListViewColumnBySubItem(browse_.values().hwnd(), column);
    int width = display_index >= 0 ? ListView_GetColumnWidth(browse_.values().hwnd(), display_index) : browse_.columns().widths[static_cast<size_t>(column)];
    if (width > 0) {
      browse_.columns().widths[static_cast<size_t>(column)] = width;
    }
    browse_.columns().visible[static_cast<size_t>(column)] = false;
  }
  ApplyValueColumns();
}

void MainWindow::Impl::ToggleHistoryColumn(int column, bool visible) {
  if (column < 0 || static_cast<size_t>(column) >= history_column_visible_.size()) {
    return;
  }
  if (visible == history_column_visible_[static_cast<size_t>(column)]) {
    return;
  }

  if (visible) {
    int width = history_column_widths_[static_cast<size_t>(column)];
    if (width <= 0) {
      width = history_columns_[static_cast<size_t>(column)].width;
    }
    history_column_visible_[static_cast<size_t>(column)] = true;
    history_column_widths_[static_cast<size_t>(column)] = width;
  } else {
    int display_index = FindListViewColumnBySubItem(history_list_, column);
    int width = display_index >= 0 ? ListView_GetColumnWidth(history_list_, display_index) : history_column_widths_[static_cast<size_t>(column)];
    if (width > 0) {
      history_column_widths_[static_cast<size_t>(column)] = width;
    }
    history_column_visible_[static_cast<size_t>(column)] = false;
  }
  ApplyHistoryColumns();
}

void MainWindow::Impl::ToggleSearchColumn(int column, bool visible) {
  bool compare = IsCompareTabSelected();
  auto& columns = compare ? compare_columns_ : search_columns_;
  auto& widths = compare ? compare_column_widths_ : search_column_widths_;
  auto& visibility = compare ? compare_column_visible_ : search_column_visible_;
  if (column < 0 || static_cast<size_t>(column) >= visibility.size() || static_cast<size_t>(column) >= columns.size()) {
    return;
  }
  if (visible == visibility[static_cast<size_t>(column)]) {
    return;
  }

  if (visible) {
    int width = widths[static_cast<size_t>(column)];
    if (width <= 0) {
      width = columns[static_cast<size_t>(column)].width;
    }
    visibility[static_cast<size_t>(column)] = true;
    widths[static_cast<size_t>(column)] = width;
  } else {
    int display_index = FindListViewColumnBySubItem(search_results_list_, column);
    int width = display_index >= 0 ? ListView_GetColumnWidth(search_results_list_, display_index) : widths[static_cast<size_t>(column)];
    if (width > 0) {
      widths[static_cast<size_t>(column)] = width;
    }
    visibility[static_cast<size_t>(column)] = false;
  }
  ApplySearchColumns(compare);
}

void MainWindow::Impl::DrawAddressButton(const DRAWITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  const Theme& theme = Theme::Current();
  HDC hdc = info->hDC;
  RECT rect = info->rcItem;
  bool pressed = (info->itemState & ODS_SELECTED) != 0;

  COLORREF bg_color = pressed ? theme.HoverColor() : theme.SurfaceColor();
  FillRect(hdc, &rect, appearance::CachedBrush(bg_color));

  HPEN pen = appearance::CachedPen(theme.BorderColor());
  HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
  MoveToEx(hdc, rect.left, rect.top + 3, nullptr);
  LineTo(hdc, rect.left, rect.bottom - 3);
  SelectObject(hdc, old_pen);

  if (info->CtlID == kAddressGoId) {
    if (address_go_icon_) {
      UINT dpi = win32::DpiForWindow(hwnd_);
      int icon_size = util::ScaleForDpi(kToolbarGlyphSize, dpi);
      int icon_x = rect.left + (rect.right - rect.left - icon_size) / 2;
      int icon_y = rect.top + (rect.bottom - rect.top - icon_size) / 2;
      DrawIconEx(hdc, icon_x, icon_y, address_go_icon_, icon_size, icon_size, 0, nullptr, DI_NORMAL);
    } else {
      POINT pts[3] = {
          {rect.left + 8, rect.top + 6},
          {rect.left + 8, rect.bottom - 6},
          {rect.right - 6, (rect.top + rect.bottom) / 2},
      };
      COLORREF arrow_color = theme.MutedTextColor();
      HBRUSH arrow_brush = appearance::CachedBrush(arrow_color);
      HBRUSH old_brush = reinterpret_cast<HBRUSH>(SelectObject(hdc, arrow_brush));
      HPEN arrow_pen = appearance::CachedPen(arrow_color);
      HPEN old_arrow = reinterpret_cast<HPEN>(SelectObject(hdc, arrow_pen));
      Polygon(hdc, pts, 3);
      SelectObject(hdc, old_arrow);
      SelectObject(hdc, old_brush);
    }
  }
}

void MainWindow::Impl::DrawHeaderCloseButton(const DRAWITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  const Theme& theme = Theme::Current();
  HDC hdc = info->hDC;
  RECT rect = info->rcItem;
  bool pressed = (info->itemState & ODS_SELECTED) != 0;

  COLORREF bg_color = pressed ? theme.HoverColor() : theme.HeaderColor();
  FillRect(hdc, &rect, appearance::CachedBrush(bg_color));

  const UINT dpi = win32::DpiForWindow(info->hwndItem);
  const int radius = util::ScaleForDpi(3, dpi);
  const int pen_width = std::max(1, util::ScaleForDpi(1, dpi));
  const int center_x = (rect.left + rect.right) / 2;
  const int center_y = (rect.top + rect.bottom) / 2;
  HPEN pen = appearance::CachedPen(theme.MutedTextColor(), pen_width);
  HGDIOBJ old_pen = SelectObject(hdc, pen);
  MoveToEx(hdc, center_x - radius, center_y - radius, nullptr);
  LineTo(hdc, center_x + radius + 1, center_y + radius + 1);
  MoveToEx(hdc, center_x + radius, center_y - radius, nullptr);
  LineTo(hdc, center_x - radius - 1, center_y + radius + 1);
  SelectObject(hdc, old_pen);
}

bool MainWindow::Impl::SelectTreePath(const std::wstring& path) {
  if (!browse_.tree().hwnd()) {
    return false;
  }
  std::vector<std::wstring> parts = BuildVisibleTreePathParts(path);
  if (parts.empty()) {
    return false;
  }

  HTREEITEM root = TreeView_GetRoot(browse_.tree().hwnd());
  HTREEITEM current = root;
  for (const auto& part : parts) {
    TreeView_Expand(browse_.tree().hwnd(), current, TVE_EXPAND);
    HTREEITEM child = FindChildByText(browse_.tree().hwnd(), current, part);
    if (!child) {
      RefreshTreeItem(current);
      child = FindChildByText(browse_.tree().hwnd(), current, part);
    }
    if (!child) {
      return false;
    }
    current = child;
  }

  if (current) {
    TreeView_SelectItem(browse_.tree().hwnd(), current);
    TreeView_EnsureVisible(browse_.tree().hwnd(), current);
    return true;
  }
  return false;
}

bool MainWindow::Impl::SelectValueByName(const std::wstring& name) {
  return browse_.SelectValue(name);
}

void MainWindow::Impl::HandleTypeToSelectList(wchar_t ch) {
  browse_.TypeSelectValues(ch, GetTickCount());
}

void MainWindow::Impl::HandleTypeToSelectTree(wchar_t ch) {
  browse_.TypeSelectTree(ch, GetTickCount());
}

} // namespace regkit
