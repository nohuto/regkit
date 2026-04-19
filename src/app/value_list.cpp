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

#include "app/value_list.h"

namespace regkit {

namespace {
bool ContainsInsensitive(const std::wstring& text, const std::wstring& needle) {
  if (needle.empty()) {
    return true;
  }
  if (text.empty()) {
    return false;
  }
  return FindStringOrdinal(FIND_FROMSTART, text.c_str(), static_cast<int>(text.size()), needle.c_str(), static_cast<int>(needle.size()), TRUE) >= 0;
}

} // namespace

void ValueList::Create(HWND parent, HINSTANCE instance, int control_id) {
  hwnd_ = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | LVS_EDITLABELS, 0, 0, 100, 100, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)), instance, nullptr);
  DWORD ex_mask = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_BORDERSELECT | LVS_EX_TRACKSELECT | LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE | LVS_EX_UNDERLINEHOT;
  DWORD ex_style = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER;
  ListView_SetExtendedListViewStyleEx(hwnd_, ex_mask, ex_style);
  SendMessageW(hwnd_, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
}

HWND ValueList::hwnd() const {
  return hwnd_;
}

void ValueList::SetColumns(const std::vector<ColumnInfo>& columns) {
  int count = Header_GetItemCount(ListView_GetHeader(hwnd_));
  for (int i = count - 1; i >= 0; --i) {
    ListView_DeleteColumn(hwnd_, i);
  }

  int index = 0;
  for (const auto& column : columns) {
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
    col.pszText = const_cast<wchar_t*>(column.title.c_str());
    col.cx = column.width;
    col.fmt = column.fmt;
    ListView_InsertColumn(hwnd_, index++, &col);
  }
}

void ValueList::SetRows(std::vector<ListRow> rows) {
  rows_ = std::move(rows);
  RebuildFilter();
}

void ValueList::SetImageList(HIMAGELIST image_list) {
  ListView_SetImageList(hwnd_, image_list, LVSIL_SMALL);
}

void ValueList::Clear() {
  rows_.clear();
  visible_indices_.clear();
  ListView_SetItemCountEx(hwnd_, 0, LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
  InvalidateRect(hwnd_, nullptr, TRUE);
}

void ValueList::SetFilter(const std::wstring& text) {
  if (filter_text_ == text) {
    return;
  }
  filter_text_ = text;
  RebuildFilter();
}

void ValueList::RebuildFilter() {
  visible_indices_.clear();
  visible_indices_.reserve(rows_.size());
  if (filter_text_.empty()) {
    for (size_t i = 0; i < rows_.size(); ++i) {
      visible_indices_.push_back(static_cast<int>(i));
    }
  } else {
    for (size_t i = 0; i < rows_.size(); ++i) {
      const auto& row = rows_[i];
      if (ContainsInsensitive(row.name, filter_text_) || ContainsInsensitive(row.type, filter_text_) || ContainsInsensitive(row.data, filter_text_) || ContainsInsensitive(row.default_data, filter_text_) || ContainsInsensitive(row.read_on_boot, filter_text_) || ContainsInsensitive(row.extra, filter_text_) || ContainsInsensitive(row.size, filter_text_) || ContainsInsensitive(row.date, filter_text_) || ContainsInsensitive(row.details, filter_text_) || ContainsInsensitive(row.comment, filter_text_)) {
        visible_indices_.push_back(static_cast<int>(i));
      }
    }
  }
  ListView_SetItemCountEx(hwnd_, static_cast<int>(visible_indices_.size()), LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
  InvalidateRect(hwnd_, nullptr, TRUE);
}

bool ValueList::HasFilter() const {
  return !filter_text_.empty();
}

size_t ValueList::RowCount() const {
  return visible_indices_.size();
}

const ListRow* ValueList::RowAt(int index) const {
  if (index < 0 || static_cast<size_t>(index) >= visible_indices_.size()) {
    return nullptr;
  }
  int mapped = visible_indices_[static_cast<size_t>(index)];
  if (mapped < 0 || static_cast<size_t>(mapped) >= rows_.size()) {
    return nullptr;
  }
  return &rows_[static_cast<size_t>(mapped)];
}

ListRow* ValueList::MutableRowAt(int index) {
  if (index < 0 || static_cast<size_t>(index) >= visible_indices_.size()) {
    return nullptr;
  }
  int mapped = visible_indices_[static_cast<size_t>(index)];
  if (mapped < 0 || static_cast<size_t>(mapped) >= rows_.size()) {
    return nullptr;
  }
  return &rows_[static_cast<size_t>(mapped)];
}

std::vector<ListRow>& ValueList::rows() {
  return rows_;
}

const std::vector<ListRow>& ValueList::rows() const {
  return rows_;
}

} // namespace regkit
