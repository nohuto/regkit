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

#include "../../include/app/value_list.h"

#include <algorithm>
#include <cwctype>

#include "../../include/win32/win32_helpers.h"

namespace regkit {

namespace {
void AppendSearchField(std::wstring* out, const std::wstring& text) {
  if (!out || text.empty()) {
    return;
  }
  if (!out->empty()) {
    out->push_back(L'\x1f');
  }
  out->reserve(out->size() + text.size());
  for (wchar_t ch : text) {
    out->push_back(static_cast<wchar_t>(towlower(ch)));
  }
}

std::wstring BuildSearchText(const ListRow& row) {
  std::wstring text;
  size_t reserve = row.name.size() + row.type.size() + row.data.size() + row.default_data.size() + row.read_on_boot.size() + row.extra.size() + row.size.size() + row.date.size() + row.details.size() + row.comment.size() + 10;
  text.reserve(reserve);
  AppendSearchField(&text, row.name);
  AppendSearchField(&text, row.type);
  AppendSearchField(&text, row.data);
  AppendSearchField(&text, row.default_data);
  AppendSearchField(&text, row.read_on_boot);
  AppendSearchField(&text, row.extra);
  AppendSearchField(&text, row.size);
  AppendSearchField(&text, row.date);
  AppendSearchField(&text, row.details);
  AppendSearchField(&text, row.comment);
  return text;
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

void ValueList::SetRows(std::vector<ListRow> rows) {
  rows_ = std::move(rows);
  filter_cache_.clear();
  filter_cache_valid_.clear();
  filter_cache_.resize(rows_.size());
  filter_cache_valid_.resize(rows_.size(), false);
  RebuildFilter();
}

void ValueList::SetImageList(HIMAGELIST image_list) {
  ListView_SetImageList(hwnd_, image_list, LVSIL_SMALL);
}

void ValueList::Clear() {
  rows_.clear();
  visible_indices_.clear();
  filter_cache_.clear();
  filter_cache_valid_.clear();
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
    if (filter_cache_.size() != rows_.size()) {
      filter_cache_.assign(rows_.size(), std::wstring());
      filter_cache_valid_.assign(rows_.size(), false);
    }
    std::wstring filter = util::ToLower(filter_text_);
    for (size_t i = 0; i < rows_.size(); ++i) {
      if (!filter_cache_valid_[i]) {
        filter_cache_[i] = BuildSearchText(rows_[i]);
        filter_cache_valid_[i] = true;
      }
      if (filter_cache_[i].find(filter) != std::wstring::npos) {
        visible_indices_.push_back(static_cast<int>(i));
      }
    }
  }
  ListView_SetItemCountEx(hwnd_, static_cast<int>(visible_indices_.size()), LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
  InvalidateRect(hwnd_, nullptr, TRUE);
}

void ValueList::InvalidateFilterCache() {
  std::fill(filter_cache_valid_.begin(), filter_cache_valid_.end(), false);
}

void ValueList::InvalidateFilterCache(const ListRow* row) {
  if (!row || rows_.empty()) {
    return;
  }
  const ListRow* first = rows_.data();
  const ListRow* last = first + rows_.size();
  if (row < first || row >= last) {
    return;
  }
  size_t index = static_cast<size_t>(row - first);
  if (index < filter_cache_valid_.size()) {
    filter_cache_valid_[index] = false;
  }
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
