// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>
#include <commctrl.h>

#include <cstdint>
#include <string>
#include <vector>

namespace regkit {

struct ColumnInfo {
  std::wstring title;
  int width = 0;
  int fmt = LVCFMT_LEFT;
};

struct ListRow {
  std::wstring name;
  std::wstring type;
  std::wstring data;
  std::wstring default_data;
  std::wstring read_on_boot;
  std::wstring extra;
  std::wstring size;
  std::wstring date;
  std::wstring details;
  std::wstring comment;
  uint64_t size_value = 0;
  uint64_t date_value = 0;
  uint64_t detail_key_count = 0;
  uint64_t detail_value_count = 0;
  bool has_size = false;
  bool has_date = false;
  bool has_details = false;
  DWORD value_type = 0;
  DWORD value_data_size = 0;
  bool data_ready = false;
  bool simulated = false;
  int image_index = 0;
  LPARAM kind = 0;
};

class ValueList {
public:
  void Create(HWND parent, HINSTANCE instance, int control_id);
  HWND hwnd() const;
  void SetRows(std::vector<ListRow> rows);
  int AppendRow(ListRow row, bool force_visible = false);
  void SetImageList(HIMAGELIST image_list);
  void Clear();
  void SetFilter(const std::wstring& text);
  void RebuildFilter();
  void InvalidateFilterCache();
  void InvalidateFilterCache(const ListRow* row);
  bool HasFilter() const;
  size_t RowCount() const;
  const ListRow* RowAt(int index) const;
  ListRow* MutableRowAt(int index);
  std::vector<ListRow>& rows();
  const std::vector<ListRow>& rows() const;

private:
  HWND hwnd_ = nullptr;
  std::vector<ListRow> rows_;
  std::vector<int> visible_indices_;
  std::vector<std::wstring> filter_cache_;
  std::vector<bool> filter_cache_valid_;
  std::wstring filter_text_;
};

} // namespace regkit
