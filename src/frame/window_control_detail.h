// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "frame/window_registry_detail.h"
#include "appearance/font_metrics.h"

#include "frame/window_impl.h"

#include <algorithm>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <cwchar>
#include <cwctype>
#include <exception>
#include <functional>
#include <limits>
#include <regex>

#include <commdlg.h>
#include <pathcch.h>
#include <richedit.h>
#include <shellapi.h>
#include <shldisp.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <windowsx.h>
#include <winternl.h>

#include "frame/command_ids.h"
#include "registry/security_dialog.h"
#include "appearance/feedback.h"
#include "appearance/gdi_cache.h"
#include "registry/registry_store.h"
#include "appearance/icon_loader.h"

#include "defaults/default_loader.h"
#include "editors/comment_editor.h"
#include "editors/value_editor.h"
#include "frame/message_dispatch.h"
#include "frame/message_ids.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "search/result_file.h"
#include "trace/trace_loader.h"
#include "trace/trace_parser.h"
#include "workspace/settings.h"
#include "workspace/tab_state.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/shell_paths.h"
#include "resource.h"

namespace regkit::window_detail {
inline int CalcEditHeight(HWND hwnd, HFONT font, int min_height) {
  int height = min_height;
  if (!hwnd || !font) {
    return height;
  }
  HDC hdc = GetDC(hwnd);
  HFONT old = reinterpret_cast<HFONT>(SelectObject(hdc, font));
  TEXTMETRICW tm = {};
  if (GetTextMetricsW(hdc, &tm)) {
    int metric_height = static_cast<int>(tm.tmHeight + tm.tmExternalLeading + 6);
    height = std::max(height, metric_height);
  }
  SelectObject(hdc, old);
  ReleaseDC(hwnd, hdc);
  return height;
}

inline void SetEditMargins(HWND hwnd, int left, int right) {
  if (!hwnd) {
    return;
  }
  SendMessageW(hwnd, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELONG(left, right));
}

inline void SetEditVerticalRect(HWND hwnd, HFONT font, int min_pad, int left_pad, int right_pad) {
  if (!hwnd) {
    return;
  }
  RECT rect = {};
  GetClientRect(hwnd, &rect);
  rect.left += left_pad;
  rect.right -= right_pad;
  int pad = min_pad;
  if (font) {
    HDC hdc = GetDC(hwnd);
    HFONT old = reinterpret_cast<HFONT>(SelectObject(hdc, font));
    TEXTMETRICW tm = {};
    if (GetTextMetricsW(hdc, &tm)) {
      int line_height = static_cast<int>(tm.tmHeight + tm.tmExternalLeading);
      int available = rect.bottom - rect.top;
      int centered = (available - line_height) / 2;
      if (centered > pad) {
        pad = centered;
      }
      int max_line = std::max(1, available - pad * 2);
      if (line_height > max_line) {
        line_height = max_line;
      }
      rect.top += pad;
      rect.bottom = rect.top + line_height;
      SelectObject(hdc, old);
      ReleaseDC(hwnd, hdc);
      SendMessageW(hwnd, EM_SETRECT, 0, reinterpret_cast<LPARAM>(&rect));
      return;
    }
    SelectObject(hdc, old);
    ReleaseDC(hwnd, hdc);
  }
  rect.top += pad;
  rect.bottom -= pad;
  SendMessageW(hwnd, EM_SETRECT, 0, reinterpret_cast<LPARAM>(&rect));
}

inline void DrawToolbarButtonBackground(HDC hdc, const RECT& rect, COLORREF fill, COLORREF border) {
  if (!hdc) {
    return;
  }
  RECT draw = rect;
  InflateRect(&draw, -1, -1);
  HBRUSH brush = appearance::CachedBrush(fill);
  HPEN pen = appearance::CachedPen(border);
  HGDIOBJ old_brush = SelectObject(hdc, brush);
  HGDIOBJ old_pen = SelectObject(hdc, pen);
  RoundRect(hdc, draw.left, draw.top, draw.right, draw.bottom, 4, 4);
  SelectObject(hdc, old_pen);
  SelectObject(hdc, old_brush);
}

inline RegistryNode MakeChildNode(const RegistryNode& parent, const std::wstring& name) {
  RegistryNode child = parent;
  if (child.subkey.empty()) {
    child.subkey = name;
  } else {
    child.subkey = child.subkey + L"\\" + name;
  }
  return child;
}

inline std::wstring LeafName(const RegistryNode& node) {
  if (node.subkey.empty()) {
    return node.root_name.empty() ? registry_path::RootName(node.root) : node.root_name;
  }
  return registry_path::Leaf(node.subkey);
}

inline bool UseBinaryValueIcon(DWORD type) {
  switch (type) {
  case REG_NONE:
  case REG_BINARY:
  case REG_DWORD:
  case REG_DWORD_BIG_ENDIAN:
  case REG_QWORD:
  case REG_RESOURCE_LIST:
  case REG_FULL_RESOURCE_DESCRIPTOR:
  case REG_RESOURCE_REQUIREMENTS_LIST:
  case REG_LINK:
    return true;
  default:
    return false;
  }
}

inline ListRow MakeValueListRow(const std::wstring& name, DWORD type,
                                const BYTE* data, DWORD data_size) {
  ListRow row;
  row.name = name.empty() ? L"(Default)" : name;
  row.type = value_format::TypeName(type);
  row.data_ready = data_size == 0 || data != nullptr;
  if (row.data_ready && data_size > 0) {
    row.data = value_format::DisplayData(type, data, data_size);
  }
  row.image_index = UseBinaryValueIcon(type) ? kBinaryIconIndex
                                              : kValueIconIndex;
  row.kind = rowkind::kValue;
  row.extra = name;
  row.size = std::to_wstring(data_size);
  row.size_value = data_size;
  row.has_size = true;
  row.value_type = type;
  row.value_data_size = data_size;
  return row;
}

inline void UpdateLeafName(RegistryNode* node, const std::wstring& new_name) {
  if (!node || node->subkey.empty()) {
    return;
  }
  size_t pos = node->subkey.rfind(L'\\');
  if (pos == std::wstring::npos) {
    node->subkey = new_name;
  } else {
    node->subkey = node->subkey.substr(0, pos + 1) + new_name;
  }
}

inline std::wstring FormatFileTime(const FILETIME& filetime) {
  if (filetime.dwLowDateTime == 0 && filetime.dwHighDateTime == 0) {
    return L"";
  }
  FILETIME local = {};
  SYSTEMTIME st = {};
  if (!FileTimeToLocalFileTime(&filetime, &local) || !FileTimeToSystemTime(&local, &st)) {
    return L"";
  }
  wchar_t buffer[64] = {};
  swprintf_s(buffer, L"%d/%d/%d %d:%02d", st.wMonth, st.wDay, st.wYear, st.wHour, st.wMinute);
  return buffer;
}

inline std::wstring FormatCommentDisplay(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  bool last_space = false;
  for (wchar_t ch : text) {
    if (ch == L'\r' || ch == L'\n' || ch == L'\t') {
      ch = L' ';
    }
    if (ch == L' ') {
      if (last_space) {
        continue;
      }
      last_space = true;
    } else {
      last_space = false;
    }
    out.push_back(ch);
  }
  return out;
}

inline uint64_t FileTimeToUint64(const FILETIME& filetime) {
  ULARGE_INTEGER value = {};
  value.LowPart = filetime.dwLowDateTime;
  value.HighPart = filetime.dwHighDateTime;
  return value.QuadPart;
}

inline int CompareTextInsensitive(const std::wstring& left, const std::wstring& right) {
  if (left.empty()) {
    return right.empty() ? 0 : 1;
  }
  if (right.empty()) {
    return -1;
  }
  int result = CompareStringOrdinal(left.c_str(), static_cast<int>(left.size()), right.c_str(), static_cast<int>(right.size()), TRUE);
  if (result == CSTR_LESS_THAN) {
    return -1;
  }
  if (result == CSTR_GREATER_THAN) {
    return 1;
  }
  return 0;
}

inline int CompareUint64(uint64_t left, uint64_t right) {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
}

constexpr int kCellTooltipPadding = 8;
constexpr size_t kCellTooltipMeasureLimit = 512;
constexpr size_t kCellTextDrawLimit = 512;
constexpr size_t kValuePreviewLimit = 4096;

inline const std::wstring& ValueRowFieldText(const ListRow& row, int subitem) {
  switch (subitem) {
  case kValueColName:
    return row.name;
  case kValueColType:
    return row.type;
  case kValueColData:
    return row.data;
  case kValueColDefault:
    return row.default_data;
  case kValueColReadOnBoot:
    return row.read_on_boot;
  case kValueColSize:
    return row.size;
  case kValueColDate:
    return row.date;
  case kValueColDetails:
    return row.details;
  case kValueColComment:
    return row.comment;
  default:
    return row.extra;
  }
}

inline bool CellTextIsClipped(HWND list, const std::wstring& text, int available) {
  if (text.size() > kCellTooltipMeasureLimit) {
    return true;
  }
  HDC dc = GetDC(list);
  if (!dc) {
    return false;
  }
  HFONT font = reinterpret_cast<HFONT>(SendMessageW(list, WM_GETFONT, 0, 0));
  HGDIOBJ old_font = font ? SelectObject(dc, font) : nullptr;
  SIZE size = {};
  GetTextExtentPoint32W(dc, text.c_str(), static_cast<int>(text.size()), &size);
  if (old_font) {
    SelectObject(dc, old_font);
  }
  ReleaseDC(list, dc);
  return size.cx > available;
}

inline int CompareValueRow(const ListRow& left, const ListRow& right, int column) {
  if (left.kind != right.kind) {
    return (left.kind == rowkind::kKey) ? -1 : 1;
  }
  switch (column) {
  case kValueColName:
    return CompareTextInsensitive(left.name, right.name);
  case kValueColType:
    return CompareTextInsensitive(left.type, right.type);
  case kValueColData:
    return CompareTextInsensitive(left.data, right.data);
  case kValueColDefault:
    return CompareTextInsensitive(left.default_data, right.default_data);
  case kValueColReadOnBoot:
    return CompareTextInsensitive(left.read_on_boot, right.read_on_boot);
  case kValueColSize:
    if (left.has_size != right.has_size) {
      return left.has_size ? -1 : 1;
    }
    return CompareUint64(left.size_value, right.size_value);
  case kValueColDate:
    if (left.has_date != right.has_date) {
      return left.has_date ? -1 : 1;
    }
    return CompareUint64(left.date_value, right.date_value);
  case kValueColDetails:
    if (left.has_details != right.has_details) {
      return left.has_details ? -1 : 1;
    }
    if (left.detail_key_count != right.detail_key_count) {
      return CompareUint64(left.detail_key_count, right.detail_key_count);
    }
    return CompareUint64(left.detail_value_count, right.detail_value_count);
  case kValueColComment:
    return CompareTextInsensitive(left.comment, right.comment);
  default:
    return CompareTextInsensitive(left.name, right.name);
  }
}

inline void SortValueRows(std::vector<ListRow>* rows, int column, bool ascending) {
  if (!rows || rows->size() < 2) {
    return;
  }
  std::stable_sort(rows->begin(), rows->end(), [column, ascending](const ListRow& left, const ListRow& right) {
    int result = CompareValueRow(left, right, column);
    if (result == 0) {
      return false;
    }
    return ascending ? (result < 0) : (result > 0);
  });
}

inline void UpdateListViewSort(HWND list, int column, bool ascending) {
  if (!list) {
    return;
  }
  HWND header = ListView_GetHeader(list);
  if (!header) {
    return;
  }
  int count = Header_GetItemCount(header);
  for (int i = 0; i < count; ++i) {
    HDITEMW item = {};
    item.mask = HDI_FORMAT;
    if (!Header_GetItem(header, i, &item)) {
      continue;
    }
    const int current = item.fmt & (HDF_SORTUP | HDF_SORTDOWN);
    int wanted = 0;
    if (column >= 0 && GetListViewColumnSubItem(list, i) == column) {
      wanted = ascending ? HDF_SORTUP : HDF_SORTDOWN;
    }
    if (current == wanted) {
      continue;
    }
    item.fmt = (item.fmt & ~(HDF_SORTUP | HDF_SORTDOWN)) | wanted;
    Header_SetItem(header, i, &item);
  }
}

inline HFONT CreateUIFont() {
  return static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
}

inline HFONT CreateIconFont(int point_size) {
  int height = appearance::FontHeight(point_size);
  return CreateFontW(height, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe MDL2 Assets");
}

inline void ApplyFont(HWND hwnd, HFONT font) {
  if (hwnd && font) {
    SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
  }
}

inline HTREEITEM FindChildByText(HWND tree, HTREEITEM parent, const std::wstring& text) {
  wchar_t buffer[256] = {};
  HTREEITEM child = TreeView_GetChild(tree, parent);
  while (child) {
    TVITEMW item = {};
    item.mask = TVIF_TEXT;
    item.hItem = child;
    item.pszText = buffer;
    item.cchTextMax = static_cast<int>(_countof(buffer));
    if (TreeView_GetItem(tree, &item)) {
      if (EqualsInsensitive(text, buffer)) {
        return child;
      }
    }
    child = TreeView_GetNextSibling(tree, child);
  }
  return nullptr;
}

} // namespace regkit::window_detail
