// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

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
#include "win32/system_error.h"
#include "win32/text_transform.h"
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
constexpr int kToolbarId = 100;
constexpr int kAddressEditId = 0;
constexpr int kTreeId = 1;
constexpr int kValueListId = 2;
constexpr int kRegeditCompatEditId = 101;
constexpr int kRegeditCompatTreeId = 104;
constexpr int kRegeditCompatListId = 105;
constexpr int kTabId = 103;
constexpr int kHistoryListId = 106;
constexpr int kHistoryLabelId = 107;
constexpr int kTreeHeaderId = 108;
constexpr int kAddressGoId = 109;
constexpr int kTreeHeaderCloseId = 110;
constexpr int kStatusBarId = 111;
constexpr int kSearchResultsListId = 112;
constexpr int kSearchProgressId = 113;
constexpr int kHistoryHeaderCloseId = 114;
constexpr int kFilterEditId = 115;
constexpr int kToolbarIconSize = 16;
constexpr int kToolbarGlyphSize = 16;
constexpr wchar_t kRestartSystemArg[] = L"--restart-system";
constexpr wchar_t kRestartTiArg[] = L"--restart-ti";
template <typename T>
inline T ClampValue(T value, T low, T high) {
  return value < low ? low : (high < value ? high : value);
}

inline char* MutableData(std::string& text) {
  return text.empty() ? nullptr : &text[0];
}

inline wchar_t* MutableData(std::wstring& text) {
  return text.empty() ? nullptr : &text[0];
}

template <typename T>
inline void ReleasePostedPayload(std::unique_ptr<T>& payload) {
  T* posted_payload = payload.release();
  (void)posted_payload;
}

constexpr wchar_t kStandardGroupLabel[] = L"Standart Hives";
constexpr wchar_t kRealGroupLabel[] = L"REGISTRY";

constexpr size_t kSearchResultsBatch = 1024;
constexpr DWORD kSearchResultsMaxMs = 15;
constexpr DWORD kSearchResultsRefreshMs = 1000;
constexpr DWORD kSearchProgressUiMs = 500;
constexpr size_t kSearchQueueBatch = 128;
constexpr UINT_PTR kRegeditCompatApplyTimerId = 34;
constexpr UINT kRegeditCompatApplyDelayMs = 75;
constexpr ULONG_PTR kExternalJumpCopyDataId = 0x52474A54;
constexpr ULONG_PTR kRegeditCompatActivateCopyDataId = 0x52474354;
constexpr UINT_PTR kAddressSubclassId = 1;
constexpr UINT_PTR kTabSubclassId = 2;
constexpr UINT_PTR kHeaderSubclassId = 3;
constexpr UINT_PTR kListViewSubclassId = 4;
constexpr UINT_PTR kTreeViewSubclassId = 5;
constexpr UINT_PTR kAutoCompletePopupSubclassId = 6;
constexpr UINT_PTR kAutoCompleteListBoxSubclassId = 7;
constexpr UINT_PTR kFilterSubclassId = 8;
constexpr wchar_t kRegeditWindowClassName[] = L"RegEdit_RegEdit";
constexpr wchar_t kRegKitWindowProperty[] = L"RegKitMainWindow";
constexpr wchar_t kRegeditIfeoPath[] = L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\regedit.exe";
using ui::ListViewItemSelected;
using util::FormatWin32Error;
using util::ToLower;
using util::TrimWhitespace;

constexpr int kToolbarSepGroup1 = 30001;
constexpr int kToolbarSepGroup2 = 30002;
constexpr int kToolbarSepGroup3 = 30003;

constexpr int kFolderIconIndex = 0;
constexpr int kSymlinkIconIndex = 1;
constexpr int kDatabaseIconIndex = 2;
constexpr int kFolderSimIconIndex = 3;
constexpr int kValueIconIndex = 4;
constexpr int kBinaryIconIndex = 5;
constexpr int kHeaderTextPadding = 6;
constexpr int kTabMinWidth = 90;
constexpr int kTabInsetX = 2;
constexpr int kTabInsetY = 2;
constexpr int kTabTextPaddingX = 10;
constexpr int kTabCloseSize = 14;
constexpr int kTabCloseGap = 6;
constexpr int kSplitterWidth = 6;
constexpr int kHistorySplitterHeight = 4;
constexpr int kMinTreeWidth = 160;
constexpr int kMinValueListWidth = 240;
constexpr int kMinHistoryHeight = 80;
constexpr int kHistoryMaxPadding = 140;
constexpr int kHistoryGap = 2;
constexpr int kBorderInflate = 1;
constexpr int kValueColName = 0;
constexpr int kValueColType = 1;
constexpr int kValueColData = 2;
constexpr int kValueColDefault = 3;
constexpr int kValueColReadOnBoot = 4;
constexpr int kValueColSize = 5;
constexpr int kValueColDate = 6;
constexpr int kValueColDetails = 7;
constexpr int kValueColComment = 8;

struct TraceParseBatch : work::MoveOnly {
  uint64_t generation = 0;
  std::wstring source_lower;
  std::vector<KeyValueDialogEntry> entries;
  std::unordered_set<std::wstring> affected_keys;
  std::wstring error;
  bool done = false;
  bool cancelled = false;
};

struct DefaultParseBatch : work::MoveOnly {
  uint64_t generation = 0;
  std::wstring source_lower;
  std::vector<KeyValueDialogEntry> entries;
  std::unordered_set<std::wstring> affected_keys;
  std::wstring error;
  bool done = false;
  bool cancelled = false;
};

struct ValueListPayload : work::MoveOnly {
  uint64_t generation = 0;
  std::vector<ListRow> rows;
  int key_count = 0;
  int value_count = 0;
};

std::wstring NormalizeTraceKeyPathBasic(const std::wstring& text);
std::wstring ResolveRegistryLinkPath(const std::wstring& path);

inline bool GetChildRectInParent(HWND parent, HWND child, RECT* rect) {
  if (!parent || !child || !rect) {
    return false;
  }
  if (!GetWindowRect(child, rect)) {
    return false;
  }
  MapWindowPoints(nullptr, parent, reinterpret_cast<POINT*>(rect), 2);
  return true;
}

inline RECT InflateCopy(RECT rect, int dx, int dy) {
  InflateRect(&rect, dx, dy);
  return rect;
}

inline void DrawOutlineRect(HDC hdc, const RECT& rect, int inflate) {
  if (!hdc) {
    return;
  }
  RECT draw = InflateCopy(rect, inflate, inflate);
  Rectangle(hdc, draw.left, draw.top, draw.right, draw.bottom);
}

inline RECT AdjustTabDrawRect(const RECT& item_rect, int header_bottom, bool selected) {
  RECT rect = item_rect;
  rect.left += kTabInsetX;
  rect.right -= kTabInsetX;
  rect.top += kTabInsetY;
  rect.bottom = header_bottom - 1;
  if (selected) {
    rect.top -= 1;
    rect.bottom = header_bottom;
  }
  return rect;
}

inline bool CalcTabCloseRect(const RECT& tab_rect, RECT* close_rect) {
  if (!close_rect) {
    return false;
  }
  int height = tab_rect.bottom - tab_rect.top;
  int size = std::min(kTabCloseSize, std::max(8, height - 6));
  if (size <= 0) {
    return false;
  }
  int right = tab_rect.right - kTabCloseGap;
  close_rect->right = right;
  close_rect->left = right - size;
  close_rect->top = tab_rect.top + (height - size) / 2;
  close_rect->bottom = close_rect->top + size;
  return close_rect->left < close_rect->right;
}

inline int GetListViewColumnSubItem(HWND list, int display_index) {
  if (!list || display_index < 0) {
    return display_index;
  }
  LVCOLUMNW col = {};
  col.mask = LVCF_SUBITEM;
  if (ListView_GetColumn(list, display_index, &col)) {
    return col.iSubItem;
  }
  return display_index;
}

inline bool DrawSearchMatchSubItem(const search::Result& result, int subitem, HDC hdc, const RECT& rect, HFONT font) {
  bool match_subitem = false;
  if (result.match_field == search::MatchField::kPath && subitem == 0) {
    match_subitem = true;
  } else if (result.match_field == search::MatchField::kName && subitem == 1) {
    match_subitem = true;
  } else if (result.match_field == search::MatchField::kData && subitem == 3) {
    match_subitem = true;
  }
  if (!match_subitem || result.match_start < 0 || result.match_length <= 0) {
    return false;
  }

  const wchar_t* text = L"";
  if (subitem == 0) {
    text = result.key_path.c_str();
  } else if (subitem == 1) {
    text = result.display_name.c_str();
  } else if (subitem == 3) {
    text = result.data.c_str();
  }
  size_t text_len = wcslen(text);
  if (result.match_start < 0 || result.match_start >= static_cast<int>(text_len)) {
    return false;
  }
  size_t match_end = static_cast<size_t>(result.match_start + result.match_length);
  if (match_end > text_len) {
    match_end = text_len;
  }
  if (match_end <= static_cast<size_t>(result.match_start)) {
    return false;
  }

  const Theme& theme = Theme::Current();
  COLORREF fg = theme.TextColor();
  HBRUSH bg_brush = appearance::CachedBrush(theme.PanelColor());
  FillRect(hdc, &rect, bg_brush);

  HFONT old_font = nullptr;
  if (font) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, font));
  }
  SetBkMode(hdc, TRANSPARENT);
  int padding = 6;
  RECT clip = rect;
  clip.left += padding;
  clip.right -= padding;
  int x = clip.left;
  SIZE size = {};
  GetTextExtentPoint32W(hdc, L"Ag", 2, &size);
  int y = rect.top + (rect.bottom - rect.top - size.cy) / 2;
  size = {};
  std::wstring prefix(text, text + result.match_start);
  std::wstring match(text + result.match_start, text + match_end);
  std::wstring suffix(text + match_end);

  auto draw_segment = [&](const std::wstring& segment, COLORREF color) {
    if (segment.empty()) {
      return;
    }
    SIZE seg_size = {};
    GetTextExtentPoint32W(hdc, segment.c_str(), static_cast<int>(segment.size()), &seg_size);
    SetTextColor(hdc, color);
    ExtTextOutW(hdc, x, y, ETO_CLIPPED, &clip, segment.c_str(), static_cast<UINT>(segment.size()), nullptr);
    x += seg_size.cx;
  };
  draw_segment(prefix, fg);
  draw_segment(match, theme.FocusColor());
  draw_segment(suffix, fg);

  if (old_font) {
    SelectObject(hdc, old_font);
  }
  return true;
}

inline void DrawHistoryListItem(HWND list, HDC hdc, int item_index, bool hot, HFONT font) {
  if (!list || !hdc || item_index < 0) {
    return;
  }
  RECT row_rect = {};
  if (!ListView_GetItemRect(list, item_index, &row_rect, LVIR_BOUNDS)) {
    return;
  }
  const Theme& theme = Theme::Current();
  COLORREF bg = theme.PanelColor();
  COLORREF fg = theme.TextColor();
  if (hot) {
    bg = theme.HoverColor();
  }
  FillRect(hdc, &row_rect, appearance::CachedBrush(bg));

  HWND header = ListView_GetHeader(list);
  int column_count = header ? Header_GetItemCount(header) : 0;
  if (column_count <= 0) {
    return;
  }

  HFONT old_font = nullptr;
  if (font) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, font));
  }
  int old_bk_mode = SetBkMode(hdc, TRANSPARENT);
  COLORREF old_color = SetTextColor(hdc, fg);

  for (int display_index = 0; display_index < column_count; ++display_index) {
    LVCOLUMNW col = {};
    col.mask = LVCF_FMT | LVCF_SUBITEM;
    if (!ListView_GetColumn(list, display_index, &col)) {
      continue;
    }
    int subitem = col.iSubItem;
    RECT cell_rect = {};
    if (!ListView_GetSubItemRect(list, item_index, subitem, LVIR_LABEL, &cell_rect)) {
      continue;
    }
    wchar_t text[512] = {};
    ListView_GetItemText(list, item_index, subitem, text, static_cast<int>(_countof(text)));
    if (text[0] == L'\0') {
      continue;
    }
    UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
    if (col.fmt & LVCFMT_RIGHT) {
      format |= DT_RIGHT;
    } else if (col.fmt & LVCFMT_CENTER) {
      format |= DT_CENTER;
    }
    RECT text_rect = cell_rect;
    text_rect.left += kHeaderTextPadding;
    text_rect.right -= kHeaderTextPadding;
    DrawTextW(hdc, text, -1, &text_rect, format);
  }

  bool show_grid = (ListView_GetExtendedListViewStyle(list) & LVS_EX_GRIDLINES) != 0;
  COLORREF grid = theme.BorderColor();
  HPEN pen = appearance::CachedPen(grid, 1);
  HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
  for (int display_index = 0; header && display_index < column_count; ++display_index) {
    RECT header_rect = {};
    if (!Header_GetItemRect(header, display_index, &header_rect)) {
      continue;
    }
    MapWindowPoints(header, list, reinterpret_cast<POINT*>(&header_rect), 2);
    int x = header_rect.right - 1;
    if (x <= row_rect.left || x >= row_rect.right) {
      continue;
    }
    MoveToEx(hdc, x, row_rect.top, nullptr);
    LineTo(hdc, x, row_rect.bottom);
  }
  if (show_grid) {
    int y = row_rect.bottom - 1;
    MoveToEx(hdc, row_rect.left, y, nullptr);
    LineTo(hdc, row_rect.right, y);
  }
  SelectObject(hdc, old_pen);

  SetTextColor(hdc, old_color);
  SetBkMode(hdc, old_bk_mode);
  if (old_font) {
    SelectObject(hdc, old_font);
  }
}

inline LRESULT HandleHistoryListCustomDraw(HWND list, NMLVCUSTOMDRAW* draw) {
  if (!list || !draw) {
    return CDRF_DODEFAULT;
  }
  switch (draw->nmcd.dwDrawStage) {
  case CDDS_PREPAINT:
    return CDRF_NOTIFYITEMDRAW;
  case CDDS_ITEMPREPAINT: {
    int item_index = static_cast<int>(draw->nmcd.dwItemSpec);
    bool selected = ListViewItemSelected(list, item_index);
    if (selected) {
      return CDRF_DODEFAULT;
    }
    bool hot = (draw->nmcd.uItemState & CDIS_HOT) != 0;
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(list, WM_GETFONT, 0, 0));
    DrawHistoryListItem(list, draw->nmcd.hdc, item_index, hot, font);
    return CDRF_SKIPDEFAULT;
  }
  default:
    break;
  }
  return CDRF_DODEFAULT;
}

inline int FindListViewColumnBySubItem(HWND list, int subitem) {
  if (!list || subitem < 0) {
    return -1;
  }
  HWND header = ListView_GetHeader(list);
  int count = header ? Header_GetItemCount(header) : 0;
  for (int i = 0; i < count; ++i) {
    if (GetListViewColumnSubItem(list, i) == subitem) {
      return i;
    }
  }
  return -1;
}

inline int FetchListViewItemText(HWND list, int index, int column, std::wstring* buffer) {
  if (!list || !buffer) {
    return 0;
  }
  if (buffer->empty()) {
    buffer->resize(1);
  }
  LVITEMW item = {};
  item.iSubItem = column;
  item.pszText = MutableData(*buffer);
  item.cchTextMax = static_cast<int>(buffer->size());
  int length = static_cast<int>(SendMessageW(list, LVM_GETITEMTEXTW, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&item)));
  if (length >= static_cast<int>(buffer->size() - 1)) {
    buffer->resize(static_cast<size_t>(length) + 2);
    item.pszText = MutableData(*buffer);
    item.cchTextMax = static_cast<int>(buffer->size());
    length = static_cast<int>(SendMessageW(list, LVM_GETITEMTEXTW, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&item)));
  }
  return length;
}

inline int CalcListViewColumnFitWidth(HWND list, int column, int min_width) {
  if (!list || column < 0) {
    return min_width;
  }
  int display_index = FindListViewColumnBySubItem(list, column);
  if (display_index < 0) {
    return min_width;
  }
  int width = min_width;
  wchar_t header_text[256] = {};
  LVCOLUMNW col = {};
  col.mask = LVCF_TEXT;
  col.pszText = header_text;
  col.cchTextMax = static_cast<int>(_countof(header_text));
  if (ListView_GetColumn(list, display_index, &col)) {
    int header_width = ListView_GetStringWidth(list, header_text) + 18;
    width = std::max(width, header_width);
  }

  int count = ListView_GetItemCount(list);
  std::wstring buffer;
  buffer.resize(256);
  for (int i = 0; i < count; ++i) {
    int length = FetchListViewItemText(list, i, column, &buffer);
    if (length > 0) {
      int text_width = ListView_GetStringWidth(list, buffer.c_str()) + 18;
      if (text_width > width) {
        width = text_width;
      }
    }
  }
  return width;
}

inline int FindLastVisibleColumn(const std::vector<bool>& visible) {
  for (int i = static_cast<int>(visible.size()) - 1; i >= 0; --i) {
    if (visible[static_cast<size_t>(i)]) {
      return i;
    }
  }
  return -1;
}

} // namespace regkit::window_detail
