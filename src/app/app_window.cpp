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

#include "../../include/app/app_window.h"

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

#include "../../include/app/command_ids.h"
#include "../../include/app/registry_security.h"
#include "../../include/app/ui_helpers.h"
#include "../../include/app/value_dialogs.h"
#include "../../include/registry/registry_provider.h"
#include "../../include/win32/icon_resources.h"
#include "../../include/win32/win32_helpers.h"
#include "defaults/default_loader.h"
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
#include "../../resources/resource.h"

namespace regkit {

struct MainWindow::TraceLoadPayload : work::MoveOnly {
  uint64_t generation = 0;
  std::vector<ActiveTrace> traces;
  std::unordered_map<std::wstring, trace::Selection> selection_cache;
};

struct MainWindow::DefaultLoadPayload : work::MoveOnly {
  uint64_t generation = 0;
  std::vector<ActiveDefault> defaults;
};

struct MainWindow::StartupCachePayload : work::MoveOnly {
  uint64_t generation = 0;
  std::vector<HistoryEntry> history_entries;
  std::vector<changes::CommentEntry> value_comments;
  std::vector<changes::CommentEntry> name_comments;
  std::wstring tree_selected_path;
  std::vector<std::wstring> tree_expanded_paths;
  bool history_loaded = false;
  bool comments_loaded = false;
  bool tree_state_loaded = false;
};

struct MainWindow::ReplacePayload : work::MoveOnly {
  struct Change {
    changes::UndoOperation undo;
    HistoryEntry history;
  };

  uint64_t generation = 0;
  std::vector<Change> changes;
  int failures = 0;
  bool cancelled = false;
};

namespace {
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
T ClampValue(T value, T low, T high) {
  return value < low ? low : (high < value ? high : value);
}

char* MutableData(std::string& text) {
  return text.empty() ? nullptr : &text[0];
}

wchar_t* MutableData(std::wstring& text) {
  return text.empty() ? nullptr : &text[0];
}

template <typename T>
void ReleasePostedPayload(std::unique_ptr<T>& payload) {
  T* posted_payload = payload.release();
  (void)posted_payload;
}

bool ParseBundledTraceLabel(const std::wstring& path, std::wstring* label);
constexpr wchar_t kStandardGroupLabel[] = L"Standart Hives";
constexpr wchar_t kRealGroupLabel[] = L"REGISTRY";

constexpr UINT kAddressEnterMessage = WM_APP + 10;
constexpr UINT kFocusAddressBarMessage = WM_APP + 11;
constexpr UINT kSearchResultsMessage = WM_APP + 20;
constexpr UINT kSearchFinishedMessage = WM_APP + 21;
constexpr UINT kSearchFailedMessage = WM_APP + 22;
constexpr UINT kSearchProgressMessage = WM_APP + 23;
constexpr UINT kLoadTracesMessage = WM_APP + 24;
constexpr UINT kTraceLoadReadyMessage = WM_APP + 25;
constexpr UINT kLoadDefaultsMessage = WM_APP + 26;
constexpr UINT kDefaultLoadReadyMessage = WM_APP + 27;
constexpr size_t kSearchResultsBatch = 1024;
constexpr DWORD kSearchResultsMaxMs = 15;
constexpr DWORD kSearchResultsRefreshMs = 1000;
constexpr DWORD kSearchProgressUiMs = 500;
constexpr size_t kSearchQueueBatch = 128;
constexpr UINT kValueListReadyMessage = WM_APP + 30;
constexpr UINT kTraceParseBatchMessage = WM_APP + 31;
constexpr UINT kDefaultParseBatchMessage = WM_APP + 32;
constexpr UINT kRegFileLoadReadyMessage = WM_APP + 33;
constexpr UINT kDeferredStartupMessage = WM_APP + 34;
constexpr UINT kStartupCacheReadyMessage = WM_APP + 35;
constexpr UINT kReplaceReadyMessage = WM_APP + 36;
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
constexpr DWORD kTypeSelectTimeoutMs = 1000;
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

HBRUSH GetCachedBrush(COLORREF color);
HPEN GetCachedPen(COLORREF color, int width = 1);
std::wstring NormalizeTraceKeyPathBasic(const std::wstring& text);
std::wstring ResolveRegistryLinkPath(const std::wstring& path);

bool GetChildRectInParent(HWND parent, HWND child, RECT* rect) {
  if (!parent || !child || !rect) {
    return false;
  }
  if (!GetWindowRect(child, rect)) {
    return false;
  }
  MapWindowPoints(nullptr, parent, reinterpret_cast<POINT*>(rect), 2);
  return true;
}

RECT InflateCopy(RECT rect, int dx, int dy) {
  InflateRect(&rect, dx, dy);
  return rect;
}

void DrawOutlineRect(HDC hdc, const RECT& rect, int inflate) {
  if (!hdc) {
    return;
  }
  RECT draw = InflateCopy(rect, inflate, inflate);
  Rectangle(hdc, draw.left, draw.top, draw.right, draw.bottom);
}

RECT AdjustTabDrawRect(const RECT& item_rect, int header_bottom, bool selected) {
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

bool CalcTabCloseRect(const RECT& tab_rect, RECT* close_rect) {
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

int GetListViewColumnSubItem(HWND list, int display_index) {
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

bool ListViewItemSelected(HWND list, int item_index) {
  return item_index >= 0 && (ListView_GetItemState(list, item_index, LVIS_SELECTED) & LVIS_SELECTED) != 0;
}

bool DrawSearchMatchSubItem(const search::Result& result, int subitem, HDC hdc, const RECT& rect, HFONT font) {
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
  HBRUSH bg_brush = GetCachedBrush(theme.PanelColor());
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

void DrawHistoryListItem(HWND list, HDC hdc, int item_index, bool hot, HFONT font) {
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
  FillRect(hdc, &row_rect, GetCachedBrush(bg));

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
  HPEN pen = GetCachedPen(grid, 1);
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

LRESULT HandleHistoryListCustomDraw(HWND list, NMLVCUSTOMDRAW* draw) {
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

int FindListViewColumnBySubItem(HWND list, int subitem) {
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

int FetchListViewItemText(HWND list, int index, int column, std::wstring* buffer) {
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

int CalcListViewColumnFitWidth(HWND list, int column, int min_width) {
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

int FindLastVisibleColumn(const std::vector<bool>& visible) {
  for (int i = static_cast<int>(visible.size()) - 1; i >= 0; --i) {
    if (visible[static_cast<size_t>(i)]) {
      return i;
    }
  }
  return -1;
}

bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  std::wstring buffer(32768, L'\0');
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer.data();
  ofn.nMaxFile = static_cast<DWORD>(buffer.size());
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
  if (!GetOpenFileNameW(&ofn)) {
    return false;
  }
  *path = buffer.c_str();
  return true;
}

bool PromptSaveFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  std::wstring buffer(32768, L'\0');
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer.data();
  ofn.nMaxFile = static_cast<DWORD>(buffer.size());
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
  if (!GetSaveFileNameW(&ofn)) {
    return false;
  }
  *path = buffer.c_str();
  return true;
}

std::wstring TrimTrailingSeparators(const std::wstring& path) {
  std::wstring result = path;
  while (!result.empty() && (result.back() == L'\\' || result.back() == L'/')) {
    result.pop_back();
  }
  return result;
}

bool IsDirectoryPath(const std::wstring& path) {
  DWORD attrs = GetFileAttributesW(path.c_str());
  return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

constexpr wchar_t kIconSetDefault[] = L"default";
constexpr wchar_t kIconSetTabler[] = L"tabler";
constexpr wchar_t kIconSetFluentUi[] = L"fluentui";
constexpr wchar_t kIconSetLucide[] = L"lucide";
constexpr wchar_t kIconSetMaterialSymbols[] = L"materialsymbols";
constexpr wchar_t kIconSetCustom[] = L"custom";

bool IsIconSetName(const std::wstring& value, const wchar_t* name) {
  return _wcsicmp(value.c_str(), name) == 0;
}

bool IsKnownIconSetName(const std::wstring& value) {
  return IsIconSetName(value, kIconSetDefault) || IsIconSetName(value, kIconSetTabler) || IsIconSetName(value, kIconSetFluentUi) || IsIconSetName(value, kIconSetMaterialSymbols) || IsIconSetName(value, kIconSetCustom);
}

std::wstring FindAssetsIconsRoot() {
  std::wstring base = util::GetModuleDirectory();
  for (int i = 0; i < 6; ++i) {
    if (base.empty()) {
      break;
    }
    std::wstring candidate = util::JoinPath(base, L"assets\\icons");
    if (IsDirectoryPath(candidate)) {
      return candidate;
    }
    base = registry_path::Parent(base);
  }
  DWORD len = GetCurrentDirectoryW(0, nullptr);
  if (len > 0) {
    std::wstring cwd(len, L'\0');
    DWORD written = GetCurrentDirectoryW(len, MutableData(cwd));
    if (written != 0) {
      if (written < cwd.size() && cwd[written] == L'\0') {
        cwd.resize(written);
      }
      base = cwd;
      for (int i = 0; i < 3; ++i) {
        if (base.empty()) {
          break;
        }
        std::wstring candidate = util::JoinPath(base, L"assets\\icons");
        if (IsDirectoryPath(candidate)) {
          return candidate;
        }
        base = registry_path::Parent(base);
      }
    }
  }
  return L"";
}

std::wstring AssetsIconsRoot() {
  static std::wstring cached;
  static bool cached_set = false;
  if (!cached_set) {
    cached = FindAssetsIconsRoot();
    cached_set = true;
  }
  return cached;
}

constexpr DWORD kOfflinePickFolderButtonId = 0x2001;

std::wstring ShellItemPath(IShellItem* item) {
  if (!item) {
    return L"";
  }
  PWSTR raw = nullptr;
  if (FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &raw)) || !raw) {
    return L"";
  }
  std::wstring result(raw);
  CoTaskMemFree(raw);
  return result;
}

class OfflinePickerEvents final : public IFileDialogEvents, public IFileDialogControlEvents {
public:
  explicit OfflinePickerEvents(IFileDialog* dialog) : dialog_(dialog) {
    if (dialog_) {
      dialog_->AddRef();
    }
  }

  std::wstring picked_path() const { return picked_path_; }

  IFACEMETHODIMP QueryInterface(REFIID riid, void** result) override {
    if (!result) {
      return E_POINTER;
    }
    *result = nullptr;
    if (riid == IID_IUnknown || riid == IID_IFileDialogEvents) {
      *result = static_cast<IFileDialogEvents*>(this);
    } else if (riid == IID_IFileDialogControlEvents) {
      *result = static_cast<IFileDialogControlEvents*>(this);
    } else {
      return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
  }

  IFACEMETHODIMP_(ULONG) AddRef() override { return static_cast<ULONG>(InterlockedIncrement(&ref_)); }

  IFACEMETHODIMP_(ULONG) Release() override {
    ULONG ref = static_cast<ULONG>(InterlockedDecrement(&ref_));
    if (ref == 0) {
      delete this;
    }
    return ref;
  }

  IFACEMETHODIMP OnFileOk(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnFolderChanging(IFileDialog*, IShellItem*) override { return S_OK; }
  IFACEMETHODIMP OnFolderChange(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnSelectionChange(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnShareViolation(IFileDialog*, IShellItem*, FDE_SHAREVIOLATION_RESPONSE* response) override {
    if (response) {
      *response = FDESVR_DEFAULT;
    }
    return S_OK;
  }
  IFACEMETHODIMP OnTypeChange(IFileDialog*) override { return S_OK; }
  IFACEMETHODIMP OnOverwrite(IFileDialog*, IShellItem*, FDE_OVERWRITE_RESPONSE* response) override {
    if (response) {
      *response = FDEOR_DEFAULT;
    }
    return S_OK;
  }

  IFACEMETHODIMP OnItemSelected(IFileDialogCustomize*, DWORD, DWORD) override { return S_OK; }
  IFACEMETHODIMP OnButtonClicked(IFileDialogCustomize*, DWORD id) override {
    if (id != kOfflinePickFolderButtonId || !dialog_) {
      return S_OK;
    }
    picked_path_.clear();
    IShellItem* selection = nullptr;
    if (SUCCEEDED(dialog_->GetCurrentSelection(&selection)) && selection) {
      SFGAOF attrs = 0;
      if (SUCCEEDED(selection->GetAttributes(SFGAO_FOLDER, &attrs)) && (attrs & SFGAO_FOLDER)) {
        picked_path_ = ShellItemPath(selection);
      }
      selection->Release();
    }
    if (picked_path_.empty()) {
      IShellItem* folder = nullptr;
      if (SUCCEEDED(dialog_->GetFolder(&folder)) && folder) {
        picked_path_ = ShellItemPath(folder);
        folder->Release();
      }
    }
    if (!picked_path_.empty()) {
      dialog_->Close(S_OK);
    }
    return S_OK;
  }
  IFACEMETHODIMP OnCheckButtonToggled(IFileDialogCustomize*, DWORD, BOOL) override { return S_OK; }
  IFACEMETHODIMP OnControlActivating(IFileDialogCustomize*, DWORD) override { return S_OK; }

private:
  ~OfflinePickerEvents() {
    if (dialog_) {
      dialog_->Release();
    }
  }

  LONG ref_ = 1;
  IFileDialog* dialog_ = nullptr;
  std::wstring picked_path_;
};

bool PromptOpenFolderOrFile(HWND owner, const wchar_t* title, std::wstring* path) {
  if (!path) {
    return false;
  }
  HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
  bool uninit = false;
  if (SUCCEEDED(hr)) {
    uninit = true;
  } else if (hr != RPC_E_CHANGED_MODE) {
    return false;
  }

  IFileOpenDialog* dialog = nullptr;
  hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog));
  if (FAILED(hr) || !dialog) {
    if (uninit) {
      CoUninitialize();
    }
    return false;
  }

  DWORD options = 0;
  dialog->GetOptions(&options);
  options |= FOS_FORCEFILESYSTEM | FOS_FILEMUSTEXIST | FOS_PATHMUSTEXIST;
  dialog->SetOptions(options);
  if (title) {
    dialog->SetTitle(title);
  }
  const COMDLG_FILTERSPEC filters[] = {
      {L"Registry Hive Files (*.dat;*.hiv;*.hive;*.sav;SYSTEM;SOFTWARE;SAM;SECURITY;DEFAULT;NTUSER.DAT;USRCLASS.DAT)", L"*.dat;*.hiv;*.hive;*.sav;SYSTEM;SOFTWARE;SAM;SECURITY;DEFAULT;NTUSER.DAT;USRCLASS.DAT"},
      {L"All Files (*.*)", L"*.*"},
  };
  dialog->SetFileTypes(static_cast<UINT>(_countof(filters)), filters);
  dialog->SetFileTypeIndex(1);

  IFileDialogCustomize* customize = nullptr;
  if (SUCCEEDED(dialog->QueryInterface(IID_PPV_ARGS(&customize))) && customize) {
    customize->AddPushButton(kOfflinePickFolderButtonId, L"Select Folder");
    customize->Release();
  }

  OfflinePickerEvents* events = new OfflinePickerEvents(dialog);
  DWORD cookie = 0;
  if (events) {
    dialog->Advise(events, &cookie);
  }

  hr = dialog->Show(owner);
  if (cookie) {
    dialog->Unadvise(cookie);
  }

  std::wstring selected;
  if (events) {
    selected = events->picked_path();
    events->Release();
  }

  if (selected.empty() && SUCCEEDED(hr)) {
    IShellItem* item = nullptr;
    if (SUCCEEDED(dialog->GetResult(&item)) && item) {
      selected = ShellItemPath(item);
      item->Release();
    }
  }

  dialog->Release();
  if (uninit) {
    CoUninitialize();
  }

  if (selected.empty() || hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) {
    return false;
  }
  *path = selected;
  return true;
}

bool HasRegExtension(const std::wstring& path) {
  size_t dot = path.find_last_of(L'.');
  if (dot == std::wstring::npos) {
    return false;
  }
  std::wstring ext = path.substr(dot);
  return _wcsicmp(ext.c_str(), L".reg") == 0;
}

std::wstring EnsureRegExtension(std::wstring path) {
  if (path.empty() || HasRegExtension(path)) {
    return path;
  }
  path.append(L".reg");
  return path;
}

bool IsWhitespaceOnly(const std::wstring& text) {
  for (wchar_t ch : text) {
    if (!iswspace(static_cast<wint_t>(ch))) {
      return false;
    }
  }
  return true;
}

std::wstring NormalizeMachineName(const std::wstring& text) {
  std::wstring trimmed = TrimWhitespace(text);
  while (!trimmed.empty() && (trimmed.back() == L'\\' || trimmed.back() == L'/')) {
    trimmed.pop_back();
  }
  if (trimmed.empty()) {
    return trimmed;
  }
  if (trimmed.rfind(L"\\\\", 0) == 0) {
    return trimmed;
  }
  return L"\\\\" + trimmed;
}

std::wstring StripMachinePrefix(const std::wstring& machine) {
  if (machine.rfind(L"\\\\", 0) == 0) {
    return machine.substr(2);
  }
  return machine;
}

bool FileExists(const std::wstring& path) {
  DWORD attrs = GetFileAttributesW(path.c_str());
  return attrs != INVALID_FILE_ATTRIBUTES && !(attrs & FILE_ATTRIBUTE_DIRECTORY);
}

bool EqualsInsensitive(const std::wstring& left, const std::wstring& right) {
  return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

bool StartsWithInsensitive(const std::wstring& text, const std::wstring& prefix) {
  if (prefix.empty()) {
    return true;
  }
  if (text.size() < prefix.size()) {
    return false;
  }
  return CompareStringOrdinal(text.c_str(), static_cast<int>(prefix.size()), prefix.c_str(), static_cast<int>(prefix.size()), TRUE) == CSTR_EQUAL;
}

std::wstring ReadFontSubstitute(const wchar_t* value_name) {
  if (!value_name || !*value_name) {
    return L"";
  }
  const wchar_t* subkey = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\FontSubstitutes";
  auto query = [&](REGSAM sam) -> std::wstring {
    HKEY key = nullptr;
    LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, subkey, 0, sam, &key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    DWORD type = 0;
    DWORD bytes = 0;
    result = RegQueryValueExW(key, value_name, nullptr, &type, nullptr, &bytes);
    if (result != ERROR_SUCCESS || bytes == 0 || (type != REG_SZ && type != REG_EXPAND_SZ)) {
      RegCloseKey(key);
      return L"";
    }
    std::vector<wchar_t> buffer(bytes / sizeof(wchar_t) + 1, L'\0');
    result = RegQueryValueExW(key, value_name, nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &bytes);
    RegCloseKey(key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    std::wstring value(buffer.data());
    while (!value.empty() && value.back() == L'\0') {
      value.pop_back();
    }
    if (value.empty()) {
      return L"";
    }
    if (type == REG_EXPAND_SZ) {
      std::wstring expanded = util::ExpandEnvironmentStringsDynamic(value);
      if (!expanded.empty()) {
        value = std::move(expanded);
      }
    }
    return value;
  };

  std::wstring value = query(KEY_READ | KEY_WOW64_64KEY);
  if (!value.empty()) {
    return value;
  }
  return query(KEY_READ);
}

bool WindowClassEquals(HWND hwnd, const wchar_t* class_name) {
  if (!hwnd || !class_name) {
    return false;
  }
  wchar_t buffer[64] = {};
  if (!GetClassNameW(hwnd, buffer, static_cast<int>(_countof(buffer)))) {
    return false;
  }
  return _wcsicmp(buffer, class_name) == 0;
}

std::wstring GetDefaultRegeditPath() {
  DWORD needed = GetWindowsDirectoryW(nullptr, 0);
  if (needed == 0) {
    return L"";
  }
  std::wstring windows_dir(needed, L'\0');
  DWORD written = GetWindowsDirectoryW(windows_dir.data(), needed);
  if (written == 0 || written >= needed) {
    return L"";
  }
  windows_dir.resize(written);
  return util::JoinPath(windows_dir, L"regedit.exe");
}

bool LaunchDefaultRegeditProcess(const std::wstring& regedit_path, DWORD* error_code) {
  if (error_code) {
    *error_code = ERROR_SUCCESS;
  }
  if (regedit_path.empty()) {
    if (error_code) {
      *error_code = ERROR_FILE_NOT_FOUND;
    }
    return false;
  }
  std::wstring command_line = L"\"" + regedit_path + L"\" /m";
  STARTUPINFOW startup = {};
  startup.cb = sizeof(startup);
  PROCESS_INFORMATION process = {};
  if (!CreateProcessW(nullptr, command_line.data(), nullptr, nullptr, FALSE, 0, nullptr, nullptr, &startup, &process)) {
    if (error_code) {
      *error_code = GetLastError();
    }
    return false;
  }
  CloseHandle(process.hThread);
  CloseHandle(process.hProcess);
  return true;
}

struct ParsedRegFileRoot {
  std::wstring name;
  std::shared_ptr<VirtualRegistryData> data;
};

struct RegFileParsePayload : work::MoveOnly {
  uint64_t generation = 0;
  std::wstring source_path;
  std::wstring source_lower;
  std::vector<ParsedRegFileRoot> roots;
  std::wstring error;
  bool cancelled = false;
};

VirtualRegistryKey* EnsureVirtualKey(VirtualRegistryKey* root, const std::wstring& subkey) {
  if (!root) {
    return nullptr;
  }
  if (subkey.empty()) {
    return root;
  }
  auto parts = registry_path::Split(subkey);
  VirtualRegistryKey* current = root;
  for (const auto& part : parts) {
    std::wstring lower = ToLower(part);
    auto it = current->children.find(lower);
    if (it == current->children.end()) {
      auto child = std::make_unique<VirtualRegistryKey>();
      child->name = part;
      it = current->children.emplace(lower, std::move(child)).first;
    }
    current = it->second.get();
  }
  return current;
}

bool ParseRegFileToVirtualRoots(const std::wstring& path, std::vector<ParsedRegFileRoot>* roots, std::wstring* error, const std::atomic_bool* cancel, bool* cancelled) {
  if (!roots) {
    return false;
  }
  roots->clear();

  regfile::Document document;
  if (!regfile::Load(path, &document, error, cancel, cancelled)) {
    return false;
  }

  std::unordered_map<std::wstring, size_t> root_lookup;
  auto ensure_root = [&](const std::wstring& root_name) {
    const std::wstring lower = ToLower(root_name);
    auto existing = root_lookup.find(lower);
    if (existing != root_lookup.end()) {
      return roots->at(existing->second).data.get();
    }
    ParsedRegFileRoot root;
    root.name = root_name;
    root.data = std::make_shared<VirtualRegistryData>();
    root.data->root_name = root_name;
    root.data->root = std::make_unique<VirtualRegistryKey>();
    root.data->root->name = root_name;
    roots->push_back(std::move(root));
    root_lookup.emplace(lower, roots->size() - 1);
    return roots->back().data.get();
  };

  for (const auto& source_path : document.key_order) {
    if (cancel && cancel->load()) {
      if (cancelled) {
        *cancelled = true;
      }
      return false;
    }
    const std::wstring normalized = NormalizeTraceKeyPathBasic(source_path);
    const std::wstring key_path = normalized.empty() ? source_path : normalized;
    const size_t slash = key_path.find(L'\\');
    const std::wstring root_name = key_path.substr(0, slash);
    const std::wstring subkey = slash == std::wstring::npos
                                    ? L""
                                    : key_path.substr(slash + 1);
    if (root_name.empty()) {
      continue;
    }
    auto source = document.keys.find(ToLower(source_path));
    if (source == document.keys.end()) {
      continue;
    }
    VirtualRegistryData* root = ensure_root(root_name);
    VirtualRegistryKey* target =
        root ? EnsureVirtualKey(root->root.get(), subkey) : nullptr;
    if (!target) {
      continue;
    }
    target->values.reserve(source->second.values.size());
    for (const auto& entry : source->second.values) {
      VirtualRegistryValue value;
      value.name = entry.second.name;
      value.type = entry.second.type;
      value.data = entry.second.data;
      target->values.emplace(entry.first, std::move(value));
    }
  }
  return true;
}
struct TextMatch {
  bool matched = false;
  size_t start = std::wstring::npos;
  size_t length = 0;
};

class TextMatcher {
public:
  TextMatcher(const std::wstring& query, bool use_regex, bool match_case, bool match_whole, bool* ok) : query_(query), use_regex_(use_regex), match_case_(match_case), match_whole_(match_whole) {
    if (use_regex_) {
      try {
        auto flags = std::regex_constants::ECMAScript;
        if (!match_case_) {
          flags |= std::regex_constants::icase;
        }
        regex_ = std::wregex(query_, flags);
      } catch (const std::regex_error&) {
        if (ok) {
          *ok = false;
        }
      }
    }
  }

  TextMatch Match(const std::wstring& text) const {
    TextMatch match;
    if (text.empty()) {
      return match;
    }
    if (use_regex_) {
      std::wsmatch regex_match;
      if (match_whole_) {
        if (std::regex_match(text, regex_match, regex_)) {
          match.matched = true;
          match.start = 0;
          match.length = regex_match.length();
        }
      } else if (std::regex_search(text, regex_match, regex_)) {
        match.matched = true;
        match.start = regex_match.position();
        match.length = regex_match.length();
      }
      return match;
    }

    if (match_whole_) {
      if (match_case_) {
        if (text == query_) {
          match.matched = true;
          match.start = 0;
          match.length = text.size();
        }
      } else if (CompareStringOrdinal(text.c_str(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE) == CSTR_EQUAL) {
        match.matched = true;
        match.start = 0;
        match.length = text.size();
      }
      return match;
    }

    if (match_case_) {
      size_t pos = text.find(query_);
      if (pos != std::wstring::npos) {
        match.matched = true;
        match.start = pos;
        match.length = query_.size();
      }
    } else {
      int pos = FindStringOrdinal(FIND_FROMSTART, text.c_str(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE);
      if (pos >= 0) {
        match.matched = true;
        match.start = static_cast<size_t>(pos);
        match.length = query_.size();
      }
    }
    return match;
  }

private:
  std::wstring query_;
  bool use_regex_ = false;
  bool match_case_ = false;
  bool match_whole_ = false;
  std::wregex regex_;
};

std::wstring ResolveDevicePath(const std::wstring& path) {
  if (!StartsWithInsensitive(path, L"\\Device\\")) {
    return path;
  }
  wchar_t drives[512] = {};
  DWORD drive_len = GetLogicalDriveStringsW(static_cast<DWORD>(_countof(drives) - 1), drives);
  if (drive_len == 0 || drive_len >= _countof(drives)) {
    return path;
  }
  for (const wchar_t* drive = drives; *drive; drive += wcslen(drive) + 1) {
    wchar_t device[MAX_PATH] = {};
    wchar_t drive_root[4] = {};
    wcsncpy_s(drive_root, drive, _TRUNCATE);
    size_t root_len = wcslen(drive_root);
    if (root_len >= 2 && drive_root[1] == L':') {
      drive_root[2] = L'\0';
    }
    if (!QueryDosDeviceW(drive_root, device, static_cast<DWORD>(_countof(device)))) {
      continue;
    }
    size_t device_len = wcslen(device);
    if (_wcsnicmp(path.c_str(), device, device_len) != 0) {
      continue;
    }
    std::wstring rest = path.substr(device_len);
    if (!rest.empty() && rest.front() != L'\\') {
      rest.insert(rest.begin(), L'\\');
    }
    std::wstring mapped = drive_root;
    mapped += rest;
    return mapped;
  }
  return path;
}

std::wstring NormalizeHiveFilePath(const std::wstring& raw_path) {
  if (raw_path.empty()) {
    return raw_path;
  }
  std::wstring path = raw_path;
  if (StartsWithInsensitive(path, L"\\??\\")) {
    path.erase(0, 4);
  } else if (StartsWithInsensitive(path, L"\\\\?\\")) {
    path.erase(0, 4);
  } else if (StartsWithInsensitive(path, L"\\DosDevices\\")) {
    path.erase(0, wcslen(L"\\DosDevices\\"));
  }
  if (StartsWithInsensitive(path, L"\\SystemRoot")) {
    wchar_t windows_dir[MAX_PATH] = {};
    UINT len = GetWindowsDirectoryW(windows_dir, _countof(windows_dir));
    if (len > 0 && len < _countof(windows_dir)) {
      std::wstring suffix = path.substr(wcslen(L"\\SystemRoot"));
      path = std::wstring(windows_dir) + suffix;
    }
  }
  std::wstring expanded = util::ExpandEnvironmentStringsDynamic(path);
  if (!expanded.empty()) {
    path = std::move(expanded);
  }
  path = ResolveDevicePath(path);
  return path;
}

std::wstring CurrentControlSetSegment() {
  static std::wstring cached;
  static bool loaded = false;
  if (loaded) {
    return cached;
  }
  loaded = true;
  RegistryNode node;
  node.root = HKEY_LOCAL_MACHINE;
  node.subkey = L"SYSTEM\\Select";
  ValueEntry entry;
  if (RegistryProvider::QueryValue(node, L"Current", &entry) && entry.type == REG_DWORD && entry.data.size() >= sizeof(DWORD)) {
    DWORD current = 0;
    std::memcpy(&current, entry.data.data(), sizeof(DWORD));
    wchar_t buffer[32] = {};
    swprintf_s(buffer, L"ControlSet%03u", current);
    cached = buffer;
  }
  return cached;
}

std::wstring ReplaceControlSetSegment(const std::wstring& path, const std::wstring& from, const std::wstring& to) {
  if (path.empty() || from.empty() || to.empty()) {
    return L"";
  }
  std::vector<std::wstring> parts = registry_path::Split(path);
  if (parts.size() < 3) {
    return L"";
  }
  bool is_hklm = EqualsInsensitive(parts[0], L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(parts[0], L"HKLM");
  if (!is_hklm && parts.size() > 1) {
    if (EqualsInsensitive(parts[0], L"REGISTRY") && EqualsInsensitive(parts[1], L"MACHINE")) {
      is_hklm = true;
    }
  }
  if (!is_hklm) {
    return L"";
  }
  for (size_t i = 0; i + 1 < parts.size(); ++i) {
    if (EqualsInsensitive(parts[i], L"SYSTEM") && EqualsInsensitive(parts[i + 1], from)) {
      parts[i + 1] = to;
      return registry_path::Join(parts);
    }
  }
  return L"";
}

std::wstring NormalizeCurrentControlSet(const std::wstring& path) {
  std::wstring current = CurrentControlSetSegment();
  if (current.empty()) {
    return path;
  }
  std::wstring replaced = ReplaceControlSetSegment(path, L"CurrentControlSet", current);
  return replaced.empty() ? path : replaced;
}

bool IsControlSetSegment(const std::wstring& text) {
  constexpr wchar_t kPrefix[] = L"ControlSet";
  size_t prefix_len = wcslen(kPrefix);
  if (text.size() <= prefix_len || !StartsWithInsensitive(text, kPrefix)) {
    return false;
  }
  for (size_t i = prefix_len; i < text.size(); ++i) {
    if (!iswdigit(text[i])) {
      return false;
    }
  }
  return true;
}

std::wstring MapControlSetToCurrent(const std::wstring& path) {
  std::wstring current = CurrentControlSetSegment();
  if (current.empty()) {
    return L"";
  }
  std::vector<std::wstring> parts = registry_path::Split(path);
  if (parts.size() < 3) {
    return L"";
  }
  bool is_hklm = EqualsInsensitive(parts[0], L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(parts[0], L"HKLM");
  if (!is_hklm && parts.size() > 1) {
    if (EqualsInsensitive(parts[0], L"REGISTRY") && EqualsInsensitive(parts[1], L"MACHINE")) {
      is_hklm = true;
    }
  }
  if (!is_hklm) {
    return L"";
  }
  for (size_t i = 0; i + 1 < parts.size(); ++i) {
    if (EqualsInsensitive(parts[i], L"SYSTEM") && IsControlSetSegment(parts[i + 1])) {
      if (EqualsInsensitive(parts[i + 1], current)) {
        return L"";
      }
      parts[i + 1] = current;
      return registry_path::Join(parts);
    }
  }
  return L"";
}

std::wstring CleanTraceKeyText(const std::wstring& text, const std::wstring& sid) {
  std::wstring path = text;
  if (!sid.empty()) {
    const std::wstring marker = L"<CURRENT_USER_SID>";
    size_t pos = path.find(marker);
    while (pos != std::wstring::npos) {
      path.replace(pos, marker.size(), sid);
      pos = path.find(marker, pos + sid.size());
    }
  }
  path = registry_path::Clean(path);
  if (path.empty()) {
    return {};
  }
  wchar_t machine[MAX_COMPUTERNAME_LENGTH + 1] = {};
  DWORD machine_len = static_cast<DWORD>(_countof(machine));
  if (GetComputerNameW(machine, &machine_len) && machine_len > 0) {
    const std::wstring prefix = std::wstring(machine, machine_len) + L"\\";
    if (registry_path::StartsWith(path, prefix)) {
      path.erase(0, prefix.size());
    }
  }
  return path;
}

std::wstring NormalizeTraceKeyPathBasic(const std::wstring& text) {
  const std::wstring sid = util::GetCurrentUserSidString();
  std::wstring path = CleanTraceKeyText(text, sid);
  if (path.empty()) {
    return L"";
  }
  path = registry_path::Normalize(path, sid);
  const std::wstring current_user = L"HKEY_USERS\\" + sid;
  if (!sid.empty() &&
      (registry_path::Equals(path, current_user) ||
       registry_path::StartsWith(path, current_user + L"\\"))) {
    path.replace(0, current_user.size(), L"HKEY_CURRENT_USER");
  }
  RegistryNode node;
  if (registry_path::ParseRoot(path, &node)) {
    return NormalizeCurrentControlSet(path);
  }
  return L"";
}

std::wstring NormalizeTraceKeyPath(const std::wstring& text) {
  std::wstring path = NormalizeTraceKeyPathBasic(text);
  if (path.empty()) {
    return path;
  }
  return ResolveRegistryLinkPath(path);
}

std::wstring NormalizeTraceSelectionPath(const std::wstring& text) {
  std::wstring sid = util::GetCurrentUserSidString();
  std::wstring path = CleanTraceKeyText(text, sid);
  if (path.empty()) {
    return L"";
  }
  if (StartsWithInsensitive(path, L"REGISTRY")) {
    std::wstring rest = path.substr(wcslen(L"REGISTRY"));
    while (!rest.empty() && rest.front() == L'\\') {
      rest.erase(rest.begin());
    }
    return rest.empty() ? L"REGISTRY" : L"REGISTRY\\" + rest;
  }
  return path;
}

trace::Normalizers TraceNormalizers() {
  trace::Normalizers normalizers;
  normalizers.key = [](const std::wstring& path) {
    return NormalizeTraceKeyPath(path);
  };
  normalizers.display = [](const std::wstring& path) {
    return NormalizeTraceSelectionPath(path);
  };
  return normalizers;
}

struct LinkTargetCache {
  std::mutex mutex;
  std::unordered_map<std::wstring, std::wstring> targets;
  std::unordered_set<std::wstring> misses;
};

LinkTargetCache& GetLinkTargetCache() {
  static LinkTargetCache cache;
  return cache;
}

bool ParseRegistryRoot(const std::wstring& input, RegistryNode* node, std::wstring* root_label) {
  if (!node || !root_label) {
    return false;
  }
  if (!registry_path::ParseRoot(input, node)) {
    return false;
  }
  *root_label = node->root_name;
  return true;
}

bool QueryLinkTargetCached(const std::wstring& path, const RegistryNode& node, std::wstring* target) {
  if (!target) {
    return false;
  }
  *target = L"";
  std::wstring key = ToLower(path);
  LinkTargetCache& cache = GetLinkTargetCache();
  {
    std::lock_guard<std::mutex> lock(cache.mutex);
    auto it = cache.targets.find(key);
    if (it != cache.targets.end()) {
      *target = it->second;
      return true;
    }
    if (cache.misses.find(key) != cache.misses.end()) {
      return false;
    }
  }
  std::wstring resolved;
  if (RegistryProvider::QuerySymbolicLinkTarget(node, &resolved)) {
    std::lock_guard<std::mutex> lock(cache.mutex);
    cache.targets.emplace(std::move(key), resolved);
    *target = resolved;
    return true;
  }
  {
    std::lock_guard<std::mutex> lock(cache.mutex);
    cache.misses.insert(std::move(key));
  }
  return false;
}

std::wstring ResolveRegistryLinkPath(const std::wstring& path) {
  if (path.empty()) {
    return path;
  }
  std::wstring current = path;
  std::unordered_set<std::wstring> visited;
  for (int depth = 0; depth < 8; ++depth) {
    std::wstring current_lower = ToLower(current);
    if (!visited.insert(current_lower).second) {
      break;
    }
    RegistryNode root_node;
    std::wstring root_label;
    if (!ParseRegistryRoot(current, &root_node, &root_label)) {
      break;
    }
    std::vector<std::wstring> parts = registry_path::Split(root_node.subkey);
    if (parts.empty()) {
      break;
    }
    std::wstring prefix;
    bool resolved = false;
    for (size_t i = 0; i < parts.size(); ++i) {
      if (!prefix.empty()) {
        prefix.append(L"\\");
      }
      prefix.append(parts[i]);
      RegistryNode node = root_node;
      node.subkey = prefix;
      std::wstring prefix_path = root_label;
      if (!prefix.empty()) {
        prefix_path.append(L"\\");
        prefix_path.append(prefix);
      }
      std::wstring target;
      if (!QueryLinkTargetCached(prefix_path, node, &target)) {
        continue;
      }
      std::wstring mapped_target = NormalizeTraceKeyPathBasic(target);
      if (mapped_target.empty()) {
        continue;
      }
      std::wstring remaining = registry_path::Join(parts, i + 1);
      std::wstring next = mapped_target;
      if (!remaining.empty()) {
        next.append(L"\\");
        next.append(remaining);
      }
      current = next;
      resolved = true;
      break;
    }
    if (!resolved) {
      break;
    }
  }
  return current;
}

std::wstring FileNameOnly(const std::wstring& path) {
  size_t pos = path.find_last_of(L"\\/");
  return (pos == std::wstring::npos) ? path : path.substr(pos + 1);
}

std::wstring FileBaseName(const std::wstring& path) {
  size_t pos = path.find_last_of(L"\\/");
  std::wstring name = (pos == std::wstring::npos) ? path : path.substr(pos + 1);
  size_t dot = name.find_last_of(L'.');
  if (dot != std::wstring::npos) {
    name = name.substr(0, dot);
  }
  return name;
}

struct OfflineHiveCandidate {
  std::wstring path;
  std::wstring label;
};

bool IsFilePath(const std::wstring& path) {
  DWORD attrs = GetFileAttributesW(path.c_str());
  return attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

void AddOfflineHiveCandidate(std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen, const std::wstring& path, const std::wstring& label) {
  if (!out || !seen || !IsFilePath(path)) {
    return;
  }
  std::wstring key = ToLower(path);
  if (!seen->insert(key).second) {
    return;
  }
  std::wstring use_label = TrimWhitespace(label);
  if (use_label.empty()) {
    use_label = TrimWhitespace(FileBaseName(path));
    if (use_label.empty()) {
      use_label = L"OfflineHive";
    }
  }
  out->push_back({path, use_label});
}

std::wstring TopLevelFolderLabel(const std::wstring& base, const std::wstring& folder) {
  std::wstring prefix = base;
  if (!prefix.empty() && prefix.back() != L'\\' && prefix.back() != L'/') {
    prefix.push_back(L'\\');
  }
  if (StartsWithInsensitive(folder, prefix)) {
    std::wstring relative = folder.substr(prefix.size());
    size_t sep = relative.find_first_of(L"\\/");
    if (sep != std::wstring::npos) {
      return relative.substr(0, sep);
    }
    if (!relative.empty()) {
      return relative;
    }
  }
  return FileBaseName(folder);
}

void CollectUserHiveCandidates(const std::wstring& folder, const std::wstring& base, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen) {
  std::wstring label = TopLevelFolderLabel(base, folder);
  std::wstring ntuser = util::JoinPath(folder, L"NTUSER.DAT");
  AddOfflineHiveCandidate(out, seen, ntuser, label);
  std::wstring usrclass = util::JoinPath(folder, L"USRCLASS.DAT");
  std::wstring class_label = label.empty() ? L"" : (label + L"_Classes");
  AddOfflineHiveCandidate(out, seen, usrclass, class_label);
}

void CollectUserHivesRecursive(const std::wstring& folder, const std::wstring& base, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen) {
  WIN32_FIND_DATAW data = {};
  std::wstring search = util::JoinPath(folder, L"*");
  HANDLE find = FindFirstFileW(search.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
      continue;
    }
    if (wcscmp(data.cFileName, L".") == 0 || wcscmp(data.cFileName, L"..") == 0) {
      continue;
    }
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
      continue;
    }
    std::wstring subdir = util::JoinPath(folder, data.cFileName);
    CollectUserHiveCandidates(subdir, base, out, seen);
    CollectUserHivesRecursive(subdir, base, out, seen);
  } while (FindNextFileW(find, &data));
  FindClose(find);
}

bool ShouldIncludeOfflineHiveFile(const std::wstring& name) {
  size_t dot = name.find_last_of(L'.');
  if (dot == std::wstring::npos) {
    return true;
  }
  std::wstring ext = name.substr(dot);
  return _wcsicmp(ext.c_str(), L".dat") == 0;
}

void CollectLooseHivesInFolder(const std::wstring& folder, std::vector<OfflineHiveCandidate>* out, std::unordered_set<std::wstring>* seen) {
  WIN32_FIND_DATAW data = {};
  std::wstring search = util::JoinPath(folder, L"*");
  HANDLE find = FindFirstFileW(search.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if ((data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
      continue;
    }
    if (!ShouldIncludeOfflineHiveFile(data.cFileName)) {
      continue;
    }
    std::wstring candidate = util::JoinPath(folder, data.cFileName);
    std::wstring label = FileBaseName(data.cFileName);
    AddOfflineHiveCandidate(out, seen, candidate, label);
  } while (FindNextFileW(find, &data));
  FindClose(find);
}

void CollectOfflineHivesInFolder(const std::wstring& folder, std::vector<OfflineHiveCandidate>* out) {
  if (!out) {
    return;
  }
  out->clear();
  std::unordered_set<std::wstring> seen;
  static const wchar_t* kMachineHives[] = {
      L"SYSTEM", L"SOFTWARE", L"SAM", L"SECURITY", L"DEFAULT",
  };
  for (const auto* name : kMachineHives) {
    std::wstring candidate = util::JoinPath(folder, name);
    AddOfflineHiveCandidate(out, &seen, candidate, name);
  }
  CollectUserHiveCandidates(folder, folder, out, &seen);
  CollectLooseHivesInFolder(folder, out, &seen);
  CollectUserHivesRecursive(folder, folder, out, &seen);
}

std::wstring ResolveOfflineRootName(const std::wstring& path, bool is_dir, const RegistryNode* current_node) {
  std::wstring base = FileBaseName(path);
  if (is_dir) {
    if (EqualsInsensitive(base, L"HKEY_USERS") || EqualsInsensitive(base, L"HKU")) {
      return L"HKEY_USERS";
    }
    if (EqualsInsensitive(base, L"HKEY_LOCAL_MACHINE") || EqualsInsensitive(base, L"HKLM")) {
      return L"HKEY_LOCAL_MACHINE";
    }
  } else {
    if (EqualsInsensitive(base, L"NTUSER") || EqualsInsensitive(base, L"USRCLASS")) {
      return L"HKEY_USERS";
    }
    if (EqualsInsensitive(base, L"SYSTEM") || EqualsInsensitive(base, L"SOFTWARE") || EqualsInsensitive(base, L"SAM") || EqualsInsensitive(base, L"SECURITY") || EqualsInsensitive(base, L"DEFAULT") || EqualsInsensitive(base, L"COMPONENTS") || EqualsInsensitive(base, L"BCD")) {
      return L"HKEY_LOCAL_MACHINE";
    }
  }
  if (current_node && (current_node->root == HKEY_LOCAL_MACHINE || current_node->root == HKEY_USERS)) {
    std::wstring root_name = registry_path::RootName(current_node->root);
    if (!root_name.empty()) {
      return root_name;
    }
  }
  return L"HKEY_LOCAL_MACHINE";
}

int CalcEditHeight(HWND hwnd, HFONT font, int min_height) {
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

void SetEditMargins(HWND hwnd, int left, int right) {
  if (!hwnd) {
    return;
  }
  SendMessageW(hwnd, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELONG(left, right));
}

void SetEditVerticalRect(HWND hwnd, HFONT font, int min_pad, int left_pad, int right_pad) {
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

void DrawToolbarButtonBackground(HDC hdc, const RECT& rect, COLORREF fill, COLORREF border) {
  if (!hdc) {
    return;
  }
  RECT draw = rect;
  InflateRect(&draw, -1, -1);
  HBRUSH brush = GetCachedBrush(fill);
  HPEN pen = GetCachedPen(border);
  HGDIOBJ old_brush = SelectObject(hdc, brush);
  HGDIOBJ old_pen = SelectObject(hdc, pen);
  RoundRect(hdc, draw.left, draw.top, draw.right, draw.bottom, 4, 4);
  SelectObject(hdc, old_pen);
  SelectObject(hdc, old_brush);
}

RegistryNode MakeChildNode(const RegistryNode& parent, const std::wstring& name) {
  RegistryNode child = parent;
  if (child.subkey.empty()) {
    child.subkey = name;
  } else {
    child.subkey = child.subkey + L"\\" + name;
  }
  return child;
}

std::wstring LeafName(const RegistryNode& node) {
  if (node.subkey.empty()) {
    return node.root_name.empty() ? registry_path::RootName(node.root) : node.root_name;
  }
  return registry_path::Leaf(node.subkey);
}

bool UseBinaryValueIcon(DWORD type) {
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

void UpdateLeafName(RegistryNode* node, const std::wstring& new_name) {
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

std::wstring FormatFileTime(const FILETIME& filetime) {
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

std::wstring FormatCommentDisplay(const std::wstring& text) {
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

uint64_t FileTimeToUint64(const FILETIME& filetime) {
  ULARGE_INTEGER value = {};
  value.LowPart = filetime.dwLowDateTime;
  value.HighPart = filetime.dwHighDateTime;
  return value.QuadPart;
}

int CompareTextInsensitive(const std::wstring& left, const std::wstring& right) {
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

int CompareUint64(uint64_t left, uint64_t right) {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
}

int CompareValueRow(const ListRow& left, const ListRow& right, int column) {
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

void SortValueRows(std::vector<ListRow>* rows, int column, bool ascending) {
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

void UpdateListViewSort(HWND list, int column, bool ascending) {
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
    item.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
    if (column >= 0 && GetListViewColumnSubItem(list, i) == column) {
      item.fmt |= ascending ? HDF_SORTUP : HDF_SORTDOWN;
    }
    Header_SetItem(header, i, &item);
  }
}

HBRUSH GetCachedBrush(COLORREF color) {
  struct Entry {
    COLORREF color = CLR_INVALID;
    HBRUSH brush = nullptr;
  };
  static Entry cache[4];
  for (auto& entry : cache) {
    if (entry.brush && entry.color == color) {
      return entry.brush;
    }
  }
  for (auto& entry : cache) {
    if (!entry.brush) {
      entry.color = color;
      entry.brush = CreateSolidBrush(color);
      return entry.brush;
    }
  }
  static size_t next = 0;
  if (cache[next].brush) {
    DeleteObject(cache[next].brush);
  }
  cache[next].color = color;
  cache[next].brush = CreateSolidBrush(color);
  HBRUSH result = cache[next].brush;
  next = (next + 1) % (sizeof(cache) / sizeof(cache[0]));
  return result;
}

HPEN GetCachedPen(COLORREF color, int width) {
  struct Entry {
    COLORREF color = CLR_INVALID;
    int width = 0;
    HPEN pen = nullptr;
  };
  static Entry cache[4];
  for (auto& entry : cache) {
    if (entry.pen && entry.color == color && entry.width == width) {
      return entry.pen;
    }
  }
  for (auto& entry : cache) {
    if (!entry.pen) {
      entry.color = color;
      entry.width = width;
      entry.pen = CreatePen(PS_SOLID, width, color);
      return entry.pen;
    }
  }
  static size_t next = 0;
  if (cache[next].pen) {
    DeleteObject(cache[next].pen);
  }
  cache[next].color = color;
  cache[next].width = width;
  cache[next].pen = CreatePen(PS_SOLID, width, color);
  HPEN result = cache[next].pen;
  next = (next + 1) % (sizeof(cache) / sizeof(cache[0]));
  return result;
}

UINT GetWindowDpi(HWND hwnd) {
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (user32) {
    auto get_dpi_for_window = reinterpret_cast<UINT(WINAPI*)(HWND)>(GetProcAddress(user32, "GetDpiForWindow"));
    if (get_dpi_for_window && hwnd) {
      return get_dpi_for_window(hwnd);
    }
    auto get_dpi_for_system = reinterpret_cast<UINT(WINAPI*)()>(GetProcAddress(user32, "GetDpiForSystem"));
    if (get_dpi_for_system) {
      return get_dpi_for_system();
    }
  }
  HDC hdc = GetDC(hwnd);
  int dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSX) : 96;
  if (hdc) {
    ReleaseDC(hwnd, hdc);
  }
  return dpi > 0 ? static_cast<UINT>(dpi) : 96;
}

HFONT CreateUIFont() {
  return static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
}

HFONT CreateIconFont(int point_size) {
  HDC hdc = GetDC(nullptr);
  int height = -MulDiv(point_size, GetDeviceCaps(hdc, LOGPIXELSY), 72);
  ReleaseDC(nullptr, hdc);
  return CreateFontW(height, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe MDL2 Assets");
}

int FontPointSize(const LOGFONTW& font) {
  HDC hdc = GetDC(nullptr);
  int size = 9;
  if (font.lfHeight != 0) {
    size = MulDiv(-font.lfHeight, 72, GetDeviceCaps(hdc, LOGPIXELSY));
  }
  ReleaseDC(nullptr, hdc);
  return size;
}

int FontHeightFromPointSize(int point_size) {
  HDC hdc = GetDC(nullptr);
  int height = -MulDiv(point_size, GetDeviceCaps(hdc, LOGPIXELSY), 72);
  ReleaseDC(nullptr, hdc);
  return height;
}

void ApplyFont(HWND hwnd, HFONT font) {
  if (hwnd && font) {
    SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
  }
}

HTREEITEM FindChildByText(HWND tree, HTREEITEM parent, const std::wstring& text) {
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

} // namespace

class RegistryAddressEnum : public ::IEnumString, public ::IACList {
public:
  RegistryAddressEnum(MainWindow* owner, HWND edit) : owner_(owner), edit_(edit) {}

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** out) override {
    if (!out) {
      return E_POINTER;
    }
    *out = nullptr;
    if (riid == IID_IUnknown || riid == IID_IEnumString) {
      *out = static_cast<::IEnumString*>(this);
      AddRef();
      return S_OK;
    }
    if (riid == IID_IACList) {
      *out = static_cast<::IACList*>(this);
      AddRef();
      return S_OK;
    }
    return E_NOINTERFACE;
  }

  ULONG STDMETHODCALLTYPE AddRef() override { return static_cast<ULONG>(InterlockedIncrement(&ref_count_)); }

  ULONG STDMETHODCALLTYPE Release() override {
    ULONG count = static_cast<ULONG>(InterlockedDecrement(&ref_count_));
    if (count == 0) {
      delete this;
    }
    return count;
  }

  HRESULT STDMETHODCALLTYPE Next(ULONG celt, LPOLESTR* rgelt, ULONG* pceltFetched) override {
    if (!rgelt) {
      return E_POINTER;
    }
    if (celt > 1 && !pceltFetched) {
      return E_POINTER;
    }
    UpdateSuggestionsIfNeeded();
    ULONG fetched = 0;
    for (; fetched < celt && index_ < suggestions_.size(); ++fetched, ++index_) {
      const std::wstring& item = suggestions_[index_];
      size_t bytes = (item.size() + 1) * sizeof(wchar_t);
      wchar_t* buffer = static_cast<wchar_t*>(CoTaskMemAlloc(bytes));
      if (!buffer) {
        for (ULONG i = 0; i < fetched; ++i) {
          CoTaskMemFree(rgelt[i]);
        }
        if (pceltFetched) {
          *pceltFetched = 0;
        }
        return E_OUTOFMEMORY;
      }
      wcscpy_s(buffer, item.size() + 1, item.c_str());
      rgelt[fetched] = buffer;
    }
    if (pceltFetched) {
      *pceltFetched = fetched;
    }
    return fetched == celt ? S_OK : S_FALSE;
  }

  HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override {
    UpdateSuggestionsIfNeeded();
    if (index_ + celt >= suggestions_.size()) {
      index_ = suggestions_.size();
      return S_FALSE;
    }
    index_ += celt;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE Reset() override {
    UpdateSuggestionsIfNeeded();
    index_ = 0;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE Clone(IEnumString** out) override {
    if (!out) {
      return E_POINTER;
    }
    auto* clone = new RegistryAddressEnum(owner_, edit_);
    clone->suggestions_ = suggestions_;
    clone->index_ = index_;
    clone->last_text_ = last_text_;
    clone->query_override_ = query_override_;
    *out = clone;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE Expand(PCWSTR text) noexcept override {
    if (!text) {
      query_override_.clear();
      return S_OK;
    }
    query_override_ = text;
    suggestions_.clear();
    index_ = 0;
    last_text_.clear();
    return S_OK;
  }

private:
  std::wstring ReadEditText() const {
    if (!edit_) {
      return L"";
    }
    int length = GetWindowTextLengthW(edit_);
    if (length <= 0) {
      return L"";
    }
    std::wstring text(static_cast<size_t>(length) + 1, L'\0');
    GetWindowTextW(edit_, MutableData(text), length + 1);
    text.resize(static_cast<size_t>(length));
    return text;
  }

  void UpdateSuggestionsIfNeeded() {
    if (!owner_) {
      suggestions_.clear();
      index_ = 0;
      last_text_.clear();
      return;
    }
    std::wstring query = query_override_.empty() ? ReadEditText() : query_override_;
    if (query_override_.empty() && edit_ && !query.empty()) {
      DWORD sel_start = 0;
      DWORD sel_end = 0;
      SendMessageW(edit_, EM_GETSEL, reinterpret_cast<WPARAM>(&sel_start), reinterpret_cast<LPARAM>(&sel_end));
      if (sel_end > sel_start && sel_end == query.size()) {
        query = query.substr(0, sel_start);
      }
    }
    if (query == last_text_) {
      return;
    }
    last_text_ = query;
    suggestions_ = owner_->BuildAddressSuggestions(query);
    index_ = 0;
  }

  ~RegistryAddressEnum() = default;

  LONG ref_count_ = 1;
  MainWindow* owner_ = nullptr;
  HWND edit_ = nullptr;
  std::vector<std::wstring> suggestions_;
  size_t index_ = 0;
  std::wstring last_text_;
  std::wstring query_override_;
};

struct AutoCompleteThemeContext {
  HWND owner = nullptr;
  const Theme* theme = nullptr;
};

LRESULT CALLBACK AutoCompleteListBoxSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, AutoCompleteListBoxSubclassProc, kAutoCompleteListBoxSubclassId);
    break;
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK AutoCompletePopupSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR) {
  switch (msg) {
  case WM_NOTIFY: {
    auto* header = reinterpret_cast<NMHDR*>(lparam);
    if (header && header->code == NM_CUSTOMDRAW && WindowClassEquals(header->hwndFrom, WC_LISTVIEWW)) {
      auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);
      const Theme& theme = Theme::Current();
      switch (draw->nmcd.dwDrawStage) {
      case CDDS_PREPAINT:
        return CDRF_NOTIFYITEMDRAW;
      case CDDS_ITEMPREPAINT: {
        if (draw->nmcd.uItemState & CDIS_SELECTED) {
          return CDRF_DODEFAULT;
        }
        COLORREF text = theme.TextColor();
        COLORREF background = theme.SurfaceColor();
        if (draw->nmcd.uItemState & CDIS_HOT) {
          background = theme.HoverColor();
        }
        draw->clrText = text;
        draw->clrTextBk = background;
        return CDRF_NEWFONT;
      }
      default:
        break;
      }
    }
    break;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().SurfaceBrush());
    return TRUE;
  }
  case WM_CTLCOLORLISTBOX:
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    int type = CTLCOLOR_STATIC;
    if (msg == WM_CTLCOLOREDIT) {
      type = CTLCOLOR_EDIT;
    } else if (msg == WM_CTLCOLORLISTBOX) {
      type = CTLCOLOR_LISTBOX;
    }
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, type));
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

BOOL CALLBACK ApplyAutoCompleteThemeProc(HWND hwnd, LPARAM lparam) {
  auto* ctx = reinterpret_cast<AutoCompleteThemeContext*>(lparam);
  if (!ctx || !ctx->theme) {
    return TRUE;
  }
  DWORD pid = 0;
  GetWindowThreadProcessId(hwnd, &pid);
  if (pid != GetCurrentProcessId()) {
    return TRUE;
  }

  bool is_dropdown = WindowClassEquals(hwnd, L"Auto-Suggest Dropdown") || WindowClassEquals(hwnd, L"Autocomplete") || WindowClassEquals(hwnd, L"AutoComplete");
  if (!is_dropdown) {
    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    if ((style & WS_POPUP) == 0) {
      return TRUE;
    }
    bool has_list_child = false;
    EnumChildWindows(
        hwnd,
        [](HWND child, LPARAM param) -> BOOL {
          auto* found = reinterpret_cast<bool*>(param);
          if (!found || *found) {
            return TRUE;
          }
          if (WindowClassEquals(child, WC_LISTVIEWW) || WindowClassEquals(child, WC_LISTBOXW)) {
            *found = true;
          }
          return TRUE;
        },
        reinterpret_cast<LPARAM>(&has_list_child));
    if (!has_list_child) {
      return TRUE;
    }
  }

  ctx->theme->ApplyToWindow(hwnd);
  if (!GetWindowSubclass(hwnd, AutoCompletePopupSubclassProc, kAutoCompletePopupSubclassId, nullptr)) {
    SetWindowSubclass(hwnd, AutoCompletePopupSubclassProc, kAutoCompletePopupSubclassId, 0);
  }
  EnumChildWindows(
      hwnd,
      [](HWND child, LPARAM param) -> BOOL {
        auto* theme = reinterpret_cast<const Theme*>(param);
        if (!theme) {
          return TRUE;
        }
        if (WindowClassEquals(child, WC_LISTVIEWW)) {
          theme->ApplyToListView(child);
        } else if (WindowClassEquals(child, WC_LISTBOXW) || WindowClassEquals(child, L"ComboLBox")) {
          AllowDarkModeForWindow(child, Theme::UseDarkMode());
          const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
          SetWindowTheme(child, theme_name, nullptr);
          if (!GetWindowSubclass(child, AutoCompleteListBoxSubclassProc, kAutoCompleteListBoxSubclassId, nullptr)) {
            SetWindowSubclass(child, AutoCompleteListBoxSubclassProc, kAutoCompleteListBoxSubclassId, 0);
          }
        } else {
          AllowDarkModeForWindow(child, Theme::UseDarkMode());
          const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
          SetWindowTheme(child, theme_name, nullptr);
        }
        return TRUE;
      },
      reinterpret_cast<LPARAM>(ctx->theme));
  InvalidateRect(hwnd, nullptr, TRUE);
  return TRUE;
}

MainWindow::~MainWindow() = default;

bool MainWindow::Create(HINSTANCE instance) {
  instance_ = instance;
  last_search_.criteria.search_keys = false;

  WNDCLASSEXW wc = {};
  wc.cbSize = sizeof(wc);
  wc.lpfnWndProc = MainWindow::WndProc;
  wc.hInstance = instance;
  wc.lpszClassName = kRegeditWindowClassName;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
  wc.hIconSm = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
  wc.hbrBackground = nullptr;

  RegisterClassExW(&wc);

  std::wstring title = L"RegKit V";
  title.append(REGKIT_VERSION_STR_W);
  if (util::IsProcessTrustedInstaller()) {
    title.append(L" - [TrustedInstaller]");
  } else if (util::IsProcessSystem()) {
    title.append(L" - [SYSTEM]");
  } else if (util::IsProcessElevated()) {
    title.append(L" - [Administrator]");
  }
  hwnd_ = CreateWindowExW(0, wc.lpszClassName, title.c_str(), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, CW_USEDEFAULT, CW_USEDEFAULT, 1200, 800, nullptr, nullptr, instance, this);
  if (hwnd_) {
    SetPropW(hwnd_, kRegKitWindowProperty, reinterpret_cast<HANDLE>(static_cast<INT_PTR>(1)));
    DragAcceptFiles(hwnd_, TRUE);
  }
  return hwnd_ != nullptr;
}

void MainWindow::ActivateRegeditCompatibilityMode() {
  bool was_enabled = regedit_compatibility_mode_;
  regedit_compatibility_mode_ = true;
  EnsureRegeditCompatControls();
  if (!was_enabled) {
    RefreshVisibleRegistryTreeLayout(true);
  }
  FocusAddressBarForExternalJump(true);
}

void MainWindow::Show(int cmd_show) {
  int show_cmd = cmd_show;
  if (window_placement_loaded_ && window_width_ > 0 && window_height_ > 0) {
    show_cmd = window_maximized_ ? SW_MAXIMIZE : SW_SHOWNORMAL;
  } else if (window_placement_loaded_ && window_maximized_) {
    show_cmd = SW_MAXIMIZE;
  }
  ShowWindow(hwnd_, show_cmd);
  UpdateWindow(hwnd_);
  if (regedit_compatibility_mode_) {
    EnsureRegeditCompatControls();
    FocusAddressBarForExternalJump(true);
  }
  PostMessageW(hwnd_, kDeferredStartupMessage, 0, 0);
  PostMessageW(hwnd_, kLoadTracesMessage, 0, 0);
  PostMessageW(hwnd_, kLoadDefaultsMessage, 0, 0);
}

void MainWindow::QueueExternalJump(const std::wstring& target) {
  queued_external_jump_target_ = target;
}

void MainWindow::CreateRegeditCompatControls() {
  if (regedit_compat_edit_ || regedit_compat_tree_ || regedit_compat_list_) {
    return;
  }
  DWORD base_style = WS_CHILD | WS_TABSTOP;
  regedit_compat_edit_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", base_style | ES_AUTOHSCROLL, -2000, -2000, 8, 8, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRegeditCompatEditId)), instance_, nullptr);
  regedit_compat_tree_ = CreateWindowExW(WS_EX_CLIENTEDGE, WC_TREEVIEWW, L"", base_style | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS, -2000, -2000, 8, 8, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRegeditCompatTreeId)), instance_, nullptr);
  regedit_compat_list_ = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"", base_style | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SINGLESEL, -2000, -2000, 8, 8, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRegeditCompatListId)), instance_, nullptr);
  if (regedit_compat_edit_) {
    SetWindowSubclass(regedit_compat_edit_, AddressEditProc, kAddressSubclassId, reinterpret_cast<DWORD_PTR>(this));
  }
  if (regedit_compat_list_) {
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH;
    col.cx = 240;
    col.pszText = const_cast<wchar_t*>(L"Name");
    ListView_InsertColumn(regedit_compat_list_, 0, &col);
  }
  PopulateRegeditCompatTree();
}

void MainWindow::EnsureRegeditCompatControls() {
  if (!hwnd_ || (regedit_compat_edit_ && regedit_compat_tree_ && regedit_compat_list_)) {
    return;
  }
  CreateRegeditCompatControls();
  ApplyFont(regedit_compat_edit_, ui_font_);
  ApplyFont(regedit_compat_tree_, ui_font_);
  ApplyFont(regedit_compat_list_, ui_font_);
  const Theme& theme = Theme::Current();
  theme.ApplyToTreeView(regedit_compat_tree_);
  theme.ApplyToListView(regedit_compat_list_);
  if (regedit_compat_edit_) {
    SetWindowTheme(regedit_compat_edit_, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(regedit_compat_edit_, 6, 6);
    SetEditVerticalRect(regedit_compat_edit_, ui_font_, 2, 6, 6);
  }
}

void MainWindow::PopulateRegeditCompatTree() {
  if (!regedit_compat_tree_) {
    return;
  }
  TreeView_DeleteAllItems(regedit_compat_tree_);
  regedit_compat_nodes_.clear();

  auto add_node = [&](HTREEITEM parent, const std::wstring& text, const std::wstring& path, bool has_placeholder) {
    auto node = std::make_unique<RegeditCompatTreeNode>();
    node->path = path;
    RegeditCompatTreeNode* raw = node.get();
    regedit_compat_nodes_.push_back(std::move(node));
    TVINSERTSTRUCTW insert = {};
    insert.hParent = parent;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM;
    insert.item.pszText = const_cast<wchar_t*>(text.c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(raw);
    HTREEITEM item = TreeView_InsertItem(regedit_compat_tree_, &insert);
    if (item && has_placeholder) {
      TVINSERTSTRUCTW placeholder = {};
      placeholder.hParent = item;
      placeholder.hInsertAfter = TVI_LAST;
      placeholder.item.mask = TVIF_TEXT;
      placeholder.item.pszText = const_cast<wchar_t*>(L"");
      TreeView_InsertItem(regedit_compat_tree_, &placeholder);
    }
    return item;
  };

  HTREEITEM computer = add_node(TVI_ROOT, L"Computer", L"Computer", true);
  static const wchar_t* kCompatHives[] = {L"HKEY_CLASSES_ROOT", L"HKEY_CURRENT_USER", L"HKEY_LOCAL_MACHINE", L"HKEY_USERS", L"HKEY_CURRENT_CONFIG"};
  for (const wchar_t* hive : kCompatHives) {
    add_node(computer, hive, hive, true);
  }
  TreeView_Expand(regedit_compat_tree_, computer, TVE_EXPAND);
  regedit_compat_selected_key_path_.clear();
  if (regedit_compat_edit_) {
    SetWindowTextW(regedit_compat_edit_, L"");
  }
}

void MainWindow::PopulateRegeditCompatChildren(HTREEITEM item) {
  if (!regedit_compat_tree_ || !item) {
    return;
  }
  TVITEMW tvi = {};
  tvi.mask = TVIF_PARAM;
  tvi.hItem = item;
  if (!TreeView_GetItem(regedit_compat_tree_, &tvi)) {
    return;
  }
  auto* node = reinterpret_cast<RegeditCompatTreeNode*>(tvi.lParam);
  if (!node || node->populated) {
    return;
  }
  node->populated = true;

  while (HTREEITEM child = TreeView_GetChild(regedit_compat_tree_, item)) {
    TreeView_DeleteItem(regedit_compat_tree_, child);
  }

  if (EqualsInsensitive(node->path, L"Computer")) {
    static const wchar_t* kCompatHives[] = {L"HKEY_CLASSES_ROOT", L"HKEY_CURRENT_USER", L"HKEY_LOCAL_MACHINE", L"HKEY_USERS", L"HKEY_CURRENT_CONFIG"};
    for (const wchar_t* hive : kCompatHives) {
      auto child_node = std::make_unique<RegeditCompatTreeNode>();
      child_node->path = hive;
      RegeditCompatTreeNode* raw = child_node.get();
      regedit_compat_nodes_.push_back(std::move(child_node));
      TVINSERTSTRUCTW insert = {};
      insert.hParent = item;
      insert.hInsertAfter = TVI_LAST;
      insert.item.mask = TVIF_TEXT | TVIF_PARAM;
      insert.item.pszText = const_cast<wchar_t*>(hive);
      insert.item.lParam = reinterpret_cast<LPARAM>(raw);
      HTREEITEM child_item = TreeView_InsertItem(regedit_compat_tree_, &insert);
      if (child_item) {
        TVINSERTSTRUCTW placeholder = {};
        placeholder.hParent = child_item;
        placeholder.hInsertAfter = TVI_LAST;
        placeholder.item.mask = TVIF_TEXT;
        placeholder.item.pszText = const_cast<wchar_t*>(L"");
        TreeView_InsertItem(regedit_compat_tree_, &placeholder);
      }
    }
    return;
  }

  RegistryNode registry_node;
  if (!ResolvePathToNode(node->path, &registry_node)) {
    return;
  }
  auto subkeys = RegistryProvider::EnumSubKeyNames(registry_node, true);
  for (const auto& subkey : subkeys) {
    auto child_node = std::make_unique<RegeditCompatTreeNode>();
    child_node->path = node->path + L"\\" + subkey;
    RegeditCompatTreeNode* raw = child_node.get();
    regedit_compat_nodes_.push_back(std::move(child_node));
    TVINSERTSTRUCTW insert = {};
    insert.hParent = item;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM;
    insert.item.pszText = const_cast<wchar_t*>(subkey.c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(raw);
    HTREEITEM child_item = TreeView_InsertItem(regedit_compat_tree_, &insert);
    if (child_item) {
      TVINSERTSTRUCTW placeholder = {};
      placeholder.hParent = child_item;
      placeholder.hInsertAfter = TVI_LAST;
      placeholder.item.mask = TVIF_TEXT;
      placeholder.item.pszText = const_cast<wchar_t*>(L"");
      TreeView_InsertItem(regedit_compat_tree_, &placeholder);
    }
  }
}

void MainWindow::UpdateRegeditCompatList(const std::wstring& key_path) {
  if (!regedit_compat_list_) {
    return;
  }
  regedit_compat_list_names_.clear();
  ListView_DeleteAllItems(regedit_compat_list_);

  RegistryNode node;
  if (!ResolvePathToNode(key_path, &node)) {
    return;
  }
  RegistryProvider::KeyEnumResult enum_result;
  bool names_reserved = false;
  int index = 0;
  RegistryProvider::EnumKeyStreaming(
      node, true, false, false, &enum_result,
      [&](const ValueInfo& value, const BYTE*, DWORD) {
        if (!names_reserved) {
          if (enum_result.info_valid) {
            regedit_compat_list_names_.reserve(enum_result.info.value_count);
          }
          names_reserved = true;
        }
        std::wstring display_name = value.name.empty() ? L"(Default)" : value.name;
        regedit_compat_list_names_.push_back(value.name);
        LVITEMW item = {};
        item.mask = LVIF_TEXT;
        item.iItem = index++;
        item.pszText = display_name.data();
        ListView_InsertItem(regedit_compat_list_, &item);
        return true;
      },
      {});
}

HTREEITEM MainWindow::FindRegeditCompatItemByPath(const std::wstring& key_path) {
  if (!regedit_compat_tree_) {
    return nullptr;
  }
  HTREEITEM current = TreeView_GetRoot(regedit_compat_tree_);
  if (!current) {
    return nullptr;
  }
  std::vector<std::wstring> parts = registry_path::Split(key_path);
  if (!parts.empty() && EqualsInsensitive(parts.front(), L"Computer")) {
    parts.erase(parts.begin());
  }
  if (parts.empty()) {
    return current;
  }
  for (const auto& part : parts) {
    PopulateRegeditCompatChildren(current);
    TreeView_Expand(regedit_compat_tree_, current, TVE_EXPAND);
    HTREEITEM child = FindChildByText(regedit_compat_tree_, current, part);
    if (!child) {
      return nullptr;
    }
    current = child;
  }
  return current;
}

void MainWindow::SyncRegeditCompatControls(const std::wstring& key_path, const std::wstring& value_name) {
  if (!regedit_compat_edit_ || !regedit_compat_tree_ || !regedit_compat_list_) {
    return;
  }
  syncing_regedit_compat_controls_ = true;
  auto clear_sync = [&]() {
    syncing_regedit_compat_controls_ = false;
  };

  regedit_compat_selected_key_path_ = key_path;
  if (key_path.empty()) {
    SetWindowTextW(regedit_compat_edit_, L"Computer");
    HTREEITEM root = TreeView_GetRoot(regedit_compat_tree_);
    if (root) {
      TreeView_SelectItem(regedit_compat_tree_, root);
    }
    regedit_compat_list_names_.clear();
    ListView_DeleteAllItems(regedit_compat_list_);
    clear_sync();
    return;
  }

  SetWindowTextW(regedit_compat_edit_, key_path.c_str());
  HTREEITEM item = FindRegeditCompatItemByPath(key_path);
  if (item) {
    TreeView_SelectItem(regedit_compat_tree_, item);
    TreeView_EnsureVisible(regedit_compat_tree_, item);
  }
  UpdateRegeditCompatList(key_path);
  if (!value_name.empty()) {
    for (size_t i = 0; i < regedit_compat_list_names_.size(); ++i) {
      if (regedit_compat_list_names_[i] == value_name) {
        ListView_SetItemState(regedit_compat_list_, static_cast<int>(i), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(regedit_compat_list_, static_cast<int>(i), FALSE);
        break;
      }
    }
  }
  clear_sync();
}

void MainWindow::HandleRegeditCompatTreeSelection(NMTREEVIEWW* info) {
  if (syncing_regedit_compat_controls_) {
    return;
  }
  if (!info) {
    return;
  }
  auto* node = reinterpret_cast<RegeditCompatTreeNode*>(info->itemNew.lParam);
  if (!node) {
    return;
  }
  if (EqualsInsensitive(node->path, L"Computer")) {
    regedit_compat_selected_key_path_.clear();
    regedit_compat_pending_key_path_.clear();
    regedit_compat_pending_value_name_.clear();
    KillTimer(hwnd_, kRegeditCompatApplyTimerId);
    if (regedit_compat_edit_) {
      SetWindowTextW(regedit_compat_edit_, L"Computer");
    }
    return;
  }
  regedit_compat_selected_key_path_ = node->path;
  regedit_compat_pending_value_name_.clear();
  QueueRegeditCompatNavigation(node->path);
  if (regedit_compat_edit_) {
    SetWindowTextW(regedit_compat_edit_, node->path.c_str());
  }
  UpdateRegeditCompatList(node->path);
}

void MainWindow::HandleRegeditCompatListSelection(NMLISTVIEW* info) {
  if (syncing_regedit_compat_controls_) {
    return;
  }
  if (!info || info->iItem < 0 || (info->uNewState & LVIS_SELECTED) == 0) {
    return;
  }
  if (static_cast<size_t>(info->iItem) >= regedit_compat_list_names_.size()) {
    return;
  }
  regedit_compat_pending_value_name_ = regedit_compat_list_names_[static_cast<size_t>(info->iItem)];
  if (!regedit_compat_selected_key_path_.empty()) {
    QueueRegeditCompatNavigation(regedit_compat_selected_key_path_);
  }
}

void MainWindow::QueueRegeditCompatNavigation(const std::wstring& key_path) {
  if (!hwnd_) {
    return;
  }
  regedit_compat_pending_key_path_ = key_path;
  KillTimer(hwnd_, kRegeditCompatApplyTimerId);
  SetTimer(hwnd_, kRegeditCompatApplyTimerId, kRegeditCompatApplyDelayMs, nullptr);
}

void MainWindow::BeginJumpUiBatch() {
  if (jump_ui_batch_active_) {
    return;
  }
  jump_ui_batch_active_ = true;
  if (tree_.hwnd()) {
    SendMessageW(tree_.hwnd(), WM_SETREDRAW, FALSE, 0);
  }
  if (value_list_.hwnd()) {
    SendMessageW(value_list_.hwnd(), WM_SETREDRAW, FALSE, 0);
  }
}

void MainWindow::EndJumpUiBatch() {
  if (!jump_ui_batch_active_) {
    return;
  }
  jump_ui_batch_active_ = false;
  if (tree_.hwnd()) {
    SendMessageW(tree_.hwnd(), WM_SETREDRAW, TRUE, 0);
    InvalidateRect(tree_.hwnd(), nullptr, TRUE);
  }
  if (value_list_.hwnd()) {
    SendMessageW(value_list_.hwnd(), WM_SETREDRAW, TRUE, 0);
    InvalidateRect(value_list_.hwnd(), nullptr, TRUE);
  }
}

void MainWindow::ApplyTreeSelectionEffects(RegistryNode* node) {
  UpdateAddressBar(node);
  UpdateValueListForNode(node);
  MarkTreeStateDirty();
}

void MainWindow::ApplyPendingRegeditCompatNavigation() {
  std::wstring key_path = std::move(regedit_compat_pending_key_path_);
  std::wstring value_name = std::move(regedit_compat_pending_value_name_);
  regedit_compat_pending_key_path_.clear();
  regedit_compat_pending_value_name_.clear();
  if (key_path.empty()) {
    return;
  }
  std::wstring resolved_key_path;
  std::wstring resolved_value_name;
  if (!ResolveExternalJumpTarget(key_path, &resolved_key_path, &resolved_value_name)) {
    return;
  }
  if (value_name.empty()) {
    value_name = resolved_value_name;
  }
  if (!NavigateToResolvedExternalJump(resolved_key_path, value_name, false)) {
    return;
  }
  if (value_name.empty()) {
    return;
  }
  if (!SelectValueByName(value_name) && current_node_) {
    pending_external_value_key_path_ = resolved_key_path;
    pending_external_value_name_ = value_name;
    if (!value_list_loading_) {
      UpdateValueListForNode(current_node_);
    }
  }
}

void MainWindow::FocusAddressBarForExternalJump(bool defer_if_needed) {
  if (!hwnd_) {
    return;
  }
  HWND target = nullptr;
  if (regedit_compatibility_mode_) {
    if (regedit_compat_edit_) {
      target = regedit_compat_edit_;
    } else if (tree_.hwnd()) {
      target = tree_.hwnd();
    }
  } else if (address_edit_) {
    target = address_edit_;
  }
  if (!target) {
    return;
  }
  if (IsWindowVisible(hwnd_) && !IsIconic(hwnd_)) {
    SetFocus(target);
    if (target == address_edit_ || target == regedit_compat_edit_) {
      SendMessageW(target, EM_SETSEL, 0, -1);
    }
    return;
  }
  if (defer_if_needed) {
    PostMessageW(hwnd_, kFocusAddressBarMessage, 0, 0);
  }
}

bool MainWindow::TranslateAccelerator(const MSG& msg) {
  if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN) {
    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
    HWND focus = GetFocus();
    auto is_text_input = [](HWND hwnd) -> bool {
      if (!hwnd) {
        return false;
      }
      wchar_t cls[64] = {};
      GetClassNameW(hwnd, cls, static_cast<int>(_countof(cls)));
      if (_wcsicmp(cls, L"Edit") == 0) {
        return true;
      }
      if (_wcsicmp(cls, L"RichEdit20W") == 0 || _wcsicmp(cls, L"RichEdit20A") == 0) {
        return true;
      }
      if (_wcsicmp(cls, L"ComboBox") == 0 || _wcsicmp(cls, L"ComboBoxEx32") == 0) {
        return true;
      }
      HWND parent = GetParent(hwnd);
      if (parent) {
        GetClassNameW(parent, cls, static_cast<int>(_countof(cls)));
        if (_wcsicmp(cls, L"ComboBox") == 0 || _wcsicmp(cls, L"ComboBoxEx32") == 0) {
          return true;
        }
      }
      return false;
    };
    const bool focus_edit = is_text_input(focus);

    if ((alt && msg.wParam == 'D') || (ctrl && !alt && msg.wParam == 'L')) {
      if (address_edit_) {
        SetFocus(address_edit_);
        SendMessageW(address_edit_, EM_SETSEL, 0, -1);
        return true;
      }
    }

    if (ctrl && !alt) {
      if (shift && msg.wParam == 'C' && !focus_edit) {
        HandleMenuCommand(cmd::kEditCopyKey);
        return true;
      }
      switch (msg.wParam) {
      case 'A':
        if (SelectAllInFocusedList()) {
          return true;
        }
        if (focus_edit && focus) {
          SendMessageW(focus, EM_SETSEL, 0, -1);
          return true;
        }
        break;
      case 'C':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditCopy);
          return true;
        }
        return false;
      case 'V':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditPaste);
          return true;
        }
        return false;
      case 'X':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditDelete);
          return true;
        }
        return false;
      case 'Z':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditUndo);
          return true;
        }
        return false;
      case 'Y':
        if (focus_edit && focus) {
          SendMessageW(focus, EM_REDO, 0, 0);
          return true;
        }
        HandleMenuCommand(cmd::kEditRedo);
        return true;
      case 'F':
        HandleMenuCommand(cmd::kEditFind);
        return true;
      case 'G':
        HandleMenuCommand(cmd::kEditGoTo);
        return true;
      case 'H':
        HandleMenuCommand(cmd::kEditReplace);
        return true;
      case 'S':
        HandleMenuCommand(cmd::kFileSave);
        return true;
      case 'E':
        HandleMenuCommand(cmd::kFileExport);
        return true;
      case 'N':
        OpenLocalRegistryTab();
        return true;
      }
    }

    if (!ctrl && !alt) {
      if (msg.wParam == VK_DELETE && !focus_edit) {
        HandleMenuCommand(cmd::kEditDelete);
        return true;
      }
      if (msg.wParam == VK_F2 && !focus_edit) {
        HandleMenuCommand(cmd::kEditRename);
        return true;
      }
      if (msg.wParam == VK_F5) {
        HandleMenuCommand(cmd::kViewRefresh);
        return true;
      }
    }

    if (focus_edit) {
      if (msg.wParam == VK_DELETE || msg.wParam == VK_BACK) {
        return false;
      }
      if (ctrl && !alt) {
        switch (msg.wParam) {
        case 'C':
        case 'V':
        case 'X':
        case 'Z':
        case 'Y':
          return false;
        default:
          break;
        }
      }
    }
  }
  if (accelerators_) {
    return ::TranslateAcceleratorW(hwnd_, accelerators_, const_cast<MSG*>(&msg)) != 0;
  }
  return false;
}

LRESULT CALLBACK MainWindow::WndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
  if (message == WM_NCCREATE) {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    auto* self = static_cast<MainWindow*>(create->lpCreateParams);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    self->hwnd_ = hwnd;
  }

  auto* self = reinterpret_cast<MainWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  if (!self) {
    return DefWindowProcW(hwnd, message, wparam, lparam);
  }
  return self->HandleMessage(message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::AddressEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow*>(ref_data);
  if (message == WM_SETTEXT) {
    const wchar_t* text = reinterpret_cast<const wchar_t*>(lparam);
    if (self && hwnd == self->regedit_compat_edit_ && text && *text) {
      if (self->syncing_regedit_compat_controls_) {
        return DefSubclassProc(hwnd, message, wparam, lparam);
      }
      if (_wcsicmp(text, L"Computer") == 0) {
        self->regedit_compat_selected_key_path_.clear();
        self->regedit_compat_pending_key_path_.clear();
        self->regedit_compat_pending_value_name_.clear();
        KillTimer(self->hwnd_, kRegeditCompatApplyTimerId);
      } else {
        self->regedit_compat_selected_key_path_ = text;
        self->regedit_compat_pending_value_name_.clear();
        self->QueueRegeditCompatNavigation(text);
      }
    }
  }
  if (message == WM_KEYDOWN && wparam == VK_RETURN) {
    if (self && self->hwnd_) {
      if (hwnd == self->regedit_compat_edit_) {
        wchar_t buffer[2048] = {};
        GetWindowTextW(hwnd, buffer, static_cast<int>(_countof(buffer)));
        self->regedit_compat_selected_key_path_ = buffer;
        self->regedit_compat_pending_value_name_.clear();
        self->regedit_compat_pending_key_path_ = buffer;
        KillTimer(self->hwnd_, kRegeditCompatApplyTimerId);
        self->ApplyPendingRegeditCompatNavigation();
      } else {
        SendMessageW(self->hwnd_, kAddressEnterMessage, 0, 0);
      }
    }
    return 0;
  }
  if (message == WM_CHAR && wparam == VK_RETURN) {
    return 0;
  }
  if (message == WM_SETFOCUS) {
    LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
    SendMessageW(hwnd, EM_SETSEL, 0, -1);
    return result;
  }
  if (message == WM_KEYUP) {
    LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
    auto* window = reinterpret_cast<MainWindow*>(ref_data);
    if (window) {
      window->ApplyAutoCompleteTheme();
    }
    return result;
  }
  if (message == WM_LBUTTONDOWN) {
    if (GetFocus() != hwnd) {
      LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
      SendMessageW(hwnd, EM_SETSEL, 0, -1);
      return result;
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::FilterEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR /*ref_data*/) {
  if (message == WM_KEYDOWN && wparam == VK_RETURN) {
    return 0;
  }
  if (message == WM_CHAR && wparam == VK_RETURN) {
    return 0;
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::TabProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow*>(ref_data);
  if (!self) {
    return DefSubclassProc(hwnd, message, wparam, lparam);
  }

  switch (message) {
  case WM_ERASEBKGND:
    return 1;
  case WM_MOUSEMOVE: {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    self->UpdateTabHotState(hwnd, pt);
    if (!self->tab_mouse_tracking_) {
      TRACKMOUSEEVENT tme = {};
      tme.cbSize = sizeof(tme);
      tme.dwFlags = TME_LEAVE;
      tme.hwndTrack = hwnd;
      TrackMouseEvent(&tme);
      self->tab_mouse_tracking_ = true;
    }
    return 0;
  }
  case WM_MOUSELEAVE:
    self->tab_mouse_tracking_ = false;
    if (self->tab_hot_index_ != -1 || self->tab_close_hot_index_ != -1) {
      self->tab_hot_index_ = -1;
      self->tab_close_hot_index_ = -1;
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    return 0;
  case WM_LBUTTONDOWN: {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    TCHITTESTINFO hit = {};
    hit.pt = pt;
    int index = TabCtrl_HitTest(hwnd, &hit);
    RECT close_rect = {};
    if (self->GetTabCloseRect(index, &close_rect) && PtInRect(&close_rect, pt)) {
      self->tab_close_down_index_ = index;
      SetCapture(hwnd);
      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }
    if (self->tab_close_down_index_ != -1) {
      self->tab_close_down_index_ = -1;
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    break;
  }
  case WM_LBUTTONUP: {
    if (self->tab_close_down_index_ >= 0) {
      POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      int close_index = self->tab_close_down_index_;
      self->tab_close_down_index_ = -1;
      ReleaseCapture();
      RECT close_rect = {};
      if (self->GetTabCloseRect(close_index, &close_rect) && PtInRect(&close_rect, pt)) {
        self->CloseTab(close_index);
        self->tab_hot_index_ = -1;
        self->tab_close_hot_index_ = -1;
      }
      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }
    break;
  }
  case WM_CAPTURECHANGED:
    if (self->tab_close_down_index_ >= 0) {
      self->tab_close_down_index_ = -1;
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    break;
  case WM_PAINT: {
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd, &ps);
    self->PaintTabControl(hwnd, hdc);
    EndPaint(hwnd, &ps);
    return 0;
  }
  default:
    break;
  }

  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::ListViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow*>(ref_data);
  if (message == WM_LBUTTONDOWN && self && hwnd == self->value_list_.hwnd()) {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    LVHITTESTINFO hit = {};
    hit.pt = pt;
    int index = ListView_HitTest(hwnd, &hit);
    DWORD now = GetTickCount();
    if (index >= 0 && index == self->last_value_click_index_) {
      self->last_value_click_delta_ = now - self->last_value_click_time_;
      self->last_value_click_delta_valid_ = true;
    } else {
      self->last_value_click_delta_valid_ = false;
    }
    self->last_value_click_time_ = now;
    self->last_value_click_index_ = index;
  }
  if (message == WM_KEYDOWN && self && hwnd == self->value_list_.hwnd()) {
    if (wparam == VK_RETURN) {
      self->value_activate_from_key_ = true;
      self->last_value_click_delta_valid_ = false;
    }
  }
  if (message == WM_CHAR && self && hwnd == self->value_list_.hwnd()) {
    wchar_t ch = static_cast<wchar_t>(wparam);
    if (ch == L'\b' || (iswprint(ch) && ch != L'\r' && ch != L'\n' && ch != L'\t')) {
      self->HandleTypeToSelectList(ch);
      return 0;
    }
  }
  if (message == WM_SETFOCUS || message == WM_KILLFOCUS) {
    SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
    if (self && hwnd == self->history_list_) {
      ListView_SetItemState(hwnd, -1, 0, LVIS_FOCUSED);
    }
  }
  if (message == WM_UPDATEUISTATE) {
    LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
    SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
    if (self && hwnd == self->history_list_) {
      ListView_SetItemState(hwnd, -1, 0, LVIS_FOCUSED);
    }
    return result;
  }
  if (message == WM_ERASEBKGND) {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().PanelBrush());
    return 1;
  }
  if (message == WM_CTLCOLOREDIT) {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    SetTextColor(hdc, Theme::Current().TextColor());
    SetBkColor(hdc, Theme::Current().PanelColor());
    return reinterpret_cast<LRESULT>(Theme::Current().PanelBrush());
  }
  if (message == WM_PRINTCLIENT) {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().PanelBrush());
  }
  if (message == WM_THEMECHANGED) {
    if (self) {
      InvalidateRect(hwnd, nullptr, TRUE);
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::TreeViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow*>(ref_data);
  if (message == WM_KEYDOWN && self && hwnd == self->tree_.hwnd()) {
    switch (wparam) {
    case VK_RIGHT:
      self->type_buffer_tree_.clear();
      self->tree_type_to_select_descend_ = true;
      break;
    case VK_LEFT:
    case VK_UP:
    case VK_DOWN:
    case VK_HOME:
    case VK_END:
    case VK_PRIOR:
    case VK_NEXT:
      self->type_buffer_tree_.clear();
      self->tree_type_to_select_descend_ = false;
      break;
    default:
      break;
    }
  }
  if (message == WM_CHAR && self && hwnd == self->tree_.hwnd()) {
    wchar_t ch = static_cast<wchar_t>(wparam);
    if (ch == L'\b' || (iswprint(ch) && ch != L'\r' && ch != L'\n' && ch != L'\t')) {
      self->HandleTypeToSelectTree(ch);
      return 0;
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::HeaderProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow*>(ref_data);
  if (message == WM_ERASEBKGND) {
    return 1;
  }
  if (message == WM_PAINT) {
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd, &ps);
    const Theme& theme = Theme::Current();
    RECT client = {};
    GetClientRect(hwnd, &client);
    FillRect(hdc, &client, theme.HeaderBrush());

    HFONT old_font = nullptr;
    if (self && self->ui_font_) {
      old_font = reinterpret_cast<HFONT>(SelectObject(hdc, self->ui_font_));
    }

    HTHEME header_theme = OpenThemeData(hwnd, VSCLASS_HEADER);
    SIZE arrow_size = {0, 0};
    if (header_theme) {
      GetThemePartSize(header_theme, hdc, HP_HEADERSORTARROW, HSAS_SORTEDUP, nullptr, TS_TRUE, &arrow_size);
    }
    if (arrow_size.cx <= 0 || arrow_size.cy <= 0) {
      arrow_size.cx = 8;
      arrow_size.cy = 8;
    }

    int count = Header_GetItemCount(hwnd);
    for (int i = 0; i < count; ++i) {
      RECT rect = {};
      if (!Header_GetItemRect(hwnd, i, &rect)) {
        continue;
      }

      wchar_t text[128] = {};
      HDITEMW item = {};
      item.mask = HDI_TEXT | HDI_FORMAT;
      item.pszText = text;
      item.cchTextMax = static_cast<int>(_countof(text));
      Header_GetItem(hwnd, i, &item);

      bool sorted_up = (item.fmt & HDF_SORTUP) != 0;
      bool sorted_down = (item.fmt & HDF_SORTDOWN) != 0;

      FillRect(hdc, &rect, theme.HeaderBrush());

      RECT text_rect = rect;
      text_rect.left += 8;
      text_rect.right -= 8;
      if (sorted_up || sorted_down) {
        text_rect.right -= arrow_size.cx + 6;
      }

      UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
      if (item.fmt & HDF_RIGHT) {
        format |= DT_RIGHT;
      } else if (item.fmt & HDF_CENTER) {
        format |= DT_CENTER;
      }

      SetBkMode(hdc, TRANSPARENT);
      SetTextColor(hdc, theme.TextColor());
      DrawTextW(hdc, text, -1, &text_rect, format);

      if ((sorted_up || sorted_down) && header_theme) {
        RECT arrow_rect = rect;
        arrow_rect.right -= 6;
        arrow_rect.left = arrow_rect.right - arrow_size.cx;
        arrow_rect.top = rect.top + (rect.bottom - rect.top - arrow_size.cy) / 2;
        arrow_rect.bottom = arrow_rect.top + arrow_size.cy;
        int arrow_state = sorted_up ? HSAS_SORTEDUP : HSAS_SORTEDDOWN;
        DrawThemeBackground(header_theme, hdc, HP_HEADERSORTARROW, arrow_state, &arrow_rect, nullptr);
      }
    }

    if (header_theme) {
      CloseThemeData(header_theme);
    }

    if (old_font) {
      SelectObject(hdc, old_font);
    }
    EndPaint(hwnd, &ps);
    return 0;
  }
  if (message == WM_THEMECHANGED) {
    InvalidateRect(hwnd, nullptr, TRUE);
  }
  if (message == WM_CONTEXTMENU) {
    if (self) {
      HWND value_header = ListView_GetHeader(self->value_list_.hwnd());
      HWND history_header = ListView_GetHeader(self->history_list_);
      HWND search_header = ListView_GetHeader(self->search_results_list_);
      if (hwnd == value_header) {
        POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        if (screen_pt.x == -1 && screen_pt.y == -1) {
          RECT rect = {};
          GetWindowRect(hwnd, &rect);
          screen_pt.x = rect.left + 12;
          screen_pt.y = rect.bottom - 4;
        }
        self->ShowValueHeaderMenu(screen_pt);
        return 0;
      }
      if (hwnd == history_header) {
        POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        if (screen_pt.x == -1 && screen_pt.y == -1) {
          RECT rect = {};
          GetWindowRect(hwnd, &rect);
          screen_pt.x = rect.left + 12;
          screen_pt.y = rect.bottom - 4;
        }
        self->ShowHistoryHeaderMenu(screen_pt);
        return 0;
      }
      if (hwnd == search_header) {
        POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        if (screen_pt.x == -1 && screen_pt.y == -1) {
          RECT rect = {};
          GetWindowRect(hwnd, &rect);
          screen_pt.x = rect.left + 12;
          screen_pt.y = rect.bottom - 4;
        }
        self->ShowSearchHeaderMenu(screen_pt);
        return 0;
      }
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT MainWindow::HandleMessage(UINT message, WPARAM wparam, LPARAM lparam) {
  switch (message) {
  case WM_CREATE:
    return OnCreate() ? 0 : -1;
  case WM_DESTROY:
    OnDestroy();
    PostQuitMessage(0);
    return 0;
  case WM_NCPAINT: {
    LRESULT result = DefWindowProcW(hwnd_, message, wparam, lparam);
    PaintMenuBarSeparator();
    return result;
  }
  case WM_NCACTIVATE: {
    LRESULT result = DefWindowProcW(hwnd_, message, wparam, lparam);
    PaintMenuBarSeparator();
    return result;
  }
  case WM_GETMINMAXINFO: {
    auto* info = reinterpret_cast<MINMAXINFO*>(lparam);
    if (info) {
      info->ptMinTrackSize.x = std::max<LONG>(info->ptMinTrackSize.x, 400);
      info->ptMinTrackSize.y = std::max<LONG>(info->ptMinTrackSize.y, 200);
    }
    return 0;
  }
  case WM_SIZE:
    OnSize(LOWORD(lparam), HIWORD(lparam));
    return 0;
  case WM_DPICHANGED: {
    const RECT* suggested = reinterpret_cast<const RECT*>(lparam);
    if (suggested) {
      SetWindowPos(hwnd_, nullptr, suggested->left, suggested->top, suggested->right - suggested->left, suggested->bottom - suggested->top, SWP_NOZORDER | SWP_NOACTIVATE);
    }
    UpdateUIFont();
    ReloadThemeIcons();
    return 0;
  }
  case WM_DPICHANGED_AFTERPARENT:
    UpdateUIFont();
    ReloadThemeIcons();
    return 0;
  case WM_ACTIVATE:
    if (LOWORD(wparam) == WA_ACTIVE && tree_.hwnd()) {
      int sel = TabCtrl_GetCurSel(tab_);
      if (sel >= 0 && static_cast<size_t>(sel) < tabs_.size() && tabs_[static_cast<size_t>(sel)].kind == TabEntry::Kind::kRegistry) {
        if (regedit_compatibility_mode_ && regedit_compat_edit_) {
          SetFocus(regedit_compat_edit_);
          SendMessageW(regedit_compat_edit_, EM_SETSEL, 0, -1);
        } else {
          SetFocus(tree_.hwnd());
        }
      }
    }
    break;
  case WM_DROPFILES: {
    HDROP drop = reinterpret_cast<HDROP>(wparam);
    UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
    std::vector<std::wstring> reg_paths;
    std::wstring offline_candidate;
    for (UINT index = 0; index < count; ++index) {
      UINT path_len = DragQueryFileW(drop, index, nullptr, 0);
      if (path_len == 0) {
        continue;
      }
      std::wstring path(path_len + 1, L'\0');
      if (DragQueryFileW(drop, index, path.data(), static_cast<UINT>(path.size())) == 0) {
        continue;
      }
      path.resize(path_len);
      if (HasRegExtension(path)) {
        reg_paths.push_back(path);
      } else if (offline_candidate.empty()) {
        offline_candidate = path;
      }
    }
    DragFinish(drop);
    if (!reg_paths.empty()) {
      for (const auto& path : reg_paths) {
        OpenRegFileTab(path);
      }
    }
    if (!offline_candidate.empty()) {
      LoadOfflineRegistryFromPath(offline_candidate, true);
    }
    return 0;
  }
  case WM_LBUTTONDOWN: {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    if (show_tree_ && PtInRect(&splitter_rect_, pt)) {
      splitter_start_x_ = pt.x;
      splitter_start_width_ = tree_width_;
      BeginSplitterDrag();
      return 0;
    }
    if (show_history_ && PtInRect(&history_splitter_rect_, pt)) {
      history_splitter_start_y_ = pt.y;
      history_splitter_start_height_ = history_height_;
      BeginHistorySplitterDrag();
      return 0;
    }
    break;
  }
  case WM_LBUTTONUP:
    if (splitter_dragging_) {
      EndSplitterDrag(true);
      return 0;
    }
    if (history_splitter_dragging_) {
      EndHistorySplitterDrag(true);
      return 0;
    }
    break;
  case WM_MOUSEMOVE: {
    if (splitter_dragging_) {
      POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      UpdateSplitterTrack(pt.x);
      return 0;
    }
    if (history_splitter_dragging_) {
      POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      UpdateHistorySplitterTrack(pt.y);
      return 0;
    }
    break;
  }
  case WM_CAPTURECHANGED:
    if (splitter_dragging_) {
      EndSplitterDrag(false);
      return 0;
    }
    if (history_splitter_dragging_) {
      EndHistorySplitterDrag(false);
      return 0;
    }
    break;
  case WM_SETCURSOR: {
    if (splitter_dragging_) {
      SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
      return TRUE;
    }
    if (history_splitter_dragging_) {
      SetCursor(LoadCursorW(nullptr, IDC_SIZENS));
      return TRUE;
    }
    if (show_tree_) {
      POINT pt = {};
      GetCursorPos(&pt);
      ScreenToClient(hwnd_, &pt);
      if (PtInRect(&splitter_rect_, pt)) {
        SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
        return TRUE;
      }
    }
    if (show_history_) {
      POINT pt = {};
      GetCursorPos(&pt);
      ScreenToClient(hwnd_, &pt);
      if (PtInRect(&history_splitter_rect_, pt)) {
        SetCursor(LoadCursorW(nullptr, IDC_SIZENS));
        return TRUE;
      }
    }
    break;
  }
  case kSearchResultsMessage: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    std::vector<PendingSearchResult> pending;
    {
      std::lock_guard<std::mutex> lock(search_mutex_);
      if (search_pending_.empty()) {
        search_posted_.store(false);
        return 0;
      }
      pending.swap(search_pending_);
      search_posted_.store(false);
    }
    bool should_refresh = false;
    if (IsSearchTabIndex(active_search_tab_index_)) {
      int index = SearchIndexFromTab(active_search_tab_index_);
      if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size()) {
        uint64_t start_tick = GetTickCount64();
        size_t processed = 0;
        size_t stop_at = pending.size();
        for (size_t i = 0; i < pending.size(); ++i) {
          auto& item = pending[i];
          if (item.generation != generation) {
            continue;
          }
          search_tabs_[static_cast<size_t>(index)].results.push_back(std::move(item.result));
          ++processed;
          if (processed >= kSearchResultsBatch || (GetTickCount64() - start_tick) >= kSearchResultsMaxMs) {
            stop_at = i + 1;
            break;
          }
        }
        auto& tab = search_tabs_[static_cast<size_t>(index)];
        if (processed > 0 && tab.sort_column >= 0) {
          tab.sort_dirty = true;
        }
        if (stop_at < pending.size()) {
          std::vector<PendingSearchResult> remainder;
          remainder.reserve(pending.size() - stop_at);
          for (size_t i = stop_at; i < pending.size(); ++i) {
            if (pending[i].generation != generation) {
              continue;
            }
            remainder.push_back(std::move(pending[i]));
          }
          if (!remainder.empty()) {
            std::lock_guard<std::mutex> lock(search_mutex_);
            std::vector<PendingSearchResult> merged;
            merged.reserve(remainder.size() + search_pending_.size());
            for (auto& item : remainder) {
              merged.push_back(std::move(item));
            }
            for (auto& item : search_pending_) {
              merged.push_back(std::move(item));
            }
            search_pending_.swap(merged);
            if (!search_posted_.exchange(true)) {
              PostMessageW(hwnd_, kSearchResultsMessage, static_cast<WPARAM>(generation), 0);
            }
          }
        }
        if (TabCtrl_GetCurSel(tab_) == active_search_tab_index_) {
          should_refresh = true;
        }
      }
    }
    if (should_refresh) {
      uint64_t now = GetTickCount64();
      if (now - search_last_refresh_tick_ >= kSearchResultsRefreshMs) {
        search_last_refresh_tick_ = now;
        UpdateSearchResultsView();
        UpdateStatus();
      }
    }
    return 0;
  }
  case kSearchProgressMessage: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    search_progress_posted_.store(false);
    UpdateStatus();
    return 0;
  }
  case kSearchFinishedMessage: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    bool results_pending = search_posted_.load();
    {
      std::lock_guard<std::mutex> lock(search_mutex_);
      results_pending = results_pending || !search_pending_.empty();
    }
    if (results_pending) {
      PostMessageW(hwnd_, kSearchFinishedMessage, wparam, 0);
      return 0;
    }
    search_session_.Join();
    search_running_ = false;
    for (auto& search_tab : search_tabs_) {
      if (search_tab.generation == generation && search_tab.sort_dirty) {
        SortSearchTabResults(&search_tab);
        break;
      }
    }
    if (search_session_.IsCurrent(generation) &&
        search_start_tick_ != 0) {
      search_duration_ms_ = GetTickCount64() - search_start_tick_;
      search_duration_valid_ = true;
    } else {
      search_duration_ms_ = 0;
      search_duration_valid_ = false;
    }
    if (IsSearchTabIndex(TabCtrl_GetCurSel(tab_))) {
      search_last_refresh_tick_ = GetTickCount64();
      UpdateSearchResultsView();
    }
    ApplyViewVisibility();
    UpdateStatus();
    return 0;
  }
  case kSearchFailedMessage: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (!search_session_.IsCurrent(generation)) {
      return 0;
    }
    search_session_.Join();
    search_running_ = false;
    search_duration_ms_ = 0;
    search_duration_valid_ = false;
    ui::ShowError(hwnd_, L"Invalid regex.");
    ApplyViewVisibility();
    UpdateStatus();
    return 0;
  }
  case kReplaceReadyMessage:
    ApplyReplacePayload(reinterpret_cast<ReplacePayload*>(lparam));
    return 0;
  case kLoadTracesMessage:
    StartTraceLoadWorker();
    return 0;
  case kLoadDefaultsMessage:
    StartDefaultLoadWorker();
    return 0;
  case kTraceLoadReadyMessage: {
    auto* payload = reinterpret_cast<TraceLoadPayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<TraceLoadPayload> owned(payload);
    if (!trace_load_session_.IsCurrent(owned->generation)) {
      return 0;
    }
    trace_load_session_.Join();
    active_traces_ = std::move(owned->traces);
    trace_selection_cache_ = std::move(owned->selection_cache);
    BuildMenus();
    RefreshTreeSelection();
    UpdateValueListForNode(current_node_);
    return 0;
  }
  case kDefaultLoadReadyMessage: {
    auto* payload = reinterpret_cast<DefaultLoadPayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<DefaultLoadPayload> owned(payload);
    if (!default_load_session_.IsCurrent(owned->generation)) {
      return 0;
    }
    default_load_session_.Join();
    active_defaults_ = std::move(owned->defaults);
    BuildMenus();
    UpdateValueListForNode(current_node_);
    return 0;
  }
  case kDeferredStartupMessage:
    RunDeferredStartup();
    return 0;
  case kStartupCacheReadyMessage:
    ApplyStartupCachePayload(reinterpret_cast<StartupCachePayload*>(lparam));
    return 0;
  case kRegFileLoadReadyMessage: {
    auto* payload = reinterpret_cast<RegFileParsePayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<RegFileParsePayload> owned(payload);
    auto session_it = reg_file_parse_sessions_.find(owned->source_lower);
    if (session_it == reg_file_parse_sessions_.end()) {
      return 0;
    }
    if (!session_it->second ||
        !session_it->second->work.IsCurrent(owned->generation)) {
      return 0;
    }
    session_it->second->work.Join();
    reg_file_parse_sessions_.erase(session_it);
    if (owned->cancelled) {
      return 0;
    }

    int tab_index = -1;
    for (size_t i = 0; i < tabs_.size(); ++i) {
      const TabEntry& entry = tabs_[i];
      if (entry.kind != TabEntry::Kind::kRegFile) {
        continue;
      }
      if (EqualsInsensitive(entry.reg_file_path, owned->source_path)) {
        tab_index = static_cast<int>(i);
        break;
      }
    }
    if (tab_index < 0 || static_cast<size_t>(tab_index) >= tabs_.size()) {
      return 0;
    }
    TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
    entry.reg_file_loading = false;
    if (!owned->error.empty()) {
      ui::ShowError(hwnd_, owned->error.c_str());
      UpdateStatus();
      return 0;
    }
    if (entry.reg_file_dirty) {
      UpdateStatus();
      return 0;
    }

    ReleaseRegFileRoots(&entry);
    std::vector<TabEntry::RegFileRoot> roots;
    roots.reserve(owned->roots.size());
    for (auto& parsed : owned->roots) {
      if (!parsed.data) {
        continue;
      }
      TabEntry::RegFileRoot root;
      root.name = parsed.name;
      root.data = parsed.data;
      root.root = RegistryProvider::RegisterVirtualRoot(root.name, root.data);
      if (root.root) {
        roots.push_back(std::move(root));
      }
    }
    entry.reg_file_roots = std::move(roots);
    entry.reg_file_dirty = false;
    if (tab_ && TabCtrl_GetCurSel(tab_) == tab_index) {
      SyncRegFileTabSelection();
      ApplyViewVisibility();
      UpdateStatus();
    }
    return 0;
  }
  case kTraceParseBatchMessage: {
    auto* payload = reinterpret_cast<TraceParseBatch*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<TraceParseBatch> owned(payload);
    auto it = trace_parse_sessions_.find(owned->source_lower);
    if (it == trace_parse_sessions_.end()) {
      return 0;
    }
    TraceParseSession* session = it->second.get();
    if (!session || !session->work.IsCurrent(owned->generation)) {
      return 0;
    }
    const bool touches_current =
        current_node_ &&
        owned->affected_keys.find(TracePathLowerForNode(*current_node_)) !=
            owned->affected_keys.end();
    if (session->dialog && IsWindow(session->dialog) && !owned->entries.empty()) {
      auto dialog_entries = std::make_unique<std::vector<KeyValueDialogEntry>>(std::move(owned->entries));
      TraceDialogPostEntries(session->dialog, dialog_entries.release());
    }
    if (session->added_to_active && touches_current && current_node_) {
      uint64_t now = GetTickCount64();
      if (owned->done || (now - last_trace_refresh_tick_) >= 100) {
        last_trace_refresh_tick_ = now;
        UpdateValueListForNode(current_node_);
      }
    }
    if (owned->done) {
      session->parsing_done = true;
      if (session->data) {
        std::unique_lock<std::shared_mutex> data_lock(*session->data->mutex);
        trace::Selection normalized = session->selection;
        trace::NormalizeSelection(*session->data, &normalized);
        session->selection = normalized;
        trace_selection_cache_[session->source_lower] = normalized;
        if (session->added_to_active) {
          for (auto& trace : active_traces_) {
            if (EqualsInsensitive(trace.source_path, session->source_path)) {
              trace.selection =
                  std::make_shared<trace::Selection>(normalized);
              break;
            }
          }
        }
      }
      if (!owned->error.empty()) {
        HWND error_owner = session->dialog && IsWindow(session->dialog) ? session->dialog : hwnd_;
        ui::ShowError(error_owner, owned->error.c_str());
        if (session->dialog && IsWindow(session->dialog)) {
          PostMessageW(session->dialog, WM_CLOSE, 0, 0);
        }
        if (session->added_to_active) {
          active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(), [&](const ActiveTrace& trace) { return EqualsInsensitive(trace.source_path, session->source_path); }), active_traces_.end());
          trace_selection_cache_.erase(session->source_lower);
          SaveActiveTraces();
          SaveTraceSettings();
          BuildMenus();
          RefreshTreeSelection();
          UpdateValueListForNode(current_node_);
          SaveSettings();
          session->added_to_active = false;
        }
      } else if (session->dialog && IsWindow(session->dialog)) {
        TraceDialogPostDone(session->dialog, true);
      }
      if (session->added_to_active && current_node_) {
        UpdateValueListForNode(current_node_);
      }
      session->work.Join();
      if (!session->dialog || !IsWindow(session->dialog)) {
        trace_parse_sessions_.erase(it);
      }
    }
    return 0;
  }
  case kDefaultParseBatchMessage: {
    auto* payload = reinterpret_cast<DefaultParseBatch*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<DefaultParseBatch> owned(payload);
    auto it = default_parse_sessions_.find(owned->source_lower);
    if (it == default_parse_sessions_.end()) {
      return 0;
    }
    DefaultParseSession* session = it->second.get();
    if (!session || !session->work.IsCurrent(owned->generation)) {
      return 0;
    }
    bool touches_current = false;
    if (current_node_) {
      std::wstring path = registry_path::Build(*current_node_);
      std::wstring normalized = NormalizeTraceKeyPathBasic(path);
      if (normalized.empty()) {
        normalized = path;
      }
      touches_current =
          owned->affected_keys.find(ToLower(normalized)) !=
          owned->affected_keys.end();
    }
    if (session->dialog && IsWindow(session->dialog) && !owned->entries.empty()) {
      auto dialog_entries = std::make_unique<std::vector<KeyValueDialogEntry>>(std::move(owned->entries));
      TraceDialogPostEntries(session->dialog, dialog_entries.release());
    }
    if (session->added_to_active && touches_current && current_node_) {
      uint64_t now = GetTickCount64();
      if (owned->done || (now - last_default_refresh_tick_) >= 100) {
        last_default_refresh_tick_ = now;
        UpdateValueListForNode(current_node_);
      }
    }
    if (owned->done) {
      session->parsing_done = true;
      if (!owned->error.empty()) {
        if (session->show_errors) {
          HWND error_owner = session->dialog && IsWindow(session->dialog) ? session->dialog : hwnd_;
          ui::ShowError(error_owner, owned->error.c_str());
        }
        if (session->dialog && IsWindow(session->dialog)) {
          PostMessageW(session->dialog, WM_CLOSE, 0, 0);
        }
        if (session->added_to_active) {
          active_defaults_.erase(std::remove_if(active_defaults_.begin(), active_defaults_.end(), [&](const ActiveDefault& defaults) { return EqualsInsensitive(defaults.source_path, session->source_path); }), active_defaults_.end());
          SaveActiveDefaults();
          BuildMenus();
          UpdateValueListForNode(current_node_);
          SaveSettings();
          session->added_to_active = false;
        }
      } else if (session->dialog && IsWindow(session->dialog)) {
        TraceDialogPostDone(session->dialog, true);
      }
      if (session->added_to_active && current_node_) {
        UpdateValueListForNode(current_node_);
      }
      session->work.Join();
      if (!session->dialog || !IsWindow(session->dialog)) {
        default_parse_sessions_.erase(it);
      }
    }
    return 0;
  }
  case kValueListReadyMessage: {
    auto* payload = reinterpret_cast<ValueListPayload*>(lparam);
    if (!payload) {
      return 0;
    }
    std::unique_ptr<ValueListPayload> owned(payload);
    if (payload->generation != value_list_generation_.load()) {
      return 0;
    }
    HWND list_hwnd = value_list_.hwnd();
    if (list_hwnd) {
      SendMessageW(list_hwnd, WM_SETREDRAW, FALSE, 0);
    }
    value_list_.SetRows(std::move(payload->rows));
    RefreshValueListComments();
    current_key_count_ = payload->key_count;
    current_value_count_ = payload->value_count;
    if (list_hwnd && !jump_ui_batch_active_) {
      SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
      InvalidateRect(list_hwnd, nullptr, TRUE);
    }
    value_list_loading_ = false;
    UpdateStatus();
    StartPendingValueListRename();
    if (!pending_external_value_name_.empty() && current_node_) {
      std::wstring current_path = registry_path::Build(*current_node_);
      if (EqualsInsensitive(current_path, pending_external_value_key_path_)) {
        SelectValueByName(pending_external_value_name_);
        pending_external_value_key_path_.clear();
        pending_external_value_name_.clear();
      }
    }
    return 0;
  }
  case WM_TIMER:
    if (wparam == kRegeditCompatApplyTimerId) {
      KillTimer(hwnd_, kRegeditCompatApplyTimerId);
      ApplyPendingRegeditCompatNavigation();
      return 0;
    }
    break;
  case WM_COPYDATA: {
    auto* data = reinterpret_cast<const COPYDATASTRUCT*>(lparam);
    if (!data) {
      return 0;
    }
    if (data->dwData == kRegeditCompatActivateCopyDataId) {
      ActivateRegeditCompatibilityMode();
      ShowWindow(hwnd_, SW_RESTORE);
      SetForegroundWindow(hwnd_);
      FocusAddressBarForExternalJump(true);
      return TRUE;
    }
    if (data->dwData != kExternalJumpCopyDataId || !data->lpData || data->cbData < sizeof(wchar_t)) {
      return 0;
    }
    ActivateRegeditCompatibilityMode();
    size_t length = data->cbData / sizeof(wchar_t);
    const wchar_t* text = reinterpret_cast<const wchar_t*>(data->lpData);
    std::wstring target(text, text + length);
    while (!target.empty() && target.back() == L'\0') {
      target.pop_back();
    }
    if (target.empty()) {
      return 0;
    }
    // External jump requests are authoritative; clear any in-flight compat edit debounce.
    KillTimer(hwnd_, kRegeditCompatApplyTimerId);
    regedit_compat_pending_key_path_.clear();
    regedit_compat_pending_value_name_.clear();
    if (deferred_startup_complete_) {
      NavigateToExternalJump(target);
    } else {
      QueueExternalJump(target);
    }
    ShowWindow(hwnd_, SW_RESTORE);
    SetForegroundWindow(hwnd_);
    FocusAddressBarForExternalJump(true);
    return TRUE;
  }
  case WM_SETFOCUS:
    break;
  case WM_ERASEBKGND: {
    return 1;
  }
  case WM_PAINT:
    OnPaint();
    return 0;
  case WM_SETTINGCHANGE: {
    if (applying_theme_ || theme_mode_ != ThemeMode::kSystem) {
      return 0;
    }
    if (!Theme::UpdateFromSystem()) {
      return 0;
    }
    applying_theme_ = true;
    Theme::Current().ApplyToWindow(hwnd_);
    ApplyThemeToChildren();
    ReloadThemeIcons();
    if (hwnd_) {
      InvalidateRect(hwnd_, nullptr, TRUE);
    }
    applying_theme_ = false;
    return 0;
  }
  case WM_THEMECHANGED:
    if (applying_theme_ || theme_mode_ != ThemeMode::kSystem) {
      return 0;
    }
    if (!Theme::UpdateFromSystem()) {
      return 0;
    }
    applying_theme_ = true;
    Theme::Current().ApplyToWindow(hwnd_);
    ApplyThemeToChildren();
    ReloadThemeIcons();
    if (hwnd_) {
      InvalidateRect(hwnd_, nullptr, TRUE);
    }
    applying_theme_ = false;
    return 0;
  case WM_CTLCOLORSTATIC: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    const Theme& theme = Theme::Current();
    COLORREF color = theme.TextColor();
    COLORREF background = theme.PanelColor();
    HBRUSH brush = theme.PanelBrush();
    if (target == history_label_ || target == tree_header_) {
      background = theme.HeaderColor();
      brush = theme.HeaderBrush();
    }
    SetTextColor(hdc, color);
    SetBkColor(hdc, background);
    return reinterpret_cast<LRESULT>(brush);
  }
  case WM_INITMENUPOPUP: {
    HMENU menu = reinterpret_cast<HMENU>(wparam);
    UINT state = current_node_ ? MF_ENABLED : MF_GRAYED;
    EnableMenuItem(menu, cmd::kEditPermissions, MF_BYCOMMAND | state);
    const int selected_count = value_list_.hwnd() ? ListView_GetSelectedCount(value_list_.hwnd()) : 0;
    const int selected_index = selected_count == 1 ? ListView_GetNextItem(value_list_.hwnd(), -1, LVNI_SELECTED) : -1;
    const ListRow* selected_row = selected_index >= 0 ? value_list_.RowAt(selected_index) : nullptr;
    const bool can_modify_value = !read_only_ && selected_row &&
                                  selected_row->kind == rowkind::kValue &&
                                  !selected_row->simulated;
    const UINT modify_state = can_modify_value ? MF_ENABLED : MF_GRAYED;
    EnableMenuItem(menu, cmd::kEditModify, MF_BYCOMMAND | modify_state);
    EnableMenuItem(menu, cmd::kEditModifyBinary, MF_BYCOMMAND | modify_state);
    return 0;
  }
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    SetTextColor(hdc, Theme::Current().TextColor());
    SetBkColor(hdc, Theme::Current().SurfaceColor());
    return reinterpret_cast<LRESULT>(Theme::Current().SurfaceBrush());
  }
  case WM_CLOSE: {
    SaveSettings();
    DestroyWindow(hwnd_);
    return 0;
  }
  case WM_COMMAND: {
    if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kTreeHeaderCloseId) {
      show_tree_ = false;
      ApplyViewVisibility();
      BuildMenus();
      return 0;
    }
    if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kHistoryHeaderCloseId) {
      show_history_ = false;
      SaveSettings();
      ApplyViewVisibility();
      BuildMenus();
      return 0;
    }
    if (HIWORD(wparam) == BN_CLICKED && LOWORD(wparam) == kAddressGoId) {
      NavigateToAddress();
      return 0;
    }
    if (HIWORD(wparam) == EN_CHANGE && LOWORD(wparam) == kFilterEditId) {
      wchar_t buffer[256] = {};
      GetWindowTextW(filter_edit_, buffer, static_cast<int>(_countof(buffer)));
      bool needs_full_data = buffer[0] != L'\0';
      if (needs_full_data) {
        needs_full_data = std::any_of(value_list_.rows().begin(), value_list_.rows().end(), [](const ListRow& row) {
          return row.kind == rowkind::kValue && !row.data_ready;
        });
      }
      value_list_.SetFilter(buffer);
      if (needs_full_data && current_node_) {
        UpdateValueListForNode(current_node_);
      }
      UpdateStatus();
      return 0;
    }
    if (HIWORD(wparam) == 0 && HandleMenuCommand(LOWORD(wparam))) {
      return 0;
    }
    return 0;
  }
  case WM_CONTEXTMENU: {
    HWND source = reinterpret_cast<HWND>(wparam);
    HWND header_hwnd = ListView_GetHeader(value_list_.hwnd());
    if (source == header_hwnd) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(header_hwnd, &rect);
        screen_pt.x = rect.left + 12;
        screen_pt.y = rect.bottom - 4;
      }
      ShowValueHeaderMenu(screen_pt);
      return 0;
    }
    if (source == tree_.hwnd()) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(tree_.hwnd(), &rect);
        screen_pt.x = rect.left + 16;
        screen_pt.y = rect.top + 16;
      }
      ShowTreeContextMenu(screen_pt);
      return 0;
    }
    if (source == value_list_.hwnd()) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(value_list_.hwnd(), &rect);
        screen_pt.x = rect.left + 24;
        screen_pt.y = rect.top + 24;
      }
      ShowValueContextMenu(screen_pt);
      return 0;
    }
    if (source == history_list_) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(history_list_, &rect);
        screen_pt.x = rect.left + 24;
        screen_pt.y = rect.top + 24;
      }
      ShowHistoryContextMenu(screen_pt);
      return 0;
    }
    if (source == search_results_list_) {
      POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (screen_pt.x == -1 && screen_pt.y == -1) {
        RECT rect = {};
        GetWindowRect(search_results_list_, &rect);
        screen_pt.x = rect.left + 24;
        screen_pt.y = rect.top + 24;
      }
      ShowSearchResultContextMenu(screen_pt);
      return 0;
    }
    break;
  }
  case WM_DRAWITEM: {
    auto* draw = reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
    if (draw && draw->CtlType == ODT_MENU) {
      OnDrawMenuItem(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_BUTTON && draw->CtlID == kAddressGoId) {
      DrawAddressButton(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_BUTTON && draw->CtlID == kTreeHeaderCloseId) {
      DrawHeaderCloseButton(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_BUTTON && draw->CtlID == kHistoryHeaderCloseId) {
      DrawHeaderCloseButton(draw);
      return TRUE;
    }
    if (draw && draw->CtlType == ODT_STATIC && (draw->CtlID == kTreeHeaderId || draw->CtlID == kHistoryLabelId)) {
      const Theme& theme = Theme::Current();
      HDC hdc = draw->hDC;
      RECT rect = draw->rcItem;
      UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
      FillRect(hdc, &rect, theme.HeaderBrush());

      wchar_t text[128] = {};
      GetWindowTextW(draw->hwndItem, text, static_cast<int>(_countof(text)));
      HFONT old_font = nullptr;
      if (ui_font_) {
        old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
      }
      SetBkMode(hdc, TRANSPARENT);
      SetTextColor(hdc, theme.TextColor());
      RECT text_rect = rect;
      text_rect.left += kHeaderTextPadding;
      text_rect.right -= kHeaderTextPadding;
      DrawTextW(hdc, text, -1, &text_rect, format);
      if (old_font) {
        SelectObject(hdc, old_font);
      }
      return TRUE;
    }
    break;
  }
  case WM_MEASUREITEM: {
    auto* measure = reinterpret_cast<MEASUREITEMSTRUCT*>(lparam);
    if (measure && measure->CtlType == ODT_MENU) {
      OnMeasureMenuItem(measure);
      return TRUE;
    }
    break;
  }
  case WM_NOTIFY: {
    auto* header = reinterpret_cast<NMHDR*>(lparam);
    if (!header) {
      return 0;
    }
    if (header->code == TTN_GETDISPINFOW || header->code == TTN_NEEDTEXTW) {
      auto* info = reinterpret_cast<LPTOOLTIPTEXTW>(lparam);
      if (info) {
        int command_id = static_cast<int>(info->hdr.idFrom);
        std::wstring tip = CommandTooltipText(command_id);
        if (!tip.empty()) {
          std::wstring shortcut = CommandShortcutText(command_id);
          if (!shortcut.empty()) {
            tip.append(L" (");
            tip.append(shortcut);
            tip.append(L")");
          }
          static std::wstring tip_storage;
          tip_storage = tip;
          info->lpszText = const_cast<wchar_t*>(tip_storage.c_str());
          return 0;
        }
      }
    }
      if (header->hwndFrom == toolbar_.hwnd() && header->code == NM_CUSTOMDRAW) {
        auto* draw = reinterpret_cast<NMTBCUSTOMDRAW*>(lparam);
        if (!draw || !Theme::UseDarkMode()) {
          return CDRF_DODEFAULT;
        }
        const Theme& theme = Theme::Current();
        switch (draw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT:
          FillRect(draw->nmcd.hdc, &draw->nmcd.rc, theme.BackgroundBrush());
          return CDRF_NOTIFYITEMDRAW;
        case CDDS_ITEMPREPAINT: {
          bool is_separator = false;
          int command_id = static_cast<int>(draw->nmcd.dwItemSpec);
          int index = static_cast<int>(SendMessageW(toolbar_.hwnd(), TB_COMMANDTOINDEX, command_id, 0));
          if (index >= 0) {
            TBBUTTON button = {};
            if (SendMessageW(toolbar_.hwnd(), TB_GETBUTTON, index, reinterpret_cast<LPARAM>(&button))) {
              is_separator = (button.fsStyle & BTNS_SEP) != 0;
            }
          }
          if (is_separator) {
            return CDRF_DODEFAULT;
          }

          POINT cursor = {};
          GetCursorPos(&cursor);
          ScreenToClient(toolbar_.hwnd(), &cursor);
          bool is_hovered = ((draw->nmcd.uItemState & CDIS_HOT) == CDIS_HOT) || PtInRect(&draw->nmcd.rc, cursor);

          draw->hbrMonoDither = theme.BackgroundBrush();
          draw->hbrLines = theme.BackgroundBrush();
          draw->hpenLines = GetCachedPen(theme.BorderColor(), 1);
          draw->clrText = theme.TextColor();
          draw->clrTextHighlight = theme.TextColor();
          draw->clrBtnFace = theme.BackgroundColor();
          draw->clrBtnHighlight = theme.SurfaceColor();
          draw->clrHighlightHotTrack = theme.HoverColor();
          draw->nStringBkMode = TRANSPARENT;
          draw->nHLStringBkMode = TRANSPARENT;

          if (is_hovered) {
            DrawToolbarButtonBackground(draw->nmcd.hdc, draw->nmcd.rc, theme.HoverColor(), theme.BorderColor());
            draw->nmcd.uItemState &= ~(CDIS_HOT | CDIS_CHECKED);
          } else if ((draw->nmcd.uItemState & CDIS_CHECKED) == CDIS_CHECKED) {
            DrawToolbarButtonBackground(draw->nmcd.hdc, draw->nmcd.rc, theme.SurfaceColor(), theme.BorderColor());
            draw->nmcd.uItemState &= ~CDIS_CHECKED;
          }

          LRESULT lr = TBCDRF_USECDCOLORS;
          if ((draw->nmcd.uItemState & CDIS_SELECTED) == CDIS_SELECTED) {
            lr |= TBCDRF_NOBACKGROUND;
          }
          return lr;
        }
      default:
        break;
      }
      return CDRF_DODEFAULT;
    }
    if (header->hwndFrom == tab_ && header->code == TCN_SELCHANGING) {
      if (!suppress_tab_change_ && tab_) {
        int current = TabCtrl_GetCurSel(tab_);
        if (!IsSearchTabIndex(current) && !IsRegFileTabIndex(current)) {
          CaptureRegistryTabState(current);
        }
        last_tab_index_ = current;
      }
      return 0;
    }
    if (header->hwndFrom == tab_ && header->code == TCN_SELCHANGE) {
      if (suppress_tab_change_) {
        ApplyViewVisibility();
        UpdateSearchResultsView();
        UpdateStatus();
        return 0;
      }
      int sel = TabCtrl_GetCurSel(tab_);
      ApplyTabSelection(sel);
      ApplyViewVisibility();
      UpdateSearchResultsView();
      UpdateStatus();
      return 0;
    }
    if (header->hwndFrom == tree_.hwnd()) {
      if (header->code == TVN_ITEMEXPANDINGW) {
        tree_.OnItemExpanding(reinterpret_cast<NMTREEVIEWW*>(lparam));
        return 0;
      }
      if (header->code == TVN_ITEMEXPANDEDW) {
        if (!jump_ui_batch_active_) {
          MarkTreeStateDirty();
        }
        return 0;
      }
      if (header->code == TVN_BEGINLABELEDITW) {
        if (read_only_) {
          return TRUE;
        }
        auto* disp = reinterpret_cast<NMTVDISPINFOW*>(lparam);
        if (!disp) {
          return TRUE;
        }
        RegistryNode* node = tree_.NodeFromItem(disp->item.hItem);
        if (!node || node->subkey.empty()) {
          return TRUE;
        }
        HWND edit = TreeView_GetEditControl(tree_.hwnd());
        if (edit) {
          Theme::Current().ApplyToWindow(edit);
          Theme::Current().ApplyToChildren(edit);
          const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
          SetWindowTheme(edit, theme_name, nullptr);
        }
        return FALSE;
      }
      if (header->code == TVN_ENDLABELEDITW) {
        if (read_only_) {
          return FALSE;
        }
        auto* disp = reinterpret_cast<NMTVDISPINFOW*>(lparam);
        if (!disp || !disp->item.pszText) {
          return FALSE;
        }
        RegistryNode* node = tree_.NodeFromItem(disp->item.hItem);
        if (!node || node->subkey.empty()) {
          return FALSE;
        }
        std::wstring new_name = TrimWhitespace(disp->item.pszText);
        std::wstring old_name = LeafName(*node);
        if (new_name.empty() || _wcsicmp(new_name.c_str(), old_name.c_str()) == 0) {
          return FALSE;
        }
        if (!RegistryProvider::RenameKey(*node, new_name)) {
          ui::ShowError(hwnd_, L"Failed to rename key.");
          return FALSE;
        }
        UpdateLeafName(node, new_name);
        if (current_node_ && SameNode(*current_node_, *node)) {
          UpdateAddressBar(current_node_);
        }
        AppendHistoryEntry(L"Rename key", old_name, new_name);
        MarkOfflineDirty();
        RegistryNode parent = *node;
        if (!parent.subkey.empty()) {
          size_t pos = parent.subkey.rfind(L'\\');
          parent.subkey = (pos == std::wstring::npos) ? L"" : parent.subkey.substr(0, pos);
        }
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kRenameKey;
        op.node = parent;
        op.name = old_name;
        op.new_name = new_name;
        PushUndo(std::move(op));
        RefreshTreeSelection();
        UpdateValueListForNode(current_node_);
        return TRUE;
      }
      if (header->code == TVN_SELCHANGEDW) {
        auto* info = reinterpret_cast<NMTREEVIEWW*>(lparam);
        RegistryNode* previous_node = current_node_;
        RegistryNode* node = tree_.OnSelectionChanged(info);
        if (startup_tree_restore_pending_ && !applying_startup_tree_restore_) {
          startup_tree_restore_pending_ = false;
          tree_state_restored_ = true;
        }
        current_node_ = node;
        if (previous_node && node && SameNode(*previous_node, *node)) {
          return 0;
        }
        if (!jump_ui_batch_active_) {
          ApplyTreeSelectionEffects(node);
        }
        return 0;
      }
      if (header->code == NM_CUSTOMDRAW) {
        if (!Theme::UseDarkMode()) {
          return CDRF_DODEFAULT;
        }
        auto* draw = reinterpret_cast<NMTVCUSTOMDRAW*>(lparam);
        if (!draw) {
          return CDRF_DODEFAULT;
        }
        switch (draw->nmcd.dwDrawStage) {
        case CDDS_PREPAINT:
          return CDRF_NOTIFYITEMDRAW;
        case CDDS_ITEMPREPAINT: {
          if (draw->nmcd.uItemState & CDIS_SELECTED) {
            return CDRF_DODEFAULT;
          }
          const Theme& theme = Theme::Current();
          bool hot = (draw->nmcd.uItemState & CDIS_HOT) != 0;
          if (hot) {
            draw->clrText = theme.TextColor();
            draw->clrTextBk = theme.HoverColor();
          } else {
            draw->clrText = theme.TextColor();
            draw->clrTextBk = theme.PanelColor();
          }
          return CDRF_NEWFONT;
        }
        default:
          break;
        }
      }
    }
    if (header->hwndFrom == regedit_compat_tree_) {
      if (header->code == TVN_ITEMEXPANDINGW) {
        auto* info = reinterpret_cast<NMTREEVIEWW*>(lparam);
        if (info) {
          PopulateRegeditCompatChildren(info->itemNew.hItem);
        }
        return 0;
      }
      if (header->code == TVN_SELCHANGEDW) {
        HandleRegeditCompatTreeSelection(reinterpret_cast<NMTREEVIEWW*>(lparam));
        return 0;
      }
    }
    HWND value_header = ListView_GetHeader(value_list_.hwnd());
    HWND history_header = ListView_GetHeader(history_list_);
    HWND search_header = ListView_GetHeader(search_results_list_);
    if (header->hwndFrom == value_header && (header->code == HDN_ENDTRACKW || header->code == HDN_ENDTRACKA || header->code == HDN_ITEMCHANGEDW || header->code == HDN_ITEMCHANGEDA)) {
      auto* info = reinterpret_cast<NMHEADERW*>(lparam);
      if (info && info->iItem >= 0 && info->pitem && (info->pitem->mask & HDI_WIDTH)) {
        int subitem = GetListViewColumnSubItem(value_list_.hwnd(), info->iItem);
        if (subitem >= 0 && static_cast<size_t>(subitem) < value_column_widths_.size()) {
          value_column_widths_[static_cast<size_t>(subitem)] = info->pitem->cxy;
          SaveSettings();
        }
      }
    }
    if (header->hwndFrom == history_header && (header->code == HDN_ENDTRACKW || header->code == HDN_ENDTRACKA || header->code == HDN_ITEMCHANGEDW || header->code == HDN_ITEMCHANGEDA)) {
      auto* info = reinterpret_cast<NMHEADERW*>(lparam);
      if (info && info->iItem >= 0 && info->pitem && (info->pitem->mask & HDI_WIDTH)) {
        int subitem = GetListViewColumnSubItem(history_list_, info->iItem);
        if (subitem >= 0 && static_cast<size_t>(subitem) < history_column_widths_.size()) {
          history_column_widths_[static_cast<size_t>(subitem)] = info->pitem->cxy;
        }
      }
    }
    if (header->hwndFrom == search_header && (header->code == HDN_ENDTRACKW || header->code == HDN_ENDTRACKA || header->code == HDN_ITEMCHANGEDW || header->code == HDN_ITEMCHANGEDA)) {
      auto* info = reinterpret_cast<NMHEADERW*>(lparam);
      if (info && info->iItem >= 0 && info->pitem && (info->pitem->mask & HDI_WIDTH)) {
        int subitem = GetListViewColumnSubItem(search_results_list_, info->iItem);
        bool compare = IsCompareTabSelected();
        auto& widths = compare ? compare_column_widths_ : search_column_widths_;
        if (subitem >= 0 && static_cast<size_t>(subitem) < widths.size()) {
          widths[static_cast<size_t>(subitem)] = info->pitem->cxy;
        }
      }
    }
    if (header->hwndFrom == value_list_.hwnd() && header->code == LVN_GETDISPINFOW) {
      auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
      ListRow* mutable_row = value_list_.MutableRowAt(disp->item.iItem);
      const ListRow* row = mutable_row;
      if (!row) {
        if (disp->item.mask & LVIF_TEXT) {
          if (disp->item.pszText && disp->item.cchTextMax > 0) {
            disp->item.pszText[0] = L'\0';
          }
        }
        if (disp->item.mask & LVIF_IMAGE) {
          disp->item.iImage = 0;
        }
        return 0;
      }
      if (disp->item.mask & LVIF_TEXT) {
        if (disp->item.iSubItem == kValueColData && mutable_row) {
          EnsureValueRowData(mutable_row);
        }
        const wchar_t* text = L"";
        switch (disp->item.iSubItem) {
        case kValueColName:
          text = row->name.c_str();
          break;
        case kValueColType:
          text = row->type.c_str();
          break;
        case kValueColData:
          text = row->data.c_str();
          break;
        case kValueColDefault:
          text = row->default_data.c_str();
          break;
        case kValueColReadOnBoot:
          text = row->read_on_boot.c_str();
          break;
        case kValueColSize:
          text = row->size.c_str();
          break;
        case kValueColDate:
          text = row->date.c_str();
          break;
        case kValueColDetails:
          text = row->details.c_str();
          break;
        case kValueColComment:
          text = row->comment.c_str();
          break;
        default:
          text = row->extra.c_str();
          break;
        }
        if (disp->item.pszText && disp->item.cchTextMax > 0) {
          wcsncpy_s(disp->item.pszText, disp->item.cchTextMax, text, _TRUNCATE);
        }
      }
      if (disp->item.mask & LVIF_IMAGE) {
        disp->item.iImage = row->image_index;
      }
      return 0;
    }
    if (header->hwndFrom == value_list_.hwnd() && header->code == LVN_BEGINLABELEDITW) {
      if (read_only_) {
        return TRUE;
      }
      auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
      if (!disp) {
        return TRUE;
      }
      const ListRow* row = value_list_.RowAt(disp->item.iItem);
      if (!row || row->extra.empty() || (row->kind != rowkind::kValue && row->kind != rowkind::kKey)) {
        return TRUE;
      }
      HWND edit = ListView_GetEditControl(value_list_.hwnd());
      if (edit) {
        Theme::Current().ApplyToWindow(edit);
        Theme::Current().ApplyToChildren(edit);
        const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
        SetWindowTheme(edit, theme_name, nullptr);
      }
      return FALSE;
    }
    if (header->hwndFrom == value_list_.hwnd() && header->code == LVN_ENDLABELEDITW) {
      if (read_only_) {
        return FALSE;
      }
      auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
      if (!disp || !disp->item.pszText || !current_node_) {
        return FALSE;
      }
      const ListRow* row = value_list_.RowAt(disp->item.iItem);
      if (!row || row->extra.empty()) {
        return FALSE;
      }
      std::wstring new_name = TrimWhitespace(disp->item.pszText);
      std::wstring old_name = row->extra;
      if (new_name.empty() || _wcsicmp(new_name.c_str(), old_name.c_str()) == 0) {
        return FALSE;
      }
      if (row->kind == rowkind::kKey) {
        RegistryNode child = MakeChildNode(*current_node_, old_name);
        if (!RegistryProvider::RenameKey(child, new_name)) {
          ui::ShowError(hwnd_, L"Failed to rename key.");
          return FALSE;
        }
        AppendHistoryEntry(L"Rename key " + old_name, old_name, new_name);
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kRenameKey;
        op.node = *current_node_;
        op.name = old_name;
        op.new_name = new_name;
        PushUndo(std::move(op));
        RefreshTreeSelection();
        UpdateValueListForNode(current_node_);
        return TRUE;
      }
      if (!RegistryProvider::RenameValue(*current_node_, old_name, new_name)) {
        ui::ShowError(hwnd_, L"Failed to rename value.");
        return FALSE;
      }
      AppendHistoryEntry(L"Rename value " + old_name, old_name, new_name);
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kRenameValue;
      op.node = *current_node_;
      op.name = old_name;
      op.new_name = new_name;
      PushUndo(std::move(op));
      UpdateValueListForNode(current_node_);
      SelectValueByName(new_name);
      return TRUE;
    }
    if (header->hwndFrom == search_results_list_ && header->code == LVN_GETDISPINFOW) {
      auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
      search::Result* result = nullptr;
      int sel = TabCtrl_GetCurSel(tab_);
      int index = SearchIndexFromTab(sel);
      if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size()) {
        if (disp->item.iItem >= 0 && static_cast<size_t>(disp->item.iItem) < search_tabs_[static_cast<size_t>(index)].results.size()) {
          result = &search_tabs_[static_cast<size_t>(index)].results[static_cast<size_t>(disp->item.iItem)];
        }
      }
      if (!result) {
        if (disp->item.mask & LVIF_TEXT) {
          if (disp->item.pszText && disp->item.cchTextMax > 0) {
            disp->item.pszText[0] = L'\0';
          }
        }
        if (disp->item.mask & LVIF_IMAGE) {
          disp->item.iImage = 0;
        }
        return 0;
      }
      bool compare = false;
      if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size()) {
        compare = search_tabs_[static_cast<size_t>(index)].is_compare;
      }
      if (!compare && disp->item.iSubItem == 3) {
        EnsureSearchResultDataLoaded(result);
      }
      if (disp->item.mask & LVIF_TEXT) {
        const wchar_t* text = L"";
        if (compare) {
          switch (disp->item.iSubItem) {
          case 0:
            text = result->key_path.c_str();
            break;
          case 1:
            text = result->display_name.c_str();
            break;
          case 2:
            text = result->type_text.c_str();
            break;
          case 3:
            text = result->data.c_str();
            break;
          default:
            text = L"";
            break;
          }
        } else {
          switch (disp->item.iSubItem) {
          case 0:
            text = result->key_path.c_str();
            break;
          case 1:
            text = result->display_name.c_str();
            break;
          case 2:
            text = result->type_text.c_str();
            break;
          case 3:
            text = result->data.c_str();
            break;
          case 4:
            text = result->size_text.c_str();
            break;
          case 5:
            text = result->date_text.c_str();
            break;
          default:
            text = L"";
            break;
          }
        }
        if (disp->item.pszText && disp->item.cchTextMax > 0) {
          wcsncpy_s(disp->item.pszText, disp->item.cchTextMax, text, _TRUNCATE);
        }
      }
      if (disp->item.mask & LVIF_IMAGE) {
        if (result->is_key) {
          disp->item.iImage = kFolderIconIndex;
        } else if (UseBinaryValueIcon(result->type)) {
          disp->item.iImage = kBinaryIconIndex;
        } else {
          disp->item.iImage = kValueIconIndex;
        }
      }
      return 0;
    }
    if (header->hwndFrom == value_list_.hwnd() && header->code == LVN_ITEMCHANGED) {
      UpdateStatus();
      return 0;
    }
    if (header->hwndFrom == regedit_compat_list_ && header->code == LVN_ITEMCHANGED) {
      HandleRegeditCompatListSelection(reinterpret_cast<NMLISTVIEW*>(lparam));
      return 0;
    }
    if (header->hwndFrom == history_list_ && header->code == LVN_ITEMCHANGED) {
      ListView_SetItemState(history_list_, -1, 0, LVIS_FOCUSED);
      return 0;
    }
    if (header->hwndFrom == value_list_.hwnd() && header->code == LVN_COLUMNCLICK) {
      auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
      if (info) {
        SortValueList(info->iSubItem, true);
      }
      return 0;
    }
    if (header->hwndFrom == history_list_ && header->code == LVN_COLUMNCLICK) {
      auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
      if (info) {
        SortHistoryList(info->iSubItem, true);
      }
      return 0;
    }
    if (header->hwndFrom == search_results_list_ && header->code == LVN_COLUMNCLICK) {
      auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
      if (info) {
        SortSearchResults(info->iSubItem, true);
      }
      return 0;
    }
    if (header->hwndFrom == value_list_.hwnd() && (header->code == NM_DBLCLK || header->code == LVN_ITEMACTIVATE)) {
      auto* activate = reinterpret_cast<NMITEMACTIVATE*>(lparam);
      if (activate && activate->iItem >= 0 && current_node_) {
        const ListRow* row = value_list_.RowAt(activate->iItem);
        bool fast_activate = false;
        if (header->code == LVN_ITEMACTIVATE) {
          if (!value_activate_from_key_) {
            return 0;
          }
          value_activate_from_key_ = false;
          if (last_value_click_delta_valid_) {
            return 0;
          }
          fast_activate = true;
        }
        if (header->code == NM_DBLCLK) {
          fast_activate = true;
        }
        last_value_click_delta_valid_ = false;

        if (row && row->kind == rowkind::kKey) {
          if (fast_activate) {
            std::wstring path = registry_path::Build(*current_node_);
            if (!row->extra.empty()) {
              path.append(L"\\");
              path.append(row->extra);
            }
            SelectTreePath(path);
          }
          return 0;
        }
        if (row && row->kind == rowkind::kValue) {
          if (activate->iSubItem == kValueColComment) {
            HandleMenuCommand(cmd::kEditModifyComment);
          } else {
            HandleMenuCommand(cmd::kEditModify);
          }
          return 0;
        }
      }
      return 0;
    }
    if (header->hwndFrom == search_results_list_ && (header->code == NM_DBLCLK || header->code == LVN_ITEMACTIVATE)) {
      if (IsCompareTabSelected()) {
        return 0;
      }
      auto* activate = reinterpret_cast<NMITEMACTIVATE*>(lparam);
      if (activate && activate->iItem >= 0) {
        int sel = TabCtrl_GetCurSel(tab_);
        int index = SearchIndexFromTab(sel);
        if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size() && static_cast<size_t>(activate->iItem) < search_tabs_[static_cast<size_t>(index)].results.size()) {
          const auto& result = search_tabs_[static_cast<size_t>(index)].results[static_cast<size_t>(activate->iItem)];
          int registry_tab = FindFirstRegistryTabIndex();
          if (registry_tab >= 0) {
            TabCtrl_SetCurSel(tab_, registry_tab);
          }
          ApplyViewVisibility();
          UpdateStatus();
          SelectTreePath(result.key_path);
          if (!result.is_key) {
            SelectValueByName(result.value_name);
          }
        }
      }
      return 0;
    }
    if (header->code == NM_CUSTOMDRAW) {
      if (header->hwndFrom == search_results_list_) {
        auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);
        if (!draw) {
          return CDRF_DODEFAULT;
        }
        if (draw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) {
          int item_index = static_cast<int>(draw->nmcd.dwItemSpec);
          int sel_tab = TabCtrl_GetCurSel(tab_);
          int tab_index = SearchIndexFromTab(sel_tab);
          if (item_index >= 0 && tab_index >= 0 && static_cast<size_t>(tab_index) < search_tabs_.size() && static_cast<size_t>(item_index) < search_tabs_[static_cast<size_t>(tab_index)].results.size()) {
            const search::Result& result = search_tabs_[static_cast<size_t>(tab_index)].results[static_cast<size_t>(item_index)];
            bool selected = ListViewItemSelected(search_results_list_, item_index);
            if (!selected && DrawSearchMatchSubItem(result, draw->iSubItem, draw->nmcd.hdc, draw->nmcd.rc, ui_font_)) {
              return CDRF_SKIPDEFAULT;
            }
          }
        }
        return ui::HandleThemedListViewCustomDraw(search_results_list_, draw);
      }
      if (header->hwndFrom == history_list_) {
        return HandleHistoryListCustomDraw(history_list_, reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
      }
      if (header->hwndFrom == value_list_.hwnd()) {
        return ui::HandleThemedListViewCustomDraw(header->hwndFrom, reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
      }
      if (header->hwndFrom == value_header || header->hwndFrom == history_header || header->hwndFrom == search_header) {
        return CDRF_DODEFAULT;
      }
    }
    return 0;
  }
  case kAddressEnterMessage:
    NavigateToAddress();
    return 0;
  case kFocusAddressBarMessage:
    FocusAddressBarForExternalJump(false);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd_, message, wparam, lparam);
}

bool MainWindow::OnCreate() {
  ChangeWindowMessageFilterEx(hwnd_, WM_COPYDATA, MSGFLT_ALLOW, nullptr);
  ui_font_ = CreateUIFont();
  icon_font_ = CreateIconFont(10);
  custom_font_ = DefaultLogFont();
  LoadSettings();
  if (theme_mode_ == ThemeMode::kCustom) {
    LoadThemePresets();
  }
  ApplySavedWindowPlacement();
  if (theme_mode_ == ThemeMode::kCustom && ApplyThemePresetByName(active_theme_preset_, false)) {
    // Applied by preset.
  } else {
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
  }
  UpdateUIFont();
  BuildMenus();
  BuildAccelerators();

  toolbar_.Create(hwnd_, instance_, kToolbarId);

  std::vector<TBBUTTON> buttons;
  buttons.push_back({0, cmd::kRegistryLocal, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({1, cmd::kRegistryNetwork, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({2, cmd::kRegistryOffline, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({6, kToolbarSepGroup1, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0});
  buttons.push_back({3, cmd::kEditFind, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({4, cmd::kEditReplace, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({5, cmd::kFileExport, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({6, kToolbarSepGroup2, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0});
  buttons.push_back({6, cmd::kEditUndo, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({7, cmd::kEditRedo, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({8, cmd::kEditCopy, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({9, cmd::kEditPaste, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({10, cmd::kEditDelete, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({11, cmd::kViewRefresh, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({6, kToolbarSepGroup3, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0});
  buttons.push_back({12, cmd::kNavBack, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({13, cmd::kNavForward, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({14, cmd::kNavUp, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  toolbar_.AddButtons(buttons);

  address_edit_ = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_MULTILINE, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kAddressEditId)), instance_, nullptr);
  SetWindowSubclass(address_edit_, AddressEditProc, kAddressSubclassId, reinterpret_cast<DWORD_PTR>(this));
  SendMessageW(address_edit_, EM_SETCUEBANNER, TRUE, reinterpret_cast<LPARAM>(L"Registry path"));

  address_go_btn_ = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kAddressGoId)), instance_, nullptr);
  tab_ = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | TCS_TABS | TCS_FOCUSNEVER, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTabId)), instance_, nullptr);
  ApplyFont(tab_, ui_font_);
  TabCtrl_SetPadding(tab_, kTabTextPaddingX, kTabInsetY);
  SetWindowSubclass(tab_, TabProc, kTabSubclassId, reinterpret_cast<DWORD_PTR>(this));

  filter_edit_ = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_MULTILINE, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kFilterEditId)), instance_, nullptr);
  SetWindowSubclass(filter_edit_, FilterEditProc, kFilterSubclassId, 0);

  tree_header_ = CreateWindowExW(0, L"STATIC", L"Key Tree", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTreeHeaderId)), instance_, nullptr);
  tree_close_btn_ = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTreeHeaderCloseId)), instance_, nullptr);

  tree_.Create(hwnd_, instance_, kTreeId, false);
  SetWindowSubclass(tree_.hwnd(), TreeViewProc, kTreeViewSubclassId, reinterpret_cast<DWORD_PTR>(this));
  tree_.SetIconResolver([this](const RegistryNode& node) { return KeyIconIndex(node, nullptr, nullptr); });
  tree_.SetVirtualChildProvider([this](const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) { AppendTraceChildren(node, existing_lower, out); });
  value_list_.Create(hwnd_, instance_, kValueListId);
  search_results_list_ = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kSearchResultsListId)), instance_, nullptr);
  LoadTabs();

  history_label_ = CreateWindowExW(0, L"STATIC", L"History", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryLabelId)), instance_, nullptr);
  history_close_btn_ = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryHeaderCloseId)), instance_, nullptr);
  status_bar_ = CreateWindowExW(0, STATUSCLASSNAMEW, L"", WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kStatusBarId)), instance_, nullptr);
  if (status_bar_) {
    int parts[4] = {0, 0, 0, 0};
    SendMessageW(status_bar_, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));
  }
  search_progress_ = CreateWindowExW(0, PROGRESS_CLASSW, L"", WS_CHILD | PBS_MARQUEE, 0, 0, 0, 0, status_bar_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kSearchProgressId)), instance_, nullptr);
  if (search_progress_) {
    SendMessageW(search_progress_, PBM_SETMARQUEE, TRUE, 30);
    SendMessageW(search_progress_, PBM_SETRANGE32, 0, 1);
    ShowWindow(search_progress_, SW_HIDE);
  }
  history_list_ = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryListId)), instance_, nullptr);

  DWORD ex_mask = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_BORDERSELECT | LVS_EX_TRACKSELECT | LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE | LVS_EX_UNDERLINEHOT;
  DWORD ex_style = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER;
  ListView_SetExtendedListViewStyleEx(history_list_, ex_mask, ex_style);
  ListView_SetExtendedListViewStyleEx(search_results_list_, ex_mask, ex_style);
  SendMessageW(search_results_list_, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
  SendMessageW(history_list_, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
  SetWindowSubclass(value_list_.hwnd(), ListViewProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this));
  SetWindowSubclass(history_list_, ListViewProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this));
  SetWindowSubclass(search_results_list_, ListViewProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this));

  ApplyUIFontToControls();

  ApplyThemeToChildren();
  CreateValueColumns();
  CreateHistoryColumns();
  CreateSearchColumns();
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, 0);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, 0);
    if (read_only_) {
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditPaste, 0);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditDelete, 0);
    }
  }

  roots_ = RegistryProvider::DefaultRoots(show_extra_hives_);
  AppendRealRegistryRoot(&roots_);
  tree_.SetRegeditLayout(UseRegeditVisibleTreeLayout());
  tree_.SetRootLabel(TreeRootLabel());
  tree_.PopulateRoots(roots_);

  int initial_tab = tab_ ? TabCtrl_GetCurSel(tab_) : -1;
  if (initial_tab >= 0 && !IsSearchTabIndex(initial_tab) && !IsRegFileTabIndex(initial_tab)) {
    RestoreRegistryTabState(initial_tab);
  } else {
    SelectDefaultTreeItem();
  }
  StartValueListWorker();

  ApplyViewVisibility();
  ApplyAlwaysOnTop();
  UpdateStatus();
  return true;
}

void MainWindow::RunDeferredStartup() {
  if (deferred_startup_complete_) {
    return;
  }
  deferred_startup_complete_ = true;
  const bool has_external_jump = !queued_external_jump_target_.empty();
  bool use_global_tree_state = (!has_external_jump && save_tree_state_);
  if (use_global_tree_state && tab_) {
    int active_tab = TabCtrl_GetCurSel(tab_);
    if (active_tab >= 0 && static_cast<size_t>(active_tab) < tabs_.size()) {
      const TabEntry& entry = tabs_[static_cast<size_t>(active_tab)];
      if (entry.kind == TabEntry::Kind::kRegistry && (!entry.selected_path.empty() || !entry.expanded_paths.empty())) {
        use_global_tree_state = false;
      }
    }
  }
  startup_tree_restore_pending_ = use_global_tree_state;

  EnableAddressAutoComplete();
  ReloadThemeIcons();

  UpdateSearchResultsView();
  if (has_external_jump) {
    tree_state_restored_ = true;
  } else if (!use_global_tree_state) {
    tree_state_restored_ = true;
  }
  StartStartupCacheLoad(use_global_tree_state);
  ApplyQueuedExternalJump();
  StartTreeStateWorker();
  MarkTreeStateDirty();

  BuildMenus();
  UpdateStatus();
}

void MainWindow::StartStartupCacheLoad(bool include_tree_state) {
  StopStartupCacheLoad();
  const bool load_tree_state = include_tree_state && save_tree_state_;
  int history_max_rows = history_max_rows_;
  int history_sort_column = history_sort_column_;
  bool history_sort_ascending = history_sort_ascending_;
  const HWND hwnd = hwnd_;
  startup_cache_session_.Start(
      [this, load_tree_state, history_max_rows, history_sort_column,
       history_sort_ascending, hwnd](
          uint64_t generation, const std::atomic_bool& cancel) {
    auto payload = std::make_unique<StartupCachePayload>();
    payload->generation = generation;

    std::wstring comments_path = CommentsPath();
    std::wstring comments_content;
    if (!comments_path.empty() &&
        util::ReadTextFile(
            comments_path, &comments_content, nullptr,
            static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
      changes::CommentDocument comments =
          changes::ParseComments(comments_content);
      payload->value_comments = std::move(comments.value_entries);
      payload->name_comments = std::move(comments.name_entries);
      payload->comments_loaded = true;
    }
    if (cancel.load()) {
      return;
    }

    std::wstring history_path = HistoryCachePath();
    std::wstring history_content;
    if (!history_path.empty() &&
        util::ReadTextFile(
            history_path, &history_content, nullptr,
            static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
      changes::ChangeHistory history;
      history.Replace(
          std::move(changes::ParseHistory(history_content).entries),
          static_cast<size_t>(history_max_rows));
      history.Sort(history_sort_column, history_sort_ascending);
      payload->history_entries = std::move(history.entries());
      payload->history_loaded = true;
    } else {
      payload->history_loaded = true;
    }
    if (cancel.load()) {
      return;
    }

    if (load_tree_state) {
      std::wstring tree_path = TreeStatePath();
      std::wstring tree_content;
      if (!tree_path.empty() &&
          util::ReadTextFile(
              tree_path, &tree_content, nullptr,
              static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
        workspace::TreeState state =
            workspace::ParseTreeState(tree_content);
        payload->tree_selected_path = std::move(state.selected_path);
        payload->tree_expanded_paths = std::move(state.expanded_paths);
      }
      payload->tree_state_loaded = true;
    }

    if (cancel.load()) {
      return;
    }
    if (hwnd && IsWindow(hwnd) &&
        PostMessageW(hwnd, kStartupCacheReadyMessage,
                     static_cast<WPARAM>(generation),
                     reinterpret_cast<LPARAM>(payload.get()))) {
      ReleasePostedPayload(payload);
    }
  });
}

void MainWindow::StopStartupCacheLoad() {
  startup_cache_session_.CancelAndJoin();
}

void MainWindow::ApplyStartupCachePayload(StartupCachePayload* payload) {
  if (!payload) {
    return;
  }
  std::unique_ptr<StartupCachePayload> owned(payload);
  if (!startup_cache_session_.IsCurrent(owned->generation)) {
    return;
  }
  startup_cache_session_.Join();

  if (owned->comments_loaded) {
    std::unordered_map<std::wstring, changes::CommentEntry>
        merged_value_comments;
    std::unordered_map<std::wstring, changes::CommentEntry>
        merged_name_comments;
    for (auto& entry : owned->value_comments) {
      merged_value_comments[changes::ValueComments::ValueKey(entry.path, entry.name, entry.type)] = std::move(entry);
    }
    for (auto& entry : owned->name_comments) {
      merged_name_comments[changes::ValueComments::NameKey(entry.name, entry.type)] = std::move(entry);
    }
    for (auto& pair : value_comments_.value_entries()) {
      merged_value_comments[pair.first] = std::move(pair.second);
    }
    for (auto& pair : value_comments_.name_entries()) {
      merged_name_comments[pair.first] = std::move(pair.second);
    }
    value_comments_.value_entries() = std::move(merged_value_comments);
    value_comments_.name_entries() = std::move(merged_name_comments);
    RefreshValueListComments();
    UpdateSearchResultsView();
  }

  if (owned->history_loaded) {
    std::vector<HistoryEntry> pending_session_entries;
    if (!history_loaded_ && !change_history_.entries().empty()) {
      pending_session_entries = change_history_.entries();
    }
    if (!change_history_.entries().empty()) {
      owned->history_entries.insert(owned->history_entries.end(), change_history_.entries().begin(), change_history_.entries().end());
    }
    change_history_.Replace(std::move(owned->history_entries),
                            static_cast<size_t>(history_max_rows_));
    change_history_.Sort(history_sort_column_, history_sort_ascending_);
    history_loaded_ = true;
    for (const auto& entry : pending_session_entries) {
      AppendHistoryCache(entry);
    }
    RebuildHistoryList();
  }

  if (owned->tree_state_loaded) {
    saved_tree_state_.selected_path = std::move(owned->tree_selected_path);
    saved_tree_state_.expanded_paths =
        std::move(owned->tree_expanded_paths);
    if (startup_tree_restore_pending_ && !tree_state_restored_) {
      applying_startup_tree_restore_ = true;
      RestoreTreeState();
      applying_startup_tree_restore_ = false;
    } else if (!tree_state_restored_) {
      tree_state_restored_ = true;
    }
    startup_tree_restore_pending_ = false;
  }
}

void MainWindow::OnDestroy() {
  if (hwnd_) {
    RemovePropW(hwnd_, kRegKitWindowProperty);
    KillTimer(hwnd_, kRegeditCompatApplyTimerId);
  }
  EndJumpUiBatch();
  StopStartupCacheLoad();
  StopReplace();
  StopTraceParseSessions();
  StopDefaultParseSessions();
  StopRegFileParseSessions();
  StopTraceLoadWorker();
  StopDefaultLoadWorker();
  StopValueListWorker();
  StopTreeStateWorker();
  CancelSearch();
  DiscardWorkerMessages();
  for (auto& entry : tabs_) {
    if (entry.kind == TabEntry::Kind::kRegFile) {
      ReleaseRegFileRoots(&entry);
    }
  }
  if (clear_tabs_on_exit_) {
    ClearTabsCache();
  } else if (save_tabs_) {
    SaveTabs();
  }
  ClearHistoryItems(false);
  if (clear_history_on_exit_) {
    std::wstring history_path = HistoryCachePath();
    if (!history_path.empty()) {
      DeleteFileW(history_path.c_str());
    }
  }
  UnloadOfflineRegistry(nullptr);
  ReleaseRemoteRegistry();
  if (ui_font_ && ui_font_owned_) {
    DeleteObject(ui_font_);
  }
  ui_font_ = nullptr;
  ui_font_owned_ = false;
  if (icon_font_) {
    DeleteObject(icon_font_);
    icon_font_ = nullptr;
  }
  if (tree_images_) {
    ImageList_Destroy(tree_images_);
    tree_images_ = nullptr;
  }
  if (list_images_) {
    ImageList_Destroy(list_images_);
    list_images_ = nullptr;
  }
  if (address_go_icon_) {
    DestroyIcon(address_go_icon_);
    address_go_icon_ = nullptr;
  }
  if (address_autocomplete_) {
    address_autocomplete_->Release();
    address_autocomplete_ = nullptr;
  }
  if (address_autocomplete_source_) {
    address_autocomplete_source_->Release();
    address_autocomplete_source_ = nullptr;
  }
  if (accelerators_) {
    DestroyAcceleratorTable(accelerators_);
    accelerators_ = nullptr;
  }
  menu_items_.clear();
}

void MainWindow::DiscardWorkerMessages() {
  if (!hwnd_) {
    return;
  }
  MSG message = {};
  const UINT payload_messages[] = {
      kTraceLoadReadyMessage,    kDefaultLoadReadyMessage,
      kStartupCacheReadyMessage, kRegFileLoadReadyMessage,
      kTraceParseBatchMessage,   kDefaultParseBatchMessage,
      kValueListReadyMessage,    kReplaceReadyMessage};
  for (const UINT id : payload_messages) {
    while (PeekMessageW(&message, hwnd_, id, id, PM_REMOVE)) {
      switch (id) {
      case kTraceLoadReadyMessage:
        delete reinterpret_cast<TraceLoadPayload*>(message.lParam);
        break;
      case kDefaultLoadReadyMessage:
        delete reinterpret_cast<DefaultLoadPayload*>(message.lParam);
        break;
      case kStartupCacheReadyMessage:
        delete reinterpret_cast<StartupCachePayload*>(message.lParam);
        break;
      case kRegFileLoadReadyMessage:
        delete reinterpret_cast<RegFileParsePayload*>(message.lParam);
        break;
      case kTraceParseBatchMessage:
        delete reinterpret_cast<TraceParseBatch*>(message.lParam);
        break;
      case kDefaultParseBatchMessage:
        delete reinterpret_cast<DefaultParseBatch*>(message.lParam);
        break;
      case kValueListReadyMessage:
        delete reinterpret_cast<ValueListPayload*>(message.lParam);
        break;
      case kReplaceReadyMessage:
        delete reinterpret_cast<ReplacePayload*>(message.lParam);
        break;
      default:
        break;
      }
    }
  }
}

void MainWindow::OnSize(int width, int height) {
  LayoutControls(width, height);
}

void MainWindow::OnPaint() {
  PAINTSTRUCT ps = {};
  HDC hdc = BeginPaint(hwnd_, &ps);
  RECT client = {};
  GetClientRect(hwnd_, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  if (width <= 0 || height <= 0) {
    EndPaint(hwnd_, &ps);
    return;
  }

  const Theme& theme = Theme::Current();
  HDC mem_dc = CreateCompatibleDC(hdc);
  HBITMAP buffer = CreateCompatibleBitmap(hdc, width, height);
  HGDIOBJ old_bitmap = SelectObject(mem_dc, buffer);

  FillRect(mem_dc, &client, theme.BackgroundBrush());

  HPEN pen = GetCachedPen(theme.BorderColor(), 1);
  HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(mem_dc, pen));
  HBRUSH old_brush = reinterpret_cast<HBRUSH>(SelectObject(mem_dc, GetStockObject(NULL_BRUSH)));

  RECT rect = {};
  auto draw_border = [&](HWND child) {
    if (!GetChildRectInParent(hwnd_, child, &rect)) {
      return;
    }
    DrawOutlineRect(mem_dc, rect, kBorderInflate);
  };
  auto draw_panel = [&](HWND header, HWND body) {
    if (!header || !body) {
      return;
    }
    RECT header_rect = {};
    RECT body_rect = {};
    if (!GetChildRectInParent(hwnd_, header, &header_rect)) {
      return;
    }
    if (!GetChildRectInParent(hwnd_, body, &body_rect)) {
      return;
    }
    RECT combined = {};
    combined.left = std::min(header_rect.left, body_rect.left);
    combined.top = std::min(header_rect.top, body_rect.top);
    combined.right = std::max(header_rect.right, body_rect.right);
    combined.bottom = std::max(header_rect.bottom, body_rect.bottom);
    DrawOutlineRect(mem_dc, combined, kBorderInflate);
    MoveToEx(mem_dc, combined.left, header_rect.bottom, nullptr);
    LineTo(mem_dc, combined.right, header_rect.bottom);
  };

  bool show_search = IsSearchTabSelected();
  if (show_value_ && !show_search) {
    draw_border(value_list_.hwnd());
  }
  if (show_tree_ && !show_search) {
    draw_panel(tree_header_, tree_.hwnd());
  }
  if (show_history_ && !show_search) {
    draw_panel(history_label_, history_list_);
  }
  if (show_tree_ && show_value_ && splitter_rect_.right > splitter_rect_.left) {
    RECT split = splitter_rect_;
    FillRect(mem_dc, &split, theme.PanelBrush());
    int mid_x = (split.left + split.right) / 2;
    MoveToEx(mem_dc, mid_x, split.top + 4, nullptr);
    LineTo(mem_dc, mid_x, split.bottom - 4);
  }
  if (show_history_ && history_splitter_rect_.bottom > history_splitter_rect_.top) {
    RECT split = history_splitter_rect_;
    FillRect(mem_dc, &split, theme.PanelBrush());
    int mid_y = (split.top + split.bottom) / 2;
    MoveToEx(mem_dc, split.left + 4, mid_y, nullptr);
    LineTo(mem_dc, split.right - 4, mid_y);
  }

  if (address_edit_ && address_go_btn_) {
    RECT left = {};
    RECT right = {};
    if (GetChildRectInParent(hwnd_, address_edit_, &left) && GetChildRectInParent(hwnd_, address_go_btn_, &right)) {
      RECT combined = left;
      combined.right = right.right;
      DrawOutlineRect(mem_dc, combined, kBorderInflate);
    }
  }
  if (filter_edit_ && IsWindowVisible(filter_edit_)) {
    RECT filter_rect = {};
    if (GetChildRectInParent(hwnd_, filter_edit_, &filter_rect)) {
      DrawOutlineRect(mem_dc, filter_rect, kBorderInflate);
    }
  }

  HPEN top_pen = GetCachedPen(theme.BorderColor(), 1);
  HGDIOBJ old_top = SelectObject(mem_dc, top_pen);
  MoveToEx(mem_dc, 0, 0, nullptr);
  LineTo(mem_dc, client.right, 0);
  SelectObject(mem_dc, old_top);

  SelectObject(mem_dc, old_brush);
  SelectObject(mem_dc, old_pen);

  BitBlt(hdc, 0, 0, width, height, mem_dc, 0, 0, SRCCOPY);
  SelectObject(mem_dc, old_bitmap);
  DeleteObject(buffer);
  DeleteDC(mem_dc);

  EndPaint(hwnd_, &ps);
}

void MainWindow::PaintMenuBarSeparator() {
  if (!hwnd_ || !GetMenu(hwnd_)) {
    return;
  }

  MENUBARINFO menu_info = {};
  menu_info.cbSize = sizeof(menu_info);
  if (!GetMenuBarInfo(hwnd_, OBJID_MENU, 0, &menu_info)) {
    return;
  }

  RECT window_rect = {};
  if (!GetWindowRect(hwnd_, &window_rect)) {
    return;
  }

  RECT separator = menu_info.rcBar;
  OffsetRect(&separator, -window_rect.left, -window_rect.top);
  separator.top = separator.bottom - 1;
  separator.bottom += 1;
  if (separator.bottom <= separator.top) {
    return;
  }

  HDC hdc = GetWindowDC(hwnd_);
  if (!hdc) {
    return;
  }
  FillRect(hdc, &separator, Theme::Current().BackgroundBrush());
  ReleaseDC(hwnd_, hdc);
}

void MainWindow::ApplyThemeToChildren() {
  const Theme& theme = Theme::Current();

  theme.ApplyToToolbar(toolbar_.hwnd());
  theme.ApplyToTreeView(tree_.hwnd());
  theme.ApplyToTreeView(regedit_compat_tree_);
  theme.ApplyToListView(value_list_.hwnd());
  theme.ApplyToListView(regedit_compat_list_);
  theme.ApplyToListView(history_list_);
  theme.ApplyToListView(search_results_list_);
  theme.ApplyToTabControl(tab_);
  theme.ApplyToStatusBar(status_bar_);

  if (address_edit_) {
    SetWindowTheme(address_edit_, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(address_edit_, 6, 6);
    SetEditVerticalRect(address_edit_, ui_font_, 2, 6, 6);
  }
  if (regedit_compat_edit_) {
    SetWindowTheme(regedit_compat_edit_, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(regedit_compat_edit_, 6, 6);
    SetEditVerticalRect(regedit_compat_edit_, ui_font_, 2, 6, 6);
  }
  if (filter_edit_) {
    SetWindowTheme(filter_edit_, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(filter_edit_, 6, 6);
    SetEditVerticalRect(filter_edit_, ui_font_, 2, 6, 6);
  }
  if (tree_header_) {
    SetWindowTheme(tree_header_, L"", L"");
  }
  ApplyAutoCompleteTheme();
  DrawMenuBar(hwnd_);
}

void MainWindow::ApplySystemTheme() {
  if (applying_theme_) {
    return;
  }
  applying_theme_ = true;
  Theme::UpdateFromSystem();
  Theme::Current().ApplyToWindow(hwnd_);
  ApplyThemeToChildren();
  ReloadThemeIcons();
  if (hwnd_) {
    InvalidateRect(hwnd_, nullptr, TRUE);
  }
  applying_theme_ = false;
}

void MainWindow::LoadThemePresets() {
  std::vector<ThemePreset> presets;
  bool loaded = ThemePresetStore::Load(&presets);
  bool updated_builtins = false;
  if (!loaded || presets.empty()) {
    presets = ThemePresetStore::BuiltInPresets();
  } else {
    std::vector<ThemePreset> builtins = ThemePresetStore::BuiltInPresets();
    auto same_colors = [](const ThemeColors& left, const ThemeColors& right) { return left.background == right.background && left.panel == right.panel && left.surface == right.surface && left.header == right.header && left.border == right.border && left.text == right.text && left.muted_text == right.muted_text && left.accent == right.accent && left.selection == right.selection && left.selection_text == right.selection_text && left.hover == right.hover && left.focus == right.focus; };
    auto same_preset = [&](const ThemePreset& left, const ThemePreset& right) { return left.is_dark == right.is_dark && same_colors(left.colors, right.colors); };
    for (const auto& builtin : builtins) {
      auto it = std::find_if(presets.begin(), presets.end(), [&](const ThemePreset& existing) { return _wcsicmp(existing.name.c_str(), builtin.name.c_str()) == 0; });
      if (it == presets.end()) {
        presets.push_back(builtin);
        updated_builtins = true;
      } else if (!same_preset(*it, builtin)) {
        *it = builtin;
        updated_builtins = true;
      }
    }
  }
  theme_presets_ = std::move(presets);
  if (theme_presets_.empty()) {
    return;
  }
  if (active_theme_preset_.empty()) {
    active_theme_preset_ = theme_presets_.front().name;
  }
  auto it = std::find_if(theme_presets_.begin(), theme_presets_.end(), [&](const ThemePreset& preset) { return _wcsicmp(preset.name.c_str(), active_theme_preset_.c_str()) == 0; });
  if (it == theme_presets_.end()) {
    active_theme_preset_ = theme_presets_.front().name;
  }
  if (!loaded || updated_builtins) {
    SaveThemePresets();
  }
}

void MainWindow::SaveThemePresets() const {
  ThemePresetStore::Save(theme_presets_, nullptr);
}

bool MainWindow::ApplyThemePresetByName(const std::wstring& name, bool persist) {
  if (theme_presets_.empty()) {
    return false;
  }
  auto it = std::find_if(theme_presets_.begin(), theme_presets_.end(), [&](const ThemePreset& preset) { return _wcsicmp(preset.name.c_str(), name.c_str()) == 0; });
  if (it == theme_presets_.end()) {
    it = theme_presets_.begin();
  }
  Theme::SetCustomColors(it->colors, it->is_dark);
  theme_mode_ = ThemeMode::kCustom;
  active_theme_preset_ = it->name;
  Theme::SetMode(theme_mode_);
  ApplySystemTheme();
  if (persist) {
    SaveSettings();
    BuildMenus();
  }
  return true;
}

void MainWindow::UpdateThemePresets(const std::vector<ThemePreset>& presets, const std::wstring& active_name, bool apply_now) {
  theme_presets_ = presets;
  active_theme_preset_ = active_name;
  SaveThemePresets();
  if (apply_now) {
    ApplyThemePresetByName(active_theme_preset_, true);
  } else {
    SaveSettings();
    BuildMenus();
  }
}

void MainWindow::ApplyAlwaysOnTop() {
  if (!hwnd_) {
    return;
  }
  SetWindowPos(hwnd_, always_on_top_ ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
}

void MainWindow::UpdateUIFont() {
  HFONT next_font = nullptr;
  bool next_owned = false;
  if (use_custom_font_) {
    next_font = CreateFontIndirectW(&custom_font_);
    next_owned = next_font != nullptr;
  } else {
    LOGFONTW lf = DefaultLogFont();
    next_font = CreateFontIndirectW(&lf);
    next_owned = next_font != nullptr;
  }
  if (!next_font) {
    next_font = CreateUIFont();
    next_owned = false;
  }
  if (ui_font_ && ui_font_owned_) {
    DeleteObject(ui_font_);
  }
  ui_font_ = next_font;
  ui_font_owned_ = next_owned;
  ApplyUIFontToControls();
}

void MainWindow::ApplyUIFontToControls() {
  if (!ui_font_) {
    return;
  }
  ApplyFont(toolbar_.hwnd(), ui_font_);
  ApplyFont(address_edit_, ui_font_);
  ApplyFont(regedit_compat_edit_, ui_font_);
  ApplyFont(address_go_btn_, ui_font_);
  ApplyFont(filter_edit_, ui_font_);
  ApplyFont(tab_, ui_font_);
  ApplyFont(tree_header_, ui_font_);
  ApplyFont(tree_close_btn_, ui_font_);
  ApplyFont(tree_.hwnd(), ui_font_);
  ApplyFont(regedit_compat_tree_, ui_font_);
  ApplyFont(value_list_.hwnd(), ui_font_);
  ApplyFont(regedit_compat_list_, ui_font_);
  ApplyFont(history_close_btn_, ui_font_);
  ApplyFont(history_label_, ui_font_);
  ApplyFont(history_list_, ui_font_);
  ApplyFont(status_bar_, ui_font_);
  ApplyFont(search_results_list_, ui_font_);
  UpdateTabWidth();
  if (hwnd_) {
    DrawMenuBar(hwnd_);
  }
  InvalidateRect(hwnd_, nullptr, TRUE);
  if (hwnd_) {
    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    LayoutControls(rect.right, rect.bottom);
  }
}

void MainWindow::ComputeSplitterLimits(int* min_width, int* max_width) const {
  if (!min_width || !max_width || !hwnd_) {
    return;
  }
  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  int width = rect.right - rect.left;
  int available_width = std::max(0, width);
  int max_tree = std::max(kMinTreeWidth, available_width - kMinValueListWidth - kSplitterWidth);
  *min_width = kMinTreeWidth;
  *max_width = max_tree;
}

void MainWindow::ComputeHistorySplitterLimits(int* min_height, int* max_height) const {
  if (!min_height || !max_height || !hwnd_) {
    return;
  }
  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  int height = rect.bottom - rect.top;

  const int gap = 6;
  const int top_offset = 4;
  UINT dpi = GetWindowDpi(hwnd_);
  const int address_height = CalcEditHeight(address_edit_, ui_font_, util::ScaleForDpi(16, dpi));
  const int tabs_height = std::max(20, tab_height_);
  int status_height = 0;
  if (status_bar_ && show_status_bar_) {
    RECT sb_rect = {};
    GetWindowRect(status_bar_, &sb_rect);
    status_height = sb_rect.bottom - sb_rect.top;
    if (status_height <= 0) {
      status_height = 20;
    }
  }

  int y = top_offset;
  if (show_toolbar_) {
    SendMessageW(toolbar_.hwnd(), TB_AUTOSIZE, 0, 0);
    RECT tb_rect = {};
    GetWindowRect(toolbar_.hwnd(), &tb_rect);
    int toolbar_height = tb_rect.bottom - tb_rect.top;
    y += toolbar_height;
  }
  y += address_height + gap;
  y += 4;
  y += tabs_height + gap;

  int status_top = height - status_height;
  int content_total_height = std::max(0, status_top - y);
  int max_history = std::max(kMinHistoryHeight, content_total_height - kHistoryMaxPadding);
  *min_height = kMinHistoryHeight;
  *max_height = max_history;
}

void MainWindow::InitDragLayout() {
  if (!hwnd_) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd_, &client);
  drag_client_width_ = client.right - client.left;
  drag_client_height_ = client.bottom - client.top;
  drag_content_left_ = 0;
  drag_content_right_ = drag_client_width_;

  drag_content_top_ = splitter_rect_.top;
  if (drag_content_top_ <= 0) {
    RECT rect = {};
    HWND target = value_list_.hwnd() ? value_list_.hwnd() : search_results_list_;
    if (target && GetWindowRect(target, &rect)) {
      MapWindowPoints(nullptr, hwnd_, reinterpret_cast<POINT*>(&rect), 2);
      drag_content_top_ = rect.top;
    }
  }
  if (drag_content_top_ <= 0) {
    drag_content_top_ = 0;
  }

  drag_status_top_ = drag_client_height_;
  if (show_status_bar_ && status_bar_) {
    RECT rect = {};
    if (GetWindowRect(status_bar_, &rect)) {
      MapWindowPoints(nullptr, hwnd_, reinterpret_cast<POINT*>(&rect), 2);
      drag_status_top_ = rect.top;
    }
  }

  drag_tree_header_height_ = 20;
  if (tree_header_) {
    RECT rect = {};
    if (GetWindowRect(tree_header_, &rect)) {
      drag_tree_header_height_ = rect.bottom - rect.top;
    }
  }
  drag_history_label_height_ = 18;
  if (history_label_) {
    RECT rect = {};
    if (GetWindowRect(history_label_, &rect)) {
      drag_history_label_height_ = rect.bottom - rect.top;
    }
  }
  drag_layout_valid_ = true;
}

void MainWindow::ApplyDragLayout() {
  if (!hwnd_) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd_, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  if (!drag_layout_valid_ || width != drag_client_width_ || height != drag_client_height_) {
    InitDragLayout();
  }

  const int gap = 6;
  const bool show_search = IsSearchTabSelected();
  const bool show_tree = show_tree_ && !show_search;
  const bool show_history = show_history_ && !show_search;
  const bool show_value = show_value_ && !show_search;

  int content_left = drag_content_left_;
  int content_right = drag_content_right_;
  int y = drag_content_top_;
  int status_top = drag_status_top_;
  int content_total_height = std::max(0, status_top - y);
  int min_history = kMinHistoryHeight;
  int max_history = std::max(min_history, content_total_height - kHistoryMaxPadding);
  int history_height = show_history ? ClampValue(history_height_, min_history, max_history) : 0;
  if (show_history) {
    history_height_ = history_height;
  }
  int history_top = status_top - history_height;

  int history_splitter_height = show_history ? kHistorySplitterHeight : 0;
  int history_gap = show_history ? kHistoryGap : 0;
  int splitter_bottom = show_history ? (history_top - history_gap) : history_top;
  int splitter_top = show_history ? (splitter_bottom - history_splitter_height) : history_top;
  if (show_history) {
    history_splitter_rect_.left = content_left;
    history_splitter_rect_.right = content_right;
    history_splitter_rect_.top = splitter_top;
    history_splitter_rect_.bottom = splitter_bottom;
  } else {
    history_splitter_rect_ = {};
  }
  int content_bottom = show_history ? splitter_top : (status_top - gap);
  int content_height = std::max(0, content_bottom - y);

  int available_width = content_right - content_left;
  int min_tree = kMinTreeWidth;
  int min_list = kMinValueListWidth;
  int max_tree = std::max(min_tree, available_width - min_list - kSplitterWidth);
  int tree_width = show_tree ? ClampValue(tree_width_, min_tree, max_tree) : 0;
  if (show_tree) {
    tree_width_ = tree_width;
  }

  int tree_header_height = drag_tree_header_height_;
  int history_label_height = drag_history_label_height_;
  int list_x = show_tree ? (content_left + tree_width + kSplitterWidth) : content_left;
  int list_width = content_right - list_x;
  int tree_content_height = std::max(0, content_height - (show_tree ? tree_header_height : 0));

  auto get_panel_rect = [&](HWND header, HWND body, RECT* rect) {
    if (!rect) {
      return false;
    }
    RECT header_rect = {};
    RECT body_rect = {};
    bool has_header = GetChildRectInParent(hwnd_, header, &header_rect);
    bool has_body = GetChildRectInParent(hwnd_, body, &body_rect);
    if (!has_header && !has_body) {
      return false;
    }
    if (!has_body) {
      *rect = header_rect;
      return true;
    }
    if (!has_header) {
      *rect = body_rect;
      return true;
    }
    UnionRect(rect, &header_rect, &body_rect);
    return true;
  };

  RECT old_tree_panel_rect = {};
  RECT old_history_panel_rect = {};
  RECT old_value_rect = {};
  RECT old_splitter_rect = splitter_rect_;
  RECT old_history_splitter_rect = history_splitter_rect_;
  bool had_old_tree_panel = get_panel_rect(tree_header_, tree_.hwnd(), &old_tree_panel_rect);
  bool had_old_history_panel = get_panel_rect(history_label_, history_list_, &old_history_panel_rect);
  bool had_old_value = GetChildRectInParent(hwnd_, value_list_.hwnd(), &old_value_rect);

  int window_count = 0;
  if (show_tree) {
    window_count += 3;
  }
  if (show_history) {
    window_count += 2;
  }
  if (show_search || show_value) {
    window_count += 1;
  }
  HDWP hdwp = BeginDeferWindowPos(std::max(1, window_count));
  auto defer = [&](HWND target, int x, int y_pos, int w, int h) {
    if (!target) {
      return;
    }
    if (hdwp) {
      hdwp = DeferWindowPos(hdwp, target, nullptr, x, y_pos, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
    } else {
      SetWindowPos(target, nullptr, x, y_pos, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
    }
  };

  if (show_history) {
    int history_width = content_right - content_left;
    defer(history_label_, content_left, history_top, history_width, history_label_height);
    defer(history_close_btn_, content_left + history_width - 18, history_top + 1, 16, 16);
    defer(history_list_, content_left, history_top + history_label_height + 2, history_width, history_height - history_label_height - 2);
  }

  if (show_tree) {
    defer(tree_header_, content_left, y, tree_width, tree_header_height);
    defer(tree_close_btn_, content_left + tree_width - 18, y + 2, 16, 16);
    defer(tree_.hwnd(), content_left, y + tree_header_height, tree_width, tree_content_height);
    splitter_rect_.left = content_left + tree_width;
    splitter_rect_.right = splitter_rect_.left + kSplitterWidth;
    splitter_rect_.top = y;
    splitter_rect_.bottom = y + content_height;
  } else {
    splitter_rect_ = {};
  }

  if (show_search) {
    defer(search_results_list_, content_left, y, content_right - content_left, content_height);
  } else if (show_value) {
    defer(value_list_.hwnd(), list_x, y, list_width, content_height);
  }

  if (hdwp) {
    EndDeferWindowPos(hdwp);
  }

  RECT new_tree_panel_rect = {};
  RECT new_history_panel_rect = {};
  RECT new_value_rect = {};
  bool has_new_tree_panel = get_panel_rect(tree_header_, tree_.hwnd(), &new_tree_panel_rect);
  bool has_new_history_panel = get_panel_rect(history_label_, history_list_, &new_history_panel_rect);
  bool has_new_value = GetChildRectInParent(hwnd_, value_list_.hwnd(), &new_value_rect);

  RECT dirty_layout = {};
  bool has_dirty_layout = false;
  auto extend_dirty = [&](const RECT& rect, bool has_rect) {
    if (!has_rect) {
      return;
    }
    if (!has_dirty_layout) {
      dirty_layout = rect;
      has_dirty_layout = true;
      return;
    }
    RECT combined = {};
    UnionRect(&combined, &dirty_layout, &rect);
    dirty_layout = combined;
  };
  extend_dirty(old_tree_panel_rect, had_old_tree_panel);
  extend_dirty(new_tree_panel_rect, has_new_tree_panel);
  extend_dirty(old_history_panel_rect, had_old_history_panel);
  extend_dirty(new_history_panel_rect, has_new_history_panel);
  extend_dirty(old_value_rect, had_old_value);
  extend_dirty(new_value_rect, has_new_value);
  extend_dirty(old_splitter_rect, old_splitter_rect.right > old_splitter_rect.left);
  extend_dirty(splitter_rect_, splitter_rect_.right > splitter_rect_.left);
  extend_dirty(old_history_splitter_rect, old_history_splitter_rect.bottom > old_history_splitter_rect.top);
  extend_dirty(history_splitter_rect_, history_splitter_rect_.bottom > history_splitter_rect_.top);

  if (has_dirty_layout) {
    InvalidateRect(hwnd_, &dirty_layout, FALSE);
  }
  if (tree_header_) {
    RedrawWindow(tree_header_, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
  }
  if (tree_close_btn_) {
    RedrawWindow(tree_close_btn_, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
  }
  if (history_label_) {
    RedrawWindow(history_label_, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
  }
  if (history_close_btn_) {
    RedrawWindow(history_close_btn_, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
  }
}

void MainWindow::BeginSplitterDrag() {
  splitter_dragging_ = true;
  ComputeSplitterLimits(&splitter_min_width_, &splitter_max_width_);
  drag_layout_valid_ = false;
  SetCapture(hwnd_);
}

void MainWindow::BeginHistorySplitterDrag() {
  history_splitter_dragging_ = true;
  ComputeHistorySplitterLimits(&history_splitter_min_height_, &history_splitter_max_height_);
  drag_layout_valid_ = false;
  SetCapture(hwnd_);
}

void MainWindow::UpdateSplitterTrack(int client_x) {
  if (!splitter_dragging_) {
    return;
  }
  int desired = splitter_start_width_ + (client_x - splitter_start_x_);
  desired = ClampValue(desired, splitter_min_width_, splitter_max_width_);
  if (desired == tree_width_) {
    return;
  }
  tree_width_ = desired;
  ApplyDragLayout();
}

void MainWindow::UpdateHistorySplitterTrack(int client_y) {
  if (!history_splitter_dragging_) {
    return;
  }
  int desired = history_splitter_start_height_ - (client_y - history_splitter_start_y_);
  desired = ClampValue(desired, history_splitter_min_height_, history_splitter_max_height_);
  if (desired == history_height_) {
    return;
  }
  history_height_ = desired;
  ApplyDragLayout();
}

void MainWindow::EndSplitterDrag(bool apply) {
  if (!splitter_dragging_) {
    return;
  }
  splitter_dragging_ = false;
  if (GetCapture() == hwnd_) {
    ReleaseCapture();
  }
  if (apply) {
    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    LayoutControls(rect.right, rect.bottom);
  }
}

void MainWindow::EndHistorySplitterDrag(bool apply) {
  if (!history_splitter_dragging_) {
    return;
  }
  history_splitter_dragging_ = false;
  if (GetCapture() == hwnd_) {
    ReleaseCapture();
  }
  if (apply) {
    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    LayoutControls(rect.right, rect.bottom);
  }
}

void MainWindow::ApplyViewVisibility() {
  bool show_search = IsSearchTabSelected();
  bool show_tree = show_tree_ && !show_search;
  bool show_value = show_value_ && !show_search;
  bool show_history = show_history_ && !show_search;
  ShowWindow(toolbar_.hwnd(), show_toolbar_ ? SW_SHOW : SW_HIDE);
  ShowWindow(address_edit_, show_address_bar_ ? SW_SHOW : SW_HIDE);
  ShowWindow(address_go_btn_, show_address_bar_ ? SW_SHOW : SW_HIDE);
  ShowWindow(tab_, show_tab_control_ ? SW_SHOW : SW_HIDE);
  ShowWindow(filter_edit_, (show_value && show_filter_bar_) ? SW_SHOW : SW_HIDE);
  ShowWindow(tree_header_, show_tree ? SW_SHOW : SW_HIDE);
  ShowWindow(tree_close_btn_, show_tree ? SW_SHOW : SW_HIDE);
  ShowWindow(tree_.hwnd(), show_tree ? SW_SHOW : SW_HIDE);
  ShowWindow(value_list_.hwnd(), show_value ? SW_SHOW : SW_HIDE);
  ShowWindow(history_label_, show_history ? SW_SHOW : SW_HIDE);
  ShowWindow(history_close_btn_, show_history ? SW_SHOW : SW_HIDE);
  ShowWindow(history_list_, show_history ? SW_SHOW : SW_HIDE);
  ShowWindow(search_results_list_, show_search ? SW_SHOW : SW_HIDE);
  ShowWindow(regedit_compat_edit_, SW_HIDE);
  ShowWindow(regedit_compat_tree_, SW_HIDE);
  ShowWindow(regedit_compat_list_, SW_HIDE);
  if (show_search && search_results_list_) {
    LONG_PTR style = GetWindowLongPtrW(search_results_list_, GWL_STYLE);
    if (style & LVS_SINGLESEL) {
      SetWindowLongPtrW(search_results_list_, GWL_STYLE, style & ~LVS_SINGLESEL);
      SetWindowPos(search_results_list_, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    }
  }
  ShowWindow(status_bar_, show_status_bar_ ? SW_SHOW : SW_HIDE);
  if (search_progress_) {
    bool show_progress = show_status_bar_ && show_search && search_running_ && !IsCompareTabSelected();
    ShowWindow(search_progress_, show_progress ? SW_SHOW : SW_HIDE);
  }

  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  LayoutControls(rect.right, rect.bottom);
}

void MainWindow::ApplyTabSelection(int index) {
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  const TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind == TabEntry::Kind::kRegistry) {
    switch (entry.registry_mode) {
    case RegistryMode::kLocal: {
      SwitchToLocalRegistry();
      break;
    }
    case RegistryMode::kOffline:
      if (!entry.offline_path.empty()) {
        LoadOfflineRegistryFromPath(entry.offline_path, false);
      }
      break;
    case RegistryMode::kRemote:
      if (!entry.remote_machine.empty()) {
        remote_machine_ = entry.remote_machine;
      }
      if (registry_mode_ != RegistryMode::kRemote) {
        SwitchToRemoteRegistry();
      }
      break;
    }
    RestoreRegistryTabState(index);
  } else if (entry.kind == TabEntry::Kind::kSearch) {
    EnsureSearchTabResultsLoaded(entry.search_index);
  } else if (entry.kind == TabEntry::Kind::kRegFile) {
    SyncRegFileTabSelection();
  }
}

void MainWindow::ResetHiveListCache() {
  hive_list_loaded_ = false;
  hive_list_.clear();
  hive_roots_.reset();
}

void MainWindow::EnsureHiveListLoaded() {
  if (hive_list_loaded_) {
    return;
  }
  hive_list_loaded_ = true;
  hive_list_.clear();
  auto hive_roots = std::make_shared<std::unordered_set<std::wstring>>();
  hive_roots_ = hive_roots;

  HKEY hklm = nullptr;
  for (const auto& root : roots_) {
    if (EqualsInsensitive(root.display_name, L"HKEY_LOCAL_MACHINE")) {
      hklm = root.root;
      break;
    }
  }
  if (!hklm) {
    return;
  }
  util::UniqueHKey hive_key;
  if (RegOpenKeyExW(hklm, L"SYSTEM\\CurrentControlSet\\Control\\hivelist", 0, KEY_READ, hive_key.put()) != ERROR_SUCCESS) {
    return;
  }

  DWORD value_count = 0;
  DWORD max_name_len = 0;
  DWORD max_data_len = 0;
  if (RegQueryInfoKeyW(hive_key.get(), nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, &value_count, &max_name_len, &max_data_len, nullptr, nullptr) != ERROR_SUCCESS) {
    return;
  }

  std::vector<wchar_t> name_buffer(max_name_len + 1, L'\0');
  std::vector<BYTE> data_buffer(max_data_len > 0 ? max_data_len : 1);

  for (DWORD i = 0; i < value_count; ++i) {
    DWORD name_len = static_cast<DWORD>(name_buffer.size());
    DWORD data_len = static_cast<DWORD>(data_buffer.size());
    DWORD type = 0;
    LONG result = RegEnumValueW(hive_key.get(), i, name_buffer.data(), &name_len, nullptr, &type, data_buffer.data(), &data_len);
    if (result != ERROR_SUCCESS || name_len == 0 || data_len == 0) {
      continue;
    }
    if (type != REG_SZ && type != REG_EXPAND_SZ) {
      continue;
    }
    std::wstring name(name_buffer.data(), name_len);
    std::wstring data(reinterpret_cast<wchar_t*>(data_buffer.data()), data_len / sizeof(wchar_t));
    while (!data.empty() && data.back() == L'\0') {
      data.pop_back();
    }
    if (data.empty()) {
      continue;
    }
    data = NormalizeHiveFilePath(data);
    if (data.empty()) {
      continue;
    }
    std::wstring name_lower = ToLower(name);
    hive_list_.emplace(name_lower, std::move(data));
    hive_roots->insert(std::move(name_lower));
  }
}

std::wstring MainWindow::LookupHivePath(const RegistryNode& node, bool* is_root) {
  if (is_root) {
    *is_root = false;
  }
  EnsureHiveListLoaded();
  if (hive_list_.empty()) {
    return L"";
  }
  std::wstring nt_path = registry_path::BuildNative(node);
  if (nt_path.empty() && !node.root_name.empty()) {
    auto equals_root = [&](const wchar_t* name) -> bool { return EqualsInsensitive(node.root_name, name); };
    if (equals_root(L"REGISTRY")) {
      nt_path = L"\\REGISTRY";
    } else if (equals_root(L"HKLM") || equals_root(L"HKEY_LOCAL_MACHINE")) {
      nt_path = L"\\REGISTRY\\MACHINE";
    } else if (equals_root(L"HKU") || equals_root(L"HKEY_USERS")) {
      nt_path = L"\\REGISTRY\\USER";
    } else if (equals_root(L"HKCU") || equals_root(L"HKEY_CURRENT_USER")) {
      std::wstring sid = util::GetCurrentUserSidString();
      if (!sid.empty()) {
        nt_path = L"\\REGISTRY\\USER\\" + sid;
      }
    } else if (equals_root(L"HKCC") || equals_root(L"HKEY_CURRENT_CONFIG")) {
      nt_path = L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Hardware\\Profiles\\Current";
    } else if (equals_root(L"HKCR") || equals_root(L"HKEY_CLASSES_ROOT")) {
      nt_path = L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes";
    }
    if (!nt_path.empty() && !node.subkey.empty()) {
      nt_path += L"\\" + node.subkey;
    }
  }
  if (nt_path.empty()) {
    return L"";
  }
  std::wstring nt_lower = ToLower(nt_path);
  size_t best_len = 0;
  std::wstring best_path;
  for (const auto& entry : hive_list_) {
    const std::wstring& hive_key = entry.first;
    if (nt_lower.size() < hive_key.size()) {
      continue;
    }
    if (nt_lower.compare(0, hive_key.size(), hive_key) != 0) {
      continue;
    }
    if (nt_lower.size() > hive_key.size() && nt_lower[hive_key.size()] != L'\\') {
      continue;
    }
    if (hive_key.size() > best_len) {
      best_len = hive_key.size();
      best_path = entry.second;
    }
  }
  if (best_len > 0 && is_root) {
    *is_root = nt_lower.size() == best_len;
  }
  return best_path;
}

int MainWindow::KeyIconIndex(const RegistryNode& node, bool* is_link, bool* is_hive_root) {
  if (is_link) {
    *is_link = false;
  }
  if (is_hive_root) {
    *is_hive_root = false;
  }
  if (node.simulated) {
    return kFolderSimIconIndex;
  }
  std::wstring link_target;
  if (RegistryProvider::QuerySymbolicLinkTarget(node, &link_target)) {
    if (is_link) {
      *is_link = true;
    }
    return kSymlinkIconIndex;
  }
  bool hive_root = false;
  std::wstring hive_path = LookupHivePath(node, &hive_root);
  if (!hive_path.empty() && hive_root && node.subkey.empty()) {
    if (node.root == HKEY_CURRENT_USER || EqualsInsensitive(node.root_name, L"HKEY_CURRENT_USER")) {
      hive_root = false;
    }
  }
  if (!hive_path.empty() && hive_root) {
    if (is_hive_root) {
      *is_hive_root = true;
    }
    return kDatabaseIconIndex;
  }
  return kFolderIconIndex;
}

std::wstring MainWindow::ResolveIconPath(const wchar_t* filename, bool use_light) const {
  if (!filename || !*filename) {
    return L"";
  }
  if (icon_set_.empty() || IsIconSetName(icon_set_, kIconSetDefault)) {
    return L"";
  }
  if (IsIconSetName(icon_set_, kIconSetCustom)) {
    std::wstring root = util::JoinPath(util::GetAppDataFolder(), L"icons");
    if (root.empty()) {
      return L"";
    }
    std::wstring dark_dir = util::JoinPath(root, L"dark");
    std::wstring light_dir = util::JoinPath(root, L"light");
    std::wstring dir;
    if (IsDirectoryPath(dark_dir) && IsDirectoryPath(light_dir)) {
      dir = use_light ? light_dir : dark_dir;
    } else if (IsDirectoryPath(root)) {
      dir = root;
    } else {
      return L"";
    }
    return util::JoinPath(dir, filename);
  }
  if (!IsKnownIconSetName(icon_set_)) {
    return L"";
  }
  std::wstring base = AssetsIconsRoot();
  if (base.empty()) {
    return L"";
  }
  base = util::JoinPath(base, icon_set_);
  std::wstring dir = util::JoinPath(base, use_light ? L"light" : L"dark");
  if (!IsDirectoryPath(dir)) {
    return L"";
  }
  return util::JoinPath(dir, filename);
}

bool MainWindow::ShouldUseLightIcons() const {
  switch (theme_mode_) {
  case ThemeMode::kDark:
    return true;
  case ThemeMode::kLight:
    return false;
  case ThemeMode::kSystem:
    return Theme::IsSystemDarkMode();
  case ThemeMode::kCustom:
  default:
    return Theme::UseDarkMode();
  }
}

HICON MainWindow::LoadThemeIcon(const wchar_t* filename, int light_id, int dark_id, int size, UINT dpi) const {
  bool use_light = ShouldUseLightIcons();
  std::wstring path = ResolveIconPath(filename, use_light);
  HICON icon = nullptr;
  if (!path.empty()) {
    icon = util::LoadIconFromFile(path, size, dpi);
  }
  if (!icon) {
    int resource_id = use_light ? light_id : dark_id;
    icon = util::LoadIconResource(resource_id, size, dpi);
  }
  return icon;
}

ToolbarIcon MainWindow::MakeToolbarIcon(const wchar_t* filename, int light_id, int dark_id, bool use_light) const {
  ToolbarIcon icon;
  icon.resource_id = use_light ? light_id : dark_id;
  std::wstring path = ResolveIconPath(filename, use_light);
  if (!path.empty()) {
    icon.path = path;
  }
  return icon;
}

void MainWindow::ReloadThemeIcons() {
  UINT dpi = GetWindowDpi(hwnd_);
  bool use_light = ShouldUseLightIcons();
  auto set_redraw = [](HWND hwnd, bool enable) {
    if (!hwnd) {
      return;
    }
    SendMessageW(hwnd, WM_SETREDRAW, enable ? TRUE : FALSE, 0);
  };
  set_redraw(toolbar_.hwnd(), false);
  set_redraw(tree_.hwnd(), false);
  set_redraw(value_list_.hwnd(), false);
  set_redraw(search_results_list_, false);
  set_redraw(address_go_btn_, false);

  toolbar_.LoadIcons(
      {
          MakeToolbarIcon(L"local-registry.ico", IDI_ICON_LIGHT_LOCAL_REGISTRY, IDI_ICON_DARK_LOCAL_REGISTRY, use_light),
          MakeToolbarIcon(L"remote-registry.ico", IDI_ICON_LIGHT_REMOTE_REGISTRY, IDI_ICON_DARK_REMOTE_REGISTRY, use_light),
          MakeToolbarIcon(L"offline-registry.ico", IDI_ICON_LIGHT_OFFLINE_REGISTRY, IDI_ICON_DARK_OFFLINE_REGISTRY, use_light),
          MakeToolbarIcon(L"search.ico", IDI_ICON_LIGHT_SEARCH, IDI_ICON_DARK_SEARCH, use_light),
          MakeToolbarIcon(L"replace.ico", IDI_ICON_LIGHT_REPLACE, IDI_ICON_DARK_REPLACE, use_light),
          MakeToolbarIcon(L"export.ico", IDI_ICON_LIGHT_EXPORT, IDI_ICON_DARK_EXPORT, use_light),
          MakeToolbarIcon(L"undo.ico", IDI_ICON_LIGHT_UNDO, IDI_ICON_DARK_UNDO, use_light),
          MakeToolbarIcon(L"redo.ico", IDI_ICON_LIGHT_REDO, IDI_ICON_DARK_REDO, use_light),
          MakeToolbarIcon(L"copy.ico", IDI_ICON_LIGHT_COPY, IDI_ICON_DARK_COPY, use_light),
          MakeToolbarIcon(L"paste.ico", IDI_ICON_LIGHT_PASTE, IDI_ICON_DARK_PASTE, use_light),
          MakeToolbarIcon(L"delete.ico", IDI_ICON_LIGHT_DELETE, IDI_ICON_DARK_DELETE, use_light),
          MakeToolbarIcon(L"refresh.ico", IDI_ICON_LIGHT_REFRESH, IDI_ICON_DARK_REFRESH, use_light),
          MakeToolbarIcon(L"back.ico", IDI_ICON_LIGHT_BACK, IDI_ICON_DARK_BACK, use_light),
          MakeToolbarIcon(L"forward.ico", IDI_ICON_LIGHT_FORWARD, IDI_ICON_DARK_FORWARD, use_light),
          MakeToolbarIcon(L"up.ico", IDI_ICON_LIGHT_UP, IDI_ICON_DARK_UP, use_light),
      },
      kToolbarIconSize, kToolbarGlyphSize);

  BuildImageLists();
  if (tree_.hwnd()) {
    tree_.SetImageList(tree_images_);
  }
  if (value_list_.hwnd()) {
    value_list_.SetImageList(list_images_);
  }
  if (search_results_list_) {
    ListView_SetImageList(search_results_list_, list_images_, LVSIL_SMALL);
  }

  if (address_go_icon_) {
    DestroyIcon(address_go_icon_);
    address_go_icon_ = nullptr;
  }
  address_go_icon_ = LoadThemeIcon(L"forward.ico", IDI_ICON_LIGHT_FORWARD, IDI_ICON_DARK_FORWARD, kToolbarGlyphSize, dpi);

  set_redraw(toolbar_.hwnd(), true);
  set_redraw(tree_.hwnd(), true);
  set_redraw(value_list_.hwnd(), true);
  set_redraw(search_results_list_, true);
  set_redraw(address_go_btn_, true);
  if (toolbar_.hwnd()) {
    InvalidateRect(toolbar_.hwnd(), nullptr, TRUE);
  }
  if (tree_.hwnd()) {
    InvalidateRect(tree_.hwnd(), nullptr, TRUE);
  }
  if (value_list_.hwnd()) {
    InvalidateRect(value_list_.hwnd(), nullptr, TRUE);
  }
  if (search_results_list_) {
    InvalidateRect(search_results_list_, nullptr, TRUE);
  }
  if (address_go_btn_) {
    InvalidateRect(address_go_btn_, nullptr, TRUE);
  }
}

void MainWindow::LayoutControls(int width, int height) {
  if (width <= 0 || height <= 0) {
    return;
  }

  const int padding = 8;
  const int gap = 6;
  const int splitter_width = kSplitterWidth;
  const int top_offset = 4;
  UINT dpi = GetWindowDpi(hwnd_);
  const int address_height = CalcEditHeight(address_edit_, ui_font_, util::ScaleForDpi(18, dpi));
  const int address_btn_width = std::max(util::ScaleForDpi(18, dpi), address_height);
  const int tabs_height = std::max(20, tab_height_);
  const int filter_height = address_height;
  const int filter_min_width = 160;
  const int filter_max_width = 260;
  const int filter_gap = 6;
  const int tree_header_height = 20;
  const int history_label_height = 18;
  int status_height = 0;
  if (status_bar_ && show_status_bar_) {
    RECT sb_rect = {};
    GetWindowRect(status_bar_, &sb_rect);
    status_height = sb_rect.bottom - sb_rect.top;
    if (status_height <= 0) {
      status_height = 20;
    }
  }
  const bool show_search = IsSearchTabSelected();
  const bool show_tree = show_tree_ && !show_search;
  const bool show_history = show_history_ && !show_search;
  const bool show_value = show_value_ && !show_search;

  int y = top_offset;

  const bool dragging_splitter = splitter_dragging_ || history_splitter_dragging_;
  auto place = [&](HWND hwnd, int x, int y_pos, int w, int h) {
    if (!hwnd) {
      return;
    }
    UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;
    if (!dragging_splitter) {
      flags |= SWP_NOREDRAW;
    }
    SetWindowPos(hwnd, nullptr, x, y_pos, w, h, flags);
  };

  if (show_toolbar_) {
    SendMessageW(toolbar_.hwnd(), TB_AUTOSIZE, 0, 0);
    RECT tb_rect = {};
    GetWindowRect(toolbar_.hwnd(), &tb_rect);
    int toolbar_height = tb_rect.bottom - tb_rect.top;
    int toolbar_width = tb_rect.right - tb_rect.left;
    int toolbar_area_width = width - padding * 2;
    if (toolbar_area_width < 0) {
      toolbar_area_width = 0;
    }
    toolbar_width = std::min(toolbar_width, toolbar_area_width);
    place(toolbar_.hwnd(), padding, y + 2, toolbar_width, toolbar_height);

    y += toolbar_height;
  }
  if (show_address_bar_) {
    int address_width = width - padding * 2 - address_btn_width - 2;
    if (address_width < 120) {
      address_width = 120;
    }
    place(address_edit_, padding, y, address_width, address_height);
    place(address_go_btn_, padding + address_width, y, address_btn_width, address_height);
    SetEditMargins(address_edit_, 6, 6);
    SetEditVerticalRect(address_edit_, ui_font_, 2, 6, 6);
    y += address_height + gap;
  }

  int tabs_width = width - padding * 2;
  bool show_tabs = show_tab_control_ && tab_;
  bool show_filter = show_value && show_filter_bar_ && filter_edit_;
  bool show_tab_row = show_tabs || show_filter;
  if (show_tab_row) {
    y += 4;
    if (show_tabs && show_filter) {
      int available = std::max(0, tabs_width);
      int min_needed = kTabMinWidth + filter_min_width + filter_gap;
      if (available >= min_needed) {
        int target_width = ClampValue(available / 4, filter_min_width, filter_max_width);
        int filter_width = std::min(target_width, std::max(filter_min_width, available - kTabMinWidth - filter_gap));
        tabs_width = std::max(kTabMinWidth, available - filter_width - filter_gap);
        int filter_y = y + std::max(0, (tabs_height - filter_height) / 2);
        place(tab_, padding, y, tabs_width, tabs_height);
        place(filter_edit_, padding + tabs_width + filter_gap, filter_y, filter_width, filter_height);
        SetEditMargins(filter_edit_, 6, 6);
        SetEditVerticalRect(filter_edit_, ui_font_, 2, 6, 6);
        ShowWindow(filter_edit_, SW_SHOW);
      } else {
        show_filter = false;
      }
    }
    if (show_tabs && !show_filter) {
      place(tab_, padding, y, tabs_width, tabs_height);
      if (filter_edit_) {
        ShowWindow(filter_edit_, SW_HIDE);
      }
    } else if (!show_tabs && show_filter) {
      int available = std::max(0, tabs_width);
      int filter_width = ClampValue(available, filter_min_width, filter_max_width);
      int filter_y = y + std::max(0, (tabs_height - filter_height) / 2);
      int filter_x = padding + std::max(0, tabs_width - filter_width);
      place(filter_edit_, filter_x, filter_y, filter_width, filter_height);
      SetEditMargins(filter_edit_, 6, 6);
      SetEditVerticalRect(filter_edit_, ui_font_, 2, 6, 6);
      ShowWindow(filter_edit_, SW_SHOW);
    }
    y += tabs_height + gap;
  } else {
    if (tab_) {
      ShowWindow(tab_, SW_HIDE);
    }
    if (filter_edit_) {
      ShowWindow(filter_edit_, SW_HIDE);
    }
  }

  int status_top = height - status_height;
  int content_left = 0;
  int content_right = width;
  if (show_status_bar_ && status_bar_) {
    place(status_bar_, content_left, status_top, content_right - content_left, status_height);
    SendMessageW(status_bar_, WM_SIZE, 0, 0);
  }

  int history_splitter_height = show_history ? kHistorySplitterHeight : 0;
  int history_gap = show_history ? kHistoryGap : 0;
  int content_total_height = std::max(0, status_top - y);
  int min_history = kMinHistoryHeight;
  int max_history = std::max(min_history, content_total_height - kHistoryMaxPadding);
  int history_height = show_history ? ClampValue(history_height_, min_history, max_history) : 0;
  if (show_history) {
    history_height_ = history_height;
  }
  int history_top = status_top - history_height;
  if (show_history) {
    int history_width = content_right - content_left;
    place(history_label_, content_left, history_top, history_width, history_label_height);
    place(history_close_btn_, content_left + history_width - 18, history_top + 1, 16, 16);
    place(history_list_, content_left, history_top + history_label_height + 2, history_width, history_height - history_label_height - 2);
  }

  int splitter_bottom = show_history ? (history_top - history_gap) : history_top;
  int splitter_top = show_history ? (splitter_bottom - history_splitter_height) : history_top;
  if (show_history) {
    history_splitter_rect_.left = content_left;
    history_splitter_rect_.right = content_right;
    history_splitter_rect_.top = splitter_top;
    history_splitter_rect_.bottom = splitter_bottom;
  } else {
    history_splitter_rect_ = {};
  }
  int content_bottom = show_history ? splitter_top : (status_top - gap);
  int available_width = content_right - content_left;
  int min_tree = kMinTreeWidth;
  int min_list = kMinValueListWidth;
  int max_tree = std::max(min_tree, available_width - min_list - splitter_width);
  int tree_width = show_tree ? std::min(tree_width_, max_tree) : 0;
  tree_width = show_tree ? std::max(tree_width, min_tree) : 0;
  int list_x = show_tree ? (content_left + tree_width + splitter_width) : content_left;
  int list_width = content_right - list_x;
  int content_height = std::max(0, content_bottom - y);
  int tree_content_height = std::max(0, content_height - (show_tree ? tree_header_height : 0));
  if (show_tree) {
    place(tree_header_, content_left, y, tree_width, tree_header_height);
    place(tree_close_btn_, content_left + tree_width - 18, y + 2, 16, 16);
    place(tree_.hwnd(), content_left, y + tree_header_height, tree_width, tree_content_height);
    splitter_rect_.left = content_left + tree_width;
    splitter_rect_.right = splitter_rect_.left + splitter_width;
    splitter_rect_.top = y;
    splitter_rect_.bottom = y + content_height;
  } else {
    splitter_rect_ = {};
  }
  if (show_search) {
    place(search_results_list_, content_left, y, content_right - content_left, content_height);
  } else {
    place(value_list_.hwnd(), list_x, y, list_width, content_height);
  }

  UpdateStatus();
  if (!dragging_splitter) {
    UINT redraw_flags = RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_ERASE;
    RedrawWindow(hwnd_, nullptr, nullptr, redraw_flags);
  }
  drag_layout_valid_ = false;
}

void MainWindow::BuildImageLists() {
  if (tree_images_) {
    ImageList_Destroy(tree_images_);
    tree_images_ = nullptr;
  }
  if (list_images_) {
    ImageList_Destroy(list_images_);
    list_images_ = nullptr;
  }

  UINT dpi = GetWindowDpi(hwnd_);

  const int base_icon_size = kToolbarIconSize;
  const int icon_size = util::ScaleForDpi(base_icon_size, dpi);
  auto create_list = [&](int count) -> HIMAGELIST {
    HIMAGELIST list = ImageList_Create(icon_size, icon_size, ILC_COLOR32, count, 2);
    if (list) {
      int list_cx = 0;
      int list_cy = 0;
      if (ImageList_GetIconSize(list, &list_cx, &list_cy) && (list_cx != icon_size || list_cy != icon_size)) {
        ImageList_Destroy(list);
        list = ImageList_Create(icon_size, icon_size, ILC_COLOR32, count, 2);
      }
    }
    return list;
  };
  tree_images_ = create_list(4);
  list_images_ = create_list(6);
  ImageList_SetBkColor(tree_images_, CLR_NONE);
  ImageList_SetBkColor(list_images_, CLR_NONE);

  auto add_icon = [&](HIMAGELIST list, const wchar_t* name, int light_id, int dark_id) {
    HICON icon = LoadThemeIcon(name, light_id, dark_id, base_icon_size, dpi);
    if (icon) {
      ImageList_AddIcon(list, icon);
      DestroyIcon(icon);
    }
  };

  add_icon(tree_images_, L"folder.ico", IDI_ICON_LIGHT_FOLDER, IDI_ICON_DARK_FOLDER);
  add_icon(tree_images_, L"symlink.ico", IDI_ICON_LIGHT_SYMLINK, IDI_ICON_DARK_SYMLINK);
  add_icon(tree_images_, L"database.ico", IDI_ICON_LIGHT_DATABASE, IDI_ICON_DARK_DATABASE);
  add_icon(tree_images_, L"folder-sim.ico", IDI_ICON_LIGHT_FOLDER_SIM, IDI_ICON_DARK_FOLDER_SIM);
  add_icon(list_images_, L"folder.ico", IDI_ICON_LIGHT_FOLDER, IDI_ICON_DARK_FOLDER);
  add_icon(list_images_, L"symlink.ico", IDI_ICON_LIGHT_SYMLINK, IDI_ICON_DARK_SYMLINK);
  add_icon(list_images_, L"database.ico", IDI_ICON_LIGHT_DATABASE, IDI_ICON_DARK_DATABASE);
  add_icon(list_images_, L"folder-sim.ico", IDI_ICON_LIGHT_FOLDER_SIM, IDI_ICON_DARK_FOLDER_SIM);
  add_icon(list_images_, L"text.ico", IDI_ICON_LIGHT_TEXT, IDI_ICON_DARK_TEXT);
  add_icon(list_images_, L"binary.ico", IDI_ICON_LIGHT_BINARY, IDI_ICON_DARK_BINARY);
}

void MainWindow::CreateValueColumns() {
  value_columns_ = {
      {L"Name", 260, LVCFMT_LEFT}, {L"Type", 120, LVCFMT_LEFT}, {L"Data", 160, LVCFMT_LEFT}, {L"Default", 200, LVCFMT_LEFT}, {L"Read on boot", 110, LVCFMT_LEFT}, {L"Size", 70, LVCFMT_RIGHT}, {L"Date Modified", 140, LVCFMT_LEFT}, {L"Details", 160, LVCFMT_LEFT}, {L"Comment", 220, LVCFMT_LEFT},
  };
  value_column_widths_.clear();
  value_column_visible_.clear();
  value_column_widths_.reserve(value_columns_.size());
  value_column_visible_.reserve(value_columns_.size());
  for (const auto& column : value_columns_) {
    value_column_widths_.push_back(column.width);
    value_column_visible_.push_back(true);
  }
  if (saved_value_columns_loaded_) {
    auto patch_widths = [&](std::vector<int>& widths) {
      if (widths.size() == value_columns_.size() - 1) {
        widths.insert(widths.begin() + kValueColDefault, value_columns_[kValueColDefault].width);
      } else if (widths.size() == value_columns_.size() - 2) {
        widths.insert(widths.begin() + kValueColDefault, value_columns_[kValueColDefault].width);
        widths.push_back(value_columns_[kValueColComment].width);
      } else if (widths.size() == value_columns_.size() - 3) {
        widths.insert(widths.begin() + kValueColDefault, value_columns_[kValueColDefault].width);
        widths.push_back(value_columns_[kValueColDetails].width);
        widths.push_back(value_columns_[kValueColComment].width);
      } else if (widths.size() == value_columns_.size() - 4) {
        widths.insert(widths.begin() + kValueColDefault, value_columns_[kValueColDefault].width);
        widths.insert(widths.begin() + kValueColReadOnBoot, value_columns_[kValueColReadOnBoot].width);
        widths.push_back(value_columns_[kValueColDetails].width);
        widths.push_back(value_columns_[kValueColComment].width);
      }
    };
    auto patch_visible = [&](std::vector<bool>& visible) {
      if (visible.size() == value_columns_.size() - 1) {
        visible.insert(visible.begin() + kValueColDefault, true);
      } else if (visible.size() == value_columns_.size() - 2) {
        visible.insert(visible.begin() + kValueColDefault, true);
        visible.push_back(true);
      } else if (visible.size() == value_columns_.size() - 3) {
        visible.insert(visible.begin() + kValueColDefault, true);
        visible.push_back(true);
        visible.push_back(true);
      } else if (visible.size() == value_columns_.size() - 4) {
        visible.insert(visible.begin() + kValueColDefault, true);
        visible.insert(visible.begin() + kValueColReadOnBoot, true);
        visible.push_back(true);
        visible.push_back(true);
      }
    };
    patch_widths(saved_value_column_widths_);
    patch_visible(saved_value_column_visible_);
    for (size_t i = 0; i < value_columns_.size(); ++i) {
      if (i < saved_value_column_widths_.size() && saved_value_column_widths_[i] > 0) {
        value_column_widths_[i] = saved_value_column_widths_[i];
        value_columns_[i].width = saved_value_column_widths_[i];
      }
      if (i < saved_value_column_visible_.size()) {
        value_column_visible_[i] = saved_value_column_visible_[i];
      }
    }
  }
  ApplyValueColumns();
}

void MainWindow::CreateHistoryColumns() {
  history_columns_ = {
      {L"Time", 140, LVCFMT_LEFT},
      {L"Action", 280, LVCFMT_LEFT},
      {L"Old Data", 220, LVCFMT_LEFT},
      {L"New Data", 220, LVCFMT_LEFT},
  };
  history_column_widths_.clear();
  history_column_visible_.clear();
  history_column_widths_.reserve(history_columns_.size());
  history_column_visible_.reserve(history_columns_.size());
  for (const auto& column : history_columns_) {
    history_column_widths_.push_back(column.width);
    history_column_visible_.push_back(true);
  }
  ApplyHistoryColumns();
}

void MainWindow::ApplyValueColumns() {
  HWND list = value_list_.hwnd();
  if (!list) {
    return;
  }
  HWND header = ListView_GetHeader(list);
  int count = header ? Header_GetItemCount(header) : 0;
  for (int i = count - 1; i >= 0; --i) {
    ListView_DeleteColumn(list, i);
  }

  int insert_index = 0;
  for (size_t i = 0; i < value_columns_.size(); ++i) {
    if (i < value_column_visible_.size() && !value_column_visible_[i]) {
      continue;
    }
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT | LVCF_SUBITEM;
    col.pszText = const_cast<wchar_t*>(value_columns_[i].title.c_str());
    int width = value_column_widths_[i];
    if (width <= 0) {
      width = value_columns_[i].width;
    }
    col.cx = width;
    col.fmt = value_columns_[i].fmt;
    col.iSubItem = static_cast<int>(i);
    ListView_InsertColumn(list, insert_index++, &col);
  }

  UpdateListViewSort(list, value_sort_column_, value_sort_ascending_);
  header = ListView_GetHeader(list);
  if (header) {
    int size_display = FindListViewColumnBySubItem(list, kValueColSize);
    if (size_display >= 0) {
      HDITEMW item = {};
      item.mask = HDI_FORMAT;
      if (Header_GetItem(header, size_display, &item)) {
        item.fmt |= HDF_RIGHT;
        Header_SetItem(header, size_display, &item);
      }
    }
    if (!GetWindowSubclass(header, HeaderProc, kHeaderSubclassId, nullptr)) {
      SetWindowSubclass(header, HeaderProc, kHeaderSubclassId, reinterpret_cast<DWORD_PTR>(this));
    }
  }
}

void MainWindow::ApplyHistoryColumns() {
  if (!history_list_) {
    return;
  }
  HWND header = ListView_GetHeader(history_list_);
  int count = header ? Header_GetItemCount(header) : 0;
  for (int i = count - 1; i >= 0; --i) {
    ListView_DeleteColumn(history_list_, i);
  }

  int insert_index = 0;
  for (size_t i = 0; i < history_columns_.size(); ++i) {
    if (i < history_column_visible_.size() && !history_column_visible_[i]) {
      continue;
    }
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT | LVCF_SUBITEM;
    col.pszText = const_cast<wchar_t*>(history_columns_[i].title.c_str());
    int width = history_column_widths_[i];
    if (width <= 0) {
      width = history_columns_[i].width;
    }
    col.cx = width;
    col.fmt = history_columns_[i].fmt;
    col.iSubItem = static_cast<int>(i);
    ListView_InsertColumn(history_list_, insert_index++, &col);
  }

  UpdateListViewSort(history_list_, history_sort_column_, history_sort_ascending_);
  header = ListView_GetHeader(history_list_);
  if (header && !GetWindowSubclass(header, HeaderProc, kHeaderSubclassId, nullptr)) {
    SetWindowSubclass(header, HeaderProc, kHeaderSubclassId, reinterpret_cast<DWORD_PTR>(this));
  }
}

void MainWindow::CreateSearchColumns() {
  if (!search_results_list_) {
    return;
  }
  search_columns_ = {
      {L"Path", 320, LVCFMT_LEFT}, {L"Value", 180, LVCFMT_LEFT}, {L"Type", 110, LVCFMT_LEFT}, {L"Data", 360, LVCFMT_LEFT}, {L"Size", 80, LVCFMT_RIGHT}, {L"Data Modified", 150, LVCFMT_LEFT},
  };
  search_column_widths_.clear();
  search_column_visible_.clear();
  search_column_widths_.reserve(search_columns_.size());
  search_column_visible_.reserve(search_columns_.size());
  for (const auto& column : search_columns_) {
    search_column_widths_.push_back(column.width);
    search_column_visible_.push_back(true);
  }
  compare_columns_ = {
      {L"Path", 320, LVCFMT_LEFT},
      {L"Value", 180, LVCFMT_LEFT},
      {L"First Entry", 320, LVCFMT_LEFT},
      {L"Second Entry", 320, LVCFMT_LEFT},
  };
  compare_column_widths_.clear();
  compare_column_visible_.clear();
  compare_column_widths_.reserve(compare_columns_.size());
  compare_column_visible_.reserve(compare_columns_.size());
  for (const auto& column : compare_columns_) {
    compare_column_widths_.push_back(column.width);
    compare_column_visible_.push_back(true);
  }
  ApplySearchColumns(false);
  HWND header = ListView_GetHeader(search_results_list_);
  if (header && !GetWindowSubclass(header, HeaderProc, kHeaderSubclassId, nullptr)) {
    SetWindowSubclass(header, HeaderProc, kHeaderSubclassId, reinterpret_cast<DWORD_PTR>(this));
  }
}

void MainWindow::ApplySearchColumns(bool compare) {
  if (!search_results_list_) {
    return;
  }
  const auto& columns = compare ? compare_columns_ : search_columns_;
  auto& widths = compare ? compare_column_widths_ : search_column_widths_;
  auto& visible = compare ? compare_column_visible_ : search_column_visible_;
  HWND header = ListView_GetHeader(search_results_list_);
  int count = header ? Header_GetItemCount(header) : 0;
  for (int i = count - 1; i >= 0; --i) {
    ListView_DeleteColumn(search_results_list_, i);
  }

  int insert_index = 0;
  for (size_t i = 0; i < columns.size(); ++i) {
    if (i < visible.size() && !visible[i]) {
      continue;
    }
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT | LVCF_SUBITEM;
    col.pszText = const_cast<wchar_t*>(columns[i].title.c_str());
    int width = widths[i];
    if (width <= 0) {
      width = columns[i].width;
    }
    col.cx = width;
    col.fmt = columns[i].fmt;
    col.iSubItem = static_cast<int>(i);
    ListView_InsertColumn(search_results_list_, insert_index++, &col);
  }

  header = ListView_GetHeader(search_results_list_);
  if (header && !GetWindowSubclass(header, HeaderProc, kHeaderSubclassId, nullptr)) {
    SetWindowSubclass(header, HeaderProc, kHeaderSubclassId, reinterpret_cast<DWORD_PTR>(this));
  }
  compare_columns_active_ = compare;
}

void MainWindow::UpdateValueListForNode(RegistryNode* node) {
  if (updating_value_list_) {
    return;
  }
  updating_value_list_ = true;
  uint64_t generation = value_list_generation_.fetch_add(1) + 1;
  HWND list_hwnd = value_list_.hwnd();
  if (list_hwnd) {
    SendMessageW(list_hwnd, WM_SETREDRAW, FALSE, 0);
  }

  value_list_.Clear();
  current_key_count_ = 0;
  current_value_count_ = 0;
  if (!node) {
    if (list_hwnd && !jump_ui_batch_active_) {
      SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
      InvalidateRect(list_hwnd, nullptr, TRUE);
    }
    UpdateStatus();
    updating_value_list_ = false;
    value_list_loading_ = false;
    return;
  }

  RegistryNode snapshot = *node;
  std::wstring path = registry_path::Build(snapshot);
  RecordNavigation(path);
  std::wstring trace_path = NormalizeTraceKeyPath(path);
  if (trace_path.empty()) {
    trace_path = path;
  }
  std::wstring trace_path_lower = ToLower(trace_path);
  std::wstring default_path = NormalizeTraceKeyPathBasic(path);
  if (default_path.empty()) {
    default_path = path;
  }
  std::wstring default_path_lower = ToLower(default_path);
  bool is_reg_file = IsRegFileTabSelected();
  auto trace_data_list = is_reg_file ? std::vector<ActiveTrace>() : active_traces_;
  auto default_data_list = is_reg_file ? std::vector<ActiveDefault>() : active_defaults_;
  bool show_simulated_keys = show_simulated_keys_ && !is_reg_file;
  constexpr size_t kDateColumn = static_cast<size_t>(kValueColDate);
  bool include_dates = (value_sort_column_ == static_cast<int>(kDateColumn));
  if (kDateColumn < value_column_visible_.size() && value_column_visible_[kDateColumn]) {
    include_dates = true;
  }
  constexpr size_t kDetailsColumn = static_cast<size_t>(kValueColDetails);
  bool include_details = (value_sort_column_ == static_cast<int>(kDetailsColumn));
  if (kDetailsColumn < value_column_visible_.size() && value_column_visible_[kDetailsColumn]) {
    include_details = true;
  }

  if (list_hwnd && !jump_ui_batch_active_) {
    SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(list_hwnd, nullptr, TRUE);
  }
  UpdateStatus();
  updating_value_list_ = false;
  value_list_loading_ = true;

  int sort_column = value_sort_column_;
  bool sort_ascending = value_sort_ascending_;
  bool show_keys_in_list = show_keys_in_list_;
  if (show_keys_in_list) {
    EnsureHiveListLoaded();
  }
  if (!value_loader_.running()) {
    StartValueListWorker();
  }
  auto task = std::make_unique<ValueListTask>();
  task->generation = generation;
  task->snapshot = snapshot;
  task->trace_path_lower = std::move(trace_path_lower);
  task->default_path_lower = std::move(default_path_lower);
  task->include_dates = include_dates;
  task->sort_column = sort_column;
  task->sort_ascending = sort_ascending;
  task->show_keys_in_list = show_keys_in_list;
  task->include_details = include_details;
  task->show_simulated_keys = show_simulated_keys;
  task->include_all_value_data = sort_column == kValueColData || value_list_.HasFilter();
  task->hwnd = hwnd_;
  task->trace_data_list = std::move(trace_data_list);
  task->default_data_list = std::move(default_data_list);
  task->hive_roots = show_keys_in_list ? hive_roots_ : nullptr;
  value_loader_.Submit(std::move(task));
}

void MainWindow::ScheduleValueListRename(LPARAM kind, const std::wstring& name) {
  pending_value_list_kind_ = kind;
  pending_value_list_name_ = name;
}

void MainWindow::StartPendingValueListRename() {
  if (pending_value_list_name_.empty() || !value_list_.hwnd()) {
    return;
  }
  if (pending_value_list_kind_ == rowkind::kKey && !show_keys_in_list_) {
    pending_value_list_kind_ = 0;
    pending_value_list_name_.clear();
    return;
  }
  int index = -1;
  for (size_t i = 0; i < value_list_.RowCount(); ++i) {
    const ListRow* row = value_list_.RowAt(static_cast<int>(i));
    if (!row || row->kind != pending_value_list_kind_) {
      continue;
    }
    if (row->extra == pending_value_list_name_) {
      index = static_cast<int>(i);
      break;
    }
  }
  if (index >= 0 && IsWindowVisible(value_list_.hwnd())) {
    SetFocus(value_list_.hwnd());
    ListView_SetItemState(value_list_.hwnd(), -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetItemState(value_list_.hwnd(), index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(value_list_.hwnd(), index, FALSE);
    ListView_EditLabel(value_list_.hwnd(), index);
  }
  pending_value_list_kind_ = 0;
  pending_value_list_name_.clear();
}

void MainWindow::EnsureValueRowData(ListRow* row) {
  if (!row || row->kind != rowkind::kValue || row->data_ready) {
    return;
  }
  if (row->value_data_size == 0) {
    row->data.clear();
    row->data_ready = true;
    value_list_.InvalidateFilterCache(row);
    return;
  }
  if (!current_node_) {
    return;
  }

  ValueEntry entry;
  if (!RegistryProvider::QueryValue(*current_node_, row->extra, &entry)) {
    row->data.clear();
    row->data_ready = true;
    value_list_.InvalidateFilterCache(row);
    return;
  }

  row->value_type = entry.type;
  row->value_data_size = static_cast<DWORD>(entry.data.size());
  row->data = value_format::DisplayData(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
  row->data_ready = true;
  if (row->type.empty()) {
    row->type = value_format::TypeName(entry.type);
  }
  row->size_value = row->value_data_size;
  row->has_size = true;
  if (row->size.empty() && row->value_data_size > 0) {
    row->size = std::to_wstring(row->value_data_size);
  }
  value_list_.InvalidateFilterCache(row);
}

void MainWindow::UpdateAddressBar(RegistryNode* node) {
  if (!node || !address_edit_) {
    return;
  }
  std::wstring path = registry_path::Build(*node);
  SetWindowTextW(address_edit_, path.c_str());
  AddAddressHistory(path);
}

void MainWindow::EnableAddressAutoComplete() {
  if (!address_edit_ || address_autocomplete_) {
    return;
  }
  ::IAutoComplete2* autocomplete = nullptr;
  HRESULT hr = CoCreateInstance(CLSID_AutoComplete, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&autocomplete));
  if (FAILED(hr) || !autocomplete) {
    return;
  }

  ::IEnumString* source = new RegistryAddressEnum(this, address_edit_);
  hr = autocomplete->Init(address_edit_, source, nullptr, nullptr);
  if (FAILED(hr)) {
    source->Release();
    autocomplete->Release();
    return;
  }
  DWORD options = ACO_AUTOSUGGEST | ACO_AUTOAPPEND | ACO_UPDOWNKEYDROPSLIST | ACO_FILTERPREFIXES;
  autocomplete->SetOptions(options);
  address_autocomplete_ = autocomplete;
  address_autocomplete_source_ = source;
}

std::vector<std::wstring> MainWindow::BuildAddressSuggestions(const std::wstring& input) const {
  std::vector<std::wstring> items;
  std::wstring text = TrimWhitespace(input);
  for (auto& ch : text) {
    if (ch == L'/') {
      ch = L'\\';
    }
  }
  bool trailing_sep = !text.empty() && text.back() == L'\\';
  if (trailing_sep) {
    text.pop_back();
  }

  constexpr size_t kMaxSuggestions = 200;
  auto add_unique = [&](const std::wstring& value, std::unordered_set<std::wstring>* seen) {
    if (value.empty()) {
      return;
    }
    std::wstring key = ToLower(value);
    if (seen->insert(key).second) {
      items.push_back(value);
    }
  };

  size_t sep = text.find_last_of(L'\\');
  if (trailing_sep) {
    sep = text.size();
  }
  if (sep == std::wstring::npos) {
    std::unordered_set<std::wstring> seen;
    const std::wstring prefix = text;
    for (const auto& root : roots_) {
      if (prefix.empty() || StartsWithInsensitive(root.path_name, prefix)) {
        add_unique(root.path_name, &seen);
      }
    }
    struct RootAlias {
      const wchar_t* short_name;
      const wchar_t* full_name;
    };
    const RootAlias aliases[] = {
        {L"HKCR", L"HKEY_CLASSES_ROOT"}, {L"HKCU", L"HKEY_CURRENT_USER"}, {L"HKLM", L"HKEY_LOCAL_MACHINE"}, {L"HKU", L"HKEY_USERS"}, {L"HKCC", L"HKEY_CURRENT_CONFIG"},
    };
    for (const auto& alias : aliases) {
      if (prefix.empty() || StartsWithInsensitive(alias.short_name, prefix)) {
        add_unique(alias.short_name, &seen);
        add_unique(alias.full_name, &seen);
      }
    }
    if (items.size() > kMaxSuggestions) {
      items.resize(kMaxSuggestions);
    }
    return items;
  }

  std::wstring prefix = text.substr(0, sep);
  std::wstring partial;
  if (sep < text.size()) {
    partial = text.substr(sep + 1);
  }
  if (prefix.empty()) {
    prefix = text;
  }
  std::wstring normalized_prefix = NormalizeRegistryPath(prefix);
  std::wstring display_prefix = prefix;
  if (display_prefix.empty()) {
    display_prefix = normalized_prefix;
  }
  RegistryNode node;
  if (!ResolvePathToNode(normalized_prefix, &node)) {
    return items;
  }
  KeyInfo info = {};
  if (!RegistryProvider::QueryKeyInfo(node, &info)) {
    return items;
  }
  auto subkeys = RegistryProvider::EnumSubKeyNames(node, true);
  items.reserve(std::min(subkeys.size(), kMaxSuggestions));
  for (const auto& name : subkeys) {
    if (!partial.empty() && !StartsWithInsensitive(name, partial)) {
      continue;
    }
    items.push_back(display_prefix + L"\\" + name);
    if (items.size() >= kMaxSuggestions) {
      break;
    }
  }
  return items;
}

void MainWindow::ApplyAutoCompleteTheme() {
  if (!Theme::UseDarkMode()) {
    return;
  }
  AutoCompleteThemeContext ctx;
  ctx.owner = hwnd_;
  ctx.theme = &Theme::Current();
  EnumThreadWindows(GetCurrentThreadId(), ApplyAutoCompleteThemeProc, reinterpret_cast<LPARAM>(&ctx));
}

std::wstring MainWindow::NormalizeRegistryPath(const std::wstring& input) const {
  const std::wstring sid = util::GetCurrentUserSidString();
  std::wstring path = registry_path::Normalize(input, sid);
  auto strip_context = [&](const std::wstring& label) {
    if (label.empty()) {
      return;
    }
    const std::wstring prefix = label + L"\\";
    if (registry_path::StartsWith(path, prefix)) {
      path.erase(0, prefix.size());
    }
  };
  strip_context(TreeRootLabel());
  if (registry_mode_ == RegistryMode::kRemote) {
    strip_context(StripMachinePrefix(remote_machine_));
  }
  return registry_path::Normalize(path, sid);
}

std::wstring MainWindow::FormatRegistryPath(const std::wstring& path, RegistryPathFormat format) const {
  const std::wstring normalized = NormalizeRegistryPath(path);
  if (normalized.empty()) {
    return {};
  }
  std::wstring tree_root =
      registry_mode_ == RegistryMode::kLocal ? L"Computer" : TreeRootLabel();
  registry_path::Style style = registry_path::Style::kFull;
  switch (format) {
  case RegistryPathFormat::kAbbrev:
    style = registry_path::Style::kAbbreviated;
    break;
  case RegistryPathFormat::kRegedit:
    style = registry_path::Style::kRegeditAddress;
    break;
  case RegistryPathFormat::kRegFile:
    style = registry_path::Style::kRegFileHeader;
    break;
  case RegistryPathFormat::kPowerShellDrive:
    style = registry_path::Style::kPowerShellDrive;
    break;
  case RegistryPathFormat::kPowerShellProvider:
    style = registry_path::Style::kPowerShellProvider;
    break;
  case RegistryPathFormat::kEscaped:
    style = registry_path::Style::kEscaped;
    break;
  case RegistryPathFormat::kFull:
    break;
  }
  return registry_path::Format(normalized, style, tree_root);
}
bool MainWindow::FindNearestExistingPath(const std::wstring& path, std::wstring* nearest_path) const {
  return changes::FindNearestExistingPath(
      path,
      [this](const std::wstring& candidate) {
        RegistryNode node;
        KeyInfo info = {};
        return ResolvePathToNode(candidate, &node) &&
               RegistryProvider::QueryKeyInfo(node, &info);
      },
      nearest_path);
}

bool MainWindow::CreateRegistryPath(const std::wstring& path) {
  RegistryNode node;
  if (!ResolvePathToNode(path, &node)) {
    return false;
  }
  if (node.subkey.empty()) {
    return true;
  }
  std::vector<std::wstring> parts = registry_path::Split(node.subkey);
  RegistryNode current = node;
  current.subkey.clear();
  bool created = false;
  for (const auto& part : parts) {
    if (!RegistryProvider::CreateKey(current, part)) {
      return false;
    }
    created = true;
    if (current.subkey.empty()) {
      current.subkey = part;
    } else {
      current.subkey += L"\\" + part;
    }
  }
  if (created) {
    MarkOfflineDirty();
  }
  return true;
}

void MainWindow::UpdateStatus() {
  if (!status_bar_) {
    return;
  }
  RECT rc = {};
  GetClientRect(status_bar_, &rc);
  int total_width = rc.right - rc.left;
  if (total_width < 0) {
    total_width = 0;
  }
  LONG_PTR sb_style = GetWindowLongPtrW(status_bar_, GWL_STYLE);
  if (sb_style & SBARS_SIZEGRIP) {
    int grip = GetSystemMetrics(SM_CXVSCROLL);
    total_width = std::max(total_width - grip, 0);
  }
  auto measure_text = [&](HDC hdc, const std::wstring& text) -> int {
    if (!hdc || text.empty()) {
      return 0;
    }
    SIZE size = {};
    GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &size);
    return size.cx + 20;
  };
  if (IsSearchTabSelected()) {
    bool compare_selected = IsCompareTabSelected();
    int sel = TabCtrl_GetCurSel(tab_);
    int tab_index = SearchIndexFromTab(sel);
    size_t count = 0;
    if (tab_index >= 0 && static_cast<size_t>(tab_index) < search_tabs_.size()) {
      count = search_tabs_[static_cast<size_t>(tab_index)].results.size();
    }
    unsigned long long count_value = static_cast<unsigned long long>(count);
    wchar_t buffer[256] = {};
    if (compare_selected) {
      swprintf_s(buffer, L"Differences: %llu", count_value);
    } else if (search_running_) {
      uint64_t searched = search_progress_searched_.load();
      if (searched > 0) {
        swprintf_s(buffer, L"Searching... Results: ~%llu | Scanned: %llu", count_value, searched);
      } else {
        swprintf_s(buffer, L"Searching... Results: ~%llu", count_value);
      }
    } else if (search_duration_valid_ && search_duration_ms_ > 0) {
      double seconds = static_cast<double>(search_duration_ms_) / 1000.0;
      swprintf_s(buffer, L"Results: %llu (%.2fs)", count_value, seconds);
    } else {
      swprintf_s(buffer, L"Results: %llu", count_value);
    }
    int part = total_width;
    SendMessageW(status_bar_, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(&part));
    SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(buffer));
    return;
  }
  if (IsRegFileTabSelected()) {
    int sel = TabCtrl_GetCurSel(tab_);
    if (IsRegFileTabIndex(sel) && static_cast<size_t>(sel) < tabs_.size()) {
      const TabEntry& entry = tabs_[static_cast<size_t>(sel)];
      if (entry.reg_file_loading) {
        std::wstring label = entry.reg_file_label.empty() ? L"registry file" : entry.reg_file_label;
        std::wstring text = L"Loading " + label + L"...";
        int part = total_width;
        SendMessageW(status_bar_, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(&part));
        SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(text.c_str()));
        return;
      }
    }
  }

  int selected = ListView_GetSelectedCount(value_list_.hwnd());
  wchar_t buffer[256] = {};
  std::wstring keys_text;
  std::wstring values_text;
  std::wstring selected_text;
  std::wstring path_text;
  if (current_node_) {
    path_text = registry_path::Build(*current_node_);
  }
  swprintf_s(buffer, L"Keys: %d", current_key_count_);
  keys_text = buffer;
  swprintf_s(buffer, L"Values: %d", current_value_count_);
  values_text = buffer;
  swprintf_s(buffer, L"Selected: %d", selected);
  selected_text = buffer;

  HDC hdc = GetDC(status_bar_);
  HFONT old_font = nullptr;
  if (hdc && ui_font_) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }
  int values_width = measure_text(hdc, values_text);
  int selected_width = measure_text(hdc, selected_text);
  int keys_width = measure_text(hdc, keys_text);
  if (old_font) {
    SelectObject(hdc, old_font);
  }
  if (hdc) {
    ReleaseDC(status_bar_, hdc);
  }

  int part3 = total_width;
  int part2 = std::max(part3 - keys_width, 0);
  int part1 = std::max(part2 - selected_width, 0);
  int part0 = std::max(part1 - values_width, 0);
  int parts[4] = {part0, part1, part2, part3};
  SendMessageW(status_bar_, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));
  SendMessageW(status_bar_, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(path_text.c_str()));
  SendMessageW(status_bar_, SB_SETTEXTW, 1, reinterpret_cast<LPARAM>(values_text.c_str()));
  SendMessageW(status_bar_, SB_SETTEXTW, 2, reinterpret_cast<LPARAM>(selected_text.c_str()));
  SendMessageW(status_bar_, SB_SETTEXTW, 3, reinterpret_cast<LPARAM>(keys_text.c_str()));
}

bool MainWindow::IsSearchTabSelected() const {
  if (!tab_) {
    return false;
  }
  int index = TabCtrl_GetCurSel(tab_);
  return IsSearchTabIndex(index);
}

bool MainWindow::IsRegFileTabSelected() const {
  if (!tab_) {
    return false;
  }
  int index = TabCtrl_GetCurSel(tab_);
  return IsRegFileTabIndex(index);
}

bool MainWindow::IsCompareTabSelected() const {
  if (!tab_) {
    return false;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (!IsSearchTabIndex(index)) {
    return false;
  }
  int search_index = SearchIndexFromTab(index);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return false;
  }
  return search_tabs_[static_cast<size_t>(search_index)].is_compare;
}

bool MainWindow::IsSearchTabIndex(int index) const {
  if (index < 0) {
    return false;
  }
  if (static_cast<size_t>(index) >= tabs_.size()) {
    return false;
  }
  return tabs_[static_cast<size_t>(index)].kind == TabEntry::Kind::kSearch;
}

bool MainWindow::IsRegFileTabIndex(int index) const {
  if (index < 0) {
    return false;
  }
  if (static_cast<size_t>(index) >= tabs_.size()) {
    return false;
  }
  return tabs_[static_cast<size_t>(index)].kind == TabEntry::Kind::kRegFile;
}

int MainWindow::SearchIndexFromTab(int index) const {
  if (!IsSearchTabIndex(index)) {
    return -1;
  }
  return tabs_[static_cast<size_t>(index)].search_index;
}

int MainWindow::FindFirstRegistryTabIndex() const {
  for (size_t i = 0; i < tabs_.size(); ++i) {
    if (tabs_[i].kind == TabEntry::Kind::kRegistry) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

void MainWindow::SyncRegFileTabSelection() {
  if (!tab_) {
    return;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (!IsRegFileTabIndex(index)) {
    return;
  }
  if (static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  const TabEntry& entry = tabs_[static_cast<size_t>(index)];
  registry_mode_ = RegistryMode::kLocal;
  std::vector<RegistryRootEntry> roots;
  roots.reserve(entry.reg_file_roots.size());
  for (const auto& root : entry.reg_file_roots) {
    if (!root.root) {
      continue;
    }
    RegistryRootEntry reg_root;
    reg_root.root = root.root;
    reg_root.display_name = root.name;
    reg_root.path_name = root.name;
    reg_root.subkey_prefix = L"";
    reg_root.group = RegistryRootGroup::kStandard;
    roots.push_back(std::move(reg_root));
  }
  ApplyRegistryRoots(roots);
}

void MainWindow::UpdateSearchResultsView() {
  if (!search_results_list_) {
    return;
  }
  int sel = TabCtrl_GetCurSel(tab_);
  if (!IsSearchTabIndex(sel)) {
    ListView_SetItemCountEx(search_results_list_, 0, LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
    search_results_view_tab_index_ = -1;
    return;
  }
  int search_index = SearchIndexFromTab(sel);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return;
  }
  EnsureSearchTabResultsLoaded(search_index);
  bool force_redraw = (search_results_view_tab_index_ != sel);
  search_results_view_tab_index_ = sel;
  auto& tab = search_tabs_[static_cast<size_t>(search_index)];
  bool compare = tab.is_compare;
  if (compare != compare_columns_active_) {
    ApplySearchColumns(compare);
    force_redraw = true;
  }
  int max_sort_col = compare ? 3 : 5;
  if (tab.sort_column > max_sort_col) {
    tab.sort_column = -1;
  }
  UpdateListViewSort(search_results_list_, tab.sort_column, tab.sort_ascending);
  HWND header = ListView_GetHeader(search_results_list_);
  if (header) {
    InvalidateRect(header, nullptr, TRUE);
  }
  size_t count = tab.results.size();
  size_t old_count = tab.last_ui_count;
  if (force_redraw || count != old_count) {
    ListView_SetItemCountEx(search_results_list_, static_cast<int>(count), LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);
    if (force_redraw || count < old_count) {
      InvalidateRect(search_results_list_, nullptr, TRUE);
    } else if (count > old_count) {
      int first = static_cast<int>(old_count);
      int last = static_cast<int>(count - 1);
      ListView_RedrawItems(search_results_list_, first, last);
    }
    tab.last_ui_count = count;
  }
}

void MainWindow::StartSearch(const SearchDialogResult& options) {
  if (options.criteria.query.empty()) {
    ui::ShowWarning(hwnd_, L"Enter text to find.");
    return;
  }

  bool matcher_ok = true;
  TextMatcher matcher(options.criteria.query, options.criteria.use_regex, options.criteria.match_case, options.criteria.match_whole, &matcher_ok);
  if (!matcher_ok) {
    ui::ShowError(hwnd_, L"Invalid regex.");
    return;
  }

  bool want_registry = options.search_standard_hives || options.search_registry_root;
  bool want_trace = options.search_trace_values && !active_traces_.empty();
  std::wstring registry_scope_path;
  std::wstring scope_path;
  if (options.scope == SearchScope::kCurrentKey) {
    if (!options.start_key.empty()) {
      registry_scope_path = options.start_key;
      scope_path = NormalizeRegistryPath(options.start_key);
    } else if (current_node_) {
      registry_scope_path = registry_path::Build(*current_node_);
      scope_path = NormalizeRegistryPath(registry_scope_path);
    } else {
      ui::ShowError(hwnd_, L"Select a starting key first.");
      return;
    }
  }

  std::vector<RegistryNode> start_nodes;
  if (want_registry) {
    if (options.scope == SearchScope::kCurrentKey) {
      if (!registry_scope_path.empty()) {
        RegistryNode node;
        if (ResolvePathToNode(registry_scope_path, &node)) {
          start_nodes.push_back(node);
        } else {
          std::wstring normalized = NormalizeRegistryPath(registry_scope_path);
          if (!normalized.empty() && ResolvePathToNode(normalized, &node)) {
            start_nodes.push_back(node);
          } else {
            ui::ShowError(hwnd_, L"Starting key path was not found.");
            return;
          }
        }
      } else if (current_node_) {
        start_nodes.push_back(*current_node_);
      } else {
        ui::ShowError(hwnd_, L"Select a starting key first.");
        return;
      }
    } else {
      std::unordered_set<std::wstring> seen;
      auto add_root = [&](const RegistryRootEntry& entry) {
        std::wstring key = ToLower(entry.path_name.empty() ? entry.display_name : entry.path_name);
        if (key.empty()) {
          return;
        }
        if (!seen.insert(key).second) {
          return;
        }
        RegistryNode node;
        node.root = entry.root;
        node.root_name = entry.path_name;
        node.subkey = entry.subkey_prefix;
        start_nodes.push_back(std::move(node));
      };

      if (options.search_standard_hives) {
        for (const auto& path : options.root_paths) {
          for (const auto& root : roots_) {
            if (_wcsicmp(root.path_name.c_str(), path.c_str()) == 0 || _wcsicmp(root.display_name.c_str(), path.c_str()) == 0) {
              add_root(root);
              break;
            }
          }
        }
        if (start_nodes.empty()) {
          for (const auto& root : roots_) {
            if (root.group == RegistryRootGroup::kStandard) {
              add_root(root);
            }
          }
        }
      }
      if (options.search_registry_root) {
        for (const auto& root : roots_) {
          if (_wcsicmp(root.path_name.c_str(), L"REGISTRY") == 0 || _wcsicmp(root.display_name.c_str(), L"REGISTRY") == 0) {
            add_root(root);
            break;
          }
        }
      }
    }
  }

  if (want_registry && start_nodes.empty()) {
    ui::ShowError(hwnd_, L"Select at least one top-level key.");
    return;
  }
  if (!want_registry && !want_trace) {
    return;
  }
  if (!tab_) {
    return;
  }

  CancelSearch();

  search::Criteria criteria = options.criteria;
  criteria.start_nodes = start_nodes;
  criteria.exclude_paths = options.exclude_paths;

  std::wstring label = L"Find";
  if (!criteria.query.empty()) {
    label = L"Find: " + criteria.query;
    constexpr size_t kMaxLabel = 48;
    if (label.size() > kMaxLabel) {
      label.resize(kMaxLabel - 3);
      label.append(L"...");
    }
  }

  int tab_index = -1;
  int search_index = -1;
  bool reuse_tab = options.result_mode == SearchResultMode::kReuseTab;
  if (reuse_tab) {
    int sel = TabCtrl_GetCurSel(tab_);
    int candidate = IsSearchTabIndex(sel) ? sel : active_search_tab_index_;
    if (IsSearchTabIndex(candidate)) {
      int index = SearchIndexFromTab(candidate);
      if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size() && !search_tabs_[static_cast<size_t>(index)].is_compare) {
        tab_index = candidate;
        search_index = index;
      }
    }
  }

  if (search_index >= 0) {
    SearchTab& tab = search_tabs_[static_cast<size_t>(search_index)];
    tab.label = label;
    tab.results.clear();
    tab.last_ui_count = 0;
    tab.is_compare = false;
    tab.sort_dirty = false;
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(tab.label.c_str());
    TabCtrl_SetItem(tab_, tab_index, &item);
  } else {
    SearchTab tab;
    tab.label = label;
    tab.is_compare = false;
    search_tabs_.push_back(std::move(tab));
    search_index = static_cast<int>(search_tabs_.size() - 1);
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(search_tabs_.back().label.c_str());
    tab_index = TabCtrl_GetItemCount(tab_);
    TabCtrl_InsertItem(tab_, tab_index, &item);
    tabs_.push_back({TabEntry::Kind::kSearch, search_index});
  }

  UpdateTabWidth();
  TabCtrl_SetCurSel(tab_, tab_index);
  active_search_tab_index_ = tab_index;
  search_results_view_tab_index_ = -1;
  search_progress_searched_.store(0);
  search_progress_total_.store(0);
  search_progress_percent_ = 0;
  search_progress_posted_.store(false);
  search_posted_.store(false);
  {
    std::lock_guard<std::mutex> lock(search_mutex_);
    std::vector<PendingSearchResult>().swap(search_pending_);
  }
  search_last_refresh_tick_ = 0;
  search_start_tick_ = GetTickCount64();
  search_duration_ms_ = 0;
  search_duration_valid_ = false;
  search_running_ = true;

  if (search_progress_) {
    SendMessageW(search_progress_, PBM_SETMARQUEE, TRUE, 30);
  }

  ApplyViewVisibility();
  UpdateSearchResultsView();
  UpdateStatus();

  std::vector<ActiveTrace> traces = active_traces_;
  std::vector<std::wstring> exclude_paths = options.exclude_paths;
  std::wstring scope_lower = ToLower(scope_path);
  bool scope_recursive = criteria.recursive;
  bool trace_enabled = want_trace;
  bool registry_enabled = want_registry && !criteria.start_nodes.empty();

  const uint64_t generation = search_session_.Start(
      [this, criteria, traces, exclude_paths, scope_lower, scope_recursive,
       trace_enabled, registry_enabled, matcher](
          uint64_t generation, std::atomic_bool& cancel) mutable {
    auto should_stop = [&]() { return cancel.load(); };

    std::vector<PendingSearchResult> batch;
    batch.reserve(kSearchQueueBatch);
    std::mutex batch_mutex;

    auto flush = [&]() {
      std::vector<PendingSearchResult> pending;
      {
        std::lock_guard<std::mutex> lock(batch_mutex);
        if (batch.empty()) {
          return;
        }
        pending.swap(batch);
      }
      {
        std::lock_guard<std::mutex> lock(search_mutex_);
        search_pending_.reserve(search_pending_.size() + pending.size());
        for (auto& item : pending) {
          search_pending_.push_back(std::move(item));
        }
      }
      if (!search_posted_.exchange(true)) {
        PostMessageW(hwnd_, kSearchResultsMessage, static_cast<WPARAM>(generation), 0);
      }
    };

    auto queue_result = [&](search::Result&& result) {
      bool should_flush = false;
      {
        std::lock_guard<std::mutex> lock(batch_mutex);
        PendingSearchResult pending;
        pending.generation = generation;
        pending.result = std::move(result);
        batch.push_back(std::move(pending));
        should_flush = batch.size() >= kSearchQueueBatch;
      }
      if (should_flush) {
        flush();
      }
    };

    auto is_excluded = [&](const std::wstring& path) {
      if (exclude_paths.empty()) {
        return false;
      }
      for (const auto& exclude : exclude_paths) {
        if (exclude.empty()) {
          continue;
        }
        if (FindStringOrdinal(FIND_FROMSTART, path.c_str(), static_cast<int>(path.size()), exclude.c_str(), static_cast<int>(exclude.size()), TRUE) >= 0) {
          return true;
        }
      }
      return false;
    };

    auto key_in_scope = [&](const std::wstring& key_lower) {
      if (scope_lower.empty()) {
        return true;
      }
      if (key_lower == scope_lower) {
        return true;
      }
      if (!scope_recursive) {
        return false;
      }
      if (key_lower.size() <= scope_lower.size()) {
        return false;
      }
      if (key_lower.compare(0, scope_lower.size(), scope_lower) != 0) {
        return false;
      }
      return key_lower[scope_lower.size()] == L'\\';
    };

    if (trace_enabled) {
      for (const auto& trace : traces) {
        if (should_stop()) {
          break;
        }
        if (!trace.data) {
          continue;
        }
        std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
        for (const auto& key_path : trace.data->key_paths) {
          if (should_stop()) {
            break;
          }
          if (key_path.empty()) {
            continue;
          }
          if (is_excluded(key_path)) {
            continue;
          }
          std::wstring key_lower = ToLower(key_path);
          if (!trace.selection ||
              !trace::IncludesKey(*trace.selection, key_lower)) {
            continue;
          }
          if (!key_in_scope(key_lower)) {
            continue;
          }
          std::wstring key_name = registry_path::Leaf(key_path);

          if (criteria.search_keys) {
            TextMatch match = matcher.Match(key_name);
            if (match.matched) {
              search::Result result;
              result.key_path = key_path;
              result.key_name = key_name;
              result.type_text = L"Trace Key";
              result.is_key = true;
              size_t path_start = key_path.size() >= key_name.size() ? key_path.size() - key_name.size() : 0;
              result.match_field = search::MatchField::kPath;
              result.match_start = static_cast<int>(path_start + match.start);
              result.match_length = static_cast<int>(match.length);
              queue_result(std::move(result));
            }
          }

          if (criteria.search_values) {
            auto it = trace.data->values_by_key.find(key_lower);
            if (it != trace.data->values_by_key.end()) {
              for (const auto& value_name : it->second.values_display) {
                if (should_stop()) {
                  break;
                }
                std::wstring value_lower = ToLower(value_name);
                if (!trace.selection ||
                    !trace::IncludesValue(*trace.selection, key_lower,
                                          value_lower)) {
                  continue;
                }
                std::wstring display = value_name.empty() ? L"(Default)" : value_name;
                TextMatch match = matcher.Match(display);
                if (!match.matched) {
                  continue;
                }
                search::Result result;
                result.key_path = key_path;
                result.key_name = key_name;
                result.value_name = value_name;
                result.display_name = display;
                result.type_text = L"Trace Value";
                result.is_key = false;
                result.match_field = search::MatchField::kName;
                result.match_start = static_cast<int>(match.start);
                result.match_length = static_cast<int>(match.length);
                queue_result(std::move(result));
              }
            }
          }
        }
      }
      flush();
    }

    if (!should_stop() && registry_enabled) {
      std::atomic<uint64_t> last_progress_tick{0};
      auto progress_cb = [&](uint64_t searched, uint64_t total) {
        search_progress_searched_.store(searched);
        search_progress_total_.store(total);
        uint64_t now = GetTickCount64();
        uint64_t last = last_progress_tick.load();
        if (now - last < kSearchProgressUiMs && searched < total) {
          return;
        }
        if (last_progress_tick.compare_exchange_strong(last, now)) {
          if (!search_progress_posted_.exchange(true)) {
            PostMessageW(hwnd_, kSearchProgressMessage, static_cast<WPARAM>(generation), 0);
          }
        }
      };
      bool ok = search::Run(
          criteria, &cancel,
          [&](search::Result&& result) -> bool {
            if (should_stop()) {
              return false;
            }
            queue_result(std::move(result));
            return !should_stop();
          },
          progress_cb, false);
      flush();
      if (!ok) {
        PostMessageW(hwnd_, kSearchFailedMessage, static_cast<WPARAM>(generation), 0);
        return;
      }
    }

    flush();
    PostMessageW(hwnd_, kSearchFinishedMessage, static_cast<WPARAM>(generation), 0);
  });
  search_tabs_[static_cast<size_t>(search_index)].generation = generation;
}

void MainWindow::StartReplace(const ReplaceDialogResult& options) {
  if (read_only_) {
    ui::ShowWarning(hwnd_, L"Read-only mode is enabled.");
    return;
  }
  if (options.find_text.empty()) {
    return;
  }

  RegistryNode start;
  if (!options.start_key.empty()) {
    if (!ResolvePathToNode(options.start_key, &start)) {
      ui::ShowError(hwnd_, L"Starting key path was not found.");
      return;
    }
  } else if (current_node_) {
    start = *current_node_;
  } else {
    ui::ShowError(hwnd_, L"Select a starting key first.");
    return;
  }

  search::Replacer matcher(options);
  if (!matcher.valid()) {
    ui::ShowError(hwnd_, L"Invalid replace pattern.");
    return;
  }

  if (replace_result_pending_) {
    ui::ShowWarning(hwnd_, L"Replace is already running.");
    return;
  }

  const HWND hwnd = hwnd_;
  replace_result_pending_ = true;
  replace_session_.Start(
      [this, start, options, matcher, hwnd](
          uint64_t generation, std::atomic_bool& cancel) mutable {
        auto payload = std::make_unique<ReplacePayload>();
        payload->generation = generation;
        std::vector<RegistryNode> stack;
        stack.push_back(start);

        while (!stack.empty() && !cancel.load()) {
          RegistryNode node = std::move(stack.back());
          stack.pop_back();

          std::vector<ValueEntry> values;
          RegistryProvider::KeyEnumResult enum_result;
          bool values_reserved = false;
          RegistryProvider::EnumKeyStreaming(
              node, true, true, false, &enum_result,
              [&](const ValueInfo& info, const BYTE* data, DWORD data_size) {
                if (!values_reserved) {
                  if (enum_result.info_valid) {
                    values.reserve(enum_result.info.value_count);
                  }
                  values_reserved = true;
                }
                ValueEntry value;
                value.name = info.name;
                value.type = info.type;
                if (data_size > 0 && data) {
                  value.data.assign(data, data + data_size);
                }
                values.push_back(std::move(value));
                return !cancel.load();
              },
              {});

          for (const auto& value : values) {
            if (cancel.load()) {
              break;
            }

            std::wstring current_name = value.name;
            std::wstring replaced_name;
            if (!current_name.empty() &&
                matcher.Replace(current_name, &replaced_name) &&
                replaced_name != current_name) {
              if (replaced_name.empty()) {
                continue;
              }
              std::wstring unique =
                  MakeUniqueValueName(node, replaced_name);
              if (!RegistryProvider::RenameValue(node, current_name, unique)) {
                ++payload->failures;
              } else {
                ReplacePayload::Change change;
                change.undo.type =
                    changes::UndoOperation::Type::kRenameValue;
                change.undo.node = node;
                change.undo.name = current_name;
                change.undo.new_name = unique;
                change.history.action = L"Rename value " + current_name;
                change.history.old_data = current_name;
                change.history.new_data = unique;
                change.history.key_path = registry_path::Build(node);
                change.history.value_name = unique;
                payload->changes.push_back(std::move(change));
                current_name = unique;
              }
            }

            if (value.type != REG_SZ &&
                value.type != REG_EXPAND_SZ &&
                value.type != REG_MULTI_SZ) {
              continue;
            }

            std::vector<BYTE> new_data = value.data;
            bool changed = false;
            if (value.type == REG_MULTI_SZ) {
              auto parts = value_format::MultiStringItems(value.data);
              for (auto& part : parts) {
                std::wstring updated;
                if (matcher.Replace(part, &updated) && updated != part) {
                  part = std::move(updated);
                  changed = true;
                }
              }
              if (changed) {
                new_data = value_format::MultiStringData(parts);
              }
            } else {
              const std::wstring text = value_format::Data(
                  value.type, value.data.data(),
                  static_cast<DWORD>(value.data.size()));
              std::wstring updated;
              if (matcher.Replace(text, &updated) && updated != text) {
                new_data = value_format::StringData(updated);
                changed = true;
              }
            }

            if (!changed) {
              continue;
            }
            if (!RegistryProvider::SetValue(node, current_name, value.type,
                                            new_data)) {
              ++payload->failures;
              continue;
            }

            ValueEntry old_value = value;
            old_value.name = current_name;
            ValueEntry new_value = value;
            new_value.name = current_name;
            new_value.data = new_data;
            ReplacePayload::Change change;
            change.undo.type =
                changes::UndoOperation::Type::kModifyValue;
            change.undo.node = node;
            change.undo.old_value = old_value;
            change.undo.new_value = new_value;
            change.history.action = L"Modify value " + current_name;
            change.history.old_data = value_format::Data(
                value.type, value.data.data(),
                static_cast<DWORD>(value.data.size()));
            change.history.new_data = value_format::Data(
                value.type, new_data.data(),
                static_cast<DWORD>(new_data.size()));
            change.history.key_path = registry_path::Build(node);
            change.history.value_name = current_name;
            change.history.revert_kind =
                HistoryEntry::RevertKind::kSetValue;
            change.history.revert_value = std::move(old_value);
            payload->changes.push_back(std::move(change));
          }

          if (options.recursive && !cancel.load()) {
            auto subkeys =
                RegistryProvider::EnumSubKeyNames(node, false);
            for (const auto& name : subkeys) {
              stack.push_back(MakeChildNode(node, name));
            }
          }
        }

        payload->cancelled = cancel.load();
        if (hwnd && IsWindow(hwnd) &&
            PostMessageW(hwnd, kReplaceReadyMessage,
                         static_cast<WPARAM>(generation),
                         reinterpret_cast<LPARAM>(payload.get()))) {
          ReleasePostedPayload(payload);
        }
      });
}

void MainWindow::ApplyReplacePayload(ReplacePayload* payload) {
  if (!payload) {
    return;
  }
  std::unique_ptr<ReplacePayload> owned(payload);
  if (!replace_session_.IsCurrent(owned->generation)) {
    return;
  }
  replace_session_.Join();
  replace_result_pending_ = false;
  CommitReplacePayload(std::move(owned), true);
}

void MainWindow::CommitReplacePayload(
    std::unique_ptr<ReplacePayload> payload, bool show_failures) {
  if (!payload) {
    return;
  }
  for (auto& change : payload->changes) {
    PushUndo(std::move(change.undo));
    AppendHistoryEntry(std::move(change.history));
  }
  if (!payload->changes.empty()) {
    MarkOfflineDirty();
  }
  if (current_node_) {
    UpdateValueListForNode(current_node_);
  }
  if (show_failures && payload->failures > 0) {
    const std::wstring message =
        L"Replace finished with some failures.\nReplaced: " +
        std::to_wstring(payload->changes.size()) + L"\nFailed: " +
        std::to_wstring(payload->failures);
    ui::ShowError(hwnd_, message);
  }
}

void MainWindow::StopReplace() {
  replace_session_.CancelAndJoin();
  MSG message = {};
  while (PeekMessageW(&message, hwnd_, kReplaceReadyMessage,
                      kReplaceReadyMessage, PM_REMOVE)) {
    CommitReplacePayload(
        std::unique_ptr<ReplacePayload>(
            reinterpret_cast<ReplacePayload*>(message.lParam)),
        false);
  }
  replace_result_pending_ = false;
}

void MainWindow::CancelSearch() {
  search_session_.CancelAndJoin();
  search_running_ = false;
  search_start_tick_ = 0;
  search_duration_ms_ = 0;
  search_duration_valid_ = false;
  search_progress_percent_ = 0;
  search_progress_searched_.store(0);
  search_progress_total_.store(0);
  search_progress_posted_.store(false);
  search_posted_.store(false);
  {
    std::lock_guard<std::mutex> lock(search_mutex_);
    std::vector<PendingSearchResult>().swap(search_pending_);
  }
  if (search_progress_) {
    SendMessageW(search_progress_, PBM_SETMARQUEE, FALSE, 0);
  }
  ApplyViewVisibility();
  UpdateStatus();
}

void MainWindow::CloseSearchTab(int tab_index) {
  if (!tab_ || !IsSearchTabIndex(tab_index)) {
    return;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (tab_index >= count) {
    return;
  }
  if (search_running_ && active_search_tab_index_ == tab_index) {
    CancelSearch();
  }
  int search_index = SearchIndexFromTab(tab_index);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return;
  }

  bool was_active = TabCtrl_GetCurSel(tab_) == tab_index;

  search_tabs_.erase(search_tabs_.begin() + search_index);
  tabs_.erase(tabs_.begin() + tab_index);
  for (auto& entry : tabs_) {
    if (entry.kind == TabEntry::Kind::kSearch && entry.search_index > search_index) {
      --entry.search_index;
    }
  }
  TabCtrl_DeleteItem(tab_, tab_index);
  if (active_search_tab_index_ == tab_index) {
    active_search_tab_index_ = -1;
  } else if (active_search_tab_index_ > tab_index) {
    --active_search_tab_index_;
  }

  int new_count = TabCtrl_GetItemCount(tab_);
  if (was_active && new_count > 0) {
    int next = std::min(tab_index, new_count - 1);
    TabCtrl_SetCurSel(tab_, next);
  }
  UpdateTabWidth();
  UpdateSearchResultsView();
  ApplyViewVisibility();
  UpdateStatus();
}

void MainWindow::SortValueList(int column, bool toggle) {
  if (column < 0 || static_cast<size_t>(column) >= value_columns_.size()) {
    return;
  }
  if (toggle) {
    if (value_sort_column_ == column) {
      value_sort_ascending_ = !value_sort_ascending_;
    } else {
      value_sort_column_ = column;
      value_sort_ascending_ = true;
    }
  } else {
    value_sort_column_ = column;
  }

  if (value_list_loading_ && current_node_) {
    UpdateValueListForNode(current_node_);
    return;
  }

  auto& rows = value_list_.rows();
  if (value_sort_column_ == kValueColData) {
    bool needs_data = false;
    for (const auto& row : rows) {
      if (row.kind == rowkind::kValue && !row.data_ready) {
        needs_data = true;
        break;
      }
    }
    if (needs_data && current_node_) {
      UpdateValueListForNode(current_node_);
      return;
    }
    for (auto& row : rows) {
      EnsureValueRowData(&row);
    }
  }
  SortValueRows(&rows, value_sort_column_, value_sort_ascending_);
  value_list_.RebuildFilter();

  HWND header = ListView_GetHeader(value_list_.hwnd());
  if (header) {
    UpdateListViewSort(value_list_.hwnd(), value_sort_column_, value_sort_ascending_);
    InvalidateRect(header, nullptr, TRUE);
  }
}

void MainWindow::SortHistoryList(int column, bool toggle) {
  if (!history_list_ || column < 0) {
    return;
  }
  if (toggle) {
    if (history_sort_column_ == column) {
      history_sort_ascending_ = !history_sort_ascending_;
    } else {
      history_sort_column_ = column;
      history_sort_ascending_ = true;
    }
  } else {
    history_sort_column_ = column;
  }

  change_history_.Sort(history_sort_column_, history_sort_ascending_);
  RebuildHistoryList();

  HWND header = ListView_GetHeader(history_list_);
  if (header) {
    UpdateListViewSort(history_list_, history_sort_column_, history_sort_ascending_);
    InvalidateRect(header, nullptr, TRUE);
  }
}

void MainWindow::SortSearchTabResults(SearchTab* tab) {
  if (!tab || tab->sort_column < 0) {
    if (tab) {
      tab->sort_dirty = false;
    }
    return;
  }
  if (!tab->is_compare && tab->sort_column == 3) {
    for (auto& result : tab->results) {
      EnsureSearchResultDataLoaded(&result);
    }
  }
  search::SortResults(&tab->results, tab->sort_column,
                      tab->sort_ascending, tab->is_compare);
  tab->sort_dirty = false;
}

void MainWindow::SortSearchResults(int column, bool toggle) {
  if (!search_results_list_ || column < 0) {
    return;
  }
  int sel = TabCtrl_GetCurSel(tab_);
  int index = SearchIndexFromTab(sel);
  if (index < 0 || static_cast<size_t>(index) >= search_tabs_.size()) {
    return;
  }
  EnsureSearchTabResultsLoaded(index);
  auto& tab = search_tabs_[static_cast<size_t>(index)];
  if (toggle) {
    if (tab.sort_column == column) {
      tab.sort_ascending = !tab.sort_ascending;
    } else {
      tab.sort_column = column;
      tab.sort_ascending = true;
    }
  } else {
    tab.sort_column = column;
  }
  SortSearchTabResults(&tab);
  UpdateListViewSort(search_results_list_, tab.sort_column, tab.sort_ascending);
  HWND header = ListView_GetHeader(search_results_list_);
  if (header) {
    InvalidateRect(header, nullptr, TRUE);
  }
  InvalidateRect(search_results_list_, nullptr, TRUE);
}

void MainWindow::ClearHistoryItems(bool delete_cache) {
  if (!history_list_) {
    return;
  }
  change_history_.entries().clear();
  ListView_DeleteAllItems(history_list_);

  if (delete_cache) {
    std::wstring path = HistoryCachePath();
    if (!path.empty()) {
      DeleteFileW(path.c_str());
    }
  }
}

void MainWindow::RebuildHistoryList() {
  if (!history_list_) {
    return;
  }
  SendMessageW(history_list_, WM_SETREDRAW, FALSE, 0);
  ListView_DeleteAllItems(history_list_);

  int index = 0;
  for (const auto& entry : change_history_.entries()) {
    LVITEMW item = {};
    item.mask = LVIF_TEXT | LVIF_PARAM;
    item.iItem = index;
    item.pszText = const_cast<wchar_t*>(entry.time_text.c_str());
    item.lParam = static_cast<LPARAM>(index);
    int inserted = ListView_InsertItem(history_list_, &item);
    if (inserted >= 0) {
      ListView_SetItemText(history_list_, inserted, 1, const_cast<wchar_t*>(entry.action.c_str()));
      ListView_SetItemText(history_list_, inserted, 2, const_cast<wchar_t*>(entry.old_data.c_str()));
      ListView_SetItemText(history_list_, inserted, 3, const_cast<wchar_t*>(entry.new_data.c_str()));
    }
    ++index;
  }

  SendMessageW(history_list_, WM_SETREDRAW, TRUE, 0);
  InvalidateRect(history_list_, nullptr, TRUE);
}

void MainWindow::ResetNavigationState() {
  nav_history_.clear();
  nav_index_ = -1;
  nav_is_programmatic_ = false;
  UpdateNavigationButtons();
}

void MainWindow::UpdateTabText(const std::wstring& text) {
  if (!tab_) {
    return;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (!IsSearchTabIndex(index) && !IsRegFileTabIndex(index)) {
    // keep the current registry tab label up to date
  } else {
    index = FindFirstRegistryTabIndex();
  }
  if (index < 0) {
    return;
  }
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(text.c_str());
  TabCtrl_SetItem(tab_, index, &item);
  UpdateTabWidth();
  InvalidateRect(tab_, nullptr, FALSE);
}

void MainWindow::MarkOfflineDirty() {
  if (IsRegFileTabSelected()) {
    int index = TabCtrl_GetCurSel(tab_);
    if (index >= 0 && static_cast<size_t>(index) < tabs_.size() && IsRegFileTabIndex(index)) {
      bool was_dirty = tabs_[static_cast<size_t>(index)].reg_file_dirty;
      tabs_[static_cast<size_t>(index)].reg_file_dirty = true;
      if (!was_dirty) {
        BuildMenus();
      }
    }
    return;
  }
  if (registry_mode_ != RegistryMode::kOffline) {
    return;
  }
  int index = CurrentRegistryTabIndex();
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kOffline) {
    return;
  }
  if (!entry.offline_dirty) {
    entry.offline_dirty = true;
    BuildMenus();
  }
}

void MainWindow::ClearOfflineDirty() {
  if (registry_mode_ != RegistryMode::kOffline) {
    return;
  }
  int index = CurrentRegistryTabIndex();
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kOffline) {
    return;
  }
  if (entry.offline_dirty) {
    entry.offline_dirty = false;
    BuildMenus();
  }
}

bool MainWindow::ConfirmCloseTab(int tab_index) {
  if (!tab_ || tab_index < 0 || static_cast<size_t>(tab_index) >= tabs_.size()) {
    return false;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
  if (entry.kind == TabEntry::Kind::kRegFile && entry.reg_file_dirty) {
    std::wstring message = L"The registry file has unsaved changes.\nSave "
                           L"before closing the tab?";
    int result = ui::PromptChoice(hwnd_, message, L"Unsaved changes", L"Save", L"Don't Save", L"Cancel");
    if (result == IDCANCEL) {
      return false;
    }
    if (result == IDNO) {
      return true;
    }
    if (SaveRegFileTab(tab_index)) {
      entry.reg_file_dirty = false;
      return true;
    }
    return false;
  }
  if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kOffline || !entry.offline_dirty) {
    return true;
  }
  if (tab_index != CurrentRegistryTabIndex()) {
    return true;
  }
  std::wstring message = L"The offline registry has unsaved changes.\nSave "
                         L"before closing the tab?";
  int result = ui::PromptChoice(hwnd_, message, L"Unsaved changes", L"Save", L"Don't Save", L"Cancel");
  if (result == IDCANCEL) {
    return false;
  }
  if (result == IDNO) {
    return true;
  }
  if (SaveOfflineRegistry()) {
    entry.offline_dirty = false;
    return true;
  }
  return false;
}

void MainWindow::CloseTab(int tab_index) {
  if (!tab_) {
    return;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (count <= 1 || tab_index < 0 || tab_index >= count) {
    return;
  }
  if (IsSearchTabIndex(tab_index)) {
    CloseSearchTab(tab_index);
    return;
  }
  if (!ConfirmCloseTab(tab_index)) {
    return;
  }

  if (IsRegFileTabIndex(tab_index)) {
    TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
    if (entry.reg_file_loading && !entry.reg_file_path.empty()) {
      std::wstring lower = ToLower(entry.reg_file_path);
      auto it = reg_file_parse_sessions_.find(lower);
      if (it != reg_file_parse_sessions_.end() && it->second) {
        it->second->work.CancelAndJoin();
        reg_file_parse_sessions_.erase(it);
      }
    }
    ReleaseRegFileRoots(&entry);
  }
  tabs_.erase(tabs_.begin() + tab_index);
  TabCtrl_DeleteItem(tab_, tab_index);

  if (active_search_tab_index_ == tab_index) {
    active_search_tab_index_ = -1;
  } else if (active_search_tab_index_ > tab_index) {
    --active_search_tab_index_;
  }

  int new_count = TabCtrl_GetItemCount(tab_);
  if (new_count > 0) {
    int new_index = std::min(tab_index, new_count - 1);
    TabCtrl_SetCurSel(tab_, new_index);
    ApplyTabSelection(new_index);
  }
  RefreshRegistryTabLabels();
  ApplyViewVisibility();
  UpdateSearchResultsView();
  UpdateStatus();
}

void MainWindow::OpenLocalRegistryTab() {
  if (!tab_) {
    return;
  }
  int current = TabCtrl_GetCurSel(tab_);
  if (current >= 0 && !IsSearchTabIndex(current) && !IsRegFileTabIndex(current)) {
    CaptureRegistryTabState(current);
  }
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(L"Local Registry");
  int index = TabCtrl_GetItemCount(tab_);
  TabCtrl_InsertItem(tab_, index, &item);
  TabEntry entry;
  entry.kind = TabEntry::Kind::kRegistry;
  entry.registry_mode = RegistryMode::kLocal;
  tabs_.push_back(std::move(entry));
  RefreshRegistryTabLabels();
  TabCtrl_SetCurSel(tab_, index);
  SwitchToLocalRegistry();
  RestoreRegistryTabState(index);
  ApplyViewVisibility();
  UpdateSearchResultsView();
  UpdateStatus();
}

int MainWindow::CurrentRegistryTabIndex() const {
  if (!tab_) {
    return -1;
  }
  int index = TabCtrl_GetCurSel(tab_);
  if (index < 0) {
    return -1;
  }
  if (!IsSearchTabIndex(index) && !IsRegFileTabIndex(index)) {
    return index;
  }
  return FindFirstRegistryTabIndex();
}

void MainWindow::UpdateRegistryTabEntry(RegistryMode mode, const std::wstring& offline_path, const std::wstring& remote_machine) {
  int index = CurrentRegistryTabIndex();
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry) {
    return;
  }
  entry.registry_mode = mode;
  entry.offline_path = offline_path;
  entry.remote_machine = remote_machine;
}

void MainWindow::UpdateTabWidth() {
  if (!tab_) {
    return;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (count <= 0) {
    return;
  }
  bool has_close = count > 1;
  int pad_x = kTabTextPaddingX + (has_close ? (kTabCloseSize + kTabCloseGap) : 0);
  int pad_y = kTabInsetY + 2;
  TabCtrl_SetPadding(tab_, pad_x, pad_y);
  int text_height = 0;
  HDC hdc = GetDC(tab_);
  HFONT font = reinterpret_cast<HFONT>(SendMessageW(tab_, WM_GETFONT, 0, 0));
  HFONT old_font = nullptr;
  if (hdc && font) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, font));
  }
  if (hdc) {
    TEXTMETRICW tm = {};
    if (GetTextMetricsW(hdc, &tm)) {
      text_height = tm.tmHeight;
    }
  }

  if (hdc) {
    if (old_font) {
      SelectObject(hdc, old_font);
    }
    ReleaseDC(tab_, hdc);
  }

  int min_height = std::max<int>(24, text_height + pad_y * 2 + 2);
  SendMessageW(tab_, TCM_SETMINTABWIDTH, 0, static_cast<LPARAM>(kTabMinWidth));
  RECT item_rect = {};
  if (TabCtrl_GetItemRect(tab_, 0, &item_rect)) {
    int item_height = static_cast<int>(item_rect.bottom - item_rect.top);
    tab_height_ = std::max<int>(min_height, item_height);
  } else {
    tab_height_ = min_height;
  }
  InvalidateRect(tab_, nullptr, FALSE);
  if (hwnd_) {
    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    if (rect.right > 0 && rect.bottom > 0) {
      LayoutControls(rect.right, rect.bottom);
    }
  }
}

void MainWindow::BuildAccelerators() {
  if (accelerators_) {
    DestroyAcceleratorTable(accelerators_);
    accelerators_ = nullptr;
  }
  ACCEL accels[] = {
      {FVIRTKEY | FCONTROL, 'C', cmd::kEditCopy}, {FVIRTKEY | FCONTROL, 'V', cmd::kEditPaste}, {FVIRTKEY | FCONTROL, 'A', cmd::kViewSelectAll}, {FVIRTKEY | FCONTROL, 'Z', cmd::kEditUndo}, {FVIRTKEY | FCONTROL, 'Y', cmd::kEditRedo}, {FVIRTKEY | FCONTROL, 'F', cmd::kEditFind}, {FVIRTKEY | FCONTROL, 'G', cmd::kEditGoTo}, {FVIRTKEY | FCONTROL, 'H', cmd::kEditReplace}, {FVIRTKEY | FCONTROL, 'S', cmd::kFileSave}, {FVIRTKEY | FCONTROL, 'E', cmd::kFileExport}, {FVIRTKEY | FCONTROL | FSHIFT, 'C', cmd::kEditCopyKey}, {FVIRTKEY, VK_DELETE, cmd::kEditDelete}, {FVIRTKEY, VK_F2, cmd::kEditRename}, {FVIRTKEY, VK_F5, cmd::kViewRefresh}, {FVIRTKEY | FALT, VK_LEFT, cmd::kNavBack}, {FVIRTKEY | FALT, VK_RIGHT, cmd::kNavForward}, {FVIRTKEY | FALT, VK_UP, cmd::kNavUp},
  };
  accelerators_ = CreateAcceleratorTableW(accels, static_cast<int>(sizeof(accels) / sizeof(accels[0])));
}

bool MainWindow::SelectAllInFocusedList() {
  HWND focus = GetFocus();
  if (!focus) {
    return false;
  }
  if (focus != value_list_.hwnd() && focus != history_list_ && focus != search_results_list_) {
    return false;
  }
  int count = ListView_GetItemCount(focus);
  if (count <= 0) {
    return true;
  }
  ListView_SetItemState(focus, -1, LVIS_SELECTED, LVIS_SELECTED);
  ListView_SetItemState(focus, 0, LVIS_FOCUSED, LVIS_FOCUSED);
  ListView_EnsureVisible(focus, 0, FALSE);
  return true;
}

bool MainWindow::InvertSelectionInFocusedList() {
  HWND focus = GetFocus();
  if (!focus) {
    return false;
  }
  if (focus != value_list_.hwnd() && focus != history_list_ && focus != search_results_list_) {
    return false;
  }
  int count = ListView_GetItemCount(focus);
  if (count <= 0) {
    return true;
  }
  SendMessageW(focus, WM_SETREDRAW, FALSE, 0);
  int first_selected = -1;
  for (int i = 0; i < count; ++i) {
    UINT state = ListView_GetItemState(focus, i, LVIS_SELECTED);
    if (state & LVIS_SELECTED) {
      ListView_SetItemState(focus, i, 0, LVIS_SELECTED);
    } else {
      ListView_SetItemState(focus, i, LVIS_SELECTED, LVIS_SELECTED);
      if (first_selected < 0) {
        first_selected = i;
      }
    }
  }
  if (first_selected < 0) {
    first_selected = 0;
  }
  ListView_SetItemState(focus, first_selected, LVIS_FOCUSED, LVIS_FOCUSED);
  ListView_EnsureVisible(focus, first_selected, FALSE);
  SendMessageW(focus, WM_SETREDRAW, TRUE, 0);
  InvalidateRect(focus, nullptr, TRUE);
  return true;
}

void MainWindow::UpdateTabHotState(HWND hwnd, POINT pt) {
  int new_hot = -1;
  int new_close_hot = -1;

  TCHITTESTINFO hit = {};
  hit.pt = pt;
  int index = TabCtrl_HitTest(hwnd, &hit);
  if (index >= 0) {
    new_hot = index;
    RECT close_rect = {};
    if (GetTabCloseRect(index, &close_rect) && PtInRect(&close_rect, pt)) {
      new_close_hot = index;
    }
  }

  if (new_hot != tab_hot_index_ || new_close_hot != tab_close_hot_index_) {
    tab_hot_index_ = new_hot;
    tab_close_hot_index_ = new_close_hot;
    InvalidateRect(hwnd, nullptr, FALSE);
  }
}

bool MainWindow::GetTabCloseRect(int index, RECT* rect) const {
  if (!tab_ || !rect || index < 0) {
    return false;
  }
  int count = TabCtrl_GetItemCount(tab_);
  if (count <= 1) {
    return false;
  }
  RECT item_rect = {};
  if (!TabCtrl_GetItemRect(tab_, index, &item_rect)) {
    return false;
  }
  int header_bottom = item_rect.bottom + 1;
  RECT draw_rect = AdjustTabDrawRect(item_rect, header_bottom, false);
  RECT close_area = draw_rect;
  close_area.left = item_rect.left;
  close_area.right = item_rect.right;
  return CalcTabCloseRect(close_area, rect);
}

void MainWindow::DrawTabItem(HDC hdc, int index, const RECT& item_rect, int header_bottom, bool selected) {
  const Theme& theme = Theme::Current();
  RECT draw_rect = AdjustTabDrawRect(item_rect, header_bottom, selected);

  bool is_hot = (index == tab_hot_index_);
  bool close_hot = (index == tab_close_hot_index_);
  bool close_down = (index == tab_close_down_index_);

  COLORREF fill = selected ? theme.SurfaceColor() : theme.PanelColor();
  if (is_hot) {
    fill = theme.HoverColor();
  }
  HBRUSH fill_brush = GetCachedBrush(fill);
  FillRect(hdc, &draw_rect, fill_brush);

  HPEN border_pen = GetCachedPen(theme.BorderColor(), 1);
  HGDIOBJ old_pen = SelectObject(hdc, border_pen);
  MoveToEx(hdc, draw_rect.left, draw_rect.bottom, nullptr);
  LineTo(hdc, draw_rect.left, draw_rect.top);
  LineTo(hdc, draw_rect.right, draw_rect.top);
  LineTo(hdc, draw_rect.right, draw_rect.bottom);
  if (!selected) {
    LineTo(hdc, draw_rect.left, draw_rect.bottom);
  }
  SelectObject(hdc, old_pen);

  RECT close_rect = {};
  RECT close_area = draw_rect;
  close_area.left = item_rect.left;
  close_area.right = item_rect.right;
  bool has_close = TabCtrl_GetItemCount(tab_) > 1 && CalcTabCloseRect(close_area, &close_rect);

  RECT text_rect = draw_rect;
  text_rect.left = item_rect.left + kTabTextPaddingX;
  text_rect.right = item_rect.right - kTabTextPaddingX;
  if (has_close) {
    text_rect.right = std::max(text_rect.left, close_rect.left - kTabCloseGap);
  }

  COLORREF text_color = selected || is_hot ? theme.TextColor() : theme.MutedTextColor();
  SetTextColor(hdc, text_color);
  SetBkMode(hdc, TRANSPARENT);

  wchar_t text[256] = {};
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = text;
  item.cchTextMax = static_cast<int>(_countof(text));
  if (TabCtrl_GetItem(tab_, index, &item)) {
    DrawTextW(hdc, text, -1, &text_rect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
  }

  if (has_close) {
    if (close_down) {
      HBRUSH down_brush = GetCachedBrush(theme.SelectionColor());
      FillRect(hdc, &close_rect, down_brush);
    } else if (close_hot) {
      HBRUSH hot_brush = GetCachedBrush(theme.HoverColor());
      FillRect(hdc, &close_rect, hot_brush);
    }

    COLORREF close_color = close_down ? theme.SelectionTextColor() : theme.TextColor();
    if (icon_font_) {
      HFONT old_font = reinterpret_cast<HFONT>(SelectObject(hdc, icon_font_));
      SetTextColor(hdc, close_color);
      SetBkMode(hdc, TRANSPARENT);
      DrawTextW(hdc, L"\xE711", -1, &close_rect, DT_SINGLELINE | DT_VCENTER | DT_CENTER);
      SelectObject(hdc, old_font);
    } else {
      HPEN close_pen = GetCachedPen(close_color, 2);
      HGDIOBJ old_close_pen = SelectObject(hdc, close_pen);
      int pad = std::max<int>(2, static_cast<int>((close_rect.right - close_rect.left) / 4));
      MoveToEx(hdc, close_rect.left + pad, close_rect.top + pad, nullptr);
      LineTo(hdc, close_rect.right - pad, close_rect.bottom - pad);
      MoveToEx(hdc, close_rect.right - pad, close_rect.top + pad, nullptr);
      LineTo(hdc, close_rect.left + pad, close_rect.bottom - pad);
      SelectObject(hdc, old_close_pen);
    }
  }
}

void MainWindow::PaintTabControl(HWND hwnd, HDC hdc) {
  RECT client = {};
  GetClientRect(hwnd, &client);
  const Theme& theme = Theme::Current();
  FillRect(hdc, &client, theme.BackgroundBrush());

  HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
  HGDIOBJ old_font = nullptr;
  if (font) {
    old_font = SelectObject(hdc, font);
  }

  int count = TabCtrl_GetItemCount(hwnd);
  int current = TabCtrl_GetCurSel(hwnd);

  int header_bottom = client.top;
  RECT first_rect = {};
  if (count > 0 && TabCtrl_GetItemRect(hwnd, 0, &first_rect)) {
    int row_height = first_rect.bottom - first_rect.top;
    int rows = std::max(1, TabCtrl_GetRowCount(hwnd));
    header_bottom = first_rect.top + row_height * rows + 1;
  }

  if (header_bottom > client.top) {
    HPEN line_pen = GetCachedPen(theme.BorderColor(), 1);
    HGDIOBJ old_pen = SelectObject(hdc, line_pen);
    MoveToEx(hdc, client.left, header_bottom, nullptr);
    LineTo(hdc, client.right, header_bottom);
    SelectObject(hdc, old_pen);
  }

  for (int i = 0; i < count; ++i) {
    if (i == current) {
      continue;
    }
    RECT item_rect = {};
    if (TabCtrl_GetItemRect(hwnd, i, &item_rect)) {
      DrawTabItem(hdc, i, item_rect, header_bottom, false);
    }
  }
  if (current >= 0) {
    RECT item_rect = {};
    if (TabCtrl_GetItemRect(hwnd, current, &item_rect)) {
      DrawTabItem(hdc, current, item_rect, header_bottom, true);
    }
  }

  if (old_font) {
    SelectObject(hdc, old_font);
  }
}

void MainWindow::ReleaseRemoteRegistry() {
  if (remote_hklm_) {
    RegCloseKey(remote_hklm_);
    remote_hklm_ = nullptr;
  }
  if (remote_hku_) {
    RegCloseKey(remote_hku_);
    remote_hku_ = nullptr;
  }
  remote_machine_.clear();
}

bool MainWindow::UnloadOfflineRegistry(std::wstring* error) {
  if (error) {
    error->clear();
  }
  if (offline_roots_.empty()) {
    return true;
  }
  ClearOfflineDirty();
  for (HKEY root : offline_roots_) {
    if (!RegistryProvider::CloseOfflineHive(root, error)) {
      return false;
    }
  }
  RegistryProvider::SetOfflineRoots({});
  offline_roots_.clear();
  offline_root_labels_.clear();
  offline_root_paths_.clear();
  offline_root_ = nullptr;
  offline_mount_.clear();
  offline_root_name_.clear();
  return true;
}

void MainWindow::ApplyRegistryRoots(const std::vector<RegistryRootEntry>& roots) {
  roots_ = roots;
  ResetHiveListCache();
  current_node_ = nullptr;
  value_list_.Clear();
  current_key_count_ = 0;
  current_value_count_ = 0;
  tree_.SetRegeditLayout(UseRegeditVisibleTreeLayout());
  tree_.SetRootLabel(TreeRootLabel());
  tree_.PopulateRoots(roots_);
  ResetNavigationState();
  UpdateStatus();

  SelectDefaultTreeItem();
}

bool MainWindow::UseRegeditVisibleTreeLayout() const {
  return regedit_compatibility_mode_ && registry_mode_ == RegistryMode::kLocal;
}

std::vector<std::wstring> MainWindow::BuildVisibleTreePathParts(const std::wstring& path) const {
  std::vector<std::wstring> parts = registry_path::Split(path);
  if (parts.empty()) {
    return parts;
  }

  std::wstring root_label = TreeRootLabel();
  if (!root_label.empty() && !parts.empty() && EqualsInsensitive(parts.front(), root_label)) {
    parts.erase(parts.begin());
  }
  if (!parts.empty() && EqualsInsensitive(parts.front(), L"Computer")) {
    parts.erase(parts.begin());
  }

  auto is_standard_root = [](const std::wstring& name) -> bool {
    if (StartsWithInsensitive(name, L"HKEY_")) {
      return true;
    }
    return EqualsInsensitive(name, L"HKLM") || EqualsInsensitive(name, L"HKCU") || EqualsInsensitive(name, L"HKCR") || EqualsInsensitive(name, L"HKU") || EqualsInsensitive(name, L"HKCC");
  };
  if (!parts.empty() && EqualsInsensitive(parts.front(), L"Registry")) {
    if (parts.size() > 1 && is_standard_root(parts[1])) {
      if (UseRegeditVisibleTreeLayout()) {
        parts.erase(parts.begin());
      } else {
        parts.front() = kStandardGroupLabel;
      }
    } else if (!UseRegeditVisibleTreeLayout()) {
      parts.front() = kRealGroupLabel;
    }
  } else if (!parts.empty() && EqualsInsensitive(parts.front(), L"Real Registry")) {
    if (UseRegeditVisibleTreeLayout()) {
      parts.clear();
      return parts;
    }
    parts.front() = kRealGroupLabel;
    if (parts.size() > 1 && EqualsInsensitive(parts[1], kRealGroupLabel)) {
      parts.erase(parts.begin() + 1);
    }
  }

  if (registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty()) {
    std::wstring machine = StripMachinePrefix(remote_machine_);
    if (!machine.empty() && !parts.empty() && EqualsInsensitive(parts.front(), machine)) {
      parts.erase(parts.begin());
    }
  }
  if (registry_mode_ == RegistryMode::kOffline && !offline_root_labels_.empty() && parts.size() >= 2) {
    std::wstring root_name = offline_root_name_;
    auto is_offline_label = [&](const std::wstring& name) {
      for (const auto& label : offline_root_labels_) {
        if (EqualsInsensitive(label, name)) {
          return true;
        }
      }
      return false;
    };
    if (!root_name.empty() && EqualsInsensitive(parts[0], root_name) && is_offline_label(parts[1])) {
      parts.erase(parts.begin());
    }
  }

  if (UseRegeditVisibleTreeLayout()) {
    if (!parts.empty() && EqualsInsensitive(parts.front(), kStandardGroupLabel)) {
      parts.erase(parts.begin());
    }
    if (!parts.empty() && EqualsInsensitive(parts.front(), kRealGroupLabel)) {
      parts.clear();
      return parts;
    }
    return parts;
  }

  if (!parts.empty()) {
    if (!EqualsInsensitive(parts.front(), kStandardGroupLabel) && !EqualsInsensitive(parts.front(), kRealGroupLabel)) {
      if (EqualsInsensitive(parts.front(), L"REGISTRY")) {
        parts.insert(parts.begin(), kRealGroupLabel);
      } else {
        parts.insert(parts.begin(), kStandardGroupLabel);
      }
    }
  }
  return parts;
}

void MainWindow::RefreshVisibleRegistryTreeLayout(bool preserve_selection) {
  if (!tree_.hwnd() || roots_.empty()) {
    return;
  }

  std::wstring selected_path;
  std::vector<std::wstring> expanded_paths;
  if (preserve_selection) {
    CaptureTreeState(&selected_path, &expanded_paths);
  }

  tree_.SetRegeditLayout(UseRegeditVisibleTreeLayout());
  tree_.SetRootLabel(TreeRootLabel());
  tree_.PopulateRoots(roots_);

  if (!preserve_selection) {
    SelectDefaultTreeItem();
    return;
  }

  std::sort(expanded_paths.begin(), expanded_paths.end(), [](const std::wstring& left, const std::wstring& right) {
    if (left.size() != right.size()) {
      return left.size() < right.size();
    }
    return _wcsicmp(left.c_str(), right.c_str()) < 0;
  });
  for (const auto& path : expanded_paths) {
    ExpandTreePath(path);
  }
  if (!selected_path.empty() && SelectTreePath(selected_path)) {
    return;
  }
  SelectDefaultTreeItem();
}

std::wstring MainWindow::TreeRootLabel() const {
  if (UseRegeditVisibleTreeLayout()) {
    return L"Computer";
  }
  if (registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty()) {
    return StripMachinePrefix(remote_machine_);
  }
  wchar_t buffer[MAX_COMPUTERNAME_LENGTH + 1] = {};
  DWORD size = static_cast<DWORD>(_countof(buffer));
  if (GetComputerNameW(buffer, &size) && size > 0) {
    return std::wstring(buffer, size);
  }
  return L"Computer";
}

void MainWindow::SelectDefaultTreeItem() {
  if (!tree_.hwnd()) {
    return;
  }
  HTREEITEM root = TreeView_GetRoot(tree_.hwnd());
  if (!root) {
    return;
  }
  if (UseRegeditVisibleTreeLayout()) {
    HTREEITEM child = TreeView_GetChild(tree_.hwnd(), root);
    while (child) {
      if (tree_.NodeFromItem(child)) {
        TreeView_SelectItem(tree_.hwnd(), child);
        return;
      }
      child = TreeView_GetNextSibling(tree_.hwnd(), child);
    }
    return;
  }
  HTREEITEM group = TreeView_GetChild(tree_.hwnd(), root);
  HTREEITEM standard_group = nullptr;
  while (group) {
    wchar_t text[128] = {};
    TVITEMW tvi = {};
    tvi.mask = TVIF_TEXT;
    tvi.hItem = group;
    tvi.pszText = text;
    tvi.cchTextMax = static_cast<int>(_countof(text));
    if (TreeView_GetItem(tree_.hwnd(), &tvi)) {
      if (_wcsicmp(text, kStandardGroupLabel) == 0) {
        standard_group = group;
        break;
      }
    }
    group = TreeView_GetNextSibling(tree_.hwnd(), group);
  }
  if (standard_group) {
    TreeView_SelectItem(tree_.hwnd(), standard_group);
    return;
  }
  group = TreeView_GetChild(tree_.hwnd(), root);
  while (group) {
    RegistryNode* node = tree_.NodeFromItem(group);
    if (node) {
      TreeView_SelectItem(tree_.hwnd(), group);
      return;
    }
    HTREEITEM child = TreeView_GetChild(tree_.hwnd(), group);
    if (child) {
      TreeView_SelectItem(tree_.hwnd(), child);
      return;
    }
    group = TreeView_GetNextSibling(tree_.hwnd(), group);
  }
}

void MainWindow::CaptureRegistryTabState(int index) {
  if (!tree_.hwnd() || index < 0 ||
      static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry) {
    return;
  }
  CaptureTreeState(&entry.selected_path, &entry.expanded_paths);
}

void MainWindow::ResetRegistryTreeState() {
  if (!tree_.hwnd()) {
    return;
  }
  HTREEITEM root = TreeView_GetRoot(tree_.hwnd());
  if (!root) {
    return;
  }

  SendMessageW(tree_.hwnd(), WM_SETREDRAW, FALSE, 0);
  std::function<void(HTREEITEM)> collapse = [&](HTREEITEM item) {
    while (item) {
      HTREEITEM child = TreeView_GetChild(tree_.hwnd(), item);
      if (child) {
        collapse(child);
      }
      TreeView_Expand(tree_.hwnd(), item, TVE_COLLAPSE);
      item = TreeView_GetNextSibling(tree_.hwnd(), item);
    }
  };
  HTREEITEM child = TreeView_GetChild(tree_.hwnd(), root);
  if (child) {
    collapse(child);
  }
  TreeView_SelectItem(tree_.hwnd(), root);
  SendMessageW(tree_.hwnd(), WM_SETREDRAW, TRUE, 0);
  RedrawWindow(tree_.hwnd(), nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN);
}

void MainWindow::RestoreRegistryTabState(int index) {
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size() || !tree_.hwnd()) {
    return;
  }
  const TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind != TabEntry::Kind::kRegistry) {
    return;
  }
  ResetRegistryTreeState();
  if (entry.expanded_paths.empty() && entry.selected_path.empty()) {
    SelectDefaultTreeItem();
    return;
  }
  std::vector<std::wstring> expanded;
  expanded.reserve(entry.expanded_paths.size());
  std::unordered_set<std::wstring> seen;
  seen.reserve(entry.expanded_paths.size());
  for (const auto& path : entry.expanded_paths) {
    if (path.empty()) {
      continue;
    }
    std::wstring key = ToLower(path);
    if (seen.insert(key).second) {
      expanded.push_back(path);
    }
  }
  std::sort(expanded.begin(), expanded.end(), [](const std::wstring& left, const std::wstring& right) {
    if (left.size() != right.size()) {
      return left.size() < right.size();
    }
    return _wcsicmp(left.c_str(), right.c_str()) < 0;
  });
  for (const auto& path : expanded) {
    ExpandTreePath(path);
  }
  if (!entry.selected_path.empty() && SelectTreePath(entry.selected_path)) {
    return;
  }
  SelectDefaultTreeItem();
}

std::wstring MainWindow::LocalRegistryTabLabel(int index) const {
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return L"Local Registry";
  }
  int local_count = 0;
  int local_index = 0;
  for (size_t i = 0; i < tabs_.size(); ++i) {
    const TabEntry& entry = tabs_[i];
    if (entry.kind != TabEntry::Kind::kRegistry || entry.registry_mode != RegistryMode::kLocal) {
      continue;
    }
    ++local_count;
    if (static_cast<int>(i) == index) {
      local_index = local_count;
    }
  }
  if (local_count <= 1 || local_index <= 1) {
    return L"Local Registry";
  }
  return L"Local Registry (" + std::to_wstring(local_index) + L")";
}

void MainWindow::RefreshRegistryTabLabels() {
  if (!tab_) {
    return;
  }
  for (size_t i = 0; i < tabs_.size(); ++i) {
    const TabEntry& entry = tabs_[i];
    if (entry.kind != TabEntry::Kind::kRegistry) {
      continue;
    }
    std::wstring label;
    if (entry.registry_mode == RegistryMode::kLocal) {
      label = LocalRegistryTabLabel(static_cast<int>(i));
    } else {
      continue;
    }
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(label.c_str());
    TabCtrl_SetItem(tab_, static_cast<int>(i), &item);
  }
  UpdateTabWidth();
  InvalidateRect(tab_, nullptr, FALSE);
}

void MainWindow::AppendRealRegistryRoot(std::vector<RegistryRootEntry>* roots) {
  if (!roots || registry_mode_ != RegistryMode::kLocal) {
    return;
  }
  if (!registry_root_.get()) {
    registry_root_ = util::OpenNativeRegistryRoot();
  }
  if (!registry_root_.get()) {
    return;
  }
  RegistryRootEntry entry;
  entry.root = registry_root_.get();
  entry.display_name = L"REGISTRY";
  entry.path_name = L"REGISTRY";
  entry.subkey_prefix = L"";
  entry.group = RegistryRootGroup::kReal;
  roots->push_back(std::move(entry));
}

bool MainWindow::SwitchToLocalRegistry() {
  bool needs_reload = registry_mode_ != RegistryMode::kLocal;
  if (!needs_reload) {
    if (roots_.empty()) {
      needs_reload = true;
    } else if (RegistryProvider::IsVirtualRoot(roots_.front().root)) {
      needs_reload = true;
    } else {
      auto has_root = [&](HKEY root) -> bool {
        for (const auto& entry_root : roots_) {
          if (entry_root.root == root) {
            return true;
          }
        }
        return false;
      };
      if (!has_root(HKEY_CLASSES_ROOT) || !has_root(HKEY_CURRENT_USER) || !has_root(HKEY_LOCAL_MACHINE) || !has_root(HKEY_USERS) || !has_root(HKEY_CURRENT_CONFIG)) {
        needs_reload = true;
      }
    }
  }
  if (!needs_reload) {
    return true;
  }
  if (registry_mode_ == RegistryMode::kOffline) {
    std::wstring error;
    if (!UnloadOfflineRegistry(&error)) {
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      return false;
    }
  }
  ReleaseRemoteRegistry();
  registry_mode_ = RegistryMode::kLocal;
  UpdateRegistryTabEntry(RegistryMode::kLocal, L"", L"");
  std::vector<RegistryRootEntry> roots = RegistryProvider::DefaultRoots(show_extra_hives_);
  AppendRealRegistryRoot(&roots);
  ApplyRegistryRoots(roots);
  RefreshRegistryTabLabels();
  return true;
}

bool MainWindow::SwitchToRemoteRegistry() {
  std::wstring machine = remote_machine_;
  if (!PromptForValueText(hwnd_, L"", L"Connect to Remote Registry", L"Computer name (e.g. \\\\MACHINE):", &machine)) {
    return false;
  }
  machine = NormalizeMachineName(machine);
  if (machine.empty()) {
    ui::ShowError(hwnd_, L"Computer name is required.");
    return false;
  }

  HKEY hklm = nullptr;
  LONG result = RegConnectRegistryW(machine.c_str(), HKEY_LOCAL_MACHINE, &hklm);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }

  HKEY hku = nullptr;
  LONG hku_result = RegConnectRegistryW(machine.c_str(), HKEY_USERS, &hku);

  if (registry_mode_ == RegistryMode::kOffline) {
    std::wstring error;
    if (!UnloadOfflineRegistry(&error)) {
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      if (hku) {
        RegCloseKey(hku);
      }
      RegCloseKey(hklm);
      return false;
    }
  }

  ReleaseRemoteRegistry();
  registry_mode_ = RegistryMode::kRemote;
  remote_machine_ = machine;
  remote_hklm_ = hklm;
  remote_hku_ = hku;
  UpdateRegistryTabEntry(RegistryMode::kRemote, L"", remote_machine_);

  std::wstring prefix = machine + L"\\";
  std::vector<RegistryRootEntry> roots;
  roots.push_back({remote_hklm_, L"HKEY_LOCAL_MACHINE", prefix + L"HKEY_LOCAL_MACHINE", L""});
  if (remote_hku_) {
    roots.push_back({remote_hku_, L"HKEY_USERS", prefix + L"HKEY_USERS", L""});
  }

  UpdateTabText(L"Remote Registry (" + StripMachinePrefix(machine) + L")");
  ApplyRegistryRoots(roots);
  RefreshRegistryTabLabels();

  if (hku_result != ERROR_SUCCESS) {
    std::wstring message = L"Connected to HKEY_LOCAL_MACHINE, but HKEY_USERS was unavailable.\n";
    message += FormatWin32Error(hku_result);
    ui::ShowError(hwnd_, message);
  }
  return true;
}

bool MainWindow::SwitchToOfflineRegistry() {
  std::wstring hive_path;
  if (!PromptOpenFolderOrFile(hwnd_, L"Select Offline Hive Folder or File", &hive_path)) {
    return false;
  }
  return LoadOfflineRegistryFromPath(hive_path, true);
}

bool MainWindow::LoadOfflineRegistryFromPath(const std::wstring& path, bool open_new_tab) {
  if (registry_mode_ == RegistryMode::kOffline && !offline_roots_.empty()) {
    std::wstring error;
    if (!UnloadOfflineRegistry(&error)) {
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      return false;
    }
  }

  std::wstring selection_path = TrimTrailingSeparators(path);
  if (selection_path.empty()) {
    return false;
  }

  bool is_dir = IsDirectoryPath(selection_path);
  std::vector<OfflineHiveCandidate> candidates;
  if (is_dir) {
    CollectOfflineHivesInFolder(selection_path, &candidates);
    if (candidates.empty()) {
      ui::ShowError(hwnd_, L"The selected folder does not contain a registry hive file.");
      return false;
    }
  } else {
    std::wstring mount_name = TrimWhitespace(FileBaseName(selection_path));
    if (mount_name.empty()) {
      mount_name = L"OfflineHive";
    }
    candidates.push_back({selection_path, mount_name});
  }

  offline_root_name_ = ResolveOfflineRootName(selection_path, is_dir, current_node_);
  if (offline_root_name_.empty()) {
    offline_root_name_ = L"HKEY_LOCAL_MACHINE";
  }

  std::wstring error;
  std::vector<HKEY> handles;
  std::vector<std::wstring> labels;
  std::vector<std::wstring> paths;
  std::vector<RegistryRootEntry> roots;
  handles.reserve(candidates.size());
  labels.reserve(candidates.size());
  paths.reserve(candidates.size());
  roots.reserve(candidates.size());
  auto close_handles = [&](std::vector<HKEY>* to_close) {
    if (!to_close) {
      return;
    }
    for (HKEY root : *to_close) {
      RegistryProvider::CloseOfflineHive(root, nullptr);
    }
  };
  for (const auto& candidate : candidates) {
    HKEY hive_handle = nullptr;
    if (!RegistryProvider::OpenOfflineHive(candidate.path, &hive_handle, &error)) {
      close_handles(&handles);
      if (!error.empty()) {
        ui::ShowError(hwnd_, error);
      }
      return false;
    }
    std::wstring label = TrimWhitespace(candidate.label);
    if (label.empty()) {
      label = TrimWhitespace(FileBaseName(candidate.path));
      if (label.empty()) {
        label = L"OfflineHive";
      }
    }
    std::wstring path_name = offline_root_name_ + L"\\" + label;
    roots.push_back({hive_handle, label, path_name, L""});
    handles.push_back(hive_handle);
    labels.push_back(label);
    paths.push_back(candidate.path);
  }

  if (tab_ && open_new_tab) {
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(L"Offline Registry");
    int index = TabCtrl_GetItemCount(tab_);
    TabCtrl_InsertItem(tab_, index, &item);
    TabEntry entry;
    entry.kind = TabEntry::Kind::kRegistry;
    entry.registry_mode = RegistryMode::kOffline;
    entry.offline_path = selection_path;
    tabs_.push_back(std::move(entry));
    UpdateTabWidth();
    suppress_tab_change_ = true;
    TabCtrl_SetCurSel(tab_, index);
    suppress_tab_change_ = false;
  }

  ReleaseRemoteRegistry();
  registry_mode_ = RegistryMode::kOffline;
  offline_roots_ = std::move(handles);
  offline_root_labels_ = std::move(labels);
  offline_root_paths_ = std::move(paths);
  if (offline_roots_.size() == 1) {
    offline_root_ = offline_roots_.front();
    offline_mount_ = offline_root_labels_.front();
  } else {
    offline_root_ = nullptr;
    offline_mount_.clear();
  }
  RegistryProvider::SetOfflineRoots(offline_roots_);

  std::wstring tab_text = L"Offline Registry";
  if (offline_roots_.size() == 1 && !offline_root_name_.empty() && !offline_mount_.empty()) {
    tab_text = L"Offline Registry (" + offline_root_name_ + L"\\" + offline_mount_ + L")";
  } else if (!offline_root_name_.empty()) {
    tab_text = L"Offline Registry (" + offline_root_name_ + L")";
  }
  UpdateTabText(tab_text);
  UpdateRegistryTabEntry(RegistryMode::kOffline, selection_path, L"");
  ApplyRegistryRoots(roots);
  RefreshRegistryTabLabels();
  HistoryEntry history;
  history.action = L"Load offline registry";
  history.new_data = selection_path;
  AppendHistoryEntry(std::move(history));
  return true;
}

bool MainWindow::SaveOfflineRegistry() {
  if (registry_mode_ != RegistryMode::kOffline || offline_roots_.empty()) {
    ui::ShowError(hwnd_, L"No offline registry is loaded.");
    return false;
  }
  if (offline_roots_.size() > 1) {
    if (offline_root_paths_.size() != offline_roots_.size()) {
      ui::ShowError(hwnd_, L"Failed to resolve offline hive paths for saving.");
      return false;
    }
    for (size_t i = 0; i < offline_roots_.size(); ++i) {
      const std::wstring& path = offline_root_paths_[i];
      if (path.empty()) {
        ui::ShowError(hwnd_, L"Failed to resolve offline hive path for saving.");
        return false;
      }
      DWORD attrs = GetFileAttributesW(path.c_str());
      if (attrs != INVALID_FILE_ATTRIBUTES) {
        if (!DeleteFileW(path.c_str())) {
          ui::ShowError(hwnd_, FormatWin32Error(GetLastError()));
          return false;
        }
      }
      std::wstring error;
      if (!RegistryProvider::SaveOfflineHive(offline_roots_[i], path, &error)) {
        ui::ShowError(hwnd_, error.empty() ? L"Failed to save offline hive." : error);
        return false;
      }
    }
    ClearOfflineDirty();
    HistoryEntry history;
    history.action = L"Save offline registry";
    history.new_data = std::to_wstring(offline_roots_.size()) + L" hives";
    AppendHistoryEntry(std::move(history));
    return true;
  }
  if (!offline_root_) {
    ui::ShowError(hwnd_, L"No offline registry is loaded.");
    return false;
  }

  std::wstring path;
  if (!PromptSaveFile(hwnd_, L"Hive Files (*.*)\0*.*\0", &path)) {
    return false;
  }

  DWORD attrs = GetFileAttributesW(path.c_str());
  if (attrs != INVALID_FILE_ATTRIBUTES) {
    if (!DeleteFileW(path.c_str())) {
      ui::ShowError(hwnd_, FormatWin32Error(GetLastError()));
      return false;
    }
  }

  std::wstring error;
  if (!RegistryProvider::SaveOfflineHive(offline_root_, path, &error)) {
    ui::ShowError(hwnd_, error.empty() ? L"Failed to save offline hive." : error);
    return false;
  }
  ClearOfflineDirty();
  HistoryEntry history;
  history.action = L"Save offline registry";
  history.new_data = path;
  AppendHistoryEntry(std::move(history));
  return true;
}

void MainWindow::NavigateToAddress() {
  wchar_t buffer[512] = {};
  GetWindowTextW(address_edit_, buffer, static_cast<int>(_countof(buffer)));
  std::wstring path = NormalizeRegistryPath(buffer);
  if (path.empty()) {
    return;
  }
  std::wstring jump_key_path;
  std::wstring jump_value_name;
  if (ResolveExternalJumpTarget(path, &jump_key_path, &jump_value_name) && !jump_value_name.empty()) {
    if (NavigateToResolvedExternalJump(jump_key_path, jump_value_name, true)) {
      AddAddressHistory(jump_key_path);
    }
    return;
  }
  if (SelectTreePath(path)) {
    AddAddressHistory(path);
  } else {
    std::wstring nearest;
    if (!FindNearestExistingPath(path, &nearest) || nearest.empty()) {
      ui::ShowWarning(hwnd_, L"Registry path not found.");
      return;
    }
    std::wstring message = L"The registry key \"" + path + L"\" does not exist.";
    if (read_only_) {
      message += L"\nRead-only mode is enabled.";
      int result = ui::PromptChoice(hwnd_, message, L"Registry path not found", L"Go nearest key", L"Cancel", L"Cancel");
      if (result == IDYES) {
        if (SelectTreePath(nearest)) {
          AddAddressHistory(nearest);
        }
      }
      return;
    }
    int result = ui::PromptChoice(hwnd_, message, L"Registry path not found", L"Go nearest key", L"Create key", L"Cancel");
    if (result == IDYES) {
      if (SelectTreePath(nearest)) {
        AddAddressHistory(nearest);
      }
      return;
    }
    if (result == IDNO) {
      if (!CreateRegistryPath(path)) {
        ui::ShowError(hwnd_, L"Failed to create registry key.");
        return;
      }
      if (SelectTreePath(path)) {
        AddAddressHistory(path);
      }
    }
  }
}

void MainWindow::ApplyQueuedExternalJump() {
  if (queued_external_jump_target_.empty()) {
    return;
  }
  std::wstring target = std::move(queued_external_jump_target_);
  queued_external_jump_target_.clear();
  NavigateToExternalJump(target);
}

bool MainWindow::ResolveExternalJumpTarget(const std::wstring& target, std::wstring* key_path, std::wstring* value_name) const {
  if (!key_path || !value_name) {
    return false;
  }
  key_path->clear();
  value_name->clear();

  std::wstring normalized = NormalizeRegistryPath(target);
  if (normalized.empty()) {
    return false;
  }

  RegistryNode node;
  auto has_value = [&](const RegistryNode& candidate, const std::wstring& candidate_value) -> bool {
    if (candidate_value.empty()) {
      return false;
    }
    bool found = false;
    RegistryProvider::EnumKeyStreaming(
        candidate, true, false, false, nullptr,
        [&](const ValueInfo& value, const BYTE*, DWORD) {
          found = found || EqualsInsensitive(value.name, candidate_value);
          return true;
        },
        {});
    return found;
  };
  if (ResolvePathToNode(normalized, &node)) {
    KeyInfo info = {};
    if (RegistryProvider::QueryKeyInfo(node, &info)) {
      *key_path = std::move(normalized);
      return true;
    }
  }

  size_t search_pos = normalized.size();
  while (search_pos > 0) {
    size_t slash = normalized.rfind(L'\\', search_pos - 1);
    if (slash == std::wstring::npos || slash + 1 >= normalized.size()) {
      break;
    }
    std::wstring candidate_key = normalized.substr(0, slash);
    std::wstring candidate_value = normalized.substr(slash + 1);
    if (!candidate_value.empty() && ResolvePathToNode(candidate_key, &node)) {
      KeyInfo info = {};
      if (RegistryProvider::QueryKeyInfo(node, &info)) {
        if (!has_value(node, candidate_value)) {
          if (slash == 0) {
            break;
          }
          search_pos = slash;
          continue;
        }
        *key_path = std::move(candidate_key);
        *value_name = std::move(candidate_value);
        return true;
      }
    }
    if (slash == 0) {
      break;
    }
    search_pos = slash;
  }

  std::wstring nearest;
  if (FindNearestExistingPath(normalized, &nearest) && !nearest.empty()) {
    *key_path = std::move(nearest);
    return true;
  }
  return false;
}

bool MainWindow::NavigateToExternalJump(const std::wstring& target) {
  std::wstring key_path;
  std::wstring value_name;
  if (!ResolveExternalJumpTarget(target, &key_path, &value_name)) {
    return false;
  }
  return NavigateToResolvedExternalJump(key_path, value_name, true);
}

bool MainWindow::NavigateToResolvedExternalJump(const std::wstring& key_path, const std::wstring& value_name, bool sync_compat_controls) {
  if (key_path.empty()) {
    return false;
  }

  int sel = TabCtrl_GetCurSel(tab_);
  if (sel < 0 || static_cast<size_t>(sel) >= tabs_.size() || tabs_[static_cast<size_t>(sel)].kind != TabEntry::Kind::kRegistry) {
    int registry_tab = FindFirstRegistryTabIndex();
    if (registry_tab >= 0) {
      suppress_tab_change_ = true;
      TabCtrl_SetCurSel(tab_, registry_tab);
      suppress_tab_change_ = false;
      ApplyTabSelection(registry_tab);
    } else {
      OpenLocalRegistryTab();
    }
  }

  if (registry_mode_ != RegistryMode::kLocal) {
    if (!SwitchToLocalRegistry()) {
      return false;
    }
  }

  ApplyViewVisibility();
  UpdateStatus();

  BeginJumpUiBatch();
  if (!SelectTreePath(key_path)) {
    EndJumpUiBatch();
    return false;
  }
  ApplyTreeSelectionEffects(current_node_);
  EndJumpUiBatch();

  AddAddressHistory(key_path);
  if (sync_compat_controls && !syncing_regedit_compat_controls_) {
    EnsureRegeditCompatControls();
    SyncRegeditCompatControls(key_path, value_name);
  }

  pending_external_value_key_path_.clear();
  pending_external_value_name_.clear();
  if (!value_name.empty()) {
    pending_external_value_key_path_ = key_path;
    pending_external_value_name_ = value_name;
    if (!SelectValueByName(value_name) && current_node_ && !value_list_loading_) {
      UpdateValueListForNode(current_node_);
    }
  }
  return true;
}

void MainWindow::AppendHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data) {
  HistoryEntry entry;
  entry.action = action;
  entry.old_data = old_data;
  entry.new_data = new_data;
  if (current_node_) {
    entry.key_path = registry_path::Build(*current_node_);
  }
  AppendHistoryEntry(std::move(entry));
}

void MainWindow::AppendValueHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data, const RegistryNode& node, const std::wstring& value_name, HistoryEntry::RevertKind revert_kind, const ValueEntry* revert_value) {
  HistoryEntry entry;
  entry.action = action;
  entry.old_data = old_data;
  entry.new_data = new_data;
  entry.key_path = registry_path::Build(node);
  entry.value_name = value_name;
  entry.revert_kind = revert_kind;
  if (revert_value) {
    entry.revert_value = *revert_value;
  }
  AppendHistoryEntry(std::move(entry));
}

void MainWindow::AppendHistoryEntry(HistoryEntry entry) {
  if (!history_list_) {
    return;
  }

  const HistoryEntry appended =
      change_history_.Append(std::move(entry),
                             static_cast<size_t>(history_max_rows_));
  if (history_loaded_) {
    AppendHistoryCache(appended);
  }
  change_history_.Sort(history_sort_column_, history_sort_ascending_);
  RebuildHistoryList();
}

bool MainWindow::PrepareHistoryRevert(const HistoryEntry& entry, HistoryEntry* prepared) const {
  return changes::PrepareRevert(
      entry,
      [this](const std::wstring& path, const std::wstring& name,
             ValueEntry* value) {
        RegistryNode node;
        return ResolvePathToNode(path, &node) &&
               RegistryProvider::QueryValue(node, name, value);
      },
      prepared);
}

bool MainWindow::OpenHistoryTarget(const HistoryEntry& entry) {
  if (entry.key_path.empty()) {
    return false;
  }
  RegistryNode node;
  std::wstring target = entry.key_path;
  KeyInfo info = {};
  if (!ResolvePathToNode(target, &node) || !RegistryProvider::QueryKeyInfo(node, &info)) {
    if (!FindNearestExistingPath(target, &target) || target.empty()) {
      return false;
    }
    return NavigateToResolvedExternalJump(target, L"", true);
  }
  return NavigateToResolvedExternalJump(target, entry.value_name, true);
}

bool MainWindow::RevertHistoryEntry(const HistoryEntry& entry) {
  HistoryEntry prepared;
  if (!EnsureWritable() || !PrepareHistoryRevert(entry, &prepared)) {
    return false;
  }

  bool ok = false;
  is_replaying_ = true;
  switch (prepared.revert_kind) {
  case HistoryEntry::RevertKind::kSetValue: {
    RegistryNode node;
    if (ResolvePathToNode(prepared.key_path, &node)) {
      ok = RegistryProvider::SetValue(node, prepared.revert_value.name, prepared.revert_value.type, prepared.revert_value.data);
    }
    break;
  }
  case HistoryEntry::RevertKind::kDeleteValue: {
    RegistryNode node;
    if (ResolvePathToNode(prepared.key_path, &node)) {
      ok = RegistryProvider::DeleteValue(node, prepared.value_name);
    }
    break;
  }
  case HistoryEntry::RevertKind::kDeleteKey: {
    RegistryNode node;
    if (ResolvePathToNode(prepared.key_path, &node)) {
      std::wstring name = LeafName(node);
      if (!name.empty() && ui::ConfirmDelete(hwnd_, L"Revert Key Creation", name)) {
        ok = RegistryProvider::DeleteKey(node);
      }
    }
    break;
  }
  default:
    break;
  }
  is_replaying_ = false;

  if (!ok) {
    ui::ShowError(hwnd_, L"Failed to revert history entry.");
    return false;
  }

  MarkOfflineDirty();
  HistoryEntry revert_entry;
  revert_entry.action = L"Revert: " + entry.action;
  revert_entry.key_path = prepared.key_path;
  revert_entry.value_name = prepared.value_name;
  AppendHistoryEntry(std::move(revert_entry));
  RefreshTreeSelection();
  if (current_node_) {
    UpdateValueListForNode(current_node_);
  }
  return true;
}

bool MainWindow::EnsureSearchResultDataLoaded(search::Result* result) {
  if (!result || result->is_key || result->data_loaded) {
    return true;
  }
  result->data_loaded = true;

  RegistryNode node;
  if (!ResolvePathToNode(result->key_path, &node)) {
    result->data.clear();
    return false;
  }

  ValueEntry entry;
  if (!RegistryProvider::QueryValue(node, result->value_name, &entry)) {
    result->data.clear();
    return false;
  }

  result->type = entry.type;
  result->type_text = value_format::TypeName(entry.type);
  result->size_text = std::to_wstring(entry.data.size());
  constexpr size_t kMaxDisplaySize = 1024 * 1024;
  if (entry.data.size() > kMaxDisplaySize) {
    result->data.clear();
    return true;
  }
  result->data = value_format::DisplayData(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
  return true;
}

void MainWindow::AppendHistoryCache(const HistoryEntry& entry) {
  changes::AppendHistoryFile(HistoryCachePath(), entry);
}

std::wstring MainWindow::CacheFolderPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  std::wstring cache = util::JoinPath(folder, L"cache");
  if (!cache.empty()) {
    SHCreateDirectoryExW(nullptr, cache.c_str(), nullptr);
  }

  return cache;
}

std::wstring MainWindow::HistoryCachePath() const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"history.tsv");
}

std::wstring MainWindow::TabsCachePath() const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"tabs.ini");
}

std::wstring MainWindow::SearchTabCachePath(const std::wstring& file) const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  if (file.empty()) {
    return L"";
  }
  return util::JoinPath(folder, file);
}

bool MainWindow::EnsureSearchTabResultsLoaded(int search_index) {
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return false;
  }
  SearchTab& tab = search_tabs_[static_cast<size_t>(search_index)];
  if (tab.results_loaded) {
    return true;
  }
  tab.last_ui_count = 0;
  if (tab.cache_file.empty()) {
    tab.results_loaded = true;
    return true;
  }

  std::wstring result_path = SearchTabCachePath(tab.cache_file);
  std::vector<search::Result> loaded_results;
  bool ok = search::LoadResults(result_path, &loaded_results);
  if (!ok) {
    tab.results.clear();
    return false;
  }
  tab.results = std::move(loaded_results);
  if (tab.sort_column >= 0) {
    search::SortResults(&tab.results, tab.sort_column,
                        tab.sort_ascending, tab.is_compare);
  }
  tab.sort_dirty = false;
  tab.results_loaded = true;
  return true;
}

void MainWindow::ClearTabsCache() {
  std::wstring tabs_path = TabsCachePath();
  if (!tabs_path.empty()) {
    DeleteFileW(tabs_path.c_str());
  }
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return;
  }
  std::wstring pattern = util::JoinPath(folder, L"search_*.tsv");
  WIN32_FIND_DATAW data = {};
  HANDLE find = FindFirstFileW(pattern.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
      continue;
    }
    std::wstring path = util::JoinPath(folder, data.cFileName);
    DeleteFileW(path.c_str());
  } while (FindNextFileW(find, &data) != 0);
  FindClose(find);
}

void MainWindow::LoadTabs() {
  if (!tab_) {
    return;
  }
  tabs_.clear();
  search_tabs_.clear();
  active_search_tab_index_ = -1;
  TabCtrl_DeleteAllItems(tab_);

  int active_index = 0;
  bool loaded = false;
  if (save_tabs_) {
    workspace::TabState state;
    loaded = workspace::LoadTabs(TabsCachePath(), &state);
    if (loaded) {
      active_index = state.active_index;
      for (workspace::PersistedTab& saved : state.tabs) {
        std::wstring label = std::move(saved.label);
        if (saved.kind == workspace::PersistedTab::Kind::kRegistry) {
          if (label.empty()) {
            label = L"Local Registry";
          }
          TCITEMW item = {};
          item.mask = TCIF_TEXT;
          item.pszText = label.data();
          TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
          TabEntry entry;
          entry.kind = TabEntry::Kind::kRegistry;
          entry.registry_mode = RegistryMode::kLocal;
          entry.selected_path = std::move(saved.selected_path);
          entry.expanded_paths = std::move(saved.expanded_paths);
          tabs_.push_back(std::move(entry));
        } else {
          SearchTab search_tab;
          search_tab.label = label.empty() ? L"Find" : std::move(label);
          search_tab.cache_file = std::move(saved.search_cache_file);
          search_tab.results_loaded = search_tab.cache_file.empty();
          search_tab.is_compare =
              StartsWithInsensitive(search_tab.label, L"Compare:");
          search_tabs_.push_back(std::move(search_tab));
          const int search_index =
              static_cast<int>(search_tabs_.size() - 1);
          TCITEMW item = {};
          item.mask = TCIF_TEXT;
          item.pszText = search_tabs_.back().label.data();
          TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
          tabs_.push_back({TabEntry::Kind::kSearch, search_index});
        }
      }
    }
  }

  if (!loaded || tabs_.empty()) {
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = const_cast<wchar_t*>(L"Local Registry");
    TabCtrl_InsertItem(tab_, 0, &item);
    TabEntry entry;
    entry.kind = TabEntry::Kind::kRegistry;
    entry.registry_mode = RegistryMode::kLocal;
    tabs_.push_back(std::move(entry));
    active_index = 0;
  }

  int count = TabCtrl_GetItemCount(tab_);
  if (count > 0) {
    int sel = ClampValue(active_index, 0, count - 1);
    TabCtrl_SetCurSel(tab_, sel);
    if (IsSearchTabIndex(sel)) {
      active_search_tab_index_ = sel;
    }
  }
  RefreshRegistryTabLabels();
}

void MainWindow::SaveTabs() {
  if (!tab_) {
    return;
  }
  int current_registry_tab = CurrentRegistryTabIndex();
  if (current_registry_tab >= 0 && !IsSearchTabIndex(current_registry_tab) && !IsRegFileTabIndex(current_registry_tab)) {
    CaptureRegistryTabState(current_registry_tab);
  }
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return;
  }

  std::unordered_set<std::wstring> referenced_files;
  std::unordered_set<std::wstring> reserved_files;
  for (const auto& search_tab : search_tabs_) {
    if (!search_tab.cache_file.empty()) {
      reserved_files.insert(search_tab.cache_file);
    }
  }
  workspace::TabState state;
  int active_index = TabCtrl_GetCurSel(tab_);
  int saved_active_index = -1;

  int search_file_index = 0;
  int tab_count = TabCtrl_GetItemCount(tab_);
  int saved_index = 0;
  for (int i = 0; i < tab_count; ++i) {
    if (static_cast<size_t>(i) >= tabs_.size()) {
      break;
    }
    const auto& entry = tabs_[static_cast<size_t>(i)];
    if (entry.kind == TabEntry::Kind::kRegFile) {
      continue;
    }
    if (i == active_index) {
      saved_active_index = saved_index;
    }
    wchar_t text[256] = {};
    TCITEMW item = {};
    item.mask = TCIF_TEXT;
    item.pszText = text;
    item.cchTextMax = static_cast<int>(_countof(text));
    std::wstring label;
    if (TabCtrl_GetItem(tab_, i, &item)) {
      label = text;
    }
    if (entry.kind == TabEntry::Kind::kSearch) {
      int search_index = entry.search_index;
      if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
        continue;
      }
      SearchTab& search_tab = search_tabs_[static_cast<size_t>(search_index)];
      std::wstring file_name = search_tab.cache_file;
      if (file_name.empty()) {
        do {
          file_name = L"search_" + std::to_wstring(search_file_index++) + L".tsv";
        } while (referenced_files.find(file_name) != referenced_files.end() || reserved_files.find(file_name) != reserved_files.end());
        search_tab.cache_file = file_name;
        reserved_files.insert(file_name);
      }
      if (search_tab.results_loaded) {
        std::wstring result_path = SearchTabCachePath(file_name);
        search::SaveResults(result_path, search_tab.results);
      }
      referenced_files.insert(file_name);
      if (label.empty()) {
        label = search_tab.label;
      }
      workspace::PersistedTab saved;
      saved.kind = workspace::PersistedTab::Kind::kSearch;
      saved.label = std::move(label);
      saved.search_cache_file = std::move(file_name);
      state.tabs.push_back(std::move(saved));
    } else {
      const TabEntry& registry_entry = tabs_[static_cast<size_t>(i)];
      if (label.empty()) {
        if (registry_entry.registry_mode == RegistryMode::kLocal) {
          label = LocalRegistryTabLabel(i);
        } else {
          label = L"Local Registry";
        }
      }
      workspace::PersistedTab saved;
      saved.kind = workspace::PersistedTab::Kind::kRegistry;
      saved.label = std::move(label);
      saved.selected_path = registry_entry.selected_path;
      saved.expanded_paths = registry_entry.expanded_paths;
      state.tabs.push_back(std::move(saved));
    }
    ++saved_index;
  }
  if (saved_active_index < 0) {
    saved_active_index = 0;
  }
  state.active_index = saved_active_index;
  workspace::SaveTabs(TabsCachePath(), state);

  std::wstring pattern = util::JoinPath(folder, L"search_*.tsv");
  WIN32_FIND_DATAW data = {};
  HANDLE find = FindFirstFileW(pattern.c_str(), &data);
  if (find != INVALID_HANDLE_VALUE) {
    do {
      if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
        continue;
      }
      if (referenced_files.find(data.cFileName) != referenced_files.end()) {
        continue;
      }
      std::wstring stale = util::JoinPath(folder, data.cFileName);
      DeleteFileW(stale.c_str());
    } while (FindNextFileW(find, &data) != 0);
    FindClose(find);
  }
}

std::wstring MainWindow::CommentsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"comments.tsv");
}

void MainWindow::LoadComments() {
  value_comments_.Load(CommentsPath());
}

void MainWindow::SaveComments() const {
  value_comments_.Save(CommentsPath());
}

bool MainWindow::ImportCommentsFromFile(const std::wstring& path) {
  if (!value_comments_.Import(path)) {
    return false;
  }
  value_comments_.Save(CommentsPath());
  RefreshValueListComments();
  return true;
}

bool MainWindow::ExportCommentsToFile(const std::wstring& path) const {
  return value_comments_.Export(path);
}

void MainWindow::RefreshValueListComments() {
  if (!current_node_) {
    return;
  }
  std::wstring path = registry_path::Build(*current_node_);
  bool changed = false;
  for (auto& row : value_list_.rows()) {
    if (row.kind != rowkind::kValue) {
      if (!row.comment.empty()) {
        row.comment.clear();
        changed = true;
      }
      continue;
    }
    std::wstring value_key = changes::ValueComments::ValueKey(path, row.extra, row.value_type);
    std::wstring text;
    auto it = value_comments_.value_entries().find(value_key);
    if (it != value_comments_.value_entries().end()) {
      text = it->second.text;
    } else {
      std::wstring name_key = changes::ValueComments::NameKey(row.extra, row.value_type);
      auto it2 = value_comments_.name_entries().find(name_key);
      if (it2 != value_comments_.name_entries().end()) {
        text = it2->second.text;
      }
    }
    std::wstring display = FormatCommentDisplay(text);
    if (row.comment != display) {
      row.comment = display;
      changed = true;
    }
  }
  if (value_sort_column_ == kValueColComment) {
    SortValueRows(&value_list_.rows(), value_sort_column_, value_sort_ascending_);
    changed = true;
  }
  if (changed) {
    value_list_.InvalidateFilterCache();
  }
  if (value_list_.HasFilter()) {
    value_list_.RebuildFilter();
  } else if (changed && value_list_.hwnd()) {
    InvalidateRect(value_list_.hwnd(), nullptr, TRUE);
  }
}

bool MainWindow::EditValueComments(const std::vector<ListRow>& rows) {
  if (!current_node_ || rows.empty()) {
    return false;
  }
  std::wstring path = registry_path::Build(*current_node_);

  auto resolve_comment = [&](const ListRow& row, bool* from_name) {
    std::wstring value_key = changes::ValueComments::ValueKey(path, row.extra, row.value_type);
    auto value_it = value_comments_.value_entries().find(value_key);
    if (value_it != value_comments_.value_entries().end()) {
      if (from_name) {
        *from_name = false;
      }
      return value_it->second.text;
    }
    std::wstring name_key = changes::ValueComments::NameKey(row.extra, row.value_type);
    auto name_it = value_comments_.name_entries().find(name_key);
    if (name_it != value_comments_.name_entries().end()) {
      if (from_name) {
        *from_name = true;
      }
      return name_it->second.text;
    }
    if (from_name) {
      *from_name = false;
    }
    return std::wstring();
  };

  std::wstring initial;
  bool apply_all = true;
  bool have_initial = false;
  bool comments_match = true;
  for (const auto& row : rows) {
    if (row.kind != rowkind::kValue || row.simulated) {
      return false;
    }
    bool from_name = false;
    std::wstring comment = resolve_comment(row, &from_name);
    if (!have_initial) {
      initial = std::move(comment);
      have_initial = true;
    } else if (comment != initial) {
      comments_match = false;
    }
    apply_all = apply_all && from_name;
  }
  if (!comments_match) {
    initial.clear();
    apply_all = false;
  }

  std::wstring updated = initial;
  bool apply_all_out = apply_all;
  if (!PromptForComment(hwnd_, updated, apply_all_out, &updated, &apply_all_out)) {
    return false;
  }
  if (IsWhitespaceOnly(updated)) {
    updated.clear();
  }

  for (const auto& row : rows) {
    std::wstring value_key = changes::ValueComments::ValueKey(path, row.extra, row.value_type);
    std::wstring name_key = changes::ValueComments::NameKey(row.extra, row.value_type);
    if (updated.empty()) {
      value_comments_.value_entries().erase(value_key);
      value_comments_.name_entries().erase(name_key);
    } else if (apply_all_out) {
      changes::CommentEntry entry;
      entry.name = row.extra;
      entry.type = row.value_type;
      entry.text = updated;
      value_comments_.name_entries()[name_key] = std::move(entry);
      value_comments_.value_entries().erase(value_key);
    } else {
      changes::CommentEntry entry;
      entry.path = path;
      entry.name = row.extra;
      entry.type = row.value_type;
      entry.text = updated;
      value_comments_.value_entries()[value_key] = std::move(entry);
    }
  }
  SaveComments();
  RefreshValueListComments();
  return true;
}

void MainWindow::LoadSettings() {
  workspace::Settings settings;
  settings.clear_history_on_exit = clear_history_on_exit_;
  settings.clear_tabs_on_exit = clear_tabs_on_exit_;
  settings.show_toolbar = show_toolbar_;
  settings.show_address_bar = show_address_bar_;
  settings.show_filter_bar = show_filter_bar_;
  settings.show_tab_control = show_tab_control_;
  settings.show_tree = show_tree_;
  settings.show_history = show_history_;
  settings.show_status_bar = show_status_bar_;
  settings.show_keys_in_list = show_keys_in_list_;
  settings.show_simulated_keys = show_simulated_keys_;
  settings.show_extra_hives = show_extra_hives_;
  settings.save_tree_state = save_tree_state_;
  settings.save_tabs = save_tabs_;
  settings.always_run_as_admin = always_run_as_admin_;
  settings.always_run_as_system = always_run_as_system_;
  settings.always_run_as_trustedinstaller = always_run_as_trustedinstaller_;
  settings.always_on_top = always_on_top_;
  settings.replace_regedit = replace_regedit_;
  settings.single_instance = single_instance_;
  settings.read_only = read_only_;
  settings.window_x = window_x_;
  settings.window_y = window_y_;
  settings.window_width = window_width_;
  settings.window_height = window_height_;
  settings.window_maximized = window_maximized_;
  settings.tree_width = tree_width_;
  settings.history_height = history_height_;
  settings.theme_preset = active_theme_preset_;
  settings.icon_set = icon_set_;
  settings.use_custom_font = use_custom_font_;
  settings.font_face = custom_font_.lfFaceName;
  settings.font_size = FontPointSize(custom_font_);
  settings.font_weight = custom_font_.lfWeight;
  settings.font_italic = custom_font_.lfItalic != FALSE;
  settings.value_column_widths = saved_value_column_widths_;
  settings.value_column_visible = saved_value_column_visible_;

  if (!workspace::LoadSettings(SettingsPath(), &settings)) {
    return;
  }

  clear_history_on_exit_ = settings.clear_history_on_exit;
  clear_tabs_on_exit_ = settings.clear_tabs_on_exit;
  show_toolbar_ = settings.show_toolbar;
  show_address_bar_ = settings.show_address_bar;
  show_filter_bar_ = settings.show_filter_bar;
  show_tab_control_ = settings.show_tab_control;
  show_tree_ = settings.show_tree;
  show_history_ = settings.show_history;
  show_status_bar_ = settings.show_status_bar;
  show_keys_in_list_ = settings.show_keys_in_list;
  show_simulated_keys_ = settings.show_simulated_keys;
  show_extra_hives_ = settings.show_extra_hives;
  save_tree_state_ = settings.save_tree_state;
  save_tabs_ = settings.save_tabs;
  always_run_as_admin_ = settings.always_run_as_admin;
  always_run_as_system_ = settings.always_run_as_system;
  always_run_as_trustedinstaller_ = settings.always_run_as_trustedinstaller;
  always_on_top_ = settings.always_on_top;
  replace_regedit_ = settings.replace_regedit;
  single_instance_ = settings.single_instance;
  read_only_ = settings.read_only;
  window_placement_loaded_ = settings.window_placement_present;
  window_x_ = settings.window_x;
  window_y_ = settings.window_y;
  window_width_ = settings.window_width;
  window_height_ = settings.window_height;
  window_maximized_ = settings.window_maximized;
  tree_width_ = settings.tree_width;
  history_height_ = settings.history_height;
  if (_wcsicmp(settings.theme_mode.c_str(), L"dark") == 0) {
    theme_mode_ = ThemeMode::kDark;
  } else if (_wcsicmp(settings.theme_mode.c_str(), L"light") == 0) {
    theme_mode_ = ThemeMode::kLight;
  } else if (_wcsicmp(settings.theme_mode.c_str(), L"custom") == 0) {
    theme_mode_ = ThemeMode::kCustom;
  } else {
    theme_mode_ = ThemeMode::kSystem;
  }
  active_theme_preset_ = std::move(settings.theme_preset);
  icon_set_ = IsKnownIconSetName(settings.icon_set) &&
                      !IsIconSetName(settings.icon_set, kIconSetLucide)
                  ? std::move(settings.icon_set)
                  : kIconSetDefault;
  use_custom_font_ = settings.use_custom_font;
  if (!settings.font_face.empty()) {
    wcsncpy_s(custom_font_.lfFaceName, settings.font_face.c_str(), _TRUNCATE);
  }
  if (settings.font_size > 0) {
    custom_font_.lfHeight = FontHeightFromPointSize(settings.font_size);
  }
  custom_font_.lfWeight = settings.font_weight;
  custom_font_.lfItalic = settings.font_italic ? TRUE : FALSE;
  recent_trace_paths_.Replace(std::move(settings.recent_traces));
  recent_default_paths_.Replace(std::move(settings.recent_defaults));
  saved_value_column_widths_ = std::move(settings.value_column_widths);
  saved_value_column_visible_ = std::move(settings.value_column_visible);
  saved_value_columns_loaded_ = !saved_value_column_widths_.empty() ||
                                !saved_value_column_visible_.empty();
  if (!save_tree_state_) {
    saved_tree_state_.Clear();
  }
}
void MainWindow::SaveSettings() const {
  workspace::Settings settings;
  settings.clear_history_on_exit = clear_history_on_exit_;
  settings.clear_tabs_on_exit = clear_tabs_on_exit_;
  settings.show_toolbar = show_toolbar_;
  settings.show_address_bar = show_address_bar_;
  settings.show_filter_bar = show_filter_bar_;
  settings.show_tab_control = show_tab_control_;
  settings.show_tree = show_tree_;
  settings.show_history = show_history_;
  settings.show_status_bar = show_status_bar_;
  settings.show_keys_in_list = show_keys_in_list_;
  settings.show_simulated_keys = show_simulated_keys_;
  settings.show_extra_hives = show_extra_hives_;
  settings.save_tree_state = save_tree_state_;
  settings.save_tabs = save_tabs_;
  settings.always_run_as_admin = always_run_as_admin_;
  settings.always_run_as_system = always_run_as_system_;
  settings.always_run_as_trustedinstaller = always_run_as_trustedinstaller_;
  settings.always_on_top = always_on_top_;
  settings.replace_regedit = replace_regedit_;
  settings.single_instance = single_instance_;
  settings.read_only = read_only_;
  settings.window_x = window_x_;
  settings.window_y = window_y_;
  settings.window_width = window_width_;
  settings.window_height = window_height_;
  settings.window_maximized = window_maximized_;
  if (hwnd_ && IsWindow(hwnd_)) {
    WINDOWPLACEMENT placement = {};
    placement.length = sizeof(placement);
    if (GetWindowPlacement(hwnd_, &placement)) {
      const RECT& normal = placement.rcNormalPosition;
      const int width = normal.right - normal.left;
      const int height = normal.bottom - normal.top;
      if (width > 0 && height > 0) {
        settings.window_x = normal.left;
        settings.window_y = normal.top;
        settings.window_width = width;
        settings.window_height = height;
      }
      settings.window_maximized = placement.showCmd == SW_SHOWMAXIMIZED;
    }
  }
  settings.tree_width = tree_width_;
  settings.history_height = history_height_;
  switch (theme_mode_) {
  case ThemeMode::kDark:
    settings.theme_mode = L"dark";
    break;
  case ThemeMode::kLight:
    settings.theme_mode = L"light";
    break;
  case ThemeMode::kCustom:
    settings.theme_mode = L"custom";
    break;
  default:
    settings.theme_mode = L"system";
    break;
  }
  settings.theme_preset = active_theme_preset_;
  settings.icon_set = IsKnownIconSetName(icon_set_) ? icon_set_ : kIconSetDefault;
  settings.use_custom_font = use_custom_font_;
  settings.font_face = custom_font_.lfFaceName;
  settings.font_size = FontPointSize(custom_font_);
  settings.font_weight = custom_font_.lfWeight;
  settings.font_italic = custom_font_.lfItalic != FALSE;
  settings.recent_traces = recent_trace_paths_.items();
  settings.recent_defaults = recent_default_paths_.items();
  settings.value_column_widths = value_column_widths_;
  settings.value_column_visible = value_column_visible_;
  settings.value_column_widths.resize(value_columns_.size(), 0);
  settings.value_column_visible.resize(value_columns_.size(), true);
  workspace::SaveSettings(SettingsPath(), settings);
}
std::wstring MainWindow::SettingsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"settings.ini");
}

std::wstring MainWindow::TreeStatePath() const {
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"tree_state.ini");
}

void MainWindow::LoadTreeState() {
  saved_tree_state_.Clear();
  if (!save_tree_state_) {
    return;
  }
  workspace::LoadTreeState(TreeStatePath(), &saved_tree_state_);
}

void MainWindow::StartTreeStateWorker() {
  if (!save_tree_state_ || tree_state_saver_.running()) {
    return;
  }
  tree_state_saver_.Start(
      std::chrono::seconds(2),
      [this](workspace::TreeState state) {
        SaveTreeStateFile(state.selected_path, state.expanded_paths);
      });
}

void MainWindow::StopTreeStateWorker() {
  tree_state_saver_.Stop();
  if (save_tree_state_ && tree_.hwnd() && IsWindow(tree_.hwnd())) {
    std::wstring selected;
    std::vector<std::wstring> expanded;
    CaptureTreeState(&selected, &expanded);
    SaveTreeStateFile(selected, expanded);
  }
}

void MainWindow::StartValueListWorker() {
  if (value_loader_.running()) {
    return;
  }
  value_loader_.Start(
      [this](std::unique_ptr<ValueListTask> task,
             const std::atomic_bool& stopping) {
      if (!task || stopping.load() ||
          task->generation != value_list_generation_.load()) {
        return;
      }

      auto payload = std::make_unique<ValueListPayload>();
      payload->generation = task->generation;
      struct KeyMetadata {
        int image_index = kFolderIconIndex;
        bool is_link = false;
      };
      std::unordered_map<std::wstring, KeyMetadata> key_metadata_cache;
      key_metadata_cache.reserve(256);

      auto resolve_key_icon = [&](const RegistryNode& node, bool* is_link) -> int {
        if (is_link) {
          *is_link = false;
        }
        if (node.simulated) {
          return kFolderSimIconIndex;
        }
        std::wstring cache_key = registry_path::Build(node);
        auto cached = key_metadata_cache.find(cache_key);
        if (cached != key_metadata_cache.end()) {
          if (is_link) {
            *is_link = cached->second.is_link;
          }
          return cached->second.image_index;
        }
        KeyMetadata metadata;
        std::wstring nt_path = registry_path::BuildNative(node);
        if (!nt_path.empty() && task->hive_roots && task->hive_roots->find(ToLower(nt_path)) != task->hive_roots->end()) {
          metadata.image_index = kDatabaseIconIndex;
          key_metadata_cache.emplace(std::move(cache_key), metadata);
          return metadata.image_index;
        }
        std::wstring link_target;
        if (RegistryProvider::QuerySymbolicLinkTarget(node, &link_target)) {
          if (is_link) {
            *is_link = true;
          }
          metadata.image_index = kSymlinkIconIndex;
          metadata.is_link = true;
          key_metadata_cache.emplace(std::move(cache_key), metadata);
          return metadata.image_index;
        }
        key_metadata_cache.emplace(std::move(cache_key), metadata);
        return metadata.image_index;
      };

      auto subkeys = RegistryProvider::EnumSubKeyNames(task->snapshot, false);
      std::unordered_set<std::wstring> existing_keys;
      existing_keys.reserve(subkeys.size());
      for (const auto& name : subkeys) {
        existing_keys.insert(ToLower(name));
      }
      std::vector<std::wstring> simulated_subkeys;
      auto append_trace_children = [&](const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) {
        if (!out) {
          return;
        }
        out->clear();
        if (!task->show_simulated_keys) {
          return;
        }
        if (task->trace_data_list.empty()) {
          return;
        }
        if (!node.root_name.empty() && EqualsInsensitive(node.root_name, L"REGISTRY")) {
          return;
        }
        std::wstring path = registry_path::Build(node);
        std::wstring trace_path = NormalizeTraceKeyPath(path);
        if (trace_path.empty()) {
          trace_path = path;
        }
        std::wstring key_lower = ToLower(trace_path);
        std::unordered_set<std::wstring> seen;
        for (const auto& trace : task->trace_data_list) {
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
      };
      append_trace_children(task->snapshot, existing_keys, &simulated_subkeys);
      payload->key_count = static_cast<int>(subkeys.size() + simulated_subkeys.size());
      payload->rows.reserve((task->show_keys_in_list ? subkeys.size() + simulated_subkeys.size() : 0) + 16);

      struct TraceMatch {
        std::wstring label;
        trace::KeyValues values;
        const trace::Selection* selection = nullptr;
      };
      std::vector<TraceMatch> trace_matches;
      if (!task->trace_data_list.empty()) {
        for (const auto& trace : task->trace_data_list) {
          if (!trace.data) {
            continue;
          }
          std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
          if (!trace.selection ||
              !trace::IncludesKey(*trace.selection,
                                  task->trace_path_lower)) {
            continue;
          }
          auto it = trace.data->values_by_key.find(task->trace_path_lower);
          if (it == trace.data->values_by_key.end()) {
            continue;
          }
          TraceMatch match;
          match.label = trace.label.empty() ? L"Trace" : trace.label;
          match.values = it->second;
          match.selection = trace.selection.get();
          trace_matches.push_back(std::move(match));
        }
      }

      struct DefaultMatch {
        defaults::Key values;
        const trace::Selection* selection = nullptr;
      };
      std::vector<DefaultMatch> default_keys;
      if (!task->default_data_list.empty() && !task->default_path_lower.empty()) {
        default_keys.reserve(task->default_data_list.size());
        for (const auto& defaults : task->default_data_list) {
          if (!defaults.data) {
            continue;
          }
          std::shared_lock<std::shared_mutex> defaults_lock(*defaults.data->mutex);
          if (!defaults.selection ||
              !trace::IncludesKey(*defaults.selection,
                                  task->default_path_lower)) {
            continue;
          }
          auto it = defaults.data->values_by_key.find(task->default_path_lower);
          if (it == defaults.data->values_by_key.end()) {
            continue;
          }
          default_keys.push_back({it->second, defaults.selection.get()});
        }
      }
      auto resolve_default_data = [&](const std::wstring& value_name) -> std::wstring {
        if (default_keys.empty()) {
          return {};
        }
        std::wstring value_lower = ToLower(value_name);
        bool applies = false;
        for (const auto& match : default_keys) {
          if (match.selection &&
              !trace::IncludesValue(*match.selection,
                                    task->default_path_lower,
                                    value_lower)) {
            continue;
          }
          applies = true;
          auto it = match.values.values.find(value_lower);
          if (it != match.values.values.end()) {
            return it->second.data;
          }
        }
        return applies ? L"(Missing)" : std::wstring();
      };

      if (task->show_keys_in_list) {
        for (const auto& name : subkeys) {
          ListRow row;
          row.name = name;
          bool is_link = false;
          RegistryNode child = task->snapshot;
          child.subkey = task->snapshot.subkey.empty() ? name : task->snapshot.subkey + L"\\" + name;
          row.image_index = resolve_key_icon(child, &is_link);
          row.type = is_link ? L"Link" : L"Key";
          row.extra = name;
          row.kind = rowkind::kKey;
          if (task->include_dates || task->include_details) {
            KeyInfo info = {};
            if (RegistryProvider::QueryKeyInfo(child, &info)) {
              if (task->include_dates) {
                row.date = FormatFileTime(info.last_write);
                row.date_value = FileTimeToUint64(info.last_write);
                row.has_date = (row.date_value != 0);
              }
              if (task->include_details) {
                row.detail_key_count = info.subkey_count;
                row.detail_value_count = info.value_count;
                row.has_details = true;
                row.details = L"Keys: " + std::to_wstring(info.subkey_count) + L", Values: " + std::to_wstring(info.value_count);
              }
            }
          }
          payload->rows.emplace_back(std::move(row));
        }
        for (const auto& name : simulated_subkeys) {
          if (name.empty()) {
            continue;
          }
          ListRow row;
          row.name = name;
          RegistryNode child = task->snapshot;
          child.subkey = task->snapshot.subkey.empty() ? name : task->snapshot.subkey + L"\\" + name;
          child.simulated = true;
          row.image_index = resolve_key_icon(child, nullptr);
          row.simulated = true;
          row.type = L"Key";
          row.extra = name;
          row.kind = rowkind::kKey;
          payload->rows.emplace_back(std::move(row));
        }
      }

      std::wstring link_target;
      bool has_link = RegistryProvider::QuerySymbolicLinkTarget(task->snapshot, &link_target);
      bool track_existing = (!trace_matches.empty()) || has_link;
      std::unordered_set<std::wstring> existing_values;
      if (track_existing) {
        existing_values.reserve(64);
      }

      auto gather_labels = [&](const std::wstring& value_lower) -> std::vector<std::wstring> {
        std::vector<std::wstring> labels;
        for (const auto& match : trace_matches) {
          if (match.values.values_lower.find(value_lower) != match.values.values_lower.end()) {
            if (match.selection &&
                !trace::IncludesValue(*match.selection,
                                      task->trace_path_lower,
                                      value_lower)) {
              continue;
            }
            labels.push_back(match.label);
          }
        }
        if (labels.size() < 2) {
          return labels;
        }
        std::vector<std::wstring> unique;
        unique.reserve(labels.size());
        std::unordered_set<std::wstring> seen;
        for (const auto& label : labels) {
          std::wstring key = ToLower(label);
          if (seen.insert(key).second) {
            unique.push_back(label);
          }
        }
        return unique;
      };
      auto format_read_on_boot = [&](const std::vector<std::wstring>& labels) -> std::wstring {
        if (labels.empty()) {
          return L"No";
        }
        std::wstring out = L"Yes (";
        for (size_t i = 0; i < labels.size(); ++i) {
          if (i > 0) {
            out.append(L", ");
          }
          out.append(labels[i]);
        }
        out.push_back(L')');
        return out;
      };
      bool have_traces = !trace_matches.empty();

      bool has_default = false;
      bool has_symbolic_value = false;
      constexpr DWORD kEagerValueDataLimit = 64u * 1024u;
      const DWORD max_data_size = task->include_all_value_data ? MAXDWORD : kEagerValueDataLimit;
      auto append_value = [&](const ValueInfo& value, const BYTE* data, DWORD data_size) -> bool {
        if (value.name.empty()) {
          has_default = true;
        }
        if (EqualsInsensitive(value.name, L"SymbolicLinkValue")) {
          has_symbolic_value = true;
        }
        ListRow row;
        row.name = value.name.empty() ? L"(Default)" : value.name;
        row.type = value_format::TypeName(value.type);
        row.data_ready = data_size == 0 || data != nullptr;
        if (row.data_ready && data_size > 0) {
          row.data = value_format::DisplayData(value.type, data, data_size);
        }
        row.default_data = resolve_default_data(value.name);
        row.image_index = UseBinaryValueIcon(value.type) ? kBinaryIconIndex : kValueIconIndex;
        row.kind = rowkind::kValue;
        row.extra = value.name;
        row.size = std::to_wstring(data_size);
        row.size_value = data_size;
        row.has_size = true;
        row.value_type = value.type;
        row.value_data_size = data_size;
        if (!have_traces) {
          row.read_on_boot.clear();
        } else {
          std::wstring lower = ToLower(value.name);
          row.read_on_boot = format_read_on_boot(gather_labels(lower));
          if (track_existing) {
            existing_values.insert(lower);
          }
        }
        payload->rows.emplace_back(std::move(row));
        ++payload->value_count;
        return !stopping.load() &&
               task->generation == value_list_generation_.load();
      };
      RegistryProvider::EnumKeyStreaming(task->snapshot, true, true, false, nullptr, append_value, RegistryProvider::SubkeyStreamCallback(), max_data_size);

      if (!has_symbolic_value && has_link && !link_target.empty()) {
        ListRow row;
        row.name = L"SymbolicLinkValue";
        row.type = L"REG_LINK";
        row.data = link_target;
        row.data_ready = true;
        row.image_index = UseBinaryValueIcon(REG_LINK) ? kBinaryIconIndex : kValueIconIndex;
        row.kind = rowkind::kValue;
        row.extra = L"SymbolicLinkValue";
        row.default_data = resolve_default_data(row.extra);
        DWORD link_bytes = static_cast<DWORD>((link_target.size() + 1) * sizeof(wchar_t));
        row.size = std::to_wstring(link_bytes);
        row.size_value = link_bytes;
        row.value_data_size = link_bytes;
        row.has_size = true;
        row.value_type = REG_LINK;
        row.read_on_boot = have_traces ? L"No" : L"";
        row.simulated = true;
        payload->rows.emplace_back(std::move(row));
      }

      if (!has_default) {
        ListRow row;
        row.name = L"(Default)";
        row.type = L"REG_SZ";
        row.data = L"(value not set)";
        row.data_ready = true;
        row.image_index = kValueIconIndex;
        row.kind = rowkind::kValue;
        row.extra = L"";
        row.default_data = resolve_default_data(row.extra);
        row.size = L"0";
        row.size_value = 0;
        row.has_size = true;
        row.value_type = REG_SZ;
        if (!have_traces) {
          row.read_on_boot.clear();
        } else {
          row.read_on_boot = format_read_on_boot(gather_labels(L""));
          if (track_existing) {
            existing_values.insert(L"");
          }
        }
        payload->rows.emplace_back(std::move(row));
        payload->value_count += 1;
      }

      size_t trace_added = 0;
      if (!trace_matches.empty()) {
        for (const auto& match : trace_matches) {
          payload->rows.reserve(payload->rows.size() + match.values.values_display.size());
          for (const auto& value_name : match.values.values_display) {
            std::wstring value_lower = ToLower(value_name);
            if (match.selection &&
                !trace::IncludesValue(*match.selection,
                                      task->trace_path_lower,
                                      value_lower)) {
              continue;
            }
            if (existing_values.find(value_lower) != existing_values.end()) {
              continue;
            }
            ListRow row;
            row.name = value_name.empty() ? L"(Default)" : value_name;
            row.type = L"TRACE";
            row.data = L"(value not set)";
            row.read_on_boot = format_read_on_boot(gather_labels(value_lower));
            row.image_index = kValueIconIndex;
            row.kind = rowkind::kValue;
            row.extra = value_name;
            row.data_ready = true;
            row.default_data = resolve_default_data(value_name);
            payload->rows.emplace_back(std::move(row));
            ++trace_added;
            existing_values.insert(value_lower);
          }
        }
        payload->value_count += static_cast<int>(trace_added);
      }

      if (task->sort_column != kValueColComment) {
        SortValueRows(&payload->rows, task->sort_column, task->sort_ascending);
      }
      if (stopping.load() ||
          task->generation != value_list_generation_.load()) {
        return;
      }
      if (PostMessageW(task->hwnd, kValueListReadyMessage, static_cast<WPARAM>(task->generation), reinterpret_cast<LPARAM>(payload.get())) != 0) {
        ReleasePostedPayload(payload);
      }
  });
}

void MainWindow::StopValueListWorker() {
  value_loader_.Stop();
}

void MainWindow::MergeTraceEntries(
    TraceParseSession* session,
    const std::vector<KeyValueDialogEntry>& entries,
    std::unordered_set<std::wstring>* affected_keys) {
  if (!session || !session->data || entries.empty()) {
    return;
  }
  std::vector<trace::Entry> parsed;
  parsed.reserve(entries.size());
  for (const auto& entry : entries) {
    trace::Entry value;
    value.key_path = entry.key_path;
    value.display_path = entry.display_path;
    value.has_value = entry.has_value;
    value.value_name = entry.value_name;
    parsed.push_back(std::move(value));
  }
  trace::Merge(session->data.get(), parsed, affected_keys);
}

void MainWindow::MergeDefaultEntries(
    DefaultParseSession* session,
    const std::vector<KeyValueDialogEntry>& entries,
    std::unordered_set<std::wstring>* affected_keys) {
  if (!session || !session->data || entries.empty()) {
    return;
  }
  std::vector<defaults::Entry> parsed;
  parsed.reserve(entries.size());
  for (const auto& entry : entries) {
    defaults::Entry value;
    value.key_path = entry.key_path;
    value.has_value = entry.has_value;
    value.value_name = entry.value_name;
    value.type = entry.value_type;
    value.data = entry.value_data;
    parsed.push_back(std::move(value));
  }
  defaults::Merge(session->data.get(), parsed,
                  [](const std::wstring& path) {
                    return MapControlSetToCurrent(path);
                  },
                  affected_keys);
}

void MainWindow::StartTraceParseThread(TraceParseSession* session) {
  if (!session || session->work.running()) {
    return;
  }
  HWND hwnd = hwnd_;
  std::wstring source = session->source_path;
  std::wstring source_lower = session->source_lower;
  session->work.Start(
      [this, session, hwnd, source, source_lower](
          uint64_t generation, std::atomic_bool& cancel) {
    constexpr size_t kBatchSize = 256;
    constexpr DWORD kBatchMs = 50;
    auto post_batch = [&](std::vector<KeyValueDialogEntry>* entries, bool done, const std::wstring& error, bool cancelled) {
      auto payload = std::make_unique<TraceParseBatch>();
      payload->generation = generation;
      payload->source_lower = source_lower;
      if (entries) {
        MergeTraceEntries(session, *entries, &payload->affected_keys);
        payload->entries = std::move(*entries);
      }
      payload->done = done;
      payload->error = error;
      payload->cancelled = cancelled;
      if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, kTraceParseBatchMessage, 0, reinterpret_cast<LPARAM>(payload.get()))) {
        return;
      }
      ReleasePostedPayload(payload);
    };

    std::vector<KeyValueDialogEntry> entries;
    entries.reserve(kBatchSize);
    uint64_t last_post = GetTickCount64();
    std::wstring parse_error;
    const bool parsed = trace::LoadEntries(
        source, TraceNormalizers(),
        [&](trace::Entry&& parsed_entry) {
          KeyValueDialogEntry entry;
          entry.key_path = std::move(parsed_entry.key_path);
          entry.display_path = std::move(parsed_entry.display_path);
          entry.has_value = true;
          entry.value_name = std::move(parsed_entry.value_name);
          entries.push_back(std::move(entry));
          const uint64_t now = GetTickCount64();
          if (entries.size() >= kBatchSize ||
              now - last_post >= kBatchMs) {
            post_batch(&entries, false, L"", false);
            entries.clear();
            last_post = now;
          }
          return !cancel.load();
        },
        &parse_error, &cancel);
    if (!parsed) {
      post_batch(nullptr, true, cancel.load() ? L"" : parse_error,
                 cancel.load());
      return;
    }
    if (!entries.empty()) {
      post_batch(&entries, false, L"", false);
      entries.clear();
    }
    trace::Sort(session->data.get());
    post_batch(nullptr, true, L"", false);
  });
}

void MainWindow::StartDefaultParseThread(DefaultParseSession* session) {
  if (!session || session->work.running()) {
    return;
  }
  HWND hwnd = hwnd_;
  std::wstring source = session->source_path;
  std::wstring source_lower = session->source_lower;
  session->work.Start(
      [this, session, hwnd, source, source_lower](
          uint64_t generation, std::atomic_bool& cancel) {
    constexpr size_t kBatchSize = 256;
    constexpr DWORD kBatchMs = 50;
    auto post_batch = [&](std::vector<KeyValueDialogEntry>* entries, bool done, const std::wstring& error, bool cancelled) {
      auto payload = std::make_unique<DefaultParseBatch>();
      payload->generation = generation;
      payload->source_lower = source_lower;
      if (entries) {
        MergeDefaultEntries(session, *entries, &payload->affected_keys);
        payload->entries = std::move(*entries);
      }
      payload->done = done;
      payload->error = error;
      payload->cancelled = cancelled;
      if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, kDefaultParseBatchMessage, 0, reinterpret_cast<LPARAM>(payload.get()))) {
        return;
      }
      ReleasePostedPayload(payload);
    };

    std::vector<defaults::Entry> parsed_entries;
    std::wstring parse_error;
    if (!defaults::Load(
            source,
            [](const std::wstring& path) {
              return NormalizeTraceKeyPathBasic(path);
            },
            nullptr, &parsed_entries, &parse_error, &cancel)) {
      post_batch(nullptr, true, cancel.load() ? L"" : parse_error,
                 cancel.load());
      return;
    }

    std::vector<KeyValueDialogEntry> entries;
    entries.reserve(kBatchSize);
    uint64_t last_post = GetTickCount64();
    auto flush_if_needed = [&] {
      const uint64_t now = GetTickCount64();
      if (entries.size() >= kBatchSize || now - last_post >= kBatchMs) {
        post_batch(&entries, false, L"", false);
        entries.clear();
        last_post = now;
      }
    };

    for (auto& parsed : parsed_entries) {
      if (cancel.load()) {
        post_batch(nullptr, true, L"", true);
        return;
      }
      KeyValueDialogEntry entry;
      entry.key_path = std::move(parsed.key_path);
      entry.display_path =
          NormalizeTraceSelectionPath(parsed.source_path);
      if (entry.display_path.empty()) {
        entry.display_path = entry.key_path;
      }
      entry.has_value = parsed.has_value;
      entry.value_name = std::move(parsed.value_name);
      entry.value_type = parsed.type;
      entry.value_data = std::move(parsed.data);
      entries.push_back(std::move(entry));
      flush_if_needed();
    }

    if (!entries.empty()) {
      post_batch(&entries, false, L"", false);
    }
    post_batch(nullptr, true, L"", false);
  });
}

void MainWindow::StartTraceLoadWorker() {
  if (trace_load_session_.running()) {
    return;
  }
  LoadTraceSettings();
  std::unordered_map<std::wstring, trace::Selection> selection_cache =
      trace_selection_cache_;
  std::wstring active_path = ActiveTracesPath();
  const HWND hwnd = hwnd_;
  trace_load_session_.StartIfIdle(
      [this, selection_cache = std::move(selection_cache), active_path, hwnd](
          uint64_t generation, const std::atomic_bool& cancel) mutable {
    auto payload = std::make_unique<TraceLoadPayload>();
    payload->generation = generation;
    payload->selection_cache = std::move(selection_cache);
    std::wstring content;
    if (!util::ReadTextFile(
            active_path, &content, nullptr,
            static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
      return;
    }

    std::vector<std::wstring> entries;
    size_t start = 0;
    while (start < content.size()) {
      size_t end = content.find(L'\n', start);
      if (end == std::wstring::npos) {
        end = content.size();
      }
      std::wstring line = content.substr(start, end - start);
      if (!line.empty() && line.back() == L'\r') {
        line.pop_back();
      }
      start = end + 1;
      line = TrimWhitespace(line);
      if (line.empty() || line.front() == L'#') {
        continue;
      }
      if (StartsWithInsensitive(line, L"trace=")) {
        line.erase(0, wcslen(L"trace="));
      }
      line = TrimWhitespace(line);
      if (!line.empty()) {
        entries.push_back(std::move(line));
      }
    }

    std::unordered_set<std::wstring> loaded;
    for (const auto& entry : entries) {
      if (cancel.load()) {
        return;
      }
      std::wstring source = entry;
      std::wstring use_label;
      if (!FileExists(source)) {
        std::wstring bundled = ResolveBundledTracePath(source);
        if (!bundled.empty() && FileExists(bundled)) {
          source = bundled;
          use_label = entry;
        } else {
          continue;
        }
      }
      if (use_label.empty()) {
        use_label = FileBaseName(source);
      }
      if (use_label.empty()) {
        use_label = L"Trace";
      }
      std::wstring source_lower = ToLower(source);
      if (!loaded.insert(source_lower).second) {
        continue;
      }
      trace::Data data;
      if (!trace::Load(use_label, source, TraceNormalizers(), &data,
                       nullptr, &cancel)) {
        continue;
      }
      std::shared_ptr<const trace::Data> trace_data =
          std::make_shared<trace::Data>(std::move(data));
      trace::Selection selection = {};
      selection.select_all = true;
      selection.recursive = true;
      auto it = payload->selection_cache.find(source_lower);
      if (it != payload->selection_cache.end()) {
        selection = it->second;
      }
      trace::NormalizeSelection(*trace_data, &selection);
      payload->selection_cache[source_lower] = selection;
      payload->traces.push_back(
          {trace_data->label, source, trace_data,
           std::make_shared<trace::Selection>(selection)});
    }

    if (cancel.load()) {
      return;
    }
    if (hwnd && IsWindow(hwnd) &&
        PostMessageW(hwnd, kTraceLoadReadyMessage, 0,
                     reinterpret_cast<LPARAM>(payload.get()))) {
      ReleasePostedPayload(payload);
    }
  });
}

void MainWindow::StopTraceLoadWorker() {
  trace_load_session_.CancelAndJoin();
}

void MainWindow::StopTraceParseSessions() {
  for (auto& entry : trace_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->work.CancelAndJoin();
  }
  trace_parse_sessions_.clear();
}

void MainWindow::StartDefaultLoadWorker() {
  if (default_load_session_.running()) {
    return;
  }
  std::wstring active_path = ActiveDefaultsPath();
  const HWND hwnd = hwnd_;
  default_load_session_.StartIfIdle(
      [this, active_path, hwnd](uint64_t generation,
                                const std::atomic_bool& cancel) {
    auto payload = std::make_unique<DefaultLoadPayload>();
    payload->generation = generation;
    std::wstring content;
    if (!util::ReadTextFile(
            active_path, &content, nullptr,
            static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
      return;
    }

    std::vector<std::wstring> entries;
    size_t start = 0;
    while (start < content.size()) {
      size_t end = content.find(L'\n', start);
      if (end == std::wstring::npos) {
        end = content.size();
      }
      std::wstring line = content.substr(start, end - start);
      if (!line.empty() && line.back() == L'\r') {
        line.pop_back();
      }
      start = end + 1;
      line = TrimWhitespace(line);
      if (line.empty() || line.front() == L'#') {
        continue;
      }
      if (StartsWithInsensitive(line, L"default=")) {
        line.erase(0, wcslen(L"default="));
      }
      line = TrimWhitespace(line);
      if (!line.empty()) {
        entries.push_back(std::move(line));
      }
    }

    std::unordered_set<std::wstring> loaded;
    for (const auto& entry : entries) {
      if (cancel.load()) {
        return;
      }
      std::wstring source = entry;
      std::wstring use_label;
      if (!FileExists(source)) {
        std::wstring bundled = ResolveBundledDefaultPath(source);
        if (!bundled.empty() && FileExists(bundled)) {
          source = bundled;
          use_label = entry;
        } else {
          continue;
        }
      }
      if (use_label.empty()) {
        use_label = FileBaseName(source);
      }
      if (use_label.empty()) {
        use_label = L"Default";
      }
      std::wstring source_lower = ToLower(source);
      if (!loaded.insert(source_lower).second) {
        continue;
      }
      defaults::Data data;
      if (!defaults::Load(
              source,
              [](const std::wstring& path) {
                return NormalizeTraceKeyPathBasic(path);
              },
              &data, nullptr, nullptr,
              &cancel)) {
        continue;
      }
      std::shared_ptr<const defaults::Data> default_data =
          std::make_shared<defaults::Data>(std::move(data));
      trace::Selection selection = {};
      selection.select_all = true;
      selection.recursive = true;
      payload->defaults.push_back(
          {use_label, source, default_data,
           std::make_shared<trace::Selection>(selection)});
    }

    if (cancel.load()) {
      return;
    }
    if (hwnd && IsWindow(hwnd) &&
        PostMessageW(hwnd, kDefaultLoadReadyMessage, 0,
                     reinterpret_cast<LPARAM>(payload.get()))) {
      ReleasePostedPayload(payload);
    }
  });
}

void MainWindow::StopDefaultLoadWorker() {
  default_load_session_.CancelAndJoin();
}

void MainWindow::StopDefaultParseSessions() {
  for (auto& entry : default_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->work.CancelAndJoin();
  }
  default_parse_sessions_.clear();
}

void MainWindow::StopRegFileParseSessions() {
  for (auto& entry : reg_file_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->work.CancelAndJoin();
  }
  reg_file_parse_sessions_.clear();
}

void MainWindow::MarkTreeStateDirty() {
  if (!save_tree_state_ || !tree_.hwnd() || !IsWindow(tree_.hwnd())) {
    return;
  }
  std::wstring selected;
  std::vector<std::wstring> expanded;
  CaptureTreeState(&selected, &expanded);
  workspace::TreeState state;
  state.selected_path = std::move(selected);
  state.expanded_paths = std::move(expanded);
  tree_state_saver_.Submit(std::move(state));
}

void MainWindow::SaveTreeStateFile(const std::wstring& selected, const std::vector<std::wstring>& expanded) const {
  workspace::TreeState state;
  state.selected_path = selected;
  state.expanded_paths = expanded;
  workspace::SaveTreeState(TreeStatePath(), state);
}

std::wstring MainWindow::ActiveTracesPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"active_traces.ini");
}

std::wstring MainWindow::ActiveDefaultsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"active_defaults.ini");
}

std::wstring MainWindow::TraceSettingsPath() const {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"trace_settings.ini");
}

void MainWindow::LoadTraceSettings() {
  trace_selection_cache_.clear();
  std::wstring path = TraceSettingsPath();
  if (path.empty()) {
    return;
  }
  auto parse_bool = [](const std::wstring& value) -> bool { return (_wcsicmp(value.c_str(), L"1") == 0 || _wcsicmp(value.c_str(), L"true") == 0 || _wcsicmp(value.c_str(), L"yes") == 0); };
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, MutableData(buffer), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(buffer);
  if (content.empty()) {
    return;
  }

  trace::Selection selection = {};
  selection.select_all = true;
  selection.recursive = true;
  std::wstring current_path;
  std::wstring current_label;
  bool has_entry = false;

  auto normalize_selection = [&]() {
    std::vector<std::wstring> cleaned;
    cleaned.reserve(selection.key_paths.size());
    std::unordered_set<std::wstring> seen;
    for (auto& path : selection.key_paths) {
      std::wstring trimmed = TrimWhitespace(path);
      if (trimmed.empty()) {
        continue;
      }
      std::wstring lower = ToLower(trimmed);
      if (seen.insert(lower).second) {
        cleaned.push_back(std::move(trimmed));
      }
    }
    for (const auto& entry : selection.values_by_key) {
      if (entry.first.empty()) {
        continue;
      }
      if (seen.insert(entry.first).second) {
        cleaned.push_back(entry.first);
      }
    }
    selection.key_paths.swap(cleaned);
    if (selection.key_paths.empty() && selection.values_by_key.empty()) {
      selection.select_all = true;
    }
  };

  auto flush_entry = [&]() {
    if (!has_entry) {
      return;
    }
    normalize_selection();
    std::wstring key = current_path;
    if (key.empty() && !current_label.empty()) {
      std::wstring resolved = ResolveBundledTracePath(current_label);
      key = resolved.empty() ? current_label : resolved;
    }
    key = TrimWhitespace(key);
    if (!key.empty()) {
      trace_selection_cache_[ToLower(key)] = selection;
    }
    selection = {};
    selection.select_all = true;
    selection.recursive = true;
    current_path.clear();
    current_label.clear();
    has_entry = false;
  };

  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    line = TrimWhitespace(line);
    if (line.empty()) {
      flush_entry();
      continue;
    }
    if (line.front() == L'#') {
      continue;
    }
    if (line.front() == L'[') {
      flush_entry();
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = TrimWhitespace(line.substr(0, sep));
    std::wstring value = line.substr(sep + 1);
    if (key.empty()) {
      continue;
    }
    has_entry = true;
    if (EqualsInsensitive(key, L"path")) {
      current_path = value;
    } else if (EqualsInsensitive(key, L"label")) {
      current_label = value;
    } else if (EqualsInsensitive(key, L"select_all")) {
      selection.select_all = parse_bool(value);
    } else if (EqualsInsensitive(key, L"recursive")) {
      selection.recursive = parse_bool(value);
    } else if (EqualsInsensitive(key, L"key_path") || EqualsInsensitive(key, L"key")) {
      selection.key_paths.push_back(value);
    } else if (EqualsInsensitive(key, L"value")) {
      size_t bar = value.find(L'|');
      if (bar == std::wstring::npos) {
        continue;
      }
      std::wstring key_part = TrimWhitespace(value.substr(0, bar));
      std::wstring value_part = TrimWhitespace(value.substr(bar + 1));
      if (key_part.empty()) {
        continue;
      }
      if (value_part == L"@") {
        value_part.clear();
      }
      std::wstring key_lower = ToLower(key_part);
      std::wstring value_lower = ToLower(value_part);
      selection.values_by_key[key_lower].insert(value_lower);
    }
  }
  flush_entry();
}

void MainWindow::SaveTraceSettings() const {
  std::wstring path = TraceSettingsPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& trace : active_traces_) {
    if (!trace.data || !trace.selection) {
      continue;
    }
    content.append(L"[trace]\n");
    if (!trace.label.empty()) {
      content.append(L"label=");
      content.append(trace.label);
      content.push_back(L'\n');
    }
    if (!trace.source_path.empty()) {
      content.append(L"path=");
      content.append(trace.source_path);
      content.push_back(L'\n');
    }
    content.append(L"select_all=");
    content.append(trace.selection->select_all ? L"1\n" : L"0\n");
    content.append(L"recursive=");
    content.append(trace.selection->recursive ? L"1\n" : L"0\n");
    for (const auto& key_path : trace.selection->key_paths) {
      if (key_path.empty()) {
        continue;
      }
      content.append(L"key=");
      content.append(key_path);
      content.push_back(L'\n');
    }
    for (const auto& entry : trace.selection->values_by_key) {
      if (entry.first.empty()) {
        continue;
      }
      for (const auto& value_name : entry.second) {
        content.append(L"value=");
        content.append(entry.first);
        content.push_back(L'|');
        if (value_name.empty()) {
          content.append(L"@");
        } else {
          content.append(value_name);
        }
        content.push_back(L'\n');
      }
    }
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  if (!utf8.empty()) {
    DWORD written = 0;
    WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  }
  CloseHandle(file);
}

void MainWindow::SaveActiveTraces() const {
  std::wstring path = ActiveTracesPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& trace : active_traces_) {
    if (trace.source_path.empty()) {
      continue;
    }
    std::wstring entry = trace.source_path;
    if (!trace.label.empty()) {
      std::wstring bundled = ResolveBundledTracePath(trace.label);
      if (!bundled.empty() && EqualsInsensitive(bundled, trace.source_path)) {
        entry = trace.label;
      }
    }
    content.append(entry);
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  if (utf8.empty() && !content.empty()) {
    return;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  if (!utf8.empty()) {
    DWORD written = 0;
    WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  }
  CloseHandle(file);
}

void MainWindow::SaveActiveDefaults() const {
  std::wstring path = ActiveDefaultsPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& defaults : active_defaults_) {
    if (defaults.source_path.empty()) {
      continue;
    }
    std::wstring entry = defaults.source_path;
    if (!defaults.label.empty()) {
      std::wstring bundled = ResolveBundledDefaultPath(defaults.label);
      if (!bundled.empty() && EqualsInsensitive(bundled, defaults.source_path)) {
        entry = defaults.label;
      }
    }
    content.append(entry);
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  if (utf8.empty() && !content.empty()) {
    return;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  if (!utf8.empty()) {
    DWORD written = 0;
    WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  }
  CloseHandle(file);
}

bool MainWindow::HasActiveTraces() const {
  return !active_traces_.empty();
}

bool MainWindow::RemoveTraceByPath(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::wstring target = TrimWhitespace(path);
  if (target.empty()) {
    return false;
  }
  std::wstring target_lower = ToLower(target);
  auto session_it = trace_parse_sessions_.find(target_lower);
  if (session_it != trace_parse_sessions_.end()) {
    if (session_it->second) {
      session_it->second->work.CancelAndJoin();
    }
    trace_parse_sessions_.erase(session_it);
  }
  size_t removed = 0;
  active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(),
                                      [&](const ActiveTrace& trace) {
                                        if (!EqualsInsensitive(trace.source_path, target)) {
                                          return false;
                                        }
                                        ++removed;
                                        return true;
                                      }),
                       active_traces_.end());
  if (removed == 0) {
    return false;
  }
  trace_selection_cache_.erase(target_lower);
  SaveActiveTraces();
  SaveTraceSettings();
  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(current_node_);
  SaveSettings();
  return true;
}

bool MainWindow::RemoveTraceByLabel(const std::wstring& label) {
  if (label.empty()) {
    return false;
  }
  for (auto it = trace_parse_sessions_.begin(); it != trace_parse_sessions_.end();) {
    if (it->second && _wcsicmp(it->second->label.c_str(), label.c_str()) == 0) {
      it->second->work.CancelAndJoin();
      it = trace_parse_sessions_.erase(it);
      continue;
    }
    ++it;
  }
  size_t removed = 0;
  active_traces_.erase(std::remove_if(active_traces_.begin(), active_traces_.end(),
                                      [&](const ActiveTrace& trace) {
                                        if (_wcsicmp(trace.label.c_str(), label.c_str()) != 0) {
                                          return false;
                                        }
                                        ++removed;
                                        return true;
                                      }),
                       active_traces_.end());
  if (removed == 0) {
    return false;
  }
  trace_selection_cache_.clear();
  for (const auto& trace : active_traces_) {
    if (!trace.source_path.empty()) {
      if (trace.selection) {
        trace_selection_cache_[ToLower(trace.source_path)] = *trace.selection;
      }
    }
  }
  SaveActiveTraces();
  SaveTraceSettings();
  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(current_node_);
  SaveSettings();
  return true;
}

bool MainWindow::RemoveDefaultByPath(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  std::wstring target = TrimWhitespace(path);
  if (target.empty()) {
    return false;
  }
  std::wstring target_lower = ToLower(target);
  auto session_it = default_parse_sessions_.find(target_lower);
  if (session_it != default_parse_sessions_.end()) {
    if (session_it->second) {
      session_it->second->work.CancelAndJoin();
    }
    default_parse_sessions_.erase(session_it);
  }
  size_t removed = 0;
  active_defaults_.erase(std::remove_if(active_defaults_.begin(), active_defaults_.end(),
                                        [&](const ActiveDefault& defaults) {
                                          if (!EqualsInsensitive(defaults.source_path, target)) {
                                            return false;
                                          }
                                          ++removed;
                                          return true;
                                        }),
                         active_defaults_.end());
  if (removed == 0) {
    return false;
  }
  SaveActiveDefaults();
  BuildMenus();
  UpdateValueListForNode(current_node_);
  SaveSettings();
  return true;
}

void MainWindow::ShowPermissionsDialog(const RegistryNode& node) {
  ShowRegistryPermissions(hwnd_, node);
}

bool MainWindow::RestartAsAdmin() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }
  HINSTANCE result = ShellExecuteW(hwnd_, L"runas", exe_path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  if (reinterpret_cast<INT_PTR>(result) <= 32) {
    ui::ShowError(hwnd_, L"Failed to restart with administrator rights.");
    return false;
  }
  PostMessageW(hwnd_, WM_CLOSE, 0, 0);
  return true;
}

bool MainWindow::RestartAsSystem() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }

  if (!util::IsProcessElevated()) {
    HINSTANCE result = ShellExecuteW(hwnd_, L"runas", exe_path.c_str(), kRestartSystemArg, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
      ui::ShowError(hwnd_, L"Failed to request SYSTEM restart.");
      return false;
    }
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  }

  std::wstring command_line = L"\"";
  command_line += exe_path;
  command_line += L"\" ";
  command_line += kRestartSystemArg;
  DWORD error = 0;
  if (!util::LaunchProcessAsSystem(command_line, L"", &error)) {
    std::wstring message = L"Failed to restart with SYSTEM rights.";
    std::wstring detail = FormatWin32Error(error);
    if (!detail.empty()) {
      message += L"\n";
      message += detail;
    }
    ui::ShowError(hwnd_, message);
    return false;
  }

  PostMessageW(hwnd_, WM_CLOSE, 0, 0);
  return true;
}

bool MainWindow::RestartAsTrustedInstaller() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }

  if (!util::IsProcessElevated()) {
    HINSTANCE result = ShellExecuteW(hwnd_, L"runas", exe_path.c_str(), kRestartTiArg, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) <= 32) {
      ui::ShowError(hwnd_, L"Failed to request TrustedInstaller restart.");
      return false;
    }
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  }

  std::wstring command_line = L"\"";
  command_line += exe_path;
  command_line += L"\" ";
  command_line += kRestartTiArg;
  DWORD error = 0;
  if (!util::LaunchProcessAsTrustedInstaller(command_line, L"", &error)) {
    std::wstring message = L"Failed to restart with TrustedInstaller rights.";
    std::wstring detail = FormatWin32Error(error);
    if (!detail.empty()) {
      message += L"\n";
      message += detail;
    }
    ui::ShowError(hwnd_, message);
    return false;
  }

  PostMessageW(hwnd_, WM_CLOSE, 0, 0);
  return true;
}

void MainWindow::SyncReplaceRegeditState() {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    return;
  }

  HKEY base = nullptr;
  LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\regedit.exe", 0, KEY_QUERY_VALUE, &base);
  if (result != ERROR_SUCCESS) {
    replace_regedit_ = false;
    return;
  }

  DWORD type = 0;
  DWORD size = 0;
  result = RegQueryValueExW(base, L"Debugger", nullptr, &type, nullptr, &size);
  if (result != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || size == 0) {
    RegCloseKey(base);
    replace_regedit_ = false;
    return;
  }

  std::wstring debugger;
  debugger.resize(size / sizeof(wchar_t));
  result = RegQueryValueExW(base, L"Debugger", nullptr, &type, reinterpret_cast<LPBYTE>(MutableData(debugger)), &size);
  RegCloseKey(base);
  if (result != ERROR_SUCCESS) {
    replace_regedit_ = false;
    return;
  }

  while (!debugger.empty() && debugger.back() == L'\0') {
    debugger.pop_back();
  }
  if (debugger.empty()) {
    replace_regedit_ = false;
    return;
  }

  std::wstring expanded = debugger;
  if (type == REG_EXPAND_SZ) {
    std::wstring resolved = util::ExpandEnvironmentStringsDynamic(debugger);
    if (!resolved.empty()) {
      expanded = std::move(resolved);
    }
  }

  const wchar_t* start = expanded.c_str();
  while (*start && iswspace(*start)) {
    ++start;
  }
  std::wstring path;
  if (*start == L'\"') {
    ++start;
    const wchar_t* end = wcschr(start, L'\"');
    if (end) {
      path.assign(start, static_cast<size_t>(end - start));
    } else {
      path.assign(start);
    }
  } else {
    const wchar_t* end = start;
    while (*end && !iswspace(*end)) {
      ++end;
    }
    path.assign(start, static_cast<size_t>(end - start));
  }

  if (path.empty()) {
    replace_regedit_ = false;
    return;
  }

  replace_regedit_ = (_wcsicmp(path.c_str(), exe_path.c_str()) == 0);
}

void MainWindow::ReplaceRegedit(bool enable) {
  std::wstring exe_path = util::GetModulePath();
  if (exe_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return;
  }

  HKEY base = nullptr;
  DWORD base_disp = 0;
  LONG result = RegCreateKeyExW(HKEY_LOCAL_MACHINE, L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options", 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, &base, &base_disp);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return;
  }

  std::wstring subkey = L"regedit.exe";
  if (enable) {
    HKEY app_key = nullptr;
    DWORD disposition = 0;
    result = RegCreateKeyExW(base, subkey.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, nullptr, &app_key, &disposition);
    if (result != ERROR_SUCCESS) {
      RegCloseKey(base);
      ui::ShowError(hwnd_, FormatWin32Error(result));
      return;
    }
    std::wstring debugger = L"\"" + exe_path + L"\"";
    result = RegSetValueExW(app_key, L"Debugger", 0, REG_SZ, reinterpret_cast<const BYTE*>(debugger.c_str()), static_cast<DWORD>((debugger.size() + 1) * sizeof(wchar_t)));
    RegCloseKey(app_key);
    RegCloseKey(base);
    if (result != ERROR_SUCCESS) {
      ui::ShowError(hwnd_, FormatWin32Error(result));
      return;
    }
    replace_regedit_ = true;
  } else {
    HKEY app_key = nullptr;
    result = RegOpenKeyExW(base, subkey.c_str(), 0, KEY_READ | KEY_WRITE, &app_key);
    if (result == ERROR_SUCCESS) {
      RegDeleteValueW(app_key, L"Debugger");
      DWORD subkeys = 0;
      DWORD values = 0;
      if (RegQueryInfoKeyW(app_key, nullptr, nullptr, nullptr, &subkeys, nullptr, nullptr, &values, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS && subkeys == 0 && values == 0) {
        RegCloseKey(app_key);
        RegDeleteKeyW(base, subkey.c_str());
      } else {
        RegCloseKey(app_key);
      }
    }
    RegCloseKey(base);
    replace_regedit_ = false;
  }

  SaveSettings();
  BuildMenus();
}

bool MainWindow::OpenDefaultRegedit() {
  if (!util::IsProcessElevated() && !util::IsProcessSystem() &&
      !util::IsProcessTrustedInstaller()) {
    ui::ShowError(hwnd_, L"Administrator rights are required to open the default Regedit.");
    return false;
  }

  std::wstring regedit_path = GetDefaultRegeditPath();
  if (regedit_path.empty()) {
    ui::ShowError(hwnd_, L"Failed to locate the default Regedit executable.");
    return false;
  }
  auto launch_regedit = [&]() -> bool {
    DWORD launch_error = ERROR_SUCCESS;
    if (LaunchDefaultRegeditProcess(regedit_path, &launch_error)) {
      return true;
    }
    ui::ShowError(hwnd_, FormatWin32Error(launch_error));
    return false;
  };

  util::UniqueHKey key;
  LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, kRegeditIfeoPath, 0, KEY_READ | KEY_WRITE | KEY_WOW64_64KEY, key.put());
  if (result == ERROR_FILE_NOT_FOUND) {
    return launch_regedit();
  }
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }

  DWORD type = 0;
  DWORD size = 0;
  result = RegQueryValueExW(key.get(), L"Debugger", nullptr, &type, nullptr, &size);
  if (result == ERROR_FILE_NOT_FOUND) {
    return launch_regedit();
  }
  if (result != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || size == 0) {
    ui::ShowError(hwnd_, L"Failed to read the Regedit debugger value.");
    return false;
  }

  std::vector<BYTE> data(size);
  result = RegQueryValueExW(key.get(), L"Debugger", nullptr, &type, data.data(), &size);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }
  data.resize(size);

  std::wstring temp_name = L"Debugger_RegKitTemp";
  DWORD temp_type = 0;
  DWORD temp_size = 0;
  int suffix = 0;
  while (RegQueryValueExW(key.get(), temp_name.c_str(), nullptr, &temp_type, nullptr, &temp_size) == ERROR_SUCCESS) {
    ++suffix;
    temp_name = L"Debugger_RegKitTemp_" + std::to_wstring(suffix);
    if (suffix > 100) {
      ui::ShowError(hwnd_, L"Failed to prepare a temporary Regedit debugger value.");
      return false;
    }
  }

  result = RegSetValueExW(key.get(), temp_name.c_str(), 0, type, data.data(), size);
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }
  result = RegDeleteValueW(key.get(), L"Debugger");
  if (result != ERROR_SUCCESS) {
    RegDeleteValueW(key.get(), temp_name.c_str());
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }

  DWORD launch_error = ERROR_SUCCESS;
  bool launched = LaunchDefaultRegeditProcess(regedit_path, &launch_error);

  LONG restore = RegSetValueExW(key.get(), L"Debugger", 0, type, data.data(), size);
  RegDeleteValueW(key.get(), temp_name.c_str());
  if (restore != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(restore));
    return false;
  }
  if (!launched) {
    ui::ShowError(hwnd_, FormatWin32Error(launch_error));
    return false;
  }
  return true;
}

void MainWindow::OpenHiveFileDir() {
  if (registry_mode_ == RegistryMode::kRemote) {
    ui::ShowError(hwnd_, L"Hive files are not available for remote registries.");
    return;
  }
  RegistryNode* node = current_node_;
  if (!node && tree_.hwnd()) {
    HTREEITEM selected = TreeView_GetSelection(tree_.hwnd());
    if (selected) {
      node = tree_.NodeFromItem(selected);
    }
  }
  if (!node) {
    return;
  }
  RegistryNode target = *node;
  int index = ListView_GetNextItem(value_list_.hwnd(), -1, LVNI_SELECTED);
  if (index >= 0) {
    const ListRow* row = value_list_.RowAt(index);
    if (row && row->kind == rowkind::kKey && !row->extra.empty()) {
      target = MakeChildNode(*node, row->extra);
    }
  }
  bool is_root = false;
  std::wstring hive_path = LookupHivePath(target, &is_root);
  if (hive_path.empty()) {
    ui::ShowError(hwnd_, L"No hive file was found for this key.");
    return;
  }
  std::wstring args = L"/select,\"" + hive_path + L"\"";
  std::wstring folder = hive_path;
  folder.push_back(L'\0');
  if (SUCCEEDED(PathCchRemoveFileSpec(folder.data(), folder.size()))) {
    folder.resize(wcsnlen_s(folder.data(), folder.size()));
    ShellExecuteW(hwnd_, L"open", L"explorer.exe", args.c_str(), folder.c_str(), SW_SHOWNORMAL);
  } else {
    ShellExecuteW(hwnd_, L"open", L"explorer.exe", args.c_str(), nullptr, SW_SHOWNORMAL);
  }
}

LOGFONTW MainWindow::DefaultLogFont() const {
  LOGFONTW lf = {};
  std::wstring face = ReadFontSubstitute(L"Segoe UI");
  if (face.empty()) {
    face = L"Segoe UI";
  }
  lf.lfHeight = FontHeightFromPointSize(9);
  lf.lfWeight = FW_NORMAL;
  lf.lfCharSet = DEFAULT_CHARSET;
  wcsncpy_s(lf.lfFaceName, face.c_str(), _TRUNCATE);
  return lf;
}

void MainWindow::RefreshTreeSelection() {
  if (!tree_.hwnd()) {
    return;
  }
  HTREEITEM item = TreeView_GetSelection(tree_.hwnd());
  if (!item) {
    return;
  }
  RegistryNode* node = tree_.NodeFromItem(item);
  if (!node) {
    return;
  }
  HTREEITEM child = TreeView_GetChild(tree_.hwnd(), item);
  while (child) {
    HTREEITEM next = TreeView_GetNextSibling(tree_.hwnd(), child);
    TreeView_DeleteItem(tree_.hwnd(), child);
    child = next;
  }
  node->children_loaded = false;
  NMTREEVIEWW info = {};
  info.action = TVE_EXPAND;
  info.itemNew.hItem = item;
  tree_.OnItemExpanding(&info);
  TreeView_Expand(tree_.hwnd(), item, TVE_EXPAND);
  MarkTreeStateDirty();
}

void MainWindow::UpdateSimulatedChain(HTREEITEM item) {
  if (!tree_.hwnd() || !item) {
    return;
  }
  while (item) {
    RegistryNode* node = tree_.NodeFromItem(item);
    if (node && node->simulated) {
      KeyInfo info = {};
      if (RegistryProvider::QueryKeyInfo(*node, &info)) {
        node->simulated = false;
        int icon = KeyIconIndex(*node, nullptr, nullptr);
        TVITEMW tvi = {};
        tvi.mask = TVIF_IMAGE | TVIF_SELECTEDIMAGE;
        tvi.hItem = item;
        tvi.iImage = icon;
        tvi.iSelectedImage = icon;
        TreeView_SetItem(tree_.hwnd(), &tvi);
      }
    }
    item = TreeView_GetParent(tree_.hwnd(), item);
  }
}

void MainWindow::CaptureTreeState(std::wstring* selected_path, std::vector<std::wstring>* expanded_paths) const {
  if (selected_path) {
    selected_path->clear();
  }
  if (expanded_paths) {
    expanded_paths->clear();
  }
  if (!tree_.hwnd()) {
    return;
  }
  if (selected_path) {
    RegistryNode* node = current_node_;
    if (!node) {
      HTREEITEM selected = TreeView_GetSelection(tree_.hwnd());
      if (selected) {
        TVITEMW tvi = {};
        tvi.hItem = selected;
        tvi.mask = TVIF_PARAM;
        if (TreeView_GetItem(tree_.hwnd(), &tvi)) {
          node = reinterpret_cast<RegistryNode*>(tvi.lParam);
        }
      }
    }
    if (node) {
      *selected_path = registry_path::Build(*node);
    }
  }
  if (!expanded_paths) {
    return;
  }
  HTREEITEM root = TreeView_GetRoot(tree_.hwnd());
  if (!root) {
    return;
  }
  std::function<void(HTREEITEM, bool)> walk = [&](HTREEITEM item, bool ancestors_expanded) {
    while (item) {
      TVITEMW tvi = {};
      tvi.hItem = item;
      tvi.mask = TVIF_STATE | TVIF_PARAM;
      tvi.stateMask = TVIS_EXPANDED;
      if (TreeView_GetItem(tree_.hwnd(), &tvi)) {
        bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
        if (ancestors_expanded && expanded) {
          RegistryNode* node = reinterpret_cast<RegistryNode*>(tvi.lParam);
          if (node) {
            expanded_paths->push_back(registry_path::Build(*node));
          }
        }
      }
      HTREEITEM child = TreeView_GetChild(tree_.hwnd(), item);
      if (child) {
        bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
        if (ancestors_expanded && expanded) {
          walk(child, true);
        }
      }
      item = TreeView_GetNextSibling(tree_.hwnd(), item);
    }
  };
  walk(root, true);
}

void MainWindow::RestoreTreeState() {
  if (tree_state_restored_) {
    return;
  }
  if (!save_tree_state_) {
    return;
  }
  tree_state_restored_ = true;
  if (!tree_.hwnd()) {
    return;
  }
  workspace::TreeState state = saved_tree_state_;
  state.Normalize();
  for (const auto& path : state.expanded_paths) {
    ExpandTreePath(path);
  }
  if (!state.selected_path.empty()) {
    SelectTreePath(state.selected_path);
  }
}

void MainWindow::ApplySavedWindowPlacement() {
  if (!window_placement_loaded_ || !hwnd_) {
    return;
  }
  const int min_width = 640;
  const int min_height = 480;
  int width = std::max(window_width_, min_width);
  int height = std::max(window_height_, min_height);
  if (window_width_ <= 0 || window_height_ <= 0) {
    return;
  }
  SetWindowPos(hwnd_, nullptr, window_x_, window_y_, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
}

bool MainWindow::ExpandTreePath(const std::wstring& path) {
  if (!tree_.hwnd()) {
    return false;
  }
  std::vector<std::wstring> parts = BuildVisibleTreePathParts(path);
  if (parts.empty()) {
    return false;
  }
  HTREEITEM root = TreeView_GetRoot(tree_.hwnd());
  HTREEITEM current = root;
  for (const auto& part : parts) {
    TreeView_Expand(tree_.hwnd(), current, TVE_EXPAND);
    HTREEITEM child = FindChildByText(tree_.hwnd(), current, part);
    if (!child) {
      return false;
    }
    current = child;
  }
  if (current) {
    TreeView_Expand(tree_.hwnd(), current, TVE_EXPAND);
    return true;
  }
  return false;
}

void MainWindow::PushUndo(changes::UndoOperation operation) {
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

void MainWindow::ClearRedo() {
  undo_stack_.ClearRedo();
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                 undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, 0);
  }
}

bool MainWindow::ApplyUndoOperation(const changes::UndoOperation& operation,
                                    bool redo) {
  if (!current_node_) {
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
        ok = RegistryProvider::CreateKey(operation.node, operation.name);
      }
      if (ok) {
        RefreshTreeSelection();
      }
    } else {
      RegistryNode child = MakeChildNode(operation.node, operation.name);
      ok = RegistryProvider::DeleteKey(child);
      if (ok) {
        RefreshTreeSelection();
      }
    }
    break;
  }
  case changes::UndoOperation::Type::kDeleteKey: {
    if (redo) {
      RegistryNode child = MakeChildNode(operation.node, operation.name);
      ok = RegistryProvider::DeleteKey(child);
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
    ok = RegistryProvider::RenameKey(child, to);
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
      ok = RegistryProvider::SetValue(operation.node, operation.new_value.name, operation.new_value.type, operation.new_value.data);
    } else {
      ok = RegistryProvider::DeleteValue(operation.node, operation.name);
    }
    break;
  }
  case changes::UndoOperation::Type::kDeleteValue: {
    if (redo) {
      ok = RegistryProvider::DeleteValue(operation.node, operation.old_value.name);
    } else {
      ok = RegistryProvider::SetValue(operation.node, operation.old_value.name, operation.old_value.type, operation.old_value.data);
    }
    break;
  }
  case changes::UndoOperation::Type::kModifyValue: {
    const ValueEntry& value = redo ? operation.new_value : operation.old_value;
    ok = RegistryProvider::SetValue(operation.node, value.name, value.type, value.data);
    break;
  }
  case changes::UndoOperation::Type::kRenameValue: {
    std::wstring from = redo ? operation.name : operation.new_name;
    std::wstring to = redo ? operation.new_name : operation.name;
    ok = RegistryProvider::RenameValue(operation.node, from, to);
    break;
  }
  default:
    break;
  }
  is_replaying_ = false;

  if (ok) {
    MarkOfflineDirty();
  }
  if (ok && current_node_) {
    UpdateValueListForNode(current_node_);
  }
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                 undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo,
                 undo_stack_.CanRedo() ? TBSTATE_ENABLED : 0);
  }
  return ok;
}

bool MainWindow::SameNode(const RegistryNode& left, const RegistryNode& right) const {
  if (left.root != right.root) {
    return false;
  }
  if (!EqualsInsensitive(left.subkey, right.subkey)) {
    return false;
  }
  return EqualsInsensitive(left.root_name, right.root_name);
}

std::wstring MainWindow::MakeUniqueValueName(const RegistryNode& node, const std::wstring& base) const {
  std::unordered_set<std::wstring> value_names;
  RegistryProvider::KeyEnumResult enum_result;
  bool names_reserved = false;
  RegistryProvider::EnumKeyStreaming(
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

std::wstring MainWindow::MakeUniqueKeyName(const RegistryNode& node, const std::wstring& base) const {
  auto keys = RegistryProvider::EnumSubKeyNames(node, false);
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

bool MainWindow::ResolvePathToNode(const std::wstring& path, RegistryNode* node) const {
  if (!node || path.empty()) {
    return false;
  }
  for (const auto& root_entry : roots_) {
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

void MainWindow::ShowValueHeaderMenu(POINT screen_pt) {
  HWND header_hwnd = ListView_GetHeader(value_list_.hwnd());
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

  for (size_t i = 0; i < value_columns_.size(); ++i) {
    UINT state = value_column_visible_[i] ? MF_CHECKED : MF_UNCHECKED;
    AppendMenuW(menu, MF_STRING | state, cmd::kHeaderToggleBase + static_cast<int>(i), value_columns_[i].title.c_str());
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command == cmd::kHeaderSizeToFit && last_header_column_ >= 0) {
    int subitem = GetListViewColumnSubItem(value_list_.hwnd(), last_header_column_);
    ListView_SetColumnWidth(value_list_.hwnd(), last_header_column_, LVSCW_AUTOSIZE_USEHEADER);
    int width = ListView_GetColumnWidth(value_list_.hwnd(), last_header_column_);
    if (subitem >= 0 && static_cast<size_t>(subitem) < value_column_widths_.size()) {
      value_column_widths_[static_cast<size_t>(subitem)] = width;
    }
    SaveSettings();
    return;
  }
  if (command == cmd::kHeaderSizeAll) {
    int last_visible = FindLastVisibleColumn(value_column_visible_);
    for (size_t i = 0; i < value_columns_.size(); ++i) {
      if (i < value_column_visible_.size() && !value_column_visible_[i]) {
        continue;
      }
      int display_index = FindListViewColumnBySubItem(value_list_.hwnd(), static_cast<int>(i));
      if (display_index < 0) {
        continue;
      }
      int width = 0;
      if (static_cast<int>(i) == last_visible) {
        width = CalcListViewColumnFitWidth(value_list_.hwnd(), static_cast<int>(i), value_columns_[i].width);
        ListView_SetColumnWidth(value_list_.hwnd(), display_index, width);
      } else {
        ListView_SetColumnWidth(value_list_.hwnd(), display_index, LVSCW_AUTOSIZE_USEHEADER);
        width = ListView_GetColumnWidth(value_list_.hwnd(), display_index);
      }
      value_column_widths_[i] = width;
    }
    SaveSettings();
    return;
  }
  if (command >= cmd::kHeaderToggleBase) {
    int index = command - cmd::kHeaderToggleBase;
    if (index >= 0 && static_cast<size_t>(index) < value_columns_.size()) {
      ToggleValueColumn(index, !value_column_visible_[static_cast<size_t>(index)]);
      SaveSettings();
    }
  }
}

void MainWindow::ShowHistoryHeaderMenu(POINT screen_pt) {
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

void MainWindow::ShowSearchHeaderMenu(POINT screen_pt) {
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

void MainWindow::ToggleValueColumn(int column, bool visible) {
  if (column < 0 || static_cast<size_t>(column) >= value_column_visible_.size()) {
    return;
  }
  if (visible == value_column_visible_[static_cast<size_t>(column)]) {
    return;
  }

  if (visible) {
    int width = value_column_widths_[static_cast<size_t>(column)];
    if (width <= 0) {
      width = value_columns_[static_cast<size_t>(column)].width;
    }
    value_column_visible_[static_cast<size_t>(column)] = true;
    value_column_widths_[static_cast<size_t>(column)] = width;
  } else {
    int display_index = FindListViewColumnBySubItem(value_list_.hwnd(), column);
    int width = display_index >= 0 ? ListView_GetColumnWidth(value_list_.hwnd(), display_index) : value_column_widths_[static_cast<size_t>(column)];
    if (width > 0) {
      value_column_widths_[static_cast<size_t>(column)] = width;
    }
    value_column_visible_[static_cast<size_t>(column)] = false;
  }
  ApplyValueColumns();
}

void MainWindow::ToggleHistoryColumn(int column, bool visible) {
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

void MainWindow::ToggleSearchColumn(int column, bool visible) {
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

void MainWindow::DrawAddressButton(const DRAWITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  const Theme& theme = Theme::Current();
  HDC hdc = info->hDC;
  RECT rect = info->rcItem;
  bool pressed = (info->itemState & ODS_SELECTED) != 0;

  COLORREF bg_color = pressed ? theme.HoverColor() : theme.SurfaceColor();
  FillRect(hdc, &rect, GetCachedBrush(bg_color));

  HPEN pen = GetCachedPen(theme.BorderColor());
  HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
  MoveToEx(hdc, rect.left, rect.top + 3, nullptr);
  LineTo(hdc, rect.left, rect.bottom - 3);
  SelectObject(hdc, old_pen);

  if (info->CtlID == kAddressGoId) {
    if (address_go_icon_) {
      UINT dpi = GetWindowDpi(hwnd_);
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
      HBRUSH arrow_brush = GetCachedBrush(arrow_color);
      HBRUSH old_brush = reinterpret_cast<HBRUSH>(SelectObject(hdc, arrow_brush));
      HPEN arrow_pen = GetCachedPen(arrow_color);
      HPEN old_arrow = reinterpret_cast<HPEN>(SelectObject(hdc, arrow_pen));
      Polygon(hdc, pts, 3);
      SelectObject(hdc, old_arrow);
      SelectObject(hdc, old_brush);
    }
  }
}

void MainWindow::DrawHeaderCloseButton(const DRAWITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  const Theme& theme = Theme::Current();
  HDC hdc = info->hDC;
  RECT rect = info->rcItem;
  bool pressed = (info->itemState & ODS_SELECTED) != 0;

  COLORREF bg_color = pressed ? theme.HoverColor() : theme.HeaderColor();
  FillRect(hdc, &rect, GetCachedBrush(bg_color));

  if (icon_font_) {
    HFONT old_font = reinterpret_cast<HFONT>(SelectObject(hdc, icon_font_));
    SetTextColor(hdc, theme.MutedTextColor());
    SetBkMode(hdc, TRANSPARENT);
    DrawTextW(hdc, L"\xE711", -1, &rect, DT_SINGLELINE | DT_VCENTER | DT_CENTER);
    SelectObject(hdc, old_font);
  }
}

void MainWindow::AddAddressHistory(const std::wstring& path) {
  if (path.empty()) {
    return;
  }
  auto it = std::find(address_history_.begin(), address_history_.end(), path);
  if (it != address_history_.end()) {
    address_history_.erase(it);
  }
  address_history_.insert(address_history_.begin(), path);
  if (address_history_.size() > 20) {
    address_history_.resize(20);
  }
}

bool MainWindow::SelectTreePath(const std::wstring& path) {
  if (!tree_.hwnd()) {
    return false;
  }
  std::vector<std::wstring> parts = BuildVisibleTreePathParts(path);
  if (parts.empty()) {
    return false;
  }

  HTREEITEM root = TreeView_GetRoot(tree_.hwnd());
  HTREEITEM current = root;
  for (const auto& part : parts) {
    TreeView_Expand(tree_.hwnd(), current, TVE_EXPAND);
    HTREEITEM child = FindChildByText(tree_.hwnd(), current, part);
    if (!child) {
      return false;
    }
    current = child;
  }

  if (current) {
    TreeView_SelectItem(tree_.hwnd(), current);
    TreeView_EnsureVisible(tree_.hwnd(), current);
    return true;
  }
  return false;
}

bool MainWindow::SelectValueByName(const std::wstring& name) {
  if (!value_list_.hwnd()) {
    return false;
  }
  for (size_t i = 0; i < value_list_.RowCount(); ++i) {
    const ListRow* row = value_list_.RowAt(static_cast<int>(i));
    if (!row || row->kind != rowkind::kValue) {
      continue;
    }
    if (row->extra == name) {
      ListView_SetItemState(value_list_.hwnd(), -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
      ListView_SetItemState(value_list_.hwnd(), static_cast<int>(i), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
      ListView_EnsureVisible(value_list_.hwnd(), static_cast<int>(i), FALSE);
      return true;
    }
  }
  return false;
}

void MainWindow::HandleTypeToSelectList(wchar_t ch) {
  if (!value_list_.hwnd()) {
    return;
  }
  DWORD now = GetTickCount();
  if (now - type_buffer_list_tick_ > kTypeSelectTimeoutMs) {
    type_buffer_list_.clear();
  }
  type_buffer_list_tick_ = now;
  if (ch == L'\b') {
    if (!type_buffer_list_.empty()) {
      type_buffer_list_.pop_back();
    }
  } else {
    type_buffer_list_.push_back(ch);
  }
  if (type_buffer_list_.empty()) {
    return;
  }
  int count = static_cast<int>(value_list_.RowCount());
  if (count <= 0) {
    return;
  }

  int match_index = -1;
  for (int i = 0; i < count; ++i) {
    const ListRow* row = value_list_.RowAt(i);
    if (!row) {
      continue;
    }
    if (StartsWithInsensitive(row->name, type_buffer_list_)) {
      match_index = i;
      break;
    }
  }

  if (match_index < 0) {
    int nearest_index = -1;
    std::wstring nearest_text;
    for (int i = 0; i < count; ++i) {
      const ListRow* row = value_list_.RowAt(i);
      if (!row) {
        continue;
      }
      if (CompareTextInsensitive(row->name, type_buffer_list_) >= 0) {
        if (nearest_index < 0 || CompareTextInsensitive(row->name, nearest_text) < 0) {
          nearest_index = i;
          nearest_text = row->name;
        }
      }
    }
    if (nearest_index >= 0) {
      match_index = nearest_index;
    } else {
      match_index = count - 1;
    }
  }

  ListView_SetItemState(value_list_.hwnd(), -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
  ListView_SetItemState(value_list_.hwnd(), match_index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
  ListView_EnsureVisible(value_list_.hwnd(), match_index, FALSE);
}

void MainWindow::HandleTypeToSelectTree(wchar_t ch) {
  if (!tree_.hwnd()) {
    return;
  }
  DWORD now = GetTickCount();
  if (now - type_buffer_tree_tick_ > kTypeSelectTimeoutMs) {
    type_buffer_tree_.clear();
  }
  type_buffer_tree_tick_ = now;
  if (ch == L'\b') {
    if (!type_buffer_tree_.empty()) {
      type_buffer_tree_.pop_back();
    }
  } else {
    type_buffer_tree_.push_back(ch);
  }
  if (type_buffer_tree_.empty()) {
    return;
  }
  HTREEITEM selected = TreeView_GetSelection(tree_.hwnd());
  if (!selected) {
    return;
  }

  auto ensure_children_loaded = [&](HTREEITEM item) {
    RegistryNode* node = tree_.NodeFromItem(item);
    if (!node || node->children_loaded) {
      return;
    }
    NMTREEVIEWW info = {};
    info.action = TVE_EXPAND;
    info.itemNew.hItem = item;
    tree_.OnItemExpanding(&info);
  };

  auto collect_children = [&](HTREEITEM parent, std::vector<HTREEITEM>* out) {
    for (HTREEITEM child = TreeView_GetChild(tree_.hwnd(), parent); child; child = TreeView_GetNextSibling(tree_.hwnd(), child)) {
      out->push_back(child);
    }
  };

  auto get_item_text = [&](HTREEITEM item, std::wstring* text_out) -> bool {
    wchar_t text[256] = {};
    TVITEMW tvi = {};
    tvi.hItem = item;
    tvi.mask = TVIF_TEXT;
    tvi.pszText = text;
    tvi.cchTextMax = static_cast<int>(_countof(text));
    if (!TreeView_GetItem(tree_.hwnd(), &tvi)) {
      return false;
    }
    if (text_out) {
      *text_out = text;
    }
    return true;
  };

  auto find_match = [&](const std::vector<HTREEITEM>& items, HTREEITEM start_item) -> HTREEITEM {
    if (items.empty()) {
      return nullptr;
    }
    size_t start_index = 0;
    for (size_t i = 0; i < items.size(); ++i) {
      if (items[i] == start_item) {
        start_index = i;
        break;
      }
    }
    for (size_t offset = 0; offset < items.size(); ++offset) {
      size_t idx = (start_index + offset) % items.size();
      std::wstring text;
      if (get_item_text(items[idx], &text) && EqualsInsensitive(text, type_buffer_tree_)) {
        return items[idx];
      }
    }
    for (size_t offset = 0; offset < items.size(); ++offset) {
      size_t idx = (start_index + offset) % items.size();
      std::wstring text;
      if (get_item_text(items[idx], &text) && StartsWithInsensitive(text, type_buffer_tree_)) {
        return items[idx];
      }
    }
    return nullptr;
  };

  HTREEITEM target = nullptr;
  if (tree_type_to_select_descend_) {
    std::vector<HTREEITEM> children;
    ensure_children_loaded(selected);
    collect_children(selected, &children);
    target = find_match(children, selected);
  }
  HTREEITEM parent = TreeView_GetParent(tree_.hwnd(), selected);
  if (!target && parent) {
    std::vector<HTREEITEM> siblings;
    ensure_children_loaded(parent);
    collect_children(parent, &siblings);
    target = find_match(siblings, selected);
  }
  if (!target && !tree_type_to_select_descend_) {
    std::vector<HTREEITEM> children;
    ensure_children_loaded(selected);
    collect_children(selected, &children);
    target = find_match(children, selected);
  }
  if (!target) {
    return;
  }
  tree_type_to_select_descend_ = false;
  TreeView_SelectItem(tree_.hwnd(), target);
  TreeView_EnsureVisible(tree_.hwnd(), target);
}

namespace {

bool ParseBundledTraceLabel(const std::wstring& path, std::wstring* label) {
  if (label) {
    label->clear();
  }
  if (path.size() < 4) {
    return false;
  }
  if (!StartsWithInsensitive(path, L"res:")) {
    return false;
  }
  std::wstring key = path.substr(4);
  if (key.empty()) {
    return false;
  }
  if (label) {
    *label = key;
  }
  return true;
}

} // namespace

void MainWindow::StartTraceDialogLoad(HWND hwnd, void* context) {
  auto* ctx = reinterpret_cast<TraceDialogStartContext*>(context);
  if (!ctx || !ctx->window || !ctx->session) {
    return;
  }
  ctx->session->dialog = hwnd;
  ctx->window->StartTraceParseThread(ctx->session);
}

void MainWindow::StartDefaultDialogLoad(HWND hwnd, void* context) {
  auto* ctx = reinterpret_cast<DefaultDialogStartContext*>(context);
  if (!ctx || !ctx->window || !ctx->session) {
    return;
  }
  ctx->session->dialog = hwnd;
  ctx->window->StartDefaultParseThread(ctx->session);
}

bool MainWindow::AllowTraceSimulation(const RegistryNode& node) const {
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

std::wstring MainWindow::TracePathLowerForNode(const RegistryNode& node) const {
  std::wstring path = registry_path::Build(node);
  std::wstring trace_path = NormalizeTraceKeyPath(path);
  if (trace_path.empty()) {
    trace_path = path;
  }
  return ToLower(trace_path);
}

void MainWindow::AppendTraceChildren(const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) const {
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

std::wstring MainWindow::ResolveBundledTracePath(const std::wstring& label) const {
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
  std::wstring records = util::JoinPath(module_dir, L"records");
  return util::JoinPath(records, file);
}

bool MainWindow::LoadBundledTrace(
    const std::wstring& label,
    const trace::Selection* selection_override) {
  std::wstring path = ResolveBundledTracePath(label);
  if (path.empty()) {
    return false;
  }
  return LoadTraceFromFile(label, path, selection_override);
}

std::wstring MainWindow::ResolveBundledDefaultPath(const std::wstring& label) const {
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
  return util::JoinPath(defaults, file);
}

bool MainWindow::AddTraceFromFile(
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
    UpdateValueListForNode(current_node_);
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

bool MainWindow::LoadTraceFromFile(
    const std::wstring& label, const std::wstring& path,
    const trace::Selection* selection_override) {
  return AddTraceFromFile(label, path, selection_override, true, true);
}

bool MainWindow::LoadTraceFromPrompt() {
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

void MainWindow::ClearTrace() {
  StopTraceParseSessions();
  active_traces_.clear();
  trace_selection_cache_.clear();
  SaveActiveTraces();
  SaveTraceSettings();
  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(current_node_);
  SaveSettings();
}

bool MainWindow::AddDefaultFromFile(const std::wstring& label, const std::wstring& path, bool show_error, bool prompt_for_selection, bool update_ui) {
  if (path.empty()) {
    return false;
  }
  std::wstring source = path;
  std::wstring use_label = label;
  if (!FileExists(source)) {
    std::wstring bundled = ResolveBundledDefaultPath(path);
    if (!bundled.empty() && FileExists(bundled)) {
      source = bundled;
      if (use_label.empty()) {
        use_label = path;
      }
    } else {
      if (show_error) {
        ui::ShowError(hwnd_, L"Default file not found.");
      }
      return false;
    }
  }
  if (use_label.empty()) {
    use_label = FileBaseName(source);
  }
  if (use_label.empty()) {
    use_label = L"Default";
  }
  for (const auto& defaults : active_defaults_) {
    if (EqualsInsensitive(defaults.source_path, source)) {
      return false;
    }
  }
  std::wstring source_lower = ToLower(source);
  if (default_parse_sessions_.find(source_lower) != default_parse_sessions_.end()) {
    return false;
  }

  trace::Selection selection = {};
  selection.select_all = true;
  selection.recursive = true;

  auto session = std::make_unique<DefaultParseSession>();
  session->label = use_label;
  session->source_path = source;
  session->source_lower = source_lower;
  session->data = std::make_shared<defaults::Data>();
  session->selection = selection;
  session->show_errors = show_error;

  DefaultParseSession* session_ptr = session.get();
  default_parse_sessions_.emplace(source_lower, std::move(session));

  if (prompt_for_selection) {
    trace::Selection dialog_selection = selection;
    TraceDialogOptions options;
    options.title = use_label.empty() ? L"Default entries" : L"Default entries - " + use_label;
    options.prompt = L"";
    options.show_values = true;
    DefaultDialogStartContext context;
    context.window = this;
    context.session = session_ptr;
    if (!ShowTraceDialog(hwnd_, options, &dialog_selection, StartDefaultDialogLoad, &context)) {
      session_ptr->work.CancelAndJoin();
      default_parse_sessions_.erase(source_lower);
      return false;
    }
    session_ptr->dialog = nullptr;
    session_ptr->selection = std::move(dialog_selection);
  } else {
    StartDefaultParseThread(session_ptr);
  }

  if (!session_ptr->selection.select_all && session_ptr->selection.key_paths.empty() && session_ptr->selection.values_by_key.empty()) {
    session_ptr->selection.select_all = true;
  }

  session_ptr->added_to_active = true;
  active_defaults_.push_back(
      {use_label, source, session_ptr->data,
       std::make_shared<trace::Selection>(session_ptr->selection)});
  if (update_ui) {
    SaveActiveDefaults();
    BuildMenus();
    UpdateValueListForNode(current_node_);
    SaveSettings();
  }
  if (session_ptr->parsing_done) {
    session_ptr->work.Join();
  }
  if (session_ptr->parsing_done && !session_ptr->dialog) {
    default_parse_sessions_.erase(source_lower);
  }
  return true;
}

bool MainWindow::SaveRegFileTab(int tab_index) {
  if (!IsRegFileTabIndex(tab_index) || static_cast<size_t>(tab_index) >= tabs_.size()) {
    return false;
  }
  TabEntry& entry = tabs_[static_cast<size_t>(tab_index)];
  if (entry.reg_file_path.empty()) {
    return false;
  }
  std::wstring content;
  if (!BuildRegFileContent(entry, &content)) {
    return false;
  }
  if (!util::WriteTextFile(entry.reg_file_path, content, true)) {
    ui::ShowError(hwnd_, L"Failed to save registry file.");
    return false;
  }
  if (entry.reg_file_dirty) {
    entry.reg_file_dirty = false;
    BuildMenus();
  }
  AppendHistoryEntry(L"Save .reg file " + FileNameOnly(entry.reg_file_path), L"", entry.reg_file_path);
  return true;
}

bool MainWindow::ExportRegFileTab(int tab_index, const std::wstring& path) {
  if (!IsRegFileTabIndex(tab_index) || static_cast<size_t>(tab_index) >= tabs_.size()) {
    return false;
  }
  if (path.empty()) {
    return false;
  }
  std::wstring content;
  if (!BuildRegFileContent(tabs_[static_cast<size_t>(tab_index)], &content)) {
    return false;
  }
  std::wstring target = EnsureRegExtension(path);
  if (!util::WriteTextFile(target, content, true)) {
    ui::ShowError(hwnd_, L"Failed to export registry file.");
    return false;
  }
  return true;
}

bool MainWindow::BuildRegFileContent(const TabEntry& entry, std::wstring* out) const {
  if (!out) {
    return false;
  }
  out->clear();
  if (entry.kind != TabEntry::Kind::kRegFile) {
    return false;
  }

  regfile::Writer writer;
  std::function<void(const VirtualRegistryKey&,
                     const std::wstring&)>
      append_key;
  append_key = [&](const VirtualRegistryKey& key,
                    const std::wstring& full_path) {
    if (!key.values.empty()) {
      std::vector<const VirtualRegistryValue*> values;
      values.reserve(key.values.size());
      for (const auto& source : key.values) {
        values.push_back(&source.second);
      }
      writer.AppendKey(full_path, std::move(values));
    }
    std::vector<const VirtualRegistryKey*> children;
    children.reserve(key.children.size());
    for (const auto& child : key.children) {
      if (child.second) {
        children.push_back(child.second.get());
      }
    }
    std::sort(children.begin(), children.end(), [](const VirtualRegistryKey* left, const VirtualRegistryKey* right) {
      if (!left || !right) {
        return left != nullptr;
      }
      return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
    });
    for (const auto* child : children) {
      if (!child) {
        continue;
      }
      append_key(*child, full_path + L"\\" + child->name);
    }
  };

  for (const auto& root : entry.reg_file_roots) {
    if (!root.data || !root.data->root) {
      continue;
    }
    std::wstring root_name = root.name;
    if (root_name.empty()) {
      root_name = root.data->root_name;
    }
    if (root_name.empty()) {
      continue;
    }
    append_key(*root.data->root, root_name);
  }
  *out = std::move(writer).Finish();
  return true;
}

void MainWindow::ReleaseRegFileRoots(TabEntry* entry) {
  if (!entry) {
    return;
  }
  for (auto& root : entry->reg_file_roots) {
    if (root.root) {
      RegistryProvider::UnregisterVirtualRoot(root.root);
      root.root = nullptr;
    }
    root.data.reset();
  }
  entry->reg_file_roots.clear();
}

bool MainWindow::OpenRegFileTab(const std::wstring& path) {
  if (!tab_ || path.empty()) {
    return false;
  }
  if (!FileExists(path)) {
    ui::ShowError(hwnd_, L"Registry file not found.");
    return false;
  }
  std::wstring label = FileNameOnly(path);
  if (label.empty()) {
    label = L"Registry File";
  }
  std::wstring path_lower = ToLower(path);
  auto start_parse = [&]() {
    if (reg_file_parse_sessions_.find(path_lower) != reg_file_parse_sessions_.end()) {
      return;
    }
    auto session = std::make_unique<RegFileParseSession>();
    session->source_path = path;
    session->source_lower = path_lower;
    HWND hwnd = hwnd_;
    RegFileParseSession* session_ptr = session.get();
    session->work.Start([this, session_ptr, hwnd](
                            uint64_t generation,
                            std::atomic_bool& cancel) {
      auto payload = std::make_unique<RegFileParsePayload>();
      payload->generation = generation;
      payload->source_path = session_ptr->source_path;
      payload->source_lower = session_ptr->source_lower;
      std::wstring parse_error;
      std::vector<ParsedRegFileRoot> parsed_roots;
      bool cancelled = false;
      if (!ParseRegFileToVirtualRoots(payload->source_path, &parsed_roots,
                                      &parse_error, &cancel,
                                      &cancelled)) {
        if (!cancelled && parse_error.empty()) {
          parse_error = L"Failed to read registry file.";
        }
      }
      payload->roots = std::move(parsed_roots);
      payload->error = std::move(parse_error);
      payload->cancelled = cancelled;
      if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, kRegFileLoadReadyMessage, 0, reinterpret_cast<LPARAM>(payload.get()))) {
        return;
      }
      ReleasePostedPayload(payload);
    });
    reg_file_parse_sessions_.emplace(path_lower, std::move(session));
  };

  for (size_t i = 0; i < tabs_.size(); ++i) {
    TabEntry& entry = tabs_[i];
    if (entry.kind != TabEntry::Kind::kRegFile) {
      continue;
    }
    if (EqualsInsensitive(entry.reg_file_path, path)) {
      entry.reg_file_path = path;
      entry.reg_file_label = label;
      entry.reg_file_loading = true;
      TCITEMW item = {};
      item.mask = TCIF_TEXT;
      item.pszText = const_cast<wchar_t*>(label.c_str());
      TabCtrl_SetItem(tab_, static_cast<int>(i), &item);
      TabCtrl_SetCurSel(tab_, static_cast<int>(i));
      SyncRegFileTabSelection();
      ApplyViewVisibility();
      UpdateStatus();
      start_parse();
      AppendHistoryEntry(L"Open .reg file " + FileNameOnly(path), L"", path);
      return true;
    }
  }

  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(label.c_str());
  int index = TabCtrl_GetItemCount(tab_);
  TabCtrl_InsertItem(tab_, index, &item);
  TabEntry entry;
  entry.kind = TabEntry::Kind::kRegFile;
  entry.reg_file_path = path;
  entry.reg_file_label = label;
  entry.reg_file_dirty = false;
  entry.reg_file_loading = true;
  tabs_.push_back(std::move(entry));
  UpdateTabWidth();
  TabCtrl_SetCurSel(tab_, index);
  SyncRegFileTabSelection();
  ApplyViewVisibility();
  UpdateStatus();
  start_parse();
  AppendHistoryEntry(L"Open .reg file " + FileNameOnly(path), L"", path);
  return true;
}

bool MainWindow::LoadDefaultFromFile(const std::wstring& label, const std::wstring& path) {
  return AddDefaultFromFile(label, path, true, true, true);
}

bool MainWindow::LoadDefaultFromPrompt() {
  std::wstring path;
  if (!PromptOpenFile(hwnd_, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
    return false;
  }
  std::wstring label = FileBaseName(path);
  if (label.empty()) {
    label = L"Custom";
  }
  if (!LoadDefaultFromFile(label, path)) {
    return false;
  }
  AddRecentDefaultPath(path);
  BuildMenus();
  SaveSettings();
  return true;
}

void MainWindow::ClearDefaults() {
  StopDefaultParseSessions();
  active_defaults_.clear();
  SaveActiveDefaults();
  BuildMenus();
  UpdateValueListForNode(current_node_);
  SaveSettings();
}

void MainWindow::NormalizeRecentTraceList() {
  recent_trace_paths_.Normalize();
}

void MainWindow::NormalizeRecentDefaultList() {
  recent_default_paths_.Normalize();
}

void MainWindow::AddRecentTracePath(const std::wstring& path) {
  recent_trace_paths_.Add(path);
}

void MainWindow::AddRecentDefaultPath(const std::wstring& path) {
  recent_default_paths_.Add(path);
}

} // namespace regkit
