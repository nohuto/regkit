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

#include "app/app_window.h"

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

#include "app/command_ids.h"
#include "app/registry_io.h"
#include "app/registry_security.h"
#include "app/ui_helpers.h"
#include "app/value_dialogs.h"
#include "registry/registry_provider.h"
#include "resource.h"
#include "win32/icon_resources.h"
#include "win32/win32_helpers.h"

namespace regkit {

struct MainWindow::TraceLoadPayload {
  std::vector<ActiveTrace> traces;
  std::unordered_map<std::wstring, TraceSelection> selection_cache;
};

struct MainWindow::DefaultLoadPayload {
  std::vector<ActiveDefault> defaults;
};

namespace {
constexpr int kToolbarId = 100;
constexpr int kAddressEditId = 101;
constexpr int kTabId = 103;
constexpr int kTreeId = 104;
constexpr int kValueListId = 105;
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
constexpr int kMaxRecentTraces = cmd::kTraceRecentMax - cmd::kTraceRecentBase + 1;
constexpr int kMaxRecentDefaults = cmd::kDefaultRecentMax - cmd::kDefaultRecentBase + 1;

bool ParseBundledTraceLabel(const std::wstring& path, std::wstring* label);
constexpr wchar_t kStandardGroupLabel[] = L"Standart Hives";
constexpr wchar_t kRealGroupLabel[] = L"REGISTRY";

constexpr UINT kAddressEnterMessage = WM_APP + 10;
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
constexpr UINT_PTR kAddressSubclassId = 1;
constexpr UINT_PTR kTabSubclassId = 2;
constexpr UINT_PTR kHeaderSubclassId = 3;
constexpr UINT_PTR kListViewSubclassId = 4;
constexpr UINT_PTR kTreeViewSubclassId = 5;
constexpr UINT_PTR kAutoCompletePopupSubclassId = 6;
constexpr UINT_PTR kAutoCompleteListBoxSubclassId = 7;
constexpr UINT_PTR kFilterSubclassId = 8;

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

struct TraceParseBatch {
  std::wstring source_lower;
  std::vector<KeyValueDialogEntry> entries;
  std::wstring error;
  bool done = false;
  bool cancelled = false;
};

struct DefaultParseBatch {
  std::wstring source_lower;
  std::vector<KeyValueDialogEntry> entries;
  std::wstring error;
  bool done = false;
  bool cancelled = false;
};

struct ValueListPayload {
  uint64_t generation = 0;
  std::vector<ListRow> rows;
  int key_count = 0;
  int value_count = 0;
};

HBRUSH GetCachedBrush(COLORREF color);
HPEN GetCachedPen(COLORREF color, int width = 1);
std::vector<std::wstring> SplitPath(const std::wstring& path);
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

bool DrawSearchMatchSubItem(const SearchResult& result, int subitem, bool selected, HDC hdc, const RECT& rect, HFONT font) {
  bool match_subitem = false;
  if (result.match_field == SearchMatchField::kPath && subitem == 0) {
    match_subitem = true;
  } else if (result.match_field == SearchMatchField::kName && subitem == 1) {
    match_subitem = true;
  } else if (result.match_field == SearchMatchField::kData && subitem == 3) {
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
  COLORREF bg = selected ? theme.SelectionColor() : theme.PanelColor();
  COLORREF fg = selected ? theme.SelectionTextColor() : theme.TextColor();
  HBRUSH bg_brush = GetCachedBrush(bg);
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

void DrawHistoryListItem(HWND list, HDC hdc, int item_index, bool selected, bool hot, HFONT font) {
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
  if (selected) {
    bg = theme.SelectionColor();
    fg = theme.SelectionTextColor();
  } else if (hot) {
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
    bool hot = (draw->nmcd.uItemState & CDIS_HOT) != 0;
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(list, WM_GETFONT, 0, 0));
    DrawHistoryListItem(list, draw->nmcd.hdc, item_index, selected, hot, font);
    if (selected) {
      const Theme& theme = Theme::Current();
      COLORREF border = (GetFocus() == list) ? theme.FocusColor() : theme.BorderColor();
      ui::DrawListViewFocusBorder(list, draw->nmcd.hdc, item_index, border);
    }
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
  item.pszText = buffer->data();
  item.cchTextMax = static_cast<int>(buffer->size());
  int length = static_cast<int>(SendMessageW(list, LVM_GETITEMTEXTW, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&item)));
  if (length >= static_cast<int>(buffer->size() - 1)) {
    buffer->resize(static_cast<size_t>(length) + 2);
    item.pszText = buffer->data();
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

std::wstring FormatWin32Error(DWORD code) {
  if (code == 0) {
    return L"";
  }
  wchar_t buffer[512] = {};
  DWORD len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, code, 0, buffer, static_cast<DWORD>(_countof(buffer)), nullptr);
  if (len == 0) {
    return L"Unknown error.";
  }
  return buffer;
}

bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  wchar_t buffer[MAX_PATH] = {};
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer;
  ofn.nMaxFile = static_cast<DWORD>(_countof(buffer));
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
  if (!GetOpenFileNameW(&ofn)) {
    return false;
  }
  *path = buffer;
  return true;
}

bool PromptSaveFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  wchar_t buffer[MAX_PATH] = {};
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer;
  ofn.nMaxFile = static_cast<DWORD>(_countof(buffer));
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
  if (!GetSaveFileNameW(&ofn)) {
    return false;
  }
  *path = buffer;
  return true;
}

std::wstring TrimTrailingSeparators(const std::wstring& path) {
  std::wstring result = path;
  while (!result.empty() && (result.back() == L'\\' || result.back() == L'/')) {
    result.pop_back();
  }
  return result;
}

std::wstring ParentPath(const std::wstring& path) {
  std::wstring trimmed = TrimTrailingSeparators(path);
  size_t pos = trimmed.find_last_of(L"\\/");
  if (pos == std::wstring::npos) {
    return L"";
  }
  return trimmed.substr(0, pos);
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
    base = ParentPath(base);
  }
  DWORD len = GetCurrentDirectoryW(0, nullptr);
  if (len > 0) {
    std::wstring cwd(len, L'\0');
    DWORD written = GetCurrentDirectoryW(len, cwd.data());
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
        base = ParentPath(base);
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

std::wstring TrimWhitespace(const std::wstring& text) {
  size_t start = 0;
  while (start < text.size() && (text[start] == L' ' || text[start] == L'\t')) {
    ++start;
  }
  if (start == text.size()) {
    return L"";
  }
  size_t end = text.size() - 1;
  while (end > start && (text[end] == L' ' || text[end] == L'\t')) {
    --end;
  }
  return text.substr(start, end - start + 1);
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

bool ReadFileBinary(const std::wstring& path, std::string* buffer) {
  if (!buffer) {
    return false;
  }
  buffer->clear();
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return false;
  }
  std::string temp(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, temp.data(), static_cast<DWORD>(temp.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return false;
  }
  temp.resize(read);
  *buffer = std::move(temp);
  return true;
}

bool ReadFileUtf8(const std::wstring& path, std::wstring* content) {
  if (!content) {
    return false;
  }
  content->clear();
  std::string buffer;
  if (!ReadFileBinary(path, &buffer)) {
    return false;
  }
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring wide = util::Utf8ToWide(buffer);
  if (wide.empty()) {
    return false;
  }
  *content = std::move(wide);
  return true;
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
      wchar_t expanded[512] = {};
      DWORD expanded_len = ExpandEnvironmentStringsW(value.c_str(), expanded, static_cast<DWORD>(_countof(expanded)));
      if (expanded_len > 0 && expanded_len < _countof(expanded)) {
        value.assign(expanded, expanded_len - 1);
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

std::wstring ToLower(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (wchar_t ch : text) {
    out.push_back(static_cast<wchar_t>(towlower(ch)));
  }
  return out;
}

bool ReadRegFileText(const std::wstring& path, std::wstring* out) {
  if (!out) {
    return false;
  }
  out->clear();
  std::string buffer;
  if (!ReadFileBinary(path, &buffer)) {
    return false;
  }
  if (buffer.size() >= 2 && static_cast<unsigned char>(buffer[0]) == 0xFF && static_cast<unsigned char>(buffer[1]) == 0xFE) {
    size_t wchar_count = (buffer.size() - 2) / sizeof(wchar_t);
    out->assign(reinterpret_cast<const wchar_t*>(buffer.data() + 2), wchar_count);
    return !out->empty();
  }
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  *out = util::Utf8ToWide(buffer);
  return !out->empty();
}

bool ParseQuotedString(const std::wstring& text, std::wstring* out, size_t* end_pos) {
  if (!out || text.empty() || text.front() != L'"') {
    return false;
  }
  out->clear();
  bool escape = false;
  for (size_t i = 1; i < text.size(); ++i) {
    wchar_t ch = text[i];
    if (escape) {
      switch (ch) {
      case L'\\':
        out->push_back(L'\\');
        break;
      case L'"':
        out->push_back(L'"');
        break;
      case L'n':
        out->push_back(L'\n');
        break;
      case L'r':
        out->push_back(L'\r');
        break;
      case L't':
        out->push_back(L'\t');
        break;
      case L'0':
        out->push_back(L'\0');
        break;
      default:
        out->push_back(ch);
        break;
      }
      escape = false;
      continue;
    }
    if (ch == L'\\') {
      escape = true;
      continue;
    }
    if (ch == L'"') {
      if (end_pos) {
        *end_pos = i + 1;
      }
      return true;
    }
    out->push_back(ch);
  }
  return false;
}

bool ParseHexBytes(const std::wstring& text, std::vector<BYTE>* out) {
  if (!out) {
    return false;
  }
  out->clear();
  int nibble = -1;
  for (wchar_t ch : text) {
    if (iswxdigit(ch)) {
      int value = 0;
      if (ch >= L'0' && ch <= L'9') {
        value = ch - L'0';
      } else if (ch >= L'a' && ch <= L'f') {
        value = 10 + (ch - L'a');
      } else if (ch >= L'A' && ch <= L'F') {
        value = 10 + (ch - L'A');
      }
      if (nibble < 0) {
        nibble = value;
      } else {
        out->push_back(static_cast<BYTE>((nibble << 4) | value));
        nibble = -1;
      }
    }
  }
  return nibble < 0;
}

std::vector<BYTE> StringToRegData(const std::wstring& text) {
  std::vector<BYTE> data((text.size() + 1) * sizeof(wchar_t));
  memcpy(data.data(), text.c_str(), data.size());
  return data;
}

bool DecodeRegString(const std::vector<BYTE>& data, std::wstring* out) {
  if (!out) {
    return false;
  }
  out->clear();
  if (data.empty()) {
    return true;
  }
  if (data.size() % sizeof(wchar_t) != 0) {
    return false;
  }
  size_t wchar_count = data.size() / sizeof(wchar_t);
  const wchar_t* raw = reinterpret_cast<const wchar_t*>(data.data());
  std::wstring text(raw, wchar_count);
  while (!text.empty() && text.back() == L'\0') {
    text.pop_back();
  }
  if (text.find(L'\0') != std::wstring::npos) {
    return false;
  }
  *out = std::move(text);
  return true;
}

std::wstring EscapeRegString(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (wchar_t ch : text) {
    switch (ch) {
    case L'\\':
      out.append(L"\\\\");
      break;
    case L'"':
      out.append(L"\\\"");
      break;
    case L'\n':
      out.append(L"\\n");
      break;
    case L'\r':
      out.append(L"\\r");
      break;
    case L'\t':
      out.append(L"\\t");
      break;
    case L'\0':
      out.append(L"\\0");
      break;
    default:
      out.push_back(ch);
      break;
    }
  }
  return out;
}

std::wstring FormatHexBytes(const std::vector<BYTE>& data) {
  std::wstring out;
  if (!data.empty()) {
    out.reserve(data.size() * 3);
  }
  for (size_t i = 0; i < data.size(); ++i) {
    if (i > 0) {
      out.push_back(L',');
    }
    wchar_t buffer[4] = {};
    swprintf_s(buffer, L"%02x", data[i]);
    out.append(buffer);
  }
  return out;
}

DWORD RegTypeCode(DWORD type) {
  DWORD base = RegistryProvider::NormalizeValueType(type);
  switch (base) {
  case REG_NONE:
    return 0x0;
  case REG_SZ:
    return 0x1;
  case REG_EXPAND_SZ:
    return 0x2;
  case REG_BINARY:
    return 0x3;
  case REG_DWORD:
    return 0x4;
  case REG_DWORD_BIG_ENDIAN:
    return 0x5;
  case REG_LINK:
    return 0x6;
  case REG_MULTI_SZ:
    return 0x7;
  case REG_RESOURCE_LIST:
    return 0x8;
  case REG_FULL_RESOURCE_DESCRIPTOR:
    return 0x9;
  case REG_RESOURCE_REQUIREMENTS_LIST:
    return 0xA;
  case REG_QWORD:
    return 0xB;
  default:
    return base;
  }
}

std::wstring FormatRegValueData(DWORD type, const std::vector<BYTE>& data) {
  DWORD base = RegistryProvider::NormalizeValueType(type);
  if (base == REG_SZ) {
    std::wstring text;
    if (DecodeRegString(data, &text)) {
      return L"\"" + EscapeRegString(text) + L"\"";
    }
  }
  if (base == REG_DWORD && data.size() >= sizeof(DWORD)) {
    DWORD value = 0;
    memcpy(&value, data.data(), sizeof(DWORD));
    wchar_t buffer[16] = {};
    swprintf_s(buffer, L"dword:%08x", value);
    return buffer;
  }
  std::wstring hex = FormatHexBytes(data);
  if (base == REG_BINARY && type == REG_BINARY) {
    return L"hex:" + hex;
  }
  DWORD code = RegTypeCode(type);
  wchar_t type_buffer[16] = {};
  swprintf_s(type_buffer, L"%x", code);
  return L"hex(" + std::wstring(type_buffer) + L"):" + hex;
}

bool WriteRegFileText(const std::wstring& path, const std::wstring& text) {
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  const wchar_t bom = 0xFEFF;
  DWORD written = 0;
  if (!WriteFile(file, &bom, static_cast<DWORD>(sizeof(bom)), &written, nullptr)) {
    CloseHandle(file);
    return false;
  }
  if (!text.empty()) {
    DWORD bytes = static_cast<DWORD>(text.size() * sizeof(wchar_t));
    if (!WriteFile(file, text.data(), bytes, &written, nullptr)) {
      CloseHandle(file);
      return false;
    }
  }
  CloseHandle(file);
  return true;
}

struct ParsedRegFileRoot {
  std::wstring name;
  std::shared_ptr<RegistryProvider::VirtualRegistryData> data;
};

struct RegFileParsePayload {
  std::wstring source_path;
  std::wstring source_lower;
  std::vector<ParsedRegFileRoot> roots;
  std::wstring error;
  bool cancelled = false;
};

RegistryProvider::VirtualRegistryKey* EnsureVirtualKey(RegistryProvider::VirtualRegistryKey* root, const std::wstring& subkey) {
  if (!root) {
    return nullptr;
  }
  if (subkey.empty()) {
    return root;
  }
  auto parts = SplitPath(subkey);
  RegistryProvider::VirtualRegistryKey* current = root;
  for (const auto& part : parts) {
    std::wstring lower = ToLower(part);
    auto it = current->children.find(lower);
    if (it == current->children.end()) {
      auto child = std::make_unique<RegistryProvider::VirtualRegistryKey>();
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
  if (cancelled) {
    *cancelled = false;
  }
  auto is_cancelled = [&]() -> bool {
    if (cancel && cancel->load()) {
      if (cancelled) {
        *cancelled = true;
      }
      return true;
    }
    return false;
  };
  if (is_cancelled()) {
    return false;
  }
  std::wstring content;
  if (!ReadRegFileText(path, &content)) {
    if (error) {
      *error = L"Failed to read registry file.";
    }
    return false;
  }

  std::vector<std::wstring> lines;
  std::wstring current;
  size_t start = 0;
  while (start < content.size()) {
    if (is_cancelled()) {
      return false;
    }
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    if (current.empty()) {
      current = line;
    } else {
      current.append(line);
    }
    std::wstring trimmed_right = current;
    while (!trimmed_right.empty() && (trimmed_right.back() == L' ' || trimmed_right.back() == L'\t')) {
      trimmed_right.pop_back();
    }
    if (!trimmed_right.empty() && trimmed_right.back() == L'\\') {
      trimmed_right.pop_back();
      current = trimmed_right;
      continue;
    }
    lines.push_back(current);
    current.clear();
  }
  if (!current.empty()) {
    lines.push_back(current);
  }

  std::unordered_map<std::wstring, size_t> root_lookup;
  auto ensure_root = [&](const std::wstring& root_name) -> RegistryProvider::VirtualRegistryData* {
    std::wstring lower = ToLower(root_name);
    auto it = root_lookup.find(lower);
    if (it != root_lookup.end()) {
      return roots->at(it->second).data.get();
    }
    ParsedRegFileRoot root;
    root.name = root_name;
    root.data = std::make_shared<RegistryProvider::VirtualRegistryData>();
    root.data->root_name = root_name;
    root.data->root = std::make_unique<RegistryProvider::VirtualRegistryKey>();
    root.data->root->name = root_name;
    roots->push_back(std::move(root));
    root_lookup.emplace(lower, roots->size() - 1);
    return roots->back().data.get();
  };

  RegistryProvider::VirtualRegistryKey* current_key = nullptr;
  for (const auto& raw : lines) {
    if (is_cancelled()) {
      return false;
    }
    std::wstring line = TrimWhitespace(raw);
    if (line.empty()) {
      continue;
    }
    if (line[0] == L';') {
      continue;
    }
    if (StartsWithInsensitive(line, L"Windows Registry Editor") || StartsWithInsensitive(line, L"REGEDIT4")) {
      continue;
    }
    if (line.front() == L'[' && line.back() == L']') {
      std::wstring key = line.substr(1, line.size() - 2);
      key = TrimWhitespace(key);
      bool delete_key = !key.empty() && key.front() == L'-';
      if (delete_key) {
        current_key = nullptr;
        continue;
      }
      std::wstring normalized = NormalizeTraceKeyPathBasic(key);
      std::wstring key_path = normalized.empty() ? key : normalized;
      size_t slash = key_path.find(L'\\');
      std::wstring root_name = (slash == std::wstring::npos) ? key_path : key_path.substr(0, slash);
      std::wstring subkey = (slash == std::wstring::npos) ? L"" : key_path.substr(slash + 1);
      if (root_name.empty()) {
        current_key = nullptr;
        continue;
      }
      RegistryProvider::VirtualRegistryData* data = ensure_root(root_name);
      current_key = data ? EnsureVirtualKey(data->root.get(), subkey) : nullptr;
      continue;
    }

    if (!current_key) {
      continue;
    }
    size_t eq = line.find(L'=');
    if (eq == std::wstring::npos) {
      continue;
    }
    std::wstring name_part = TrimWhitespace(line.substr(0, eq));
    std::wstring data_part = TrimWhitespace(line.substr(eq + 1));
    if (name_part.empty() || data_part.empty()) {
      continue;
    }
    if (data_part == L"-") {
      continue;
    }

    std::wstring value_name;
    if (name_part == L"@") {
      value_name.clear();
    } else if (name_part.front() == L'"') {
      size_t end_pos = 0;
      if (!ParseQuotedString(name_part, &value_name, &end_pos)) {
        continue;
      }
    } else {
      continue;
    }

    DWORD type = REG_NONE;
    std::vector<BYTE> data;
    if (data_part.front() == L'"') {
      std::wstring text;
      size_t end_pos = 0;
      if (!ParseQuotedString(data_part, &text, &end_pos)) {
        continue;
      }
      type = REG_SZ;
      data = StringToRegData(text);
    } else if (StartsWithInsensitive(data_part, L"dword:")) {
      std::wstring hex = TrimWhitespace(data_part.substr(6));
      if (hex.empty()) {
        continue;
      }
      DWORD number = static_cast<DWORD>(wcstoul(hex.c_str(), nullptr, 16));
      type = REG_DWORD;
      data.resize(sizeof(DWORD));
      memcpy(data.data(), &number, sizeof(DWORD));
    } else if (StartsWithInsensitive(data_part, L"hex")) {
      type = REG_BINARY;
      size_t colon = data_part.find(L':');
      if (colon == std::wstring::npos) {
        continue;
      }
      size_t open = data_part.find(L'(');
      size_t close = data_part.find(L')');
      if (open != std::wstring::npos && close != std::wstring::npos && close > open) {
        std::wstring code = data_part.substr(open + 1, close - open - 1);
        unsigned long parsed = wcstoul(code.c_str(), nullptr, 16);
        switch (parsed) {
        case 0x0:
          type = REG_NONE;
          break;
        case 0x1:
          type = REG_SZ;
          break;
        case 0x2:
          type = REG_EXPAND_SZ;
          break;
        case 0x3:
          type = REG_BINARY;
          break;
        case 0x4:
          type = REG_DWORD;
          break;
        case 0x5:
          type = REG_DWORD_BIG_ENDIAN;
          break;
        case 0x7:
          type = REG_MULTI_SZ;
          break;
        case 0x8:
          type = REG_RESOURCE_LIST;
          break;
        case 0x9:
          type = REG_FULL_RESOURCE_DESCRIPTOR;
          break;
        case 0xA:
          type = REG_RESOURCE_REQUIREMENTS_LIST;
          break;
        case 0xB:
          type = REG_QWORD;
          break;
        default:
          type = REG_BINARY;
          break;
        }
      }
      std::wstring hex = data_part.substr(colon + 1);
      if (!ParseHexBytes(hex, &data)) {
        continue;
      }
    } else {
      continue;
    }

    RegistryProvider::VirtualRegistryValue value;
    value.name = value_name;
    value.type = type;
    value.data = std::move(data);
    std::wstring name_lower = ToLower(value_name);
    current_key->values[name_lower] = std::move(value);
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

std::wstring KeyLeafFromPath(const std::wstring& path) {
  if (path.empty()) {
    return {};
  }
  size_t pos = path.find_last_of(L"\\/");
  if (pos == std::wstring::npos) {
    return path;
  }
  return path.substr(pos + 1);
}

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif

using NtOpenKeyFn = NTSTATUS(NTAPI*)(PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES);

NtOpenKeyFn LoadNtOpenKey() {
  HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
  if (!ntdll) {
    return nullptr;
  }
  return reinterpret_cast<NtOpenKeyFn>(GetProcAddress(ntdll, "NtOpenKey"));
}

NtOpenKeyFn GetNtOpenKey() {
  static NtOpenKeyFn open_fn = LoadNtOpenKey();
  return open_fn;
}

util::UniqueHKey OpenRegistryRootKey() {
  NtOpenKeyFn open_fn = GetNtOpenKey();
  if (!open_fn) {
    return util::UniqueHKey();
  }
  std::wstring path = L"\\REGISTRY";
  UNICODE_STRING name = {};
  name.Buffer = const_cast<PWSTR>(path.c_str());
  name.Length = static_cast<USHORT>(path.size() * sizeof(wchar_t));
  name.MaximumLength = name.Length;
  OBJECT_ATTRIBUTES attrs = {};
  InitializeObjectAttributes(&attrs, &name, OBJ_CASE_INSENSITIVE, nullptr, nullptr);

  HANDLE handle = nullptr;
  NTSTATUS status = open_fn(&handle, KEY_READ, &attrs);
  if (!NT_SUCCESS(status) || !handle) {
    return util::UniqueHKey();
  }
  return util::UniqueHKey(reinterpret_cast<HKEY>(handle));
}

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
  wchar_t expanded[4096] = {};
  DWORD expanded_len = ExpandEnvironmentStringsW(path.c_str(), expanded, static_cast<DWORD>(_countof(expanded)));
  if (expanded_len > 0 && expanded_len < _countof(expanded)) {
    path.assign(expanded, expanded_len - 1);
  }
  path = ResolveDevicePath(path);
  return path;
}

std::wstring StripOuterQuotes(const std::wstring& text) {
  if (text.size() < 2) {
    return text;
  }
  if ((text.front() == L'"' && text.back() == L'"') || (text.front() == L'\'' && text.back() == L'\'')) {
    return text.substr(1, text.size() - 2);
  }
  return text;
}

std::wstring StripRegFileKeySyntax(const std::wstring& text) {
  std::wstring trimmed = TrimWhitespace(text);
  if (trimmed.empty()) {
    return trimmed;
  }
  if (trimmed.front() == L'[' && trimmed.back() == L']' && trimmed.size() >= 2) {
    std::wstring inner = trimmed.substr(1, trimmed.size() - 2);
    inner = TrimWhitespace(inner);
    if (!inner.empty() && inner.front() == L'-') {
      inner.erase(inner.begin());
      inner = TrimWhitespace(inner);
    }
    return inner;
  }
  if (!trimmed.empty() && trimmed.front() == L'-') {
    trimmed.erase(trimmed.begin());
    trimmed = TrimWhitespace(trimmed);
  }
  return trimmed;
}

std::wstring CollapseBackslashes(const std::wstring& text) {
  if (text.empty()) {
    return text;
  }
  std::wstring out;
  out.reserve(text.size());
  bool last_slash = false;
  for (wchar_t ch : text) {
    if (ch == L'\\') {
      if (!last_slash) {
        out.push_back(ch);
      }
      last_slash = true;
    } else {
      last_slash = false;
      out.push_back(ch);
    }
  }
  return out;
}

std::wstring EscapeBackslashes(const std::wstring& text) {
  if (text.empty()) {
    return text;
  }
  std::wstring out;
  out.reserve(text.size() * 2);
  for (wchar_t ch : text) {
    if (ch == L'\\') {
      out.push_back(L'\\');
    }
    out.push_back(ch);
  }
  return out;
}

std::wstring MapNativeRegistryPath(const std::wstring& path, const std::wstring& sid) {
  if (!StartsWithInsensitive(path, L"REGISTRY")) {
    return L"";
  }
  std::wstring rest = path.substr(wcslen(L"REGISTRY"));
  while (!rest.empty() && rest.front() == L'\\') {
    rest.erase(rest.begin());
  }
  if (rest.empty()) {
    return L"REGISTRY";
  }

  auto strip_prefix = [&](const wchar_t* prefix) -> std::wstring {
    size_t len = wcslen(prefix);
    if (!StartsWithInsensitive(rest, prefix)) {
      return L"";
    }
    std::wstring tail = rest.substr(len);
    while (!tail.empty() && tail.front() == L'\\') {
      tail.erase(tail.begin());
    }
    return tail;
  };

  std::wstring tail = strip_prefix(L"MACHINE");
  if (!tail.empty() || StartsWithInsensitive(rest, L"MACHINE")) {
    std::wstring machine_tail = tail;
    const wchar_t* classes_prefix = L"SOFTWARE\\Classes";
    if (StartsWithInsensitive(machine_tail, classes_prefix)) {
      std::wstring classes_tail = machine_tail.substr(wcslen(classes_prefix));
      while (!classes_tail.empty() && classes_tail.front() == L'\\') {
        classes_tail.erase(classes_tail.begin());
      }
      return classes_tail.empty() ? L"HKEY_CLASSES_ROOT" : L"HKEY_CLASSES_ROOT\\" + classes_tail;
    }
    const wchar_t* cc_prefix = L"SYSTEM\\CurrentControlSet\\Hardware Profiles\\Current";
    if (StartsWithInsensitive(machine_tail, cc_prefix)) {
      std::wstring cc_tail = machine_tail.substr(wcslen(cc_prefix));
      while (!cc_tail.empty() && cc_tail.front() == L'\\') {
        cc_tail.erase(cc_tail.begin());
      }
      return cc_tail.empty() ? L"HKEY_CURRENT_CONFIG" : L"HKEY_CURRENT_CONFIG\\" + cc_tail;
    }
    return machine_tail.empty() ? L"HKEY_LOCAL_MACHINE" : L"HKEY_LOCAL_MACHINE\\" + machine_tail;
  }

  tail = strip_prefix(L"USER");
  if (!tail.empty() || StartsWithInsensitive(rest, L"USER")) {
    if (!sid.empty() && StartsWithInsensitive(tail, sid)) {
      std::wstring user_tail = tail.substr(sid.size());
      while (!user_tail.empty() && user_tail.front() == L'\\') {
        user_tail.erase(user_tail.begin());
      }
      return user_tail.empty() ? L"HKEY_CURRENT_USER" : L"HKEY_CURRENT_USER\\" + user_tail;
    }
    return tail.empty() ? L"HKEY_USERS" : L"HKEY_USERS\\" + tail;
  }

  return L"REGISTRY\\" + rest;
}

std::wstring JoinPathParts(const std::vector<std::wstring>& parts) {
  std::wstring out;
  for (const auto& part : parts) {
    if (part.empty()) {
      continue;
    }
    if (!out.empty()) {
      out.append(L"\\");
    }
    out.append(part);
  }
  return out;
}

std::wstring JoinPathPartsRange(const std::vector<std::wstring>& parts, size_t start) {
  std::wstring out;
  for (size_t i = start; i < parts.size(); ++i) {
    if (parts[i].empty()) {
      continue;
    }
    if (!out.empty()) {
      out.append(L"\\");
    }
    out.append(parts[i]);
  }
  return out;
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
  std::vector<std::wstring> parts = SplitPath(path);
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
      return JoinPathParts(parts);
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
  std::vector<std::wstring> parts = SplitPath(path);
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
      return JoinPathParts(parts);
    }
  }
  return L"";
}

std::wstring CleanTraceKeyText(const std::wstring& text, const std::wstring& sid) {
  std::wstring path = StripRegFileKeySyntax(text);
  path = StripOuterQuotes(path);
  path = TrimWhitespace(path);
  if (path.empty()) {
    return L"";
  }
  for (auto& ch : path) {
    if (ch == L'/') {
      ch = L'\\';
    }
  }
  path = CollapseBackslashes(path);
  if (StartsWithInsensitive(path, L"Registry::")) {
    path.erase(0, wcslen(L"Registry::"));
  }
  while (!path.empty() && path.front() == L'\\') {
    path.erase(path.begin());
  }
  if (StartsWithInsensitive(path, L"Computer\\")) {
    path.erase(0, wcslen(L"Computer\\"));
  }
  wchar_t machine[MAX_COMPUTERNAME_LENGTH + 1] = {};
  DWORD machine_len = static_cast<DWORD>(_countof(machine));
  if (GetComputerNameW(machine, &machine_len) && machine_len > 0) {
    std::wstring prefix = std::wstring(machine, machine_len) + L"\\";
    if (StartsWithInsensitive(path, prefix)) {
      path.erase(0, prefix.size());
    }
  }
  if (!sid.empty()) {
    const std::wstring marker = L"<CURRENT_USER_SID>";
    size_t pos = path.find(marker);
    while (pos != std::wstring::npos) {
      path.replace(pos, marker.size(), sid);
      pos = path.find(marker, pos + sid.size());
    }
  }
  return path;
}

std::wstring NormalizeTraceKeyPathBasic(const std::wstring& text) {
  std::wstring sid = util::GetCurrentUserSidString();
  std::wstring path = CleanTraceKeyText(text, sid);
  if (path.empty()) {
    return L"";
  }

  std::wstring native_mapped = MapNativeRegistryPath(path, sid);
  if (!native_mapped.empty()) {
    path = native_mapped;
  }

  auto map_root = [&](const std::wstring& root, const std::wstring& rest) -> std::wstring {
    std::wstring mapped;
    if (EqualsInsensitive(root, L"HKLM") || EqualsInsensitive(root, L"HKEY_LOCAL_MACHINE")) {
      mapped = L"HKEY_LOCAL_MACHINE";
    } else if (EqualsInsensitive(root, L"HKCU") || EqualsInsensitive(root, L"HKEY_CURRENT_USER")) {
      mapped = L"HKEY_CURRENT_USER";
    } else if (EqualsInsensitive(root, L"HKCR") || EqualsInsensitive(root, L"HKEY_CLASSES_ROOT")) {
      mapped = L"HKEY_CLASSES_ROOT";
    } else if (EqualsInsensitive(root, L"HKU") || EqualsInsensitive(root, L"HKEY_USERS")) {
      if (!sid.empty() && StartsWithInsensitive(rest, sid)) {
        std::wstring tail = rest.substr(sid.size());
        if (!tail.empty() && tail.front() == L'\\') {
          tail.erase(tail.begin());
        }
        mapped = L"HKEY_CURRENT_USER";
        if (!tail.empty()) {
          mapped += L"\\" + tail;
        }
        return mapped;
      }
      mapped = L"HKEY_USERS";
    } else if (EqualsInsensitive(root, L"HKCC") || EqualsInsensitive(root, L"HKEY_CURRENT_CONFIG")) {
      mapped = L"HKEY_CURRENT_CONFIG";
    } else if (EqualsInsensitive(root, L"Machine")) {
      mapped = L"HKEY_LOCAL_MACHINE";
    } else if (EqualsInsensitive(root, L"User") || EqualsInsensitive(root, L"Users")) {
      if (!sid.empty() && StartsWithInsensitive(rest, sid)) {
        std::wstring tail = rest.substr(sid.size());
        if (!tail.empty() && tail.front() == L'\\') {
          tail.erase(tail.begin());
        }
        mapped = L"HKEY_CURRENT_USER";
        if (!tail.empty()) {
          mapped += L"\\" + tail;
        }
        return mapped;
      }
      mapped = L"HKEY_USERS";
    } else {
      return L"";
    }
    if (rest.empty()) {
      return mapped;
    }
    return mapped + L"\\" + rest;
  };

  std::wstring without_prefix = path;
  if (StartsWithInsensitive(without_prefix, L"Registry\\")) {
    without_prefix.erase(0, wcslen(L"Registry\\"));
  }
  size_t slash = without_prefix.find(L'\\');
  std::wstring root = (slash == std::wstring::npos) ? without_prefix : without_prefix.substr(0, slash);
  std::wstring rest = (slash == std::wstring::npos) ? L"" : without_prefix.substr(slash + 1);
  std::wstring mapped = map_root(root, rest);
  if (!mapped.empty()) {
    return NormalizeCurrentControlSet(mapped);
  }

  slash = path.find(L'\\');
  root = (slash == std::wstring::npos) ? path : path.substr(0, slash);
  rest = (slash == std::wstring::npos) ? L"" : path.substr(slash + 1);
  mapped = map_root(root, rest);
  if (!mapped.empty()) {
    return NormalizeCurrentControlSet(mapped);
  }
  if (StartsWithInsensitive(path, L"REGISTRY")) {
    std::wstring tail = path.substr(wcslen(L"REGISTRY"));
    while (!tail.empty() && tail.front() == L'\\') {
      tail.erase(tail.begin());
    }
    return tail.empty() ? L"REGISTRY" : L"REGISTRY\\" + tail;
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

bool SelectionIncludesKey(const KeyValueSelection& selection, const std::wstring& key_lower) {
  if (selection.select_all) {
    return true;
  }
  if (key_lower.empty()) {
    return true;
  }
  if (!selection.key_paths.empty()) {
    for (const auto& path : selection.key_paths) {
      std::wstring selection_lower = ToLower(path);
      if (selection_lower.empty()) {
        continue;
      }
      if (key_lower == selection_lower) {
        return true;
      }
      if (key_lower.size() < selection_lower.size() && selection_lower.compare(0, key_lower.size(), key_lower) == 0 && selection_lower[key_lower.size()] == L'\\') {
        return true;
      }
      if (selection.recursive && key_lower.size() > selection_lower.size() && key_lower.compare(0, selection_lower.size(), selection_lower) == 0 && key_lower[selection_lower.size()] == L'\\') {
        return true;
      }
    }
    return false;
  }
  if (!selection.values_by_key.empty()) {
    return selection.values_by_key.find(key_lower) != selection.values_by_key.end();
  }
  return true;
}

bool SelectionIncludesValue(const KeyValueSelection& selection, const std::wstring& key_lower, const std::wstring& value_lower) {
  if (selection.select_all) {
    return true;
  }
  auto it = selection.values_by_key.find(key_lower);
  if (it == selection.values_by_key.end() || it->second.empty()) {
    return true;
  }
  return it->second.find(value_lower) != it->second.end();
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
  std::wstring path = input;
  while (!path.empty() && (path.front() == L'\\' || path.front() == L'/')) {
    path.erase(path.begin());
  }
  if (path.empty()) {
    return false;
  }
  size_t slash = path.find_first_of(L"\\/");
  std::wstring root = (slash == std::wstring::npos) ? path : path.substr(0, slash);
  std::wstring rest = (slash == std::wstring::npos) ? L"" : path.substr(slash + 1);
  auto set_root = [&](const wchar_t* label, HKEY root_key) -> bool {
    *root_label = label;
    node->root = root_key;
    node->root_name = label;
    node->subkey = rest;
    return true;
  };
  if (EqualsInsensitive(root, L"REGISTRY")) {
    *root_label = L"REGISTRY";
    node->root = nullptr;
    node->root_name = L"REGISTRY";
    node->subkey = rest;
    return true;
  }
  if (EqualsInsensitive(root, L"HKLM") || EqualsInsensitive(root, L"HKEY_LOCAL_MACHINE")) {
    return set_root(L"HKEY_LOCAL_MACHINE", HKEY_LOCAL_MACHINE);
  }
  if (EqualsInsensitive(root, L"HKCU") || EqualsInsensitive(root, L"HKEY_CURRENT_USER")) {
    return set_root(L"HKEY_CURRENT_USER", HKEY_CURRENT_USER);
  }
  if (EqualsInsensitive(root, L"HKCR") || EqualsInsensitive(root, L"HKEY_CLASSES_ROOT")) {
    return set_root(L"HKEY_CLASSES_ROOT", HKEY_CLASSES_ROOT);
  }
  if (EqualsInsensitive(root, L"HKU") || EqualsInsensitive(root, L"HKEY_USERS")) {
    return set_root(L"HKEY_USERS", HKEY_USERS);
  }
  if (EqualsInsensitive(root, L"HKCC") || EqualsInsensitive(root, L"HKEY_CURRENT_CONFIG")) {
    return set_root(L"HKEY_CURRENT_CONFIG", HKEY_CURRENT_CONFIG);
  }
  return false;
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
    std::vector<std::wstring> parts = SplitPath(root_node.subkey);
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
      std::wstring remaining = JoinPathPartsRange(parts, i + 1);
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
    std::wstring root_name = RegistryProvider::RootName(current_node->root);
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
    return node.root_name.empty() ? RegistryProvider::RootName(node.root) : node.root_name;
  }
  size_t pos = node.subkey.rfind(L'\\');
  if (pos == std::wstring::npos) {
    return node.subkey;
  }
  return node.subkey.substr(pos + 1);
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

std::wstring EscapeHistoryField(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (wchar_t ch : text) {
    switch (ch) {
    case L'\\':
      out.append(L"\\\\");
      break;
    case L'\t':
      out.append(L"\\t");
      break;
    case L'\r':
      out.append(L"\\r");
      break;
    case L'\n':
      out.append(L"\\n");
      break;
    default:
      out.push_back(ch);
      break;
    }
  }
  return out;
}

std::wstring UnescapeHistoryField(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (size_t i = 0; i < text.size(); ++i) {
    wchar_t ch = text[i];
    if (ch == L'\\' && i + 1 < text.size()) {
      wchar_t next = text[i + 1];
      switch (next) {
      case L'\\':
        out.push_back(L'\\');
        ++i;
        continue;
      case L't':
        out.push_back(L'\t');
        ++i;
        continue;
      case L'r':
        out.push_back(L'\r');
        ++i;
        continue;
      case L'n':
        out.push_back(L'\n');
        ++i;
        continue;
      default:
        break;
      }
    }
    out.push_back(ch);
  }
  return out;
}

std::vector<std::wstring> SplitHistoryFields(const std::wstring& line) {
  std::vector<std::wstring> fields;
  std::wstring current;
  for (wchar_t ch : line) {
    if (ch == L'\t') {
      fields.push_back(current);
      current.clear();
    } else {
      current.push_back(ch);
    }
  }
  fields.push_back(current);
  return fields;
}

std::wstring MakeValueCommentKey(const std::wstring& path, const std::wstring& name, DWORD type) {
  std::wstring key = ToLower(path);
  key.push_back(L'\t');
  key.append(ToLower(name));
  key.push_back(L'\t');
  key.append(std::to_wstring(type));
  return key;
}

std::wstring MakeNameCommentKey(const std::wstring& name, DWORD type) {
  std::wstring key = ToLower(name);
  key.push_back(L'\t');
  key.append(std::to_wstring(type));
  return key;
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

std::vector<std::wstring> MultiSzToVector(const std::vector<BYTE>& data) {
  std::vector<std::wstring> items;
  if (data.empty()) {
    return items;
  }
  const wchar_t* ptr = reinterpret_cast<const wchar_t*>(data.data());
  size_t count = data.size() / sizeof(wchar_t);
  size_t offset = 0;
  while (offset < count) {
    const wchar_t* current = ptr + offset;
    size_t len = wcsnlen_s(current, count - offset);
    if (len == 0) {
      break;
    }
    items.emplace_back(current, len);
    offset += len + 1;
  }
  return items;
}

std::vector<BYTE> VectorToMultiSz(const std::vector<std::wstring>& items) {
  size_t total_chars = 1;
  for (const auto& item : items) {
    total_chars += item.size() + 1;
  }
  std::vector<BYTE> data(total_chars * sizeof(wchar_t));
  wchar_t* out = reinterpret_cast<wchar_t*>(data.data());
  size_t offset = 0;
  for (const auto& item : items) {
    wcsncpy_s(out + offset, total_chars - offset, item.c_str(), item.size());
    offset += item.size();
    out[offset++] = L'\0';
  }
  out[offset++] = L'\0';
  data.resize(offset * sizeof(wchar_t));
  return data;
}

class ReplaceMatcher {
public:
  explicit ReplaceMatcher(const ReplaceDialogResult& options, bool* ok) : query_(options.find_text), replacement_(options.replace_text), use_regex_(options.use_regex), match_case_(options.match_case), match_whole_(options.match_whole) {
    if (query_.empty()) {
      if (ok) {
        *ok = false;
      }
      return;
    }
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

  bool Replace(const std::wstring& text, std::wstring* out) const {
    if (!out || query_.empty()) {
      return false;
    }
    if (use_regex_) {
      if (match_whole_) {
        if (!std::regex_match(text, regex_)) {
          return false;
        }
      } else {
        if (!std::regex_search(text, regex_)) {
          return false;
        }
      }
      *out = std::regex_replace(text, regex_, replacement_);
      return true;
    }

    if (match_whole_) {
      bool match = false;
      if (match_case_) {
        match = (text == query_);
      } else {
        match = CompareStringOrdinal(text.c_str(), static_cast<int>(text.size()), query_.c_str(), static_cast<int>(query_.size()), TRUE) == CSTR_EQUAL;
      }
      if (!match) {
        return false;
      }
      *out = replacement_;
      return true;
    }

    if (match_case_) {
      size_t pos = text.find(query_);
      if (pos == std::wstring::npos) {
        return false;
      }
      std::wstring result;
      size_t cursor = 0;
      while (pos != std::wstring::npos) {
        result.append(text, cursor, pos - cursor);
        result.append(replacement_);
        cursor = pos + query_.size();
        pos = text.find(query_, cursor);
      }
      result.append(text, cursor, std::wstring::npos);
      *out = std::move(result);
      return true;
    }

    size_t cursor = 0;
    std::wstring result;
    bool matched = false;
    while (cursor < text.size()) {
      int pos = FindStringOrdinal(FIND_FROMSTART, text.c_str() + cursor, static_cast<int>(text.size() - cursor), query_.c_str(), static_cast<int>(query_.size()), TRUE);
      if (pos < 0) {
        break;
      }
      size_t match_pos = cursor + static_cast<size_t>(pos);
      result.append(text, cursor, match_pos - cursor);
      result.append(replacement_);
      cursor = match_pos + query_.size();
      matched = true;
    }
    if (!matched) {
      return false;
    }
    result.append(text, cursor, std::wstring::npos);
    *out = std::move(result);
    return true;
  }

private:
  std::wstring query_;
  std::wstring replacement_;
  bool use_regex_ = false;
  bool match_case_ = false;
  bool match_whole_ = false;
  std::wregex regex_;
};

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

int CompareHistoryEntry(const HistoryEntry& left, const HistoryEntry& right, int column) {
  switch (column) {
  case 0:
    return CompareUint64(left.timestamp, right.timestamp);
  case 1:
    return CompareTextInsensitive(left.action, right.action);
  case 2:
    return CompareTextInsensitive(left.old_data, right.old_data);
  case 3:
    return CompareTextInsensitive(left.new_data, right.new_data);
  default:
    return CompareUint64(left.timestamp, right.timestamp);
  }
}

void SortHistoryEntries(std::vector<HistoryEntry>* entries, int column, bool ascending) {
  if (!entries || entries->size() < 2) {
    return;
  }
  std::stable_sort(entries->begin(), entries->end(), [column, ascending](const HistoryEntry& left, const HistoryEntry& right) {
    int result = CompareHistoryEntry(left, right, column);
    if (result == 0) {
      return false;
    }
    return ascending ? (result < 0) : (result > 0);
  });
}

int CompareSearchResult(const SearchResult& left, const SearchResult& right, int column, bool compare) {
  if (compare) {
    switch (column) {
    case 0:
      return CompareTextInsensitive(left.key_path, right.key_path);
    case 1:
      return CompareTextInsensitive(left.display_name, right.display_name);
    case 2:
      return CompareTextInsensitive(left.type_text, right.type_text);
    case 3:
      return CompareTextInsensitive(left.data, right.data);
    default:
      return CompareTextInsensitive(left.key_path, right.key_path);
    }
  }
  switch (column) {
  case 0:
    return CompareTextInsensitive(left.key_path, right.key_path);
  case 1:
    return CompareTextInsensitive(left.display_name, right.display_name);
  case 2:
    return CompareTextInsensitive(left.type_text, right.type_text);
  case 3:
    return CompareTextInsensitive(left.data, right.data);
  case 4:
    return CompareTextInsensitive(left.size_text, right.size_text);
  case 5:
    return CompareTextInsensitive(left.date_text, right.date_text);
  default:
    return CompareTextInsensitive(left.key_path, right.key_path);
  }
}

void SortSearchResultEntries(std::vector<SearchResult>* entries, int column, bool ascending, bool compare) {
  if (!entries || entries->size() < 2) {
    return;
  }
  std::stable_sort(entries->begin(), entries->end(), [column, ascending, compare](const SearchResult& left, const SearchResult& right) {
    int result = CompareSearchResult(left, right, column, compare);
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

  HRESULT STDMETHODCALLTYPE Expand(PCWSTR text) override {
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
    GetWindowTextW(edit_, text.data(), length + 1);
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
  case WM_ERASEBKGND:
    return TRUE;
  case WM_PAINT: {
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd, &ps);
    const Theme& theme = Theme::Current();
    RECT client = {};
    GetClientRect(hwnd, &client);
    FillRect(hdc, &client, theme.SurfaceBrush());

    int count = static_cast<int>(SendMessageW(hwnd, LB_GETCOUNT, 0, 0));
    int selected = static_cast<int>(SendMessageW(hwnd, LB_GETCURSEL, 0, 0));
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    HFONT old_font = font ? reinterpret_cast<HFONT>(SelectObject(hdc, font)) : nullptr;
    SetBkMode(hdc, TRANSPARENT);

    for (int i = 0; i < count; ++i) {
      RECT item_rect = {};
      if (SendMessageW(hwnd, LB_GETITEMRECT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(&item_rect)) == LB_ERR) {
        continue;
      }
      bool is_selected = (i == selected);
      COLORREF bg = is_selected ? theme.SelectionColor() : theme.SurfaceColor();
      COLORREF text = is_selected ? theme.SelectionTextColor() : theme.TextColor();
      FillRect(hdc, &item_rect, GetCachedBrush(bg));
      SetTextColor(hdc, text);

      int len = static_cast<int>(SendMessageW(hwnd, LB_GETTEXTLEN, i, 0));
      if (len > 0 && len < 8192) {
        std::wstring item_text(static_cast<size_t>(len) + 1, L'\0');
        SendMessageW(hwnd, LB_GETTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(item_text.data()));
        item_text.resize(static_cast<size_t>(len));
        RECT text_rect = item_rect;
        text_rect.left += 6;
        text_rect.right -= 6;
        DrawTextW(hdc, item_text.c_str(), -1, &text_rect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
      }
    }

    if (old_font) {
      SelectObject(hdc, old_font);
    }
    EndPaint(hwnd, &ps);
    return 0;
  }
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
        COLORREF text = theme.TextColor();
        COLORREF background = theme.SurfaceColor();
        if (draw->nmcd.uItemState & CDIS_SELECTED) {
          text = theme.SelectionTextColor();
          background = theme.SelectionColor();
        } else if (draw->nmcd.uItemState & CDIS_HOT) {
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
  wc.lpszClassName = L"RegKitMainWindow";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
  wc.hIconSm = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
  wc.hbrBackground = nullptr;

  RegisterClassExW(&wc);

  std::wstring title = L"RegKit V";
  title.append(REGKIT_VERSION_STR_W);
  if (IsProcessTrustedInstaller()) {
    title.append(L" - [TrustedInstaller]");
  } else if (IsProcessSystem()) {
    title.append(L" - [SYSTEM]");
  } else if (IsProcessElevated()) {
    title.append(L" - [Administrator]");
  }
  hwnd_ = CreateWindowExW(0, wc.lpszClassName, title.c_str(), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, CW_USEDEFAULT, CW_USEDEFAULT, 1200, 800, nullptr, nullptr, instance, this);
  if (hwnd_) {
    DragAcceptFiles(hwnd_, TRUE);
  }
  return hwnd_ != nullptr;
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
  PostMessageW(hwnd_, kLoadTracesMessage, 0, 0);
  PostMessageW(hwnd_, kLoadDefaultsMessage, 0, 0);
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
  if (message == WM_KEYDOWN && wparam == VK_RETURN) {
    auto* self = reinterpret_cast<MainWindow*>(ref_data);
    if (self && self->hwnd_) {
      SendMessageW(self->hwnd_, kAddressEnterMessage, 0, 0);
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
    auto* self = reinterpret_cast<MainWindow*>(ref_data);
    if (self) {
      self->ApplyAutoCompleteTheme();
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
  case WM_DROPFILES: {
    HDROP drop = reinterpret_cast<HDROP>(wparam);
    UINT count = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
    std::vector<std::wstring> reg_paths;
    std::wstring offline_candidate;
    for (UINT index = 0; index < count; ++index) {
      wchar_t buffer[MAX_PATH] = {};
      if (DragQueryFileW(drop, index, buffer, static_cast<UINT>(_countof(buffer))) == 0) {
        continue;
      }
      std::wstring path = buffer;
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
    if (generation != search_generation_) {
      return 0;
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
          last_search_results_.push_back(search_tabs_[static_cast<size_t>(index)].results.back());
          ++processed;
          if (processed >= kSearchResultsBatch || (GetTickCount64() - start_tick) >= kSearchResultsMaxMs) {
            stop_at = i + 1;
            break;
          }
        }
        auto& tab = search_tabs_[static_cast<size_t>(index)];
        if (processed > 0 && tab.sort_column >= 0) {
          SortSearchResultEntries(&tab.results, tab.sort_column, tab.sort_ascending, tab.is_compare);
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
    search_progress_posted_.store(false);
    if (generation != search_generation_) {
      return 0;
    }
    UpdateStatus();
    return 0;
  }
  case kSearchFinishedMessage: {
    uint64_t generation = static_cast<uint64_t>(wparam);
    if (generation != search_generation_) {
      return 0;
    }
    search_running_ = false;
    if (!search_cancel_.load() && search_start_tick_ != 0) {
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
    if (generation != search_generation_ || search_cancel_.load()) {
      return 0;
    }
    search_running_ = false;
    search_duration_ms_ = 0;
    search_duration_valid_ = false;
    ui::ShowError(hwnd_, L"Invalid regex.");
    ApplyViewVisibility();
    UpdateStatus();
    return 0;
  }
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
    active_defaults_ = std::move(owned->defaults);
    BuildMenus();
    UpdateValueListForNode(current_node_);
    return 0;
  }
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
    if (session_it->second && session_it->second->thread.joinable()) {
      session_it->second->thread.join();
    }
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
    bool touches_current = false;
    std::wstring current_key_lower;
    if (current_node_) {
      current_key_lower = TracePathLowerForNode(*current_node_);
    }
    if (session->data && !owned->entries.empty()) {
      std::unique_lock<std::shared_mutex> data_lock(*session->data->mutex);
      for (const auto& entry : owned->entries) {
        if (entry.key_path.empty()) {
          continue;
        }
        std::wstring key_lower = ToLower(entry.key_path);
        if (!current_key_lower.empty() && key_lower == current_key_lower) {
          touches_current = true;
        }
        auto map_it = session->data->values_by_key.find(key_lower);
        if (map_it == session->data->values_by_key.end()) {
          map_it = session->data->values_by_key.emplace(key_lower, TraceKeyValues()).first;
          session->data->key_paths.push_back(entry.key_path);
          std::vector<std::wstring> parts = SplitPath(entry.key_path);
          if (parts.size() > 1) {
            std::wstring current = parts.front();
            for (size_t i = 1; i < parts.size(); ++i) {
              std::wstring parent_lower = ToLower(current);
              session->data->children_by_key[parent_lower].push_back(parts[i]);
              current.append(L"\\");
              current.append(parts[i]);
            }
          }
        }
        if (!entry.display_path.empty()) {
          std::wstring display_lower = ToLower(entry.display_path);
          if (session->data->display_to_key.find(display_lower) == session->data->display_to_key.end()) {
            session->data->display_to_key.emplace(display_lower, entry.key_path);
            session->data->display_key_paths.push_back(entry.display_path);
          }
        }
        if (entry.has_value) {
          std::wstring value_lower = ToLower(entry.value_name);
          if (map_it->second.values_lower.insert(value_lower).second) {
            map_it->second.values_display.push_back(entry.value_name);
          }
        }
      }
    }
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
        std::sort(session->data->key_paths.begin(), session->data->key_paths.end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
        std::sort(session->data->display_key_paths.begin(), session->data->display_key_paths.end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
        TraceSelection normalized = session->selection;
        NormalizeSelectionForTrace(*session->data, &normalized);
        session->selection = normalized;
        trace_selection_cache_[session->source_lower] = normalized;
        if (session->added_to_active) {
          for (auto& trace : active_traces_) {
            if (EqualsInsensitive(trace.source_path, session->source_path)) {
              trace.selection = normalized;
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
      if (session->thread.joinable()) {
        session->thread.join();
      }
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
    bool touches_current = false;
    std::wstring current_key_lower;
    if (current_node_) {
      std::wstring path = RegistryProvider::BuildPath(*current_node_);
      std::wstring normalized = NormalizeTraceKeyPathBasic(path);
      if (normalized.empty()) {
        normalized = path;
      }
      current_key_lower = ToLower(normalized);
    }
    if (session->data && !owned->entries.empty()) {
      std::unique_lock<std::shared_mutex> data_lock(*session->data->mutex);
      for (const auto& entry : owned->entries) {
        if (entry.key_path.empty()) {
          continue;
        }
        std::wstring key_lower = ToLower(entry.key_path);
        if (!current_key_lower.empty() && key_lower == current_key_lower) {
          touches_current = true;
        }
        auto& key_values = session->data->values_by_key[key_lower];
        if (entry.has_value) {
          std::wstring value_lower = ToLower(entry.value_name);
          DefaultValueEntry value_entry;
          value_entry.type = entry.value_type;
          value_entry.data = entry.value_data;
          key_values.values[value_lower] = value_entry;
          std::wstring alias_path = MapControlSetToCurrent(entry.key_path);
          if (!alias_path.empty()) {
            std::wstring alias_lower = ToLower(alias_path);
            auto& alias_values = session->data->values_by_key[alias_lower];
            alias_values.values[value_lower] = value_entry;
            if (!current_key_lower.empty() && alias_lower == current_key_lower) {
              touches_current = true;
            }
          }
        }
      }
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
      if (session->thread.joinable()) {
        session->thread.join();
      }
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
    current_key_count_ = payload->key_count;
    current_value_count_ = payload->value_count;
    if (list_hwnd) {
      SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
      InvalidateRect(list_hwnd, nullptr, TRUE);
    }
    value_list_loading_ = false;
    UpdateStatus();
    StartPendingValueListRename();
    return 0;
  }
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
      value_list_.SetFilter(buffer);
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
        FillRect(draw->nmcd.hdc, &draw->nmcd.rc, theme.PanelBrush());
        return CDRF_NOTIFYITEMDRAW;
      case CDDS_ITEMPREPAINT: {
        draw->hbrMonoDither = theme.PanelBrush();
        draw->hbrLines = theme.PanelBrush();
        draw->hpenLines = GetCachedPen(theme.BorderColor(), 1);
        draw->clrText = theme.TextColor();
        draw->clrTextHighlight = theme.TextColor();
        draw->clrBtnFace = theme.PanelColor();
        draw->clrBtnHighlight = theme.SurfaceColor();
        draw->clrHighlightHotTrack = theme.HoverColor();
        draw->nStringBkMode = TRANSPARENT;
        draw->nHLStringBkMode = TRANSPARENT;

        if ((draw->nmcd.uItemState & CDIS_HOT) == CDIS_HOT) {
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
        MarkTreeStateDirty();
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
        UndoOperation op;
        op.type = UndoOperation::Type::kRenameKey;
        op.node = parent;
        op.name = old_name;
        op.new_name = new_name;
        PushUndo(std::move(op));
        RefreshTreeSelection();
        UpdateValueListForNode(current_node_);
        return TRUE;
      }
      if (header->code == TVN_SELCHANGEDW) {
        RegistryNode* node = tree_.OnSelectionChanged(reinterpret_cast<NMTREEVIEWW*>(lparam));
        current_node_ = node;
        UpdateAddressBar(node);
        UpdateValueListForNode(node);
        MarkTreeStateDirty();
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
          const Theme& theme = Theme::Current();
          bool selected = (draw->nmcd.uItemState & CDIS_SELECTED) != 0;
          bool hot = (draw->nmcd.uItemState & CDIS_HOT) != 0;
          if (selected) {
            draw->clrText = theme.SelectionTextColor();
            draw->clrTextBk = theme.SelectionColor();
          } else if (hot) {
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
        UndoOperation op;
        op.type = UndoOperation::Type::kRenameKey;
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
      UndoOperation op;
      op.type = UndoOperation::Type::kRenameValue;
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
      const SearchResult* result = nullptr;
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
      if (disp->item.mask & LVIF_TEXT) {
        const wchar_t* text = L"";
        bool compare = false;
        if (index >= 0 && static_cast<size_t>(index) < search_tabs_.size()) {
          compare = search_tabs_[static_cast<size_t>(index)].is_compare;
        }
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
            std::wstring path = RegistryProvider::BuildPath(*current_node_);
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
            EditValueComment(*row);
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
            const SearchResult& result = search_tabs_[static_cast<size_t>(tab_index)].results[static_cast<size_t>(item_index)];
            bool selected = ListViewItemSelected(search_results_list_, item_index);
            if (DrawSearchMatchSubItem(result, draw->iSubItem, selected, draw->nmcd.hdc, draw->nmcd.rc, ui_font_)) {
              return CDRF_SKIPDEFAULT;
            }
          }
        }
        return ui::HandleThemedListViewCustomDraw(search_results_list_, draw);
      }
      if (header->hwndFrom == toolbar_.hwnd()) {
        auto* draw = reinterpret_cast<NMTBCUSTOMDRAW*>(lparam);
        const Theme& theme = Theme::Current();
        if (draw->nmcd.dwDrawStage == CDDS_PREPAINT) {
          HBRUSH brush = GetCachedBrush(theme.BackgroundColor());
          FillRect(draw->nmcd.hdc, &draw->nmcd.rc, brush);
          return CDRF_NOTIFYITEMDRAW;
        }
        if (draw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
          int command_id = static_cast<int>(draw->nmcd.dwItemSpec);
          int index = static_cast<int>(SendMessageW(toolbar_.hwnd(), TB_COMMANDTOINDEX, command_id, 0));
          if (index >= 0) {
            TBBUTTON button = {};
            if (SendMessageW(toolbar_.hwnd(), TB_GETBUTTON, index, reinterpret_cast<LPARAM>(&button))) {
              if (button.fsStyle & BTNS_SEP) {
                RECT rect = draw->nmcd.rc;
                int mid_x = (rect.left + rect.right) / 2;
                HPEN pen = GetCachedPen(theme.BorderColor(), 1);
                HPEN old = reinterpret_cast<HPEN>(SelectObject(draw->nmcd.hdc, pen));
                MoveToEx(draw->nmcd.hdc, mid_x, rect.top + 4, nullptr);
                LineTo(draw->nmcd.hdc, mid_x, rect.bottom - 4);
                SelectObject(draw->nmcd.hdc, old);
                return CDRF_SKIPDEFAULT;
              }
            }
          }
          if (draw->nmcd.uItemState & (CDIS_HOT | CDIS_SELECTED)) {
            COLORREF hot = (draw->nmcd.uItemState & CDIS_SELECTED) ? theme.SelectionColor() : theme.HoverColor();
            HBRUSH brush = GetCachedBrush(hot);
            FillRect(draw->nmcd.hdc, &draw->nmcd.rc, brush);
            draw->clrBtnFace = hot;
            draw->clrBtnHighlight = hot;
          } else {
            draw->clrBtnFace = theme.BackgroundColor();
            draw->clrBtnHighlight = theme.BackgroundColor();
          }
          draw->clrText = theme.TextColor();
          return TBCDRF_NOEDGES;
        }
      }
      if (header->hwndFrom == history_list_) {
        return HandleHistoryListCustomDraw(history_list_, reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
      }
      if (header->hwndFrom == value_list_.hwnd()) {
        return ui::HandleThemedListViewCustomDraw(header->hwndFrom, reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
      }
      if (header->hwndFrom == tree_.hwnd()) {
        auto* draw = reinterpret_cast<NMTVCUSTOMDRAW*>(lparam);
        const Theme& theme = Theme::Current();
        if (draw->nmcd.dwDrawStage == CDDS_PREPAINT) {
          return CDRF_NOTIFYITEMDRAW;
        }
        if (draw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
          if (draw->nmcd.uItemState & CDIS_SELECTED) {
            draw->clrText = theme.SelectionTextColor();
            draw->clrTextBk = theme.SelectionColor();
          } else {
            draw->clrText = theme.TextColor();
            draw->clrTextBk = theme.PanelColor();
          }
          return CDRF_NEWFONT;
        }
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
  default:
    break;
  }
  return DefWindowProcW(hwnd_, message, wparam, lparam);
}

bool MainWindow::OnCreate() {
  ui_font_ = CreateUIFont();
  icon_font_ = CreateIconFont(10);
  custom_font_ = DefaultLogFont();
  LoadSettings();
  LoadThemePresets();
  LoadTreeState();
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
  EnableAddressAutoComplete();

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

  ReloadThemeIcons();
  ApplyUIFontToControls();

  ApplyThemeToChildren();
  CreateValueColumns();
  CreateHistoryColumns();
  CreateSearchColumns();
  UpdateSearchResultsView();
  LoadHistoryCache();
  LoadComments();
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
  tree_.SetRootLabel(TreeRootLabel());
  tree_.PopulateRoots(roots_);

  SelectDefaultTreeItem();
  RestoreTreeState();
  StartTreeStateWorker();
  MarkTreeStateDirty();
  StartValueListWorker();

  ApplyViewVisibility();
  ApplyAlwaysOnTop();
  UpdateStatus();

  return true;
}

void MainWindow::OnDestroy() {
  StopTraceParseSessions();
  StopDefaultParseSessions();
  StopRegFileParseSessions();
  StopTraceLoadWorker();
  StopDefaultLoadWorker();
  StopValueListWorker();
  StopTreeStateWorker();
  CancelSearch();
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

void MainWindow::ApplyThemeToChildren() {
  const Theme& theme = Theme::Current();

  theme.ApplyToToolbar(toolbar_.hwnd());
  theme.ApplyToTreeView(tree_.hwnd());
  theme.ApplyToListView(value_list_.hwnd());
  theme.ApplyToListView(history_list_);
  theme.ApplyToListView(search_results_list_);
  theme.ApplyToTabControl(tab_);
  theme.ApplyToStatusBar(status_bar_);

  if (address_edit_) {
    SetWindowTheme(address_edit_, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(address_edit_, 6, 6);
    SetEditVerticalRect(address_edit_, ui_font_, 2, 6, 6);
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
  ApplyFont(address_go_btn_, ui_font_);
  ApplyFont(filter_edit_, ui_font_);
  ApplyFont(tab_, ui_font_);
  ApplyFont(tree_header_, ui_font_);
  ApplyFont(tree_close_btn_, ui_font_);
  ApplyFont(tree_.hwnd(), ui_font_);
  ApplyFont(value_list_.hwnd(), ui_font_);
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
  const int filter_min_width = 160;
  const int filter_max_width = 260;
  const int filter_gap = 6;
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
  int history_height = show_history ? std::clamp(history_height_, min_history, max_history) : 0;
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
  int tree_width = show_tree ? std::clamp(tree_width_, min_tree, max_tree) : 0;
  if (show_tree) {
    tree_width_ = tree_width;
  }

  int tree_header_height = drag_tree_header_height_;
  int history_label_height = drag_history_label_height_;
  int list_x = show_tree ? (content_left + tree_width + kSplitterWidth) : content_left;
  int list_width = content_right - list_x;
  int tree_content_height = std::max(0, content_height - (show_tree ? tree_header_height : 0));

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
  desired = std::clamp(desired, splitter_min_width_, splitter_max_width_);
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
  desired = std::clamp(desired, history_splitter_min_height_, history_splitter_max_height_);
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
  } else if (entry.kind == TabEntry::Kind::kRegFile) {
    SyncRegFileTabSelection();
  }
}

void MainWindow::ResetHiveListCache() {
  hive_list_loaded_ = false;
  hive_list_.clear();
}

void MainWindow::EnsureHiveListLoaded() {
  if (hive_list_loaded_) {
    return;
  }
  hive_list_loaded_ = true;
  hive_list_.clear();

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
    hive_list_.emplace(ToLower(name), std::move(data));
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
  std::wstring nt_path = RegistryProvider::BuildNtPath(node);
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
        int target_width = std::clamp(available / 4, filter_min_width, filter_max_width);
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
      int filter_width = std::clamp(available, filter_min_width, filter_max_width);
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
  int history_height = show_history ? std::clamp(history_height_, min_history, max_history) : 0;
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
    if (list_hwnd) {
      SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
      InvalidateRect(list_hwnd, nullptr, TRUE);
    }
    UpdateStatus();
    updating_value_list_ = false;
    value_list_loading_ = false;
    return;
  }

  RegistryNode snapshot = *node;
  std::wstring path = RegistryProvider::BuildPath(snapshot);
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
  auto value_comments = value_comments_;
  auto name_comments = name_comments_;
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

  if (list_hwnd) {
    SendMessageW(list_hwnd, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(list_hwnd, nullptr, TRUE);
  }
  UpdateStatus();
  updating_value_list_ = false;
  value_list_loading_ = true;

  int sort_column = value_sort_column_;
  bool sort_ascending = value_sort_ascending_;
  bool show_keys_in_list = show_keys_in_list_;
  std::unordered_map<std::wstring, std::wstring> hive_list;
  if (show_keys_in_list) {
    EnsureHiveListLoaded();
    hive_list = hive_list_;
  }
  if (!value_list_thread_.joinable()) {
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
  task->hwnd = hwnd_;
  task->trace_data_list = std::move(trace_data_list);
  task->default_data_list = std::move(default_data_list);
  task->hive_list = std::move(hive_list);
  task->value_comments = std::move(value_comments);
  task->name_comments = std::move(name_comments);
  {
    std::lock_guard<std::mutex> lock(value_list_mutex_);
    value_list_task_ = std::move(task);
    value_list_pending_ = true;
  }
  value_list_cv_.notify_one();
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
    return;
  }
  if (!current_node_) {
    return;
  }

  ValueEntry entry;
  if (!RegistryProvider::QueryValue(*current_node_, row->extra, &entry)) {
    row->data.clear();
    row->data_ready = true;
    return;
  }

  row->value_type = entry.type;
  row->value_data_size = static_cast<DWORD>(entry.data.size());
  row->data = RegistryProvider::FormatValueDataForDisplay(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
  row->data_ready = true;
  if (row->type.empty()) {
    row->type = RegistryProvider::FormatValueType(entry.type);
  }
  row->size_value = row->value_data_size;
  row->has_size = true;
  if (row->size.empty() && row->value_data_size > 0) {
    row->size = std::to_wstring(row->value_data_size);
  }
}

void MainWindow::UpdateAddressBar(RegistryNode* node) {
  if (!node || !address_edit_) {
    return;
  }
  std::wstring path = RegistryProvider::BuildPath(*node);
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
  std::wstring path = StripRegFileKeySyntax(input);
  path = StripOuterQuotes(path);
  path = TrimWhitespace(path);
  if (path.empty()) {
    return path;
  }
  for (auto& ch : path) {
    if (ch == L'/') {
      ch = L'\\';
    }
  }
  path = CollapseBackslashes(path);
  if (StartsWithInsensitive(path, L"Registry::")) {
    path.erase(0, wcslen(L"Registry::"));
  }
  while (!path.empty() && path.front() == L'\\') {
    path.erase(path.begin());
  }

  if (StartsWithInsensitive(path, L"Computer\\")) {
    path.erase(0, wcslen(L"Computer\\"));
  }
  std::wstring root_label = TreeRootLabel();
  if (!root_label.empty()) {
    std::wstring prefix = root_label + L"\\";
    if (StartsWithInsensitive(path, prefix)) {
      path.erase(0, prefix.size());
    }
  }
  if (registry_mode_ == RegistryMode::kRemote && !remote_machine_.empty()) {
    std::wstring machine = StripMachinePrefix(remote_machine_);
    if (!machine.empty()) {
      std::wstring prefix = machine + L"\\";
      if (StartsWithInsensitive(path, prefix)) {
        path.erase(0, prefix.size());
      }
    }
  }

  std::wstring sid = util::GetCurrentUserSidString();
  std::wstring native_mapped = MapNativeRegistryPath(path, sid);
  if (!native_mapped.empty()) {
    path = native_mapped;
  }

  size_t split = path.find_first_of(L":\\");
  std::wstring prefix = (split == std::wstring::npos) ? path : path.substr(0, split);
  std::wstring rest = (split == std::wstring::npos) ? L"" : path.substr(split);

  auto normalize_rest = [&]() {
    if (!rest.empty() && rest.front() == L':') {
      rest.erase(rest.begin());
    }
    while (!rest.empty() && rest.front() == L'\\') {
      rest.erase(rest.begin());
    }
  };

  auto map_prefix = [&](const wchar_t* short_name, const wchar_t* full_name) -> bool {
    if (_wcsicmp(prefix.c_str(), short_name) == 0) {
      prefix = full_name;
      normalize_rest();
      return true;
    }
    return false;
  };

  map_prefix(L"REGISTRY", L"REGISTRY") || map_prefix(L"HKCR", L"HKEY_CLASSES_ROOT") || map_prefix(L"HKCU", L"HKEY_CURRENT_USER") || map_prefix(L"HKLM", L"HKEY_LOCAL_MACHINE") || map_prefix(L"HKU", L"HKEY_USERS") || map_prefix(L"HKCC", L"HKEY_CURRENT_CONFIG") || map_prefix(L"HKEY_CLASSES_ROOT", L"HKEY_CLASSES_ROOT") || map_prefix(L"HKEY_CURRENT_USER", L"HKEY_CURRENT_USER") || map_prefix(L"HKEY_LOCAL_MACHINE", L"HKEY_LOCAL_MACHINE") || map_prefix(L"HKEY_USERS", L"HKEY_USERS") || map_prefix(L"HKEY_CURRENT_CONFIG", L"HKEY_CURRENT_CONFIG");

  if (!rest.empty()) {
    return prefix + L"\\" + rest;
  }
  return prefix;
}

std::wstring MainWindow::FormatRegistryPath(const std::wstring& path, RegistryPathFormat format) const {
  std::wstring normalized = NormalizeRegistryPath(path);
  if (normalized.empty()) {
    return normalized;
  }

  size_t slash = normalized.find(L'\\');
  std::wstring root = (slash == std::wstring::npos) ? normalized : normalized.substr(0, slash);
  std::wstring rest = (slash == std::wstring::npos) ? L"" : normalized.substr(slash + 1);

  auto abbrev_root = [&](const std::wstring& full) -> std::wstring {
    if (EqualsInsensitive(full, L"HKEY_CLASSES_ROOT")) {
      return L"HKCR";
    }
    if (EqualsInsensitive(full, L"HKEY_CURRENT_USER")) {
      return L"HKCU";
    }
    if (EqualsInsensitive(full, L"HKEY_LOCAL_MACHINE")) {
      return L"HKLM";
    }
    if (EqualsInsensitive(full, L"HKEY_USERS")) {
      return L"HKU";
    }
    if (EqualsInsensitive(full, L"HKEY_CURRENT_CONFIG")) {
      return L"HKCC";
    }
    if (EqualsInsensitive(full, L"REGISTRY")) {
      return L"REGISTRY";
    }
    return full;
  };

  auto join = [&](const std::wstring& prefix, const std::wstring& suffix) -> std::wstring {
    if (suffix.empty()) {
      return prefix;
    }
    return prefix + L"\\" + suffix;
  };

  switch (format) {
  case RegistryPathFormat::kAbbrev:
    return join(abbrev_root(root), rest);
  case RegistryPathFormat::kRegedit: {
    std::wstring label = (registry_mode_ == RegistryMode::kLocal) ? L"Computer" : TreeRootLabel();
    if (label.empty()) {
      label = L"Computer";
    }
    return join(label, join(root, rest));
  }
  case RegistryPathFormat::kRegFile: {
    std::wstring inner = join(root, rest);
    return L"[" + inner + L"]";
  }
  case RegistryPathFormat::kPowerShellDrive: {
    std::wstring drive = abbrev_root(root);
    std::wstring combined = drive + L":";
    if (!rest.empty()) {
      combined += L"\\" + rest;
    }
    return combined;
  }
  case RegistryPathFormat::kPowerShellProvider: {
    std::wstring inner = join(root, rest);
    return L"Registry::" + inner;
  }
  case RegistryPathFormat::kEscaped: {
    return EscapeBackslashes(join(root, rest));
  }
  case RegistryPathFormat::kFull:
  default:
    return join(root, rest);
  }
}

bool MainWindow::FindNearestExistingPath(const std::wstring& path, std::wstring* nearest_path) const {
  if (!nearest_path) {
    return false;
  }
  RegistryNode node;
  if (!ResolvePathToNode(path, &node)) {
    return false;
  }
  if (node.subkey.empty()) {
    *nearest_path = node.root_name;
    return true;
  }
  std::vector<std::wstring> parts = SplitPath(node.subkey);
  std::wstring existing;
  for (const auto& part : parts) {
    std::wstring candidate = existing.empty() ? part : existing + L"\\" + part;
    RegistryNode test = node;
    test.subkey = candidate;
    KeyInfo info = {};
    if (RegistryProvider::QueryKeyInfo(test, &info)) {
      existing = candidate;
      continue;
    }
    break;
  }
  if (existing.empty()) {
    *nearest_path = node.root_name;
    return true;
  }
  *nearest_path = node.root_name + L"\\" + existing;
  return true;
}

bool MainWindow::CreateRegistryPath(const std::wstring& path) {
  RegistryNode node;
  if (!ResolvePathToNode(path, &node)) {
    return false;
  }
  if (node.subkey.empty()) {
    return true;
  }
  std::vector<std::wstring> parts = SplitPath(node.subkey);
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
    path_text = RegistryProvider::BuildPath(*current_node_);
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

int MainWindow::FindFirstSearchTabIndex() const {
  for (size_t i = 0; i < tabs_.size(); ++i) {
    if (tabs_[i].kind == TabEntry::Kind::kSearch) {
      return static_cast<int>(i);
    }
  }
  return -1;
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
      registry_scope_path = RegistryProvider::BuildPath(*current_node_);
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

  SearchCriteria criteria = options.criteria;
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
  last_search_results_.clear();
  last_search_index_ = 0;

  search_cancel_.store(false);
  search_progress_searched_.store(0);
  search_progress_total_.store(0);
  search_progress_percent_ = 0;
  search_progress_posted_.store(false);
  search_posted_.store(false);
  {
    std::lock_guard<std::mutex> lock(search_mutex_);
    search_pending_.clear();
  }
  search_last_refresh_tick_ = 0;
  search_start_tick_ = GetTickCount64();
  search_duration_ms_ = 0;
  search_duration_valid_ = false;
  search_running_ = true;
  search_generation_ += 1;
  uint64_t generation = search_generation_;
  search_tabs_[static_cast<size_t>(search_index)].generation = generation;

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

  search_thread_ = std::thread([this, criteria, traces, exclude_paths, scope_lower, scope_recursive, trace_enabled, registry_enabled, generation, matcher]() mutable {
    auto should_stop = [&]() { return search_cancel_.load(); };

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
        for (auto& item : pending) {
          search_pending_.push_back(std::move(item));
        }
      }
      if (!search_posted_.exchange(true)) {
        PostMessageW(hwnd_, kSearchResultsMessage, static_cast<WPARAM>(generation), 0);
      }
    };

    auto queue_result = [&](SearchResult&& result) {
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
          if (!SelectionIncludesKey(trace.selection, key_lower)) {
            continue;
          }
          if (!key_in_scope(key_lower)) {
            continue;
          }
          std::wstring key_name = KeyLeafFromPath(key_path);

          if (criteria.search_keys) {
            TextMatch match = matcher.Match(key_name);
            if (match.matched) {
              SearchResult result;
              result.key_path = key_path;
              result.key_name = key_name;
              result.type_text = L"Trace Key";
              result.is_key = true;
              size_t path_start = key_path.size() >= key_name.size() ? key_path.size() - key_name.size() : 0;
              result.match_field = SearchMatchField::kPath;
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
                if (!SelectionIncludesValue(trace.selection, key_lower, value_lower)) {
                  continue;
                }
                std::wstring display = value_name.empty() ? L"(Default)" : value_name;
                TextMatch match = matcher.Match(display);
                if (!match.matched) {
                  continue;
                }
                SearchResult result;
                result.key_path = key_path;
                result.key_name = key_name;
                result.value_name = value_name;
                result.display_name = display;
                result.type_text = L"Trace Value";
                result.is_key = false;
                result.match_field = SearchMatchField::kName;
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
      bool ok = SearchRegistryStreaming(
          criteria, &search_cancel_,
          [&](const SearchResult& result) -> bool {
            if (should_stop()) {
              return false;
            }
            SearchResult copy = result;
            queue_result(std::move(copy));
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

  bool matcher_ok = true;
  ReplaceMatcher matcher(options, &matcher_ok);
  if (!matcher_ok) {
    ui::ShowError(hwnd_, L"Invalid replace pattern.");
    return;
  }

  std::vector<RegistryNode> stack;
  stack.push_back(start);
  int replaced = 0;
  int failures = 0;

  while (!stack.empty()) {
    RegistryNode node = std::move(stack.back());
    stack.pop_back();

    auto values = RegistryProvider::EnumValues(node);
    for (const auto& value : values) {
      std::wstring current_name = value.name;
      std::wstring replaced_name;
      if (!current_name.empty() && matcher.Replace(current_name, &replaced_name) && replaced_name != current_name) {
        if (replaced_name.empty()) {
          continue;
        }
        std::wstring unique = MakeUniqueValueName(node, replaced_name);
        if (!RegistryProvider::RenameValue(node, current_name, unique)) {
          ++failures;
        } else {
          UndoOperation op;
          op.type = UndoOperation::Type::kRenameValue;
          op.node = node;
          op.name = current_name;
          op.new_name = unique;
          PushUndo(std::move(op));
          AppendHistoryEntry(L"Rename value " + current_name, current_name, unique);
          MarkOfflineDirty();
          current_name = unique;
          ++replaced;
        }
      }

      if (value.type == REG_SZ || value.type == REG_EXPAND_SZ || value.type == REG_MULTI_SZ) {
        std::vector<BYTE> new_data = value.data;
        bool changed = false;
        if (value.type == REG_MULTI_SZ) {
          auto parts = MultiSzToVector(value.data);
          for (auto& part : parts) {
            std::wstring updated;
            if (matcher.Replace(part, &updated) && updated != part) {
              part = std::move(updated);
              changed = true;
            }
          }
          if (changed) {
            new_data = VectorToMultiSz(parts);
          }
        } else {
          std::wstring text = RegistryProvider::FormatValueData(value.type, value.data.data(), static_cast<DWORD>(value.data.size()));
          std::wstring updated;
          if (matcher.Replace(text, &updated) && updated != text) {
            new_data = StringToRegData(updated);
            changed = true;
          }
        }

        if (changed) {
          if (!RegistryProvider::SetValue(node, current_name, value.type, new_data)) {
            ++failures;
          } else {
            ValueEntry old_value = value;
            old_value.name = current_name;
            ValueEntry new_value = value;
            new_value.name = current_name;
            new_value.data = new_data;
            UndoOperation op;
            op.type = UndoOperation::Type::kModifyValue;
            op.node = node;
            op.old_value = old_value;
            op.new_value = new_value;
            PushUndo(std::move(op));
            std::wstring old_text = RegistryProvider::FormatValueData(value.type, value.data.data(), static_cast<DWORD>(value.data.size()));
            std::wstring new_text = RegistryProvider::FormatValueData(value.type, new_data.data(), static_cast<DWORD>(new_data.size()));
            AppendHistoryEntry(L"Modify value " + current_name, old_text, new_text);
            MarkOfflineDirty();
            ++replaced;
          }
        }
      }
    }

    if (options.recursive) {
      auto subkeys = RegistryProvider::EnumSubKeyNames(node, false);
      for (const auto& name : subkeys) {
        stack.push_back(MakeChildNode(node, name));
      }
    }
  }

  if (current_node_) {
    UpdateValueListForNode(current_node_);
  }
  if (failures > 0) {
    std::wstring message = L"Replace finished with some failures.\nReplaced: " + std::to_wstring(replaced) + L"\nFailed: " + std::to_wstring(failures);
    ui::ShowError(hwnd_, message);
  }
}

void MainWindow::CancelSearch() {
  search_cancel_.store(true);
  if (search_thread_.joinable()) {
    search_thread_.join();
  }
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
    search_pending_.clear();
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

  SortHistoryEntries(&history_entries_, history_sort_column_, history_sort_ascending_);
  RebuildHistoryList();

  HWND header = ListView_GetHeader(history_list_);
  if (header) {
    UpdateListViewSort(history_list_, history_sort_column_, history_sort_ascending_);
    InvalidateRect(header, nullptr, TRUE);
  }
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
  SortSearchResultEntries(&tab.results, tab.sort_column, tab.sort_ascending, tab.is_compare);
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
  history_entries_.clear();
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
  for (const auto& entry : history_entries_) {
    LVITEMW item = {};
    item.mask = LVIF_TEXT;
    item.iItem = index;
    item.pszText = const_cast<wchar_t*>(entry.time_text.c_str());
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
        it->second->cancel.store(true);
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
  UpdateTabWidth();
  ApplyViewVisibility();
  UpdateSearchResultsView();
  UpdateStatus();
}

void MainWindow::OpenLocalRegistryTab() {
  if (!tab_) {
    return;
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
  UpdateTabWidth();
  TabCtrl_SetCurSel(tab_, index);
  SwitchToLocalRegistry();
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
  tree_.SetRootLabel(TreeRootLabel());
  tree_.PopulateRoots(roots_);
  ResetNavigationState();
  UpdateStatus();

  SelectDefaultTreeItem();
}

std::wstring MainWindow::TreeRootLabel() const {
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

void MainWindow::AppendRealRegistryRoot(std::vector<RegistryRootEntry>* roots) {
  if (!roots || registry_mode_ != RegistryMode::kLocal) {
    return;
  }
  if (!registry_root_.get()) {
    registry_root_ = OpenRegistryRootKey();
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
  UpdateTabText(L"Local Registry");
  UpdateRegistryTabEntry(RegistryMode::kLocal, L"", L"");
  std::vector<RegistryRootEntry> roots = RegistryProvider::DefaultRoots(show_extra_hives_);
  AppendRealRegistryRoot(&roots);
  ApplyRegistryRoots(roots);
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
  return true;
}

void MainWindow::NavigateToAddress() {
  wchar_t buffer[512] = {};
  GetWindowTextW(address_edit_, buffer, static_cast<int>(_countof(buffer)));
  std::wstring path = NormalizeRegistryPath(buffer);
  if (path.empty()) {
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

void MainWindow::AppendHistoryEntry(const std::wstring& action, const std::wstring& old_data, const std::wstring& new_data) {
  if (!history_list_) {
    return;
  }

  SYSTEMTIME st = {};
  GetLocalTime(&st);
  wchar_t time_buffer[64] = {};
  swprintf_s(time_buffer, L"%d/%d/%d %d:%02d:%02d", st.wMonth, st.wDay, st.wYear, st.wHour, st.wMinute, st.wSecond);

  FILETIME now = {};
  GetSystemTimeAsFileTime(&now);

  HistoryEntry entry;
  entry.timestamp = FileTimeToUint64(now);
  entry.time_text = time_buffer;
  entry.action = action;
  entry.old_data = old_data;
  entry.new_data = new_data;
  history_entries_.push_back(entry);
  if (history_loaded_) {
    AppendHistoryCache(entry);
  }

  if (history_entries_.size() > static_cast<size_t>(history_max_rows_)) {
    while (history_entries_.size() > static_cast<size_t>(history_max_rows_)) {
      auto oldest_it = std::min_element(history_entries_.begin(), history_entries_.end(), [](const HistoryEntry& left, const HistoryEntry& right) { return left.timestamp < right.timestamp; });
      if (oldest_it == history_entries_.end()) {
        break;
      }
      history_entries_.erase(oldest_it);
    }
  }

  SortHistoryEntries(&history_entries_, history_sort_column_, history_sort_ascending_);
  RebuildHistoryList();
}

std::wstring MainWindow::ResolveSearchComment(const SearchResult& result) const {
  if (result.is_key) {
    return L"";
  }

  std::wstring value_key = MakeValueCommentKey(result.key_path, result.value_name, result.type);
  auto it = value_comments_.find(value_key);
  if (it != value_comments_.end()) {
    return FormatCommentDisplay(it->second.text);
  }

  std::wstring name_key = MakeNameCommentKey(result.value_name, result.type);
  auto it2 = name_comments_.find(name_key);
  if (it2 != name_comments_.end()) {
    return FormatCommentDisplay(it2->second.text);
  }

  return L"";
}

void MainWindow::LoadHistoryCache() {
  if (history_loaded_) {
    return;
  }
  std::wstring path = HistoryCachePath();
  if (path.empty()) {
    history_loaded_ = true;
    return;
  }

  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    history_loaded_ = true;
    return;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    history_loaded_ = true;
    return;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    history_loaded_ = true;
    return;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(buffer);
  if (content.empty()) {
    history_loaded_ = true;
    return;
  }

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
    if (line.empty()) {
      continue;
    }
    std::vector<std::wstring> parts = SplitHistoryFields(line);
    if (parts.size() < 5) {
      continue;
    }
    HistoryEntry entry;
    try {
      entry.timestamp = std::stoull(parts[0]);
    } catch (const std::exception&) {
      continue;
    }
    entry.time_text = UnescapeHistoryField(parts[1]);
    entry.action = UnescapeHistoryField(parts[2]);
    entry.old_data = UnescapeHistoryField(parts[3]);
    entry.new_data = UnescapeHistoryField(parts[4]);
    history_entries_.push_back(std::move(entry));
  }

  if (history_entries_.size() > static_cast<size_t>(history_max_rows_)) {
    while (history_entries_.size() > static_cast<size_t>(history_max_rows_)) {
      auto oldest_it = std::min_element(history_entries_.begin(), history_entries_.end(), [](const HistoryEntry& left, const HistoryEntry& right) { return left.timestamp < right.timestamp; });
      if (oldest_it == history_entries_.end()) {
        break;
      }
      history_entries_.erase(oldest_it);
    }
  }
  SortHistoryEntries(&history_entries_, history_sort_column_, history_sort_ascending_);
  RebuildHistoryList();
  history_loaded_ = true;
}

void MainWindow::AppendHistoryCache(const HistoryEntry& entry) {
  std::wstring path = HistoryCachePath();
  if (path.empty()) {
    return;
  }
  std::wstring line = std::to_wstring(entry.timestamp);
  line.push_back(L'\t');
  line.append(EscapeHistoryField(entry.time_text));
  line.push_back(L'\t');
  line.append(EscapeHistoryField(entry.action));
  line.push_back(L'\t');
  line.append(EscapeHistoryField(entry.old_data));
  line.push_back(L'\t');
  line.append(EscapeHistoryField(entry.new_data));
  line.push_back(L'\n');

  std::string utf8 = util::WideToUtf8(line);
  if (utf8.empty()) {
    return;
  }
  HANDLE file = CreateFileW(path.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  DWORD written = 0;
  WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  CloseHandle(file);
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

bool MainWindow::ReadSearchResults(const std::wstring& path, std::vector<SearchResult>* results) const {
  if (!results) {
    return false;
  }
  results->clear();
  if (path.empty()) {
    return false;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart < 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return false;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok) {
    return false;
  }
  if (read == 0) {
    return true;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(buffer);
  if (content.empty()) {
    return true;
  }

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
    if (line.empty()) {
      continue;
    }
    std::vector<std::wstring> parts = SplitHistoryFields(line);
    if (parts.size() < 13) {
      continue;
    }
    SearchResult result;
    result.key_path = UnescapeHistoryField(parts[0]);
    result.key_name = UnescapeHistoryField(parts[1]);
    result.value_name = UnescapeHistoryField(parts[2]);
    result.display_name = UnescapeHistoryField(parts[3]);
    result.type_text = UnescapeHistoryField(parts[4]);
    result.type = static_cast<DWORD>(_wtoi(parts[5].c_str()));
    result.data = UnescapeHistoryField(parts[6]);
    result.size_text = UnescapeHistoryField(parts[7]);
    result.date_text = UnescapeHistoryField(parts[8]);
    size_t base_index = 9;
    if (parts.size() >= 14) {
      result.comment = UnescapeHistoryField(parts[9]);
      base_index = 10;
    }
    result.is_key = (_wtoi(parts[base_index].c_str()) != 0);
    int match_field = _wtoi(parts[base_index + 1].c_str());
    if (match_field < 0 || match_field > static_cast<int>(SearchMatchField::kData)) {
      result.match_field = SearchMatchField::kNone;
    } else {
      result.match_field = static_cast<SearchMatchField>(match_field);
    }
    result.match_start = _wtoi(parts[base_index + 2].c_str());
    result.match_length = _wtoi(parts[base_index + 3].c_str());
    results->push_back(std::move(result));
  }
  return true;
}

bool MainWindow::WriteSearchResults(const std::wstring& path, const std::vector<SearchResult>& results) const {
  if (path.empty()) {
    return false;
  }
  if (results.empty()) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
      return false;
    }
    CloseHandle(file);
    return true;
  }
  std::wstring content;
  for (const auto& result : results) {
    content.append(EscapeHistoryField(result.key_path));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.key_name));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.value_name));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.display_name));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.type_text));
    content.push_back(L'\t');
    content.append(std::to_wstring(result.type));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.data));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.size_text));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.date_text));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(result.comment));
    content.push_back(L'\t');
    content.append(result.is_key ? L"1" : L"0");
    content.push_back(L'\t');
    content.append(std::to_wstring(static_cast<int>(result.match_field)));
    content.push_back(L'\t');
    content.append(std::to_wstring(result.match_start));
    content.push_back(L'\t');
    content.append(std::to_wstring(result.match_length));
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  if (utf8.empty()) {
    return false;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  DWORD written = 0;
  WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  CloseHandle(file);
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
    std::wstring path = TabsCachePath();
    if (!path.empty()) {
      HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
      if (file != INVALID_HANDLE_VALUE) {
        LARGE_INTEGER size = {};
        if (GetFileSizeEx(file, &size) && size.QuadPart >= 0 && size.QuadPart <= static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
          std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
          DWORD read = 0;
          if (ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0) {
            buffer.resize(read);
            if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
              buffer.erase(0, 3);
            }
            std::wstring content = util::Utf8ToWide(buffer);
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
              if (line.empty()) {
                continue;
              }
              if (line.rfind(L"active=", 0) == 0) {
                active_index = _wtoi(line.substr(7).c_str());
                continue;
              }
              std::vector<std::wstring> parts = SplitHistoryFields(line);
              if (parts.size() < 3) {
                continue;
              }
              if (_wcsicmp(parts[0].c_str(), L"tab") != 0) {
                continue;
              }
              std::wstring type = parts[1];
              std::wstring label = UnescapeHistoryField(parts[2]);
              if (_wcsicmp(type.c_str(), L"registry") == 0) {
                if (label.empty()) {
                  label = L"Local Registry";
                }
                TCITEMW item = {};
                item.mask = TCIF_TEXT;
                item.pszText = const_cast<wchar_t*>(label.c_str());
                TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
                TabEntry entry;
                entry.kind = TabEntry::Kind::kRegistry;
                entry.registry_mode = RegistryMode::kLocal;
                tabs_.push_back(std::move(entry));
              } else if (_wcsicmp(type.c_str(), L"search") == 0 && parts.size() >= 4) {
                std::wstring file = UnescapeHistoryField(parts[3]);
                SearchTab tab;
                tab.label = label.empty() ? L"Find" : label;
                tab.is_compare = StartsWithInsensitive(tab.label, L"Compare:");
                std::wstring result_path = SearchTabCachePath(file);
                ReadSearchResults(result_path, &tab.results);
                search_tabs_.push_back(std::move(tab));
                int search_index = static_cast<int>(search_tabs_.size() - 1);
                TCITEMW item = {};
                item.mask = TCIF_TEXT;
                item.pszText = const_cast<wchar_t*>(search_tabs_.back().label.c_str());
                TabCtrl_InsertItem(tab_, TabCtrl_GetItemCount(tab_), &item);
                tabs_.push_back({TabEntry::Kind::kSearch, search_index});
              }
            }
            loaded = true;
          }
        }
        CloseHandle(file);
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
    int sel = std::clamp(active_index, 0, count - 1);
    TabCtrl_SetCurSel(tab_, sel);
    if (IsSearchTabIndex(sel)) {
      active_search_tab_index_ = sel;
    }
  }
  UpdateTabWidth();
}

void MainWindow::SaveTabs() {
  if (!tab_) {
    return;
  }
  std::wstring folder = CacheFolderPath();
  if (folder.empty()) {
    return;
  }

  std::unordered_set<std::wstring> referenced_files;
  std::wstring content;
  int active_index = TabCtrl_GetCurSel(tab_);
  int saved_active_index = -1;
  content.append(L"active=");
  content.append(std::to_wstring(active_index));
  content.push_back(L'\n');

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
      std::wstring file_name = L"search_" + std::to_wstring(search_file_index++) + L".tsv";
      std::wstring result_path = SearchTabCachePath(file_name);
      WriteSearchResults(result_path, search_tabs_[static_cast<size_t>(search_index)].results);
      referenced_files.insert(file_name);
      if (label.empty()) {
        label = search_tabs_[static_cast<size_t>(search_index)].label;
      }
      content.append(L"tab\t");
      content.append(L"search\t");
      content.append(EscapeHistoryField(label));
      content.push_back(L'\t');
      content.append(EscapeHistoryField(file_name));
      content.push_back(L'\n');
    } else {
      if (label.empty()) {
        label = L"Local Registry";
      }
      content.append(L"tab\t");
      content.append(L"registry\t");
      content.append(EscapeHistoryField(label));
      content.push_back(L'\n');
    }
    ++saved_index;
  }
  if (saved_active_index < 0) {
    saved_active_index = 0;
  }
  size_t newline = content.find(L'\n');
  if (newline != std::wstring::npos) {
    std::wstring header = L"active=" + std::to_wstring(saved_active_index);
    content.replace(0, newline, header);
  }

  std::string utf8 = util::WideToUtf8(content);
  if (!utf8.empty()) {
    std::wstring tabs_path = TabsCachePath();
    HANDLE file = CreateFileW(tabs_path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file != INVALID_HANDLE_VALUE) {
      DWORD written = 0;
      WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
      CloseHandle(file);
    }
  }

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
  value_comments_.clear();
  name_comments_.clear();
  std::wstring path = CommentsPath();
  if (path.empty()) {
    return;
  }
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
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
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
    if (line.empty()) {
      continue;
    }
    std::vector<std::wstring> parts = SplitHistoryFields(line);
    if (parts.size() < 5) {
      continue;
    }
    std::wstring scope = parts[0];
    std::wstring path_field = UnescapeHistoryField(parts[1]);
    std::wstring name_field = UnescapeHistoryField(parts[2]);
    DWORD type = static_cast<DWORD>(_wtoi(parts[3].c_str()));
    std::wstring text = UnescapeHistoryField(parts[4]);
    if (IsWhitespaceOnly(text)) {
      continue;
    }
    if (_wcsicmp(scope.c_str(), L"value") == 0) {
      CommentEntry entry;
      entry.path = path_field;
      entry.name = name_field;
      entry.type = type;
      entry.text = text;
      value_comments_[MakeValueCommentKey(path_field, name_field, type)] = std::move(entry);
    } else if (_wcsicmp(scope.c_str(), L"name") == 0) {
      CommentEntry entry;
      entry.name = name_field;
      entry.type = type;
      entry.text = text;
      name_comments_[MakeNameCommentKey(name_field, type)] = std::move(entry);
    }
  }
}

void MainWindow::SaveComments() const {
  std::wstring path = CommentsPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& pair : value_comments_) {
    const auto& entry = pair.second;
    if (IsWhitespaceOnly(entry.text)) {
      continue;
    }
    content.append(L"value\t");
    content.append(EscapeHistoryField(entry.path));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.name));
    content.push_back(L'\t');
    content.append(std::to_wstring(entry.type));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.text));
    content.push_back(L'\n');
  }
  for (const auto& pair : name_comments_) {
    const auto& entry = pair.second;
    if (IsWhitespaceOnly(entry.text)) {
      continue;
    }
    content.append(L"name\t");
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.name));
    content.push_back(L'\t');
    content.append(std::to_wstring(entry.type));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.text));
    content.push_back(L'\n');
  }
  if (content.empty()) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file != INVALID_HANDLE_VALUE) {
      CloseHandle(file);
    }
    return;
  }
  std::string utf8 = util::WideToUtf8(content);
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  DWORD written = 0;
  WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  CloseHandle(file);
}

bool MainWindow::ImportCommentsFromFile(const std::wstring& path) {
  if (path.empty()) {
    return false;
  }
  value_comments_.clear();
  name_comments_.clear();
  std::wstring original = CommentsPath();
  if (!original.empty() && _wcsicmp(path.c_str(), original.c_str()) == 0) {
    LoadComments();
  } else {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
      return false;
    }
    LARGE_INTEGER size = {};
    if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
      CloseHandle(file);
      return false;
    }
    std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
    DWORD read = 0;
    bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
    CloseHandle(file);
    if (!ok || read == 0) {
      return false;
    }
    buffer.resize(read);
    if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
      buffer.erase(0, 3);
    }
    std::wstring content = util::Utf8ToWide(buffer);
    if (content.empty()) {
      return false;
    }
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
      if (line.empty()) {
        continue;
      }
      std::vector<std::wstring> parts = SplitHistoryFields(line);
      if (parts.size() < 5) {
        continue;
      }
      std::wstring scope = parts[0];
      std::wstring path_field = UnescapeHistoryField(parts[1]);
      std::wstring name_field = UnescapeHistoryField(parts[2]);
      DWORD type = static_cast<DWORD>(_wtoi(parts[3].c_str()));
      std::wstring text = UnescapeHistoryField(parts[4]);
      if (IsWhitespaceOnly(text)) {
        continue;
      }
      if (_wcsicmp(scope.c_str(), L"value") == 0) {
        CommentEntry entry;
        entry.path = path_field;
        entry.name = name_field;
        entry.type = type;
        entry.text = text;
        value_comments_[MakeValueCommentKey(path_field, name_field, type)] = std::move(entry);
      } else if (_wcsicmp(scope.c_str(), L"name") == 0) {
        CommentEntry entry;
        entry.name = name_field;
        entry.type = type;
        entry.text = text;
        name_comments_[MakeNameCommentKey(name_field, type)] = std::move(entry);
      }
    }
  }
  SaveComments();
  RefreshValueListComments();
  return true;
}

bool MainWindow::ExportCommentsToFile(const std::wstring& path) const {
  if (path.empty()) {
    return false;
  }
  std::wstring content;
  for (const auto& pair : value_comments_) {
    const auto& entry = pair.second;
    if (IsWhitespaceOnly(entry.text)) {
      continue;
    }
    content.append(L"value\t");
    content.append(EscapeHistoryField(entry.path));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.name));
    content.push_back(L'\t');
    content.append(std::to_wstring(entry.type));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.text));
    content.push_back(L'\n');
  }
  for (const auto& pair : name_comments_) {
    const auto& entry = pair.second;
    if (IsWhitespaceOnly(entry.text)) {
      continue;
    }
    content.append(L"name\t");
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.name));
    content.push_back(L'\t');
    content.append(std::to_wstring(entry.type));
    content.push_back(L'\t');
    content.append(EscapeHistoryField(entry.text));
    content.push_back(L'\n');
  }
  if (content.empty()) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
      return false;
    }
    CloseHandle(file);
    return true;
  }
  std::string utf8 = util::WideToUtf8(content);
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  DWORD written = 0;
  WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  CloseHandle(file);
  return true;
}

void MainWindow::RefreshValueListComments() {
  if (!current_node_) {
    return;
  }
  std::wstring path = RegistryProvider::BuildPath(*current_node_);
  bool changed = false;
  for (auto& row : value_list_.rows()) {
    if (row.kind != rowkind::kValue) {
      if (!row.comment.empty()) {
        row.comment.clear();
        changed = true;
      }
      continue;
    }
    std::wstring value_key = MakeValueCommentKey(path, row.extra, row.value_type);
    std::wstring text;
    auto it = value_comments_.find(value_key);
    if (it != value_comments_.end()) {
      text = it->second.text;
    } else {
      std::wstring name_key = MakeNameCommentKey(row.extra, row.value_type);
      auto it2 = name_comments_.find(name_key);
      if (it2 != name_comments_.end()) {
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
  if (value_list_.HasFilter()) {
    value_list_.RebuildFilter();
  } else if (changed && value_list_.hwnd()) {
    InvalidateRect(value_list_.hwnd(), nullptr, TRUE);
  }
}

bool MainWindow::EditValueComment(const ListRow& row) {
  if (!current_node_ || row.kind != rowkind::kValue) {
    return false;
  }
  std::wstring path = RegistryProvider::BuildPath(*current_node_);
  std::wstring value_key = MakeValueCommentKey(path, row.extra, row.value_type);
  std::wstring name_key = MakeNameCommentKey(row.extra, row.value_type);
  bool has_value = (value_comments_.find(value_key) != value_comments_.end());
  bool has_name = (name_comments_.find(name_key) != name_comments_.end());
  std::wstring initial;
  bool apply_all = false;
  if (has_value) {
    initial = value_comments_[value_key].text;
  } else if (has_name) {
    initial = name_comments_[name_key].text;
    apply_all = true;
  }
  std::wstring updated = initial;
  bool apply_all_out = apply_all;
  if (!PromptForComment(hwnd_, updated, apply_all_out, &updated, &apply_all_out)) {
    return false;
  }
  if (IsWhitespaceOnly(updated)) {
    updated.clear();
  }
  if (updated.empty()) {
    value_comments_.erase(value_key);
    name_comments_.erase(name_key);
  } else if (apply_all_out) {
    CommentEntry entry;
    entry.name = row.extra;
    entry.type = row.value_type;
    entry.text = updated;
    name_comments_[name_key] = std::move(entry);
    value_comments_.erase(value_key);
  } else {
    CommentEntry entry;
    entry.path = path;
    entry.name = row.extra;
    entry.type = row.value_type;
    entry.text = updated;
    value_comments_[value_key] = std::move(entry);
  }
  SaveComments();
  RefreshValueListComments();
  return true;
}

void MainWindow::LoadSettings() {
  std::wstring path = SettingsPath();
  if (path.empty()) {
    return;
  }
  bool font_size_set = false;
  int font_size = 0;
  auto parse_indexed_key = [](const std::wstring& key, const wchar_t* prefix, int* index) -> bool {
    size_t prefix_len = wcslen(prefix);
    if (_wcsnicmp(key.c_str(), prefix, prefix_len) != 0) {
      return false;
    }
    const wchar_t* start = key.c_str() + prefix_len;
    if (*start == L'\0') {
      return false;
    }
    wchar_t* end = nullptr;
    long value = wcstol(start, &end, 10);
    if (end == start || *end != L'\0' || value < 0) {
      return false;
    }
    if (index) {
      *index = static_cast<int>(value);
    }
    return true;
  };
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
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
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
    if (line.empty()) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = TrimWhitespace(line.substr(0, sep));
    std::wstring value = TrimWhitespace(line.substr(sep + 1));
    int column_index = -1;
    if (_wcsicmp(key.c_str(), L"clear_history_on_exit") == 0) {
      clear_history_on_exit_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"clear_tabs_on_exit") == 0) {
      clear_tabs_on_exit_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_toolbar") == 0) {
      show_toolbar_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_address_bar") == 0) {
      show_address_bar_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_filter_bar") == 0) {
      show_filter_bar_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_tab_control") == 0) {
      show_tab_control_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_tree") == 0) {
      show_tree_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_history") == 0) {
      show_history_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_status_bar") == 0) {
      show_status_bar_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_keys_in_list") == 0) {
      show_keys_in_list_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_simulated_keys") == 0) {
      show_simulated_keys_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"view_extra_hives") == 0) {
      show_extra_hives_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"save_tree_state") == 0) {
      save_tree_state_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"save_tabs") == 0) {
      save_tabs_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"window_x") == 0) {
      window_x_ = _wtoi(value.c_str());
      window_placement_loaded_ = true;
    } else if (_wcsicmp(key.c_str(), L"window_y") == 0) {
      window_y_ = _wtoi(value.c_str());
      window_placement_loaded_ = true;
    } else if (_wcsicmp(key.c_str(), L"window_width") == 0) {
      window_width_ = _wtoi(value.c_str());
      window_placement_loaded_ = true;
    } else if (_wcsicmp(key.c_str(), L"window_height") == 0) {
      window_height_ = _wtoi(value.c_str());
      window_placement_loaded_ = true;
    } else if (_wcsicmp(key.c_str(), L"window_maximized") == 0) {
      window_maximized_ = parse_bool(value);
      window_placement_loaded_ = true;
    } else if (_wcsicmp(key.c_str(), L"always_on_top") == 0) {
      always_on_top_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"replace_regedit") == 0) {
      replace_regedit_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"single_instance") == 0) {
      single_instance_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"read_only") == 0) {
      read_only_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"always_run_as_admin") == 0) {
      always_run_as_admin_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"always_run_as_system") == 0) {
      always_run_as_system_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"always_run_as_trustedinstaller") == 0) {
      always_run_as_trustedinstaller_ = parse_bool(value);
    } else if (_wcsicmp(key.c_str(), L"theme_mode") == 0) {
      if (_wcsicmp(value.c_str(), L"dark") == 0) {
        theme_mode_ = ThemeMode::kDark;
      } else if (_wcsicmp(value.c_str(), L"light") == 0) {
        theme_mode_ = ThemeMode::kLight;
      } else if (_wcsicmp(value.c_str(), L"custom") == 0) {
        theme_mode_ = ThemeMode::kCustom;
      } else {
        theme_mode_ = ThemeMode::kSystem;
      }
    } else if (_wcsicmp(key.c_str(), L"theme_preset") == 0) {
      active_theme_preset_ = value;
    } else if (_wcsicmp(key.c_str(), L"icon_set") == 0) {
      if (IsIconSetName(value, kIconSetDefault)) {
        icon_set_ = kIconSetDefault;
      } else if (IsIconSetName(value, kIconSetLucide)) {
        icon_set_ = kIconSetDefault;
      } else if (IsIconSetName(value, kIconSetTabler)) {
        icon_set_ = kIconSetTabler;
      } else if (IsIconSetName(value, kIconSetFluentUi)) {
        icon_set_ = kIconSetFluentUi;
      } else if (IsIconSetName(value, kIconSetMaterialSymbols)) {
        icon_set_ = kIconSetMaterialSymbols;
      } else if (IsIconSetName(value, kIconSetCustom)) {
        icon_set_ = kIconSetCustom;
      } else {
        icon_set_ = kIconSetDefault;
      }
    } else if (_wcsicmp(key.c_str(), L"tree_width") == 0) {
      int width = _wtoi(value.c_str());
      if (width > 0) {
        tree_width_ = width;
      }
    } else if (_wcsicmp(key.c_str(), L"history_height") == 0) {
      int height = _wtoi(value.c_str());
      if (height > 0) {
        history_height_ = height;
      }
    } else if (parse_indexed_key(key, L"value_column_width_", &column_index)) {
      int width = _wtoi(value.c_str());
      if (column_index >= 0 && width >= 0) {
        if (static_cast<size_t>(column_index) >= saved_value_column_widths_.size()) {
          saved_value_column_widths_.resize(static_cast<size_t>(column_index) + 1, 0);
        }
        saved_value_column_widths_[static_cast<size_t>(column_index)] = width;
        saved_value_columns_loaded_ = true;
      }
    } else if (parse_indexed_key(key, L"value_column_visible_", &column_index)) {
      bool visible = parse_bool(value);
      if (column_index >= 0) {
        if (static_cast<size_t>(column_index) >= saved_value_column_visible_.size()) {
          saved_value_column_visible_.resize(static_cast<size_t>(column_index) + 1, true);
        }
        saved_value_column_visible_[static_cast<size_t>(column_index)] = visible;
        saved_value_columns_loaded_ = true;
      }
    } else if (_wcsicmp(key.c_str(), L"font_use_default") == 0) {
      bool use_default = parse_bool(value);
      use_custom_font_ = !use_default;
    } else if (_wcsicmp(key.c_str(), L"font_face") == 0) {
      if (!value.empty()) {
        wcsncpy_s(custom_font_.lfFaceName, value.c_str(), _TRUNCATE);
      }
    } else if (_wcsicmp(key.c_str(), L"font_size") == 0) {
      int size = _wtoi(value.c_str());
      if (size > 0) {
        font_size = size;
        font_size_set = true;
      }
    } else if (_wcsicmp(key.c_str(), L"font_weight") == 0) {
      int weight = _wtoi(value.c_str());
      if (weight > 0) {
        custom_font_.lfWeight = weight;
      }
    } else if (_wcsicmp(key.c_str(), L"font_italic") == 0) {
      bool italic = parse_bool(value);
      custom_font_.lfItalic = italic ? TRUE : FALSE;
    } else if (parse_indexed_key(key, L"trace_recent_", &column_index)) {
      if (column_index >= 0) {
        if (static_cast<size_t>(column_index) >= recent_trace_paths_.size()) {
          recent_trace_paths_.resize(static_cast<size_t>(column_index) + 1);
        }
        recent_trace_paths_[static_cast<size_t>(column_index)] = value;
      }
    } else if (parse_indexed_key(key, L"default_recent_", &column_index)) {
      if (column_index >= 0) {
        if (static_cast<size_t>(column_index) >= recent_default_paths_.size()) {
          recent_default_paths_.resize(static_cast<size_t>(column_index) + 1);
        }
        recent_default_paths_[static_cast<size_t>(column_index)] = value;
      }
    }
  }
  if (always_run_as_trustedinstaller_) {
    always_run_as_system_ = false;
    always_run_as_admin_ = false;
  } else if (always_run_as_system_) {
    always_run_as_admin_ = false;
  }
  if (!save_tree_state_) {
    saved_tree_selected_path_.clear();
    saved_tree_expanded_paths_.clear();
  }
  if (font_size_set) {
    custom_font_.lfHeight = FontHeightFromPointSize(font_size);
  }
  NormalizeRecentTraceList();
  NormalizeRecentDefaultList();
}

void MainWindow::SaveSettings() const {
  std::wstring path = SettingsPath();
  if (path.empty()) {
    return;
  }
  int window_x = window_x_;
  int window_y = window_y_;
  int window_w = window_width_;
  int window_h = window_height_;
  bool window_max = window_maximized_;
  if (hwnd_ && IsWindow(hwnd_)) {
    WINDOWPLACEMENT placement = {};
    placement.length = sizeof(placement);
    if (GetWindowPlacement(hwnd_, &placement)) {
      RECT normal = placement.rcNormalPosition;
      int width = normal.right - normal.left;
      int height = normal.bottom - normal.top;
      if (width > 0 && height > 0) {
        window_x = normal.left;
        window_y = normal.top;
        window_w = width;
        window_h = height;
      }
      window_max = (placement.showCmd == SW_SHOWMAXIMIZED);
    }
  }
  std::wstring content = L"clear_history_on_exit=";
  content += clear_history_on_exit_ ? L"1\n" : L"0\n";
  content += L"clear_tabs_on_exit=";
  content += clear_tabs_on_exit_ ? L"1\n" : L"0\n";
  content += L"view_toolbar=";
  content += show_toolbar_ ? L"1\n" : L"0\n";
  content += L"view_address_bar=";
  content += show_address_bar_ ? L"1\n" : L"0\n";
  content += L"view_filter_bar=";
  content += show_filter_bar_ ? L"1\n" : L"0\n";
  content += L"view_tab_control=";
  content += show_tab_control_ ? L"1\n" : L"0\n";
  content += L"view_tree=";
  content += show_tree_ ? L"1\n" : L"0\n";
  content += L"view_history=";
  content += show_history_ ? L"1\n" : L"0\n";
  content += L"view_status_bar=";
  content += show_status_bar_ ? L"1\n" : L"0\n";
  content += L"view_keys_in_list=";
  content += show_keys_in_list_ ? L"1\n" : L"0\n";
  content += L"view_simulated_keys=";
  content += show_simulated_keys_ ? L"1\n" : L"0\n";
  content += L"view_extra_hives=";
  content += show_extra_hives_ ? L"1\n" : L"0\n";
  content += L"save_tree_state=";
  content += save_tree_state_ ? L"1\n" : L"0\n";
  content += L"save_tabs=";
  content += save_tabs_ ? L"1\n" : L"0\n";
  content += L"always_run_as_admin=";
  content += always_run_as_admin_ ? L"1\n" : L"0\n";
  content += L"always_run_as_system=";
  content += always_run_as_system_ ? L"1\n" : L"0\n";
  content += L"always_run_as_trustedinstaller=";
  content += always_run_as_trustedinstaller_ ? L"1\n" : L"0\n";
  if (window_w > 0 && window_h > 0) {
    content += L"window_x=";
    content += std::to_wstring(window_x);
    content.push_back(L'\n');
    content += L"window_y=";
    content += std::to_wstring(window_y);
    content.push_back(L'\n');
    content += L"window_width=";
    content += std::to_wstring(window_w);
    content.push_back(L'\n');
    content += L"window_height=";
    content += std::to_wstring(window_h);
    content.push_back(L'\n');
    content += L"window_maximized=";
    content += window_max ? L"1\n" : L"0\n";
  }
  content += L"always_on_top=";
  content += always_on_top_ ? L"1\n" : L"0\n";
  content += L"replace_regedit=";
  content += replace_regedit_ ? L"1\n" : L"0\n";
  content += L"single_instance=";
  content += single_instance_ ? L"1\n" : L"0\n";
  content += L"read_only=";
  content += read_only_ ? L"1\n" : L"0\n";
  content += L"theme_mode=";
  if (theme_mode_ == ThemeMode::kDark) {
    content += L"dark\n";
  } else if (theme_mode_ == ThemeMode::kLight) {
    content += L"light\n";
  } else if (theme_mode_ == ThemeMode::kCustom) {
    content += L"custom\n";
  } else {
    content += L"system\n";
  }
  content += L"theme_preset=";
  content += active_theme_preset_;
  content.push_back(L'\n');
  content += L"icon_set=";
  if (IsKnownIconSetName(icon_set_)) {
    content += icon_set_;
  } else {
    content += kIconSetDefault;
  }
  content.push_back(L'\n');
  content += L"tree_width=";
  content += std::to_wstring(tree_width_);
  content.push_back(L'\n');
  content += L"history_height=";
  content += std::to_wstring(history_height_);
  content.push_back(L'\n');
  content += L"font_use_default=";
  content += use_custom_font_ ? L"0\n" : L"1\n";
  if (custom_font_.lfFaceName[0] != L'\0') {
    content += L"font_face=";
    content += custom_font_.lfFaceName;
    content.push_back(L'\n');
  }
  int font_size = FontPointSize(custom_font_);
  if (font_size > 0) {
    content += L"font_size=";
    content += std::to_wstring(font_size);
    content.push_back(L'\n');
  }
  content += L"font_weight=";
  content += std::to_wstring(custom_font_.lfWeight);
  content.push_back(L'\n');
  content += L"font_italic=";
  content += custom_font_.lfItalic ? L"1\n" : L"0\n";
  for (size_t i = 0; i < recent_trace_paths_.size(); ++i) {
    if (recent_trace_paths_[i].empty()) {
      continue;
    }
    content += L"trace_recent_";
    content += std::to_wstring(i);
    content.push_back(L'=');
    content += recent_trace_paths_[i];
    content.push_back(L'\n');
  }
  for (size_t i = 0; i < recent_default_paths_.size(); ++i) {
    if (recent_default_paths_[i].empty()) {
      continue;
    }
    content += L"default_recent_";
    content += std::to_wstring(i);
    content.push_back(L'=');
    content += recent_default_paths_[i];
    content.push_back(L'\n');
  }
  for (size_t i = 0; i < value_columns_.size(); ++i) {
    int width = 0;
    if (i < value_column_widths_.size()) {
      width = value_column_widths_[i];
    }
    content += L"value_column_width_";
    content += std::to_wstring(i);
    content.push_back(L'=');
    content += std::to_wstring(width);
    content.push_back(L'\n');
    bool visible = true;
    if (i < value_column_visible_.size()) {
      visible = value_column_visible_[i];
    }
    content += L"value_column_visible_";
    content += std::to_wstring(i);
    content.push_back(L'=');
    content += visible ? L"1\n" : L"0\n";
  }
  std::string utf8 = util::WideToUtf8(content);
  if (utf8.empty()) {
    return;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  DWORD written = 0;
  WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  CloseHandle(file);
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
  saved_tree_selected_path_.clear();
  saved_tree_expanded_paths_.clear();
  if (!save_tree_state_) {
    return;
  }
  std::wstring path = TreeStatePath();
  if (path.empty()) {
    return;
  }

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
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
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
    if (line.empty()) {
      continue;
    }
    if (line.front() == L'#') {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = TrimWhitespace(line.substr(0, sep));
    std::wstring value = line.substr(sep + 1);
    if (EqualsInsensitive(key, L"selected")) {
      saved_tree_selected_path_ = UnescapeHistoryField(value);
    } else if (EqualsInsensitive(key, L"expanded")) {
      std::wstring path_value = UnescapeHistoryField(value);
      if (!path_value.empty()) {
        saved_tree_expanded_paths_.push_back(std::move(path_value));
      }
    }
  }
}

void MainWindow::StartTreeStateWorker() {
  if (!save_tree_state_ || tree_state_thread_.joinable()) {
    return;
  }
  tree_state_stop_ = false;
  tree_state_thread_ = std::thread([this]() {
    for (;;) {
      std::unique_lock<std::mutex> lock(tree_state_mutex_);
      tree_state_cv_.wait_for(lock, std::chrono::seconds(2), [this]() { return tree_state_stop_ || tree_state_dirty_; });
      if (tree_state_stop_) {
        break;
      }
      if (!tree_state_dirty_) {
        continue;
      }
      std::wstring selected = tree_state_selected_;
      std::vector<std::wstring> expanded = tree_state_expanded_;
      tree_state_dirty_ = false;
      lock.unlock();
      SaveTreeStateFile(selected, expanded);
    }
  });
}

void MainWindow::StopTreeStateWorker() {
  if (save_tree_state_ && tree_.hwnd() && IsWindow(tree_.hwnd())) {
    std::wstring selected;
    std::vector<std::wstring> expanded;
    CaptureTreeState(&selected, &expanded);
    SaveTreeStateFile(selected, expanded);
  }
  {
    std::lock_guard<std::mutex> lock(tree_state_mutex_);
    tree_state_stop_ = true;
  }
  tree_state_cv_.notify_one();
  if (tree_state_thread_.joinable()) {
    tree_state_thread_.join();
  }
}

void MainWindow::StartValueListWorker() {
  if (value_list_thread_.joinable()) {
    return;
  }
  value_list_stop_ = false;
  value_list_thread_ = std::thread([this]() {
    for (;;) {
      std::unique_ptr<ValueListTask> task;
      {
        std::unique_lock<std::mutex> lock(value_list_mutex_);
        value_list_cv_.wait(lock, [this]() { return value_list_stop_ || value_list_pending_; });
        if (value_list_stop_) {
          break;
        }
        task = std::move(value_list_task_);
        value_list_pending_ = false;
      }
      if (!task) {
        continue;
      }
      if (task->generation != value_list_generation_.load()) {
        continue;
      }

      auto payload = std::make_unique<ValueListPayload>();
      payload->generation = task->generation;
      std::wstring node_path = RegistryProvider::BuildPath(task->snapshot);

      auto resolve_comment = [&](const std::wstring& name, DWORD type) -> std::wstring {
        std::wstring value_key = MakeValueCommentKey(node_path, name, type);
        auto it = task->value_comments.find(value_key);
        if (it != task->value_comments.end()) {
          return it->second.text;
        }
        std::wstring name_key = MakeNameCommentKey(name, type);
        auto it2 = task->name_comments.find(name_key);
        if (it2 != task->name_comments.end()) {
          return it2->second.text;
        }
        return {};
      };

      auto lookup_hive_root = [&](const RegistryNode& node) -> bool {
        if (task->hive_list.empty()) {
          return false;
        }
        std::wstring nt_path = RegistryProvider::BuildNtPath(node);
        if (nt_path.empty()) {
          return false;
        }
        std::wstring nt_lower = ToLower(nt_path);
        for (const auto& entry : task->hive_list) {
          const std::wstring& hive_key = entry.first;
          if (nt_lower == hive_key) {
            return true;
          }
        }
        return false;
      };
      auto resolve_key_icon = [&](const RegistryNode& node, bool* is_link) -> int {
        if (is_link) {
          *is_link = false;
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
        if (lookup_hive_root(node)) {
          return kDatabaseIconIndex;
        }
        return kFolderIconIndex;
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
        std::wstring path = RegistryProvider::BuildPath(node);
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
          if (!SelectionIncludesKey(trace.selection, key_lower)) {
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
        TraceKeyValues values;
        const TraceSelection* selection = nullptr;
      };
      std::vector<TraceMatch> trace_matches;
      if (!task->trace_data_list.empty()) {
        for (const auto& trace : task->trace_data_list) {
          if (!trace.data) {
            continue;
          }
          std::shared_lock<std::shared_mutex> trace_lock(*trace.data->mutex);
          if (!SelectionIncludesKey(trace.selection, task->trace_path_lower)) {
            continue;
          }
          auto it = trace.data->values_by_key.find(task->trace_path_lower);
          if (it == trace.data->values_by_key.end()) {
            continue;
          }
          TraceMatch match;
          match.label = trace.label.empty() ? L"Trace" : trace.label;
          match.values = it->second;
          match.selection = &trace.selection;
          trace_matches.push_back(std::move(match));
        }
      }

      struct DefaultMatch {
        DefaultKeyValues values;
        const KeyValueSelection* selection = nullptr;
      };
      std::vector<DefaultMatch> default_keys;
      if (!task->default_data_list.empty() && !task->default_path_lower.empty()) {
        default_keys.reserve(task->default_data_list.size());
        for (const auto& defaults : task->default_data_list) {
          if (!defaults.data) {
            continue;
          }
          std::shared_lock<std::shared_mutex> defaults_lock(*defaults.data->mutex);
          if (!SelectionIncludesKey(defaults.selection, task->default_path_lower)) {
            continue;
          }
          auto it = defaults.data->values_by_key.find(task->default_path_lower);
          if (it == defaults.data->values_by_key.end()) {
            continue;
          }
          default_keys.push_back({it->second, &defaults.selection});
        }
      }
      auto resolve_default_data = [&](const std::wstring& value_name) -> std::wstring {
        if (default_keys.empty()) {
          return {};
        }
        std::wstring value_lower = ToLower(value_name);
        bool applies = false;
        for (const auto& match : default_keys) {
          if (match.selection && !SelectionIncludesValue(*match.selection, task->default_path_lower, value_lower)) {
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
            if (match.selection && !SelectionIncludesValue(*match.selection, task->trace_path_lower, value_lower)) {
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
      auto values = RegistryProvider::EnumValues(task->snapshot);
      payload->value_count = static_cast<int>(values.size());
      if (track_existing) {
        existing_values.reserve(values.size());
      }
      payload->rows.reserve(payload->rows.size() + values.size());
      for (const auto& value : values) {
        if (value.name.empty()) {
          has_default = true;
        }
        if (EqualsInsensitive(value.name, L"SymbolicLinkValue")) {
          has_symbolic_value = true;
        }
        ListRow row;
        row.name = value.name.empty() ? L"(Default)" : value.name;
        row.type = RegistryProvider::FormatValueType(value.type);
        row.data = RegistryProvider::FormatValueDataForDisplay(value.type, value.data.data(), static_cast<DWORD>(value.data.size()));
        row.data_ready = true;
        row.default_data = resolve_default_data(value.name);
        row.image_index = UseBinaryValueIcon(value.type) ? kBinaryIconIndex : kValueIconIndex;
        row.kind = rowkind::kValue;
        row.extra = value.name;
        row.size_value = static_cast<uint64_t>(value.data.size());
        row.has_size = true;
        row.value_type = value.type;
        row.comment = FormatCommentDisplay(resolve_comment(value.name, value.type));
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
      }

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
        row.size_value = 0;
        row.has_size = true;
        row.value_type = REG_SZ;
        row.comment = FormatCommentDisplay(resolve_comment(L"", REG_SZ));
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
            if (match.selection && !SelectionIncludesValue(*match.selection, task->trace_path_lower, value_lower)) {
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
            row.comment = FormatCommentDisplay(resolve_comment(value_name, 0));
            payload->rows.emplace_back(std::move(row));
            ++trace_added;
            existing_values.insert(value_lower);
          }
        }
        payload->value_count += static_cast<int>(trace_added);
      }

      SortValueRows(&payload->rows, task->sort_column, task->sort_ascending);
      if (task->generation != value_list_generation_.load()) {
        continue;
      }
      if (PostMessageW(task->hwnd, kValueListReadyMessage, static_cast<WPARAM>(task->generation), reinterpret_cast<LPARAM>(payload.get())) != 0) {
        payload.release();
      }
    }
  });
}

void MainWindow::StopValueListWorker() {
  {
    std::lock_guard<std::mutex> lock(value_list_mutex_);
    value_list_stop_ = true;
    value_list_pending_ = false;
    value_list_task_.reset();
  }
  value_list_cv_.notify_one();
  if (value_list_thread_.joinable()) {
    value_list_thread_.join();
  }
}

void MainWindow::StartTraceParseThread(TraceParseSession* session) {
  if (!session || session->thread.joinable()) {
    return;
  }
  session->cancel.store(false);
  HWND hwnd = hwnd_;
  std::wstring source = session->source_path;
  std::wstring source_lower = session->source_lower;
  session->thread = std::thread([this, session, hwnd, source, source_lower]() {
    constexpr size_t kBatchSize = 256;
    constexpr DWORD kBatchMs = 50;
    auto post_batch = [&](std::vector<KeyValueDialogEntry>* entries, bool done, const std::wstring& error, bool cancelled) {
      auto payload = std::make_unique<TraceParseBatch>();
      payload->source_lower = source_lower;
      if (entries) {
        payload->entries = std::move(*entries);
      }
      payload->done = done;
      payload->error = error;
      payload->cancelled = cancelled;
      if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, kTraceParseBatchMessage, 0, reinterpret_cast<LPARAM>(payload.get()))) {
        return;
      }
      payload.release();
    };

    std::string buffer;
    if (!ReadFileBinary(source, &buffer)) {
      post_batch(nullptr, true, L"Failed to read trace file.", false);
      return;
    }
    if (buffer.empty()) {
      post_batch(nullptr, true, L"Trace file is empty or too large to load.", false);
      return;
    }
    if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
      buffer.erase(0, 3);
    }
    std::wstring content = util::Utf8ToWide(buffer);
    if (content.empty()) {
      post_batch(nullptr, true, L"Trace file has no readable entries.", false);
      return;
    }

    std::vector<KeyValueDialogEntry> entries;
    entries.reserve(kBatchSize);
    uint64_t last_post = GetTickCount64();
    size_t start = 0;
    while (start < content.size()) {
      if (session->cancel.load()) {
        post_batch(nullptr, true, L"", true);
        return;
      }
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
        continue;
      }
      size_t sep = line.rfind(L" : ");
      size_t sep_len = 0;
      if (sep != std::wstring::npos) {
        sep_len = 3;
      } else {
        sep = line.rfind(L':');
        sep_len = (sep == std::wstring::npos) ? 0 : 1;
      }
      if (sep == std::wstring::npos) {
        continue;
      }
      std::wstring key_text = TrimWhitespace(line.substr(0, sep));
      std::wstring value_text = TrimWhitespace(line.substr(sep + sep_len));
      if (key_text.empty()) {
        continue;
      }
      std::wstring selection_path = NormalizeTraceSelectionPath(key_text);
      if (selection_path.empty()) {
        continue;
      }
      std::wstring key_path = NormalizeTraceKeyPath(key_text);
      if (key_path.empty()) {
        key_path = selection_path;
      }
      std::wstring value_name = value_text;
      if (EqualsInsensitive(value_name, L"(Default)")) {
        value_name.clear();
      }
      KeyValueDialogEntry entry;
      entry.key_path = key_path;
      entry.display_path = selection_path;
      entry.has_value = true;
      entry.value_name = value_name;
      entries.push_back(std::move(entry));

      uint64_t now = GetTickCount64();
      if (entries.size() >= kBatchSize || (now - last_post) >= kBatchMs) {
        post_batch(&entries, false, L"", false);
        entries.clear();
        last_post = now;
      }
    }

    if (session->cancel.load()) {
      post_batch(nullptr, true, L"", true);
      return;
    }
    if (!entries.empty()) {
      post_batch(&entries, false, L"", false);
      entries.clear();
    }
    post_batch(nullptr, true, L"", false);
  });
}

void MainWindow::StartDefaultParseThread(DefaultParseSession* session) {
  if (!session || session->thread.joinable()) {
    return;
  }
  session->cancel.store(false);
  HWND hwnd = hwnd_;
  std::wstring source = session->source_path;
  std::wstring source_lower = session->source_lower;
  session->thread = std::thread([this, session, hwnd, source, source_lower]() {
    constexpr size_t kBatchSize = 256;
    constexpr DWORD kBatchMs = 50;
    auto post_batch = [&](std::vector<KeyValueDialogEntry>* entries, bool done, const std::wstring& error, bool cancelled) {
      auto payload = std::make_unique<DefaultParseBatch>();
      payload->source_lower = source_lower;
      if (entries) {
        payload->entries = std::move(*entries);
      }
      payload->done = done;
      payload->error = error;
      payload->cancelled = cancelled;
      if (!hwnd || !IsWindow(hwnd) || !PostMessageW(hwnd, kDefaultParseBatchMessage, 0, reinterpret_cast<LPARAM>(payload.get()))) {
        return;
      }
      payload.release();
    };

    std::wstring content;
    if (!ReadRegFileText(source, &content)) {
      post_batch(nullptr, true, L"Failed to read registry file.", false);
      return;
    }
    if (content.empty()) {
      post_batch(nullptr, true, L"Default file contains no usable entries.", false);
      return;
    }

    std::vector<KeyValueDialogEntry> entries;
    entries.reserve(kBatchSize);
    uint64_t last_post = GetTickCount64();
    std::wstring current_key;
    std::wstring current_display;
    std::wstring current;
    bool saw_entry = false;

    size_t start = 0;
    while (start < content.size()) {
      if (session->cancel.load()) {
        post_batch(nullptr, true, L"", true);
        return;
      }
      size_t end = content.find(L'\n', start);
      if (end == std::wstring::npos) {
        end = content.size();
      }
      std::wstring line = content.substr(start, end - start);
      if (!line.empty() && line.back() == L'\r') {
        line.pop_back();
      }
      start = end + 1;
      if (current.empty()) {
        current = line;
      } else {
        current.append(line);
      }
      std::wstring trimmed_right = current;
      while (!trimmed_right.empty() && (trimmed_right.back() == L' ' || trimmed_right.back() == L'\t')) {
        trimmed_right.pop_back();
      }
      if (!trimmed_right.empty() && trimmed_right.back() == L'\\') {
        trimmed_right.pop_back();
        current = trimmed_right;
        continue;
      }

      std::wstring raw = TrimWhitespace(current);
      current.clear();
      if (raw.empty() || raw.front() == L';') {
        continue;
      }

      if (raw.front() == L'[' && raw.back() == L']') {
        std::wstring key = raw.substr(1, raw.size() - 2);
        key = TrimWhitespace(key);
        bool delete_key = !key.empty() && key.front() == L'-';
        if (delete_key) {
          current_key.clear();
          current_display.clear();
          continue;
        }
        std::wstring normalized = NormalizeTraceKeyPathBasic(key);
        current_key = normalized.empty() ? key : normalized;
        current_display = NormalizeTraceSelectionPath(key);
        if (current_display.empty()) {
          current_display = current_key;
        }
        if (!current_key.empty()) {
          KeyValueDialogEntry entry;
          entry.key_path = current_key;
          entry.display_path = current_display;
          entry.has_value = false;
          entries.push_back(std::move(entry));
          saw_entry = true;
        }
      } else {
        if (current_key.empty()) {
          continue;
        }
        size_t eq = raw.find(L'=');
        if (eq == std::wstring::npos) {
          continue;
        }
        std::wstring name_part = TrimWhitespace(raw.substr(0, eq));
        std::wstring data_part = TrimWhitespace(raw.substr(eq + 1));
        if (name_part.empty() || data_part.empty() || data_part == L"-") {
          continue;
        }

        std::wstring value_name;
        if (name_part == L"@") {
          value_name.clear();
        } else if (name_part.front() == L'\"') {
          size_t end_pos = 0;
          if (!ParseQuotedString(name_part, &value_name, &end_pos)) {
            continue;
          }
        } else {
          continue;
        }

        DWORD type = REG_NONE;
        std::vector<BYTE> data;
        if (data_part.front() == L'\"') {
          std::wstring text;
          size_t end_pos = 0;
          if (!ParseQuotedString(data_part, &text, &end_pos)) {
            continue;
          }
          type = REG_SZ;
          data = StringToRegData(text);
        } else if (StartsWithInsensitive(data_part, L"dword:")) {
          std::wstring hex = TrimWhitespace(data_part.substr(6));
          if (hex.empty()) {
            continue;
          }
          DWORD number = static_cast<DWORD>(wcstoul(hex.c_str(), nullptr, 16));
          type = REG_DWORD;
          data.resize(sizeof(DWORD));
          memcpy(data.data(), &number, sizeof(DWORD));
        } else if (StartsWithInsensitive(data_part, L"hex")) {
          type = REG_BINARY;
          size_t colon = data_part.find(L':');
          if (colon == std::wstring::npos) {
            continue;
          }
          size_t open = data_part.find(L'(');
          size_t close = data_part.find(L')');
          if (open != std::wstring::npos && close != std::wstring::npos && close > open) {
            std::wstring code = data_part.substr(open + 1, close - open - 1);
            unsigned long parsed = wcstoul(code.c_str(), nullptr, 16);
            switch (parsed) {
            case 0x0:
              type = REG_NONE;
              break;
            case 0x1:
              type = REG_SZ;
              break;
            case 0x2:
              type = REG_EXPAND_SZ;
              break;
            case 0x3:
              type = REG_BINARY;
              break;
            case 0x4:
              type = REG_DWORD;
              break;
            case 0x5:
              type = REG_DWORD_BIG_ENDIAN;
              break;
            case 0x7:
              type = REG_MULTI_SZ;
              break;
            case 0x8:
              type = REG_RESOURCE_LIST;
              break;
            case 0x9:
              type = REG_FULL_RESOURCE_DESCRIPTOR;
              break;
            case 0xA:
              type = REG_RESOURCE_REQUIREMENTS_LIST;
              break;
            case 0xB:
              type = REG_QWORD;
              break;
            default:
              type = REG_BINARY;
              break;
            }
          }
          std::wstring hex = data_part.substr(colon + 1);
          if (!ParseHexBytes(hex, &data)) {
            continue;
          }
        } else {
          continue;
        }

        KeyValueDialogEntry entry;
        entry.key_path = current_key;
        entry.display_path = current_display;
        entry.has_value = true;
        entry.value_name = value_name;
        entry.value_type = type;
        entry.value_data = RegistryProvider::FormatValueDataForDisplay(type, data.empty() ? nullptr : data.data(), static_cast<DWORD>(data.size()));
        entries.push_back(std::move(entry));
        saw_entry = true;
      }

      uint64_t now = GetTickCount64();
      if (entries.size() >= kBatchSize || (now - last_post) >= kBatchMs) {
        post_batch(&entries, false, L"", false);
        entries.clear();
        last_post = now;
      }
    }

    if (session->cancel.load()) {
      post_batch(nullptr, true, L"", true);
      return;
    }
    if (!entries.empty()) {
      post_batch(&entries, false, L"", false);
      entries.clear();
    }
    if (!saw_entry) {
      post_batch(nullptr, true, L"Default file contains no usable entries.", false);
      return;
    }
    post_batch(nullptr, true, L"", false);
  });
}

void MainWindow::StartTraceLoadWorker() {
  if (trace_load_running_.exchange(true)) {
    return;
  }
  trace_load_stop_.store(false);
  LoadTraceSettings();
  std::unordered_map<std::wstring, TraceSelection> selection_cache = trace_selection_cache_;
  std::wstring active_path = ActiveTracesPath();
  if (trace_load_thread_.joinable()) {
    trace_load_thread_.join();
  }
  trace_load_thread_ = std::thread([this, selection_cache = std::move(selection_cache), active_path]() mutable {
    auto payload = std::make_unique<TraceLoadPayload>();
    payload->selection_cache = std::move(selection_cache);
    std::wstring content;
    if (!ReadFileUtf8(active_path, &content)) {
      trace_load_running_.store(false);
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
      if (trace_load_stop_.load()) {
        trace_load_running_.store(false);
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
      std::string buffer;
      if (!ReadFileBinary(source, &buffer)) {
        continue;
      }
      TraceData data;
      if (!this->BuildTraceDataFromBuffer(use_label, source, buffer, &data, nullptr)) {
        continue;
      }
      std::shared_ptr<const TraceData> trace = std::make_shared<TraceData>(std::move(data));
      TraceSelection selection = {};
      selection.select_all = true;
      selection.recursive = true;
      auto it = payload->selection_cache.find(source_lower);
      if (it != payload->selection_cache.end()) {
        selection = it->second;
      }
      NormalizeSelectionForTrace(*trace, &selection);
      payload->selection_cache[source_lower] = selection;
      payload->traces.push_back({trace->label, source, trace, selection});
    }

    if (trace_load_stop_.load()) {
      trace_load_running_.store(false);
      return;
    }
    if (hwnd_ && IsWindow(hwnd_)) {
      PostMessageW(hwnd_, kTraceLoadReadyMessage, 0, reinterpret_cast<LPARAM>(payload.release()));
    }
    trace_load_running_.store(false);
  });
}

void MainWindow::StopTraceLoadWorker() {
  trace_load_stop_.store(true);
  if (trace_load_thread_.joinable()) {
    trace_load_thread_.join();
  }
  trace_load_running_.store(false);
}

void MainWindow::StopTraceParseSessions() {
  for (auto& entry : trace_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->cancel.store(true);
    if (entry.second->thread.joinable()) {
      entry.second->thread.join();
    }
  }
  trace_parse_sessions_.clear();
}

void MainWindow::StartDefaultLoadWorker() {
  if (default_load_running_.exchange(true)) {
    return;
  }
  default_load_stop_.store(false);
  std::wstring active_path = ActiveDefaultsPath();
  if (default_load_thread_.joinable()) {
    default_load_thread_.join();
  }
  default_load_thread_ = std::thread([this, active_path]() {
    auto payload = std::make_unique<DefaultLoadPayload>();
    std::wstring content;
    if (!ReadFileUtf8(active_path, &content)) {
      default_load_running_.store(false);
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
      if (default_load_stop_.load()) {
        default_load_running_.store(false);
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
      DefaultData data;
      if (!this->ParseDefaultRegFile(source, &data, nullptr)) {
        continue;
      }
      std::shared_ptr<const DefaultData> defaults = std::make_shared<DefaultData>(std::move(data));
      KeyValueSelection selection = {};
      selection.select_all = true;
      selection.recursive = true;
      payload->defaults.push_back({use_label, source, defaults, selection});
    }

    if (default_load_stop_.load()) {
      default_load_running_.store(false);
      return;
    }
    if (hwnd_ && IsWindow(hwnd_)) {
      PostMessageW(hwnd_, kDefaultLoadReadyMessage, 0, reinterpret_cast<LPARAM>(payload.release()));
    }
    default_load_running_.store(false);
  });
}

void MainWindow::StopDefaultLoadWorker() {
  default_load_stop_.store(true);
  if (default_load_thread_.joinable()) {
    default_load_thread_.join();
  }
  default_load_running_.store(false);
}

void MainWindow::StopDefaultParseSessions() {
  for (auto& entry : default_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->cancel.store(true);
    if (entry.second->thread.joinable()) {
      entry.second->thread.join();
    }
  }
  default_parse_sessions_.clear();
}

void MainWindow::StopRegFileParseSessions() {
  for (auto& entry : reg_file_parse_sessions_) {
    if (!entry.second) {
      continue;
    }
    entry.second->cancel.store(true);
    if (entry.second->thread.joinable()) {
      entry.second->thread.join();
    }
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
  {
    std::lock_guard<std::mutex> lock(tree_state_mutex_);
    tree_state_selected_ = std::move(selected);
    tree_state_expanded_ = std::move(expanded);
    tree_state_dirty_ = true;
  }
  tree_state_cv_.notify_one();
}

void MainWindow::SaveTreeStateFile(const std::wstring& selected, const std::vector<std::wstring>& expanded) const {
  std::wstring path = TreeStatePath();
  if (path.empty()) {
    return;
  }

  std::wstring content;
  if (!selected.empty()) {
    content += L"selected=";
    content += EscapeHistoryField(selected);
    content.push_back(L'\n');
  }
  for (const auto& entry : expanded) {
    if (entry.empty()) {
      continue;
    }
    content += L"expanded=";
    content += EscapeHistoryField(entry);
    content.push_back(L'\n');
  }
  std::string utf8 = util::WideToUtf8(content);
  if (utf8.empty()) {
    return;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return;
  }
  DWORD written = 0;
  WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  CloseHandle(file);
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
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
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

  TraceSelection selection = {};
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
    if (!trace.data) {
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
    content.append(trace.selection.select_all ? L"1\n" : L"0\n");
    content.append(L"recursive=");
    content.append(trace.selection.recursive ? L"1\n" : L"0\n");
    for (const auto& key_path : trace.selection.key_paths) {
      if (key_path.empty()) {
        continue;
      }
      content.append(L"key=");
      content.append(key_path);
      content.push_back(L'\n');
    }
    for (const auto& entry : trace.selection.values_by_key) {
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

void MainWindow::LoadActiveTraces() {
  active_traces_.clear();
  LoadTraceSettings();
  std::wstring path = ActiveTracesPath();
  if (path.empty()) {
    return;
  }
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
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
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
      continue;
    }
    if (line.front() == L'#') {
      continue;
    }
    if (StartsWithInsensitive(line, L"trace=")) {
      line.erase(0, wcslen(L"trace="));
    }
    line = TrimWhitespace(line);
    if (line.empty()) {
      continue;
    }
    AddTraceFromFile(L"", line, nullptr, false, false);
  }

  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(current_node_);
}

void MainWindow::LoadActiveDefaults() {
  active_defaults_.clear();
  std::wstring path = ActiveDefaultsPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  if (!ReadFileUtf8(path, &content)) {
    return;
  }
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
    if (line.empty()) {
      continue;
    }
    AddDefaultFromFile(L"", line, false, false, false);
  }

  BuildMenus();
  UpdateValueListForNode(current_node_);
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
      session_it->second->cancel.store(true);
      if (session_it->second->thread.joinable()) {
        session_it->second->thread.join();
      }
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
      it->second->cancel.store(true);
      if (it->second->thread.joinable()) {
        it->second->thread.join();
      }
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
      trace_selection_cache_[ToLower(trace.source_path)] = trace.selection;
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

bool MainWindow::HasActiveDefaults() const {
  return !active_defaults_.empty();
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
      session_it->second->cancel.store(true);
      if (session_it->second->thread.joinable()) {
        session_it->second->thread.join();
      }
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

bool MainWindow::RemoveDefaultByLabel(const std::wstring& label) {
  if (label.empty()) {
    return false;
  }
  for (auto it = default_parse_sessions_.begin(); it != default_parse_sessions_.end();) {
    if (it->second && _wcsicmp(it->second->label.c_str(), label.c_str()) == 0) {
      it->second->cancel.store(true);
      if (it->second->thread.joinable()) {
        it->second->thread.join();
      }
      it = default_parse_sessions_.erase(it);
      continue;
    }
    ++it;
  }
  size_t removed = 0;
  active_defaults_.erase(std::remove_if(active_defaults_.begin(), active_defaults_.end(),
                                        [&](const ActiveDefault& defaults) {
                                          if (_wcsicmp(defaults.label.c_str(), label.c_str()) != 0) {
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

bool MainWindow::IsProcessElevated() const {
  return util::IsProcessElevated();
}

bool MainWindow::IsProcessSystem() const {
  return util::IsProcessSystem();
}

bool MainWindow::IsProcessTrustedInstaller() const {
  return util::IsProcessTrustedInstaller();
}

bool MainWindow::RestartAsAdmin() {
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }
  exe_path.resize(len);
  HINSTANCE result = ShellExecuteW(hwnd_, L"runas", exe_path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  if (reinterpret_cast<INT_PTR>(result) <= 32) {
    ui::ShowError(hwnd_, L"Failed to restart with administrator rights.");
    return false;
  }
  PostMessageW(hwnd_, WM_CLOSE, 0, 0);
  return true;
}

bool MainWindow::RestartAsSystem() {
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }
  exe_path.resize(len);

  if (!IsProcessElevated()) {
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
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return false;
  }
  exe_path.resize(len);

  if (!IsProcessElevated()) {
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
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    return;
  }
  exe_path.resize(len);

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
  result = RegQueryValueExW(base, L"Debugger", nullptr, &type, reinterpret_cast<LPBYTE>(debugger.data()), &size);
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
    wchar_t buffer[MAX_PATH * 2] = {};
    DWORD expanded_len = ExpandEnvironmentStringsW(debugger.c_str(), buffer, static_cast<DWORD>(_countof(buffer)));
    if (expanded_len > 0 && expanded_len < _countof(buffer)) {
      expanded.assign(buffer, expanded_len - 1);
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
  std::wstring exe_path;
  exe_path.resize(MAX_PATH);
  DWORD len = GetModuleFileNameW(nullptr, exe_path.data(), static_cast<DWORD>(exe_path.size()));
  if (len == 0 || len >= exe_path.size()) {
    ui::ShowError(hwnd_, L"Failed to locate the executable path.");
    return;
  }
  exe_path.resize(len);

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
  if (!IsProcessElevated() && !IsProcessSystem() && !IsProcessTrustedInstaller()) {
    ui::ShowError(hwnd_, L"Administrator rights are required to open the default Regedit.");
    return false;
  }

  const wchar_t* key_path = L"Software\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\regedit.exe";
  util::UniqueHKey key;
  LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, key_path, 0, KEY_READ | KEY_WRITE, key.put());
  if (result == ERROR_FILE_NOT_FOUND) {
    HINSTANCE launched = ShellExecuteW(hwnd_, L"open", L"regedit.exe", nullptr, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(launched) <= 32) {
      ui::ShowError(hwnd_, L"Failed to open Regedit.");
      return false;
    }
    return true;
  }
  if (result != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(result));
    return false;
  }

  DWORD type = 0;
  DWORD size = 0;
  result = RegQueryValueExW(key.get(), L"Debugger", nullptr, &type, nullptr, &size);
  if (result == ERROR_FILE_NOT_FOUND) {
    HINSTANCE launched = ShellExecuteW(hwnd_, L"open", L"regedit.exe", nullptr, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(launched) <= 32) {
      ui::ShowError(hwnd_, L"Failed to open Regedit.");
      return false;
    }
    return true;
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

  HINSTANCE launched = ShellExecuteW(hwnd_, L"open", L"regedit.exe", nullptr, nullptr, SW_SHOWNORMAL);

  LONG restore = RegSetValueExW(key.get(), L"Debugger", 0, type, data.data(), size);
  RegDeleteValueW(key.get(), temp_name.c_str());
  if (restore != ERROR_SUCCESS) {
    ui::ShowError(hwnd_, FormatWin32Error(restore));
    return false;
  }
  if (reinterpret_cast<INT_PTR>(launched) <= 32) {
    ui::ShowError(hwnd_, L"Failed to open Regedit.");
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
  wchar_t folder[MAX_PATH] = {};
  wcsncpy_s(folder, hive_path.c_str(), _TRUNCATE);
  if (SUCCEEDED(PathCchRemoveFileSpec(folder, _countof(folder)))) {
    ShellExecuteW(hwnd_, L"open", L"explorer.exe", args.c_str(), folder, SW_SHOWNORMAL);
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
  HFONT font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
  if (font && GetObjectW(font, sizeof(lf), &lf) > 0) {
    wcsncpy_s(lf.lfFaceName, face.c_str(), _TRUNCATE);
    return lf;
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
      *selected_path = RegistryProvider::BuildPath(*node);
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
            expanded_paths->push_back(RegistryProvider::BuildPath(*node));
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
  std::vector<std::wstring> expanded;
  expanded.reserve(saved_tree_expanded_paths_.size());
  std::unordered_set<std::wstring> seen;
  seen.reserve(saved_tree_expanded_paths_.size());
  for (const auto& path : saved_tree_expanded_paths_) {
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
  if (!saved_tree_selected_path_.empty()) {
    SelectTreePath(saved_tree_selected_path_);
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
  std::vector<std::wstring> parts = SplitPath(path);
  if (parts.empty()) {
    return false;
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
      parts.front() = kStandardGroupLabel;
    } else {
      parts.front() = kRealGroupLabel;
    }
  } else if (!parts.empty() && EqualsInsensitive(parts.front(), L"Real Registry")) {
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
  if (!parts.empty()) {
    if (!EqualsInsensitive(parts.front(), kStandardGroupLabel) && !EqualsInsensitive(parts.front(), kRealGroupLabel)) {
      if (EqualsInsensitive(parts.front(), L"REGISTRY")) {
        parts.insert(parts.begin(), kRealGroupLabel);
      } else {
        parts.insert(parts.begin(), kStandardGroupLabel);
      }
    }
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

void MainWindow::PushUndo(UndoOperation operation) {
  if (is_replaying_) {
    return;
  }
  undo_stack_.push_back(std::move(operation));
  ClearRedo();
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, undo_stack_.empty() ? 0 : TBSTATE_ENABLED);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, redo_stack_.empty() ? 0 : TBSTATE_ENABLED);
  }
}

void MainWindow::ClearRedo() {
  redo_stack_.clear();
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, undo_stack_.empty() ? 0 : TBSTATE_ENABLED);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, TBSTATE_ENABLED * (!redo_stack_.empty()));
  }
}

bool MainWindow::ApplyUndoOperation(const UndoOperation& operation, bool redo) {
  if (!current_node_) {
    return false;
  }
  bool ok = false;
  is_replaying_ = true;
  switch (operation.type) {
  case UndoOperation::Type::kCreateKey: {
    if (redo) {
      if (!operation.key_snapshot.name.empty()) {
        ok = RestoreKeySnapshot(operation.node, operation.key_snapshot);
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
  case UndoOperation::Type::kDeleteKey: {
    if (redo) {
      RegistryNode child = MakeChildNode(operation.node, operation.name);
      ok = RegistryProvider::DeleteKey(child);
      if (ok) {
        RefreshTreeSelection();
      }
    } else {
      ok = RestoreKeySnapshot(operation.node, operation.key_snapshot);
      if (ok) {
        RefreshTreeSelection();
      }
    }
    break;
  }
  case UndoOperation::Type::kRenameKey: {
    std::wstring from = redo ? operation.name : operation.new_name;
    std::wstring to = redo ? operation.new_name : operation.name;
    RegistryNode child = MakeChildNode(operation.node, from);
    ok = RegistryProvider::RenameKey(child, to);
    if (ok) {
      RefreshTreeSelection();
      std::wstring path = RegistryProvider::BuildPath(operation.node);
      if (!path.empty()) {
        path.append(L"\\");
        path.append(to);
        SelectTreePath(path);
      }
    }
    break;
  }
  case UndoOperation::Type::kCreateValue: {
    if (redo) {
      ok = RegistryProvider::SetValue(operation.node, operation.new_value.name, operation.new_value.type, operation.new_value.data);
    } else {
      ok = RegistryProvider::DeleteValue(operation.node, operation.name);
    }
    break;
  }
  case UndoOperation::Type::kDeleteValue: {
    if (redo) {
      ok = RegistryProvider::DeleteValue(operation.node, operation.old_value.name);
    } else {
      ok = RegistryProvider::SetValue(operation.node, operation.old_value.name, operation.old_value.type, operation.old_value.data);
    }
    break;
  }
  case UndoOperation::Type::kModifyValue: {
    const ValueEntry& value = redo ? operation.new_value : operation.old_value;
    ok = RegistryProvider::SetValue(operation.node, value.name, value.type, value.data);
    break;
  }
  case UndoOperation::Type::kRenameValue: {
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
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, undo_stack_.empty() ? 0 : TBSTATE_ENABLED);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, redo_stack_.empty() ? 0 : TBSTATE_ENABLED);
  }
  return ok;
}

MainWindow::KeySnapshot MainWindow::CaptureKeySnapshot(const RegistryNode& node) {
  KeySnapshot snapshot;
  snapshot.name = LeafName(node);
  snapshot.values = RegistryProvider::EnumValues(node);
  auto children = RegistryProvider::EnumSubKeyNames(node, false);
  snapshot.children.reserve(children.size());
  for (const auto& child_name : children) {
    RegistryNode child = MakeChildNode(node, child_name);
    snapshot.children.push_back(CaptureKeySnapshot(child));
  }
  return snapshot;
}

bool MainWindow::RestoreKeySnapshot(const RegistryNode& parent, const KeySnapshot& snapshot) {
  if (snapshot.name.empty()) {
    return false;
  }
  if (!RegistryProvider::CreateKey(parent, snapshot.name)) {
    return false;
  }
  RegistryNode node = MakeChildNode(parent, snapshot.name);
  for (const auto& value : snapshot.values) {
    if (!RegistryProvider::SetValue(node, value.name, value.type, value.data)) {
      return false;
    }
  }
  for (const auto& child : snapshot.children) {
    if (!RestoreKeySnapshot(node, child)) {
      return false;
    }
  }
  return true;
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
  auto values = RegistryProvider::EnumValues(node);
  auto exists = [&](const std::wstring& candidate) -> bool {
    for (const auto& value : values) {
      if (EqualsInsensitive(value.name, candidate)) {
        return true;
      }
    }
    return false;
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
  std::vector<std::wstring> parts = SplitPath(path);
  if (parts.empty()) {
    return false;
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
      parts.front() = kStandardGroupLabel;
    } else {
      parts.front() = kRealGroupLabel;
    }
  } else if (!parts.empty() && EqualsInsensitive(parts.front(), L"Real Registry")) {
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
  if (!parts.empty()) {
    if (!EqualsInsensitive(parts.front(), kStandardGroupLabel) && !EqualsInsensitive(parts.front(), kRealGroupLabel)) {
      if (EqualsInsensitive(parts.front(), L"REGISTRY")) {
        parts.insert(parts.begin(), kRealGroupLabel);
      } else {
        parts.insert(parts.begin(), kStandardGroupLabel);
      }
    }
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

  auto matches = [&](HTREEITEM item) -> bool {
    wchar_t text[256] = {};
    TVITEMW tvi = {};
    tvi.hItem = item;
    tvi.mask = TVIF_TEXT;
    tvi.pszText = text;
    tvi.cchTextMax = static_cast<int>(_countof(text));
    if (!TreeView_GetItem(tree_.hwnd(), &tvi)) {
      return false;
    }
    return StartsWithInsensitive(text, type_buffer_tree_);
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
      if (matches(items[idx])) {
        return items[idx];
      }
    }
    return nullptr;
  };

  HTREEITEM target = nullptr;
  std::vector<HTREEITEM> children;
  ensure_children_loaded(selected);
  collect_children(selected, &children);
  target = find_match(children, selected);

  if (!target) {
    HTREEITEM parent = TreeView_GetParent(tree_.hwnd(), selected);
    if (parent) {
      std::vector<HTREEITEM> siblings;
      ensure_children_loaded(parent);
      collect_children(parent, &siblings);
      target = find_match(siblings, selected);
    }
  }
  if (!target) {
    return;
  }
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
  std::wstring path = RegistryProvider::BuildPath(node);
  std::wstring trace_path = NormalizeTraceKeyPath(path);
  if (trace_path.empty()) {
    trace_path = path;
  }
  return ToLower(trace_path);
}

void MainWindow::NormalizeSelectionForTrace(const TraceData& trace, TraceSelection* selection) const {
  if (!selection || selection->select_all) {
    return;
  }
  auto resolve_key = [&](const std::wstring& key) -> std::wstring {
    if (key.empty()) {
      return key;
    }
    std::wstring lower = ToLower(key);
    auto it = trace.display_to_key.find(lower);
    if (it != trace.display_to_key.end()) {
      return it->second;
    }
    return key;
  };

  std::unordered_map<std::wstring, std::wstring> key_lookup;
  key_lookup.reserve(trace.key_paths.size());
  for (const auto& key_path : trace.key_paths) {
    key_lookup.emplace(ToLower(key_path), key_path);
  }

  std::vector<std::wstring> normalized_keys;
  std::unordered_set<std::wstring> seen_keys;
  for (const auto& path : selection->key_paths) {
    std::wstring resolved = resolve_key(path);
    std::wstring lower = ToLower(resolved);
    auto it = key_lookup.find(lower);
    if (it == key_lookup.end()) {
      continue;
    }
    if (seen_keys.insert(lower).second) {
      normalized_keys.push_back(it->second);
    }
  }

  std::unordered_map<std::wstring, std::unordered_set<std::wstring>> normalized_values;
  for (const auto& entry : selection->values_by_key) {
    std::wstring resolved = resolve_key(entry.first);
    std::wstring key_lower = ToLower(resolved);
    if (key_lower.empty()) {
      continue;
    }
    auto& values = normalized_values[key_lower];
    for (const auto& value_lower : entry.second) {
      values.insert(value_lower);
    }
  }

  selection->key_paths = std::move(normalized_keys);
  selection->values_by_key = std::move(normalized_values);
  if (selection->key_paths.empty() && selection->values_by_key.empty()) {
    selection->select_all = true;
  }
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
    if (!SelectionIncludesKey(trace.selection, key_lower)) {
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
  std::wstring assets = util::JoinPath(module_dir, L"assets");
  std::wstring traces = util::JoinPath(assets, L"traces");
  return util::JoinPath(traces, file);
}

bool MainWindow::LoadBundledTrace(const std::wstring& label, const TraceSelection* selection_override) {
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

bool MainWindow::LoadBundledDefault(const std::wstring& label) {
  std::wstring path = ResolveBundledDefaultPath(label);
  if (path.empty()) {
    return false;
  }
  return LoadDefaultFromFile(label, path);
}

bool MainWindow::ParseDefaultRegFile(const std::wstring& path, DefaultData* out, std::wstring* error) const {
  if (!out) {
    return false;
  }
  out->values_by_key.clear();
  std::wstring content;
  if (!ReadRegFileText(path, &content)) {
    if (error) {
      *error = L"Failed to read registry file.";
    }
    return false;
  }

  std::vector<std::wstring> lines;
  std::wstring current;
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
    if (current.empty()) {
      current = line;
    } else {
      current.append(line);
    }
    std::wstring trimmed_right = current;
    while (!trimmed_right.empty() && (trimmed_right.back() == L' ' || trimmed_right.back() == L'\t')) {
      trimmed_right.pop_back();
    }
    if (!trimmed_right.empty() && trimmed_right.back() == L'\\') {
      trimmed_right.pop_back();
      current = trimmed_right;
      continue;
    }
    lines.push_back(current);
    current.clear();
  }
  if (!current.empty()) {
    lines.push_back(current);
  }

  std::wstring current_key;
  for (const auto& raw : lines) {
    std::wstring line = TrimWhitespace(raw);
    if (line.empty()) {
      continue;
    }
    if (line[0] == L';') {
      continue;
    }
    if (line.front() == L'[' && line.back() == L']') {
      std::wstring key = line.substr(1, line.size() - 2);
      key = TrimWhitespace(key);
      bool delete_key = !key.empty() && key.front() == L'-';
      if (delete_key) {
        current_key.clear();
        continue;
      }
      std::wstring normalized = NormalizeTraceKeyPathBasic(key);
      current_key = normalized.empty() ? key : normalized;
      if (!current_key.empty()) {
        std::wstring key_lower = ToLower(current_key);
        out->values_by_key.emplace(key_lower, DefaultKeyValues());
      }
      continue;
    }
    if (current_key.empty()) {
      continue;
    }
    size_t eq = line.find(L'=');
    if (eq == std::wstring::npos) {
      continue;
    }
    std::wstring name_part = TrimWhitespace(line.substr(0, eq));
    std::wstring data_part = TrimWhitespace(line.substr(eq + 1));
    if (name_part.empty() || data_part.empty()) {
      continue;
    }
    if (data_part == L"-") {
      continue;
    }

    std::wstring value_name;
    if (name_part == L"@") {
      value_name.clear();
    } else if (name_part.front() == L'"') {
      size_t end_pos = 0;
      if (!ParseQuotedString(name_part, &value_name, &end_pos)) {
        continue;
      }
    } else {
      continue;
    }

    DWORD type = REG_NONE;
    std::vector<BYTE> data;
    if (data_part.front() == L'"') {
      std::wstring text;
      size_t end_pos = 0;
      if (!ParseQuotedString(data_part, &text, &end_pos)) {
        continue;
      }
      type = REG_SZ;
      data = StringToRegData(text);
    } else if (StartsWithInsensitive(data_part, L"dword:")) {
      std::wstring hex = TrimWhitespace(data_part.substr(6));
      if (hex.empty()) {
        continue;
      }
      DWORD number = static_cast<DWORD>(wcstoul(hex.c_str(), nullptr, 16));
      type = REG_DWORD;
      data.resize(sizeof(DWORD));
      memcpy(data.data(), &number, sizeof(DWORD));
    } else if (StartsWithInsensitive(data_part, L"hex")) {
      type = REG_BINARY;
      size_t colon = data_part.find(L':');
      if (colon == std::wstring::npos) {
        continue;
      }
      size_t open = data_part.find(L'(');
      size_t close = data_part.find(L')');
      if (open != std::wstring::npos && close != std::wstring::npos && close > open) {
        std::wstring code = data_part.substr(open + 1, close - open - 1);
        unsigned long parsed = wcstoul(code.c_str(), nullptr, 16);
        switch (parsed) {
        case 0x0:
          type = REG_NONE;
          break;
        case 0x1:
          type = REG_SZ;
          break;
        case 0x2:
          type = REG_EXPAND_SZ;
          break;
        case 0x3:
          type = REG_BINARY;
          break;
        case 0x4:
          type = REG_DWORD;
          break;
        case 0x5:
          type = REG_DWORD_BIG_ENDIAN;
          break;
        case 0x7:
          type = REG_MULTI_SZ;
          break;
        case 0x8:
          type = REG_RESOURCE_LIST;
          break;
        case 0x9:
          type = REG_FULL_RESOURCE_DESCRIPTOR;
          break;
        case 0xA:
          type = REG_RESOURCE_REQUIREMENTS_LIST;
          break;
        case 0xB:
          type = REG_QWORD;
          break;
        default:
          type = REG_BINARY;
          break;
        }
      }
      std::wstring hex = data_part.substr(colon + 1);
      if (!ParseHexBytes(hex, &data)) {
        continue;
      }
    } else {
      continue;
    }

    std::wstring key_lower = ToLower(current_key);
    auto it = out->values_by_key.find(key_lower);
    if (it == out->values_by_key.end()) {
      continue;
    }
    std::wstring name_lower = ToLower(value_name);
    DefaultValueEntry entry;
    entry.type = type;
    entry.data = RegistryProvider::FormatValueDataForDisplay(type, data.empty() ? nullptr : data.data(), static_cast<DWORD>(data.size()));
    it->second.values[name_lower] = std::move(entry);
  }
  if (out->values_by_key.empty()) {
    if (error) {
      *error = L"Default file contains no usable entries.";
    }
    return false;
  }
  return true;
}

bool MainWindow::BuildTraceDataFromBuffer(const std::wstring& label, const std::wstring& source, const std::string& buffer, TraceData* out_data, std::wstring* error) const {
  if (!out_data) {
    return false;
  }
  if (buffer.empty()) {
    if (error) {
      *error = L"Trace file is empty or too large to load.";
    }
    return false;
  }
  std::string text = buffer;
  if (text.size() >= 3 && static_cast<unsigned char>(text[0]) == 0xEF && static_cast<unsigned char>(text[1]) == 0xBB && static_cast<unsigned char>(text[2]) == 0xBF) {
    text.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(text);
  if (content.empty()) {
    if (error) {
      *error = L"Trace file has no readable entries.";
    }
    return false;
  }

  TraceData data;
  data.label = label;
  data.source_path = source;
  std::unordered_map<std::wstring, std::wstring> key_by_lower;
  std::unordered_set<std::wstring> display_lower;
  std::unordered_map<std::wstring, std::wstring> display_to_key;

  auto add_display_path = [&](const std::wstring& path, const std::wstring& key_path) {
    if (path.empty()) {
      return;
    }
    std::wstring lower = ToLower(path);
    if (display_lower.insert(lower).second) {
      data.display_key_paths.push_back(path);
    }
    if (!key_path.empty()) {
      display_to_key.emplace(lower, key_path);
    }
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
      continue;
    }
    size_t sep = line.rfind(L" : ");
    size_t sep_len = 0;
    if (sep != std::wstring::npos) {
      sep_len = 3;
    } else {
      sep = line.rfind(L':');
      sep_len = (sep == std::wstring::npos) ? 0 : 1;
    }
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key_text = TrimWhitespace(line.substr(0, sep));
    std::wstring value_text = TrimWhitespace(line.substr(sep + sep_len));
    if (key_text.empty()) {
      continue;
    }
    std::wstring selection_path = NormalizeTraceSelectionPath(key_text);
    if (selection_path.empty()) {
      continue;
    }
    std::wstring key_path = NormalizeTraceKeyPath(key_text);
    if (key_path.empty()) {
      key_path = selection_path;
    }
    std::wstring key_lower = ToLower(key_path);
    key_by_lower.emplace(key_lower, key_path);
    add_display_path(selection_path, key_path);

    std::wstring value_name = value_text;
    if (EqualsInsensitive(value_name, L"(Default)")) {
      value_name.clear();
    }
    std::wstring value_lower = ToLower(value_name);
    auto& entry = data.values_by_key[key_lower];
    if (entry.values_lower.insert(value_lower).second) {
      entry.values_display.push_back(value_name);
    }
  }

  if (key_by_lower.empty()) {
    if (error) {
      *error = L"Trace file contains no usable entries.";
    }
    return false;
  }
  data.key_paths.reserve(key_by_lower.size());
  for (const auto& item : key_by_lower) {
    data.key_paths.push_back(item.second);
  }
  data.display_to_key = std::move(display_to_key);
  std::sort(data.key_paths.begin(), data.key_paths.end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
  std::sort(data.display_key_paths.begin(), data.display_key_paths.end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });

  std::unordered_map<std::wstring, std::unordered_map<std::wstring, std::wstring>> child_map;
  child_map.reserve(data.key_paths.size());
  for (const auto& key_path : data.key_paths) {
    std::vector<std::wstring> parts = SplitPath(key_path);
    if (parts.size() < 2) {
      continue;
    }
    std::wstring current = parts.front();
    for (size_t i = 1; i < parts.size(); ++i) {
      const std::wstring& child = parts[i];
      std::wstring parent_lower = ToLower(current);
      std::wstring child_lower = ToLower(child);
      auto& children = child_map[parent_lower];
      if (children.find(child_lower) == children.end()) {
        children.emplace(child_lower, child);
      }
      current.append(L"\\");
      current.append(child);
    }
  }
  data.children_by_key.clear();
  data.children_by_key.reserve(child_map.size());
  for (auto& entry : child_map) {
    std::vector<std::wstring> children;
    children.reserve(entry.second.size());
    for (auto& child : entry.second) {
      children.push_back(child.second);
    }
    std::sort(children.begin(), children.end(), [](const std::wstring& left, const std::wstring& right) { return _wcsicmp(left.c_str(), right.c_str()) < 0; });
    data.children_by_key.emplace(entry.first, std::move(children));
  }

  *out_data = std::move(data);
  return true;
}

bool MainWindow::AddTraceFromBuffer(const std::wstring& label, const std::wstring& source, const std::string& buffer, const TraceSelection* selection_override, bool prompt_for_selection) {
  TraceData data;
  std::wstring error;
  if (!BuildTraceDataFromBuffer(label, source, buffer, &data, &error)) {
    ui::ShowError(hwnd_, error.empty() ? L"Failed to load trace file." : error.c_str());
    return false;
  }

  std::shared_ptr<const TraceData> trace = std::make_shared<TraceData>(std::move(data));
  TraceSelection selection = {};
  selection.select_all = true;
  selection.recursive = true;
  if (selection_override) {
    selection = *selection_override;
  }
  if (prompt_for_selection) {
    TraceDialogOptions options;
    options.title = trace->label.empty() ? L"Trace entries" : L"Trace entries - " + trace->label;
    options.prompt = L"";
    options.show_values = true;

    std::unordered_map<std::wstring, std::wstring> key_lookup;
    key_lookup.reserve(trace->key_paths.size());
    for (const auto& key_path : trace->key_paths) {
      key_lookup.emplace(ToLower(key_path), key_path);
    }
    std::unordered_map<std::wstring, std::wstring> display_lookup;
    display_lookup.reserve(trace->display_key_paths.size());
    for (const auto& display_path : trace->display_key_paths) {
      std::wstring display_lower = ToLower(display_path);
      auto it = trace->display_to_key.find(display_lower);
      if (it == trace->display_to_key.end()) {
        continue;
      }
      std::wstring key_lower = ToLower(it->second);
      display_lookup.emplace(key_lower, display_path);
    }

    auto entries = std::make_unique<std::vector<KeyValueDialogEntry>>();
    for (const auto& entry : trace->values_by_key) {
      std::wstring key_lower = entry.first;
      std::wstring key_path = key_lower;
      auto key_it = key_lookup.find(key_lower);
      if (key_it != key_lookup.end()) {
        key_path = key_it->second;
      }
      std::wstring display_path = key_path;
      auto display_it = display_lookup.find(key_lower);
      if (display_it != display_lookup.end()) {
        display_path = display_it->second;
      }
      for (const auto& value_name : entry.second.values_display) {
        KeyValueDialogEntry dialog_entry;
        dialog_entry.key_path = key_path;
        dialog_entry.display_path = display_path;
        dialog_entry.has_value = true;
        dialog_entry.value_name = value_name;
        entries->push_back(std::move(dialog_entry));
      }
    }

    struct DialogEntryContext {
      std::unique_ptr<std::vector<KeyValueDialogEntry>> entries;
    };
    DialogEntryContext context;
    context.entries = std::move(entries);
    auto on_ready = [](HWND dialog, void* context_ptr) {
      auto* ctx = reinterpret_cast<DialogEntryContext*>(context_ptr);
      if (!ctx) {
        return;
      }
      if (ctx->entries) {
        TraceDialogPostEntries(dialog, ctx->entries.release());
      }
      TraceDialogPostDone(dialog, true);
    };

    if (!ShowTraceDialog(hwnd_, options, &selection, on_ready, &context)) {
      return false;
    }
  }
  if (!selection.select_all && selection.key_paths.empty() && selection.values_by_key.empty()) {
    selection.select_all = true;
  }
  NormalizeSelectionForTrace(*trace, &selection);

  active_traces_.push_back({trace->label, source, trace, selection});
  trace_selection_cache_[ToLower(source)] = selection;
  return true;
}

bool MainWindow::LoadTraceFromBuffer(const std::wstring& label, const std::wstring& source, const std::string& buffer, const TraceSelection* selection_override) {
  if (!AddTraceFromBuffer(label, source, buffer, selection_override, true)) {
    return false;
  }
  SaveActiveTraces();
  SaveTraceSettings();
  BuildMenus();
  RefreshTreeSelection();
  UpdateValueListForNode(current_node_);
  SaveSettings();
  return true;
}

bool MainWindow::AddTraceFromFile(const std::wstring& label, const std::wstring& path, const TraceSelection* selection_override, bool prompt_for_selection, bool update_ui) {
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

  TraceSelection selection = {};
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
  session->data = std::make_shared<TraceData>();
  session->data->label = use_label;
  session->data->source_path = source;
  session->selection = selection;

  TraceParseSession* session_ptr = session.get();
  trace_parse_sessions_.emplace(source_lower, std::move(session));

  if (prompt_for_selection) {
    TraceSelection dialog_selection = selection;
    TraceDialogOptions options;
    options.title = use_label.empty() ? L"Trace entries" : L"Trace entries - " + use_label;
    options.prompt = L"";
    options.show_values = true;
    TraceDialogStartContext context;
    context.window = this;
    context.session = session_ptr;
    if (!ShowTraceDialog(hwnd_, options, &dialog_selection, StartTraceDialogLoad, &context)) {
      session_ptr->cancel.store(true);
      if (session_ptr->thread.joinable()) {
        session_ptr->thread.join();
      }
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
  active_traces_.push_back({use_label, source, session_ptr->data, session_ptr->selection});
  trace_selection_cache_[source_lower] = session_ptr->selection;

  if (update_ui) {
    SaveActiveTraces();
    SaveTraceSettings();
    BuildMenus();
    RefreshTreeSelection();
    UpdateValueListForNode(current_node_);
    SaveSettings();
  }
  if (session_ptr->parsing_done && session_ptr->thread.joinable()) {
    session_ptr->thread.join();
  }
  if (session_ptr->parsing_done && !session_ptr->dialog) {
    trace_parse_sessions_.erase(source_lower);
  }
  return true;
}

bool MainWindow::LoadTraceFromFile(const std::wstring& label, const std::wstring& path, const TraceSelection* selection_override) {
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

  KeyValueSelection selection = {};
  selection.select_all = true;
  selection.recursive = true;

  auto session = std::make_unique<DefaultParseSession>();
  session->label = use_label;
  session->source_path = source;
  session->source_lower = source_lower;
  session->data = std::make_shared<DefaultData>();
  session->selection = selection;
  session->show_errors = show_error;

  DefaultParseSession* session_ptr = session.get();
  default_parse_sessions_.emplace(source_lower, std::move(session));

  if (prompt_for_selection) {
    KeyValueSelection dialog_selection = selection;
    TraceDialogOptions options;
    options.title = use_label.empty() ? L"Default entries" : L"Default entries - " + use_label;
    options.prompt = L"";
    options.show_values = true;
    DefaultDialogStartContext context;
    context.window = this;
    context.session = session_ptr;
    if (!ShowTraceDialog(hwnd_, options, &dialog_selection, StartDefaultDialogLoad, &context)) {
      session_ptr->cancel.store(true);
      if (session_ptr->thread.joinable()) {
        session_ptr->thread.join();
      }
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
  active_defaults_.push_back({use_label, source, session_ptr->data, session_ptr->selection});
  if (update_ui) {
    SaveActiveDefaults();
    BuildMenus();
    UpdateValueListForNode(current_node_);
    SaveSettings();
  }
  if (session_ptr->parsing_done && session_ptr->thread.joinable()) {
    session_ptr->thread.join();
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
  if (!WriteRegFileText(entry.reg_file_path, content)) {
    ui::ShowError(hwnd_, L"Failed to save registry file.");
    return false;
  }
  if (entry.reg_file_dirty) {
    entry.reg_file_dirty = false;
    BuildMenus();
  }
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
  if (!WriteRegFileText(target, content)) {
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

  out->append(L"Windows Registry Editor Version 5.00\r\n");
  std::function<void(const RegistryProvider::VirtualRegistryKey&, const std::wstring&)> append_key;
  append_key = [&](const RegistryProvider::VirtualRegistryKey& key, const std::wstring& full_path) {
    std::vector<const RegistryProvider::VirtualRegistryValue*> values;
    values.reserve(key.values.size());
    for (const auto& entry_value : key.values) {
      values.push_back(&entry_value.second);
    }
    std::sort(values.begin(), values.end(), [](const RegistryProvider::VirtualRegistryValue* left, const RegistryProvider::VirtualRegistryValue* right) {
      if (!left || !right) {
        return left != nullptr;
      }
      bool left_default = left->name.empty();
      bool right_default = right->name.empty();
      if (left_default != right_default) {
        return left_default;
      }
      return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
    });

    if (!values.empty()) {
      out->append(L"\r\n");
      out->append(L"[");
      out->append(full_path);
      out->append(L"]\r\n");
      for (const auto* value : values) {
        if (!value) {
          continue;
        }
        if (value->name.empty()) {
          out->append(L"@=");
        } else {
          out->append(L"\"");
          out->append(EscapeRegString(value->name));
          out->append(L"\"=");
        }
        out->append(FormatRegValueData(value->type, value->data));
        out->append(L"\r\n");
      }
    }

    std::vector<const RegistryProvider::VirtualRegistryKey*> children;
    children.reserve(key.children.size());
    for (const auto& child : key.children) {
      if (child.second) {
        children.push_back(child.second.get());
      }
    }
    std::sort(children.begin(), children.end(), [](const RegistryProvider::VirtualRegistryKey* left, const RegistryProvider::VirtualRegistryKey* right) {
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
    session->thread = std::thread([this, session_ptr, hwnd]() {
      auto payload = std::make_unique<RegFileParsePayload>();
      payload->source_path = session_ptr->source_path;
      payload->source_lower = session_ptr->source_lower;
      std::wstring parse_error;
      std::vector<ParsedRegFileRoot> parsed_roots;
      bool cancelled = false;
      if (!ParseRegFileToVirtualRoots(payload->source_path, &parsed_roots, &parse_error, &session_ptr->cancel, &cancelled)) {
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
      payload.release();
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
  std::vector<std::wstring> cleaned;
  cleaned.reserve(recent_trace_paths_.size());
  for (const auto& entry : recent_trace_paths_) {
    std::wstring path = TrimWhitespace(entry);
    if (path.empty()) {
      continue;
    }
    bool duplicate = false;
    for (const auto& existing : cleaned) {
      if (EqualsInsensitive(existing, path)) {
        duplicate = true;
        break;
      }
    }
    if (!duplicate) {
      cleaned.push_back(std::move(path));
      if (static_cast<int>(cleaned.size()) >= kMaxRecentTraces) {
        break;
      }
    }
  }
  recent_trace_paths_.swap(cleaned);
}

void MainWindow::NormalizeRecentDefaultList() {
  std::vector<std::wstring> cleaned;
  cleaned.reserve(recent_default_paths_.size());
  for (const auto& entry : recent_default_paths_) {
    std::wstring path = TrimWhitespace(entry);
    if (path.empty()) {
      continue;
    }
    bool duplicate = false;
    for (const auto& existing : cleaned) {
      if (EqualsInsensitive(existing, path)) {
        duplicate = true;
        break;
      }
    }
    if (!duplicate) {
      cleaned.push_back(std::move(path));
      if (static_cast<int>(cleaned.size()) >= kMaxRecentDefaults) {
        break;
      }
    }
  }
  recent_default_paths_.swap(cleaned);
}

void MainWindow::AddRecentTracePath(const std::wstring& path) {
  std::wstring trimmed = TrimWhitespace(path);
  if (trimmed.empty()) {
    return;
  }
  auto it = std::find_if(recent_trace_paths_.begin(), recent_trace_paths_.end(), [&](const std::wstring& entry) { return EqualsInsensitive(entry, trimmed); });
  if (it != recent_trace_paths_.end()) {
    recent_trace_paths_.erase(it);
  }
  recent_trace_paths_.insert(recent_trace_paths_.begin(), std::move(trimmed));
  NormalizeRecentTraceList();
}

void MainWindow::AddRecentDefaultPath(const std::wstring& path) {
  std::wstring trimmed = TrimWhitespace(path);
  if (trimmed.empty()) {
    return;
  }
  auto it = std::find_if(recent_default_paths_.begin(), recent_default_paths_.end(), [&](const std::wstring& entry) { return EqualsInsensitive(entry, trimmed); });
  if (it != recent_default_paths_.end()) {
    recent_default_paths_.erase(it);
  }
  recent_default_paths_.insert(recent_default_paths_.begin(), std::move(trimmed));
  NormalizeRecentDefaultList();
}

} // namespace regkit
