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
#include <commctrl.h>
#include <commdlg.h>
#include <cwctype>
#include <mutex>
#include <shellapi.h>
#include <unordered_map>
#include <unordered_set>

#include "../../include/app/command_ids.h"
#include "workspace/favorites.h"
#include "../../include/app/font_dialog.h"
#include "../../include/app/registry_io.h"
#include "../../include/app/theme.h"
#include "../../include/app/ui_helpers.h"
#include "../../include/app/value_dialogs.h"
#include "../../include/registry/registry_provider.h"
#include "search/compare.h"
#include "search/search.h"
#include "../../include/win32/win32_helpers.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/process_rights.h"
#include "win32/shell_paths.h"
#include "../../resources/resource.h"

namespace regkit {

using workspace::FavoritesStore;

namespace {

constexpr wchar_t kRepoUrl[] = L"https://github.com/nohuto/regkit";

using util::ToLower;
using util::TrimWhitespace;

HMENU BuildCopyKeyPathMenu() {
  HMENU menu = CreatePopupMenu();
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathAbbrev, L"Abbreviated (HKLM)");
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathRegedit, L"Regedit Address Bar");
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathRegFile, L".reg File Header");
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathPowerShell, L"PowerShell Drive");
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathPowerShellProvider, L"PowerShell Provider");
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKeyPathEscaped, L"Escaped Backslashes");
  return menu;
}
constexpr wchar_t kOneKeyPerLineText[] = L"Each line should include one key.";

std::vector<std::wstring> SplitLines(const std::wstring& text) {
  std::vector<std::wstring> lines;
  std::wstring current;
  for (wchar_t ch : text) {
    if (ch == L'\r' || ch == L'\n') {
      if (!current.empty()) {
        lines.push_back(current);
        current.clear();
      }
    } else {
      current.push_back(ch);
    }
  }
  if (!current.empty()) {
    lines.push_back(current);
  }
  for (auto& line : lines) {
    line.erase(line.begin(), std::find_if(line.begin(), line.end(), [](wchar_t c) { return c != L' '; }));
    while (!line.empty() && line.back() == L' ') {
      line.pop_back();
    }
  }
  lines.erase(std::remove_if(lines.begin(), lines.end(), [](const std::wstring& line) { return line.empty(); }), lines.end());
  return lines;
}

std::wstring JoinLines(const std::vector<std::wstring>& lines) {
  std::wstring out;
  for (size_t i = 0; i < lines.size(); ++i) {
    if (lines[i].empty()) {
      continue;
    }
    if (!out.empty()) {
      out.append(L"\r\n");
    }
    out.append(lines[i]);
  }
  return out;
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

HPEN GetCachedPen(COLORREF color, int width = 1) {
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

const ListRow* SelectedValueRow(const ValueList& list, int* out_index) {
  if (!list.hwnd()) {
    return nullptr;
  }
  int index = ListView_GetNextItem(list.hwnd(), -1, LVNI_SELECTED);
  if (index < 0) {
    return nullptr;
  }
  if (out_index) {
    *out_index = index;
  }
  return list.RowAt(index);
}

std::vector<ListRow> SelectedListRows(const ValueList& list) {
  std::vector<ListRow> rows;
  if (!list.hwnd()) {
    return rows;
  }
  rows.reserve(static_cast<size_t>(ListView_GetSelectedCount(list.hwnd())));
  int index = -1;
  while ((index = ListView_GetNextItem(list.hwnd(), index, LVNI_SELECTED)) >= 0) {
    const ListRow* row = list.RowAt(index);
    if (row) {
      rows.push_back(*row);
    }
  }
  return rows;
}

bool GetValueEntry(const RegistryNode& node, const std::wstring& name, ValueEntry* out) {
  if (RegistryProvider::QueryValue(node, name, out)) {
    return true;
  }
  if (out && name.empty()) {
    out->name.clear();
    out->type = REG_SZ;
    out->data.clear();
    return true;
  }
  return false;
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

bool SelectValueByName(ValueList& list, const std::wstring& name) {
  for (size_t i = 0; i < list.RowCount(); ++i) {
    const ListRow* row = list.RowAt(static_cast<int>(i));
    if (!row || row->kind != rowkind::kValue) {
      continue;
    }
    if (row->extra == name) {
      ListView_SetItemState(list.hwnd(), static_cast<int>(i), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
      ListView_EnsureVisible(list.hwnd(), static_cast<int>(i), FALSE);
      return true;
    }
  }
  return false;
}

bool PromptOpenFilePath(HWND owner, const wchar_t* filter, std::wstring* path) {
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

bool PromptSaveFilePath(HWND owner, const wchar_t* filter, std::wstring* path) {
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

std::wstring FileNameOnly(const std::wstring& path) {
  size_t pos = path.find_last_of(L"\\/");
  return (pos == std::wstring::npos) ? path : path.substr(pos + 1);
}

std::wstring FileBaseName(const std::wstring& path) {
  std::wstring name = FileNameOnly(path);
  size_t dot = name.find_last_of(L'.');
  if (dot != std::wstring::npos) {
    name = name.substr(0, dot);
  }
  return name;
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
  item.pszText = &(*buffer)[0];
  item.cchTextMax = static_cast<int>(buffer->size());
  int length = static_cast<int>(SendMessageW(list, LVM_GETITEMTEXTW, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&item)));
  if (length >= static_cast<int>(buffer->size() - 1)) {
    buffer->resize(static_cast<size_t>(length) + 2);
    item.pszText = &(*buffer)[0];
    item.cchTextMax = static_cast<int>(buffer->size());
    length = static_cast<int>(SendMessageW(list, LVM_GETITEMTEXTW, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&item)));
  }
  return length;
}

bool GetListViewColumnInfo(HWND list, int display_index, int* subitem, int* width) {
  if (subitem) {
    *subitem = -1;
  }
  if (width) {
    *width = 0;
  }
  if (!list || display_index < 0) {
    return false;
  }
  LVCOLUMNW col = {};
  col.mask = LVCF_SUBITEM | LVCF_WIDTH;
  if (!ListView_GetColumn(list, display_index, &col)) {
    return false;
  }
  if (subitem) {
    *subitem = col.iSubItem;
  }
  if (width) {
    *width = col.cx;
  }
  return true;
}

std::wstring BuildSelectedListViewText(HWND list) {
  if (!list) {
    return L"";
  }
  HWND header = ListView_GetHeader(list);
  int columns = header ? Header_GetItemCount(header) : 0;
  std::vector<int> subitems;
  subitems.reserve(columns);
  for (int i = 0; i < columns; ++i) {
    int subitem = -1;
    int width = 0;
    if (!GetListViewColumnInfo(list, i, &subitem, &width)) {
      continue;
    }
    if (width <= 0 || subitem < 0) {
      continue;
    }
    subitems.push_back(subitem);
  }
  if (subitems.empty()) {
    return L"";
  }

  std::wstring output;
  std::wstring buffer(256, L'\0');
  int index = -1;
  bool first_row = true;
  while ((index = ListView_GetNextItem(list, index, LVNI_SELECTED)) >= 0) {
    if (!first_row) {
      output.append(L"\r\n");
    }
    first_row = false;
    for (size_t i = 0; i < subitems.size(); ++i) {
      if (i > 0) {
        output.append(L"\t");
      }
      buffer.assign(256, L'\0');
      int length = FetchListViewItemText(list, index, subitems[i], &buffer);
      if (length > 0) {
        output.append(buffer.c_str(), static_cast<size_t>(length));
      }
    }
  }
  return output;
}

bool EqualsInsensitive(const std::wstring& left, const std::wstring& right) {
  return _wcsicmp(left.c_str(), right.c_str()) == 0;
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

bool StartsWithInsensitive(const std::wstring& text, const std::wstring& prefix) {
  if (prefix.empty()) {
    return true;
  }
  if (text.size() < prefix.size()) {
    return false;
  }
  return _wcsnicmp(text.c_str(), prefix.c_str(), prefix.size()) == 0;
}

enum class CompareSourceType {
  kRegistry = 0,
  kRegFile = 1,
};

struct CompareDialogSelection {
  CompareSourceType type = CompareSourceType::kRegistry;
  std::wstring root;
  std::wstring path;
  std::wstring file_path;
  std::wstring key_path;
  bool recursive = true;
};

struct CompareDialogDefaults {
  std::vector<std::wstring> registry_roots;
  CompareDialogSelection left;
  CompareDialogSelection right;
};

struct CompareDialogResult {
  CompareDialogSelection left;
  CompareDialogSelection right;
};

struct CompareDialogState {
  CompareDialogDefaults data;
  HFONT ui_font = nullptr;
};

struct EditBorderState {
  bool hot = false;
  UINT dpi = 0;
  int x_edge = 1;
  int y_edge = 1;
  int x_scroll = 0;
  int y_scroll = 0;
};

using GetSystemMetricsForDpiFn = int(WINAPI*)(int, UINT);

int GetMetricForDpi(int index, UINT dpi) {
  static GetSystemMetricsForDpiFn get_for_dpi = []() -> GetSystemMetricsForDpiFn {
    HMODULE user32 = GetModuleHandleW(L"user32.dll");
    if (!user32) {
      return nullptr;
    }
    return reinterpret_cast<GetSystemMetricsForDpiFn>(GetProcAddress(user32, "GetSystemMetricsForDpi"));
  }();
  if (get_for_dpi) {
    return get_for_dpi(index, dpi);
  }
  int value = GetSystemMetrics(index);
  return MulDiv(value, static_cast<int>(dpi), 96);
}

void UpdateEditBorderMetrics(HWND hwnd, EditBorderState* state, UINT dpi_override = 0) {
  if (!state) {
    return;
  }
  UINT dpi = dpi_override ? dpi_override : (hwnd ? GetDpiForWindow(hwnd) : 96);
  state->dpi = dpi;
  state->x_edge = GetMetricForDpi(SM_CXEDGE, dpi);
  state->y_edge = GetMetricForDpi(SM_CYEDGE, dpi);
  state->x_scroll = GetMetricForDpi(SM_CXVSCROLL, dpi);
  state->y_scroll = GetMetricForDpi(SM_CYVSCROLL, dpi);
  if (state->x_edge < 1) {
    state->x_edge = 1;
  }
  if (state->y_edge < 1) {
    state->y_edge = 1;
  }
}

LRESULT CALLBACK EditBorderSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR data) {
  auto* state = reinterpret_cast<EditBorderState*>(data);
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, EditBorderSubclassProc, id);
    delete state;
    break;
  case WM_NCCALCSIZE: {
    UpdateEditBorderMetrics(hwnd, state);
    int x_edge = state ? state->x_edge : 1;
    int y_edge = state ? state->y_edge : 1;
    if (wparam) {
      auto* params = reinterpret_cast<NCCALCSIZE_PARAMS*>(lparam);
      InflateRect(&params->rgrc[0], -x_edge, -y_edge);
      return 0;
    }
    auto* rect = reinterpret_cast<RECT*>(lparam);
    InflateRect(rect, -x_edge, -y_edge);
    return 0;
  }
  case WM_NCPAINT: {
    LRESULT result = DefSubclassProc(hwnd, msg, wparam, lparam);
    HDC hdc = GetWindowDC(hwnd);
    if (!hdc) {
      return result;
    }
    UpdateEditBorderMetrics(hwnd, state);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    if (state) {
      rect.right += 2 * state->x_edge;
      rect.bottom += 2 * state->y_edge;
      LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
      if ((style & WS_VSCROLL) == WS_VSCROLL) {
        rect.right += state->x_scroll;
      }
      if ((style & WS_HSCROLL) == WS_HSCROLL) {
        rect.bottom += state->y_scroll;
      }
    }

    const Theme& theme = Theme::Current();
    RECT inner = rect;
    InflateRect(&inner, -1, -1);
    HPEN inner_pen = GetCachedPen(theme.BackgroundColor(), 1);
    HGDIOBJ old_pen = SelectObject(hdc, inner_pen);
    HGDIOBJ old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
    Rectangle(hdc, inner.left, inner.top, inner.right, inner.bottom);
    SelectObject(hdc, old_brush);
    SelectObject(hdc, old_pen);

    bool enabled = IsWindowEnabled(hwnd) != FALSE;
    COLORREF border = theme.BorderColor();
    if (enabled) {
      if (GetFocus() == hwnd) {
        border = theme.FocusColor();
      } else if (state && state->hot) {
        border = theme.HoverColor();
      }
    }
    HPEN pen = GetCachedPen(border, 1);
    old_pen = SelectObject(hdc, pen);
    old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
    Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
    SelectObject(hdc, old_brush);
    SelectObject(hdc, old_pen);
    ReleaseDC(hwnd, hdc);
    return 0;
  }
  case WM_MOUSEMOVE: {
    if (state && !state->hot) {
      state->hot = true;
      TRACKMOUSEEVENT tme = {};
      tme.cbSize = sizeof(tme);
      tme.dwFlags = TME_LEAVE;
      tme.hwndTrack = hwnd;
      TrackMouseEvent(&tme);
      SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    }
    break;
  }
  case WM_MOUSELEAVE:
    if (state && state->hot) {
      state->hot = false;
      SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    }
    break;
  case WM_SETFOCUS:
  case WM_KILLFOCUS:
  case WM_ENABLE:
    RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME);
    break;
  case WM_DPICHANGED:
  case WM_DPICHANGED_AFTERPARENT:
    UpdateEditBorderMetrics(hwnd, state, (msg == WM_DPICHANGED) ? LOWORD(wparam) : 0);
    SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    break;
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

void ApplyEditCustomBorder(HWND parent, int id) {
  HWND ctrl = GetDlgItem(parent, id);
  if (!ctrl) {
    return;
  }
  LONG_PTR ex = GetWindowLongPtrW(ctrl, GWL_EXSTYLE);
  if (ex & WS_EX_CLIENTEDGE) {
    ex &= ~static_cast<LONG_PTR>(WS_EX_CLIENTEDGE);
    SetWindowLongPtrW(ctrl, GWL_EXSTYLE, ex);
  }
  LONG_PTR style = GetWindowLongPtrW(ctrl, GWL_STYLE);
  if (style & WS_BORDER) {
    style &= ~static_cast<LONG_PTR>(WS_BORDER);
    SetWindowLongPtrW(ctrl, GWL_STYLE, style);
  }
  if (!GetWindowSubclass(ctrl, EditBorderSubclassProc, 1, nullptr)) {
    auto* state = new EditBorderState();
    if (!SetWindowSubclass(ctrl, EditBorderSubclassProc, 1, reinterpret_cast<DWORD_PTR>(state))) {
      delete state;
    }
  }
  SetWindowPos(ctrl, nullptr, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

std::vector<std::wstring> ExtractRegFileKeys(
    const regfile::Document& data) {
  std::vector<std::wstring> keys = data.key_order;
  if (keys.empty()) {
    keys.reserve(data.keys.size());
    for (const auto& entry : data.keys) {
      keys.push_back(entry.second.path);
    }
  }
  std::sort(keys.begin(), keys.end(), [](const std::wstring& a, const std::wstring& b) { return _wcsicmp(a.c_str(), b.c_str()) < 0; });
  return keys;
}

void ApplyDialogFonts(HWND hwnd, HFONT font) {
  if (!font) {
    return;
  }
  SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
  EnumChildWindows(
      hwnd,
      [](HWND child, LPARAM param) -> BOOL {
        HFONT font_handle = reinterpret_cast<HFONT>(param);
        SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(font_handle), TRUE);
        return TRUE;
      },
      reinterpret_cast<LPARAM>(font));
}

void CenterDialogToOwner(HWND dlg) {
  if (!dlg) {
    return;
  }
  RECT rect = {};
  if (!GetWindowRect(dlg, &rect)) {
    return;
  }
  int width = rect.right - rect.left;
  int height = rect.bottom - rect.top;
  HWND owner = GetWindow(dlg, GW_OWNER);
  RECT owner_rect = {};
  if (owner && GetWindowRect(owner, &owner_rect)) {
    int owner_w = owner_rect.right - owner_rect.left;
    int owner_h = owner_rect.bottom - owner_rect.top;
    int x = owner_rect.left + std::max(0, (owner_w - width) / 2);
    int y = owner_rect.top + std::max(0, (owner_h - height) / 2);
    SetWindowPos(dlg, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
    return;
  }
  RECT work_area = {};
  if (SystemParametersInfoW(SPI_GETWORKAREA, 0, &work_area, 0)) {
    int work_w = work_area.right - work_area.left;
    int work_h = work_area.bottom - work_area.top;
    int x = work_area.left + std::max(0, (work_w - width) / 2);
    int y = work_area.top + std::max(0, (work_h - height) / 2);
    SetWindowPos(dlg, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
  }
}

HFONT CreateDefaultGuiFont() {
  HFONT stock = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
  if (!stock) {
    return ui::DefaultUIFont();
  }
  LOGFONTW lf = {};
  if (GetObjectW(stock, sizeof(lf), &lf) == 0) {
    return ui::DefaultUIFont();
  }
  HFONT font = CreateFontIndirectW(&lf);
  return font ? font : ui::DefaultUIFont();
}

int ControlHeight(HWND dlg, int id) {
  HWND ctrl = GetDlgItem(dlg, id);
  if (!ctrl) {
    return 0;
  }
  RECT rect = {};
  if (!GetWindowRect(ctrl, &rect)) {
    return 0;
  }
  int height = static_cast<int>(rect.bottom - rect.top);
  if (height < 0) {
    height = 0;
  }
  return height;
}

void SetComboHeights(HWND dlg, int id, int height) {
  HWND ctrl = GetDlgItem(dlg, id);
  if (!ctrl || height <= 0) {
    return;
  }
  RECT rect = {};
  if (!GetWindowRect(ctrl, &rect)) {
    return;
  }
  int window_height = rect.bottom - rect.top;
  if (window_height <= 0) {
    return;
  }
  int target = height;
  if (target > window_height) {
    target = window_height;
  }
  SendMessageW(ctrl, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), static_cast<LPARAM>(target));
  int new_total = window_height;
  POINT pt = {rect.left, rect.top};
  ScreenToClient(dlg, &pt);
  SetWindowPos(ctrl, nullptr, pt.x, pt.y, rect.right - rect.left, new_total, SWP_NOZORDER | SWP_NOACTIVATE);
}

void PopulateCombo(HWND combo, const std::vector<std::wstring>& items) {
  if (!combo) {
    return;
  }
  SendMessageW(combo, CB_RESETCONTENT, 0, 0);
  for (const auto& item : items) {
    SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item.c_str()));
  }
}

void SetComboSelection(HWND combo, const std::wstring& value) {
  if (!combo) {
    return;
  }
  if (!value.empty()) {
    int count = static_cast<int>(SendMessageW(combo, CB_GETCOUNT, 0, 0));
    for (int i = 0; i < count; ++i) {
      wchar_t buffer[256] = {};
      SendMessageW(combo, CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer));
      if (_wcsicmp(buffer, value.c_str()) == 0) {
        SendMessageW(combo, CB_SETCURSEL, i, 0);
        return;
      }
    }
  }
  if (SendMessageW(combo, CB_GETCOUNT, 0, 0) > 0) {
    SendMessageW(combo, CB_SETCURSEL, 0, 0);
  }
}

std::wstring ReadComboText(HWND combo) {
  std::wstring text;
  if (!combo) {
    return text;
  }
  int length = GetWindowTextLengthW(combo);
  if (length <= 0) {
    return text;
  }
  text.resize(static_cast<size_t>(length) + 1);
  int copied = GetWindowTextW(combo, &text[0], length + 1);
  text.resize(static_cast<size_t>(copied));
  return text;
}

std::wstring ReadDialogText(HWND dlg, int id) {
  HWND ctrl = GetDlgItem(dlg, id);
  if (!ctrl) {
    return L"";
  }
  int length = GetWindowTextLengthW(ctrl);
  if (length <= 0) {
    return L"";
  }
  std::wstring text;
  text.resize(static_cast<size_t>(length) + 1);
  int copied = GetWindowTextW(ctrl, &text[0], length + 1);
  text.resize(static_cast<size_t>(copied));
  return text;
}

void SetDialogText(HWND dlg, int id, const std::wstring& text) {
  HWND ctrl = GetDlgItem(dlg, id);
  if (ctrl) {
    SetWindowTextW(ctrl, text.c_str());
  }
}

void ToggleCompareControls(HWND dlg, bool left, CompareSourceType type) {
  int root_id = left ? IDC_COMPARE_LEFT_ROOT : IDC_COMPARE_RIGHT_ROOT;
  int path_id = left ? IDC_COMPARE_LEFT_PATH : IDC_COMPARE_RIGHT_PATH;
  int file_id = left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE;
  int browse_id = left ? IDC_COMPARE_LEFT_BROWSE : IDC_COMPARE_RIGHT_BROWSE;
  int key_id = left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY;
  bool reg = type == CompareSourceType::kRegistry;
  EnableWindow(GetDlgItem(dlg, root_id), reg);
  EnableWindow(GetDlgItem(dlg, path_id), reg);
  EnableWindow(GetDlgItem(dlg, file_id), !reg);
  EnableWindow(GetDlgItem(dlg, browse_id), !reg);
  EnableWindow(GetDlgItem(dlg, key_id), !reg);
}

INT_PTR CALLBACK CompareDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<CompareDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<CompareDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    if (!state) {
      return TRUE;
    }
    state->ui_font = CreateDefaultGuiFont();
    ApplyDialogFonts(dlg, state->ui_font);
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    ApplyEditCustomBorder(dlg, IDC_COMPARE_LEFT_PATH);
    ApplyEditCustomBorder(dlg, IDC_COMPARE_LEFT_FILE);
    ApplyEditCustomBorder(dlg, IDC_COMPARE_RIGHT_PATH);
    ApplyEditCustomBorder(dlg, IDC_COMPARE_RIGHT_FILE);

    PopulateCombo(GetDlgItem(dlg, IDC_COMPARE_LEFT_SOURCE), {L"Registry", L"Reg File"});
    PopulateCombo(GetDlgItem(dlg, IDC_COMPARE_RIGHT_SOURCE), {L"Registry", L"Reg File"});
    PopulateCombo(GetDlgItem(dlg, IDC_COMPARE_LEFT_ROOT), state->data.registry_roots);
    PopulateCombo(GetDlgItem(dlg, IDC_COMPARE_RIGHT_ROOT), state->data.registry_roots);

    SetComboSelection(GetDlgItem(dlg, IDC_COMPARE_LEFT_SOURCE), state->data.left.type == CompareSourceType::kRegFile ? L"Reg File" : L"Registry");
    SetComboSelection(GetDlgItem(dlg, IDC_COMPARE_RIGHT_SOURCE), state->data.right.type == CompareSourceType::kRegFile ? L"Reg File" : L"Registry");
    SetComboSelection(GetDlgItem(dlg, IDC_COMPARE_LEFT_ROOT), state->data.left.root);
    SetComboSelection(GetDlgItem(dlg, IDC_COMPARE_RIGHT_ROOT), state->data.right.root);
    SetDialogText(dlg, IDC_COMPARE_LEFT_PATH, state->data.left.path);
    SetDialogText(dlg, IDC_COMPARE_RIGHT_PATH, state->data.right.path);
    SetDialogText(dlg, IDC_COMPARE_LEFT_FILE, state->data.left.file_path);
    SetDialogText(dlg, IDC_COMPARE_RIGHT_FILE, state->data.right.file_path);
    SetDialogText(dlg, IDC_COMPARE_LEFT_KEY, state->data.left.key_path);
    SetDialogText(dlg, IDC_COMPARE_RIGHT_KEY, state->data.right.key_path);
    CheckDlgButton(dlg, IDC_COMPARE_LEFT_RECURSIVE, state->data.left.recursive ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dlg, IDC_COMPARE_RIGHT_RECURSIVE, state->data.right.recursive ? BST_CHECKED : BST_UNCHECKED);

    auto populate_file_keys = [&](bool left) {
      std::wstring file_path = ReadDialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE);
      if (file_path.empty()) {
        return;
      }
      regfile::Document data;
      std::wstring error;
      if (!regfile::Load(file_path, &data, &error)) {
        return;
      }
      std::vector<std::wstring> keys = ExtractRegFileKeys(data);
      HWND combo = GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY);
      PopulateCombo(combo, keys);
      std::wstring current = ReadComboText(combo);
      if (!current.empty()) {
        SetComboSelection(combo, current);
      } else if (!keys.empty()) {
        SendMessageW(combo, CB_SETCURSEL, 0, 0);
        SetDialogText(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY, keys.front());
      }
    };
    populate_file_keys(true);
    populate_file_keys(false);

    int edit_height = ControlHeight(dlg, IDC_COMPARE_LEFT_PATH);
    if (edit_height > 0) {
      SetComboHeights(dlg, IDC_COMPARE_LEFT_SOURCE, edit_height);
      SetComboHeights(dlg, IDC_COMPARE_LEFT_ROOT, edit_height);
      SetComboHeights(dlg, IDC_COMPARE_LEFT_KEY, edit_height);
      SetComboHeights(dlg, IDC_COMPARE_RIGHT_SOURCE, edit_height);
      SetComboHeights(dlg, IDC_COMPARE_RIGHT_ROOT, edit_height);
      SetComboHeights(dlg, IDC_COMPARE_RIGHT_KEY, edit_height);
    }

    ToggleCompareControls(dlg, true, state->data.left.type);
    ToggleCompareControls(dlg, false, state->data.right.type);
    CenterDialogToOwner(dlg);
    return TRUE;
  }
  case WM_DESTROY:
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return TRUE;
  case WM_SETTINGCHANGE:
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG:
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLOREDIT:
  case WM_CTLCOLORLISTBOX:
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    int type = CTLCOLOR_STATIC;
    if (msg == WM_CTLCOLOREDIT) {
      type = CTLCOLOR_EDIT;
    } else if (msg == WM_CTLCOLORLISTBOX) {
      type = CTLCOLOR_LISTBOX;
    } else if (msg == WM_CTLCOLORBTN) {
      type = CTLCOLOR_BTN;
    } else if (msg == WM_CTLCOLORDLG) {
      type = CTLCOLOR_DLG;
    }
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_COMMAND: {
    if (!state) {
      return TRUE;
    }
    int id = LOWORD(wparam);
    int code = HIWORD(wparam);
    if (code == CBN_SELCHANGE && (id == IDC_COMPARE_LEFT_SOURCE || id == IDC_COMPARE_RIGHT_SOURCE)) {
      bool left = id == IDC_COMPARE_LEFT_SOURCE;
      HWND combo = GetDlgItem(dlg, id);
      int sel = combo ? static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0)) : 0;
      CompareSourceType type = (sel == 1) ? CompareSourceType::kRegFile : CompareSourceType::kRegistry;
      ToggleCompareControls(dlg, left, type);
      return TRUE;
    }
    if (code == BN_CLICKED && (id == IDC_COMPARE_LEFT_BROWSE || id == IDC_COMPARE_RIGHT_BROWSE)) {
      bool left = id == IDC_COMPARE_LEFT_BROWSE;
      std::wstring path;
      if (!PromptOpenFilePath(dlg, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
        return TRUE;
      }
      SetDialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE, path);
      regfile::Document data;
      std::wstring error;
      if (regfile::Load(path, &data, &error)) {
        std::vector<std::wstring> keys = ExtractRegFileKeys(data);
        if (keys.empty()) {
          ui::ShowError(dlg, L"No registry keys were found in the .reg file.");
          return TRUE;
        }
        HWND combo = GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY);
        PopulateCombo(combo, keys);
        if (!keys.empty()) {
          SendMessageW(combo, CB_SETCURSEL, 0, 0);
          SetDialogText(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY, keys.front());
        }
      } else if (!error.empty()) {
        ui::ShowError(dlg, error);
      }
      return TRUE;
    }
    if (id == IDOK) {
      CompareDialogResult result;
      auto read_side = [&](bool left, CompareDialogSelection* out) -> bool {
        out->recursive = IsDlgButtonChecked(dlg, left ? IDC_COMPARE_LEFT_RECURSIVE : IDC_COMPARE_RIGHT_RECURSIVE) == BST_CHECKED;
        HWND source_combo = GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_SOURCE : IDC_COMPARE_RIGHT_SOURCE);
        int source_index = source_combo ? static_cast<int>(SendMessageW(source_combo, CB_GETCURSEL, 0, 0)) : 0;
        out->type = (source_index == 1) ? CompareSourceType::kRegFile : CompareSourceType::kRegistry;
        if (out->type == CompareSourceType::kRegistry) {
          out->root = TrimWhitespace(ReadComboText(GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_ROOT : IDC_COMPARE_RIGHT_ROOT)));
          out->path = TrimWhitespace(ReadDialogText(dlg, left ? IDC_COMPARE_LEFT_PATH : IDC_COMPARE_RIGHT_PATH));
          if (out->root.empty()) {
            ui::ShowError(dlg, L"Registry root is required.");
            return false;
          }
          return true;
        }
        out->file_path = TrimWhitespace(ReadDialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE));
        out->key_path = TrimWhitespace(ReadComboText(GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY)));
        if (out->file_path.empty()) {
          ui::ShowError(dlg, L"Registry file path is required.");
          return false;
        }
        regfile::Document data;
        std::wstring error;
        if (!regfile::Load(out->file_path, &data, &error)) {
          ui::ShowError(dlg, error.empty() ? L"Failed to read registry file." : error);
          return false;
        }
        std::vector<std::wstring> keys = ExtractRegFileKeys(data);
        if (keys.empty()) {
          ui::ShowError(dlg, L"No registry keys were found in the .reg file.");
          return false;
        }
        if (out->key_path.empty()) {
          out->key_path = keys.front();
        }
        std::wstring key_lower = ToLower(out->key_path);
        bool found = false;
        for (const auto& key : keys) {
          if (_wcsicmp(key.c_str(), out->key_path.c_str()) == 0) {
            found = true;
            break;
          }
          std::wstring key_check = ToLower(key);
          if (StartsWithInsensitive(key_check, key_lower) || StartsWithInsensitive(key_lower, key_check)) {
            found = true;
          }
        }
        if (!found) {
          ui::ShowError(dlg, L"The selected key path was not found in the .reg file.");
          return false;
        }
        return true;
      };
      if (!read_side(true, &result.left)) {
        return TRUE;
      }
      if (!read_side(false, &result.right)) {
        return TRUE;
      }
      state->data.left = result.left;
      state->data.right = result.right;
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    if (id == IDCANCEL) {
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    }
    break;
  }
  default:
    break;
  }
  return FALSE;
}

bool ShowCompareDialog(HWND owner, const CompareDialogDefaults& defaults, CompareDialogResult* out) {
  if (!out) {
    return false;
  }
  CompareDialogState state;
  state.data = defaults;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_COMPARE), owner, CompareDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result != IDOK) {
    return false;
  }
  out->left = state.data.left;
  out->right = state.data.right;
  return true;
}

} // namespace

std::wstring MainWindow::CommandShortcutText(int command_id) const {
  switch (command_id) {
  case cmd::kEditCopy:
    return L"Ctrl+C";
  case cmd::kEditPaste:
    return L"Ctrl+V";
  case cmd::kEditUndo:
    return L"Ctrl+Z";
  case cmd::kEditRedo:
    return L"Ctrl+Y";
  case cmd::kEditFind:
    return L"Ctrl+F";
  case cmd::kEditReplace:
    return L"Ctrl+H";
  case cmd::kEditGoTo:
    return L"Ctrl+G";
  case cmd::kEditRename:
    return L"F2";
  case cmd::kEditDelete:
    return L"Del";
  case cmd::kEditCopyKey:
    return L"Ctrl+Shift+C";
  case cmd::kViewSelectAll:
    return L"Ctrl+A";
  case cmd::kFileSave:
    return L"Ctrl+S";
  case cmd::kFileExport:
    return L"Ctrl+E";
  case cmd::kViewRefresh:
    return L"F5";
  case cmd::kNavBack:
    return L"Alt+Left";
  case cmd::kNavForward:
    return L"Alt+Right";
  case cmd::kNavUp:
    return L"Alt+Up";
  default:
    return L"";
  }
}

std::wstring MainWindow::CommandTooltipText(int command_id) const {
  switch (command_id) {
  case cmd::kRegistryLocal:
    return L"Local Registry";
  case cmd::kRegistryNetwork:
    return L"Remote Registry";
  case cmd::kRegistryOffline:
    return L"Offline Registry";
  case cmd::kEditFind:
    return L"Find";
  case cmd::kEditReplace:
    return L"Replace";
  case cmd::kFileSave:
    return L"Save";
  case cmd::kFileExport:
    return L"Export";
  case cmd::kEditUndo:
    return L"Undo";
  case cmd::kEditRedo:
    return L"Redo";
  case cmd::kEditCopy:
    return L"Copy";
  case cmd::kEditPaste:
    return L"Paste";
  case cmd::kEditDelete:
    return L"Delete";
  case cmd::kViewRefresh:
    return L"Refresh";
  case cmd::kNavBack:
    return L"Back";
  case cmd::kNavForward:
    return L"Forward";
  case cmd::kNavUp:
    return L"Up";
  default:
    return L"";
  }
}

bool MainWindow::EnsureWritable() {
  if (!read_only_) {
    return true;
  }
  ui::ShowWarning(hwnd_, L"Read-only mode is enabled.");
  return false;
}

void MainWindow::BuildMenus() {
  if (deferred_startup_complete_) {
    SyncReplaceRegeditState();
  }
  if (deferred_startup_complete_ && !favorites_loaded_) {
    RefreshFavoritesCache();
  }
  if (deferred_startup_complete_ && !bundled_defaults_loaded_) {
    RefreshBundledDefaultsCache();
  }
  menu_items_.clear();
  bool can_modify = !read_only_;
  HMENU menu = CreateMenu();
  HMENU file_menu = CreatePopupMenu();
  auto append_menu = [&](HMENU target, UINT flags, int command, const wchar_t* text) {
    std::wstring shortcut = CommandShortcutText(command);
    if (!shortcut.empty()) {
      std::wstring combined = std::wstring(text) + L"\t" + shortcut;
      AppendMenuW(target, flags, command, combined.c_str());
      return;
    }
    AppendMenuW(target, flags, command, text);
  };
  bool can_save = false;
  if (can_modify && tab_) {
    int sel = TabCtrl_GetCurSel(tab_);
    if (sel >= 0 && static_cast<size_t>(sel) < tabs_.size()) {
      const auto& entry = tabs_[static_cast<size_t>(sel)];
      if (entry.kind == TabEntry::Kind::kRegFile) {
        can_save = entry.reg_file_dirty;
      } else if (entry.kind == TabEntry::Kind::kRegistry && entry.registry_mode == RegistryMode::kOffline) {
        can_save = entry.offline_dirty;
      }
    }
  }
  UINT save_flags = MF_STRING | (can_save ? 0 : MF_GRAYED);
  append_menu(file_menu, save_flags, cmd::kFileSave, L"Save");
  append_menu(file_menu, MF_STRING, cmd::kFileOpenRegFile, L"Open .reg File...");
  UINT import_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
  append_menu(file_menu, import_flags, cmd::kFileImport, L"Import...");
  append_menu(file_menu, MF_STRING, cmd::kFileExport, L"Export...");
  append_menu(file_menu, MF_STRING, cmd::kFileImportComments, L"Import Comments...");
  append_menu(file_menu, MF_STRING, cmd::kFileExportComments, L"Export Comments...");
  append_menu(file_menu, MF_STRING, cmd::kOptionsCompareRegistries, L"Compare Registries...");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  UINT hive_modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
  append_menu(file_menu, hive_modify_flags, cmd::kFileLoadHive, L"Load Hive...");
  append_menu(file_menu, hive_modify_flags, cmd::kFileUnloadHive, L"Unload Hive...");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  UINT local_flags = MF_STRING | (registry_mode_ == RegistryMode::kLocal ? MF_CHECKED : MF_UNCHECKED);
  UINT remote_flags = MF_STRING | (registry_mode_ == RegistryMode::kRemote ? MF_CHECKED : MF_UNCHECKED);
  UINT offline_flags = MF_STRING | (registry_mode_ == RegistryMode::kOffline ? MF_CHECKED : MF_UNCHECKED);
  append_menu(file_menu, local_flags, cmd::kRegistryLocal, L"Local Registry");
  append_menu(file_menu, remote_flags, cmd::kRegistryNetwork, L"Remote Registry...");
  append_menu(file_menu, offline_flags, cmd::kRegistryOffline, L"Offline Registry...");
  UINT save_offline_flags = MF_STRING | ((registry_mode_ == RegistryMode::kOffline && !offline_mount_.empty()) ? 0 : MF_GRAYED);
  append_menu(file_menu, save_offline_flags, cmd::kFileSaveOfflineHive, L"Save Offline Hive...");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  UINT clear_flags = MF_STRING | (clear_history_on_exit_ ? MF_CHECKED : MF_UNCHECKED);
  append_menu(file_menu, clear_flags, cmd::kFileClearHistoryOnExit, L"Clear History on Exit");
  UINT clear_tabs_flags = MF_STRING | (clear_tabs_on_exit_ ? MF_CHECKED : MF_UNCHECKED);
  append_menu(file_menu, clear_tabs_flags, cmd::kFileClearTabsOnExit, L"Clear Tabs on Exit");
  AppendMenuW(file_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(file_menu, MF_STRING, cmd::kFileExit, L"Exit");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(file_menu), L"File");

  HMENU edit_menu = CreatePopupMenu();
  UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
  append_menu(edit_menu, modify_flags, cmd::kEditModify, L"Modify...");
  append_menu(edit_menu, modify_flags, cmd::kEditModifyBinary, L"Modify Binary Data...");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, modify_flags, cmd::kEditUndo, L"Undo");
  append_menu(edit_menu, modify_flags, cmd::kEditRedo, L"Redo");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  HMENU edit_new = CreatePopupMenu();
  AppendMenuW(edit_new, MF_STRING, cmd::kNewKey, L"Key");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewString, L"String Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewBinary, L"Binary Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
  AppendMenuW(edit_new, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");
  AppendMenuW(edit_menu, MF_POPUP | (can_modify ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(edit_new), L"New");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, MF_STRING, cmd::kEditCopy, L"Copy");
  append_menu(edit_menu, modify_flags, cmd::kEditPaste, L"Paste");
  append_menu(edit_menu, modify_flags, cmd::kEditRename, L"Rename");
  append_menu(edit_menu, modify_flags, cmd::kEditDelete, L"Delete");
  append_menu(edit_menu, MF_STRING, cmd::kViewSelectAll, L"Select All");
  append_menu(edit_menu, MF_STRING, cmd::kEditInvertSelection, L"Invert Selection");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, MF_STRING, cmd::kEditCopyKey, L"Copy Key Name");
  append_menu(edit_menu, MF_STRING, cmd::kEditCopyKeyPath, L"Copy Key Path");
  AppendMenuW(edit_menu, MF_POPUP, reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
  UINT permissions_flags = MF_STRING | ((current_node_ && can_modify) ? 0 : MF_GRAYED);
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, MF_STRING, cmd::kEditGoTo, L"Go to...");
  append_menu(edit_menu, MF_STRING, cmd::kEditFind, L"Find...");
  append_menu(edit_menu, modify_flags, cmd::kEditReplace, L"Replace...");
  AppendMenuW(edit_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(edit_menu, permissions_flags, cmd::kEditPermissions, L"Permissions...");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(edit_menu), L"Edit");

  HMENU view_menu = CreatePopupMenu();
  append_menu(view_menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
  append_menu(view_menu, MF_STRING, cmd::kViewSelectAll, L"Select All");
  AppendMenuW(view_menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(view_menu, MF_STRING | (show_toolbar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewToolbar, L"Toolbar");
  AppendMenuW(view_menu, MF_STRING | (show_address_bar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewAddressBar, L"Address Bar");
  AppendMenuW(view_menu, MF_STRING | (show_filter_bar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewFilterBar, L"Filter Bar");
  AppendMenuW(view_menu, MF_STRING | (show_tab_control_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewTabControl, L"Tab Control");
  AppendMenuW(view_menu, MF_STRING | (show_tree_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewKeyTree, L"Key Tree");
  AppendMenuW(view_menu, MF_STRING | (show_keys_in_list_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewKeysInList, L"Keys in List");
  UINT simulated_flags = MF_STRING | (show_simulated_keys_ ? MF_CHECKED : MF_UNCHECKED);
  if (!HasActiveTraces()) {
    simulated_flags |= MF_GRAYED;
  }
  AppendMenuW(view_menu, simulated_flags, cmd::kViewSimulatedKeys, L"Simulated Keys");
  AppendMenuW(view_menu, MF_STRING | (show_history_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewHistory, L"History");
  AppendMenuW(view_menu, MF_STRING | (show_status_bar_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewStatusBar, L"Status Bar");
  UINT extra_flags = MF_STRING | (show_extra_hives_ ? MF_CHECKED : MF_UNCHECKED);
  if (registry_mode_ != RegistryMode::kLocal) {
    extra_flags |= MF_GRAYED;
  }
  AppendMenuW(view_menu, extra_flags, cmd::kViewExtraHives, L"Show Extra Hives");
  AppendMenuW(view_menu, MF_SEPARATOR, 0, nullptr);
  UINT hive_flags = MF_STRING | (registry_mode_ == RegistryMode::kLocal ? 0 : MF_GRAYED);
  append_menu(view_menu, hive_flags, cmd::kOptionsHiveFileDir, L"Open Hive File");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(view_menu), L"View");

  HMENU options_menu = CreatePopupMenu();
  HMENU theme_menu = CreatePopupMenu();
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kSystem ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeSystem, L"System");
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kLight ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeLight, L"Light");
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kDark ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeDark, L"Dark");
  AppendMenuW(theme_menu, MF_STRING | (theme_mode_ == ThemeMode::kCustom ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsThemeCustom, L"Custom");
  AppendMenuW(theme_menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(theme_menu, MF_STRING, cmd::kOptionsThemePresets, L"Theme Presets...");
  AppendMenuW(options_menu, MF_POPUP, reinterpret_cast<UINT_PTR>(theme_menu), L"Theme");
  HMENU icon_menu = CreatePopupMenu();
  auto icon_flags = [&](const wchar_t* name) -> UINT { return MF_STRING | (_wcsicmp(icon_set_.c_str(), name) == 0 ? MF_CHECKED : MF_UNCHECKED); };
  AppendMenuW(icon_menu, icon_flags(L"default"), cmd::kOptionsIconSetDefault, L"Lucide");
  AppendMenuW(icon_menu, icon_flags(L"tabler"), cmd::kOptionsIconSetTabler, L"Tabler");
  AppendMenuW(icon_menu, icon_flags(L"fluentui"), cmd::kOptionsIconSetFluentUi, L"Fluent UI");
  AppendMenuW(icon_menu, icon_flags(L"materialsymbols"), cmd::kOptionsIconSetMaterialSymbols, L"Material Symbols");
  AppendMenuW(icon_menu, icon_flags(L"custom"), cmd::kOptionsIconSetCustom, L"Custom");
  AppendMenuW(options_menu, MF_POPUP, reinterpret_cast<UINT_PTR>(icon_menu), L"Icons");
  AppendMenuW(options_menu, MF_STRING, cmd::kViewFont, L"Font...");
  AppendMenuW(options_menu, MF_SEPARATOR, 0, nullptr);
  bool is_elevated = util::IsProcessElevated();
  bool is_system = util::IsProcessSystem();
  bool is_ti = util::IsProcessTrustedInstaller();
  UINT admin_flags = MF_STRING | (is_elevated ? MF_GRAYED : 0);
  AppendMenuW(options_menu, admin_flags, cmd::kOptionsRestartAdmin, L"Restart as Admin");
  AppendMenuW(options_menu, MF_STRING | (always_run_as_admin_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsAlwaysRunAdmin, L"Always run as Admin");
  UINT system_flags = MF_STRING | (is_system ? MF_GRAYED : 0);
  AppendMenuW(options_menu, system_flags, cmd::kOptionsRestartSystem, L"Restart as SYSTEM");
  AppendMenuW(options_menu, MF_STRING | (always_run_as_system_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsAlwaysRunSystem, L"Always run as SYSTEM");
  UINT ti_flags = MF_STRING | (is_ti ? MF_GRAYED : 0);
  AppendMenuW(options_menu, ti_flags, cmd::kOptionsRestartTrustedInstaller, L"Restart as TI");
  AppendMenuW(options_menu, MF_STRING | (always_run_as_trustedinstaller_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsAlwaysRunTrustedInstaller, L"Always run as TI");
  AppendMenuW(options_menu, MF_SEPARATOR, 0, nullptr);
  UINT regedit_flags = MF_STRING | ((is_elevated || is_system || is_ti) ? 0 : MF_GRAYED);
  AppendMenuW(options_menu, regedit_flags, cmd::kOptionsOpenDefaultRegedit, L"Open Default Regedit");
  AppendMenuW(options_menu, MF_STRING | (replace_regedit_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsReplaceRegedit, L"Replace Regedit");
  AppendMenuW(options_menu, MF_STRING | (single_instance_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsSingleInstance, L"Single Instance");
  AppendMenuW(options_menu, MF_STRING | (save_tabs_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsSaveTabs, L"Save Tabs");
  AppendMenuW(options_menu, MF_STRING | (read_only_ ? MF_CHECKED : MF_UNCHECKED), cmd::kOptionsReadOnly, L"Read Only Mode");
  AppendMenuW(options_menu, MF_STRING | (save_tree_state_ ? MF_CHECKED : MF_UNCHECKED), cmd::kViewSaveTreeState, L"Save Previous Tree State");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(options_menu), L"Options");

  HMENU favorites_menu = CreatePopupMenu();
  AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesAdd, L"Add to Favorites...");
  AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesRemove, L"Remove Favorite");
  AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesEdit, L"Edit Favorites...");
  AppendMenuW(favorites_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(favorites_menu, MF_STRING, cmd::kFavoritesImport, L"Import Favorites...");
  append_menu(favorites_menu, MF_STRING, cmd::kFavoritesImportRegedit, L"Import Regedit Favorites");
  append_menu(favorites_menu, MF_STRING, cmd::kFavoritesExport, L"Export Favorites...");
  if (!favorites_cache_.empty()) {
    AppendMenuW(favorites_menu, MF_SEPARATOR, 0, nullptr);
    int limit = std::min(static_cast<int>(favorites_cache_.size()), cmd::kFavoritesItemMax - cmd::kFavoritesItemBase + 1);
    for (int i = 0; i < limit; ++i) {
      AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesItemBase + i, favorites_cache_[static_cast<size_t>(i)].c_str());
    }
  }
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(favorites_menu), L"Favorites");

  HMENU window_menu = CreatePopupMenu();
  AppendMenuW(window_menu, MF_STRING, cmd::kWindowNew, L"New Window");
  AppendMenuW(window_menu, MF_STRING, cmd::kWindowClose, L"Close Window");
  AppendMenuW(window_menu, MF_STRING | (always_on_top_ ? MF_CHECKED : MF_UNCHECKED), cmd::kWindowAlwaysOnTop, L"Always on Top");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(window_menu), L"Window");

  HMENU trace_menu = CreatePopupMenu();
  auto has_label = [&](const wchar_t* label) -> bool {
    for (const auto& trace : active_traces_) {
      if (_wcsicmp(trace.label.c_str(), label) == 0) {
        return true;
      }
    }
    return false;
  };
  auto has_path = [&](const std::wstring& path) -> bool {
    for (const auto& trace : active_traces_) {
      if (EqualsInsensitive(trace.source_path, path)) {
        return true;
      }
    }
    return false;
  };
  bool trace_23h2 = has_label(L"23H2");
  bool trace_24h2 = has_label(L"24H2");
  bool trace_25h2 = has_label(L"25H2");
  bool has_recent_trace = false;
  append_menu(trace_menu, MF_STRING | (trace_23h2 ? MF_CHECKED : MF_UNCHECKED), cmd::kTraceLoad23H2, L"23H2");
  append_menu(trace_menu, MF_STRING | (trace_24h2 ? MF_CHECKED : MF_UNCHECKED), cmd::kTraceLoad24H2, L"24H2");
  append_menu(trace_menu, MF_STRING | (trace_25h2 ? MF_CHECKED : MF_UNCHECKED), cmd::kTraceLoad25H2, L"25H2");
  int recent_limit = std::min(static_cast<int>(recent_trace_paths_.items().size()), cmd::kTraceRecentMax - cmd::kTraceRecentBase + 1);
  for (int i = 0; i < recent_limit; ++i) {
    const std::wstring& path = recent_trace_paths_.items()[static_cast<size_t>(i)];
    if (path.empty()) {
      continue;
    }
    has_recent_trace = true;
    std::wstring name = FileNameOnly(path);
    if (name.empty()) {
      name = L"Trace";
    }
    UINT flags = MF_STRING;
    if (has_path(path)) {
      flags |= MF_CHECKED;
    }
    append_menu(trace_menu, flags, cmd::kTraceRecentBase + i, name.c_str());
  }
  AppendMenuW(trace_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(trace_menu, MF_STRING, cmd::kTraceGuide, L"Guide");
  append_menu(trace_menu, MF_STRING, cmd::kTraceLoadCustom, L"Open Trace File...");
  UINT edit_recent_flags = MF_STRING | (has_recent_trace ? 0 : MF_GRAYED);
  append_menu(trace_menu, edit_recent_flags, cmd::kTraceEditRecent, L"Edit Recent Traces...");
  append_menu(trace_menu, MF_STRING, cmd::kTraceEditActive, L"Edit Active Traces...");
  UINT clear_trace_flags = MF_STRING | (!active_traces_.empty() ? 0 : MF_GRAYED);
  append_menu(trace_menu, clear_trace_flags, cmd::kTraceClear, L"Clear Trace");

  HMENU default_menu = CreatePopupMenu();
  auto has_default_path = [&](const std::wstring& path) -> bool {
    for (const auto& defaults : active_defaults_) {
      if (EqualsInsensitive(defaults.source_path, path)) {
        return true;
      }
    }
    return false;
  };
  for (size_t i = 0; i < bundled_defaults_.size(); ++i) {
    const auto& entry = bundled_defaults_[i];
    UINT flags = MF_STRING;
    if (has_default_path(entry.path)) {
      flags |= MF_CHECKED;
    }
    append_menu(default_menu, flags, cmd::kDefaultBundledBase + static_cast<int>(i), entry.label.c_str());
  }
  bool has_recent_default = false;
  int default_recent_limit = std::min(static_cast<int>(recent_default_paths_.items().size()), cmd::kDefaultRecentMax - cmd::kDefaultRecentBase + 1);
  for (int i = 0; i < default_recent_limit; ++i) {
    const std::wstring& path = recent_default_paths_.items()[static_cast<size_t>(i)];
    if (path.empty()) {
      continue;
    }
    has_recent_default = true;
    std::wstring name = FileNameOnly(path);
    if (name.empty()) {
      name = L"Default";
    }
    UINT flags = MF_STRING;
    if (has_default_path(path)) {
      flags |= MF_CHECKED;
    }
    append_menu(default_menu, flags, cmd::kDefaultRecentBase + i, name.c_str());
  }
  AppendMenuW(default_menu, MF_SEPARATOR, 0, nullptr);
  append_menu(default_menu, MF_STRING, cmd::kDefaultLoadCustom, L"Open Default File...");
  UINT edit_default_recent_flags = MF_STRING | (has_recent_default ? 0 : MF_GRAYED);
  append_menu(default_menu, edit_default_recent_flags, cmd::kDefaultEditRecent, L"Edit Recent Defaults...");
  append_menu(default_menu, MF_STRING, cmd::kDefaultEditActive, L"Edit Active Defaults...");
  UINT clear_default_flags = MF_STRING | (!active_defaults_.empty() ? 0 : MF_GRAYED);
  append_menu(default_menu, clear_default_flags, cmd::kDefaultClear, L"Clear Defaults");

  HMENU help_menu = CreatePopupMenu();
  AppendMenuW(help_menu, MF_STRING, cmd::kHelpContents, L"Help");
  AppendMenuW(help_menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(help_menu, MF_STRING, cmd::kHelpAbout, L"About RegKit");

  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(trace_menu), L"Trace");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(default_menu), L"Default");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(help_menu), L"Help");

  PrepareMenusForOwnerDraw(menu, true);

  HMENU old_menu = GetMenu(hwnd_);
  SetMenu(hwnd_, menu);
  DrawMenuBar(hwnd_);
  if (old_menu) {
    DestroyMenu(old_menu);
  }
}

void MainWindow::RefreshFavoritesCache() {
  favorites_cache_.clear();
  FavoritesStore::Load(&favorites_cache_);
  favorites_loaded_ = true;
}

void MainWindow::RefreshBundledDefaultsCache() {
  bundled_defaults_.clear();
  bundled_defaults_loaded_ = true;

  std::wstring module_dir = util::GetModuleDirectory();
  if (module_dir.empty()) {
    return;
  }
  std::wstring assets = util::JoinPath(module_dir, L"assets");
  std::wstring defaults_dir = util::JoinPath(assets, L"defaults");
  std::wstring pattern = util::JoinPath(defaults_dir, L"*.reg");
  WIN32_FIND_DATAW data = {};
  HANDLE find = FindFirstFileW(pattern.c_str(), &data);
  if (find == INVALID_HANDLE_VALUE) {
    return;
  }
  do {
    if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
      continue;
    }
    std::wstring file_name = data.cFileName;
    std::wstring label = FileBaseName(file_name);
    if (label.empty()) {
      continue;
    }
    BundledDefault entry;
    entry.label = std::move(label);
    entry.path = util::JoinPath(defaults_dir, file_name);
    bundled_defaults_.push_back(std::move(entry));
  } while (FindNextFileW(find, &data));
  FindClose(find);

  std::sort(bundled_defaults_.begin(), bundled_defaults_.end(), [](const BundledDefault& left, const BundledDefault& right) { return _wcsicmp(left.label.c_str(), right.label.c_str()) < 0; });
  size_t bundled_limit = std::min(bundled_defaults_.size(), static_cast<size_t>(cmd::kDefaultBundledMax - cmd::kDefaultBundledBase + 1));
  if (bundled_defaults_.size() > bundled_limit) {
    bundled_defaults_.resize(bundled_limit);
  }
}

bool MainWindow::HandleMenuCommand(int command_id) {
  if (command_id >= cmd::kFavoritesItemBase && command_id <= cmd::kFavoritesItemMax) {
    if (!favorites_loaded_) {
      RefreshFavoritesCache();
    }
    size_t index = static_cast<size_t>(command_id - cmd::kFavoritesItemBase);
    if (index < favorites_cache_.size()) {
      SelectTreePath(favorites_cache_[index]);
      return true;
    }
  }
  if (command_id >= cmd::kDefaultBundledBase && command_id <= cmd::kDefaultBundledMax) {
    size_t index = static_cast<size_t>(command_id - cmd::kDefaultBundledBase);
    if (index < bundled_defaults_.size()) {
      const auto& entry = bundled_defaults_[index];
      if (RemoveDefaultByPath(entry.path)) {
        return true;
      }
      LoadDefaultFromFile(entry.label, entry.path);
      return true;
    }
  }
  if (command_id >= cmd::kDefaultRecentBase && command_id <= cmd::kDefaultRecentMax) {
    size_t index = static_cast<size_t>(command_id - cmd::kDefaultRecentBase);
    if (index < recent_default_paths_.items().size()) {
      std::wstring path = recent_default_paths_.items()[index];
      std::wstring label = FileBaseName(path);
      if (label.empty()) {
        label = L"Default";
      }
      if (RemoveDefaultByPath(path)) {
        return true;
      }
      if (LoadDefaultFromFile(label, path)) {
        AddRecentDefaultPath(path);
        BuildMenus();
        SaveSettings();
      }
      return true;
    }
  }
  if (command_id >= cmd::kTraceRecentBase && command_id <= cmd::kTraceRecentMax) {
    size_t index = static_cast<size_t>(command_id - cmd::kTraceRecentBase);
    if (index < recent_trace_paths_.items().size()) {
      std::wstring path = recent_trace_paths_.items()[index];
      std::wstring label = FileBaseName(path);
      if (label.empty()) {
        label = L"Trace";
      }
      if (RemoveTraceByPath(path)) {
        return true;
      }
      if (LoadTraceFromFile(label, path)) {
        AddRecentTracePath(path);
        BuildMenus();
        SaveSettings();
        SaveActiveTraces();
      }
      return true;
    }
  }

  switch (command_id) {
  case cmd::kFileExit:
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  case cmd::kFileImport: {
    if (!EnsureWritable()) {
      return true;
    }
    std::wstring path;
    if (!PromptOpenFilePath(hwnd_, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    std::wstring error;
    if (ImportRegFileFromPath(path, &error)) {
      AppendHistoryEntry(L"Import .reg file " + FileNameOnly(path), L"", path);
    } else if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return true;
  }
  case cmd::kFileOpenRegFile: {
    std::wstring path;
    if (!PromptOpenFilePath(hwnd_, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    OpenRegFileTab(path);
    return true;
  }
  case cmd::kFileSave: {
    if (!EnsureWritable()) {
      return true;
    }
    if (IsRegFileTabSelected()) {
      int tab_index = TabCtrl_GetCurSel(tab_);
      if (tab_index >= 0 && static_cast<size_t>(tab_index) < tabs_.size()) {
        if (tabs_[static_cast<size_t>(tab_index)].reg_file_dirty) {
          SaveRegFileTab(tab_index);
        }
      }
      return true;
    }
    if (registry_mode_ == RegistryMode::kOffline) {
      int index = CurrentRegistryTabIndex();
      if (index >= 0 && static_cast<size_t>(index) < tabs_.size()) {
        if (tabs_[static_cast<size_t>(index)].offline_dirty) {
          SaveOfflineRegistry();
        }
      }
      return true;
    }
    return true;
  }
  case cmd::kFileExport: {
    if (IsRegFileTabSelected()) {
      int tab_index = TabCtrl_GetCurSel(tab_);
      std::wstring path;
      if (!PromptSaveFilePath(hwnd_, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0", &path)) {
        return true;
      }
      if (ExportRegFileTab(tab_index, path)) {
        AppendHistoryEntry(L"Export .reg tab " + FileNameOnly(path), L"", path);
      }
      return true;
    }
    if (!current_node_) {
      return true;
    }
    if (value_list_.hwnd() && GetFocus() == value_list_.hwnd()) {
      std::vector<std::wstring> selected_values;
      std::vector<std::wstring> selected_keys;
      int index = -1;
      while ((index = ListView_GetNextItem(value_list_.hwnd(), index, LVNI_SELECTED)) >= 0) {
        const ListRow* row = value_list_.RowAt(index);
        if (!row) {
          continue;
        }
        if (row->kind == rowkind::kValue) {
          selected_values.push_back(row->extra);
        } else if (row->kind == rowkind::kKey) {
          selected_keys.push_back(row->extra);
        }
      }
      if (!selected_values.empty() || !selected_keys.empty()) {
        auto dedupe = [](std::vector<std::wstring>* items) {
          if (!items) {
            return;
          }
          std::unordered_set<std::wstring> seen;
          std::vector<std::wstring> unique;
          unique.reserve(items->size());
          for (const auto& item : *items) {
            std::wstring key = ToLower(item);
            if (seen.insert(key).second) {
              unique.push_back(item);
            }
          }
          *items = std::move(unique);
        };
        dedupe(&selected_values);
        dedupe(&selected_keys);
        std::wstring error;
        std::wstring path = registry_path::Build(*current_node_);
        if (ExportRegFileSelection(hwnd_, path, selected_values, selected_keys, &error)) {
          HistoryEntry entry;
          entry.action = L"Export registry selection";
          entry.old_data = std::to_wstring(selected_keys.size()) + L" keys, " + std::to_wstring(selected_values.size()) + L" values";
          entry.key_path = path;
          AppendHistoryEntry(std::move(entry));
        } else if (!error.empty()) {
          ui::ShowError(hwnd_, error);
        }
        return true;
      }
    }
    std::wstring error;
    std::wstring path = registry_path::Build(*current_node_);
    if (ExportRegFile(hwnd_, path, &error)) {
      HistoryEntry entry;
      entry.action = L"Export registry key";
      entry.key_path = path;
      entry.new_data = path;
      AppendHistoryEntry(std::move(entry));
    } else if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return true;
  }
  case cmd::kFileImportComments: {
    std::wstring path;
    if (!PromptOpenFilePath(hwnd_, L"RegKit Comment Files (*.rkc)\0*.rkc\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    if (ImportCommentsFromFile(path)) {
      AppendHistoryEntry(L"Import comments " + FileNameOnly(path), L"", path);
    } else {
      ui::ShowError(hwnd_, L"Failed to import comments.");
    }
    return true;
  }
  case cmd::kFileExportComments: {
    std::wstring path;
    if (!PromptSaveFilePath(hwnd_, L"RegKit Comment Files (*.rkc)\0*.rkc\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    if (ExportCommentsToFile(path)) {
      AppendHistoryEntry(L"Export comments " + FileNameOnly(path), L"", path);
    } else {
      ui::ShowError(hwnd_, L"Failed to export comments.");
    }
    return true;
  }
  case cmd::kFileLoadHive: {
    if (!EnsureWritable()) {
      return true;
    }
    if (registry_mode_ == RegistryMode::kRemote) {
      ui::ShowError(hwnd_, L"Loading hives is not supported for remote registries.");
      return true;
    }
    std::wstring error;
    HKEY root = HKEY_LOCAL_MACHINE;
    if (current_node_ && (current_node_->root == HKEY_LOCAL_MACHINE || current_node_->root == HKEY_USERS)) {
      root = current_node_->root;
    }
    if (LoadHive(hwnd_, root, &error)) {
      AppendHistoryEntry(L"Load hive", L"", registry_path::RootName(root));
      UpdateValueListForNode(current_node_);
    } else if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return true;
  }
  case cmd::kFileUnloadHive: {
    if (!EnsureWritable()) {
      return true;
    }
    if (registry_mode_ == RegistryMode::kRemote) {
      ui::ShowError(hwnd_, L"Unloading hives is not supported for remote registries.");
      return true;
    }
    HKEY root = HKEY_LOCAL_MACHINE;
    std::wstring subkey;
    if (current_node_ && (current_node_->root == HKEY_LOCAL_MACHINE || current_node_->root == HKEY_USERS)) {
      root = current_node_->root;
      subkey = current_node_->subkey;
    }
    std::wstring error;
    if (!UnloadHive(hwnd_, root, subkey, &error) && !error.empty()) {
      ui::ShowError(hwnd_, error);
    } else {
      UpdateValueListForNode(current_node_);
    }
    return true;
  }
  case cmd::kFileSaveOfflineHive:
    SaveOfflineRegistry();
    return true;
  case cmd::kFileClearHistoryOnExit:
    clear_history_on_exit_ = !clear_history_on_exit_;
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kFileClearTabsOnExit:
    clear_tabs_on_exit_ = !clear_tabs_on_exit_;
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewRefresh:
    RefreshTreeSelection();
    UpdateValueListForNode(current_node_);
    return true;
  case cmd::kViewAddressBar:
    show_address_bar_ = !show_address_bar_;
    SaveSettings();
    ApplyViewVisibility();
    BuildMenus();
    return true;
  case cmd::kViewFilterBar:
    show_filter_bar_ = !show_filter_bar_;
    SaveSettings();
    ApplyViewVisibility();
    BuildMenus();
    return true;
  case cmd::kViewTabControl:
    show_tab_control_ = !show_tab_control_;
    SaveSettings();
    ApplyViewVisibility();
    BuildMenus();
    return true;
  case cmd::kTreeToggleExpand: {
    if (!tree_.hwnd()) {
      return true;
    }
    HTREEITEM item = TreeView_GetSelection(tree_.hwnd());
    if (!item) {
      return true;
    }
    TVITEMW tvi = {};
    tvi.hItem = item;
    tvi.mask = TVIF_STATE | TVIF_CHILDREN;
    tvi.stateMask = TVIS_EXPANDED;
    if (!TreeView_GetItem(tree_.hwnd(), &tvi)) {
      return true;
    }
    bool expanded = (tvi.state & TVIS_EXPANDED) != 0;
    bool has_child = TreeView_GetChild(tree_.hwnd(), item) != nullptr || tvi.cChildren != 0;
    if (!expanded && !has_child) {
      return true;
    }
    TreeView_Expand(tree_.hwnd(), item, expanded ? TVE_COLLAPSE : TVE_EXPAND);
    return true;
  }
  case cmd::kViewSelectAll:
    if (!SelectAllInFocusedList()) {
      HWND focus = GetFocus();
      if (focus) {
        SendMessageW(focus, EM_SETSEL, 0, -1);
      }
    }
    return true;
  case cmd::kEditInvertSelection:
    InvertSelectionInFocusedList();
    return true;
  case cmd::kViewToolbar:
    show_toolbar_ = !show_toolbar_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewKeyTree:
    show_tree_ = !show_tree_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewKeysInList:
    show_keys_in_list_ = !show_keys_in_list_;
    BuildMenus();
    UpdateValueListForNode(current_node_);
    SaveSettings();
    return true;
  case cmd::kViewSimulatedKeys:
    show_simulated_keys_ = !show_simulated_keys_;
    BuildMenus();
    RefreshTreeSelection();
    UpdateValueListForNode(current_node_);
    SaveSettings();
    return true;
  case cmd::kViewHistory:
    show_history_ = !show_history_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewStatusBar:
    show_status_bar_ = !show_status_bar_;
    ApplyViewVisibility();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kViewExtraHives:
    show_extra_hives_ = !show_extra_hives_;
    SaveSettings();
    BuildMenus();
    if (registry_mode_ == RegistryMode::kLocal) {
      std::vector<RegistryRootEntry> roots = RegistryProvider::DefaultRoots(show_extra_hives_);
      AppendRealRegistryRoot(&roots);
      ApplyRegistryRoots(roots);
    }
    return true;
  case cmd::kViewSaveTreeState:
    if (save_tree_state_) {
      StopTreeStateWorker();
      save_tree_state_ = false;
      saved_tree_state_.selected_path.clear();
      saved_tree_state_.expanded_paths.clear();
    } else {
      save_tree_state_ = true;
      LoadTreeState();
      tree_state_restored_ = false;
      RestoreTreeState();
      StartTreeStateWorker();
      MarkTreeStateDirty();
    }
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsSaveTabs:
    save_tabs_ = !save_tabs_;
    if (!save_tabs_) {
      ClearTabsCache();
    } else {
      SaveTabs();
    }
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsReadOnly:
    read_only_ = !read_only_;
    SaveSettings();
    BuildMenus();
    if (toolbar_.hwnd()) {
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditPaste, read_only_ ? 0 : TBSTATE_ENABLED);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditDelete, read_only_ ? 0 : TBSTATE_ENABLED);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                   !read_only_ && undo_stack_.CanUndo() ? TBSTATE_ENABLED
                                                       : 0);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo,
                   !read_only_ && undo_stack_.CanRedo() ? TBSTATE_ENABLED
                                                       : 0);
    }
    return true;
  case cmd::kOptionsCompareRegistries:
    StartCompareRegistries();
    return true;
  case cmd::kViewFont: {
    FontDialogResult result = {};
    if (ShowFontDialog(hwnd_, DefaultLogFont(), !use_custom_font_, custom_font_, &result)) {
      use_custom_font_ = !result.use_default;
      custom_font_ = result.font;
      UpdateUIFont();
      SaveSettings();
    }
    return true;
  }
  case cmd::kTraceLoad23H2:
    if (RemoveTraceByLabel(L"23H2")) {
      return true;
    }
    LoadBundledTrace(L"23H2");
    return true;
  case cmd::kTraceLoad24H2:
    if (RemoveTraceByLabel(L"24H2")) {
      return true;
    }
    LoadBundledTrace(L"24H2");
    return true;
  case cmd::kTraceLoad25H2:
    if (RemoveTraceByLabel(L"25H2")) {
      return true;
    }
    LoadBundledTrace(L"25H2");
    return true;
  case cmd::kTraceLoadCustom:
    LoadTraceFromPrompt();
    return true;
  case cmd::kTraceClear:
    ClearTrace();
    return true;
  case cmd::kDefaultLoadCustom:
    LoadDefaultFromPrompt();
    return true;
  case cmd::kDefaultClear:
    ClearDefaults();
    return true;
  case cmd::kDefaultEditActive: {
    std::vector<std::wstring> active;
    active.reserve(active_defaults_.size());
    for (const auto& defaults : active_defaults_) {
      active.push_back(defaults.source_path);
    }
    std::wstring content = JoinLines(active);
    if (PromptForMultiLineText(hwnd_, L"Edit Active Defaults", L"One default path per line.", &content)) {
      std::vector<std::wstring> lines = SplitLines(content);
      active_defaults_.clear();
      for (const auto& line : lines) {
        AddDefaultFromFile(L"", line, false, false, false);
      }
      SaveActiveDefaults();
      BuildMenus();
      UpdateValueListForNode(current_node_);
      SaveSettings();
    }
    return true;
  }
  case cmd::kTraceEditActive: {
    std::vector<std::wstring> active;
    active.reserve(active_traces_.size());
    for (const auto& trace : active_traces_) {
      active.push_back(trace.source_path);
    }
    std::wstring content = JoinLines(active);
    if (PromptForMultiLineText(hwnd_, L"Edit Active Traces", L"One trace path per line.", &content)) {
      std::vector<std::wstring> lines = SplitLines(content);
      LoadTraceSettings();
      active_traces_.clear();
      for (const auto& line : lines) {
        AddTraceFromFile(L"", line, nullptr, false, false);
      }
      SaveActiveTraces();
      SaveTraceSettings();
      BuildMenus();
      RefreshTreeSelection();
      UpdateValueListForNode(current_node_);
      SaveSettings();
    }
    return true;
  }
  case cmd::kTraceGuide:
    ShellExecuteW(hwnd_, L"open", L"https://github.com/nohuto/regkit/blob/main/guides/wpr-wpa.md", nullptr, nullptr, SW_SHOWNORMAL);
    return true;
  case cmd::kWindowNew:
    ui::LaunchNewInstance();
    return true;
  case cmd::kWindowClose:
    PostMessageW(hwnd_, WM_CLOSE, 0, 0);
    return true;
  case cmd::kWindowAlwaysOnTop:
    always_on_top_ = !always_on_top_;
    ApplyAlwaysOnTop();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeSystem:
    theme_mode_ = ThemeMode::kSystem;
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeLight:
    theme_mode_ = ThemeMode::kLight;
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeDark:
    theme_mode_ = ThemeMode::kDark;
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsThemeCustom:
    ApplyThemePresetByName(active_theme_preset_, true);
    return true;
  case cmd::kOptionsThemePresets:
    ShowThemePresetsDialog();
    return true;
  case cmd::kOptionsIconSetDefault:
    icon_set_ = L"default";
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetTabler:
    icon_set_ = L"tabler";
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetFluentUi:
    icon_set_ = L"fluentui";
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetMaterialSymbols:
    icon_set_ = L"materialsymbols";
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsIconSetCustom:
    icon_set_ = L"custom";
    ReloadThemeIcons();
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsRestartAdmin:
    RestartAsAdmin();
    return true;
  case cmd::kOptionsAlwaysRunAdmin:
    always_run_as_admin_ = !always_run_as_admin_;
    if (always_run_as_admin_) {
      always_run_as_system_ = false;
      always_run_as_trustedinstaller_ = false;
    }
    SaveSettings();
    BuildMenus();
    if (always_run_as_admin_ && !util::IsProcessElevated()) {
      RestartAsAdmin();
    }
    return true;
  case cmd::kOptionsRestartSystem:
    RestartAsSystem();
    return true;
  case cmd::kOptionsAlwaysRunSystem:
    always_run_as_system_ = !always_run_as_system_;
    if (always_run_as_system_) {
      always_run_as_admin_ = false;
      always_run_as_trustedinstaller_ = false;
    }
    SaveSettings();
    BuildMenus();
    if (always_run_as_system_ && !util::IsProcessSystem()) {
      RestartAsSystem();
    }
    return true;
  case cmd::kOptionsRestartTrustedInstaller:
    RestartAsTrustedInstaller();
    return true;
  case cmd::kOptionsAlwaysRunTrustedInstaller:
    always_run_as_trustedinstaller_ = !always_run_as_trustedinstaller_;
    if (always_run_as_trustedinstaller_) {
      always_run_as_admin_ = false;
      always_run_as_system_ = false;
    }
    SaveSettings();
    BuildMenus();
    if (always_run_as_trustedinstaller_ &&
        !util::IsProcessTrustedInstaller()) {
      RestartAsTrustedInstaller();
    }
    return true;
  case cmd::kOptionsOpenDefaultRegedit:
    OpenDefaultRegedit();
    return true;
  case cmd::kCreateSimulatedKey: {
    if (!EnsureWritable()) {
      return true;
    }
    RegistryNode target;
    bool has_target = false;
    const ListRow* row = SelectedValueRow(value_list_, nullptr);
    if (row && row->kind == rowkind::kKey && row->simulated && current_node_) {
      target = MakeChildNode(*current_node_, row->extra);
      has_target = true;
    } else if (current_node_ && current_node_->simulated) {
      target = *current_node_;
      has_target = true;
    }
    if (!has_target) {
      return true;
    }
    std::wstring path = registry_path::Build(target);
    if (!CreateRegistryPath(path)) {
      ui::ShowError(hwnd_, L"Failed to create the key.");
      return true;
    }
    UpdateSimulatedChain(TreeView_GetSelection(tree_.hwnd()));
    RefreshTreeSelection();
    UpdateValueListForNode(current_node_);
    return true;
  }
  case cmd::kOptionsReplaceRegedit:
    ReplaceRegedit(!replace_regedit_);
    return true;
  case cmd::kOptionsSingleInstance:
    single_instance_ = !single_instance_;
    SaveSettings();
    BuildMenus();
    return true;
  case cmd::kOptionsHiveFileDir:
    OpenHiveFileDir();
    return true;
  case cmd::kHelpAbout:
    ui::ShowAbout(hwnd_);
    return true;
  case cmd::kHelpContents:
    ShellExecuteW(hwnd_, L"open", kRepoUrl, nullptr, nullptr, SW_SHOWNORMAL);
    return true;
  case cmd::kFavoritesAdd: {
    if (current_node_) {
      if (FavoritesStore::Add(registry_path::Build(*current_node_))) {
        RefreshFavoritesCache();
        BuildMenus();
      }
    }
    return true;
  }
  case cmd::kFavoritesRemove: {
    if (current_node_) {
      if (FavoritesStore::Remove(registry_path::Build(*current_node_))) {
        RefreshFavoritesCache();
        BuildMenus();
      }
    }
    return true;
  }
  case cmd::kFavoritesEdit: {
    if (!favorites_loaded_) {
      RefreshFavoritesCache();
    }
    std::wstring content = JoinLines(favorites_cache_);
    if (PromptForMultiLineText(hwnd_, L"Edit Favorites", kOneKeyPerLineText, &content)) {
      std::vector<std::wstring> updated = SplitLines(content);
      if (!FavoritesStore::Save(updated)) {
        ui::ShowError(hwnd_, L"Failed to save favorites.");
        return true;
      }
      favorites_cache_ = std::move(updated);
      favorites_loaded_ = true;
      BuildMenus();
    }
    return true;
  }
  case cmd::kFavoritesImport: {
    std::wstring path;
    if (!PromptOpenFilePath(hwnd_, L"Favorites Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    if (!FavoritesStore::ImportFromFile(path)) {
      ui::ShowError(hwnd_, L"Failed to import favorites.");
    } else {
      RefreshFavoritesCache();
      AppendHistoryEntry(L"Import favorites " + FileNameOnly(path), L"", path);
    }
    BuildMenus();
    return true;
  }
  case cmd::kFavoritesImportRegedit: {
    size_t imported = 0;
    std::wstring error;
    if (!FavoritesStore::ImportFromRegedit(&imported, &error)) {
      ui::ShowError(hwnd_, error.empty() ? L"Failed to import Regedit favorites." : error);
      return true;
    }
    if (imported > 0) {
      RefreshFavoritesCache();
      BuildMenus();
    }
    AppendHistoryEntry(L"Import Regedit favorites", L"", std::to_wstring(imported) + L" favorites");
    return true;
  }
  case cmd::kFavoritesExport: {
    std::wstring path;
    if (!PromptSaveFilePath(hwnd_, L"Favorites Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0", &path)) {
      return true;
    }
    if (!FavoritesStore::ExportToFile(path)) {
      ui::ShowError(hwnd_, L"Failed to export favorites.");
    } else {
      AppendHistoryEntry(L"Export favorites " + FileNameOnly(path), L"", path);
    }
    return true;
  }
  case cmd::kDefaultEditRecent: {
    std::wstring content = JoinLines(recent_default_paths_.items());
    if (PromptForMultiLineText(hwnd_, L"Edit Recent Defaults", L"One default path per line.", &content)) {
      std::vector<std::wstring> updated = SplitLines(content);
      recent_default_paths_.Replace(std::move(updated));
      NormalizeRecentDefaultList();
      SaveSettings();
      BuildMenus();
    }
    return true;
  }
  case cmd::kTraceEditRecent: {
    std::wstring content = JoinLines(recent_trace_paths_.items());
    if (PromptForMultiLineText(hwnd_, L"Edit Recent Traces", L"One trace path per line.", &content)) {
      std::vector<std::wstring> updated = SplitLines(content);
      recent_trace_paths_.Replace(std::move(updated));
      NormalizeRecentTraceList();
      SaveSettings();
      BuildMenus();
    }
    return true;
  }
  case cmd::kEditCopyKey: {
    std::wstring name;
    int index = -1;
    const ListRow* row = SelectedValueRow(value_list_, &index);
    if (row && row->kind == rowkind::kKey) {
      name = row->extra;
    } else if (current_node_) {
      name = LeafName(*current_node_);
    }
    if (!name.empty()) {
      ui::CopyTextToClipboard(hwnd_, name);
    }
    return true;
  }
  case cmd::kEditCopyValueName: {
    std::vector<ListRow> selected_rows = SelectedListRows(value_list_);
    if (selected_rows.size() != 1 || selected_rows.front().kind != rowkind::kValue) {
      return true;
    }
    const ListRow* row = &selected_rows.front();
    std::wstring name = row->extra.empty() ? L"(Default)" : row->extra;
    ui::CopyTextToClipboard(hwnd_, name);
    return true;
  }
  case cmd::kEditCopyValueData: {
    if (!current_node_) {
      return true;
    }
    const ListRow* row = SelectedValueRow(value_list_, nullptr);
    if (!row || row->kind != rowkind::kValue) {
      return true;
    }
    if (row->simulated) {
      return true;
    }
    ValueEntry entry;
    if (!GetValueEntry(*current_node_, row->extra, &entry)) {
      ui::ShowError(hwnd_, L"Failed to read value.");
      return true;
    }
    std::wstring data = value_format::DisplayData(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
    ui::CopyTextToClipboard(hwnd_, data);
    return true;
  }
  case cmd::kEditCopyKeyPath:
  case cmd::kEditCopyKeyPathAbbrev:
  case cmd::kEditCopyKeyPathRegedit:
  case cmd::kEditCopyKeyPathRegFile:
  case cmd::kEditCopyKeyPathPowerShell:
  case cmd::kEditCopyKeyPathPowerShellProvider:
  case cmd::kEditCopyKeyPathEscaped: {
    auto build_path = [&]() -> std::wstring {
      std::wstring path;
      int index = -1;
      const ListRow* row = SelectedValueRow(value_list_, &index);
      if (row && row->kind == rowkind::kKey && current_node_) {
        path = registry_path::Build(*current_node_);
        if (!row->extra.empty()) {
          path += L"\\" + row->extra;
        }
      } else if (current_node_) {
        path = registry_path::Build(*current_node_);
      }
      return path;
    };
    std::wstring path = build_path();
    if (path.empty()) {
      return true;
    }
    RegistryPathFormat format = RegistryPathFormat::kFull;
    switch (command_id) {
    case cmd::kEditCopyKeyPathAbbrev:
      format = RegistryPathFormat::kAbbrev;
      break;
    case cmd::kEditCopyKeyPathRegedit:
      format = RegistryPathFormat::kRegedit;
      break;
    case cmd::kEditCopyKeyPathRegFile:
      format = RegistryPathFormat::kRegFile;
      break;
    case cmd::kEditCopyKeyPathPowerShell:
      format = RegistryPathFormat::kPowerShellDrive;
      break;
    case cmd::kEditCopyKeyPathPowerShellProvider:
      format = RegistryPathFormat::kPowerShellProvider;
      break;
    case cmd::kEditCopyKeyPathEscaped:
      format = RegistryPathFormat::kEscaped;
      break;
    default:
      format = RegistryPathFormat::kFull;
      break;
    }
    ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(path, format));
    return true;
  }
  case cmd::kEditCopy: {
    HWND focus = GetFocus();
    if (focus == value_list_.hwnd() || focus == search_results_list_ || focus == history_list_) {
      HWND list = focus;
      int selected = ListView_GetSelectedCount(list);
      if (selected > 0) {
        std::wstring text = BuildSelectedListViewText(list);
        if (!text.empty()) {
          ui::CopyTextToClipboard(hwnd_, text);
        }
        if (list == value_list_.hwnd() && selected == 1 && current_node_) {
          int index = -1;
          const ListRow* row = SelectedValueRow(value_list_, &index);
          if (row && row->kind == rowkind::kValue) {
            ValueEntry entry;
            if (GetValueEntry(*current_node_, row->extra, &entry)) {
              clipboard_.kind = ClipboardItem::Kind::kValue;
              clipboard_.source_parent = *current_node_;
              clipboard_.name = entry.name;
              clipboard_.value = entry;
            }
          } else if (row && row->kind == rowkind::kKey) {
            RegistryNode child = MakeChildNode(*current_node_, row->extra);
            clipboard_.kind = ClipboardItem::Kind::kKey;
            clipboard_.source_parent = *current_node_;
            clipboard_.name = row->extra;
            clipboard_.key_snapshot = changes::CaptureKey(child);
          }
        } else if (list == value_list_.hwnd()) {
          clipboard_.kind = ClipboardItem::Kind::kNone;
        }
        return true;
      }
    }
    if (!current_node_) {
      return true;
    }
    int index = -1;
    const ListRow* row = SelectedValueRow(value_list_, &index);
    if (row && row->kind == rowkind::kValue) {
      ValueEntry entry;
      if (GetValueEntry(*current_node_, row->extra, &entry)) {
        clipboard_.kind = ClipboardItem::Kind::kValue;
        clipboard_.source_parent = *current_node_;
        clipboard_.name = entry.name;
        clipboard_.value = entry;
        ui::CopyTextToClipboard(hwnd_, row->name);
      } else {
        ui::ShowError(hwnd_, L"Failed to read value.");
      }
      return true;
    }
    if (row && row->kind == rowkind::kKey) {
      RegistryNode child = MakeChildNode(*current_node_, row->extra);
      clipboard_.kind = ClipboardItem::Kind::kKey;
      clipboard_.source_parent = *current_node_;
      clipboard_.name = row->extra;
      clipboard_.key_snapshot = changes::CaptureKey(child);
      ui::CopyTextToClipboard(hwnd_, registry_path::Build(child));
      return true;
    }
    clipboard_.kind = ClipboardItem::Kind::kNone;
    ui::CopyTextToClipboard(hwnd_, registry_path::Build(*current_node_));
    return true;
  }
  case cmd::kEditGoTo:
    if (address_edit_) {
      SetFocus(address_edit_);
      SendMessageW(address_edit_, EM_SETSEL, 0, -1);
    }
    return true;
  case cmd::kEditPermissions:
    if (!EnsureWritable()) {
      return true;
    }
    if (current_node_) {
      int index = -1;
      const ListRow* row = SelectedValueRow(value_list_, &index);
      if (row && row->kind == rowkind::kKey && !row->extra.empty()) {
        RegistryNode child = MakeChildNode(*current_node_, row->extra);
        ShowPermissionsDialog(child);
      } else {
        ShowPermissionsDialog(*current_node_);
      }
    }
    return true;
  case cmd::kEditFind: {
    SearchDialogResult options = last_search_;
    bool trace_available = HasActiveTraces();
    bool registry_available = std::any_of(roots_.begin(), roots_.end(), [](const RegistryRootEntry& entry) { return _wcsicmp(entry.path_name.c_str(), L"REGISTRY") == 0; });
    if (ShowSearchDialog(hwnd_, &options, trace_available, registry_available)) {
      last_search_ = options;
      StartSearch(options);
    }
    return true;
  }
  case cmd::kEditPaste: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!current_node_ || clipboard_.kind == ClipboardItem::Kind::kNone) {
      return true;
    }
    if (clipboard_.kind == ClipboardItem::Kind::kValue) {
      bool same_parent = SameNode(*current_node_, clipboard_.source_parent);
      std::wstring base_name = clipboard_.name;
      if (same_parent) {
        if (base_name.empty()) {
          base_name = L"Default - Copy";
        } else {
          base_name += L" - Copy";
        }
      }
      std::wstring unique = MakeUniqueValueName(*current_node_, base_name);
      ValueEntry new_value = clipboard_.value;
      new_value.name = unique;
      if (!RegistryProvider::SetValue(*current_node_, unique, new_value.type, new_value.data)) {
        ui::ShowError(hwnd_, L"Failed to paste value.");
      } else {
        std::wstring data_text = value_format::Data(new_value.type, new_value.data.data(), static_cast<DWORD>(new_value.data.size()));
        AppendValueHistoryEntry(L"Create value " + unique, L"", data_text, *current_node_, unique, HistoryEntry::RevertKind::kDeleteValue);
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kCreateValue;
        op.node = *current_node_;
        op.name = unique;
        op.new_value = new_value;
        PushUndo(std::move(op));
        UpdateValueListForNode(current_node_);
      }
      return true;
    }
    if (clipboard_.kind == ClipboardItem::Kind::kKey) {
      bool same_parent = SameNode(*current_node_, clipboard_.source_parent);
      std::wstring base_name = clipboard_.name;
      if (same_parent && !base_name.empty()) {
        base_name += L" - Copy";
      }
      std::wstring unique = MakeUniqueKeyName(*current_node_, base_name);
      changes::KeySnapshot snapshot = clipboard_.key_snapshot;
      snapshot.name = unique;
      if (!changes::RestoreKey(*current_node_, snapshot)) {
        ui::ShowError(hwnd_, L"Failed to paste key.");
      } else {
        AppendHistoryEntry(L"Create key " + unique, L"", L"");
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kCreateKey;
        op.node = *current_node_;
        op.name = unique;
        op.key_snapshot = snapshot;
        PushUndo(std::move(op));
        RefreshTreeSelection();
        UpdateValueListForNode(current_node_);
      }
      return true;
    }
    return true;
  }
  case cmd::kEditReplace: {
    if (!EnsureWritable()) {
      return true;
    }
    ReplaceDialogResult options = last_replace_;
    if (options.start_key.empty() && current_node_) {
      options.start_key = registry_path::Build(*current_node_);
    }
    if (ShowReplaceDialog(hwnd_, &options)) {
      last_replace_ = options;
      StartReplace(options);
    }
    return true;
  }
  case cmd::kEditUndo: {
    if (!EnsureWritable()) {
      return true;
    }
    auto operation = undo_stack_.TakeUndo();
    if (!operation) {
      return true;
    }
    if (ApplyUndoOperation(*operation, false)) {
      undo_stack_.CompleteUndo(std::move(*operation));
    } else {
      undo_stack_.CompleteRedo(std::move(*operation));
    }
    if (toolbar_.hwnd()) {
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                   undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo,
                   undo_stack_.CanRedo() ? TBSTATE_ENABLED : 0);
    }
    return true;
  }
  case cmd::kEditRedo: {
    if (!EnsureWritable()) {
      return true;
    }
    auto operation = undo_stack_.TakeRedo();
    if (!operation) {
      return true;
    }
    if (ApplyUndoOperation(*operation, true)) {
      undo_stack_.CompleteRedo(std::move(*operation));
    } else {
      undo_stack_.CompleteUndo(std::move(*operation));
    }
    if (toolbar_.hwnd()) {
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo,
                   undo_stack_.CanUndo() ? TBSTATE_ENABLED : 0);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo,
                   undo_stack_.CanRedo() ? TBSTATE_ENABLED : 0);
    }
    return true;
  }
  case cmd::kRegistryLocal:
    if (SwitchToLocalRegistry()) {
      BuildMenus();
    }
    return true;
  case cmd::kRegistryNetwork:
    if (SwitchToRemoteRegistry()) {
      BuildMenus();
    }
    return true;
  case cmd::kRegistryOffline:
    if (SwitchToOfflineRegistry()) {
      BuildMenus();
    }
    return true;
  case cmd::kNavBack:
    NavigateBack();
    return true;
  case cmd::kNavForward:
    NavigateForward();
    return true;
  case cmd::kNavUp:
    NavigateUp();
    return true;
  case cmd::kNewKey: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!current_node_) {
      return true;
    }
    std::wstring name = MakeUniqueKeyName(*current_node_, L"New Key");
    if (name.empty()) {
      return true;
    }
    if (!RegistryProvider::CreateKey(*current_node_, name)) {
      ui::ShowError(hwnd_, L"Failed to create key.");
    } else {
      AppendHistoryEntry(L"Create key " + name, L"", L"");
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kCreateKey;
      op.node = *current_node_;
      op.name = name;
      op.key_snapshot.name = name;
      PushUndo(std::move(op));
      std::wstring path = registry_path::Build(*current_node_);
      if (!path.empty()) {
        path.append(L"\\");
        path.append(name);
      }
      HWND focus = GetFocus();
      bool edit_in_list = (focus == value_list_.hwnd()) && show_keys_in_list_ && value_list_.hwnd();
      if (edit_in_list) {
        ScheduleValueListRename(rowkind::kKey, name);
        UpdateValueListForNode(current_node_);
      } else {
        std::wstring parent_path = registry_path::Build(*current_node_);
        HTREEITEM parent_item = TreeView_GetSelection(tree_.hwnd());
        if (!parent_item && !parent_path.empty()) {
          SelectTreePath(parent_path);
          parent_item = TreeView_GetSelection(tree_.hwnd());
        }

        if (parent_item) {
          TreeView_SelectItem(tree_.hwnd(), parent_item);
        }
        RefreshTreeSelection();

        HTREEITEM target = nullptr;
        if (parent_item) {
          target = FindChildByText(tree_.hwnd(), parent_item, name);
          if (target) {
            TreeView_SelectItem(tree_.hwnd(), target);
            TreeView_EnsureVisible(tree_.hwnd(), target);
          }
        }
        if (!target && !path.empty()) {
          if (SelectTreePath(path)) {
            target = TreeView_GetSelection(tree_.hwnd());
          }
        }
        if (target) {
          SetFocus(tree_.hwnd());
          TreeView_EditLabel(tree_.hwnd(), target);
        }
        UpdateValueListForNode(current_node_);
      }
    }
    return true;
  }
  case cmd::kNewString:
  case cmd::kNewExpandString:
  case cmd::kNewBinary:
  case cmd::kNewDword:
  case cmd::kNewQword:
  case cmd::kNewMultiString: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!current_node_) {
      return true;
    }
    DWORD type = REG_SZ;
    std::wstring base_name = L"New Value";
    switch (command_id) {
    case cmd::kNewExpandString:
      type = REG_EXPAND_SZ;
      base_name = L"New Expandable String Value";
      break;
    case cmd::kNewBinary:
      type = REG_BINARY;
      base_name = L"New Binary Value";
      break;
    case cmd::kNewDword:
      type = REG_DWORD;
      base_name = L"New DWORD Value";
      break;
    case cmd::kNewQword:
      type = REG_QWORD;
      base_name = L"New QWORD Value";
      break;
    case cmd::kNewMultiString:
      type = REG_MULTI_SZ;
      base_name = L"New Multi-String Value";
      break;
    default:
      type = REG_SZ;
      break;
    }
    std::wstring value_name = MakeUniqueValueName(*current_node_, base_name);
    if (value_name.empty()) {
      return true;
    }
    std::vector<BYTE> data;
    if (type == REG_SZ || type == REG_EXPAND_SZ) {
      data.resize(sizeof(wchar_t), 0);
    } else if (type == REG_MULTI_SZ) {
      data.resize(sizeof(wchar_t) * 2, 0);
    } else if (type == REG_DWORD) {
      data.resize(sizeof(DWORD), 0);
    } else if (type == REG_QWORD) {
      data.resize(sizeof(unsigned long long), 0);
    }
    if (!RegistryProvider::SetValue(*current_node_, value_name, type, data)) {
      ui::ShowError(hwnd_, L"Failed to set value.");
    } else {
      std::wstring data_text = value_format::Data(type, data.data(), static_cast<DWORD>(data.size()));
      AppendValueHistoryEntry(L"Create value " + value_name, L"", data_text, *current_node_, value_name, HistoryEntry::RevertKind::kDeleteValue);
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kCreateValue;
      op.node = *current_node_;
      op.name = value_name;
      op.new_value.name = value_name;
      op.new_value.type = type;
      op.new_value.data = data;
      PushUndo(std::move(op));
      ScheduleValueListRename(rowkind::kValue, value_name);
      UpdateValueListForNode(current_node_);
    }
    return true;
  }
  case cmd::kEditModify:
  case cmd::kEditModifyBinary: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!current_node_) {
      return true;
    }
    std::vector<ListRow> selected_rows = SelectedListRows(value_list_);
    if (selected_rows.size() != 1 || selected_rows.front().kind != rowkind::kValue) {
      return true;
    }
    const ListRow* row = &selected_rows.front();
    ValueEntry entry;
    if (!GetValueEntry(*current_node_, row->extra, &entry)) {
      if (HasActiveTraces() && (row->type.empty() || EqualsInsensitive(row->type, L"TRACE"))) {
        bool needs_create = current_node_->simulated;
        DWORD type = REG_SZ;
        std::vector<BYTE> data;
        if (!PromptForCustomValue(hwnd_, row->extra, &type, &data)) {
          return true;
        }
        if (needs_create) {
          std::wstring path = registry_path::Build(*current_node_);
          if (!CreateRegistryPath(path)) {
            ui::ShowError(hwnd_, L"Failed to create the key.");
            return true;
          }
          UpdateSimulatedChain(TreeView_GetSelection(tree_.hwnd()));
        }
        if (!RegistryProvider::SetValue(*current_node_, row->extra, type, data)) {
          ui::ShowError(hwnd_, L"Failed to set value.");
          return true;
        }
        std::wstring display_name = row->extra.empty() ? L"(Default)" : row->extra;
        std::wstring data_text = value_format::Data(type, data.data(), static_cast<DWORD>(data.size()));
        AppendValueHistoryEntry(L"Create value " + display_name, L"", data_text, *current_node_, row->extra, HistoryEntry::RevertKind::kDeleteValue);
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kCreateValue;
        op.node = *current_node_;
        op.name = row->extra;
        op.new_value.name = row->extra;
        op.new_value.type = type;
        op.new_value.data = data;
        PushUndo(std::move(op));
        RefreshTreeSelection();
        UpdateValueListForNode(current_node_);
        return true;
      }
      ui::ShowError(hwnd_, L"Failed to read value.");
      return true;
    }
    std::wstring old_text = value_format::Data(entry.type, entry.data.data(), static_cast<DWORD>(entry.data.size()));
    DWORD base_type = value_format::NormalizeType(entry.type);
    bool supports_extended_dialog = base_type == REG_SZ || base_type == REG_EXPAND_SZ || base_type == REG_MULTI_SZ || base_type == REG_DWORD || base_type == REG_DWORD_BIG_ENDIAN || base_type == REG_QWORD || base_type == REG_LINK;
    std::vector<BYTE> new_data;
    if (command_id == cmd::kEditModifyBinary || base_type == REG_BINARY || base_type == REG_NONE || base_type == REG_RESOURCE_LIST || base_type == REG_FULL_RESOURCE_DESCRIPTOR || base_type == REG_RESOURCE_REQUIREMENTS_LIST) {
      std::wstring type_label = value_format::TypeName(entry.type);
      if (!PromptForBinary(hwnd_, entry.name, entry.data, &new_data, type_label.c_str())) {
        return true;
      }
    } else if (command_id == cmd::kEditModify && supports_extended_dialog) {
      std::wstring type_label = value_format::TypeName(entry.type);
      if (!PromptForFlaggedValue(hwnd_, entry.name, base_type, entry.data, type_label, &new_data)) {
        return true;
      }
    } else {
      std::wstring type_label = value_format::TypeName(entry.type);
      if (!PromptForBinary(hwnd_, entry.name, entry.data, &new_data, type_label.c_str())) {
        return true;
      }
    }
    if (!RegistryProvider::SetValue(*current_node_, entry.name, entry.type, new_data)) {
      ui::ShowError(hwnd_, L"Failed to update value.");
    } else {
      std::wstring new_text = value_format::Data(entry.type, new_data.data(), static_cast<DWORD>(new_data.size()));
      AppendValueHistoryEntry(L"Modify value " + entry.name, old_text, new_text, *current_node_, entry.name, HistoryEntry::RevertKind::kSetValue, &entry);
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kModifyValue;
      op.node = *current_node_;
      op.old_value = entry;
      op.new_value = entry;
      op.new_value.data = new_data;
      PushUndo(std::move(op));
      UpdateValueListForNode(current_node_);
    }
    return true;
  }
  case cmd::kEditModifyComment: {
    if (!current_node_) {
      return true;
    }
    std::vector<ListRow> selected_rows = SelectedListRows(value_list_);
    if (selected_rows.empty()) {
      return true;
    }
    if (std::any_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& row) {
          return row.kind != rowkind::kValue || row.simulated;
        })) {
      return true;
    }
    EditValueComments(selected_rows);
    return true;
  }
  case cmd::kEditRename: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!current_node_) {
      return true;
    }
    HWND focus = GetFocus();
    std::vector<ListRow> selected_rows = SelectedListRows(value_list_);
    if (focus == value_list_.hwnd() && selected_rows.size() > 1) {
      return true;
    }

    const ListRow* row = selected_rows.empty() ? nullptr : &selected_rows.front();
    if (focus == tree_.hwnd() || (!row && current_node_)) {
      if (current_node_->subkey.empty()) {
        return true;
      }
      HTREEITEM selected = TreeView_GetSelection(tree_.hwnd());
      if (selected) {
        SetFocus(tree_.hwnd());
        TreeView_EditLabel(tree_.hwnd(), selected);
      }
      return true;
    }
    if (row && row->kind == rowkind::kKey) {
      if (row->extra.empty()) {
        return true;
      }
      int index = -1;
      SelectedValueRow(value_list_, &index);
      if (focus == value_list_.hwnd() && index >= 0) {
        SetFocus(value_list_.hwnd());
        ListView_EditLabel(value_list_.hwnd(), index);
        return true;
      }
      std::wstring path = registry_path::Build(*current_node_);
      if (!row->extra.empty()) {
        path.append(L"\\");
        path.append(row->extra);
      }
      if (SelectTreePath(path)) {
        HTREEITEM selected = TreeView_GetSelection(tree_.hwnd());
        if (selected) {
          SetFocus(tree_.hwnd());
          TreeView_EditLabel(tree_.hwnd(), selected);
        }
      }
      return true;
    }
    if (row && row->kind == rowkind::kValue) {
      if (row->simulated) {
        return true;
      }
      if (row->extra.empty()) {
        return true;
      }
      int index = -1;
      SelectedValueRow(value_list_, &index);
      if (index >= 0) {
        SetFocus(value_list_.hwnd());
        ListView_EditLabel(value_list_.hwnd(), index);
      }
      return true;
    }
    return true;
  }
  case cmd::kEditDelete: {
    if (!EnsureWritable()) {
      return true;
    }
    if (!current_node_) {
      return true;
    }
    HWND focus = GetFocus();
    bool tree_focus = (focus == tree_.hwnd());
    if (tree_focus && current_node_ && !current_node_->subkey.empty()) {
      std::wstring name = LeafName(*current_node_);
      if (!ui::ConfirmDelete(hwnd_, L"Delete Key", name)) {
        return true;
      }
      RegistryNode target = *current_node_;
      RegistryNode parent = target;
      size_t pos = parent.subkey.rfind(L'\\');
      parent.subkey = (pos == std::wstring::npos) ? L"" : parent.subkey.substr(0, pos);
      changes::KeySnapshot snapshot = changes::CaptureKey(target);
      if (!RegistryProvider::DeleteKey(target)) {
        ui::ShowError(hwnd_, L"Failed to delete key.");
      } else {
        AppendHistoryEntry(L"Delete key " + name, name, L"");
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteKey;
        op.node = parent;
        op.name = name;
        op.key_snapshot = std::move(snapshot);
        PushUndo(std::move(op));
        std::wstring parent_path = registry_path::Build(parent);
        bool selected_parent = false;
        if (!parent_path.empty()) {
          selected_parent = SelectTreePath(parent_path);
        }
        RefreshTreeSelection();
        if (!selected_parent) {
          UpdateValueListForNode(current_node_);
        }
      }
      return true;
    }

    std::vector<ListRow> selected_rows = SelectedListRows(value_list_);
    if (focus == value_list_.hwnd() && selected_rows.size() > 1) {
      const bool all_values = std::all_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& selected) {
        return selected.kind == rowkind::kValue && !selected.simulated;
      });
      if (!all_values) {
        ui::ShowWarning(hwnd_, L"Bulk deletion only supports registry values.");
        return true;
      }

      std::vector<ValueEntry> entries;
      entries.reserve(selected_rows.size());
      for (const auto& selected : selected_rows) {
        ValueEntry entry;
        if (!GetValueEntry(*current_node_, selected.extra, &entry)) {
          ui::ShowError(hwnd_, L"Failed to read all selected values.");
          return true;
        }
        entries.push_back(std::move(entry));
      }

      std::wstring target = std::to_wstring(entries.size()) + L" selected values";
      if (!ui::ConfirmDelete(hwnd_, L"Delete Values", target)) {
        return true;
      }

      size_t deleted = 0;
      for (auto& entry : entries) {
        if (!RegistryProvider::DeleteValue(*current_node_, entry.name)) {
          continue;
        }
        std::wstring display_name = entry.name.empty() ? L"(Default)" : entry.name;
        AppendValueHistoryEntry(L"Delete value " + display_name, display_name, L"", *current_node_, entry.name, HistoryEntry::RevertKind::kSetValue, &entry);
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteValue;
        op.node = *current_node_;
        op.old_value = std::move(entry);
        PushUndo(std::move(op));
        ++deleted;
      }

      if (deleted > 0) {
        MarkOfflineDirty();
        UpdateValueListForNode(current_node_);
      }
      if (deleted != selected_rows.size()) {
        std::wstring message = L"Deleted " + std::to_wstring(deleted) + L" of " +
                               std::to_wstring(selected_rows.size()) + L" selected values.";
        ui::ShowError(hwnd_, message);
      }
      return true;
    }

    const ListRow* row = selected_rows.empty() ? nullptr : &selected_rows.front();
    if (row && row->kind == rowkind::kKey) {
      if (!ui::ConfirmDelete(hwnd_, L"Delete Key", row->extra)) {
        return true;
      }
      RegistryNode child = MakeChildNode(*current_node_, row->extra);
      changes::KeySnapshot snapshot = changes::CaptureKey(child);
      if (!RegistryProvider::DeleteKey(child)) {
        ui::ShowError(hwnd_, L"Failed to delete key.");
      } else {
        AppendHistoryEntry(L"Delete key " + row->extra, row->extra, L"");
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteKey;
        op.node = *current_node_;
        op.name = row->extra;
        op.key_snapshot = std::move(snapshot);
        PushUndo(std::move(op));
        RefreshTreeSelection();
        UpdateValueListForNode(current_node_);
      }
      return true;
    }
    if (row && row->kind == rowkind::kValue) {
      if (row->simulated) {
        return true;
      }
      std::wstring display_name = row->extra.empty() ? L"(Default)" : row->extra;
      if (!ui::ConfirmDelete(hwnd_, L"Delete Value", display_name)) {
        return true;
      }
      ValueEntry entry;
      if (!GetValueEntry(*current_node_, row->extra, &entry)) {
        ui::ShowError(hwnd_, L"Failed to read value.");
        return true;
      }
      if (!RegistryProvider::DeleteValue(*current_node_, row->extra)) {
        ui::ShowError(hwnd_, L"Failed to delete value.");
      } else {
        AppendValueHistoryEntry(L"Delete value " + display_name, display_name, L"", *current_node_, row->extra, HistoryEntry::RevertKind::kSetValue, &entry);
        MarkOfflineDirty();
        changes::UndoOperation op;
        op.type = changes::UndoOperation::Type::kDeleteValue;
        op.node = *current_node_;
        op.old_value = std::move(entry);
        PushUndo(std::move(op));
        UpdateValueListForNode(current_node_);
      }
      return true;
    }
    return true;
  }
  default:
    break;
  }

  return false;
}

void MainWindow::StartCompareRegistries() {
  CompareDialogDefaults defaults;
  defaults.registry_roots.reserve(roots_.size());
  std::unordered_set<std::wstring> seen_roots;
  for (const auto& root : roots_) {
    if (root.path_name.empty()) {
      continue;
    }
    std::wstring key = ToLower(root.path_name);
    if (seen_roots.insert(key).second) {
      defaults.registry_roots.push_back(root.path_name);
    }
  }
  if (defaults.registry_roots.empty()) {
    defaults.registry_roots = {L"HKEY_LOCAL_MACHINE", L"HKEY_CURRENT_USER", L"HKEY_CLASSES_ROOT", L"HKEY_USERS", L"HKEY_CURRENT_CONFIG"};
  }

  CompareDialogSelection left;
  CompareDialogSelection right;
  left.type = CompareSourceType::kRegistry;
  right.type = CompareSourceType::kRegistry;
  left.recursive = true;
  right.recursive = true;
  if (current_node_) {
    std::wstring root_name = current_node_->root_name.empty() ? registry_path::RootName(current_node_->root) : current_node_->root_name;
    left.root = root_name;
    right.root = root_name;
    left.path = current_node_->subkey;
    right.path = current_node_->subkey;
  } else if (!defaults.registry_roots.empty()) {
    left.root = defaults.registry_roots.front();
    right.root = defaults.registry_roots.front();
  }
  defaults.left = left;
  defaults.right = right;

  CompareDialogResult selection;
  if (!ShowCompareDialog(hwnd_, defaults, &selection)) {
    return;
  }

  auto normalize_base = [&](const CompareDialogSelection& sel, std::wstring* out_base) -> bool {
    if (!out_base) {
      return false;
    }
    if (sel.type == CompareSourceType::kRegistry) {
      std::wstring base;
      if (!sel.path.empty()) {
        std::wstring normalized_path = NormalizeRegistryPath(sel.path);
        if (StartsWithInsensitive(normalized_path, L"HKEY_") || StartsWithInsensitive(normalized_path, L"REGISTRY")) {
          base = normalized_path;
        }
      }
      if (base.empty()) {
        base = sel.root;
        if (!sel.path.empty()) {
          base += L"\\" + sel.path;
        }
        base = NormalizeRegistryPath(base);
      }
      if (base.empty()) {
        return false;
      }
      *out_base = base;
      return true;
    }
    std::wstring base = NormalizeRegistryPath(sel.key_path);
    if (base.empty()) {
      return false;
    }
    *out_base = base;
    return true;
  };

  auto build_snapshot =
      [&](const CompareDialogSelection& source,
          search::compare::Snapshot* snapshot,
          std::wstring* error) -> bool {
    std::wstring base;
    if (!normalize_base(source, &base)) {
      if (error) {
        *error = L"Invalid registry path.";
      }
      return false;
    }
    if (source.type == CompareSourceType::kRegFile) {
      return search::compare::LoadRegFile(
          source.file_path, base, source.recursive,
          [this](const std::wstring& path) {
            return NormalizeRegistryPath(path);
          },
          snapshot, error);
    }

    RegistryNode node;
    KeyInfo info = {};
    if (!ResolvePathToNode(base, &node) ||
        !RegistryProvider::QueryKeyInfo(node, &info)) {
      if (error) {
        *error = L"Registry path not found: " + base;
      }
      return false;
    }
    return search::compare::CaptureRegistry(
        base, node, source.recursive, snapshot);
  };

  search::compare::Snapshot left_snapshot;
  search::compare::Snapshot right_snapshot;
  std::wstring error;
  if (!build_snapshot(selection.left, &left_snapshot, &error)) {
    if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return;
  }
  error.clear();
  if (!build_snapshot(selection.right, &right_snapshot, &error)) {
    if (!error.empty()) {
      ui::ShowError(hwnd_, error);
    }
    return;
  }

  std::vector<search::Result> results =
      search::compare::Diff(left_snapshot, right_snapshot);
  std::wstring tab_label = L"Registry Comparision";

  SearchTab tab;
  tab.label = std::move(tab_label);
  tab.results = std::move(results);
  tab.is_compare = true;
  search_tabs_.push_back(std::move(tab));
  int search_index = static_cast<int>(search_tabs_.size() - 1);
  TCITEMW item = {};
  item.mask = TCIF_TEXT;
  item.pszText = const_cast<wchar_t*>(search_tabs_.back().label.c_str());
  int tab_index = TabCtrl_GetItemCount(tab_);
  TabCtrl_InsertItem(tab_, tab_index, &item);
  tabs_.push_back({TabEntry::Kind::kSearch, search_index});

  UpdateTabWidth();
  TabCtrl_SetCurSel(tab_, tab_index);
  active_search_tab_index_ = tab_index;
  UpdateSearchResultsView();
  ApplyViewVisibility();
  UpdateStatus();
}

void MainWindow::PrepareMenusForOwnerDraw(HMENU menu, bool is_menu_bar) {
  if (!menu) {
    return;
  }
  HDC hdc = GetDC(hwnd_);
  HFONT old_font = nullptr;
  if (hdc && ui_font_) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }

  auto prepare = [&](auto&& self, HMENU current, bool menu_bar) -> void {
    if (menu_bar) {
      MENUINFO menu_info = {};
      menu_info.cbSize = sizeof(menu_info);
      menu_info.fMask = MIM_BACKGROUND;
      menu_info.hbrBack = Theme::Current().BackgroundBrush();
      SetMenuInfo(current, &menu_info);
    }

    int count = GetMenuItemCount(current);
    for (int i = 0; i < count; ++i) {
      MENUITEMINFOW info = {};
      info.cbSize = sizeof(info);
      info.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_SUBMENU | MIIM_ID;
      wchar_t text[256] = {};
      info.dwTypeData = text;
      info.cch = static_cast<UINT>(_countof(text));
      if (!GetMenuItemInfoW(current, i, TRUE, &info)) {
        continue;
      }

      if (menu_bar) {
        auto data = std::make_unique<MenuItemData>();
        data->text = text;
        size_t tab_pos = data->text.find(L'\t');
        if (tab_pos != std::wstring::npos) {
          data->left_text = data->text.substr(0, tab_pos);
          data->right_text = data->text.substr(tab_pos + 1);
        } else {
          data->left_text = data->text;
        }
        data->separator = (info.fType & MFT_SEPARATOR) != 0;
        data->has_submenu = info.hSubMenu != nullptr;
        data->is_menu_bar = menu_bar;
        if (data->separator) {
          data->width = 4;
          data->height = 8;
        } else {
          SIZE left_size = {};
          SIZE right_size = {};
          if (hdc) {
            GetTextExtentPoint32W(hdc, data->left_text.c_str(), static_cast<int>(data->left_text.size()), &left_size);
            if (!data->right_text.empty()) {
              GetTextExtentPoint32W(hdc, data->right_text.c_str(), static_cast<int>(data->right_text.size()), &right_size);
            }
          }
          int height = data->is_menu_bar ? 18 : 22;
          int padding = data->is_menu_bar ? 2 : 12;
          int shortcut_gap = (!data->is_menu_bar && !data->right_text.empty()) ? 24 : 0;
          int extra = (!data->is_menu_bar && data->has_submenu) ? 12 : 6;
          data->height = height;
          data->width = static_cast<int>(left_size.cx) + static_cast<int>(right_size.cx) + padding + shortcut_gap + extra;
        }

        MenuItemData* raw = data.get();
        menu_items_.push_back(std::move(data));

        info.fMask = MIIM_FTYPE | MIIM_DATA;
        info.fType |= MFT_OWNERDRAW;
        info.dwItemData = reinterpret_cast<ULONG_PTR>(raw);
        SetMenuItemInfoW(current, i, TRUE, &info);
      }

      if (info.hSubMenu) {
        if (menu_bar) {
          self(self, info.hSubMenu, false);
        }
      }
    }
  };

  prepare(prepare, menu, is_menu_bar);

  if (hdc && old_font) {
    SelectObject(hdc, old_font);
  }
  if (hdc) {
    ReleaseDC(hwnd_, hdc);
  }
}

void MainWindow::OnMeasureMenuItem(MEASUREITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  auto* data = reinterpret_cast<MenuItemData*>(info->itemData);
  if (!data) {
    return;
  }
  if (data->width > 0 && data->height > 0) {
    info->itemWidth = static_cast<UINT>(data->width);
    info->itemHeight = static_cast<UINT>(data->height);
    return;
  }
  if (data->separator) {
    info->itemHeight = 8;
    info->itemWidth = 4;
    return;
  }

  HDC hdc = GetDC(hwnd_);
  HFONT old = nullptr;
  if (ui_font_) {
    old = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }
  SIZE size = {};
  GetTextExtentPoint32W(hdc, data->text.c_str(), static_cast<int>(data->text.size()), &size);
  if (old) {
    SelectObject(hdc, old);
  }
  ReleaseDC(hwnd_, hdc);

  int height = data->is_menu_bar ? 18 : 22;
  int padding = data->is_menu_bar ? 2 : 12;
  int extra = (!data->is_menu_bar && data->has_submenu) ? 12 : 6;
  info->itemHeight = height;
  info->itemWidth = size.cx + padding + extra;
}

void MainWindow::OnDrawMenuItem(const DRAWITEMSTRUCT* info) {
  if (!info) {
    return;
  }
  auto* data = reinterpret_cast<MenuItemData*>(info->itemData);
  if (!data) {
    return;
  }
  const Theme& theme = Theme::Current();
  HDC hdc = info->hDC;
  RECT rect = info->rcItem;

  if (data->separator) {
    HPEN pen = GetCachedPen(theme.BorderColor());
    HPEN old = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
    int y = (rect.top + rect.bottom) / 2;
    MoveToEx(hdc, rect.left + 8, y, nullptr);
    LineTo(hdc, rect.right - 8, y);
    SelectObject(hdc, old);
    return;
  }

  bool selected = (info->itemState & (ODS_SELECTED | ODS_HOTLIGHT)) != 0;
  bool disabled = (info->itemState & ODS_DISABLED) != 0;
  bool checked = (info->itemState & ODS_CHECKED) != 0;
  COLORREF bg = data->is_menu_bar ? theme.BackgroundColor() : theme.PanelColor();
  COLORREF fg = theme.TextColor();
  if (selected) {
    if (data->is_menu_bar) {
      bg = theme.HoverColor();
    } else {
      bg = theme.SelectionColor();
      fg = theme.SelectionTextColor();
    }
  } else if (disabled) {
    fg = theme.MutedTextColor();
  }

  HBRUSH bg_brush = nullptr;
  if (selected) {
    bg_brush = GetCachedBrush(bg);
  } else {
    bg_brush = data->is_menu_bar ? theme.BackgroundBrush() : theme.PanelBrush();
  }
  FillRect(hdc, &rect, bg_brush);

  SetBkMode(hdc, TRANSPARENT);
  SetTextColor(hdc, fg);
  HFONT old_font = nullptr;
  if (ui_font_) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, ui_font_));
  }
  RECT text_rect = rect;
  int left_padding = data->is_menu_bar ? 0 : 12;
  int right_padding = data->is_menu_bar ? 0 : (data->has_submenu ? 12 : 6);
  text_rect.left += left_padding;
  text_rect.right -= right_padding;
  if (checked && !data->is_menu_bar) {
    int mid_y = (rect.top + rect.bottom) / 2;
    HPEN pen = GetCachedPen(fg, 1);
    HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
    MoveToEx(hdc, rect.left + 8, mid_y, nullptr);
    LineTo(hdc, rect.left + 11, mid_y + 3);
    LineTo(hdc, rect.left + 16, mid_y - 3);
    SelectObject(hdc, old_pen);
  }
  UINT format = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS;
  if (data->is_menu_bar) {
    format |= DT_CENTER;
  }
  if (!data->right_text.empty() && !data->is_menu_bar) {
    SIZE right_size = {};
    GetTextExtentPoint32W(hdc, data->right_text.c_str(), static_cast<int>(data->right_text.size()), &right_size);
    RECT right_rect = rect;
    right_rect.right -= right_padding;
    right_rect.left = right_rect.right - right_size.cx;
    RECT left_rect = text_rect;
    left_rect.right = right_rect.left - 12;
    DrawTextW(hdc, data->left_text.c_str(), -1, &left_rect, format);
    DrawTextW(hdc, data->right_text.c_str(), -1, &right_rect, DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_RIGHT);
  } else {
    DrawTextW(hdc, data->left_text.c_str(), -1, &text_rect, format);
  }
  if (old_font) {
    SelectObject(hdc, old_font);
  }

  if (data->has_submenu && !data->is_menu_bar) {
    int rect_h = rect.bottom - rect.top;
    int arrow_size = rect_h - 8;
    if (arrow_size < 6) {
      arrow_size = 6;
    } else if (arrow_size > 10) {
      arrow_size = 10;
    }
    RECT arrow_rect = rect;
    arrow_rect.right = rect.right - 6;
    arrow_rect.left = arrow_rect.right - arrow_size;
    arrow_rect.top = rect.top + (rect_h - arrow_size) / 2;
    arrow_rect.bottom = arrow_rect.top + arrow_size;
    HFONT arrow_font = CreateFontW(-arrow_size, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_DONTCARE, L"Marlett");
    HFONT old_arrow = nullptr;
    if (arrow_font) {
      old_arrow = reinterpret_cast<HFONT>(SelectObject(hdc, arrow_font));
    }
    COLORREF arrow_color = disabled ? theme.MutedTextColor() : fg;
    SetTextColor(hdc, arrow_color);
    DrawTextW(hdc, L"8", -1, &arrow_rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_NOPREFIX | DT_NOCLIP);
    if (old_arrow) {
      SelectObject(hdc, old_arrow);
    }
    if (arrow_font) {
      DeleteObject(arrow_font);
    }
  }
}

void MainWindow::RecordNavigation(const std::wstring& path) {
  if (path.empty()) {
    return;
  }
  if (nav_is_programmatic_) {
    nav_is_programmatic_ = false;
    return;
  }
  if (nav_index_ >= 0 && nav_index_ < static_cast<int>(nav_history_.size()) && nav_history_[static_cast<size_t>(nav_index_)] == path) {
    return;
  }
  if (nav_index_ + 1 < static_cast<int>(nav_history_.size())) {
    nav_history_.erase(nav_history_.begin() + nav_index_ + 1, nav_history_.end());
  }
  nav_history_.push_back(path);
  nav_index_ = static_cast<int>(nav_history_.size()) - 1;
  UpdateNavigationButtons();
}

void MainWindow::NavigateBack() {
  if (nav_index_ <= 0) {
    return;
  }
  --nav_index_;
  nav_is_programmatic_ = true;
  SelectTreePath(nav_history_[static_cast<size_t>(nav_index_)]);
  UpdateNavigationButtons();
}

void MainWindow::NavigateForward() {
  if (nav_index_ + 1 >= static_cast<int>(nav_history_.size())) {
    return;
  }
  ++nav_index_;
  nav_is_programmatic_ = true;
  SelectTreePath(nav_history_[static_cast<size_t>(nav_index_)]);
  UpdateNavigationButtons();
}

void MainWindow::NavigateUp() {
  if (!current_node_) {
    return;
  }
  if (current_node_->subkey.empty()) {
    return;
  }
  std::wstring path = registry_path::Build(*current_node_);
  size_t pos = path.rfind(L'\\');
  if (pos == std::wstring::npos) {
    return;
  }
  std::wstring parent = path.substr(0, pos);
  nav_is_programmatic_ = true;
  SelectTreePath(parent);
  UpdateNavigationButtons();
}

void MainWindow::UpdateNavigationButtons() {
  if (!toolbar_.hwnd()) {
    return;
  }
  SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kNavBack, nav_index_ > 0 ? TBSTATE_ENABLED : 0);
  SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kNavForward, nav_index_ + 1 < static_cast<int>(nav_history_.size()) ? TBSTATE_ENABLED : 0);
  SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kNavUp, (current_node_ && !current_node_->subkey.empty()) ? TBSTATE_ENABLED : 0);
}

void MainWindow::ShowTreeContextMenu(POINT screen_pt) {
  if (!tree_.hwnd()) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(tree_.hwnd(), &client_pt);
  TVHITTESTINFO hit = {};
  hit.pt = client_pt;
  HTREEITEM item = TreeView_HitTest(tree_.hwnd(), &hit);
  if (item) {
    TreeView_SelectItem(tree_.hwnd(), item);
  }
  SetFocus(tree_.hwnd());
  HTREEITEM target = item ? item : TreeView_GetSelection(tree_.hwnd());
  RegistryNode* node = tree_.NodeFromItem(target);

  HMENU menu = CreatePopupMenu();
  bool has_node = node != nullptr;
  bool can_rename = has_node && !node->subkey.empty();
  bool is_simulated = has_node && node->simulated;
  bool can_modify = !read_only_;
  UINT edit_flags = MF_STRING | (has_node ? 0 : MF_GRAYED);
  UINT modify_flags = MF_STRING | ((has_node && can_modify) ? 0 : MF_GRAYED);
  UINT rename_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);
  UINT delete_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);

  bool expanded = false;
  bool can_toggle = false;
  if (target) {
    TVITEMW tvi = {};
    tvi.hItem = target;
    tvi.mask = TVIF_STATE | TVIF_CHILDREN;
    tvi.stateMask = TVIS_EXPANDED;
    if (TreeView_GetItem(tree_.hwnd(), &tvi)) {
      expanded = (tvi.state & TVIS_EXPANDED) != 0;
      bool has_child = TreeView_GetChild(tree_.hwnd(), target) != nullptr || tvi.cChildren != 0;
      can_toggle = expanded || has_child;
    }
  }
  std::wstring expand_label = expanded ? L"Collapse Key" : L"Expand Key";
  UINT expand_flags = MF_STRING | (can_toggle ? 0 : MF_GRAYED);
  auto equals_insensitive = [](const std::wstring& left, const wchar_t* right) -> bool { return _wcsicmp(left.c_str(), right) == 0; };
  bool can_open_hive = false;
  if (has_node) {
    bool is_root = false;
    std::wstring hive_path = LookupHivePath(*node, &is_root);
    if (!hive_path.empty() && is_root) {
      if (node->subkey.empty() && (node->root == HKEY_CURRENT_USER || equals_insensitive(node->root_name, L"HKEY_CURRENT_USER"))) {
        can_open_hive = false;
      } else {
        can_open_hive = true;
      }
    }
  }

  HMENU new_value = CreatePopupMenu();
  AppendMenuW(new_value, MF_STRING, cmd::kNewString, L"String Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewBinary, L"Binary Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
  AppendMenuW(new_value, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");

  AppendMenuW(menu, edit_flags, cmd::kEditCopyKey, L"Copy Key Name");
  AppendMenuW(menu, edit_flags, cmd::kEditCopyKeyPath, L"Copy Key Path");
  AppendMenuW(menu, MF_POPUP | (has_node ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
  if (!is_simulated) {
    AppendMenuW(menu, modify_flags, cmd::kEditPermissions, L"Permissions...");
    if (can_open_hive) {
      AppendMenuW(menu, MF_STRING, cmd::kOptionsHiveFileDir, L"Open Hive File");
    }
  }
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, expand_flags, cmd::kTreeToggleExpand, expand_label.c_str());
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  if (is_simulated) {
    AppendMenuW(menu, modify_flags, cmd::kCreateSimulatedKey, L"Create Key");
  } else {
    AppendMenuW(menu, modify_flags, cmd::kNewKey, L"New Key");
    AppendMenuW(menu, MF_POPUP | ((has_node && can_modify) ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(new_value), L"New Value");
  }
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  if (!is_simulated) {
    AppendMenuW(menu, edit_flags, cmd::kFileExport, L"Export...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  }
  AppendMenuW(menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
  if (!is_simulated) {
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, rename_flags, cmd::kEditRename, L"Rename");
    AppendMenuW(menu, delete_flags, cmd::kEditDelete, L"Delete");
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command != 0) {
    HandleMenuCommand(command);
  }
}

void MainWindow::ShowValueContextMenu(POINT screen_pt) {
  if (!value_list_.hwnd()) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(value_list_.hwnd(), &client_pt);
  LVHITTESTINFO hit = {};
  hit.pt = client_pt;
  int index = ListView_HitTest(value_list_.hwnd(), &hit);
  const ListRow* row = nullptr;
  if (index >= 0) {
    const UINT state = ListView_GetItemState(value_list_.hwnd(), index, LVIS_SELECTED);
    if ((state & LVIS_SELECTED) == 0) {
      ListView_SetItemState(value_list_.hwnd(), -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    }
    ListView_SetItemState(value_list_.hwnd(), index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    row = value_list_.RowAt(index);
  }
  SetFocus(value_list_.hwnd());

  HMENU menu = CreatePopupMenu();
  if (row && row->kind == rowkind::kKey) {
    auto equals_insensitive = [](const std::wstring& left, const wchar_t* right) -> bool { return _wcsicmp(left.c_str(), right) == 0; };
    bool is_simulated = row->simulated;
    bool can_rename = !row->extra.empty();
    bool can_modify = !read_only_;
    UINT edit_flags = MF_STRING;
    UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
    UINT rename_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);
    UINT delete_flags = MF_STRING | ((can_rename && can_modify) ? 0 : MF_GRAYED);
    UINT expand_flags = MF_STRING | MF_GRAYED;
    std::wstring expand_label = L"Expand Key";
    bool can_open_hive = false;
    if (current_node_) {
      RegistryNode target = *current_node_;
      if (!row->extra.empty()) {
        target = MakeChildNode(*current_node_, row->extra);
      }
      bool is_root = false;
      std::wstring hive_path = LookupHivePath(target, &is_root);
      if (!hive_path.empty() && is_root) {
        if (target.subkey.empty() && (target.root == HKEY_CURRENT_USER || equals_insensitive(target.root_name, L"HKEY_CURRENT_USER"))) {
          can_open_hive = false;
        } else {
          can_open_hive = true;
        }
      }
    }

    HMENU new_value = CreatePopupMenu();
    AppendMenuW(new_value, MF_STRING, cmd::kNewString, L"String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewBinary, L"Binary Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");

    AppendMenuW(menu, edit_flags, cmd::kEditCopyKey, L"Copy Key Name");
    AppendMenuW(menu, edit_flags, cmd::kEditCopyKeyPath, L"Copy Key Path");
    AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
    if (!is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kEditPermissions, L"Permissions...");
      if (can_open_hive) {
        AppendMenuW(menu, MF_STRING, cmd::kOptionsHiveFileDir, L"Open Hive File");
      }
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, expand_flags, cmd::kTreeToggleExpand, expand_label.c_str());
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kCreateSimulatedKey, L"Create Key");
    } else {
      AppendMenuW(menu, modify_flags, cmd::kNewKey, L"New Key");
      AppendMenuW(menu, MF_POPUP | (can_modify ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(new_value), L"New Value");
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (!is_simulated) {
      AppendMenuW(menu, edit_flags, cmd::kFileExport, L"Export...");
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
    if (!is_simulated) {
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
      AppendMenuW(menu, rename_flags, cmd::kEditRename, L"Rename");
      AppendMenuW(menu, delete_flags, cmd::kEditDelete, L"Delete");
    }
  } else if (row && row->kind == rowkind::kValue) {
    std::vector<ListRow> selected_rows = SelectedListRows(value_list_);
    const bool all_values = !selected_rows.empty() &&
                            std::all_of(selected_rows.begin(), selected_rows.end(), [](const ListRow& selected) {
                              return selected.kind == rowkind::kValue && !selected.simulated;
                            });
    const bool single_value = all_values && selected_rows.size() == 1;
    bool can_modify = !read_only_ && single_value;
    bool can_delete = !read_only_ && all_values;
    bool can_comment = all_values;
    bool can_export = !row->simulated && current_node_ && !current_node_->simulated;
    UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
    UINT delete_flags = MF_STRING | (can_delete ? 0 : MF_GRAYED);
    UINT single_flags = MF_STRING | (single_value ? 0 : MF_GRAYED);
    UINT export_flags = MF_STRING | (can_export ? 0 : MF_GRAYED);
    UINT comment_flags = MF_STRING | (can_comment ? 0 : MF_GRAYED);
    AppendMenuW(menu, modify_flags, cmd::kEditModify, L"Modify...");
    AppendMenuW(menu, modify_flags, cmd::kEditModifyBinary, L"Modify Binary Data...");
    AppendMenuW(menu, comment_flags, cmd::kEditModifyComment, L"Modify Comment...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, single_flags, cmd::kEditCopyValueName, L"Copy Value Name");
    AppendMenuW(menu, single_flags, cmd::kEditCopyValueData, L"Copy Value Data");
    AppendMenuW(menu, export_flags, cmd::kFileExport, L"Export...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, modify_flags, cmd::kEditRename, L"Rename");
    AppendMenuW(menu, delete_flags, cmd::kEditDelete, L"Delete");
  } else {
    bool is_simulated = current_node_ && current_node_->simulated;
    bool can_modify = !read_only_;
    UINT edit_flags = MF_STRING | (current_node_ ? 0 : MF_GRAYED);
    UINT modify_flags = MF_STRING | ((current_node_ && can_modify) ? 0 : MF_GRAYED);
    HMENU new_value = CreatePopupMenu();
    AppendMenuW(new_value, MF_STRING, cmd::kNewString, L"String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewBinary, L"Binary Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewDword, L"DWORD (32-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewQword, L"QWORD (64-bit) Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewMultiString, L"Multi-String Value");
    AppendMenuW(new_value, MF_STRING, cmd::kNewExpandString, L"Expandable String Value");

    AppendMenuW(menu, edit_flags, cmd::kEditCopyKey, L"Copy Key Name");
    AppendMenuW(menu, edit_flags, cmd::kEditCopyKeyPath, L"Copy Key Path");
    AppendMenuW(menu, MF_POPUP | (current_node_ ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(BuildCopyKeyPathMenu()), L"Copy Key Path As");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kCreateSimulatedKey, L"Create Key");
    } else {
      AppendMenuW(menu, modify_flags, cmd::kNewKey, L"New Key");
      AppendMenuW(menu, MF_POPUP | ((current_node_ && can_modify) ? 0 : MF_GRAYED), reinterpret_cast<UINT_PTR>(new_value), L"New Value");
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (!is_simulated) {
      AppendMenuW(menu, edit_flags, cmd::kFileExport, L"Export...");
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    AppendMenuW(menu, MF_STRING, cmd::kViewRefresh, L"Refresh");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    if (!is_simulated) {
      AppendMenuW(menu, modify_flags, cmd::kEditPermissions, L"Permissions...");
    }
  }

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command != 0) {
    HandleMenuCommand(command);
  }
}

void MainWindow::ShowHistoryContextMenu(POINT screen_pt) {
  if (!history_list_) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(history_list_, &client_pt);
  LVHITTESTINFO hit = {};
  hit.pt = client_pt;
  int index = ListView_HitTest(history_list_, &hit);
  if (index >= 0) {
    ListView_SetItemState(history_list_, index, LVIS_SELECTED, LVIS_SELECTED);
  }
  const HistoryEntry* entry = nullptr;
  if (index >= 0) {
    LVITEMW item = {};
    item.mask = LVIF_PARAM;
    item.iItem = index;
    if (ListView_GetItem(history_list_, &item)) {
      size_t history_index = static_cast<size_t>(item.lParam);
      if (history_index < change_history_.entries().size()) {
        entry = &change_history_.entries()[history_index];
      }
    }
  }

  HMENU menu = CreatePopupMenu();
  HistoryEntry prepared_revert;
  const bool can_revert = entry && PrepareHistoryRevert(*entry, &prepared_revert);
  UINT open_flags = MF_STRING | ((entry && !entry->key_path.empty()) ? 0 : MF_GRAYED);
  UINT revert_flags = MF_STRING | (can_revert ? 0 : MF_GRAYED);
  AppendMenuW(menu, open_flags, cmd::kHistoryOpenTarget, L"Open Entry");
  AppendMenuW(menu, revert_flags, cmd::kHistoryRevert, L"Revert");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING, cmd::kEditCopyKey, L"Copy");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING, cmd::kEditDelete, L"Clear History");

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);

  if (command == cmd::kHistoryOpenTarget && entry) {
    if (!OpenHistoryTarget(*entry)) {
      ui::ShowError(hwnd_, L"Failed to open history target.");
    }
  } else if (command == cmd::kHistoryRevert && entry) {
    RevertHistoryEntry(*entry);
  } else if (command == cmd::kEditCopyKey && index >= 0) {
    wchar_t time[128] = {};
    wchar_t action[256] = {};
    wchar_t old_data[256] = {};
    wchar_t new_data[256] = {};
    ListView_GetItemText(history_list_, index, 0, time, _countof(time));
    ListView_GetItemText(history_list_, index, 1, action, _countof(action));
    ListView_GetItemText(history_list_, index, 2, old_data, _countof(old_data));
    ListView_GetItemText(history_list_, index, 3, new_data, _countof(new_data));
    std::wstring combined = std::wstring(time) + L" | " + action + L" | " + old_data + L" | " + new_data;
    ui::CopyTextToClipboard(hwnd_, combined);
  } else if (command == cmd::kEditDelete) {
    ClearHistoryItems(true);
  }
}

void MainWindow::ShowSearchResultContextMenu(POINT screen_pt) {
  if (!search_results_list_) {
    return;
  }
  POINT client_pt = screen_pt;
  ScreenToClient(search_results_list_, &client_pt);
  LVHITTESTINFO hit = {};
  hit.pt = client_pt;
  int index = ListView_HitTest(search_results_list_, &hit);
  if (index < 0) {
    return;
  }

  int sel_tab = TabCtrl_GetCurSel(tab_);
  int search_index = SearchIndexFromTab(sel_tab);
  if (search_index < 0 || static_cast<size_t>(search_index) >= search_tabs_.size()) {
    return;
  }
  if (static_cast<size_t>(index) >= search_tabs_[static_cast<size_t>(search_index)].results.size()) {
    return;
  }

  ListView_SetItemState(search_results_list_, index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
  const search::Result& result = search_tabs_[static_cast<size_t>(search_index)].results[static_cast<size_t>(index)];
  std::wstring key_path = result.key_path;
  if (key_path.empty()) {
    return;
  }

  RegistryNode node;
  bool node_ok = ResolvePathToNode(key_path, &node);
  KeyInfo info = {};
  bool key_exists = node_ok && RegistryProvider::QueryKeyInfo(node, &info);
  bool can_modify = !read_only_;
  bool can_rename = key_exists && !node.subkey.empty() && can_modify;
  bool can_delete = key_exists && !node.subkey.empty() && can_modify;
  bool can_export = key_exists;
  bool can_permissions = key_exists && can_modify;
  bool can_open_hive = false;
  if (key_exists) {
    bool is_root = false;
    std::wstring hive_path = LookupHivePath(node, &is_root);
    if (!hive_path.empty() && is_root) {
      if (node.subkey.empty() && (node.root == HKEY_CURRENT_USER || EqualsInsensitive(node.root_name, L"HKEY_CURRENT_USER"))) {
        can_open_hive = false;
      } else {
        can_open_hive = true;
      }
    }
  }

  enum {
    kSearchOpenKey = 51000,
    kSearchOpenKeyNewTab = 51001,
    kSearchModify = 51002,
    kSearchModifyBinary = 51003,
    kSearchModifyComment = 51004,
    kSearchCopyKeyName = 51005,
    kSearchCopyKeyPath = 51006,
    kSearchCopyKeyPathAbbrev = 51013,
    kSearchCopyKeyPathRegedit = 51014,
    kSearchCopyKeyPathRegFile = 51015,
    kSearchCopyKeyPathPowerShell = 51016,
    kSearchCopyKeyPathPowerShellProvider = 51017,
    kSearchCopyKeyPathEscaped = 51018,
    kSearchPermissions = 51007,
    kSearchOpenHive = 51008,
    kSearchExport = 51009,
    kSearchRename = 51010,
    kSearchDelete = 51011,
    kSearchRefresh = 51012,
  };

  auto build_copy_path_menu = [&]() -> HMENU {
    HMENU submenu = CreatePopupMenu();
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathAbbrev, L"Abbreviated (HKLM)");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathRegedit, L"Regedit Address Bar");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathRegFile, L".reg File Header");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathPowerShell, L"PowerShell Drive");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathPowerShellProvider, L"PowerShell Provider");
    AppendMenuW(submenu, MF_STRING, kSearchCopyKeyPathEscaped, L"Escaped Backslashes");
    return submenu;
  };

  HMENU menu = CreatePopupMenu();
  AppendMenuW(menu, MF_STRING, kSearchOpenKey, L"Open Key");
  AppendMenuW(menu, MF_STRING, kSearchOpenKeyNewTab, L"Open Key in New Tab");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  if (!result.is_key) {
    UINT modify_flags = MF_STRING | (can_modify ? 0 : MF_GRAYED);
    AppendMenuW(menu, modify_flags, kSearchModify, L"Modify...");
    AppendMenuW(menu, modify_flags, kSearchModifyBinary, L"Modify Binary Data...");
    AppendMenuW(menu, MF_STRING, kSearchModifyComment, L"Modify Comment...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  }
  AppendMenuW(menu, MF_STRING, kSearchCopyKeyName, L"Copy Key Name");
  AppendMenuW(menu, MF_STRING, kSearchCopyKeyPath, L"Copy Key Path");
  AppendMenuW(menu, MF_POPUP, reinterpret_cast<UINT_PTR>(build_copy_path_menu()), L"Copy Key Path As");
  UINT permissions_flags = MF_STRING | (can_permissions ? 0 : MF_GRAYED);
  AppendMenuW(menu, permissions_flags, kSearchPermissions, L"Permissions...");
  UINT open_hive_flags = MF_STRING | (can_open_hive ? 0 : MF_GRAYED);
  AppendMenuW(menu, open_hive_flags, kSearchOpenHive, L"Open Hive File");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  UINT export_flags = MF_STRING | (can_export ? 0 : MF_GRAYED);
  AppendMenuW(menu, export_flags, kSearchExport, L"Export...");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING, kSearchRefresh, L"Refresh");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  UINT rename_flags = MF_STRING | (can_rename ? 0 : MF_GRAYED);
  UINT delete_flags = MF_STRING | (can_delete ? 0 : MF_GRAYED);
  AppendMenuW(menu, rename_flags, kSearchRename, L"Rename");
  AppendMenuW(menu, delete_flags, kSearchDelete, L"Delete");

  int command = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, screen_pt.x, screen_pt.y, 0, hwnd_, nullptr);
  DestroyMenu(menu);
  if (command == 0) {
    return;
  }

  auto open_key = [&](bool new_tab) {
    if (!tab_) {
      return;
    }
    if (new_tab) {
      OpenLocalRegistryTab();
    } else {
      int registry_tab = FindFirstRegistryTabIndex();
      if (registry_tab >= 0) {
        TabCtrl_SetCurSel(tab_, registry_tab);
      } else {
        OpenLocalRegistryTab();
      }
    }
    ApplyViewVisibility();
    UpdateStatus();
    SelectTreePath(key_path);
  };
  auto focus_key = [&]() {
    open_key(false);
    if (tree_.hwnd()) {
      SetFocus(tree_.hwnd());
    }
  };
  auto focus_value = [&]() -> bool {
    if (result.is_key) {
      return false;
    }
    open_key(false);
    if (!SelectValueByName(result.value_name)) {
      return false;
    }
    if (value_list_.hwnd()) {
      SetFocus(value_list_.hwnd());
    }
    return true;
  };

  switch (command) {
  case kSearchOpenKey:
    open_key(false);
    return;
  case kSearchOpenKeyNewTab:
    open_key(true);
    return;
  case kSearchModify:
    if (focus_value()) {
      HandleMenuCommand(cmd::kEditModify);
    }
    return;
  case kSearchModifyBinary:
    if (focus_value()) {
      HandleMenuCommand(cmd::kEditModifyBinary);
    }
    return;
  case kSearchModifyComment:
    if (focus_value()) {
      HandleMenuCommand(cmd::kEditModifyComment);
    }
    return;
  case kSearchCopyKeyName: {
    std::wstring name;
    if (node_ok) {
      name = LeafName(node);
    } else {
      size_t pos = key_path.find_last_of(L"\\/");
      name = (pos == std::wstring::npos) ? key_path : key_path.substr(pos + 1);
    }
    if (!name.empty()) {
      ui::CopyTextToClipboard(hwnd_, name);
    }
    return;
  }
  case kSearchCopyKeyPath:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, key_path);
    }
    return;
  case kSearchCopyKeyPathAbbrev:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kAbbrev));
    }
    return;
  case kSearchCopyKeyPathRegedit:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kRegedit));
    }
    return;
  case kSearchCopyKeyPathRegFile:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kRegFile));
    }
    return;
  case kSearchCopyKeyPathPowerShell:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kPowerShellDrive));
    }
    return;
  case kSearchCopyKeyPathPowerShellProvider:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kPowerShellProvider));
    }
    return;
  case kSearchCopyKeyPathEscaped:
    if (!key_path.empty()) {
      ui::CopyTextToClipboard(hwnd_, FormatRegistryPath(key_path, RegistryPathFormat::kEscaped));
    }
    return;
  case kSearchPermissions:
    if (node_ok && key_exists) {
      ShowPermissionsDialog(node);
    }
    return;
  case kSearchOpenHive:
    if (can_open_hive) {
      focus_key();
      HandleMenuCommand(cmd::kOptionsHiveFileDir);
    }
    return;
  case kSearchExport:
    if (can_export) {
      focus_key();
      HandleMenuCommand(cmd::kFileExport);
    }
    return;
  case kSearchRefresh:
    focus_key();
    HandleMenuCommand(cmd::kViewRefresh);
    return;
  case kSearchRename:
    if (can_rename) {
      focus_key();
      HandleMenuCommand(cmd::kEditRename);
    }
    return;
  case kSearchDelete:
    if (can_delete) {
      focus_key();
      HandleMenuCommand(cmd::kEditDelete);
    }
    return;
  default:
    return;
  }
}

} // namespace regkit
