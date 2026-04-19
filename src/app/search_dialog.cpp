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

#include "app/search_dialog.h"

#include <algorithm>
#include <limits>
#include <vector>

#include <commctrl.h>
#include <shlobj.h>

#include "app/registry_tree.h"
#include "app/theme.h"
#include "app/ui_helpers.h"
#include "app/value_dialogs.h"
#include "registry/registry_provider.h"
#include "win32/win32_helpers.h"

namespace regkit {

namespace {

constexpr wchar_t kDialogClass[] = L"RegKitSearchDialog";
constexpr wchar_t kAppTitle[] = L"RegKit";

enum ControlId {
  kFindLabel = 100,
  kFindCombo = 101,
  kWhereGroup = 110,
  kScopeTop = 111,
  kScopeKey = 112,
  kScopeRecursive = 113,
  kScopeCombo = 114,
  kScopeEdit = 116,
  kScopeBrowse = 117,
  kOptionsGroup = 120,
  kOptKeys = 121,
  kOptValues = 122,
  kOptData = 123,
  kOptDataTypes = 124,
  kOptMatchCase = 125,
  kOptMatchWhole = 126,
  kOptUseRegex = 127,
  kOptMinSize = 129,
  kOptMinSizeEdit = 130,
  kOptMaxSize = 131,
  kOptMaxSizeEdit = 132,
  kOptStandardHives = 133,
  kOptRegistryRoot = 134,
  kOptTraceValues = 135,
  kModifiedLabel = 140,
  kModifiedFrom = 141,
  kModifiedDash = 142,
  kModifiedTo = 143,
  kExcludeGroup = 150,
  kExcludeEnable = 151,
  kExcludeEdit = 152,
  kExcludeButton = 153,
  kResultGroup = 160,
  kResultReuse = 161,
  kResultNew = 162,
  kFindButton = IDOK,
  kCancelButton = IDCANCEL,
};

struct DataTypeItem {
  DWORD type = 0;
  std::wstring label;
};

struct BaseTypeItem {
  DWORD type = 0;
  const wchar_t* label = nullptr;
};

constexpr BaseTypeItem kBaseDataTypes[] = {
    {REG_SZ, L"REG_SZ"}, {REG_EXPAND_SZ, L"REG_EXPAND_SZ"}, {REG_MULTI_SZ, L"REG_MULTI_SZ"}, {REG_DWORD, L"DWORD (32-bit)"}, {REG_QWORD, L"QWORD (64-bit)"}, {REG_BINARY, L"REG_BINARY"}, {REG_NONE, L"REG_NONE"}, {REG_DWORD_BIG_ENDIAN, L"REG_DWORD_BIG_ENDIAN"}, {REG_LINK, L"REG_LINK"}, {REG_RESOURCE_LIST, L"REG_RESOURCE_LIST"}, {REG_FULL_RESOURCE_DESCRIPTOR, L"REG_FULL_RESOURCE_DESCRIPTOR"}, {REG_RESOURCE_REQUIREMENTS_LIST, L"REG_RESOURCE_REQUIREMENTS_LIST"},
};

constexpr DWORD kExtendedTypeFlags[] = {0x20000, 0x40000};

constexpr int kDataTypesPadding = 12;
constexpr int kDataTypesButtonHeight = 24;
constexpr int kDataTypesButtonGap = 10;
constexpr int kDataTypesColGap = 12;
constexpr int kDataTypesColCount = 3;
constexpr int kDataTypesColWidth = 270;
constexpr int kDataTypesRowHeight = 20;
constexpr int kDataTypesRowStep = 24;

std::vector<DataTypeItem> BuildDataTypeItems() {
  std::vector<DataTypeItem> items;
  items.reserve(_countof(kBaseDataTypes) * (1 + _countof(kExtendedTypeFlags)));
  for (const auto& entry : kBaseDataTypes) {
    items.push_back({entry.type, entry.label});
  }
  for (DWORD flag : kExtendedTypeFlags) {
    for (const auto& entry : kBaseDataTypes) {
      DWORD type = flag | entry.type;
      items.push_back({type, RegistryProvider::FormatValueType(type)});
    }
  }
  return items;
}

struct SearchDialogState {
  HWND hwnd = nullptr;
  HWND find_combo = nullptr;
  HWND scope_top = nullptr;
  HWND scope_key = nullptr;
  HWND scope_recursive = nullptr;
  HWND scope_combo = nullptr;
  HWND scope_edit = nullptr;
  HWND scope_browse = nullptr;
  HWND options_keys = nullptr;
  HWND options_values = nullptr;
  HWND options_data = nullptr;
  HWND options_data_types = nullptr;
  HWND options_standard = nullptr;
  HWND options_registry = nullptr;
  HWND options_trace = nullptr;
  HWND match_case = nullptr;
  HWND match_whole = nullptr;
  HWND use_regex = nullptr;
  HWND min_size = nullptr;
  HWND min_size_edit = nullptr;
  HWND max_size = nullptr;
  HWND max_size_edit = nullptr;
  HWND modified_from = nullptr;
  HWND modified_to = nullptr;
  HWND exclude_enable = nullptr;
  HWND exclude_edit = nullptr;
  HWND exclude_button = nullptr;
  HWND result_reuse = nullptr;
  HWND result_new = nullptr;
  HWND find_button = nullptr;
  HWND cancel_button = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  SearchDialogResult* out = nullptr;
  bool trace_available = false;
  bool registry_available = true;
  bool accepted = false;
  bool recursive = true;
  std::vector<std::wstring> history;
  std::vector<std::wstring> root_names;
  std::vector<bool> root_selected;
  std::vector<DWORD> data_types;
  bool owner_restored = false;
};

HFONT CreateDialogFont() {
  return ui::DefaultUIFont();
}

void ApplyFont(HWND hwnd, HFONT font) {
  if (hwnd && font) {
    SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
  }
}

int CalcDialogLineHeight(HWND hwnd, HFONT font, int min_height) {
  if (!hwnd || !font) {
    return min_height;
  }
  int height = min_height;
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

std::wstring SearchHistoryPath() {
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return L"";
  }
  return util::JoinPath(folder, L"search_history.txt");
}

std::vector<std::wstring> LoadSearchHistory() {
  std::vector<std::wstring> items;
  std::wstring path = SearchHistoryPath();
  if (path.empty()) {
    return items;
  }
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return items;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return items;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return items;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(buffer);
  if (content.empty()) {
    return items;
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
    if (!line.empty()) {
      items.push_back(line);
    }
    start = end + 1;
  }
  return items;
}

void SaveSearchHistory(const std::vector<std::wstring>& items) {
  std::wstring path = SearchHistoryPath();
  if (path.empty()) {
    return;
  }
  std::wstring content;
  for (const auto& item : items) {
    content += item;
    content += L"\n";
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

void UpdateHistoryList(std::vector<std::wstring>* items, const std::wstring& entry) {
  if (!items || entry.empty()) {
    return;
  }
  items->erase(std::remove_if(items->begin(), items->end(), [&](const std::wstring& item) { return _wcsicmp(item.c_str(), entry.c_str()) == 0; }), items->end());
  items->insert(items->begin(), entry);
  const size_t max_items = 20;
  if (items->size() > max_items) {
    items->resize(max_items);
  }
}

void PopulateHistoryCombo(HWND combo, const std::vector<std::wstring>& items) {
  if (!combo) {
    return;
  }
  SendMessageW(combo, CB_RESETCONTENT, 0, 0);
  for (const auto& item : items) {
    SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item.c_str()));
  }
}

void UpdateScopeComboText(SearchDialogState* state) {
  if (!state || !state->scope_combo) {
    return;
  }
  size_t total = state->root_selected.size();
  size_t selected = 0;
  std::wstring first;
  for (size_t i = 0; i < state->root_selected.size(); ++i) {
    if (state->root_selected[i]) {
      ++selected;
      if (first.empty() && i < state->root_names.size()) {
        first = state->root_names[i];
      }
    }
  }
  std::wstring text;
  if (selected == 0) {
    text = L"No top-level keys";
  } else if (selected == total) {
    text = L"All top-level keys";
  } else if (selected == 1) {
    text = first;
  } else {
    text = L"Multiple keys";
  }
  SendMessageW(state->scope_combo, CB_RESETCONTENT, 0, 0);
  SendMessageW(state->scope_combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text.c_str()));
  SendMessageW(state->scope_combo, CB_SETCURSEL, 0, 0);
}

void ShowRootSelectionMenu(HWND owner, SearchDialogState* state) {
  if (!owner || !state || !state->scope_combo) {
    return;
  }
  RECT rect = {};
  GetWindowRect(state->scope_combo, &rect);
  HMENU menu = CreatePopupMenu();
  for (size_t i = 0; i < state->root_names.size(); ++i) {
    UINT flags = MF_STRING;
    if (i < state->root_selected.size() && state->root_selected[i]) {
      flags |= MF_CHECKED;
    }
    AppendMenuW(menu, flags, static_cast<UINT>(1000 + i), state->root_names[i].c_str());
  }
  int cmd = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_LEFTALIGN | TPM_TOPALIGN, rect.left, rect.bottom, 0, owner, nullptr);
  DestroyMenu(menu);
  if (cmd >= 1000) {
    size_t index = static_cast<size_t>(cmd - 1000);
    if (index < state->root_selected.size()) {
      state->root_selected[index] = !state->root_selected[index];
      UpdateScopeComboText(state);
    }
  }
}

void CenterWindowToOwner(HWND hwnd, HWND owner, int width, int height) {
  RECT rect = {};
  if (owner && GetWindowRect(owner, &rect)) {
    int owner_w = rect.right - rect.left;
    int owner_h = rect.bottom - rect.top;
    int x = rect.left + std::max(0, (owner_w - width) / 2);
    int y = rect.top + std::max(0, (owner_h - height) / 2);
    SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    return;
  }
  SetWindowPos(hwnd, nullptr, CW_USEDEFAULT, CW_USEDEFAULT, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
}

void FocusFindCombo(SearchDialogState* state) {
  if (!state || !state->find_combo) {
    return;
  }
  COMBOBOXINFO info = {};
  info.cbSize = sizeof(info);
  if (GetComboBoxInfo(state->find_combo, &info) && info.hwndItem) {
    SetFocus(info.hwndItem);
    SendMessageW(info.hwndItem, EM_SETSEL, 0, -1);
    return;
  }
  SetFocus(state->find_combo);
  SendMessageW(state->find_combo, CB_SETEDITSEL, 0, MAKELPARAM(0, -1));
}

void RestoreOwnerWindow(HWND owner, bool* restored) {
  if (!owner || !restored || *restored) {
    return;
  }
  EnableWindow(owner, TRUE);
  SetActiveWindow(owner);
  SetForegroundWindow(owner);
  *restored = true;
}

bool ParseUint64(const std::wstring& text, uint64_t* out) {
  if (!out) {
    return false;
  }
  *out = 0;
  if (text.empty()) {
    return false;
  }
  wchar_t* end = nullptr;
  unsigned long long value = wcstoull(text.c_str(), &end, 10);
  if (!end || end == text.c_str() || *end != L'\0') {
    return false;
  }
  *out = static_cast<uint64_t>(value);
  return true;
}

bool GetDateTimeValue(HWND control, FILETIME* out) {
  if (!control || !out) {
    return false;
  }
  SYSTEMTIME local = {};
  DWORD result = static_cast<DWORD>(SendMessageW(control, DTM_GETSYSTEMTIME, 0, reinterpret_cast<LPARAM>(&local)));
  if (result != GDT_VALID) {
    return false;
  }
  SYSTEMTIME utc = {};
  if (!TzSpecificLocalTimeToSystemTime(nullptr, &local, &utc)) {
    return false;
  }
  return SystemTimeToFileTime(&utc, out) != 0;
}

void SetDateTimeValue(HWND control, const FILETIME& value) {
  if (!control) {
    return;
  }
  SYSTEMTIME utc = {};
  SYSTEMTIME local = {};
  if (!FileTimeToSystemTime(&value, &utc)) {
    return;
  }
  if (!SystemTimeToTzSpecificLocalTime(nullptr, &utc, &local)) {
    return;
  }
  SendMessageW(control, DTM_SETSYSTEMTIME, GDT_VALID, reinterpret_cast<LPARAM>(&local));
}

struct DataTypesDialogState {
  HWND hwnd = nullptr;
  HWND ok_button = nullptr;
  HWND cancel_button = nullptr;
  HWND select_all = nullptr;
  HWND clear_all = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  std::vector<DataTypeItem> items;
  std::vector<HWND> checks;
  std::vector<DWORD> types;
  int padding = kDataTypesPadding;
  int row_h = kDataTypesRowHeight;
  int row_step = kDataTypesRowStep;
  int col_w = kDataTypesColWidth;
  int col_gap = kDataTypesColGap;
  int rows_per_col = 12;
  int button_h = kDataTypesButtonHeight;
  int button_gap = kDataTypesButtonGap;
  bool accepted = false;
  bool owner_restored = false;
};

LRESULT CALLBACK DataTypesDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<DataTypesDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<DataTypesDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    state->font = CreateDialogFont();
    HFONT font = state->font;
    RECT client = {};
    GetClientRect(hwnd, &client);
    int btn_h = state->button_h;
    int x = state->padding;
    int y = state->padding;
    int col_w = state->col_w;
    int row_h = state->row_h;
    int row_step = state->row_step;
    int col_gap = state->col_gap;
    int rows_per_col = state->rows_per_col;
    int col = 0;
    int row = 0;
    for (const auto& item : state->items) {
      HWND check = CreateWindowExW(0, L"BUTTON", item.label.c_str(), WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, x + col * (col_w + col_gap), y + row * row_step, col_w, row_h, hwnd, nullptr, nullptr, nullptr);
      ApplyFont(check, font);
      bool checked = state->types.empty();
      if (!state->types.empty()) {
        checked = std::find(state->types.begin(), state->types.end(), item.type) != state->types.end();
      }
      SendMessageW(check, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
      state->checks.push_back(check);
      ++row;
      if (row >= rows_per_col) {
        row = 0;
        ++col;
      }
    }
    const int btn_w = 70;
    const int btn_gap = 12;
    const int aux_btn_w = 90;
    int btn_y = client.bottom - state->padding - btn_h;
    int cancel_x = client.right - state->padding - btn_w;
    int ok_x = cancel_x - btn_gap - btn_w;
    int select_x = x;
    int clear_x = select_x + aux_btn_w + btn_gap;
    state->select_all = CreateWindowExW(0, L"BUTTON", L"Select All", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, select_x, btn_y, aux_btn_w, btn_h, hwnd, reinterpret_cast<HMENU>(100), nullptr, nullptr);
    state->clear_all = CreateWindowExW(0, L"BUTTON", L"Clear All", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, clear_x, btn_y, aux_btn_w, btn_h, hwnd, reinterpret_cast<HMENU>(101), nullptr, nullptr);
    state->ok_button = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, ok_x, btn_y, btn_w, btn_h, hwnd, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);
    state->cancel_button = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, cancel_x, btn_y, btn_w, btn_h, hwnd, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);
    ApplyFont(state->select_all, font);
    ApplyFont(state->clear_all, font);
    ApplyFont(state->ok_button, font);
    ApplyFont(state->cancel_button, font);
    Theme::Current().ApplyToChildren(hwnd);
    return 0;
  }
  case WM_DESTROY:
    if (state && state->font) {
      DeleteObject(state->font);
      state->font = nullptr;
    }
    return 0;
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(hwnd);
      Theme::Current().ApplyToChildren(hwnd);
      InvalidateRect(hwnd, nullptr, TRUE);
    }
    return 0;
  }
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLORBTN:
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    int type = (msg == WM_CTLCOLORSTATIC) ? CTLCOLOR_STATIC : (msg == WM_CTLCOLORBTN) ? CTLCOLOR_BTN : CTLCOLOR_EDIT;
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return 1;
  }
  case WM_COMMAND: {
    if (!state) {
      return 0;
    }
    switch (LOWORD(wparam)) {
    case 100: { // select all
      for (HWND check : state->checks) {
        SendMessageW(check, BM_SETCHECK, BST_CHECKED, 0);
      }
      return 0;
    }
    case 101: { // clear all
      for (HWND check : state->checks) {
        SendMessageW(check, BM_SETCHECK, BST_UNCHECKED, 0);
      }
      return 0;
    }
    case IDOK: {
      state->types.clear();
      for (size_t i = 0; i < state->checks.size() && i < state->items.size(); ++i) {
        if (SendMessageW(state->checks[i], BM_GETCHECK, 0, 0) == BST_CHECKED) {
          state->types.push_back(state->items[i].type);
        }
      }
      if (state->types.empty()) {
        ui::ShowWarning(hwnd, L"Select at least one data type.");
        return 0;
      }
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    }
    case IDCANCEL:
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  }
  case WM_CLOSE:
    if (state) {
      RestoreOwnerWindow(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

bool ShowDataTypesDialog(HWND owner, std::vector<DWORD>* types) {
  if (!types) {
    return false;
  }
  HINSTANCE instance = GetModuleHandleW(nullptr);
  WNDCLASSW wc = {};
  wc.lpfnWndProc = DataTypesDialogProc;
  wc.hInstance = instance;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = L"RegKitDataTypesDialog";
  RegisterClassW(&wc);

  DataTypesDialogState state;
  state.items = BuildDataTypeItems();
  state.types = *types;
  state.owner = owner;
  int count = static_cast<int>(state.items.size());
  int rows_per_col = std::max(1, (count + kDataTypesColCount - 1) / kDataTypesColCount);
  int rows = rows_per_col;
  int content_w = kDataTypesColCount * kDataTypesColWidth + (kDataTypesColCount - 1) * kDataTypesColGap;
  int content_h = rows * kDataTypesRowStep;
  int client_w = kDataTypesPadding * 2 + content_w;
  int client_h = kDataTypesPadding * 2 + content_h + kDataTypesButtonGap + kDataTypesButtonHeight;

  RECT window_rect = {0, 0, client_w, client_h};
  DWORD style = WS_POPUP | WS_CAPTION | WS_SYSMENU;
  DWORD ex_style = WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT;
  AdjustWindowRectEx(&window_rect, style, FALSE, ex_style);
  int width = window_rect.right - window_rect.left;
  int height = window_rect.bottom - window_rect.top;

  state.row_h = kDataTypesRowHeight;
  state.row_step = kDataTypesRowStep;
  state.col_w = kDataTypesColWidth;
  state.rows_per_col = rows_per_col;

  HWND hwnd = CreateWindowExW(ex_style, wc.lpszClassName, kAppTitle, style, CW_USEDEFAULT, CW_USEDEFAULT, width, height, owner, nullptr, instance, &state);
  if (!hwnd) {
    return false;
  }
  Theme::Current().ApplyToWindow(hwnd);
  CenterWindowToOwner(hwnd, owner, width, height);

  EnableWindow(owner, FALSE);
  ShowWindow(hwnd, SW_SHOW);
  UpdateWindow(hwnd);

  MSG msg = {};
  while (IsWindow(hwnd) && GetMessageW(&msg, nullptr, 0, 0)) {
    if (!IsDialogMessageW(hwnd, &msg)) {
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  }
  RestoreOwnerWindow(owner, &state.owner_restored);

  if (state.accepted) {
    *types = state.types;
    return true;
  }
  return false;
}

struct BrowseDialogState {
  HWND hwnd = nullptr;
  HWND tree_hwnd = nullptr;
  HWND ok_button = nullptr;
  HWND cancel_button = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  RegistryTree tree;
  std::wstring selected_path;
  bool accepted = false;
  bool owner_restored = false;
};

LRESULT CALLBACK BrowseDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<BrowseDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<BrowseDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    state->font = CreateDialogFont();
    HFONT font = state->font;
    state->tree_hwnd = CreateWindowExW(0, WC_TREEVIEWW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(1), nullptr, nullptr);
    state->ok_button = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);
    state->cancel_button = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);
    ApplyFont(state->tree_hwnd, font);
    ApplyFont(state->ok_button, font);
    ApplyFont(state->cancel_button, font);

    state->tree.Create(hwnd, GetModuleHandleW(nullptr), 1);
    std::vector<RegistryRootEntry> roots = RegistryProvider::DefaultRoots();
    state->tree.PopulateRoots(roots);

    Theme::Current().ApplyToTreeView(state->tree.hwnd());
    Theme::Current().ApplyToChildren(hwnd);
    return 0;
  }
  case WM_DESTROY:
    if (state && state->font) {
      DeleteObject(state->font);
      state->font = nullptr;
    }
    return 0;
  case WM_SIZE: {
    RECT client = {};
    GetClientRect(hwnd, &client);
    int width = client.right - client.left;
    int height = client.bottom - client.top;
    int padding = 10;
    int btn_w = 80;
    int btn_h = 24;
    SetWindowPos(state->tree.hwnd(), nullptr, padding, padding, width - padding * 2, height - padding * 2 - btn_h, SWP_NOZORDER);
    int bottom_y = height - padding - btn_h;
    SetWindowPos(state->ok_button, nullptr, width - padding - btn_w * 2 - 8, bottom_y, btn_w, btn_h, SWP_NOZORDER);
    SetWindowPos(state->cancel_button, nullptr, width - padding - btn_w, bottom_y, btn_w, btn_h, SWP_NOZORDER);
    return 0;
  }
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(hwnd);
      Theme::Current().ApplyToChildren(hwnd);
      InvalidateRect(hwnd, nullptr, TRUE);
    }
    return 0;
  }
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLORBTN:
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    int type = (msg == WM_CTLCOLORSTATIC) ? CTLCOLOR_STATIC : (msg == WM_CTLCOLORBTN) ? CTLCOLOR_BTN : CTLCOLOR_EDIT;
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return 1;
  }
  case WM_NOTIFY: {
    auto* hdr = reinterpret_cast<NMHDR*>(lparam);
    if (hdr && hdr->hwndFrom == state->tree.hwnd()) {
      if (hdr->code == TVN_ITEMEXPANDINGW) {
        state->tree.OnItemExpanding(reinterpret_cast<NMTREEVIEWW*>(lparam));
        return 0;
      }
      if (hdr->code == TVN_SELCHANGEDW) {
        RegistryNode* node = state->tree.OnSelectionChanged(reinterpret_cast<NMTREEVIEWW*>(lparam));
        if (node) {
          state->selected_path = RegistryProvider::BuildPath(*node);
        }
        return 0;
      }
      if (hdr->code == NM_CUSTOMDRAW) {
        const Theme& theme = Theme::Current();
        auto* draw = reinterpret_cast<NMTVCUSTOMDRAW*>(lparam);
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
    }
    break;
  }
  case WM_COMMAND: {
    if (!state) {
      return 0;
    }
    switch (LOWORD(wparam)) {
    case IDOK: {
      if (state->selected_path.empty()) {
        ui::ShowWarning(hwnd, L"Select a key.");
        return 0;
      }
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    }
    case IDCANCEL:
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  }
  case WM_CLOSE:
    if (state) {
      RestoreOwnerWindow(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

bool ShowBrowseKeyDialogInternal(HWND owner, std::wstring* selected_path) {
  if (!selected_path) {
    return false;
  }
  HINSTANCE instance = GetModuleHandleW(nullptr);
  WNDCLASSW wc = {};
  wc.lpfnWndProc = BrowseDialogProc;
  wc.hInstance = instance;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = L"RegKitBrowseKeyDialog";
  RegisterClassW(&wc);

  BrowseDialogState state;
  state.owner = owner;
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, wc.lpszClassName, kAppTitle, WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_SIZEBOX, CW_USEDEFAULT, CW_USEDEFAULT, 420, 420, owner, nullptr, instance, &state);
  if (!hwnd) {
    return false;
  }
  Theme::Current().ApplyToWindow(hwnd);
  CenterWindowToOwner(hwnd, owner, 420, 420);

  EnableWindow(owner, FALSE);
  ShowWindow(hwnd, SW_SHOW);
  UpdateWindow(hwnd);

  MSG msg = {};
  while (IsWindow(hwnd) && GetMessageW(&msg, nullptr, 0, 0)) {
    if (!IsDialogMessageW(hwnd, &msg)) {
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  }
  RestoreOwnerWindow(owner, &state.owner_restored);

  if (state.accepted) {
    *selected_path = state.selected_path;
    return true;
  }
  return false;
}

std::vector<std::wstring> SplitExcludePaths(const std::wstring& text) {
  std::vector<std::wstring> items;
  std::wstring current;
  for (wchar_t ch : text) {
    if (ch == L'\r' || ch == L'\n' || ch == L';' || ch == L',') {
      if (!current.empty()) {
        items.push_back(current);
        current.clear();
      }
    } else {
      current.push_back(ch);
    }
  }
  if (!current.empty()) {
    items.push_back(current);
  }
  for (auto& item : items) {
    item.erase(item.begin(), std::find_if(item.begin(), item.end(), [](wchar_t c) { return c != L' '; }));
    while (!item.empty() && item.back() == L' ') {
      item.pop_back();
    }
  }
  items.erase(std::remove_if(items.begin(), items.end(), [](const std::wstring& value) { return value.empty(); }), items.end());
  return items;
}

std::wstring JoinExcludePaths(const std::vector<std::wstring>& items) {
  std::wstring out;
  for (const auto& item : items) {
    if (item.empty()) {
      continue;
    }
    if (!out.empty()) {
      out.append(L", ");
    }
    out.append(item);
  }
  return out;
}

bool IsChecked(HWND hwnd) {
  return SendMessageW(hwnd, BM_GETCHECK, 0, 0) == BST_CHECKED;
}

void UpdateDialogEnableState(SearchDialogState* state) {
  if (!state) {
    return;
  }

  bool scope_top = state->scope_top && IsChecked(state->scope_top);
  bool scope_key = state->scope_key && IsChecked(state->scope_key);
  bool standard_roots = state->options_standard && IsChecked(state->options_standard);
  bool enable_roots = scope_top;
  EnableWindow(state->scope_combo, enable_roots && standard_roots);
  EnableWindow(state->scope_edit, scope_key);
  EnableWindow(state->scope_browse, scope_key);
  EnableWindow(state->scope_recursive, scope_key);

  bool search_data = state->options_data && IsChecked(state->options_data);
  EnableWindow(state->options_data_types, search_data);
  EnableWindow(state->min_size, search_data);
  EnableWindow(state->max_size, search_data);

  bool min_checked = state->min_size && IsChecked(state->min_size);
  bool max_checked = state->max_size && IsChecked(state->max_size);
  EnableWindow(state->min_size_edit, search_data && min_checked);
  EnableWindow(state->max_size_edit, search_data && max_checked);

  bool exclude_checked = state->exclude_enable && IsChecked(state->exclude_enable);
  EnableWindow(state->exclude_edit, exclude_checked);
  EnableWindow(state->exclude_button, exclude_checked);

  if (state->options_trace) {
    EnableWindow(state->options_trace, state->trace_available);
  }
  if (state->options_registry) {
    EnableWindow(state->options_registry, state->registry_available);
  }
}

void LayoutDialog(HWND hwnd, SearchDialogState* state, HFONT font) {
  if (!hwnd || !state) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd, &client);
  int width = client.right - client.left;
  int x = 12;
  int y = 12;
  int label_w = 94;
  int edit_h = CalcDialogLineHeight(hwnd, font, 20);
  int line_h = std::max(edit_h + 2, 22);
  auto place_check = [&](HWND check, int x_pos, int y_pos, int width) {
    if (check) {
      SetWindowPos(check, nullptr, x_pos, y_pos, width, 18, SWP_NOZORDER);
    }
  };

  HWND find_label = GetDlgItem(hwnd, kFindLabel);
  SetWindowPos(find_label, nullptr, x, y + 4, label_w, 18, SWP_NOZORDER);
  int combo_w = width - x * 2 - label_w - 8;
  SetWindowPos(state->find_combo, nullptr, x + label_w + 8, y, combo_w, line_h, SWP_NOZORDER);
  y += line_h + 12;

  int group_w = width - x * 2;
  int where_h = 100;
  SetWindowPos(GetDlgItem(hwnd, kWhereGroup), nullptr, x, y, group_w, where_h, SWP_NOZORDER);
  int gx = x + 12;
  int gy = y + 20;
  int left_w = 150;
  place_check(state->scope_top, gx, gy, left_w);
  int key_offset = 5;
  place_check(state->scope_key, gx, gy + 22 + key_offset, left_w);
  int combo_x = gx + left_w + 8;
  int combo_w2 = width - combo_x - x - 8;
  SetWindowPos(state->scope_combo, nullptr, combo_x, gy, combo_w2, line_h, SWP_NOZORDER);
  int scope_edit_y = gy + 22 + key_offset + (line_h - edit_h) / 2;
  SetWindowPos(state->scope_edit, nullptr, combo_x, scope_edit_y, combo_w2 - 84, edit_h, SWP_NOZORDER);
  SetWindowPos(state->scope_browse, nullptr, combo_x + combo_w2 - 80, scope_edit_y, 80, edit_h, SWP_NOZORDER);
  place_check(state->scope_recursive, combo_x, gy + 44 + key_offset, 140);
  y += where_h + 12;

  int options_h = 160;
  SetWindowPos(GetDlgItem(hwnd, kOptionsGroup), nullptr, x, y, group_w, options_h, SWP_NOZORDER);
  gy = y + 20;
  int left_x = x + 12;
  int right_x = x + group_w / 2 + 8;
  place_check(state->options_keys, left_x, gy, 170);
  place_check(state->options_values, left_x, gy + 22, 170);
  place_check(state->options_data, left_x, gy + 44, 170);
  place_check(state->options_standard, left_x, gy + 66, 200);
  place_check(state->options_registry, left_x, gy + 88, 200);
  place_check(state->options_trace, left_x, gy + 110, 200);

  place_check(state->min_size, right_x, gy, 180);
  SetWindowPos(state->min_size_edit, nullptr, right_x + 188, gy - 4, 76, line_h, SWP_NOZORDER);
  place_check(state->max_size, right_x, gy + 22, 180);
  SetWindowPos(state->max_size_edit, nullptr, right_x + 188, gy + 18, 76, line_h, SWP_NOZORDER);
  place_check(state->match_case, right_x, gy + 44, 140);
  place_check(state->match_whole, right_x, gy + 66, 160);
  place_check(state->use_regex, right_x, gy + 88, 190);
  SetWindowPos(state->options_data_types, nullptr, right_x, gy + 110, 120, 20, SWP_NOZORDER);
  y += options_h + 8;

  int modified_label_w = 150;
  int modified_x = x + modified_label_w + 6;
  int modified_w = 150;
  SetWindowPos(GetDlgItem(hwnd, kModifiedLabel), nullptr, x, y + 4, modified_label_w, 18, SWP_NOZORDER);
  SetWindowPos(state->modified_from, nullptr, modified_x, y, modified_w, line_h, SWP_NOZORDER);
  SetWindowPos(GetDlgItem(hwnd, kModifiedDash), nullptr, modified_x + modified_w + 6, y + 4, 12, 18, SWP_NOZORDER);
  SetWindowPos(state->modified_to, nullptr, modified_x + modified_w + 24, y, modified_w, line_h, SWP_NOZORDER);
  y += line_h + 12;

  int exclude_h = 70;
  SetWindowPos(GetDlgItem(hwnd, kExcludeGroup), nullptr, x, y, group_w, exclude_h, SWP_NOZORDER);
  place_check(state->exclude_enable, x + 12, y + 20, 120);
  SetWindowPos(state->exclude_edit, nullptr, x + 12, y + 40, group_w - 100, line_h, SWP_NOZORDER);
  SetWindowPos(state->exclude_button, nullptr, x + group_w - 80, y + 40, 70, line_h, SWP_NOZORDER);
  y += exclude_h + 8;

  int result_h = 70;
  SetWindowPos(GetDlgItem(hwnd, kResultGroup), nullptr, x, y, group_w, result_h, SWP_NOZORDER);
  place_check(state->result_reuse, x + 12, y + 20, 220);
  place_check(state->result_new, x + 12, y + 42, 240);

  int btn_y = client.bottom - 30;
  SetWindowPos(state->find_button, nullptr, width - 180, btn_y, 80, 22, SWP_NOZORDER);
  SetWindowPos(state->cancel_button, nullptr, width - 90, btn_y, 80, 22, SWP_NOZORDER);

  ApplyFont(hwnd, font);
  ApplyFont(find_label, font);
  ApplyFont(state->find_combo, font);
  ApplyFont(state->scope_combo, font);
}

LRESULT CALLBACK SearchDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<SearchDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<SearchDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, L"Find");
    state->font = CreateDialogFont();
    HFONT font = state->font;

    CreateWindowExW(0, L"STATIC", L"Find what:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kFindLabel), nullptr, nullptr);
    state->find_combo = CreateWindowExW(0, WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | CBS_AUTOHSCROLL, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kFindCombo), nullptr, nullptr);

    CreateWindowExW(0, L"BUTTON", L"Where to search", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kWhereGroup), nullptr, nullptr);
    state->scope_top = CreateWindowExW(0, L"BUTTON", L"Top-level keys", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON | WS_GROUP, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kScopeTop), nullptr, nullptr);
    state->scope_key = CreateWindowExW(0, L"BUTTON", L"Specific Key", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kScopeKey), nullptr, nullptr);
    state->scope_combo = CreateWindowExW(0, WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_HASSTRINGS, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kScopeCombo), nullptr, nullptr);
    state->scope_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kScopeEdit), nullptr, nullptr);
    state->scope_browse = CreateWindowExW(0, L"BUTTON", L"Browse...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kScopeBrowse), nullptr, nullptr);
    state->scope_recursive = CreateWindowExW(0, L"BUTTON", L"Recursive", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kScopeRecursive), nullptr, nullptr);

    CreateWindowExW(0, L"BUTTON", L"Search options", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptionsGroup), nullptr, nullptr);
    state->options_keys = CreateWindowExW(0, L"BUTTON", L"Search keys", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptKeys), nullptr, nullptr);
    state->options_values = CreateWindowExW(0, L"BUTTON", L"Search values", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptValues), nullptr, nullptr);
    state->options_data = CreateWindowExW(0, L"BUTTON", L"Search data", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptData), nullptr, nullptr);
    state->options_data_types = CreateWindowExW(0, L"BUTTON", L"Data Types...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptDataTypes), nullptr, nullptr);
    state->match_case = CreateWindowExW(0, L"BUTTON", L"Match case", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptMatchCase), nullptr, nullptr);
    state->match_whole = CreateWindowExW(0, L"BUTTON", L"Match whole string", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptMatchWhole), nullptr, nullptr);
    state->use_regex = CreateWindowExW(0, L"BUTTON", L"Use regular expressions", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptUseRegex), nullptr, nullptr);
    state->min_size = CreateWindowExW(0, L"BUTTON", L"Min data size (bytes):", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptMinSize), nullptr, nullptr);
    state->min_size_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptMinSizeEdit), nullptr, nullptr);
    state->max_size = CreateWindowExW(0, L"BUTTON", L"Max data size (bytes):", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptMaxSize), nullptr, nullptr);
    state->max_size_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptMaxSizeEdit), nullptr, nullptr);
    state->options_standard = CreateWindowExW(0, L"BUTTON", L"Search Standard Hives", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptStandardHives), nullptr, nullptr);
    state->options_registry = CreateWindowExW(0, L"BUTTON", L"Search REGISTRY", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptRegistryRoot), nullptr, nullptr);
    state->options_trace = CreateWindowExW(0, L"BUTTON", L"Search Trace Values", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptTraceValues), nullptr, nullptr);

    CreateWindowExW(0, L"STATIC", L"Modified in period:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kModifiedLabel), nullptr, nullptr);
    CreateWindowExW(0, L"STATIC", L"-", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kModifiedDash), nullptr, nullptr);
    state->modified_from = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"", WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | DTS_SHOWNONE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kModifiedFrom), nullptr, nullptr);
    state->modified_to = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"", WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | DTS_SHOWNONE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kModifiedTo), nullptr, nullptr);
    SendMessageW(state->modified_from, DTM_SETFORMAT, 0, reinterpret_cast<LPARAM>(L"M/d/yyyy HH:mm"));
    SendMessageW(state->modified_to, DTM_SETFORMAT, 0, reinterpret_cast<LPARAM>(L"M/d/yyyy HH:mm"));
    SendMessageW(state->modified_from, DTM_SETSYSTEMTIME, GDT_NONE, 0);
    SendMessageW(state->modified_to, DTM_SETSYSTEMTIME, GDT_NONE, 0);

    CreateWindowExW(0, L"BUTTON", L"Exclude keys", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kExcludeGroup), nullptr, nullptr);
    state->exclude_enable = CreateWindowExW(0, L"BUTTON", L"Exclude keys", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kExcludeEnable), nullptr, nullptr);
    state->exclude_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kExcludeEdit), nullptr, nullptr);
    state->exclude_button = CreateWindowExW(0, L"BUTTON", L"Edit...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kExcludeButton), nullptr, nullptr);

    CreateWindowExW(0, L"BUTTON", L"Result options", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kResultGroup), nullptr, nullptr);
    state->result_reuse = CreateWindowExW(0, L"BUTTON", L"Reuse last Find Results window", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON | WS_GROUP, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kResultReuse), nullptr, nullptr);
    state->result_new = CreateWindowExW(0, L"BUTTON", L"Open new Find Results window", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kResultNew), nullptr, nullptr);

    state->find_button = CreateWindowExW(0, L"BUTTON", L"Find", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kFindButton), nullptr, nullptr);
    state->cancel_button = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kCancelButton), nullptr, nullptr);

    ApplyFont(hwnd, font);
    EnumChildWindows(
        hwnd,
        [](HWND child, LPARAM param) -> BOOL {
          HFONT font_handle = reinterpret_cast<HFONT>(param);
          ApplyFont(child, font_handle);
          return TRUE;
        },
        reinterpret_cast<LPARAM>(font));

    state->history = LoadSearchHistory();
    PopulateHistoryCombo(state->find_combo, state->history);
    if (state->out && !state->out->criteria.query.empty()) {
      SetWindowTextW(state->find_combo, state->out->criteria.query.c_str());
    } else if (!state->history.empty()) {
      SetWindowTextW(state->find_combo, state->history.front().c_str());
    }

    auto roots = RegistryProvider::DefaultRoots();
    state->root_names.clear();
    state->root_selected.clear();
    state->root_names.reserve(roots.size());
    for (const auto& root : roots) {
      state->root_names.push_back(root.path_name);
    }
    state->root_selected.assign(state->root_names.size(), true);
    if (state->out && !state->out->root_paths.empty()) {
      state->root_selected.assign(state->root_names.size(), false);
      for (size_t i = 0; i < state->root_names.size(); ++i) {
        for (const auto& path : state->out->root_paths) {
          if (_wcsicmp(state->root_names[i].c_str(), path.c_str()) == 0) {
            state->root_selected[i] = true;
            break;
          }
        }
      }
    }
    UpdateScopeComboText(state);

    if (state->out) {
      state->data_types = state->out->criteria.allowed_types;
      state->recursive = state->out->criteria.recursive;
    }
    const SearchDialogResult* initial = state->out;
    if (initial) {
      SendMessageW(state->options_keys, BM_SETCHECK, initial->criteria.search_keys ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->options_values, BM_SETCHECK, initial->criteria.search_values ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->options_data, BM_SETCHECK, initial->criteria.search_data ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->match_case, BM_SETCHECK, initial->criteria.match_case ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->match_whole, BM_SETCHECK, initial->criteria.match_whole ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->use_regex, BM_SETCHECK, initial->criteria.use_regex ? BST_CHECKED : BST_UNCHECKED, 0);
      if (initial->criteria.use_min_size) {
        SendMessageW(state->min_size, BM_SETCHECK, BST_CHECKED, 0);
        SetWindowTextW(state->min_size_edit, std::to_wstring(initial->criteria.min_size).c_str());
      }
      if (initial->criteria.use_max_size) {
        SendMessageW(state->max_size, BM_SETCHECK, BST_CHECKED, 0);
        SetWindowTextW(state->max_size_edit, std::to_wstring(initial->criteria.max_size).c_str());
      }
      if (initial->criteria.use_modified_from) {
        SetDateTimeValue(state->modified_from, initial->criteria.modified_from);
      }
      if (initial->criteria.use_modified_to) {
        SetDateTimeValue(state->modified_to, initial->criteria.modified_to);
      }
      bool standard_hives = initial->search_standard_hives;
      bool registry_root = initial->search_registry_root;
      bool trace_values = initial->search_trace_values;
      if (!state->registry_available) {
        registry_root = false;
      }
      if (!state->trace_available) {
        trace_values = false;
      }
      SendMessageW(state->options_standard, BM_SETCHECK, standard_hives ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->options_registry, BM_SETCHECK, registry_root ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->options_trace, BM_SETCHECK, trace_values ? BST_CHECKED : BST_UNCHECKED, 0);
      bool scope_top = initial->scope == SearchScope::kEntireRegistry;
      SendMessageW(state->scope_top, BM_SETCHECK, scope_top ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->scope_key, BM_SETCHECK, scope_top ? BST_UNCHECKED : BST_CHECKED, 0);
      if (!initial->start_key.empty()) {
        SetWindowTextW(state->scope_edit, initial->start_key.c_str());
      }
      bool new_tab = initial->result_mode == SearchResultMode::kNewTab;
      SendMessageW(state->result_reuse, BM_SETCHECK, new_tab ? BST_UNCHECKED : BST_CHECKED, 0);
      SendMessageW(state->result_new, BM_SETCHECK, new_tab ? BST_CHECKED : BST_UNCHECKED, 0);
    } else {
      SendMessageW(state->options_keys, BM_SETCHECK, BST_UNCHECKED, 0);
      SendMessageW(state->options_values, BM_SETCHECK, BST_CHECKED, 0);
      SendMessageW(state->options_data, BM_SETCHECK, BST_CHECKED, 0);
      SendMessageW(state->options_standard, BM_SETCHECK, BST_CHECKED, 0);
      SendMessageW(state->options_registry, BM_SETCHECK, state->registry_available ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->options_trace, BM_SETCHECK, state->trace_available ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->scope_top, BM_SETCHECK, BST_CHECKED, 0);
      SendMessageW(state->result_reuse, BM_SETCHECK, BST_CHECKED, 0);
    }
    SendMessageW(state->scope_recursive, BM_SETCHECK, state->recursive ? BST_CHECKED : BST_UNCHECKED, 0);

    UpdateDialogEnableState(state);

    Theme::Current().ApplyToChildren(hwnd);
    LayoutDialog(hwnd, state, font);
    return 0;
  }
  case WM_DESTROY:
    if (state && state->font) {
      DeleteObject(state->font);
      state->font = nullptr;
    }
    return 0;
  case WM_SIZE: {
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    LayoutDialog(hwnd, state, font);
    return 0;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return 1;
  }
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(hwnd);
      Theme::Current().ApplyToChildren(hwnd);
      InvalidateRect(hwnd, nullptr, TRUE);
    }
    return 0;
  }
  case WM_CTLCOLORSTATIC: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_STATIC));
  }
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_BTN));
  }
  case WM_CTLCOLORLISTBOX: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_LISTBOX));
  }
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_EDIT));
  }
  case WM_COMMAND: {
    if (!state) {
      return 0;
    }
    if (HIWORD(wparam) == CBN_DROPDOWN && LOWORD(wparam) == kScopeCombo) {
      ShowRootSelectionMenu(hwnd, state);
      SendMessageW(state->scope_combo, CB_SHOWDROPDOWN, FALSE, 0);
      return 0;
    }
    if (HIWORD(wparam) == BN_CLICKED) {
      switch (LOWORD(wparam)) {
      case kScopeTop:
      case kScopeKey:
      case kOptData:
      case kOptMinSize:
      case kOptMaxSize:
      case kOptStandardHives:
      case kOptRegistryRoot:
      case kOptTraceValues:
      case kExcludeEnable:
        UpdateDialogEnableState(state);
        break;
      default:
        break;
      }
    }
    switch (LOWORD(wparam)) {
    case kScopeBrowse: {
      std::wstring selected;
      if (ShowBrowseKeyDialog(hwnd, &selected)) {
        if (!selected.empty()) {
          SetWindowTextW(state->scope_edit, selected.c_str());
        }
        SendMessageW(state->scope_key, BM_SETCHECK, BST_CHECKED, 0);
        SendMessageW(state->scope_top, BM_SETCHECK, BST_UNCHECKED, 0);
        UpdateDialogEnableState(state);
      }
      return 0;
    }
    case kOptDataTypes:
      ShowDataTypesDialog(hwnd, &state->data_types);
      return 0;
    case kExcludeButton: {
      wchar_t buffer[2048] = {};
      GetWindowTextW(state->exclude_edit, buffer, static_cast<int>(_countof(buffer)));
      std::vector<std::wstring> items = SplitExcludePaths(buffer);
      std::wstring multiline;
      for (const auto& item : items) {
        if (item.empty()) {
          continue;
        }
        if (!multiline.empty()) {
          multiline.append(L"\r\n");
        }
        multiline.append(item);
      }
      if (PromptForMultiLineText(hwnd, L"Exclude Keys", L"Each line should include one key.", &multiline)) {
        std::vector<std::wstring> updated = SplitExcludePaths(multiline);
        std::wstring joined = JoinExcludePaths(updated);
        SetWindowTextW(state->exclude_edit, joined.c_str());
      }
      return 0;
    }
    case kFindButton: {
      wchar_t query[512] = {};
      GetWindowTextW(state->find_combo, query, static_cast<int>(_countof(query)));
      std::wstring query_text = query;
      if (query_text.empty()) {
        ui::ShowWarning(hwnd, L"Enter a search term.");
        return 0;
      }
      bool keys = SendMessageW(state->options_keys, BM_GETCHECK, 0, 0) == BST_CHECKED;
      bool values = SendMessageW(state->options_values, BM_GETCHECK, 0, 0) == BST_CHECKED;
      bool data = SendMessageW(state->options_data, BM_GETCHECK, 0, 0) == BST_CHECKED;
      if (!keys && !values && !data) {
        ui::ShowWarning(hwnd, L"Select at least one search option.");
        return 0;
      }
      bool standard_hives = state->options_standard && SendMessageW(state->options_standard, BM_GETCHECK, 0, 0) == BST_CHECKED;
      bool registry_root = state->options_registry && SendMessageW(state->options_registry, BM_GETCHECK, 0, 0) == BST_CHECKED;
      bool trace_values = state->options_trace && SendMessageW(state->options_trace, BM_GETCHECK, 0, 0) == BST_CHECKED;
      if (!state->registry_available) {
        registry_root = false;
      }
      if (!state->trace_available) {
        trace_values = false;
      }
      if (!standard_hives && !registry_root && !trace_values) {
        ui::ShowWarning(hwnd, L"Select at least one search source.");
        return 0;
      }

      SearchDialogResult result;
      result.criteria.query = query_text;
      result.criteria.search_keys = keys;
      result.criteria.search_values = values;
      result.criteria.search_data = data;
      result.criteria.match_case = SendMessageW(state->match_case, BM_GETCHECK, 0, 0) == BST_CHECKED;
      result.criteria.match_whole = SendMessageW(state->match_whole, BM_GETCHECK, 0, 0) == BST_CHECKED;
      result.criteria.use_regex = SendMessageW(state->use_regex, BM_GETCHECK, 0, 0) == BST_CHECKED;
      result.criteria.allowed_types = state->data_types;
      if (data) {
        if (state->min_size && SendMessageW(state->min_size, BM_GETCHECK, 0, 0) == BST_CHECKED) {
          wchar_t buffer[64] = {};
          GetWindowTextW(state->min_size_edit, buffer, static_cast<int>(_countof(buffer)));
          uint64_t value = 0;
          if (!ParseUint64(buffer, &value)) {
            ui::ShowWarning(hwnd, L"Enter a valid minimum data size.");
            return 0;
          }
          result.criteria.use_min_size = true;
          result.criteria.min_size = value;
        }
        if (state->max_size && SendMessageW(state->max_size, BM_GETCHECK, 0, 0) == BST_CHECKED) {
          wchar_t buffer[64] = {};
          GetWindowTextW(state->max_size_edit, buffer, static_cast<int>(_countof(buffer)));
          uint64_t value = 0;
          if (!ParseUint64(buffer, &value)) {
            ui::ShowWarning(hwnd, L"Enter a valid maximum data size.");
            return 0;
          }
          result.criteria.use_max_size = true;
          result.criteria.max_size = value;
        }
      }
      FILETIME modified_from = {};
      FILETIME modified_to = {};
      bool has_modified_from = GetDateTimeValue(state->modified_from, &modified_from);
      bool has_modified_to = GetDateTimeValue(state->modified_to, &modified_to);
      if (has_modified_from) {
        result.criteria.use_modified_from = true;
        result.criteria.modified_from = modified_from;
      }
      if (has_modified_to) {
        result.criteria.use_modified_to = true;
        result.criteria.modified_to = modified_to;
      }
      if (result.criteria.use_min_size && result.criteria.use_max_size && result.criteria.min_size > result.criteria.max_size) {
        ui::ShowWarning(hwnd, L"Minimum data size cannot exceed maximum data size.");
        return 0;
      }
      if (has_modified_from && has_modified_to && CompareFileTime(&modified_from, &modified_to) > 0) {
        ui::ShowWarning(hwnd, L"Modified date range is invalid.");
        return 0;
      }
      result.search_standard_hives = standard_hives;
      result.search_registry_root = registry_root;
      result.search_trace_values = trace_values;

      bool scope_top = SendMessageW(state->scope_top, BM_GETCHECK, 0, 0) == BST_CHECKED;
      result.scope = scope_top ? SearchScope::kEntireRegistry : SearchScope::kCurrentKey;
      state->recursive = SendMessageW(state->scope_recursive, BM_GETCHECK, 0, 0) == BST_CHECKED;
      result.criteria.recursive = scope_top ? true : state->recursive;
      result.result_mode = SendMessageW(state->result_new, BM_GETCHECK, 0, 0) == BST_CHECKED ? SearchResultMode::kNewTab : SearchResultMode::kReuseTab;

      if (SendMessageW(state->exclude_enable, BM_GETCHECK, 0, 0) == BST_CHECKED) {
        wchar_t buffer[512] = {};
        GetWindowTextW(state->exclude_edit, buffer, static_cast<int>(_countof(buffer)));
        result.exclude_paths = SplitExcludePaths(buffer);
      }

      result.root_paths.clear();
      result.start_key.clear();
      if (scope_top) {
        if (standard_hives) {
          for (size_t i = 0; i < state->root_selected.size(); ++i) {
            if (state->root_selected[i] && i < state->root_names.size()) {
              result.root_paths.push_back(state->root_names[i]);
            }
          }
          if (result.root_paths.empty()) {
            ui::ShowWarning(hwnd, L"Select at least one top-level key.");
            return 0;
          }
        }
      } else {
        wchar_t buffer[512] = {};
        GetWindowTextW(state->scope_edit, buffer, static_cast<int>(_countof(buffer)));
        result.start_key = buffer;
      }

      UpdateHistoryList(&state->history, query_text);
      SaveSearchHistory(state->history);

      if (state->out) {
        *state->out = result;
      }
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    }
    case kCancelButton:
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  }
  case WM_CLOSE:
    if (state) {
      RestoreOwnerWindow(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

HWND CreateSearchDialogWindow(HINSTANCE instance, HWND owner, SearchDialogState* state) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = SearchDialogProc;
  wc.hInstance = instance;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kDialogClass;
  RegisterClassW(&wc);

  return CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kDialogClass, L"Find", WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 600, 590, owner, nullptr, instance, state);
}

} // namespace

bool ShowBrowseKeyDialog(HWND owner, std::wstring* selected_path) {
  return ShowBrowseKeyDialogInternal(owner, selected_path);
}

bool ShowSearchDialog(HWND owner, SearchDialogResult* result, bool trace_available, bool registry_available) {
  if (!result) {
    return false;
  }
  HINSTANCE instance = GetModuleHandleW(nullptr);
  SearchDialogState state;
  state.out = result;
  state.owner = owner;
  state.trace_available = trace_available;
  state.registry_available = registry_available;
  HWND hwnd = CreateSearchDialogWindow(instance, owner, &state);
  if (!hwnd) {
    return false;
  }
  SetWindowTextW(hwnd, L"Find");

  Theme::Current().ApplyToWindow(hwnd);
  RECT rect = {};
  if (GetWindowRect(hwnd, &rect)) {
    CenterWindowToOwner(hwnd, owner, rect.right - rect.left, rect.bottom - rect.top);
  }

  EnableWindow(owner, FALSE);
  ShowWindow(hwnd, SW_SHOW);
  UpdateWindow(hwnd);
  FocusFindCombo(&state);

  MSG msg = {};
  while (IsWindow(hwnd) && GetMessageW(&msg, nullptr, 0, 0)) {
    if (!IsDialogMessageW(hwnd, &msg)) {
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  }

  RestoreOwnerWindow(owner, &state.owner_restored);
  return state.accepted;
}

} // namespace regkit
