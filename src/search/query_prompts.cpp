// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/query_prompts.h"

#include <algorithm>
#include <vector>

#include <commctrl.h>

#include "appearance/dialog_layout.h"
#include "appearance/default_font.h"
#include "appearance/feedback.h"
#include "appearance/theme.h"
#include "browse/key_tree.h"
#include "registry/registry_path.h"
#include "registry/registry_store.h"
#include "registry/value_format.h"

namespace regkit::query_prompts {

constexpr wchar_t kAppTitle[] = L"RegKit";

HFONT CreateDialogFont() {
  return ui::DefaultUIFont();
}

struct DataTypeItem {
  DWORD type = 0;
  std::wstring label;
};

struct BaseTypeItem {
  DWORD type = 0;
  const wchar_t* label = nullptr;
};

constexpr BaseTypeItem kBaseDataTypes[] = {
    {REG_SZ, L"REG_SZ"},
    {REG_EXPAND_SZ, L"REG_EXPAND_SZ"},
    {REG_MULTI_SZ, L"REG_MULTI_SZ"},
    {REG_DWORD, L"DWORD (32-bit)"},
    {REG_QWORD, L"QWORD (64-bit)"},
    {REG_BINARY, L"REG_BINARY"},
    {REG_NONE, L"REG_NONE"},
    {REG_DWORD_BIG_ENDIAN, L"REG_DWORD_BIG_ENDIAN"},
    {REG_LINK, L"REG_LINK"},
    {REG_RESOURCE_LIST, L"REG_RESOURCE_LIST"},
    {REG_FULL_RESOURCE_DESCRIPTOR, L"REG_FULL_RESOURCE_DESCRIPTOR"},
    {REG_RESOURCE_REQUIREMENTS_LIST, L"REG_RESOURCE_REQUIREMENTS_LIST"},
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
      items.push_back({type, value_format::TypeName(type)});
    }
  }
  return items;
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
      appearance::SetControlFont(check, font);
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
    appearance::SetControlFont(state->select_all, font);
    appearance::SetControlFont(state->clear_all, font);
    appearance::SetControlFont(state->ok_button, font);
    appearance::SetControlFont(state->cancel_button, font);
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
    int type = (msg == WM_CTLCOLORSTATIC) ? CTLCOLOR_STATIC : (msg == WM_CTLCOLORBTN) ? CTLCOLOR_BTN
                                                                                      : CTLCOLOR_EDIT;
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
    case 100: {
      for (HWND check : state->checks) {
        SendMessageW(check, BM_SETCHECK, BST_CHECKED, 0);
      }
      return 0;
    }
    case 101: {
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
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    }
    case IDCANCEL:
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  }
  case WM_CLOSE:
    if (state) {
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

bool ShowDataTypes(HWND owner, std::vector<DWORD>* types) {
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
  appearance::PositionDialog(hwnd, owner, width, height);

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
  appearance::RestoreDialogOwner(owner, &state.owner_restored);

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
    appearance::SetControlFont(state->tree_hwnd, font);
    appearance::SetControlFont(state->ok_button, font);
    appearance::SetControlFont(state->cancel_button, font);

    state->tree.Create(hwnd, GetModuleHandleW(nullptr), 1);
    std::vector<RegistryRootEntry> roots = RegistryStore::DefaultRoots();
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
    int type = (msg == WM_CTLCOLORSTATIC) ? CTLCOLOR_STATIC : (msg == WM_CTLCOLORBTN) ? CTLCOLOR_BTN
                                                                                      : CTLCOLOR_EDIT;
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
          state->selected_path = registry_path::Build(*node);
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
            return CDRF_DODEFAULT;
          }
          draw->clrText = theme.TextColor();
          draw->clrTextBk = theme.PanelColor();
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
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    }
    case IDCANCEL:
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  }
  case WM_CLOSE:
    if (state) {
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

bool ShowRegistryKey(HWND owner, std::wstring* selected_path) {
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
  appearance::PositionDialog(hwnd, owner, 420, 420);

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
  appearance::RestoreDialogOwner(owner, &state.owner_restored);

  if (state.accepted) {
    *selected_path = state.selected_path;
    return true;
  }
  return false;
}

} // namespace regkit::query_prompts
