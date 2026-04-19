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

#include "app/replace_dialog.h"

#include <algorithm>

#include <commctrl.h>
#include <uxtheme.h>

#include "app/search_dialog.h"
#include "app/theme.h"
#include "app/ui_helpers.h"

namespace regkit {

namespace {

constexpr wchar_t kDialogClass[] = L"RegKitReplaceDialog";
constexpr wchar_t kAppTitle[] = L"RegKit";

enum ControlId {
  kFindLabel = 100,
  kFindEdit = 101,
  kReplaceLabel = 102,
  kReplaceEdit = 103,
  kWhereGroup = 110,
  kKeyLabel = 111,
  kKeyEdit = 112,
  kKeyBrowse = 113,
  kOptionsGroup = 120,
  kRecursive = 121,
  kMatchCase = 122,
  kMatchWhole = 123,
  kUseRegex = 124,
  kReplaceButton = IDOK,
  kCancelButton = IDCANCEL,
};

struct ReplaceDialogState {
  HWND hwnd = nullptr;
  HWND find_edit = nullptr;
  HWND replace_edit = nullptr;
  HWND key_edit = nullptr;
  HWND key_browse = nullptr;
  HWND recursive = nullptr;
  HWND match_case = nullptr;
  HWND match_whole = nullptr;
  HWND use_regex = nullptr;
  HWND replace_button = nullptr;
  HWND cancel_button = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  ReplaceDialogResult* out = nullptr;
  bool accepted = false;
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

void RestoreOwnerWindow(HWND owner, bool* restored) {
  if (!owner || !restored || *restored) {
    return;
  }
  EnableWindow(owner, TRUE);
  SetActiveWindow(owner);
  SetForegroundWindow(owner);
  *restored = true;
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

void CenterWindowToOwner(HWND hwnd, HWND owner) {
  if (!hwnd) {
    return;
  }
  RECT rect = {};
  if (!GetWindowRect(hwnd, &rect)) {
    return;
  }
  int width = rect.right - rect.left;
  int height = rect.bottom - rect.top;
  RECT owner_rect = {};
  if (owner && GetWindowRect(owner, &owner_rect)) {
    int owner_w = owner_rect.right - owner_rect.left;
    int owner_h = owner_rect.bottom - owner_rect.top;
    int x = owner_rect.left + std::max(0, (owner_w - width) / 2);
    int y = owner_rect.top + std::max(0, (owner_h - height) / 2);
    SetWindowPos(hwnd, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
    return;
  }
  SetWindowPos(hwnd, nullptr, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
}

void LayoutDialog(HWND hwnd, ReplaceDialogState* state, HFONT font) {
  if (!hwnd || !state) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd, &client);
  int width = client.right - client.left;
  int x = 12;
  int y = 12;
  int label_w = 90;
  int key_label_w = 32;
  int key_gap = 4;
  int line_h = CalcDialogLineHeight(hwnd, font, 22);
  auto place_check = [&](HWND check, int x_pos, int y_pos, int width) {
    if (check) {
      SetWindowPos(check, nullptr, x_pos, y_pos, width, 18, SWP_NOZORDER);
    }
  };

  HWND find_label = GetDlgItem(hwnd, kFindLabel);
  SetWindowPos(find_label, nullptr, x, y + 4, label_w, 18, SWP_NOZORDER);
  int edit_w = width - x * 2 - label_w - 8;
  SetWindowPos(state->find_edit, nullptr, x + label_w + 8, y, edit_w, line_h, SWP_NOZORDER);
  y += line_h + 8;

  HWND replace_label = GetDlgItem(hwnd, kReplaceLabel);
  SetWindowPos(replace_label, nullptr, x, y + 4, label_w, 18, SWP_NOZORDER);
  SetWindowPos(state->replace_edit, nullptr, x + label_w + 8, y, edit_w, line_h, SWP_NOZORDER);
  y += line_h + 12;

  int group_w = width - x * 2;
  int where_h = line_h + 26;
  SetWindowPos(GetDlgItem(hwnd, kWhereGroup), nullptr, x, y, group_w, where_h, SWP_NOZORDER);
  int gx = x + 12;
  int gy = y + 18;
  int browse_w = 80;
  int key_w = group_w - 24 - key_label_w - key_gap - browse_w - 6;
  SetWindowPos(GetDlgItem(hwnd, kKeyLabel), nullptr, gx, gy + 4, key_label_w, 18, SWP_NOZORDER);
  SetWindowPos(state->key_edit, nullptr, gx + key_label_w + key_gap, gy, key_w, line_h, SWP_NOZORDER);
  SetWindowPos(state->key_browse, nullptr, gx + key_label_w + key_gap + key_w + 6, gy, browse_w, line_h, SWP_NOZORDER);
  y += where_h + 10;

  int options_h = line_h * 2 + 26;
  SetWindowPos(GetDlgItem(hwnd, kOptionsGroup), nullptr, x, y, group_w, options_h, SWP_NOZORDER);
  int ox = x + 12;
  int oy = y + 18;
  int option_col2_x = ox + 180;
  place_check(state->recursive, ox, oy, 120);
  place_check(state->match_case, option_col2_x, oy, 120);
  place_check(state->match_whole, ox, oy + 22, 160);
  place_check(state->use_regex, option_col2_x, oy + 22, 190);

  int btn_w = 90;
  int btn_h = std::max(20, line_h);
  int btn_gap = 10;
  int btn_margin = 10;
  int btn_y = client.bottom - btn_h - 6;
  int cancel_x = width - btn_margin - btn_w;
  int replace_x = cancel_x - btn_gap - btn_w;
  SetWindowPos(state->replace_button, nullptr, replace_x, btn_y, btn_w, btn_h, SWP_NOZORDER);
  SetWindowPos(state->cancel_button, nullptr, cancel_x, btn_y, btn_w, btn_h, SWP_NOZORDER);

  ApplyFont(hwnd, font);
  ApplyFont(find_label, font);
  ApplyFont(replace_label, font);
}

LRESULT CALLBACK ReplaceDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<ReplaceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<ReplaceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, L"Replace");
    state->font = CreateDialogFont();
    HFONT font = state->font;

    CreateWindowExW(0, L"STATIC", L"Find what:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kFindLabel), nullptr, nullptr);
    state->find_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kFindEdit), nullptr, nullptr);

    CreateWindowExW(0, L"STATIC", L"Replace with:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kReplaceLabel), nullptr, nullptr);
    state->replace_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kReplaceEdit), nullptr, nullptr);

    CreateWindowExW(0, L"BUTTON", L"Where to search", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kWhereGroup), nullptr, nullptr);
    CreateWindowExW(0, L"STATIC", L"Key:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kKeyLabel), nullptr, nullptr);
    state->key_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kKeyEdit), nullptr, nullptr);
    state->key_browse = CreateWindowExW(0, L"BUTTON", L"Browse...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kKeyBrowse), nullptr, nullptr);

    CreateWindowExW(0, L"BUTTON", L"Options", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptionsGroup), nullptr, nullptr);
    state->recursive = CreateWindowExW(0, L"BUTTON", L"Recursive", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kRecursive), nullptr, nullptr);
    state->match_case = CreateWindowExW(0, L"BUTTON", L"Match case", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kMatchCase), nullptr, nullptr);
    state->match_whole = CreateWindowExW(0, L"BUTTON", L"Match whole string", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kMatchWhole), nullptr, nullptr);
    state->use_regex = CreateWindowExW(0, L"BUTTON", L"Use regular expressions", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kUseRegex), nullptr, nullptr);

    state->replace_button = CreateWindowExW(0, L"BUTTON", L"Replace", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kReplaceButton), nullptr, nullptr);
    state->cancel_button = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kCancelButton), nullptr, nullptr);

    if (state->out) {
      SetWindowTextW(state->find_edit, state->out->find_text.c_str());
      SetWindowTextW(state->replace_edit, state->out->replace_text.c_str());
      SetWindowTextW(state->key_edit, state->out->start_key.c_str());
      SendMessageW(state->recursive, BM_SETCHECK, state->out->recursive ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->match_case, BM_SETCHECK, state->out->match_case ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->match_whole, BM_SETCHECK, state->out->match_whole ? BST_CHECKED : BST_UNCHECKED, 0);
      SendMessageW(state->use_regex, BM_SETCHECK, state->out->use_regex ? BST_CHECKED : BST_UNCHECKED, 0);
    } else {
      SendMessageW(state->recursive, BM_SETCHECK, BST_CHECKED, 0);
    }

    EnumChildWindows(
        hwnd,
        [](HWND child, LPARAM param) -> BOOL {
          HFONT font_handle = reinterpret_cast<HFONT>(param);
          ApplyFont(child, font_handle);
          return TRUE;
        },
        reinterpret_cast<LPARAM>(font));

    Theme::Current().ApplyToWindow(hwnd);
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
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_EDIT));
  }
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_BTN));
  }
  case WM_COMMAND: {
    if (!state) {
      return 0;
    }
    switch (LOWORD(wparam)) {
    case kKeyBrowse: {
      std::wstring selected;
      if (ShowBrowseKeyDialog(hwnd, &selected)) {
        if (!selected.empty()) {
          SetWindowTextW(state->key_edit, selected.c_str());
        }
      }
      return 0;
    }
    case kReplaceButton: {
      wchar_t find_text[512] = {};
      GetWindowTextW(state->find_edit, find_text, static_cast<int>(_countof(find_text)));
      std::wstring find_value = find_text;
      if (find_value.empty()) {
        ui::ShowError(hwnd, L"Enter text to find.");
        return 0;
      }
      wchar_t replace_text[512] = {};
      GetWindowTextW(state->replace_edit, replace_text, static_cast<int>(_countof(replace_text)));
      wchar_t key_text[512] = {};
      GetWindowTextW(state->key_edit, key_text, static_cast<int>(_countof(key_text)));

      if (state->out) {
        state->out->find_text = find_value;
        state->out->replace_text = replace_text;
        state->out->start_key = key_text;
        state->out->recursive = SendMessageW(state->recursive, BM_GETCHECK, 0, 0) == BST_CHECKED;
        state->out->match_case = SendMessageW(state->match_case, BM_GETCHECK, 0, 0) == BST_CHECKED;
        state->out->match_whole = SendMessageW(state->match_whole, BM_GETCHECK, 0, 0) == BST_CHECKED;
        state->out->use_regex = SendMessageW(state->use_regex, BM_GETCHECK, 0, 0) == BST_CHECKED;
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

HWND CreateReplaceDialogWindow(HINSTANCE instance, HWND owner, ReplaceDialogState* state) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = ReplaceDialogProc;
  wc.hInstance = instance;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kDialogClass;
  RegisterClassW(&wc);

  return CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kDialogClass, L"Replace", WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 520, 280, owner, nullptr, instance, state);
}

} // namespace

bool ShowReplaceDialog(HWND owner, ReplaceDialogResult* result) {
  if (!result) {
    return false;
  }
  HINSTANCE instance = GetModuleHandleW(nullptr);
  ReplaceDialogState state;
  state.out = result;
  state.owner = owner;
  HWND hwnd = CreateReplaceDialogWindow(instance, owner, &state);
  if (!hwnd) {
    return false;
  }
  SetWindowTextW(hwnd, L"Replace");

  Theme::Current().ApplyToWindow(hwnd);
  CenterWindowToOwner(hwnd, owner);

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
  return state.accepted;
}

} // namespace regkit
