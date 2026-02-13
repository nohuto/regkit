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

#include "app/font_dialog.h"

#include <algorithm>
#include <string>
#include <vector>

#include <commctrl.h>

#include "app/theme.h"
#include "app/ui_helpers.h"

namespace regkit {

namespace {

constexpr wchar_t kFontDialogClass[] = L"RegKitFontDialog";
constexpr wchar_t kAppTitle[] = L"RegKit";

constexpr int kIdDefault = 100;
constexpr int kIdCustom = 101;
constexpr int kIdFontList = 102;
constexpr int kIdStyleList = 103;
constexpr int kIdSizeEdit = 104;

struct FontDialogState {
  HWND hwnd = nullptr;
  HWND radio_default = nullptr;
  HWND radio_custom = nullptr;
  HWND font_list = nullptr;
  HWND style_list = nullptr;
  HWND size_edit = nullptr;
  HWND size_spin = nullptr;
  HWND ok_btn = nullptr;
  HWND cancel_btn = nullptr;
  HWND owner = nullptr;
  HFONT ui_font = nullptr;
  bool use_default = true;
  LOGFONTW font = {};
  bool accepted = false;
  std::vector<std::wstring> fonts;
  bool owner_restored = false;
};

int CALLBACK EnumFontFamExProc(const LOGFONTW* lf, const TEXTMETRICW*, DWORD, LPARAM lparam) {
  if (!lf) {
    return 1;
  }
  auto* fonts = reinterpret_cast<std::vector<std::wstring>*>(lparam);
  if (fonts) {
    fonts->push_back(lf->lfFaceName);
  }
  return 1;
}

std::vector<std::wstring> EnumerateFonts() {
  std::vector<std::wstring> fonts;
  LOGFONTW lf = {};
  lf.lfCharSet = DEFAULT_CHARSET;
  HDC hdc = GetDC(nullptr);
  EnumFontFamiliesExW(hdc, &lf, EnumFontFamExProc, reinterpret_cast<LPARAM>(&fonts), 0);
  ReleaseDC(nullptr, hdc);
  std::sort(fonts.begin(), fonts.end(), [](const std::wstring& a, const std::wstring& b) { return _wcsicmp(a.c_str(), b.c_str()) < 0; });
  fonts.erase(std::unique(fonts.begin(), fonts.end(), [](const std::wstring& a, const std::wstring& b) { return _wcsicmp(a.c_str(), b.c_str()) == 0; }), fonts.end());
  return fonts;
}

void ApplyDialogFont(HWND hwnd, HFONT font) {
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

void CenterToOwner(HWND hwnd, HWND owner, int width, int height) {
  RECT owner_rect = {};
  if (owner && GetWindowRect(owner, &owner_rect)) {
    int x = owner_rect.left + (owner_rect.right - owner_rect.left - width) / 2;
    int y = owner_rect.top + (owner_rect.bottom - owner_rect.top - height) / 2;
    SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER);
    return;
  }
  SetWindowPos(hwnd, nullptr, CW_USEDEFAULT, CW_USEDEFAULT, width, height, SWP_NOZORDER);
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

int FontPointSize(const LOGFONTW& font) {
  HDC hdc = GetDC(nullptr);
  int size = 9;
  if (font.lfHeight != 0) {
    size = MulDiv(-font.lfHeight, 72, GetDeviceCaps(hdc, LOGPIXELSY));
  }
  ReleaseDC(nullptr, hdc);
  return size;
}

void SetFontListSelection(FontDialogState* state) {
  if (!state || !state->font_list) {
    return;
  }
  int index = 0;
  int match = -1;
  for (const auto& font : state->fonts) {
    if (_wcsicmp(font.c_str(), state->font.lfFaceName) == 0) {
      match = index;
      break;
    }
    ++index;
  }
  if (match >= 0) {
    SendMessageW(state->font_list, LB_SETCURSEL, match, 0);
  }
}

int StyleIndex(const LOGFONTW& font) {
  bool bold = font.lfWeight >= FW_BOLD;
  bool italic = font.lfItalic != 0;
  if (bold && italic) {
    return 3;
  }
  if (bold) {
    return 1;
  }
  if (italic) {
    return 2;
  }
  return 0;
}

void ApplyEnableState(FontDialogState* state) {
  if (!state) {
    return;
  }
  bool enable = !state->use_default;
  if (state->font_list) {
    EnableWindow(state->font_list, enable);
  }
  if (state->style_list) {
    EnableWindow(state->style_list, enable);
  }
  if (state->size_edit) {
    EnableWindow(state->size_edit, enable);
  }
  if (state->size_spin) {
    EnableWindow(state->size_spin, enable);
  }
}

void UpdateFontFromControls(FontDialogState* state) {
  if (!state || !state->font_list || !state->style_list || !state->size_edit) {
    return;
  }
  int index = static_cast<int>(SendMessageW(state->font_list, LB_GETCURSEL, 0, 0));
  if (index != LB_ERR && index >= 0 && static_cast<size_t>(index) < state->fonts.size()) {
    wcsncpy_s(state->font.lfFaceName, state->fonts[static_cast<size_t>(index)].c_str(), _TRUNCATE);
  }
  int style = static_cast<int>(SendMessageW(state->style_list, CB_GETCURSEL, 0, 0));
  state->font.lfWeight = (style == 1 || style == 3) ? FW_BOLD : FW_NORMAL;
  state->font.lfItalic = (style == 2 || style == 3) ? TRUE : FALSE;

  wchar_t size_text[16] = {};
  GetWindowTextW(state->size_edit, size_text, static_cast<int>(_countof(size_text)));
  int size = _wtoi(size_text);
  if (size <= 0) {
    size = 9;
  }
  HDC hdc = GetDC(nullptr);
  state->font.lfHeight = -MulDiv(size, GetDeviceCaps(hdc, LOGPIXELSY), 72);
  ReleaseDC(nullptr, hdc);
}

LRESULT CALLBACK FontDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<FontDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<FontDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, kAppTitle);

    state->radio_default = CreateWindowExW(0, L"BUTTON", L"Use default system font", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON, 12, 12, 200, 16, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdDefault)), nullptr, nullptr);
    state->radio_custom = CreateWindowExW(0, L"BUTTON", L"Use custom font", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON, 12, 32, 200, 16, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdCustom)), nullptr, nullptr);
    state->font_list = CreateWindowExW(0, L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LBS_NOTIFY | WS_VSCROLL, 12, 58, 210, 200, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdFontList)), nullptr, nullptr);
    CreateWindowExW(0, L"STATIC", L"Style:", WS_CHILD | WS_VISIBLE, 232, 58, 120, 16, hwnd, nullptr, nullptr, nullptr);
    state->style_list = CreateWindowExW(0, WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | CBS_DROPDOWNLIST, 232, 74, 150, 200, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdStyleList)), nullptr, nullptr);
    CreateWindowExW(0, L"STATIC", L"Size:", WS_CHILD | WS_VISIBLE, 232, 112, 120, 16, hwnd, nullptr, nullptr, nullptr);
    state->size_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER, 232, 128, 48, 20, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kIdSizeEdit)), nullptr, nullptr);
    state->size_spin = CreateWindowExW(0, UPDOWN_CLASSW, L"", WS_CHILD | WS_VISIBLE | UDS_SETBUDDYINT | UDS_ARROWKEYS, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    if (state->size_spin) {
      SendMessageW(state->size_spin, UDM_SETRANGE32, 6, 72);
      SendMessageW(state->size_spin, UDM_SETBUDDY, reinterpret_cast<WPARAM>(state->size_edit), 0);
      SetWindowPos(state->size_spin, nullptr, 232 + 48, 128, 16, 20, SWP_NOZORDER);
    }

    state->ok_btn = CreateWindowExW(0, L"BUTTON", L"Apply", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 225, 252, 80, 22, hwnd, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);
    state->cancel_btn = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 315, 252, 80, 22, hwnd, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);

    state->ui_font = ui::DefaultUIFont();
    ApplyDialogFont(hwnd, state->ui_font);
    Theme::Current().ApplyToWindow(hwnd);
    Theme::Current().ApplyToChildren(hwnd);

    state->fonts = EnumerateFonts();
    for (const auto& font : state->fonts) {
      SendMessageW(state->font_list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(font.c_str()));
    }

    SendMessageW(state->style_list, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Regular"));
    SendMessageW(state->style_list, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Bold"));
    SendMessageW(state->style_list, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Italic"));
    SendMessageW(state->style_list, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Bold Italic"));

    SetFontListSelection(state);
    SendMessageW(state->style_list, CB_SETCURSEL, StyleIndex(state->font), 0);
    int size_value = FontPointSize(state->font);
    wchar_t size_text[16] = {};
    swprintf_s(size_text, L"%d", size_value);
    SetWindowTextW(state->size_edit, size_text);
    if (state->size_spin) {
      SendMessageW(state->size_spin, UDM_SETPOS32, 0, size_value);
    }

    CheckDlgButton(hwnd, state->use_default ? kIdDefault : kIdCustom, BST_CHECKED);
    ApplyEnableState(state);
    return 0;
  }
  case WM_DESTROY:
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return 0;
  case WM_SETTINGCHANGE:
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(hwnd);
      Theme::Current().ApplyToChildren(hwnd);
      InvalidateRect(hwnd, nullptr, TRUE);
    }
    return 0;
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, hwnd, CTLCOLOR_DLG));
  }
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
    }
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_COMMAND: {
    int id = LOWORD(wparam);
    if (id == kIdDefault || id == kIdCustom) {
      state->use_default = (id == kIdDefault);
      ApplyEnableState(state);
      return 0;
    }
    if (id == kIdFontList && HIWORD(wparam) == LBN_SELCHANGE) {
      return 0;
    }
    if (id == IDOK) {
      if (!state->use_default) {
        UpdateFontFromControls(state);
      }
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    }
    if (id == IDCANCEL) {
      state->accepted = false;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    }
    break;
  }
  case WM_CLOSE:
    if (state) {
      state->accepted = false;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

} // namespace

bool ShowFontDialog(HWND owner, bool use_default, const LOGFONTW& current, FontDialogResult* out) {
  if (!out) {
    return false;
  }
  WNDCLASSW wc = {};
  wc.lpfnWndProc = FontDialogProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kFontDialogClass;
  RegisterClassW(&wc);

  FontDialogState state;
  state.use_default = use_default;
  state.font = current;
  state.owner = owner;
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kFontDialogClass, kAppTitle, WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 420, 320, owner, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    return false;
  }

  Theme::Current().ApplyToWindow(hwnd);
  CenterToOwner(hwnd, owner, 420, 320);

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
  if (!state.accepted) {
    return false;
  }

  out->use_default = state.use_default;
  out->font = state.font;
  return true;
}

} // namespace regkit
