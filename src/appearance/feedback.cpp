// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "appearance/feedback.h"

#include "appearance/default_font.h"
#include "appearance/dialog_layout.h"

#include <algorithm>
#include <cwchar>
#include <vector>

#include <commctrl.h>
#include <shellapi.h>
#include <uxtheme.h>

#include "appearance/gdi_cache.h"
#include "appearance/theme.h"
#include "win32/file_dialog.h"
#include "win32/shell_paths.h"

namespace regkit::ui {

namespace {
#ifndef WC_LINK
#define WC_LINK L"SysLink"
#endif

constexpr wchar_t kErrorClass[] = L"RegKitErrorDialog";
constexpr wchar_t kChoiceClass[] = L"RegKitChoiceDialog";
constexpr wchar_t kAboutClass[] = L"RegKitAboutDialog";
constexpr wchar_t kAppTitle[] = L"RegKit";

struct ErrorDialogState {
  HWND hwnd = nullptr;
  HWND text = nullptr;
  HWND ok_btn = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  std::wstring title;
  std::wstring message;
  bool accepted = false;
  bool owner_restored = false;
};

struct ChoiceDialogState {
  HWND hwnd = nullptr;
  HWND icon = nullptr;
  HWND text = nullptr;
  HWND detail_edit = nullptr;
  HWND detail_tip = nullptr;
  HWND yes_btn = nullptr;
  HWND no_btn = nullptr;
  HWND cancel_btn = nullptr;
  HWND default_btn = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  std::wstring title;
  std::wstring message;
  std::wstring detail;
  std::wstring yes_label;
  std::wstring no_label;
  std::wstring cancel_label;
  int yes_button_width_dlu = 0;
  int default_id = 0;
  PCWSTR icon_id = nullptr;
  int button_width_dlu = 45;
  int result = IDCANCEL;
  bool accepted = false;
  bool owner_restored = false;
};

struct AboutDialogState {
  HWND hwnd = nullptr;
  HWND credits = nullptr;
  HWND repo_link = nullptr;
  HWND discord_link = nullptr;
  HWND website_link = nullptr;
  HWND email_link = nullptr;
  HWND ok_btn = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  bool accepted = false;
  bool owner_restored = false;
};

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
  RECT work_area = {};
  if (SystemParametersInfoW(SPI_GETWORKAREA, 0, &work_area, 0)) {
    int work_w = work_area.right - work_area.left;
    int work_h = work_area.bottom - work_area.top;
    int x = work_area.left + std::max(0, (work_w - width) / 2);
    int y = work_area.top + std::max(0, (work_h - height) / 2);
    SetWindowPos(hwnd, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
  }
}

void ApplyConfirmFonts(HWND hwnd, HFONT font) {
  if (!font) {
    return;
  }
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

int TextWidth(HWND window, HFONT font, const std::wstring& text) {
  if (text.empty()) {
    return 0;
  }
  HDC dc = GetDC(window);
  if (!dc) {
    return 0;
  }
  HGDIOBJ old_font = font ? SelectObject(dc, font) : nullptr;
  SIZE size = {};
  GetTextExtentPoint32W(dc, text.c_str(), static_cast<int>(text.size()), &size);
  if (old_font) {
    SelectObject(dc, old_font);
  }
  ReleaseDC(window, dc);
  return static_cast<int>(size.cx);
}

int TextBlockHeight(HWND window, HFONT font, const std::wstring& text, int width) {
  HDC dc = GetDC(window);
  HFONT old_font = font ? reinterpret_cast<HFONT>(SelectObject(dc, font)) : nullptr;
  RECT rect = {0, 0, width, 0};
  DrawTextW(dc, text.c_str(), -1, &rect, DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX);
  if (old_font) {
    SelectObject(dc, old_font);
  }
  ReleaseDC(window, dc);
  return rect.bottom;
}

constexpr int kMaxDetailLines = 10;

int DetailLineCount(const std::wstring& detail) {
  if (detail.empty()) {
    return 1;
  }
  return 1 + static_cast<int>(std::count(detail.begin(), detail.end(), L'\n'));
}

void FitChoiceDialogToContent(HWND hwnd, ChoiceDialogState* state) {
  if (!hwnd || !state || !state->detail_edit) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd, &client);
  const int padding = 12;
  const int icon_w = state->icon_id ? 32 : 0;
  const int icon_gap = state->icon_id ? 12 : 0;
  const int base_units = GetDialogBaseUnits();
  const int base_y = std::max(1, static_cast<int>(HIWORD(base_units)));
  const int content_w = (client.right - client.left) - padding - icon_w - icon_gap - padding;
  const int text_h = TextBlockHeight(hwnd, state->font, state->message, content_w);
  const int lines = std::min(DetailLineCount(state->detail), kMaxDetailLines);
  const int detail_h = MulDiv(12, base_y, 8) + (lines - 1) * base_y;
  const int btn_h = MulDiv(11, base_y, 8);
  const int needed = padding + text_h + MulDiv(2, base_y, 8) + detail_h +
                     MulDiv(6, base_y, 8) + btn_h + padding;
  if (needed <= client.bottom - client.top) {
    return;
  }
  RECT frame = {0, 0, client.right - client.left, needed};
  const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_STYLE));
  const DWORD ex_style = static_cast<DWORD>(GetWindowLongPtrW(hwnd, GWL_EXSTYLE));
  AdjustWindowRectEx(&frame, style, FALSE, ex_style);
  SetWindowPos(hwnd, nullptr, 0, 0, frame.right - frame.left, frame.bottom - frame.top,
               SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
}

void LayoutChoiceDialog(HWND hwnd, ChoiceDialogState* state) {
  if (!hwnd || !state) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  int padding = 12;
  int icon_w = state->icon_id ? 32 : 0;
  int icon_gap = state->icon_id ? 12 : 0;
  int base_units = GetDialogBaseUnits();
  int base_x = std::max(1, static_cast<int>(LOWORD(base_units)));
  int base_y = std::max(1, static_cast<int>(HIWORD(base_units)));
  int btn_w = MulDiv(state->button_width_dlu, base_x, 4);
  int btn_h = MulDiv(11, base_y, 8);
  int btn_y = height - padding - btn_h;
  int gap = MulDiv(6, base_x, 4);
  if (state->icon) {
    SetWindowPos(state->icon, nullptr, padding, padding + 2, 32, 32, SWP_NOZORDER);
  }
  const int text_x = padding + icon_w + icon_gap;
  const int content_w = width - text_x - padding;
  if (state->detail_edit) {
    const int text_h = TextBlockHeight(hwnd, state->font, state->message, content_w);
    const int detail_y = padding + text_h + MulDiv(2, base_y, 8);
    if (state->text) {
      SetWindowPos(state->text, nullptr, text_x, padding, content_w, text_h, SWP_NOZORDER);
    }
    const int lines = std::min(DetailLineCount(state->detail), kMaxDetailLines);
    const int detail_h = MulDiv(12, base_y, 8) + (lines - 1) * base_y;
    SetWindowPos(state->detail_edit, nullptr, text_x, detail_y, content_w, detail_h, SWP_NOZORDER);
    if (state->detail_tip) {
      SendMessageW(state->detail_tip, TTM_ACTIVATE,
                   lines == 1 && TextWidth(hwnd, state->font, state->detail) > content_w - 8, 0);
    }
  } else if (state->text) {
    SetWindowPos(state->text, nullptr, text_x, padding, content_w, btn_y - padding, SWP_NOZORDER);
  }

  HWND buttons[] = {state->yes_btn, state->no_btn, state->cancel_btn};
  const int widths[] = {state->yes_button_width_dlu > 0 ? MulDiv(state->yes_button_width_dlu, base_x, 4) : btn_w, btn_w, btn_w};
  int total_w = 0;
  int button_count = 0;
  for (int i = 0; i < 3; ++i) {
    if (buttons[i]) {
      total_w += widths[i];
      ++button_count;
    }
  }
  if (button_count == 0) {
    return;
  }
  total_w += gap * (button_count - 1);
  int x = std::max(padding, width - padding - total_w);
  for (int i = 0; i < 3; ++i) {
    if (!buttons[i]) {
      continue;
    }
    SetWindowPos(buttons[i], nullptr, x, btn_y, widths[i], btn_h, SWP_NOZORDER);
    x += widths[i] + gap;
  }
}

void LayoutErrorDialog(HWND hwnd, ErrorDialogState* state) {
  if (!hwnd || !state) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  int padding = 12;
  int btn_w = 80;
  int btn_h = 22;
  int btn_y = height - padding - btn_h;
  int text_h = btn_y - padding;
  if (state->text) {
    SetWindowPos(state->text, nullptr, padding, padding, width - padding * 2, text_h, SWP_NOZORDER);
  }
  int ok_x = width - padding - btn_w;
  if (state->ok_btn) {
    SetWindowPos(state->ok_btn, nullptr, ok_x, btn_y, btn_w, btn_h, SWP_NOZORDER);
  }
}

void LayoutAboutDialog(HWND hwnd, AboutDialogState* state) {
  if (!hwnd || !state) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  int padding = 12;
  int line_h = 20;
  int gap = 6;
  int btn_w = 80;
  int btn_h = 22;
  int btn_y = height - padding - btn_h;
  int text_w = width - padding * 2;
  int y = padding;

  if (state->credits) {
    SetWindowPos(state->credits, nullptr, padding, y, text_w, line_h, SWP_NOZORDER);
  }
  y += line_h + gap;

  if (state->repo_link) {
    SetWindowPos(state->repo_link, nullptr, padding, y, text_w, line_h, SWP_NOZORDER);
  }
  y += line_h + gap;
  if (state->discord_link) {
    SetWindowPos(state->discord_link, nullptr, padding, y, text_w, line_h, SWP_NOZORDER);
  }
  y += line_h + gap;
  if (state->website_link) {
    SetWindowPos(state->website_link, nullptr, padding, y, text_w, line_h, SWP_NOZORDER);
  }
  y += line_h + gap;
  if (state->email_link) {
    SetWindowPos(state->email_link, nullptr, padding, y, text_w, line_h, SWP_NOZORDER);
  }

  int ok_x = width - padding - btn_w;
  if (state->ok_btn) {
    SetWindowPos(state->ok_btn, nullptr, ok_x, btn_y, btn_w, btn_h, SWP_NOZORDER);
  }
}

LRESULT CALLBACK ChoiceDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<ChoiceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<ChoiceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, state->title.empty() ? kAppTitle : state->title.c_str());
    state->font = DefaultUIFont();
    if (state->icon_id) {
      state->icon = CreateWindowExW(0, L"STATIC", nullptr, WS_CHILD | WS_VISIBLE | SS_ICON, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
      HICON icon = LoadIconW(nullptr, state->icon_id);
      if (state->icon && icon) {
        SendMessageW(state->icon, STM_SETICON, reinterpret_cast<WPARAM>(icon), 0);
      }
    }
    state->text = CreateWindowExW(0, L"STATIC", state->message.c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    if (!state->detail.empty()) {
      const int detail_lines = DetailLineCount(state->detail);
      DWORD detail_style = WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_READONLY;
      if (detail_lines > 1) {
        detail_style |= ES_MULTILINE;
      }
      if (detail_lines > kMaxDetailLines) {
        detail_style |= WS_VSCROLL;
      }
      state->detail_edit = CreateWindowExW(0, L"EDIT", state->detail.c_str(),
                                           detail_style,
                                           0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
      state->detail_tip = CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr,
                                          WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
                                          CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                                          CW_USEDEFAULT, hwnd, nullptr, nullptr, nullptr);
      if (state->detail_tip && state->detail_edit) {
        TOOLINFOW info = {};
        info.cbSize = sizeof(info);
        info.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
        info.hwnd = hwnd;
        info.uId = reinterpret_cast<UINT_PTR>(state->detail_edit);
        info.lpszText = const_cast<wchar_t*>(state->detail.c_str());
        SendMessageW(state->detail_tip, TTM_ADDTOOL, 0, reinterpret_cast<LPARAM>(&info));
        SendMessageW(state->detail_tip, TTM_SETMAXTIPWIDTH, 0, 900);
        AllowDarkModeForWindow(state->detail_tip, Theme::UseDarkMode());
        SetWindowTheme(state->detail_tip, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
      }
    }
    if (!state->yes_label.empty()) {
      state->yes_btn = CreateWindowExW(0, L"BUTTON", state->yes_label.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDYES), nullptr, nullptr);
    }
    if (!state->no_label.empty()) {
      state->no_btn = CreateWindowExW(0, L"BUTTON", state->no_label.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDNO), nullptr, nullptr);
    }
    if (!state->cancel_label.empty()) {
      state->cancel_btn = CreateWindowExW(0, L"BUTTON", state->cancel_label.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);
    }
    state->default_btn = nullptr;
    if (state->default_id == IDYES) {
      state->default_btn = state->yes_btn;
    } else if (state->default_id == IDNO) {
      state->default_btn = state->no_btn;
    } else if (state->default_id == IDCANCEL) {
      state->default_btn = state->cancel_btn;
    }
    if (!state->default_btn) {
      state->default_btn = state->yes_btn ? state->yes_btn
                          : (state->no_btn ? state->no_btn : state->cancel_btn);
    }
    if (state->default_btn) {
      SetWindowLongPtrW(state->default_btn, GWL_STYLE,
                        GetWindowLongPtrW(state->default_btn, GWL_STYLE) | BS_DEFPUSHBUTTON);
    }

    ApplyConfirmFonts(hwnd, state->font);
    Theme::Current().ApplyToWindow(hwnd);
    Theme::Current().ApplyToChildren(hwnd);
    FitChoiceDialogToContent(hwnd, state);
    LayoutChoiceDialog(hwnd, state);
    if (state->default_btn) {
      SetFocus(state->default_btn);
    }
    return 0;
  }
  case WM_SIZE:
    LayoutChoiceDialog(hwnd, state);
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
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLORDLG:
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
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, type));
  }
  case DM_GETDEFID:
    if (state && state->default_btn) {
      return MAKELRESULT(GetDlgCtrlID(state->default_btn), DC_HASDEFID);
    }
    break;
  case WM_COMMAND:
    switch (LOWORD(wparam)) {
    case IDYES:
      state->result = IDYES;
      state->accepted = true;
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    case IDNO:
      state->result = IDNO;
      state->accepted = true;
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    case IDCANCEL:
      state->result = IDCANCEL;
      state->accepted = true;
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  case WM_DESTROY:
    if (state && state->font) {
      DeleteObject(state->font);
      state->font = nullptr;
    }
    return 0;
  case WM_CLOSE:
    if (state) {
      state->result = IDCANCEL;
      state->accepted = true;
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK ErrorDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<ErrorDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<ErrorDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, kAppTitle);
    state->font = DefaultUIFont();
    state->text = CreateWindowExW(0, L"STATIC", state->message.c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->ok_btn = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);

    ApplyConfirmFonts(hwnd, state->font);
    Theme::Current().ApplyToWindow(hwnd);
    Theme::Current().ApplyToChildren(hwnd);
    LayoutErrorDialog(hwnd, state);
    return 0;
  }
  case WM_SIZE:
    LayoutErrorDialog(hwnd, state);
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
  case WM_COMMAND:
    switch (LOWORD(wparam)) {
    case IDOK:
    case IDCANCEL:
      if (state) {
        state->accepted = true;
        appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      }
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  case WM_DESTROY:
    if (state && state->font) {
      DeleteObject(state->font);
      state->font = nullptr;
    }
    return 0;
  case WM_CLOSE:
    if (state) {
      state->accepted = true;
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK AboutDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<AboutDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<AboutDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, L"About RegKit");
    state->font = DefaultUIFont();
    state->credits = CreateWindowExW(0, L"STATIC", L"\x00A9 nohuto 2026", WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->repo_link = CreateWindowExW(0, WC_LINK,
                                       L"Repository: <a href=\"https://github.com/nohuto/regkit\">"
                                       L"https://github.com/nohuto/regkit</a>",
                                       WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->discord_link = CreateWindowExW(0, WC_LINK,
                                          L"Discord: <a href=\"https://discord.noverse.dev\">"
                                          L"https://discord.noverse.dev</a>",
                                          WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->website_link = CreateWindowExW(0, WC_LINK,
                                          L"Website: <a href=\"https://www.noverse.dev/\">"
                                          L"https://www.noverse.dev/</a>",
                                          WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->email_link = CreateWindowExW(0, WC_LINK,
                                        L"Email: <a href=\"mailto:contact@noverse.dev\">"
                                        L"contact@noverse.dev</a>",
                                        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->ok_btn = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);

    ApplyConfirmFonts(hwnd, state->font);
    Theme::Current().ApplyToWindow(hwnd);
    Theme::Current().ApplyToChildren(hwnd);
    LayoutAboutDialog(hwnd, state);
    return 0;
  }
  case WM_SHOWWINDOW:
    if (wparam && state && state->ok_btn) {
      SetFocus(state->ok_btn);
      return 0;
    }
    break;
  case WM_SETFOCUS:
    if (state && state->ok_btn) {
      SetFocus(state->ok_btn);
      return 0;
    }
    break;
  case WM_SIZE:
    LayoutAboutDialog(hwnd, state);
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
  case WM_CTLCOLORDLG:
  case WM_CTLCOLORSTATIC:
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    int type = CTLCOLOR_STATIC;
    if (msg == WM_CTLCOLORDLG) {
      type = CTLCOLOR_DLG;
    } else if (msg == WM_CTLCOLORBTN) {
      type = CTLCOLOR_BTN;
    }
    return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_NOTIFY: {
    auto* hdr = reinterpret_cast<NMHDR*>(lparam);
    if (hdr && (hdr->code == NM_CLICK || hdr->code == NM_RETURN)) {
      auto* link = reinterpret_cast<NMLINK*>(lparam);
      if (link && link->item.szUrl[0] != L'\0') {
        ShellExecuteW(hwnd, L"open", link->item.szUrl, nullptr, nullptr, SW_SHOWNORMAL);
        return 0;
      }
    }
    break;
  }
  case WM_COMMAND:
    switch (LOWORD(wparam)) {
    case IDOK:
    case IDCANCEL:
      if (state) {
        state->accepted = true;
        appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
      }
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  case WM_DESTROY:
    if (state && state->font) {
      DeleteObject(state->font);
      state->font = nullptr;
    }
    return 0;
  case WM_CLOSE:
    if (state) {
      state->accepted = true;
      appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

bool ShowErrorDialog(HWND owner, const std::wstring& message) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = ErrorDialogProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kErrorClass;
  RegisterClassW(&wc);

  ErrorDialogState state;
  state.owner = owner;
  state.message = message;
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kErrorClass, kAppTitle, WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 320, 120, owner, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    return false;
  }
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

  appearance::RestoreDialogOwner(owner, &state.owner_restored);
  return state.accepted;
}

bool ShowAboutDialog(HWND owner) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = AboutDialogProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kAboutClass;
  RegisterClassW(&wc);

  AboutDialogState state;
  state.owner = owner;
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kAboutClass, L"About RegKit", WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 460, 240, owner, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    return false;
  }
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

  appearance::RestoreDialogOwner(owner, &state.owner_restored);
  return state.accepted;
}

bool ShowChoiceDialog(HWND owner, const std::wstring& title, const std::wstring& message, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, int* result, PCWSTR icon_id, int width, int height, int button_width_dlu = 45, const std::wstring& detail = std::wstring(), int yes_button_width_dlu = 0, int default_id = 0) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = ChoiceDialogProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kChoiceClass;
  RegisterClassW(&wc);

  ChoiceDialogState state;
  state.owner = owner;
  state.title = title;
  state.message = message;
  state.detail = detail;
  state.yes_label = yes_label;
  state.no_label = no_label;
  state.cancel_label = cancel_label;
  state.icon_id = icon_id;
  state.button_width_dlu = button_width_dlu;
  state.yes_button_width_dlu = yes_button_width_dlu;
  state.default_id = default_id;
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kChoiceClass, kAppTitle, WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, width, height, owner, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    return false;
  }
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

  appearance::RestoreDialogOwner(owner, &state.owner_restored);
  if (!state.accepted) {
    return false;
  }
  if (result) {
    *result = state.result;
  }
  return true;
}

HRESULT CALLBACK TaskDialogCenterCallback(HWND hwnd, UINT msg, WPARAM, LPARAM, LONG_PTR ref_data) {
  if (msg == TDN_CREATED) {
    CenterWindowToOwner(hwnd, reinterpret_cast<HWND>(ref_data));
  }
  return S_OK;
}

bool ShowTaskDialog(HWND owner, const std::wstring& title, const std::wstring& message, TASKDIALOG_COMMON_BUTTON_FLAGS buttons, int* button, PCWSTR icon) {
  if (Theme::UseDarkMode()) {
    return false;
  }
  TASKDIALOGCONFIG config = {};
  config.cbSize = sizeof(config);
  config.hwndParent = owner;
  config.dwFlags = TDF_ALLOW_DIALOG_CANCELLATION;
  config.dwCommonButtons = buttons;
  config.pszWindowTitle = title.empty() ? kAppTitle : title.c_str();
  config.pszContent = message.c_str();
  config.pszMainIcon = icon;
  config.pfCallback = TaskDialogCenterCallback;
  config.lpCallbackData = reinterpret_cast<LONG_PTR>(owner);
  int clicked = 0;
  HRESULT hr = TaskDialogIndirect(&config, &clicked, nullptr, nullptr);
  if (FAILED(hr)) {
    return false;
  }
  if (button) {
    *button = clicked;
  }
  return true;
}

} // namespace

bool ListViewItemSelected(HWND list, int item_index) {
  return item_index >= 0 && (ListView_GetItemState(list, item_index, LVIS_SELECTED) & LVIS_SELECTED) != 0;
}

namespace {
void ApplyListViewThemeColors(NMLVCUSTOMDRAW* draw, const Theme& theme) {
  draw->clrText = theme.TextColor();
  draw->clrTextBk = theme.PanelColor();
}
} // namespace

LRESULT HandleThemedListViewCustomDraw(HWND list, NMLVCUSTOMDRAW* draw) {
  if (!list || !draw) {
    return CDRF_DODEFAULT;
  }
  switch (draw->nmcd.dwDrawStage) {
  case CDDS_PREPAINT:
    return CDRF_NOTIFYITEMDRAW;
  case CDDS_ITEMPREPAINT:
    draw->nmcd.uItemState &= ~(CDIS_FOCUS | CDIS_HOT);
    if (draw->nmcd.uItemState & CDIS_SELECTED) {
      return CDRF_DODEFAULT;
    }
    ApplyListViewThemeColors(draw, Theme::Current());
    return CDRF_DODEFAULT;
  default:
    break;
  }
  return CDRF_DODEFAULT;
}

bool CopyTextToClipboard(HWND owner, const std::wstring& text) {
  if (!OpenClipboard(owner)) {
    return false;
  }
  EmptyClipboard();
  size_t bytes = (text.size() + 1) * sizeof(wchar_t);
  HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, bytes);
  if (!memory) {
    CloseClipboard();
    return false;
  }
  void* data = GlobalLock(memory);
  if (!data) {
    GlobalFree(memory);
    CloseClipboard();
    return false;
  }
  memcpy(data, text.c_str(), bytes);
  GlobalUnlock(memory);
  SetClipboardData(CF_UNICODETEXT, memory);
  CloseClipboard();
  return true;
}

void ShowError(HWND owner, const std::wstring& message) {
  if (!ShowErrorDialog(owner, message)) {
    ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_ERROR_ICON);
  }
}

void ShowWarning(HWND owner, const std::wstring& message) {
  if (!ShowErrorDialog(owner, message)) {
    ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_WARNING_ICON);
  }
}

void ShowInfo(HWND owner, const std::wstring& message) {
  if (!ShowErrorDialog(owner, message)) {
    ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_INFORMATION_ICON);
  }
}

void ShowAbout(HWND owner) {
  if (ShowAboutDialog(owner)) {
    return;
  }
  ShowInfo(owner, L"\x00A9 nohuto 2026\n"
                  L"Repository: https://github.com/nohuto/regkit\n"
                  L"Discord: https://discord.noverse.dev\n"
                  L"Website: https://www.noverse.dev/\n"
                  L"Email: contact@noverse.dev");
}

bool ConfirmRegFileMerge(HWND owner, const std::wstring& path) {
  std::wstring message = L"Adding information can unintentionally change or delete values and\n"
                         L"cause components to stop working correctly. If you don't trust the\n"
                         L"source of this information in ";
  message += path;
  message += L",\ndon't add it to the registry.\n\n"
             L"Are you sure you want to continue?";
  int result = IDCANCEL;
  if (ShowChoiceDialog(owner, kAppTitle, message, L"Yes", L"No", L"", &result, IDI_WARNING, 560, 200, 36)) {
    return result == IDYES;
  }
  int clicked = 0;
  if (ShowTaskDialog(owner, kAppTitle, message, TDCBF_YES_BUTTON | TDCBF_NO_BUTTON, &clicked, TD_WARNING_ICON)) {
    return clicked == IDYES;
  }
  return false;
}

void ShowRegFileMergeSucceeded(HWND owner, const std::wstring& path) {
  std::wstring message = L"The keys and values contained in\n";
  message += path;
  message += L" have been successfully added to\nthe registry.";
  int result = IDCANCEL;
  if (ShowChoiceDialog(owner, kAppTitle, message, L"OK", L"", L"", &result, IDI_INFORMATION, 350, 150, 36)) {
    return;
  }
  if (!ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_INFORMATION_ICON)) {
    ShowInfo(owner, message);
  }
}

void ShowRegFileMergeFailed(HWND owner, const std::wstring& path, const std::wstring& detail) {
  std::wstring message = L"Can't import ";
  message += path;
  message += L".";
  if (!detail.empty()) {
    message += L"\n\n";
    message += detail;
  }
  int result = IDCANCEL;
  if (ShowChoiceDialog(owner, kAppTitle, message, L"OK", L"", L"", &result, IDI_ERROR, 520, 180, 36)) {
    return;
  }
  if (!ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_ERROR_ICON)) {
    ShowError(owner, message);
  }
}

constexpr int kKeyChoiceButtonWidthDlu = 45;

bool ConfirmDelete(HWND owner, const std::wstring& title,
                   const std::vector<std::wstring>& names) {
  if (names.empty()) {
    return false;
  }
  const bool many = names.size() > 1;
  std::wstring message;
  if (_wcsicmp(title.c_str(), L"Delete Key") == 0) {
    message = many ? L"Delete these keys and all of their subkeys?"
                   : L"Delete this key and all of its subkeys?";
  } else if (_wcsnicmp(title.c_str(), L"Delete Value", 12) == 0) {
    message = many ? L"Delete these values?" : L"Delete this value?";
  } else {
    message = many ? L"Delete these items?" : L"Delete this item?";
  }

  std::wstring detail;
  for (const std::wstring& name : names) {
    if (!detail.empty()) {
      detail.append(L"\r\n");
    }
    detail.append(name.empty() ? L"(Default)" : name);
  }

  const int lines = std::min(static_cast<int>(names.size()), kMaxDetailLines);
  int result = IDCANCEL;
  if (ShowChoiceDialog(owner, title, message, L"Delete", L"", L"Cancel", &result,
                       nullptr, 460, 128 + (lines - 1) * 16, kKeyChoiceButtonWidthDlu,
                       detail)) {
    return result == IDYES;
  }
  return false;
}

bool ConfirmDelete(HWND owner, const std::wstring& title, const std::wstring& name) {
  return ConfirmDelete(owner, title, std::vector<std::wstring>{name});
}


int PromptKeyChoice(HWND owner, const std::wstring& message, const std::wstring& key_path, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, int yes_button_width_dlu) {
  int result = IDCANCEL;
  if (ShowChoiceDialog(owner, title, message, yes_label, no_label, cancel_label,
                       &result, nullptr, 560, 140, kKeyChoiceButtonWidthDlu,
                       key_path, yes_button_width_dlu)) {
    return result;
  }
  return IDCANCEL;
}

int PromptChoice(HWND owner, const std::wstring& message, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, int button_width_dlu, int width) {
  int result = IDCANCEL;
  if (ShowChoiceDialog(owner, title, message, yes_label, no_label, cancel_label, &result, nullptr, width, 120, button_width_dlu)) {
    return result;
  }
  return IDCANCEL;
}

bool ReportFileDialogResult(HWND owner, HRESULT hr) {
  if (FAILED(hr) && !win32::DialogCancelled(hr)) {
    ShowError(owner, win32::FormatDialogError(hr));
  }
  return SUCCEEDED(hr);
}

bool LaunchNewInstance() {
  std::wstring exe = util::JoinPath(util::GetModuleDirectory(), L"RegKit.exe");
  if (exe.empty()) {
    return false;
  }
  HINSTANCE result = ShellExecuteW(nullptr, L"open", exe.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  return reinterpret_cast<intptr_t>(result) > 32;
}

} // namespace regkit::ui
