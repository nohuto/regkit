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

#include "app/theme.h"

#include <commctrl.h>
#include <dwmapi.h>
#include <richedit.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <vssym32.h>
#include <winreg.h>

namespace regkit {

namespace {
#ifndef TB_SETBKCOLOR
#define TB_SETBKCOLOR (WM_USER + 66)
#endif
constexpr DWORD kUseImmersiveDarkModeBefore20H1 = 19;
constexpr DWORD kUseImmersiveDarkMode = 20;
ThemeMode g_theme_mode = ThemeMode::kSystem;
bool g_use_dark_mode = true;
bool g_custom_is_dark = true;
ThemeColors g_custom_colors = []() -> ThemeColors {
  ThemeColors colors;
  colors.background = RGB(20, 20, 20);
  colors.panel = RGB(20, 20, 20);
  colors.surface = RGB(34, 34, 34);
  colors.header = colors.surface;
  colors.border = RGB(66, 66, 66);
  colors.text = RGB(200, 200, 200);
  colors.muted_text = RGB(170, 170, 170);
  colors.accent = RGB(90, 162, 255);
  colors.selection = RGB(20, 20, 20);
  colors.selection_text = RGB(255, 255, 255);
  colors.hover = RGB(44, 44, 44);
  colors.focus = colors.accent;
  return colors;
}();
constexpr wchar_t kPersonalizePath[] = L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";

enum class PreferredAppMode {
  kDefault = 0,
  kAllowDark = 1,
  kForceDark = 2,
  kForceLight = 3,
  kMax = 4,
};

using SetPreferredAppModeFn = PreferredAppMode(WINAPI*)(PreferredAppMode);
using AllowDarkModeForWindowFn = BOOL(WINAPI*)(HWND, BOOL);
using RefreshImmersiveColorPolicyStateFn = void(WINAPI*)();
using FlushMenuThemesFn = void(WINAPI*)();

struct ComboBoxThemeState {
  bool hot = false;
};

constexpr UINT_PTR kGroupBoxSubclassId = 2;
constexpr UINT_PTR kStatusBarSubclassId = 3;
constexpr UINT_PTR kTreeViewSubclassId = 4;
constexpr UINT_PTR kEditShortcutSubclassId = 5;

HBRUSH GetCachedBrush(COLORREF color) {
  struct Entry {
    COLORREF color = CLR_INVALID;
    HBRUSH brush = nullptr;
  };
  static Entry cache[6];
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
  static Entry cache[6];
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

LRESULT DrawThemedButton(const NMCUSTOMDRAW* draw) {
  if (!draw || !draw->hdr.hwndFrom) {
    return CDRF_DODEFAULT;
  }

  LONG_PTR style = GetWindowLongPtrW(draw->hdr.hwndFrom, GWL_STYLE);
  int type = static_cast<int>(style & BS_TYPEMASK);
  bool is_checkbox = type == BS_AUTOCHECKBOX || type == BS_CHECKBOX || type == BS_AUTO3STATE || type == BS_3STATE;
  bool is_radio = type == BS_AUTORADIOBUTTON || type == BS_RADIOBUTTON;
  if (!is_checkbox && !is_radio) {
    return CDRF_DODEFAULT;
  }

  if (draw->dwDrawStage != CDDS_PREPAINT) {
    return CDRF_DODEFAULT;
  }

  HDC hdc = draw->hdc;
  RECT rect = draw->rc;

  HWND parent = GetParent(draw->hdr.hwndFrom);
  HBRUSH bg_brush = Theme::Current().BackgroundBrush();
  if (parent) {
    HBRUSH parent_brush = reinterpret_cast<HBRUSH>(SendMessageW(parent, WM_CTLCOLORBTN, reinterpret_cast<WPARAM>(hdc), reinterpret_cast<LPARAM>(draw->hdr.hwndFrom)));
    if (parent_brush) {
      bg_brush = parent_brush;
    }
  }
  FillRect(hdc, &rect, bg_brush);

  int part = is_radio ? BP_RADIOBUTTON : BP_CHECKBOX;
  bool disabled = (draw->uItemState & CDIS_DISABLED) != 0;
  bool pressed = (draw->uItemState & CDIS_SELECTED) != 0;
  bool hot = (draw->uItemState & CDIS_HOT) != 0;
  bool focused = (draw->uItemState & CDIS_FOCUS) != 0;
  bool show_cues = (draw->uItemState & CDIS_SHOWKEYBOARDCUES) != 0;

  LRESULT check_state = SendMessageW(draw->hdr.hwndFrom, BM_GETCHECK, 0, 0);
  bool checked = check_state == BST_CHECKED;
  bool mixed = check_state == BST_INDETERMINATE;

  int state = 0;
  if (is_checkbox) {
    if (disabled) {
      state = checked ? CBS_CHECKEDDISABLED : mixed ? CBS_MIXEDDISABLED : CBS_UNCHECKEDDISABLED;
    } else if (pressed) {
      state = checked ? CBS_CHECKEDPRESSED : mixed ? CBS_MIXEDPRESSED : CBS_UNCHECKEDPRESSED;
    } else if (hot) {
      state = checked ? CBS_CHECKEDHOT : mixed ? CBS_MIXEDHOT : CBS_UNCHECKEDHOT;
    } else {
      state = checked ? CBS_CHECKEDNORMAL : mixed ? CBS_MIXEDNORMAL : CBS_UNCHECKEDNORMAL;
    }
  } else {
    if (disabled) {
      state = checked ? RBS_CHECKEDDISABLED : RBS_UNCHECKEDDISABLED;
    } else if (pressed) {
      state = checked ? RBS_CHECKEDPRESSED : RBS_UNCHECKEDPRESSED;
    } else if (hot) {
      state = checked ? RBS_CHECKEDHOT : RBS_UNCHECKEDHOT;
    } else {
      state = checked ? RBS_CHECKEDNORMAL : RBS_UNCHECKEDNORMAL;
    }
  }

  HTHEME theme = OpenThemeData(draw->hdr.hwndFrom, L"Button");
  SIZE box_size = {};
  if (theme) {
    GetThemePartSize(theme, hdc, part, state, nullptr, TS_TRUE, &box_size);
  }
  if (box_size.cx <= 0 || box_size.cy <= 0) {
    box_size.cx = GetSystemMetrics(SM_CXMENUCHECK);
    box_size.cy = GetSystemMetrics(SM_CYMENUCHECK);
  }

  RECT box = rect;
  int padding = 4;
  bool left_text = (style & BS_LEFTTEXT) != 0;
  if (left_text) {
    box.right = rect.right - padding;
    box.left = box.right - box_size.cx;
  } else {
    box.left = rect.left + padding;
    box.right = box.left + box_size.cx;
  }
  box.top = rect.top + (rect.bottom - rect.top - box_size.cy) / 2;
  box.bottom = box.top + box_size.cy;

  if (theme) {
    DrawThemeBackground(theme, hdc, part, state, &box, nullptr);
  } else {
    DrawFrameControl(hdc, &box, DFC_BUTTON, is_radio ? DFCS_BUTTONRADIO : DFCS_BUTTONCHECK);
  }

  wchar_t text[256] = {};
  GetWindowTextW(draw->hdr.hwndFrom, text, static_cast<int>(_countof(text)));
  RECT text_rect = rect;
  if (left_text) {
    text_rect.right = box.left - padding;
    text_rect.left += padding;
  } else {
    text_rect.left = box.right + padding;
    text_rect.right -= padding;
  }

  COLORREF text_color = disabled ? Theme::Current().MutedTextColor() : Theme::Current().TextColor();
  UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
  if (!show_cues) {
    format |= DT_HIDEPREFIX;
  }

  if (theme) {
    DTTOPTS opts = {};
    opts.dwSize = sizeof(opts);
    opts.dwFlags = DTT_TEXTCOLOR;
    opts.crText = text_color;
    DrawThemeTextEx(theme, hdc, part, state, text, -1, format, &text_rect, &opts);
    CloseThemeData(theme);
  } else {
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, text_color);
    DrawTextW(hdc, text, -1, &text_rect, format);
  }

  if (focused && show_cues) {
    RECT focus_rect = text_rect;
    InflateRect(&focus_rect, 1, 1);
    DrawFocusRect(hdc, &focus_rect);
  }

  return CDRF_SKIPDEFAULT;
}

void PaintGroupBox(HWND hwnd, HDC hdc) {
  if (!hwnd || !hdc) {
    return;
  }

  const Theme& theme = Theme::Current();
  RECT rc = {};
  GetClientRect(hwnd, &rc);
  FillRect(hdc, &rc, theme.BackgroundBrush());

  HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
  HFONT old_font = nullptr;
  if (font) {
    old_font = reinterpret_cast<HFONT>(SelectObject(hdc, font));
  }

  wchar_t text[256] = {};
  GetWindowTextW(hwnd, text, static_cast<int>(_countof(text)));

  SIZE text_size = {};
  if (text[0] != L'\0') {
    GetTextExtentPoint32W(hdc, text, static_cast<int>(wcslen(text)), &text_size);
  } else {
    GetTextExtentPoint32W(hdc, L"M", 1, &text_size);
  }

  RECT rc_text = rc;
  RECT rc_frame = rc;
  rc_frame.top += text_size.cy / 2;

  if (text[0] != L'\0') {
    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    bool centered = (style & BS_CENTER) == BS_CENTER;
    int text_x = centered ? ((rc.right - rc.left - text_size.cx) / 2) : 7;
    rc_text.left += text_x;
    rc_text.right = rc_text.left + text_size.cx + 4;
    rc_text.bottom = rc_text.top + text_size.cy;

    ExcludeClipRect(hdc, rc_text.left, rc_text.top, rc_text.right, rc_text.bottom);
  }

  HPEN pen = GetCachedPen(theme.BorderColor(), 1);
  HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
  HBRUSH old_brush = reinterpret_cast<HBRUSH>(SelectObject(hdc, GetStockObject(NULL_BRUSH)));
  Rectangle(hdc, rc_frame.left, rc_frame.top, rc_frame.right, rc_frame.bottom);
  SelectObject(hdc, old_brush);
  SelectObject(hdc, old_pen);

  SelectClipRgn(hdc, nullptr);

  if (text[0] != L'\0') {
    COLORREF text_color = (GetWindowLongPtrW(hwnd, GWL_STYLE) & WS_DISABLED) ? theme.MutedTextColor() : theme.TextColor();
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, text_color);
    DrawTextW(hdc, text, -1, &rc_text, DT_SINGLELINE | DT_LEFT | DT_NOPREFIX);
  }

  if (old_font) {
    SelectObject(hdc, old_font);
  }
}

LRESULT CALLBACK GroupBoxSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, GroupBoxSubclassProc, id);
    break;
  case WM_ERASEBKGND:
    return 1;
  case WM_PRINTCLIENT:
  case WM_PAINT: {
    PAINTSTRUCT ps = {};
    HDC hdc = (msg == WM_PAINT) ? BeginPaint(hwnd, &ps) : reinterpret_cast<HDC>(wparam);
    PaintGroupBox(hwnd, hdc);
    if (msg == WM_PAINT) {
      EndPaint(hwnd, &ps);
    }
    return 0;
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK StatusBarSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, StatusBarSubclassProc, id);
    break;
  case WM_ERASEBKGND: {
    if (!Theme::UseDarkMode() && Theme::Mode() != ThemeMode::kCustom) {
      break;
    }
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rc = {};
    GetClientRect(hwnd, &rc);
    FillRect(hdc, &rc, Theme::Current().BackgroundBrush());
    return 1;
  }
  case WM_PRINTCLIENT:
  case WM_PAINT: {
    if (!Theme::UseDarkMode() && Theme::Mode() != ThemeMode::kCustom) {
      break;
    }
    PAINTSTRUCT ps = {};
    HDC hdc = (msg == WM_PAINT) ? BeginPaint(hwnd, &ps) : reinterpret_cast<HDC>(wparam);
    const Theme& theme = Theme::Current();
    COLORREF status_bg = theme.SurfaceColor();
    COLORREF status_text = theme.TextColor();
    COLORREF status_border = theme.BorderColor();
    RECT rc = {};
    GetClientRect(hwnd, &rc);
    HBRUSH bg_brush = GetCachedBrush(status_bg);
    FillRect(hdc, &rc, bg_brush);

    int borders[3] = {0, 0, 0};
    SendMessageW(hwnd, SB_GETBORDERS, 0, reinterpret_cast<LPARAM>(&borders));

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    HFONT old_font = nullptr;
    if (font) {
      old_font = reinterpret_cast<HFONT>(SelectObject(hdc, font));
    }
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, status_text);

    int parts = static_cast<int>(SendMessageW(hwnd, SB_GETPARTS, 0, 0));
    if (parts <= 0) {
      parts = 1;
    }
    for (int i = 0; i < parts; ++i) {
      RECT part = {};
      if (!SendMessageW(hwnd, SB_GETRECT, i, reinterpret_cast<LPARAM>(&part))) {
        part = rc;
      }
      if (i < parts - 1) {
        HPEN pen = GetCachedPen(status_border, 1);
        HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, pen));
        MoveToEx(hdc, part.right - 1, part.top + 2, nullptr);
        LineTo(hdc, part.right - 1, part.bottom - 2);
        SelectObject(hdc, old_pen);
      }
      wchar_t text[256] = {};
      LRESULT length = SendMessageW(hwnd, SB_GETTEXTLENGTH, i, 0);
      if (length > 0) {
        SendMessageW(hwnd, SB_GETTEXT, i, reinterpret_cast<LPARAM>(text));
      }
      RECT text_rect = part;
      text_rect.left += borders[2] + 4;
      text_rect.right -= borders[2] + 4;
      DrawTextW(hdc, text, -1, &text_rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);
    }

    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    if ((style & SBARS_SIZEGRIP) == SBARS_SIZEGRIP) {
      HTHEME status_theme = OpenThemeData(hwnd, VSCLASS_STATUS);
      if (status_theme) {
        SIZE grip = {};
        GetThemePartSize(status_theme, hdc, SP_GRIPPER, 0, &rc, TS_DRAW, &grip);
        RECT grip_rc = rc;
        grip_rc.left = grip_rc.right - grip.cx;
        grip_rc.top = grip_rc.bottom - grip.cy;
        DrawThemeBackground(status_theme, hdc, SP_GRIPPER, 0, &grip_rc, nullptr);
        CloseThemeData(status_theme);
      }
    }

    HPEN border_pen = GetCachedPen(status_border, 1);
    HPEN old_pen = reinterpret_cast<HPEN>(SelectObject(hdc, border_pen));
    MoveToEx(hdc, rc.left, rc.top, nullptr);
    LineTo(hdc, rc.right, rc.top);
    SelectObject(hdc, old_pen);

    if (old_font) {
      SelectObject(hdc, old_font);
    }

    if (msg == WM_PAINT) {
      EndPaint(hwnd, &ps);
    }
    return 0;
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK TreeViewSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, TreeViewSubclassProc, id);
    break;
  case WM_ERASEBKGND: {
    if (!Theme::UseDarkMode()) {
      break;
    }
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rc = {};
    GetClientRect(hwnd, &rc);
    FillRect(hdc, &rc, Theme::Current().PanelBrush());
    return 1;
  }
  case WM_PRINTCLIENT: {
    if (!Theme::UseDarkMode()) {
      break;
    }
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rc = {};
    GetClientRect(hwnd, &rc);
    FillRect(hdc, &rc, Theme::Current().PanelBrush());
    break;
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK EditShortcutSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, EditShortcutSubclassProc, id);
    break;
  case WM_KEYDOWN: {
    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
    if (!ctrl || alt) {
      break;
    }
    switch (wparam) {
    case 'A':
      SendMessageW(hwnd, EM_SETSEL, 0, -1);
      return 0;
    case 'C':
      SendMessageW(hwnd, WM_COPY, 0, 0);
      return 0;
    case 'V':
      SendMessageW(hwnd, WM_PASTE, 0, 0);
      return 0;
    case 'X':
      SendMessageW(hwnd, WM_CUT, 0, 0);
      return 0;
    case 'Z':
      SendMessageW(hwnd, EM_UNDO, 0, 0);
      return 0;
    case 'Y':
      SendMessageW(hwnd, EM_REDO, 0, 0);
      return 0;
    default:
      break;
    }
    break;
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

SetPreferredAppModeFn GetSetPreferredAppMode() {
  static SetPreferredAppModeFn fn = []() -> SetPreferredAppModeFn {
    HMODULE theme = GetModuleHandleW(L"uxtheme.dll");
    if (!theme) {
      return nullptr;
    }
    return reinterpret_cast<SetPreferredAppModeFn>(GetProcAddress(theme, MAKEINTRESOURCEA(135)));
  }();
  return fn;
}

AllowDarkModeForWindowFn GetAllowDarkModeForWindow() {
  static AllowDarkModeForWindowFn fn = []() -> AllowDarkModeForWindowFn {
    HMODULE theme = GetModuleHandleW(L"uxtheme.dll");
    if (!theme) {
      return nullptr;
    }
    return reinterpret_cast<AllowDarkModeForWindowFn>(GetProcAddress(theme, MAKEINTRESOURCEA(133)));
  }();
  return fn;
}

RefreshImmersiveColorPolicyStateFn GetRefreshImmersiveColorPolicyState() {
  static RefreshImmersiveColorPolicyStateFn fn = []() -> RefreshImmersiveColorPolicyStateFn {
    HMODULE theme = GetModuleHandleW(L"uxtheme.dll");
    if (!theme) {
      return nullptr;
    }
    return reinterpret_cast<RefreshImmersiveColorPolicyStateFn>(GetProcAddress(theme, MAKEINTRESOURCEA(104)));
  }();
  return fn;
}

FlushMenuThemesFn GetFlushMenuThemes() {
  static FlushMenuThemesFn fn = []() -> FlushMenuThemesFn {
    HMODULE theme = GetModuleHandleW(L"uxtheme.dll");
    if (!theme) {
      return nullptr;
    }
    return reinterpret_cast<FlushMenuThemesFn>(GetProcAddress(theme, MAKEINTRESOURCEA(136)));
  }();
  return fn;
}

void ConfigureDarkModeSupport(PreferredAppMode mode) {
  static bool configured = false;
  static PreferredAppMode configured_mode = PreferredAppMode::kDefault;
  if (configured && configured_mode == mode) {
    return;
  }
  if (auto set_mode = GetSetPreferredAppMode()) {
    set_mode(mode);
  }
  if (auto refresh = GetRefreshImmersiveColorPolicyState()) {
    refresh();
  }
  if (auto flush = GetFlushMenuThemes()) {
    flush();
  }
  configured = true;
  configured_mode = mode;
}

bool ReadSystemDarkMode() {
  DWORD value = 1;
  DWORD size = sizeof(value);
  if (RegGetValueW(HKEY_CURRENT_USER, kPersonalizePath, L"AppsUseLightTheme", RRF_RT_REG_DWORD, nullptr, &value, &size) == ERROR_SUCCESS) {
    return value == 0;
  }
  return false;
}

void PaintComboBox(HWND hwnd, HDC hdc, ComboBoxThemeState* state) {
  RECT rect = {};
  GetClientRect(hwnd, &rect);
  int width = rect.right - rect.left;
  int height = rect.bottom - rect.top;
  if (width <= 0 || height <= 0) {
    return;
  }

  const Theme& theme = Theme::Current();
  bool enabled = IsWindowEnabled(hwnd) != FALSE;
  bool hot = state && state->hot && enabled;

  HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
  HGDIOBJ old_font = nullptr;
  if (font) {
    old_font = SelectObject(hdc, font);
  }
  SetBkMode(hdc, TRANSPARENT);

  COMBOBOXINFO info = {sizeof(COMBOBOXINFO)};
  RECT button_rect = rect;
  if (GetComboBoxInfo(hwnd, &info)) {
    button_rect = info.rcButton;
  } else {
    int btn_w = GetSystemMetrics(SM_CXVSCROLL);
    button_rect.left = rect.right - btn_w;
  }
  if (button_rect.left > rect.left) {
    button_rect.left -= 1;
  }

  LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
  LONG_PTR cb_style = style & CBS_DROPDOWNLIST;
  bool has_focus = (cb_style == CBS_DROPDOWNLIST && GetFocus() == hwnd) || (cb_style == CBS_DROPDOWN && info.hwndItem && GetFocus() == info.hwndItem);

  COLORREF border = theme.BorderColor();
  if (!enabled) {
    border = theme.BorderColor();
  } else if (has_focus) {
    border = theme.FocusColor();
  } else if (hot) {
    border = theme.HoverColor();
  }

  COLORREF text = enabled ? theme.TextColor() : theme.MutedTextColor();
  COLORREF fill = enabled ? (hot ? theme.HoverColor() : theme.SurfaceColor()) : theme.SurfaceColor();

  HTHEME combo_theme = OpenThemeData(hwnd, VSCLASS_COMBOBOX);
  bool has_theme = combo_theme != nullptr;

  if (cb_style == CBS_DROPDOWNLIST) {
    HBRUSH fill_brush = GetCachedBrush(fill);
    FillRect(hdc, &rect, fill_brush);
  } else if (cb_style == CBS_DROPDOWN) {
    HBRUSH fill_brush = GetCachedBrush(fill);
    FillRect(hdc, &button_rect, fill_brush);
  }

  if (cb_style != CBS_SIMPLE) {
    if (has_theme) {
      RECT themed_arrow = button_rect;
      themed_arrow.top -= 1;
      themed_arrow.bottom -= 1;
      DrawThemeBackground(combo_theme, hdc, CP_DROPDOWNBUTTONRIGHT, enabled ? CBXSR_NORMAL : CBXSR_DISABLED, &themed_arrow, nullptr);
    } else {
      COLORREF arrow_color = enabled ? text : theme.MutedTextColor();
      HPEN arrow_pen = GetCachedPen(arrow_color, 1);
      HBRUSH arrow_brush = GetCachedBrush(arrow_color);
      HPEN old_arrow_pen = reinterpret_cast<HPEN>(SelectObject(hdc, arrow_pen));
      HBRUSH old_arrow_brush = reinterpret_cast<HBRUSH>(SelectObject(hdc, arrow_brush));
      int mid_x = (button_rect.left + button_rect.right) / 2;
      int mid_y = (button_rect.top + button_rect.bottom) / 2;
      POINT arrow[3] = {
          {mid_x - 4, mid_y - 2},
          {mid_x + 4, mid_y - 2},
          {mid_x, mid_y + 3},
      };
      Polygon(hdc, arrow, 3);
      SelectObject(hdc, old_arrow_brush);
      SelectObject(hdc, old_arrow_pen);
    }
  }

  if (cb_style == CBS_DROPDOWNLIST) {
    wchar_t buffer[256] = {};
    int index = static_cast<int>(SendMessageW(hwnd, CB_GETCURSEL, 0, 0));
    if (index != CB_ERR) {
      SendMessageW(hwnd, CB_GETLBTEXT, index, reinterpret_cast<LPARAM>(buffer));
    } else {
      GetWindowTextW(hwnd, buffer, static_cast<int>(_countof(buffer)));
    }
    RECT text_rect = info.rcItem;
    InflateRect(&text_rect, -2, 0);
    if (has_theme) {
      DTTOPTS opts = {};
      opts.dwSize = sizeof(opts);
      opts.dwFlags = DTT_TEXTCOLOR;
      opts.crText = text;
      DrawThemeTextEx(combo_theme, hdc, CP_DROPDOWNITEM, enabled ? CBXSR_NORMAL : CBXSR_DISABLED, buffer, -1, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX, &text_rect, &opts);
    } else {
      SetTextColor(hdc, text);
      DrawTextW(hdc, buffer, -1, &text_rect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
    }
    if (enabled && has_focus && SendMessageW(hwnd, CB_GETDROPPEDSTATE, 0, 0) == FALSE) {
      DrawFocusRect(hdc, &info.rcItem);
    }
  } else if (cb_style == CBS_DROPDOWN) {
    POINT edge[] = {
        {button_rect.left - 1, button_rect.top},
        {button_rect.left - 1, button_rect.bottom},
    };
    HPEN edge_pen = GetCachedPen(border, 1);
    HGDIOBJ old_pen = SelectObject(hdc, edge_pen);
    Polyline(hdc, edge, _countof(edge));
    SelectObject(hdc, old_pen);
  }

  HPEN border_pen = GetCachedPen(border, 1);
  HGDIOBJ old_pen = SelectObject(hdc, border_pen);
  HGDIOBJ old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
  Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
  SelectObject(hdc, old_brush);
  SelectObject(hdc, old_pen);

  if (combo_theme) {
    CloseThemeData(combo_theme);
  }
  if (old_font) {
    SelectObject(hdc, old_font);
  }
}

LRESULT CALLBACK ComboBoxThemeSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR ref_data) {
  auto* state = reinterpret_cast<ComboBoxThemeState*>(ref_data);
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, ComboBoxThemeSubclassProc, id);
    delete state;
    break;
  case WM_MOUSEMOVE: {
    if (state && !state->hot) {
      state->hot = true;
      TRACKMOUSEEVENT tme = {};
      tme.cbSize = sizeof(tme);
      tme.dwFlags = TME_LEAVE;
      tme.hwndTrack = hwnd;
      TrackMouseEvent(&tme);
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    break;
  }
  case WM_MOUSELEAVE:
    if (state && state->hot) {
      state->hot = false;
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    break;
  case WM_SETFOCUS:
  case WM_KILLFOCUS:
  case WM_ENABLE:
    InvalidateRect(hwnd, nullptr, FALSE);
    break;
  case WM_ERASEBKGND:
    return 1;
  case WM_PAINT: {
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd, &ps);
    LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
    LONG_PTR cb_style = style & CBS_DROPDOWNLIST;
    if (cb_style != CBS_DROPDOWN) {
      RECT rect = {};
      GetClientRect(hwnd, &rect);
      int width = rect.right - rect.left;
      int height = rect.bottom - rect.top;
      if (width > 0 && height > 0) {
        HDC mem_dc = CreateCompatibleDC(hdc);
        HBITMAP mem_bmp = CreateCompatibleBitmap(hdc, width, height);
        HGDIOBJ old_bmp = SelectObject(mem_dc, mem_bmp);
        PaintComboBox(hwnd, mem_dc, state);
        BitBlt(hdc, 0, 0, width, height, mem_dc, 0, 0, SRCCOPY);
        SelectObject(mem_dc, old_bmp);
        DeleteObject(mem_bmp);
        DeleteDC(mem_dc);
      }
    } else {
      PaintComboBox(hwnd, hdc, state);
    }
    EndPaint(hwnd, &ps);
    return 0;
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK ThemeWindowSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, ThemeWindowSubclassProc, id);
    break;
  case WM_NOTIFY: {
    auto* hdr = reinterpret_cast<NMHDR*>(lparam);
    if (hdr && hdr->code == NM_CUSTOMDRAW) {
      auto* draw = reinterpret_cast<NMCUSTOMDRAW*>(lparam);
      wchar_t class_name[32] = {};
      if (GetClassNameW(hdr->hwndFrom, class_name, static_cast<int>(_countof(class_name))) && wcscmp(class_name, WC_BUTTON) == 0) {
        LRESULT result = DrawThemedButton(draw);
        if (result != CDRF_DODEFAULT) {
          return result;
        }
      }
    }
    break;
  }
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}
} // namespace

Theme& Theme::Dark() {
  ThemeColors colors;
  colors.background = RGB(20, 20, 20);
  colors.panel = RGB(20, 20, 20);
  colors.surface = RGB(34, 34, 34);
  colors.header = colors.surface;
  colors.border = RGB(66, 66, 66);
  colors.text = RGB(200, 200, 200);
  colors.muted_text = RGB(170, 170, 170);
  colors.accent = RGB(90, 162, 255);
  colors.selection = RGB(20, 20, 20);
  colors.selection_text = RGB(255, 255, 255);
  colors.hover = RGB(44, 44, 44);
  colors.focus = colors.accent;
  static Theme theme(colors, true);
  return theme;
}

Theme& Theme::Light() {
  ThemeColors colors;
  colors.background = RGB(245, 245, 245);
  colors.panel = RGB(255, 255, 255);
  colors.surface = RGB(242, 242, 242);
  colors.header = colors.surface;
  colors.border = RGB(204, 204, 204);
  colors.text = RGB(32, 32, 32);
  colors.muted_text = RGB(96, 96, 96);
  colors.accent = RGB(0, 120, 215);
  colors.selection = RGB(229, 241, 255);
  colors.selection_text = RGB(32, 32, 32);
  colors.hover = RGB(236, 236, 236);
  colors.focus = colors.accent;
  static Theme theme(colors, false);
  return theme;
}

Theme& Theme::Custom() {
  static Theme theme(g_custom_colors, g_custom_is_dark);
  return theme;
}

Theme& Theme::Current() {
  if (g_theme_mode == ThemeMode::kCustom) {
    return Custom();
  }
  return g_use_dark_mode ? Dark() : Light();
}

void Theme::SetCustomColors(const ThemeColors& colors, bool is_dark) {
  g_custom_colors = colors;
  g_custom_is_dark = is_dark;
  Theme& custom = Custom();
  custom.colors_ = colors;
  custom.is_dark_ = is_dark;
  custom.background_brush_.reset(CreateSolidBrush(colors.background));
  custom.panel_brush_.reset(CreateSolidBrush(colors.panel));
  custom.surface_brush_.reset(CreateSolidBrush(colors.surface));
  custom.header_brush_.reset(CreateSolidBrush(colors.header));
  if (g_theme_mode == ThemeMode::kCustom) {
    g_use_dark_mode = is_dark;
    InitializeDarkModeSupport();
  }
}

void Theme::SetMode(ThemeMode mode) {
  g_theme_mode = mode;
  if (mode == ThemeMode::kCustom) {
    g_use_dark_mode = g_custom_is_dark;
  } else if (mode == ThemeMode::kSystem) {
    g_use_dark_mode = ReadSystemDarkMode();
  } else {
    g_use_dark_mode = (mode == ThemeMode::kDark);
  }
  InitializeDarkModeSupport();
}

ThemeMode Theme::Mode() {
  return g_theme_mode;
}

bool Theme::UseDarkMode() {
  return g_use_dark_mode;
}

bool Theme::IsSystemDarkMode() {
  return ReadSystemDarkMode();
}

bool Theme::UpdateFromSystem() {
  if (g_theme_mode != ThemeMode::kSystem) {
    return false;
  }
  bool use_dark = ReadSystemDarkMode();
  if (use_dark == g_use_dark_mode) {
    return false;
  }
  g_use_dark_mode = use_dark;
  InitializeDarkModeSupport();
  return true;
}

void Theme::InitializeDarkModeSupport() {
  if (g_theme_mode == ThemeMode::kSystem) {
    g_use_dark_mode = ReadSystemDarkMode();
  } else if (g_theme_mode == ThemeMode::kCustom) {
    g_use_dark_mode = g_custom_is_dark;
  }
  ConfigureDarkModeSupport(g_use_dark_mode ? PreferredAppMode::kForceDark : PreferredAppMode::kForceLight);
}

Theme::Theme(const ThemeColors& colors, bool is_dark) : colors_(colors), is_dark_(is_dark) {
  background_brush_.reset(CreateSolidBrush(colors_.background));
  panel_brush_.reset(CreateSolidBrush(colors_.panel));
  surface_brush_.reset(CreateSolidBrush(colors_.surface));
  header_brush_.reset(CreateSolidBrush(colors_.header));
}

void Theme::ApplyToWindow(HWND hwnd) const {
  AllowDarkModeForWindow(hwnd, is_dark_);
  EnableImmersiveDarkMode(hwnd, is_dark_);
  const wchar_t* theme = is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
  SetWindowTheme(hwnd, theme, nullptr);
  if (!GetWindowSubclass(hwnd, ThemeWindowSubclassProc, 1, nullptr)) {
    SetWindowSubclass(hwnd, ThemeWindowSubclassProc, 1, 0);
  }
}

void Theme::ApplyToChildren(HWND hwnd) const {
  if (!hwnd) {
    return;
  }
  EnumChildWindows(
      hwnd,
      [](HWND child, LPARAM param) -> BOOL {
        auto* theme = reinterpret_cast<const Theme*>(param);
        if (!theme) {
          return TRUE;
        }
        AllowDarkModeForWindow(child, theme->is_dark_);
        const wchar_t* theme_name = theme->is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
        SetWindowTheme(child, theme_name, nullptr);

        wchar_t class_name[64] = {};
        if (GetClassNameW(child, class_name, static_cast<int>(_countof(class_name)))) {
          if (wcscmp(class_name, WC_COMBOBOXW) == 0) {
            theme->ApplyToComboBox(child);
          } else if (_wcsicmp(class_name, WC_EDITW) == 0 || _wcsnicmp(class_name, L"RICHEDIT", 8) == 0) {
            if (!GetWindowSubclass(child, EditShortcutSubclassProc, kEditShortcutSubclassId, nullptr)) {
              SetWindowSubclass(child, EditShortcutSubclassProc, kEditShortcutSubclassId, 0);
            }
          } else if (wcscmp(class_name, WC_BUTTONW) == 0) {
            LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
            if ((style & BS_GROUPBOX) == BS_GROUPBOX) {
              if (!GetWindowSubclass(child, GroupBoxSubclassProc, kGroupBoxSubclassId, nullptr)) {
                SetWindowSubclass(child, GroupBoxSubclassProc, kGroupBoxSubclassId, 0);
              }
            }
          }
        }
        return TRUE;
      },
      reinterpret_cast<LPARAM>(this));
}

void Theme::ApplyToTreeView(HWND hwnd) const {
  if (!hwnd) {
    return;
  }
  AllowDarkModeForWindow(hwnd, is_dark_);
  const wchar_t* theme = is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
  SetWindowTheme(hwnd, theme, nullptr);
  TreeView_SetBkColor(hwnd, colors_.panel);
  TreeView_SetTextColor(hwnd, colors_.text);
  TreeView_SetLineColor(hwnd, colors_.border);
  if (!GetWindowSubclass(hwnd, TreeViewSubclassProc, kTreeViewSubclassId, nullptr)) {
    SetWindowSubclass(hwnd, TreeViewSubclassProc, kTreeViewSubclassId, 0);
  }
}

void Theme::ApplyToListView(HWND hwnd) const {
  if (!hwnd) {
    return;
  }
  AllowDarkModeForWindow(hwnd, is_dark_);
  const wchar_t* theme = is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
  SetWindowTheme(hwnd, theme, nullptr);
  ListView_SetBkColor(hwnd, colors_.panel);
  ListView_SetTextBkColor(hwnd, colors_.panel);
  ListView_SetTextColor(hwnd, colors_.text);

  HWND tooltip = ListView_GetToolTips(hwnd);
  if (tooltip) {
    AllowDarkModeForWindow(tooltip, is_dark_);
    const wchar_t* tip_theme = is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
    SetWindowTheme(tooltip, tip_theme, nullptr);
  }

  HWND header = ListView_GetHeader(hwnd);
  if (header) {
    AllowDarkModeForWindow(header, is_dark_);
    SetWindowTheme(header, theme, nullptr);
    InvalidateRect(header, nullptr, TRUE);
  }
}

void Theme::ApplyToTabControl(HWND hwnd) const {
  if (!hwnd) {
    return;
  }
  AllowDarkModeForWindow(hwnd, is_dark_);
  const wchar_t* theme_name = is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
  SetWindowTheme(hwnd, theme_name, nullptr);
  InvalidateRect(hwnd, nullptr, TRUE);
}

void Theme::ApplyToToolbar(HWND hwnd) const {
  if (!hwnd) {
    return;
  }
  AllowDarkModeForWindow(hwnd, is_dark_);
  const wchar_t* theme_name = is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
  SetWindowTheme(hwnd, theme_name, nullptr);
}

void Theme::ApplyToComboBox(HWND hwnd) const {
  if (!hwnd) {
    return;
  }
  if (!GetWindowSubclass(hwnd, ComboBoxThemeSubclassProc, 1, nullptr)) {
    auto* state = new ComboBoxThemeState();
    SetWindowSubclass(hwnd, ComboBoxThemeSubclassProc, 1, reinterpret_cast<DWORD_PTR>(state));
  }
  LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
  COMBOBOXINFO info = {sizeof(COMBOBOXINFO)};
  if (GetComboBoxInfo(hwnd, &info) && info.hwndList) {
    const wchar_t* theme_name = is_dark_ ? L"CFD" : L"Explorer";
    SetWindowTheme(info.hwndList, theme_name, nullptr);
  }
  const wchar_t* theme_name = is_dark_ ? L"CFD" : L"Explorer";
  SetWindowTheme(hwnd, theme_name, nullptr);
  InvalidateRect(hwnd, nullptr, TRUE);
}

void Theme::ApplyToStatusBar(HWND hwnd) const {
  if (!hwnd) {
    return;
  }
  AllowDarkModeForWindow(hwnd, is_dark_);
  const wchar_t* theme_name = is_dark_ ? L"DarkMode_Explorer" : L"Explorer";
  SetWindowTheme(hwnd, theme_name, nullptr);
  if (!GetWindowSubclass(hwnd, StatusBarSubclassProc, kStatusBarSubclassId, nullptr)) {
    SetWindowSubclass(hwnd, StatusBarSubclassProc, kStatusBarSubclassId, 0);
  }
  InvalidateRect(hwnd, nullptr, TRUE);
}

HBRUSH Theme::BackgroundBrush() const {
  return background_brush_.get();
}

HBRUSH Theme::PanelBrush() const {
  return panel_brush_.get();
}

HBRUSH Theme::SurfaceBrush() const {
  return surface_brush_.get();
}

HBRUSH Theme::HeaderBrush() const {
  return header_brush_.get();
}

COLORREF Theme::BackgroundColor() const {
  return colors_.background;
}

COLORREF Theme::PanelColor() const {
  return colors_.panel;
}

COLORREF Theme::SurfaceColor() const {
  return colors_.surface;
}

COLORREF Theme::HeaderColor() const {
  return colors_.header;
}

COLORREF Theme::BorderColor() const {
  return colors_.border;
}

COLORREF Theme::TextColor() const {
  return colors_.text;
}

COLORREF Theme::MutedTextColor() const {
  return colors_.muted_text;
}

COLORREF Theme::SelectionColor() const {
  return colors_.selection;
}

COLORREF Theme::SelectionTextColor() const {
  return colors_.selection_text;
}

COLORREF Theme::HoverColor() const {
  return colors_.hover;
}

COLORREF Theme::FocusColor() const {
  return colors_.focus;
}

HBRUSH Theme::ControlColor(HDC hdc, HWND target, int type) const {
  if (!hdc) {
    return nullptr;
  }
  switch (type) {
  case CTLCOLOR_EDIT:
    SetTextColor(hdc, TextColor());
    SetBkColor(hdc, SurfaceColor());
    return SurfaceBrush();
  case CTLCOLOR_LISTBOX:
    SetTextColor(hdc, TextColor());
    SetBkColor(hdc, SurfaceColor());
    return SurfaceBrush();
  case CTLCOLOR_BTN: {
    COLORREF text = TextColor();
    if (target && !IsWindowEnabled(target)) {
      text = MutedTextColor();
    }
    SetTextColor(hdc, text);
    SetBkColor(hdc, SurfaceColor());
    SetBkMode(hdc, TRANSPARENT);
    return BackgroundBrush();
  }
  case CTLCOLOR_DLG:
  case CTLCOLOR_STATIC:
  default:
    SetTextColor(hdc, TextColor());
    SetBkColor(hdc, BackgroundColor());
    SetBkMode(hdc, TRANSPARENT);
    return BackgroundBrush();
  }
}

void EnableImmersiveDarkMode(HWND hwnd, bool enabled) {
  if (!hwnd) {
    return;
  }
  BOOL value = enabled ? TRUE : FALSE;
  HRESULT result = DwmSetWindowAttribute(hwnd, kUseImmersiveDarkMode, &value, sizeof(value));
  if (FAILED(result)) {
    DwmSetWindowAttribute(hwnd, kUseImmersiveDarkModeBefore20H1, &value, sizeof(value));
  }
}

void AllowDarkModeForWindow(HWND hwnd, bool enabled) {
  if (!hwnd) {
    return;
  }
  ConfigureDarkModeSupport(enabled ? PreferredAppMode::kForceDark : PreferredAppMode::kForceLight);
  if (auto allow = GetAllowDarkModeForWindow()) {
    allow(hwnd, enabled ? TRUE : FALSE);
  }
}

} // namespace regkit
