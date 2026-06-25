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

#include "../../include/app/font_dialog.h"

#include <algorithm>
#include <commdlg.h>
#include <commctrl.h>
#include <dlgs.h>
#include <uxtheme.h>

#include "../../include/app/theme.h"

namespace regkit {

namespace {

struct FontDialogHookState {
  bool dark_mode = false;
  HFONT sample_font = nullptr;
  LOGFONTW preview_base_font = {};
};

constexpr UINT_PTR kFontDialogSubclassId = 1;
constexpr UINT_PTR kFontDialogGroupBoxSubclassId = 3;
constexpr UINT_PTR kFontDialogSampleSubclassId = 4;
constexpr UINT kFontDialogUpdatePreviewMessage = WM_APP + 101;
constexpr int kFontDialogWidth = 447;
constexpr int kFontDialogHeight = 324;

LRESULT CALLBACK FontDialogGroupBoxSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR ref_data);
LRESULT CALLBACK FontDialogSampleSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR ref_data);

int FontPointSize(const LOGFONTW& font) {
  HDC hdc = GetDC(nullptr);
  int dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
  if (hdc) {
    ReleaseDC(nullptr, hdc);
  }
  if (dpi <= 0) {
    dpi = 96;
  }
  if (font.lfHeight == 0) {
    return 0;
  }
  return MulDiv(-font.lfHeight, 72, dpi);
}

int FontHeightFromPointSize(int point_size) {
  HDC hdc = GetDC(nullptr);
  int dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
  if (hdc) {
    ReleaseDC(nullptr, hdc);
  }
  if (dpi <= 0) {
    dpi = 96;
  }
  return -MulDiv(point_size, dpi, 72);
}

bool IsSameFontChoice(const LOGFONTW& left, const LOGFONTW& right) {
  return _wcsicmp(left.lfFaceName, right.lfFaceName) == 0 &&
         FontPointSize(left) == FontPointSize(right) &&
         left.lfWeight == right.lfWeight &&
         !!left.lfItalic == !!right.lfItalic;
}

void UpdateSamplePreview(HWND hwnd, FontDialogHookState* state) {
  if (!hwnd || !state) {
    return;
  }

  HWND face_combo = GetDlgItem(hwnd, cmb1);
  HWND style_combo = GetDlgItem(hwnd, cmb2);
  HWND size_combo = GetDlgItem(hwnd, cmb3);
  HWND sample_text = GetDlgItem(hwnd, stc5);
  if (!face_combo || !style_combo || !size_combo || !sample_text) {
    return;
  }

  LOGFONTW preview_font = state->preview_base_font;

  wchar_t face_name[LF_FACESIZE] = {};
  GetWindowTextW(face_combo, face_name, static_cast<int>(_countof(face_name)));
  if (face_name[0] != L'\0') {
    wcsncpy_s(preview_font.lfFaceName, face_name, _TRUNCATE);
  }

  preview_font.lfWeight = FW_NORMAL;
  preview_font.lfItalic = FALSE;
  wchar_t style_text[128] = {};
  GetWindowTextW(style_combo, style_text, static_cast<int>(_countof(style_text)));
  if (_wcsicmp(style_text, L"Bold") == 0 || wcsstr(style_text, L"Bold") != nullptr) {
    preview_font.lfWeight = FW_BOLD;
  }
  if (_wcsicmp(style_text, L"Italic") == 0 || _wcsicmp(style_text, L"Oblique") == 0 ||
      wcsstr(style_text, L"Italic") != nullptr || wcsstr(style_text, L"Oblique") != nullptr) {
    preview_font.lfItalic = TRUE;
  }

  wchar_t size_text[32] = {};
  GetWindowTextW(size_combo, size_text, static_cast<int>(_countof(size_text)));
  int point_size = _wtoi(size_text);
  if (point_size > 0) {
    preview_font.lfHeight = FontHeightFromPointSize(point_size);
  }

  HFONT next_font = CreateFontIndirectW(&preview_font);
  if (!next_font) {
    return;
  }

  HFONT previous_font = state->sample_font;
  state->sample_font = next_font;
  ShowWindow(sample_text, SW_SHOWNA);
  SendMessageW(sample_text, WM_SETFONT, reinterpret_cast<WPARAM>(next_font), TRUE);
  SetWindowTextW(sample_text, L"AaBbYyZz");
  InvalidateRect(sample_text, nullptr, TRUE);

  if (previous_font) {
    DeleteObject(previous_font);
  }
}

void ApplyNativeDarkTheme(HWND hwnd, bool dark_mode) {
  if (!hwnd) {
    return;
  }

  const wchar_t* default_theme = dark_mode ? L"DarkMode_Explorer" : L"Explorer";
  AllowDarkModeForWindow(hwnd, dark_mode);
  EnableImmersiveDarkMode(hwnd, dark_mode);
  SetWindowTheme(hwnd, default_theme, nullptr);

  EnumChildWindows(
      hwnd,
      [](HWND child, LPARAM param) -> BOOL {
        bool dark_mode = param != 0;
        int control_id = GetDlgCtrlID(child);
        if (control_id == stc5) {
          // Keep the sample preview controls on the stock theme so ChooseFontW
          // can keep driving the native preview surface.
          AllowDarkModeForWindow(child, FALSE);
          SetWindowTheme(child, nullptr, nullptr);
          return TRUE;
        }

        wchar_t class_name[32] = {};
        GetClassNameW(child, class_name, static_cast<int>(_countof(class_name)));

        const wchar_t* theme_name = dark_mode ? L"DarkMode_Explorer" : L"Explorer";
        if (wcscmp(class_name, WC_COMBOBOXW) == 0) {
          theme_name = dark_mode ? L"CFD" : L"Explorer";
          COMBOBOXINFO info = {sizeof(COMBOBOXINFO)};
          if (GetComboBoxInfo(child, &info) && info.hwndList) {
            AllowDarkModeForWindow(info.hwndList, dark_mode);
            SetWindowTheme(info.hwndList, theme_name, nullptr);
          }
        }

        AllowDarkModeForWindow(child, dark_mode);
        SetWindowTheme(child, theme_name, nullptr);
        return TRUE;
      },
      static_cast<LPARAM>(dark_mode ? 1 : 0));
}

void ApplyComboTheme(HWND combo, bool dark_mode) {
  if (!combo) {
    return;
  }

  LONG_PTR style = GetWindowLongPtrW(combo, GWL_STYLE);
  LONG_PTR combo_style = style & CBS_DROPDOWNLIST;
  bool is_simple = combo_style == CBS_SIMPLE;
  if (!is_simple) {
    Theme::Current().ApplyToComboBox(combo);
  }

  COMBOBOXINFO info = {sizeof(COMBOBOXINFO)};
  if (!GetComboBoxInfo(combo, &info)) {
    return;
  }

  const wchar_t* edit_theme = dark_mode ? L"DarkMode_CFD" : L"Explorer";
  const wchar_t* list_theme = dark_mode ? L"DarkMode_Explorer" : L"Explorer";
  if (info.hwndItem) {
    AllowDarkModeForWindow(info.hwndItem, dark_mode);
    SetWindowTheme(info.hwndItem, edit_theme, nullptr);
    InvalidateRect(info.hwndItem, nullptr, TRUE);
  }
  if (info.hwndList) {
    AllowDarkModeForWindow(info.hwndList, dark_mode);
    SetWindowTheme(info.hwndList, list_theme, nullptr);
    if (!is_simple) {
      LONG_PTR list_style = GetWindowLongPtrW(info.hwndList, GWL_STYLE);
      LONG_PTR list_ex_style = GetWindowLongPtrW(info.hwndList, GWL_EXSTYLE);
      if ((list_style & WS_BORDER) != 0) {
        SetWindowLongPtrW(info.hwndList, GWL_STYLE, list_style & ~WS_BORDER);
      }
      if ((list_ex_style & WS_EX_CLIENTEDGE) != 0) {
        SetWindowLongPtrW(info.hwndList, GWL_EXSTYLE, list_ex_style & ~WS_EX_CLIENTEDGE);
      }
      SetWindowPos(info.hwndList, nullptr, 0, 0, 0, 0,
                   SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    }
    RedrawWindow(info.hwndList, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW | RDW_FRAME);
  }
  InvalidateRect(combo, nullptr, TRUE);
}

void PaintFontDialogGroupBox(HWND hwnd, HDC hdc) {
  if (!hwnd || !hdc) {
    return;
  }

  RECT rc = {};
  GetClientRect(hwnd, &rc);
  FillRect(hdc, &rc, Theme::Current().BackgroundBrush());

  HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
  HGDIOBJ old_font = nullptr;
  if (font) {
    old_font = SelectObject(hdc, font);
  }

  wchar_t text[128] = {};
  GetWindowTextW(hwnd, text, static_cast<int>(_countof(text)));
  SIZE text_size = {};
  GetTextExtentPoint32W(hdc, text, static_cast<int>(wcslen(text)), &text_size);

  RECT text_rect = rc;
  text_rect.left += 8;
  text_rect.right = text_rect.left + text_size.cx + 6;
  text_rect.bottom = text_rect.top + text_size.cy;

  RECT frame_rect = rc;
  frame_rect.top += text_size.cy / 2;
  ExcludeClipRect(hdc, text_rect.left - 2, text_rect.top, text_rect.right + 2, text_rect.bottom);

  HPEN border_pen = CreatePen(PS_SOLID, 1, Theme::Current().BorderColor());
  HGDIOBJ old_pen = SelectObject(hdc, border_pen);
  HGDIOBJ old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
  Rectangle(hdc, frame_rect.left, frame_rect.top, frame_rect.right, frame_rect.bottom);
  SelectObject(hdc, old_brush);
  SelectObject(hdc, old_pen);
  DeleteObject(border_pen);
  SelectClipRgn(hdc, nullptr);

  SetBkMode(hdc, TRANSPARENT);
  SetTextColor(hdc, Theme::Current().TextColor());
  DrawTextW(hdc, text, -1, &text_rect, DT_SINGLELINE | DT_LEFT | DT_NOPREFIX);

  if (old_font) {
    SelectObject(hdc, old_font);
  }
}

LRESULT CALLBACK FontDialogGroupBoxSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, FontDialogGroupBoxSubclassProc, id);
    break;
  case WM_ERASEBKGND:
    return 1;
  case WM_PRINTCLIENT:
  case WM_PAINT: {
    PAINTSTRUCT ps = {};
    HDC hdc = (msg == WM_PAINT) ? BeginPaint(hwnd, &ps) : reinterpret_cast<HDC>(wparam);
    PaintFontDialogGroupBox(hwnd, hdc);
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

LRESULT CALLBACK FontDialogSampleSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, FontDialogSampleSubclassProc, id);
    break;
  case WM_ERASEBKGND:
    return 1;
  case WM_SETTEXT:
  case WM_SETFONT: {
    LRESULT result = DefSubclassProc(hwnd, msg, wparam, lparam);
    InvalidateRect(hwnd, nullptr, TRUE);
    return result;
  }
  case WM_PRINTCLIENT:
  case WM_PAINT: {
    PAINTSTRUCT ps = {};
    HDC hdc = (msg == WM_PAINT) ? BeginPaint(hwnd, &ps) : reinterpret_cast<HDC>(wparam);
    RECT rc = {};
    GetClientRect(hwnd, &rc);

    FillRect(hdc, &rc, GetSysColorBrush(COLOR_WINDOW));

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    HGDIOBJ old_font = nullptr;
    if (font) {
      old_font = SelectObject(hdc, font);
    }

    wchar_t text[128] = {};
    GetWindowTextW(hwnd, text, static_cast<int>(_countof(text)));
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));

    RECT text_rect = rc;
    DrawTextW(hdc, text, -1, &text_rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);

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

void ApplyFontDialogTheme(HWND hwnd, bool dark_mode) {
  if (!hwnd) {
    return;
  }

  ApplyNativeDarkTheme(hwnd, dark_mode);
  Theme::Current().ApplyToWindow(hwnd);

  constexpr int kComboIds[] = {cmb1, cmb2, cmb3, cmb5};
  for (int combo_id : kComboIds) {
    HWND combo = GetDlgItem(hwnd, combo_id);
    ApplyComboTheme(combo, dark_mode);
  }

  constexpr int kGroupIds[] = {grp1, grp2};
  for (int group_id : kGroupIds) {
    HWND group = GetDlgItem(hwnd, group_id);
    if (group && !GetWindowSubclass(group, FontDialogGroupBoxSubclassProc, kFontDialogGroupBoxSubclassId, nullptr)) {
      SetWindowSubclass(group, FontDialogGroupBoxSubclassProc, kFontDialogGroupBoxSubclassId, 0);
    }
  }

  HWND extra_bar = GetDlgItem(hwnd, stc6);
  if (extra_bar) {
    ShowWindow(extra_bar, SW_HIDE);
  }

  HWND sample_text = GetDlgItem(hwnd, stc5);
  if (sample_text && !GetWindowSubclass(sample_text, FontDialogSampleSubclassProc, kFontDialogSampleSubclassId, nullptr)) {
    SetWindowSubclass(sample_text, FontDialogSampleSubclassProc, kFontDialogSampleSubclassId, 0);
  }
}

LRESULT CALLBACK FontDialogSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR) {
  switch (msg) {
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, FontDialogSubclassProc, id);
    break;
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
    if (target == GetDlgItem(hwnd, stc5)) {
      break;
    }
    if (target == GetDlgItem(hwnd, grp2) || target == GetDlgItem(hwnd, grp1)) {
      SetTextColor(hdc, Theme::Current().TextColor());
      SetBkColor(hdc, Theme::Current().BackgroundColor());
      SetBkMode(hdc, TRANSPARENT);
      return reinterpret_cast<LRESULT>(Theme::Current().BackgroundBrush());
    }
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
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

UINT_PTR CALLBACK FontDialogHookProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  switch (msg) {
  case WM_INITDIALOG: {
    auto* choose = reinterpret_cast<CHOOSEFONTW*>(lparam);
    auto* state = choose ? reinterpret_cast<FontDialogHookState*>(choose->lCustData) : nullptr;
    bool dark_mode = state && state->dark_mode;
    SetWindowLongPtrW(hwnd, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    if (!GetWindowSubclass(hwnd, FontDialogSubclassProc, kFontDialogSubclassId, nullptr)) {
      SetWindowSubclass(hwnd, FontDialogSubclassProc, kFontDialogSubclassId, 0);
    }
    ApplyFontDialogTheme(hwnd, dark_mode);
    SetWindowPos(hwnd, nullptr, 0, 0, kFontDialogWidth, kFontDialogHeight,
                 SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    PostMessageW(hwnd, kFontDialogUpdatePreviewMessage, 0, reinterpret_cast<LPARAM>(state));
    return 0;
  }
  case kFontDialogUpdatePreviewMessage:
    UpdateSamplePreview(hwnd, reinterpret_cast<FontDialogHookState*>(lparam));
    return 0;
  case WM_COMMAND:
    if (HIWORD(wparam) == CBN_DROPDOWN || HIWORD(wparam) == CBN_SELCHANGE) {
      HWND combo = reinterpret_cast<HWND>(lparam);
      ApplyComboTheme(combo, Theme::UseDarkMode());
    }
    if (HIWORD(wparam) == CBN_SELCHANGE || HIWORD(wparam) == CBN_EDITCHANGE || HIWORD(wparam) == EN_CHANGE ||
        HIWORD(wparam) == LBN_SELCHANGE) {
      auto* dialog_state = reinterpret_cast<FontDialogHookState*>(GetWindowLongPtrW(hwnd, DWLP_USER));
      if (dialog_state) {
        PostMessageW(hwnd, kFontDialogUpdatePreviewMessage, 0, reinterpret_cast<LPARAM>(dialog_state));
      }
    }
    break;
  default:
    break;
  }
  return 0;
}

} // namespace

bool ShowFontDialog(HWND owner, const LOGFONTW& default_font, bool use_default, const LOGFONTW& current, FontDialogResult* out) {
  if (!out) {
    return false;
  }

  LOGFONTW chosen = use_default ? default_font : current;
  FontDialogHookState hook_state = {};
  hook_state.dark_mode = Theme::UseDarkMode();
  hook_state.preview_base_font = chosen;

  CHOOSEFONTW choose_font = {};
  choose_font.lStructSize = sizeof(choose_font);
  choose_font.hwndOwner = owner;
  choose_font.lpLogFont = &chosen;
  choose_font.Flags = CF_ENABLEHOOK | CF_FORCEFONTEXIST | CF_INITTOLOGFONTSTRUCT | CF_NOVERTFONTS | CF_SCREENFONTS;
  choose_font.lpfnHook = FontDialogHookProc;
  choose_font.lCustData = reinterpret_cast<LPARAM>(&hook_state);

  bool accepted = ChooseFontW(&choose_font) != FALSE;
  if (hook_state.sample_font) {
    DeleteObject(hook_state.sample_font);
    hook_state.sample_font = nullptr;
  }
  if (!accepted) {
    return false;
  }

  out->use_default = IsSameFontChoice(chosen, default_font);
  out->font = out->use_default ? default_font : chosen;
  return true;
}

} // namespace regkit
