// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/dialog_support.h"
#include "win32/text_transform.h"

#include "appearance/theme.h"
#include "appearance/default_font.h"
#include "appearance/feedback.h"

#include <algorithm>
#include <string>

#include <commctrl.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <windowsx.h>

namespace regkit::editors::dialog_support {

namespace {

constexpr UINT_PTR kMultilineSubclassId = 1;

LRESULT CALLBACK MultilineProc(HWND window, UINT message, WPARAM wparam,
                               LPARAM lparam, UINT_PTR, DWORD_PTR) {
  if (message == WM_GETDLGCODE && wparam == VK_RETURN) {
    return DefSubclassProc(window, message, wparam, lparam) | DLGC_WANTMESSAGE;
  }
  if (message == WM_NCDESTROY) {
    RemoveWindowSubclass(window, MultilineProc, kMultilineSubclassId);
  }
  return DefSubclassProc(window, message, wparam, lparam);
}

constexpr UINT_PTR kSingleLineSubclassId = 2;

LRESULT CALLBACK SingleLineProc(HWND window, UINT message, WPARAM wparam,
                                LPARAM lparam, UINT_PTR, DWORD_PTR) {
  if (message == WM_CHAR && (wparam == L'\r' || wparam == L'\n')) {
    return 0;
  }
  if (message == WM_PASTE) {
    std::wstring text;
    if (OpenClipboard(window)) {
      HANDLE handle = GetClipboardData(CF_UNICODETEXT);
      const wchar_t* data = handle ? static_cast<const wchar_t*>(GlobalLock(handle)) : nullptr;
      if (data) {
        text = data;
        GlobalUnlock(handle);
      }
      CloseClipboard();
    }
    if (text.find_first_of(L"\r\n") == std::wstring::npos) {
      return DefSubclassProc(window, message, wparam, lparam);
    }
    for (wchar_t& character : text) {
      if (character == L'\r' || character == L'\n') {
        character = L' ';
      }
    }
    SendMessageW(window, EM_REPLACESEL, TRUE,
                 reinterpret_cast<LPARAM>(text.c_str()));
    return 0;
  }
  if (message == WM_NCDESTROY) {
    RemoveWindowSubclass(window, SingleLineProc, kSingleLineSubclassId);
  }
  return DefSubclassProc(window, message, wparam, lparam);
}

bool SizeGripRect(HWND dialog, RECT* rect) {
  if (!rect || (GetWindowLongPtrW(dialog, GWL_STYLE) & WS_THICKFRAME) == 0) {
    return false;
  }
  RECT client = {};
  if (!GetClientRect(dialog, &client)) {
    return false;
  }
  SIZE grip = {};
  HTHEME theme = OpenThemeData(dialog, VSCLASS_STATUS);
  if (theme) {
    if (FAILED(GetThemePartSize(theme, nullptr, SP_GRIPPER, 0, nullptr, TS_TRUE, &grip))) {
      grip = {};
    }
    CloseThemeData(theme);
  }
  if (grip.cx <= 0 || grip.cy <= 0) {
    grip.cx = grip.cy = GetSystemMetrics(SM_CXVSCROLL);
  }
  rect->left = client.right - grip.cx;
  rect->top = client.bottom - grip.cy;
  rect->right = client.right;
  rect->bottom = client.bottom;
  return true;
}

void DrawSizeGrip(HWND dialog, HDC hdc) {
  RECT grip = {};
  if (!hdc || !SizeGripRect(dialog, &grip)) {
    return;
  }
  HTHEME theme = OpenThemeData(dialog, VSCLASS_STATUS);
  if (theme) {
    DrawThemeBackground(theme, hdc, SP_GRIPPER, 0, &grip, nullptr);
    CloseThemeData(theme);
    return;
  }
  DrawFrameControl(hdc, &grip, DFC_SCROLL, DFCS_SCROLLSIZEGRIP);
}

void Center(HWND dialog) {
  RECT dialog_rect = {};
  if (!GetWindowRect(dialog, &dialog_rect)) {
    return;
  }
  RECT target = {};
  const HWND owner = GetWindow(dialog, GW_OWNER);
  if ((!owner || !GetWindowRect(owner, &target)) &&
      !SystemParametersInfoW(SPI_GETWORKAREA, 0, &target, 0)) {
    return;
  }
  const int width = dialog_rect.right - dialog_rect.left;
  const int height = dialog_rect.bottom - dialog_rect.top;
  const LONG x = target.left +
                 std::max<LONG>(0, (target.right - target.left - width) / 2);
  const LONG y = target.top +
                 std::max<LONG>(0, (target.bottom - target.top - height) / 2);
  SetWindowPos(dialog, nullptr, x, y, 0, 0,
               SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
}

void ThinBorder(HWND dialog, int id) {
  const HWND edit = GetDlgItem(dialog, id);
  if (!edit) {
    return;
  }
  SetWindowLongPtrW(edit, GWL_EXSTYLE,
                    GetWindowLongPtrW(edit, GWL_EXSTYLE) &
                        ~WS_EX_CLIENTEDGE);
  SetWindowLongPtrW(edit, GWL_STYLE,
                    GetWindowLongPtrW(edit, GWL_STYLE) | WS_BORDER);
  SetWindowPos(edit, nullptr, 0, 0, 0, 0,
               SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER |
                   SWP_FRAMECHANGED);
}

} // namespace

void Initialize(HWND dialog, HFONT* owned_font,
                std::initializer_list<int> bordered_edits) {
  for (const int id : bordered_edits) {
    ThinBorder(dialog, id);
  }
  HFONT font = ui::DefaultUIFont();
  if (owned_font) {
    *owned_font = font;
  }
  if (font) {
    SendMessageW(dialog, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    EnumChildWindows(
        dialog,
        [](HWND child, LPARAM value) {
          SendMessageW(child, WM_SETFONT, static_cast<WPARAM>(value), TRUE);
          wchar_t class_name[16] = {};
          GetClassNameW(child, class_name, _countof(class_name));
          const LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
          if (_wcsicmp(class_name, L"Edit") == 0 &&
              (style & ES_MULTILINE) && !(style & ES_READONLY)) {
            if (style & ES_WANTRETURN) {
              SetWindowSubclass(child, MultilineProc, kMultilineSubclassId, 0);
            } else {
              SetWindowSubclass(child, SingleLineProc, kSingleLineSubclassId, 0);
            }
          }
          return TRUE;
        },
        reinterpret_cast<LPARAM>(font));
  }
  Theme::Current().ApplyToWindow(dialog);
  Theme::Current().ApplyToChildren(dialog);
  Center(dialog);
}

void ReleaseFont(HFONT* font) {
  if (font && *font) {
    DeleteObject(*font);
    *font = nullptr;
  }
}

bool HandleThemeMessage(HWND dialog, UINT message, WPARAM wparam,
                        LPARAM lparam, INT_PTR* result) {
  if (!result) {
    return false;
  }
  if (message == WM_SETTINGCHANGE) {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dialog);
      Theme::Current().ApplyToChildren(dialog);
      InvalidateRect(dialog, nullptr, TRUE);
    }
    *result = TRUE;
    return true;
  }
  if (message == WM_ERASEBKGND) {
    RECT rect = {};
    GetClientRect(dialog, &rect);
    FillRect(reinterpret_cast<HDC>(wparam), &rect,
             Theme::Current().BackgroundBrush());
    DrawSizeGrip(dialog, reinterpret_cast<HDC>(wparam));
    *result = TRUE;
    return true;
  }
  if (message == WM_NCHITTEST) {
    RECT grip = {};
    if (SizeGripRect(dialog, &grip)) {
      MapWindowPoints(dialog, nullptr, reinterpret_cast<POINT*>(&grip), 2);
      const POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      if (PtInRect(&grip, pt)) {
        SetWindowLongPtrW(dialog, DWLP_MSGRESULT, HTBOTTOMRIGHT);
        *result = TRUE;
        return true;
      }
    }
  }
  int color_type = 0;
  switch (message) {
  case WM_CTLCOLORDLG:
    color_type = CTLCOLOR_DLG;
    break;
  case WM_CTLCOLORSTATIC:
    color_type = CTLCOLOR_STATIC;
    break;
  case WM_CTLCOLOREDIT:
    color_type = CTLCOLOR_EDIT;
    break;
  case WM_CTLCOLORLISTBOX:
    color_type = CTLCOLOR_LISTBOX;
    break;
  case WM_CTLCOLORBTN:
    color_type = CTLCOLOR_BTN;
    break;
  default:
    return false;
  }
  *result = reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(
      reinterpret_cast<HDC>(wparam), reinterpret_cast<HWND>(lparam),
      color_type));
  return true;
}

std::wstring ReadText(HWND dialog, int control_id) {
  return util::WindowText(GetDlgItem(dialog, control_id));
}

} // namespace regkit::editors::dialog_support
