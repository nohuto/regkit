// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "search/replace_dialog.h"

#include <algorithm>

#include <commctrl.h>
#include <uxtheme.h>

#include "appearance/dialog_layout.h"
#include "appearance/dialog_metrics.h"
#include "win32/window_metrics.h"
#include "search/query_dialog.h"
#include "appearance/theme.h"
#include "appearance/default_font.h"
#include "appearance/feedback.h"
#include "win32/text_transform.h"

namespace regkit {

namespace {

constexpr wchar_t kDialogClass[] = L"RegKitReplaceDialog";
constexpr int kReplaceButtonWidth = 80;
constexpr int kNumberDecimalWidth = 144;
constexpr int kNumberHexWidth = 116;
constexpr int kCheckBoxIdealPadding = 12;

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
  kSearchKeys = 125,
  kSearchValues = 126,
  kSearchData = 127,
  kValueDataGroup = 128,
  kNumberDecimal = 129,
  kNumberHex = 130,
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
  HWND search_keys = nullptr;
  HWND search_values = nullptr;
  HWND search_data = nullptr;
  HWND number_decimal = nullptr;
  HWND number_hex = nullptr;
  HWND replace_button = nullptr;
  HWND cancel_button = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  ReplaceDialogResult* out = nullptr;
  bool accepted = false;
  bool owner_restored = false;
};

void UpdateValueDataOptions(
    HWND hwnd,
    const ReplaceDialogState* state
) {
  if (!state) {
    return;
  }
  const bool enabled =
      SendMessageW(state->search_data, BM_GETCHECK, 0, 0) == BST_CHECKED;
  EnableWindow(GetDlgItem(hwnd, kValueDataGroup), enabled);
  EnableWindow(state->number_decimal, enabled);
  EnableWindow(state->number_hex, enabled);
}

HFONT CreateDialogFont(
    HWND hwnd
) {
  return ui::DefaultUIFont(win32::DpiForWindow(hwnd));
}

int CheckBoxIdealWidth(
    HWND control,
    int fallback
) {
  SIZE ideal = {};
  if (control &&
      SendMessageW(control, BCM_GETIDEALSIZE, 0, reinterpret_cast<LPARAM>(&ideal)) &&
      ideal.cx > 0) {
    return ideal.cx +
           appearance::metrics::Scaled(kCheckBoxIdealPadding, win32::DpiForWindow(control));
  }
  return fallback;
}

void LayoutDialog(
    HWND hwnd,
    ReplaceDialogState* state,
    HFONT font
) {
  if (!hwnd || !state) {
    return;
  }
  using namespace appearance::metrics;
  RECT client = {};
  GetClientRect(hwnd, &client);
  const UINT dpi = win32::DpiForWindow(hwnd);
  const int margin = Scaled(kDialogContentMargin, dpi);
  const int block_gap = Scaled(kBlockGap, dpi);
  const int label_gap = Scaled(kLabelGap, dpi);
  const int label_inset = Scaled(kLabelInset, dpi);
  const int label_h = Scaled(kLabelHeight, dpi);
  const int line_h = Scaled(kControlHeight, dpi);
  const int control_pitch = Scaled(kControlPitch, dpi);
  const int row_pitch = Scaled(kRowPitch, dpi);
  const int check_h = Scaled(kCheckHeight, dpi);
  const int group_top = Scaled(kGroupTop, dpi);
  const int group_bottom = Scaled(kGroupBottom, dpi);
  const int group_inset = Scaled(kGroupInset, dpi);
  const int button_h = Scaled(kButtonHeight, dpi);
  const int button_gap = Scaled(kButtonGap, dpi);
  const int button_w = Scaled(kButtonMinWidth, dpi);
  const int replace_w = Scaled(kReplaceButtonWidth, dpi);
  const int right_margin = Scaled(kDialogButtonRightMargin, dpi);
  const int bottom_margin = Scaled(kDialogButtonBottomMargin, dpi);
  const int width = client.right - client.left;
  const int x = margin;
  const int label_w = Scaled(90, dpi);
  const int key_label_w = Scaled(32, dpi);
  const int browse_w = Scaled(90, dpi);
  int y = margin;

  HWND find_label = GetDlgItem(hwnd, kFindLabel);
  appearance::Place(find_label, x, y + label_inset, label_w, label_h);
  const int edit_w = width - x * 2 - label_w - label_gap;
  appearance::Place(state->find_edit, x + label_w + label_gap, y, edit_w, line_h);
  y += control_pitch;

  HWND replace_label = GetDlgItem(hwnd, kReplaceLabel);
  appearance::Place(replace_label, x, y + label_inset, label_w, label_h);
  appearance::Place(state->replace_edit, x + label_w + label_gap, y, edit_w, line_h);
  y += line_h + block_gap;

  const int group_w = width - x * 2;
  const int where_h = group_top + line_h + group_bottom;
  appearance::Place(GetDlgItem(hwnd, kWhereGroup), x, y, group_w, where_h);
  const int gx = x + group_inset;
  const int gy = y + group_top;
  const int key_w = group_w - group_inset * 2 - key_label_w - label_gap * 2 - browse_w;
  appearance::Place(GetDlgItem(hwnd, kKeyLabel), gx, gy + label_inset, key_label_w, label_h);
  appearance::Place(state->key_edit, gx + key_label_w + label_gap, gy, key_w, line_h);
  appearance::Place(state->key_browse, gx + key_label_w + label_gap * 2 + key_w, gy, browse_w, line_h);
  y += where_h + block_gap;

  const int nested_w = group_w - group_inset * 2;
  const int nested_h = group_top + check_h + group_bottom;
  const int options_h =
      group_top + row_pitch * 4 + nested_h + group_bottom;
  appearance::Place(GetDlgItem(hwnd, kOptionsGroup), x, y, group_w, options_h);
  const int ox = x + group_inset;
  const int oy = y + group_top;
  const int col_w = (group_w - group_inset * 2 - label_gap) / 2;
  const int col2_x = ox + col_w + label_gap;
  appearance::Place(state->recursive, ox, oy, col_w, check_h);
  appearance::Place(state->match_whole, ox, oy + row_pitch, col_w, check_h);
  appearance::Place(state->match_case, ox, oy + row_pitch * 2, col_w, check_h);
  appearance::Place(state->use_regex, ox, oy + row_pitch * 3, col_w, check_h);
  appearance::Place(state->search_keys, col2_x, oy, col_w, check_h);
  appearance::Place(state->search_values, col2_x, oy + row_pitch, col_w, check_h);
  appearance::Place(state->search_data, col2_x, oy + row_pitch * 2, col_w, check_h);
  const int nested_y = oy + row_pitch * 4;
  appearance::Place(GetDlgItem(hwnd, kValueDataGroup), ox, nested_y, nested_w, nested_h);
  const int ny = nested_y + group_top;
  const int half_w = (nested_w - group_inset * 2) / 2;
  const int half_x = ox + group_inset;
  const int dec_w = CheckBoxIdealWidth(state->number_decimal, Scaled(kNumberDecimalWidth, dpi));
  const int hex_w = CheckBoxIdealWidth(state->number_hex, Scaled(kNumberHexWidth, dpi));
  appearance::Place(state->number_decimal, half_x + (half_w - dec_w) / 2, ny, dec_w, check_h);
  appearance::Place(state->number_hex, half_x + half_w + (half_w - hex_w) / 2, ny, hex_w, check_h);
  y += options_h + block_gap;

  const int cancel_x = width - right_margin - button_w;
  appearance::Place(state->replace_button, cancel_x - button_gap - replace_w, y, replace_w, button_h);
  appearance::Place(state->cancel_button, cancel_x, y, button_w, button_h);
  appearance::FitDialogHeight(hwnd, y + button_h + bottom_margin);

  appearance::SetControlFont(hwnd, font);
  appearance::SetControlFont(find_label, font);
  appearance::SetControlFont(replace_label, font);
  for (HWND edit : {state->find_edit, state->replace_edit, state->key_edit}) {
    appearance::CenterEditText(edit, font, 2, 2);
  }
}

LRESULT CALLBACK ReplaceDialogProc(
    HWND hwnd,
    UINT msg,
    WPARAM wparam,
    LPARAM lparam
) {
  auto* state = reinterpret_cast<ReplaceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE:
    {
      auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
      return DefWindowProcW(hwnd, msg, wparam, lparam);
    }
  case WM_CREATE:
    {
      state = reinterpret_cast<ReplaceDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
      if (!state) {
        return -1;
      }
      state->hwnd = hwnd;
      SetWindowTextW(hwnd, L"Replace");
      state->font = CreateDialogFont(hwnd);
      HFONT font = state->font;

      CreateWindowExW(0, L"STATIC", L"Find what:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kFindLabel), nullptr, nullptr);
      state->find_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL | ES_MULTILINE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kFindEdit), nullptr, nullptr);

      CreateWindowExW(0, L"STATIC", L"Replace with:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kReplaceLabel), nullptr, nullptr);
      state->replace_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL | ES_MULTILINE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kReplaceEdit), nullptr, nullptr);

      CreateWindowExW(0, L"BUTTON", L"Where to search", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kWhereGroup), nullptr, nullptr);
      CreateWindowExW(0, L"STATIC", L"Key:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kKeyLabel), nullptr, nullptr);
      state->key_edit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_AUTOHSCROLL | ES_MULTILINE, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kKeyEdit), nullptr, nullptr);
      state->key_browse = CreateWindowExW(0, L"BUTTON", L"Browse...", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kKeyBrowse), nullptr, nullptr);

      CreateWindowExW(0, L"BUTTON", L"Options", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kOptionsGroup), nullptr, nullptr);
      state->recursive = CreateWindowExW(0, L"BUTTON", L"Recursive", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kRecursive), nullptr, nullptr);
      state->match_case = CreateWindowExW(0, L"BUTTON", L"Match case", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kMatchCase), nullptr, nullptr);
      state->match_whole = CreateWindowExW(0, L"BUTTON", L"Match whole string", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kMatchWhole), nullptr, nullptr);
      state->use_regex = CreateWindowExW(0, L"BUTTON", L"Regular expressions", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kUseRegex), nullptr, nullptr);
      state->search_keys = CreateWindowExW(0, L"BUTTON", L"Replace in key names", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kSearchKeys), nullptr, nullptr);
      state->search_values = CreateWindowExW(0, L"BUTTON", L"Replace in value names", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kSearchValues), nullptr, nullptr);
      state->search_data = CreateWindowExW(0, L"BUTTON", L"Replace in value data", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kSearchData), nullptr, nullptr);

      CreateWindowExW(0, L"BUTTON", L"Value Data", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kValueDataGroup), nullptr, nullptr);
      state->number_decimal = CreateWindowExW(0, L"BUTTON", L"Numbers as decimal", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kNumberDecimal), nullptr, nullptr);
      state->number_hex = CreateWindowExW(0, L"BUTTON", L"Numbers as hex", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kNumberHex), nullptr, nullptr);

      state->replace_button = CreateWindowExW(0, L"BUTTON", L"Replace", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kReplaceButton), nullptr, nullptr);
      state->cancel_button = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(kCancelButton), nullptr, nullptr);

      if (state->out) {
        SetWindowTextW(state->find_edit, state->out->find_text.c_str());
        SetWindowTextW(state->replace_edit, state->out->replace_text.c_str());
        SetWindowTextW(state->key_edit, state->out->start_key.c_str());
        SendMessageW(state->recursive, BM_SETCHECK, state->out->recursive ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->match_case, BM_SETCHECK, state->out->match_case ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->match_whole, BM_SETCHECK, state->out->match_whole ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->use_regex, BM_SETCHECK, state->out->use_regex ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->search_keys, BM_SETCHECK, state->out->replace_keys ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->search_values, BM_SETCHECK, state->out->replace_values ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->search_data, BM_SETCHECK, state->out->replace_data ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->number_decimal, BM_SETCHECK, state->out->number_decimal ? BST_CHECKED : BST_UNCHECKED, 0);
        SendMessageW(state->number_hex, BM_SETCHECK, state->out->number_hex ? BST_CHECKED : BST_UNCHECKED, 0);
      } else {
        SendMessageW(state->recursive, BM_SETCHECK, BST_CHECKED, 0);
        SendMessageW(state->search_values, BM_SETCHECK, BST_CHECKED, 0);
        SendMessageW(state->search_data, BM_SETCHECK, BST_CHECKED, 0);
        SendMessageW(state->number_decimal, BM_SETCHECK, BST_CHECKED, 0);
      }
      UpdateValueDataOptions(hwnd, state);

      EnumChildWindows(
          hwnd,
          [](HWND child, LPARAM param) -> BOOL {
            HFONT font_handle = reinterpret_cast<HFONT>(param);
            appearance::SetControlFont(child, font_handle);
            return TRUE;
          },
          reinterpret_cast<LPARAM>(font)
      );

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
  case WM_DPICHANGED:
    if (state) {
      appearance::RefreshDialogFont(hwnd, &state->font, LOWORD(wparam));
    }
    appearance::ApplyDpiChange(hwnd, lparam);
    return 0;
  case WM_SIZE:
    {
      HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
      LayoutDialog(hwnd, state, font);
      return 0;
    }
  case WM_ERASEBKGND:
    {
      HDC hdc = reinterpret_cast<HDC>(wparam);
      RECT rect = {};
      GetClientRect(hwnd, &rect);
      FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
      return 1;
    }
  case WM_SETTINGCHANGE:
    {
      if (Theme::UpdateFromSystem()) {
        Theme::Current().ApplyToWindow(hwnd);
        Theme::Current().ApplyToChildren(hwnd);
        InvalidateRect(hwnd, nullptr, TRUE);
      }
      return 0;
    }
  case WM_CTLCOLORSTATIC:
    {
      HDC hdc = reinterpret_cast<HDC>(wparam);
      HWND target = reinterpret_cast<HWND>(lparam);
      return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_STATIC));
    }
  case WM_CTLCOLOREDIT:
    {
      HDC hdc = reinterpret_cast<HDC>(wparam);
      HWND target = reinterpret_cast<HWND>(lparam);
      return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_EDIT));
    }
  case WM_CTLCOLORBTN:
    {
      HDC hdc = reinterpret_cast<HDC>(wparam);
      HWND target = reinterpret_cast<HWND>(lparam);
      return reinterpret_cast<LRESULT>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_BTN));
    }
  case DM_GETDEFID:
    return MAKELRESULT(IDOK, DC_HASDEFID);
  case WM_COMMAND:
    {
      if (!state) {
        return 0;
      }
      switch (LOWORD(wparam)) {
      case kSearchData:
        UpdateValueDataOptions(hwnd, state);
        return 0;
      case kKeyBrowse:
        {
          std::wstring selected;
          if (ShowBrowseKeyDialog(hwnd, &selected)) {
            if (!selected.empty()) {
              SetWindowTextW(state->key_edit, selected.c_str());
            }
          }
          return 0;
        }
      case kReplaceButton:
        {
          const std::wstring find_value = util::WindowText(state->find_edit);
          if (find_value.empty()) {
            ui::ShowError(hwnd, L"Enter text to find.");
            return 0;
          }
          const std::wstring replace_text = util::WindowText(state->replace_edit);
          const std::wstring key_text = util::WindowText(state->key_edit);

          if (state->out) {
            state->out->find_text = find_value;
            state->out->replace_text = replace_text;
            state->out->start_key = key_text;
            state->out->recursive = SendMessageW(state->recursive, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->match_case = SendMessageW(state->match_case, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->match_whole = SendMessageW(state->match_whole, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->use_regex = SendMessageW(state->use_regex, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->replace_keys = SendMessageW(state->search_keys, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->replace_values = SendMessageW(state->search_values, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->replace_data = SendMessageW(state->search_data, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->number_decimal = SendMessageW(state->number_decimal, BM_GETCHECK, 0, 0) == BST_CHECKED;
            state->out->number_hex = SendMessageW(state->number_hex, BM_GETCHECK, 0, 0) == BST_CHECKED;
            if (!state->out->replace_keys && !state->out->replace_values &&
                !state->out->replace_data) {
              ui::ShowError(hwnd, L"Select what should be replaced.");
              return 0;
            }
          }
          state->accepted = true;
          appearance::RestoreDialogOwner(state->owner, &state->owner_restored);
          DestroyWindow(hwnd);
          return 0;
        }
      case kCancelButton:
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

HWND CreateReplaceDialogWindow(
    HINSTANCE instance,
    HWND owner,
    ReplaceDialogState* state
) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = ReplaceDialogProc;
  wc.hInstance = instance;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kDialogClass;
  RegisterClassW(&wc);

  return CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kDialogClass, L"Replace", WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, appearance::metrics::Scaled(520, win32::DpiForWindow(owner)), appearance::metrics::Scaled(360, win32::DpiForWindow(owner)), owner, nullptr, instance, state);
}

} // namespace

bool ShowReplaceDialog(
    HWND owner,
    ReplaceDialogResult* result
) {
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
  appearance::CenterWindow(hwnd, owner);

  EnableWindow(owner, FALSE);
  ShowWindow(hwnd, SW_SHOW);
  UpdateWindow(hwnd);

  appearance::RunModalLoop(hwnd);

  appearance::RestoreDialogOwner(owner, &state.owner_restored);
  return state.accepted;
}

} // namespace regkit
