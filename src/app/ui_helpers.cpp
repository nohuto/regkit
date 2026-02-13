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

#include "app/ui_helpers.h"

#include <algorithm>
#include <cwchar>
#include <limits>
#include <vector>

#include <commctrl.h>
#include <shellapi.h>
#include <uxtheme.h>

#include "app/theme.h"
#include "win32/win32_helpers.h"

namespace regkit::ui {

namespace {
#ifndef WC_LINK
#define WC_LINK L"SysLink"
#endif

constexpr wchar_t kDeleteClass[] = L"RegKitDeleteDialog";
constexpr wchar_t kErrorClass[] = L"RegKitErrorDialog";
constexpr wchar_t kChoiceClass[] = L"RegKitChoiceDialog";
constexpr wchar_t kAboutClass[] = L"RegKitAboutDialog";
constexpr wchar_t kAppTitle[] = L"RegKit";

struct PenCacheEntry {
  COLORREF color = RGB(0, 0, 0);
  int width = 0;
  HPEN pen = nullptr;
};

HPEN GetCachedPen(COLORREF color, int width) {
  static PenCacheEntry cache[8];
  for (auto& entry : cache) {
    if (entry.pen && entry.color == color && entry.width == width) {
      return entry.pen;
    }
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

struct FontSettings {
  bool use_default = true;
  std::wstring face;
  int size = 0;
  int weight = 0;
  bool italic = false;
  bool italic_set = false;
};

struct DeleteDialogState {
  HWND hwnd = nullptr;
  HWND text = nullptr;
  HWND yes_btn = nullptr;
  HWND no_btn = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  std::wstring title;
  std::wstring message;
  bool result = false;
  bool accepted = false;
  bool owner_restored = false;
};

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
  HWND text = nullptr;
  HWND yes_btn = nullptr;
  HWND no_btn = nullptr;
  HWND cancel_btn = nullptr;
  HWND owner = nullptr;
  HFONT font = nullptr;
  std::wstring title;
  std::wstring message;
  std::wstring yes_label;
  std::wstring no_label;
  std::wstring cancel_label;
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

void RestoreOwnerWindow(HWND owner, bool* restored) {
  if (!owner || !restored || *restored) {
    return;
  }
  EnableWindow(owner, TRUE);
  SetActiveWindow(owner);
  SetForegroundWindow(owner);
  *restored = true;
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
  RECT work_area = {};
  if (SystemParametersInfoW(SPI_GETWORKAREA, 0, &work_area, 0)) {
    int work_w = work_area.right - work_area.left;
    int work_h = work_area.bottom - work_area.top;
    int x = work_area.left + std::max(0, (work_w - width) / 2);
    int y = work_area.top + std::max(0, (work_h - height) / 2);
    SetWindowPos(hwnd, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
  }
}

std::wstring TrimWhitespace(const std::wstring& text) {
  size_t start = 0;
  while (start < text.size() && (text[start] == L' ' || text[start] == L'\t')) {
    ++start;
  }
  if (start == text.size()) {
    return L"";
  }
  size_t end = text.size() - 1;
  while (end > start && (text[end] == L' ' || text[end] == L'\t')) {
    --end;
  }
  return text.substr(start, end - start + 1);
}

bool ParseBool(const std::wstring& value) {
  return (_wcsicmp(value.c_str(), L"1") == 0 || _wcsicmp(value.c_str(), L"true") == 0 || _wcsicmp(value.c_str(), L"yes") == 0);
}

bool LoadFontSettings(FontSettings* out) {
  if (!out) {
    return false;
  }
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return false;
  }
  std::wstring path = util::JoinPath(folder, L"settings.ini");
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return false;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return false;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(buffer);
  if (content.empty()) {
    return false;
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
    start = end + 1;
    if (line.empty()) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = TrimWhitespace(line.substr(0, sep));
    std::wstring value = TrimWhitespace(line.substr(sep + 1));
    if (_wcsicmp(key.c_str(), L"font_use_default") == 0) {
      out->use_default = ParseBool(value);
    } else if (_wcsicmp(key.c_str(), L"font_face") == 0) {
      out->face = value;
    } else if (_wcsicmp(key.c_str(), L"font_size") == 0) {
      int size_value = _wtoi(value.c_str());
      if (size_value > 0) {
        out->size = size_value;
      }
    } else if (_wcsicmp(key.c_str(), L"font_weight") == 0) {
      int weight_value = _wtoi(value.c_str());
      if (weight_value > 0) {
        out->weight = weight_value;
      }
    } else if (_wcsicmp(key.c_str(), L"font_italic") == 0) {
      out->italic = ParseBool(value);
      out->italic_set = true;
    }
  }
  return true;
}

std::wstring ReadFontSubstitute(const wchar_t* value_name) {
  if (!value_name || !*value_name) {
    return L"";
  }
  const wchar_t* subkey = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\FontSubstitutes";
  auto query = [&](REGSAM sam) -> std::wstring {
    HKEY key = nullptr;
    LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, subkey, 0, sam, &key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    DWORD type = 0;
    DWORD bytes = 0;
    result = RegQueryValueExW(key, value_name, nullptr, &type, nullptr, &bytes);
    if (result != ERROR_SUCCESS || bytes == 0 || (type != REG_SZ && type != REG_EXPAND_SZ)) {
      RegCloseKey(key);
      return L"";
    }
    std::vector<wchar_t> buffer(bytes / sizeof(wchar_t) + 1, L'\0');
    result = RegQueryValueExW(key, value_name, nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &bytes);
    RegCloseKey(key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    std::wstring value(buffer.data());
    while (!value.empty() && value.back() == L'\0') {
      value.pop_back();
    }
    if (value.empty()) {
      return L"";
    }
    if (type == REG_EXPAND_SZ) {
      wchar_t expanded[512] = {};
      DWORD expanded_len = ExpandEnvironmentStringsW(value.c_str(), expanded, static_cast<DWORD>(_countof(expanded)));
      if (expanded_len > 0 && expanded_len < _countof(expanded)) {
        value.assign(expanded, expanded_len - 1);
      }
    }
    return value;
  };

  std::wstring value = query(KEY_READ | KEY_WOW64_64KEY);
  if (!value.empty()) {
    return value;
  }
  return query(KEY_READ);
}

int FontHeightFromPointSize(int point_size) {
  HDC hdc = GetDC(nullptr);
  int height = -MulDiv(point_size, GetDeviceCaps(hdc, LOGPIXELSY), 72);
  ReleaseDC(nullptr, hdc);
  return height;
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

void LayoutChoiceDialog(HWND hwnd, ChoiceDialogState* state) {
  if (!hwnd || !state) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  int padding = 12;
  int base_units = GetDialogBaseUnits();
  int base_x = std::max(1, static_cast<int>(LOWORD(base_units)));
  int base_y = std::max(1, static_cast<int>(HIWORD(base_units)));
  int btn_w = MulDiv(45, base_x, 4);
  int btn_h = MulDiv(11, base_y, 8);
  int btn_y = height - padding - btn_h;
  int text_h = btn_y - padding;
  int gap = 8;
  if (state->text) {
    SetWindowPos(state->text, nullptr, padding, padding, width - padding * 2, text_h, SWP_NOZORDER);
  }
  int total_w = btn_w * 3 + gap * 2;
  int start_x = std::max(padding, width - padding - total_w);
  if (state->yes_btn) {
    SetWindowPos(state->yes_btn, nullptr, start_x, btn_y, btn_w, btn_h, SWP_NOZORDER);
  }
  if (state->no_btn) {
    SetWindowPos(state->no_btn, nullptr, start_x + btn_w + gap, btn_y, btn_w, btn_h, SWP_NOZORDER);
  }
  if (state->cancel_btn) {
    SetWindowPos(state->cancel_btn, nullptr, start_x + (btn_w + gap) * 2, btn_y, btn_w, btn_h, SWP_NOZORDER);
  }
}

void LayoutDeleteDialog(HWND hwnd, DeleteDialogState* state) {
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

  int no_x = width - padding - btn_w;
  int yes_x = no_x - 8 - btn_w;
  if (state->yes_btn) {
    SetWindowPos(state->yes_btn, nullptr, yes_x, btn_y, btn_w, btn_h, SWP_NOZORDER);
  }
  if (state->no_btn) {
    SetWindowPos(state->no_btn, nullptr, no_x, btn_y, btn_w, btn_h, SWP_NOZORDER);
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
    state->text = CreateWindowExW(0, L"STATIC", state->message.c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->yes_btn = CreateWindowExW(0, L"BUTTON", state->yes_label.c_str(), WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDYES), nullptr, nullptr);
    state->no_btn = CreateWindowExW(0, L"BUTTON", state->no_label.c_str(), WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDNO), nullptr, nullptr);
    state->cancel_btn = CreateWindowExW(0, L"BUTTON", state->cancel_label.c_str(), WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);

    ApplyConfirmFonts(hwnd, state->font);
    Theme::Current().ApplyToWindow(hwnd);
    Theme::Current().ApplyToChildren(hwnd);
    LayoutChoiceDialog(hwnd, state);
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
  case WM_COMMAND:
    switch (LOWORD(wparam)) {
    case IDYES:
      state->result = IDYES;
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    case IDNO:
      state->result = IDNO;
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    case IDCANCEL:
      state->result = IDCANCEL;
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
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
      RestoreOwnerWindow(state->owner, &state->owner_restored);
    }
    DestroyWindow(hwnd);
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

LRESULT CALLBACK DeleteDialogProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<DeleteDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<DeleteDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, L"RegKit");
    state->font = DefaultUIFont();
    state->text = CreateWindowExW(0, L"STATIC", state->message.c_str(), WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->yes_btn = CreateWindowExW(0, L"BUTTON", L"Yes", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDYES), nullptr, nullptr);
    state->no_btn = CreateWindowExW(0, L"BUTTON", L"No", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDNO), nullptr, nullptr);

    ApplyConfirmFonts(hwnd, state->font);
    Theme::Current().ApplyToWindow(hwnd);
    Theme::Current().ApplyToChildren(hwnd);
    LayoutDeleteDialog(hwnd, state);
    return 0;
  }
  case WM_SHOWWINDOW:
    if (wparam && state && state->yes_btn) {
      SetFocus(state->yes_btn);
      return 0;
    }
    break;
  case WM_SETFOCUS:
    if (state && state->yes_btn) {
      SetFocus(state->yes_btn);
      return 0;
    }
    break;
  case WM_SIZE:
    LayoutDeleteDialog(hwnd, state);
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
    case IDYES:
      if (state) {
        state->result = true;
        state->accepted = true;
        RestoreOwnerWindow(state->owner, &state->owner_restored);
      }
      DestroyWindow(hwnd);
      return 0;
    case IDNO:
    case IDCANCEL:
      if (state) {
        state->result = false;
        state->accepted = true;
        RestoreOwnerWindow(state->owner, &state->owner_restored);
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
      state->result = false;
      state->accepted = true;
      RestoreOwnerWindow(state->owner, &state->owner_restored);
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
        RestoreOwnerWindow(state->owner, &state->owner_restored);
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
      RestoreOwnerWindow(state->owner, &state->owner_restored);
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
    state->credits = CreateWindowExW(0, L"STATIC", L"\x00A9 Noverse (Nohuto) 2026", WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->repo_link = CreateWindowExW(0, WC_LINK,
                                       L"Repository: <a href=\"https://github.com/nohuto/regkit\">"
                                       L"https://github.com/nohuto/regkit</a>",
                                       WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->discord_link = CreateWindowExW(0, WC_LINK,
                                          L"Discord: <a href=\"https://discord.com/invite/E2ybG4j9jU\">"
                                          L"https://discord.com/invite/E2ybG4j9jU</a>",
                                          WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->website_link = CreateWindowExW(0, WC_LINK,
                                          L"Website: <a href=\"https://www.noverse.dev/\">"
                                          L"https://www.noverse.dev/</a>",
                                          WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);
    state->email_link = CreateWindowExW(0, WC_LINK,
                                        L"Email: <a href=\"mailto:nohuto@duck.com\">"
                                        L"nohuto@duck.com</a>",
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
        RestoreOwnerWindow(state->owner, &state->owner_restored);
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
      RestoreOwnerWindow(state->owner, &state->owner_restored);
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

  RestoreOwnerWindow(owner, &state.owner_restored);
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

  RestoreOwnerWindow(owner, &state.owner_restored);
  return state.accepted;
}

bool ShowChoiceDialog(HWND owner, const std::wstring& title, const std::wstring& message, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label, int* result) {
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
  state.yes_label = yes_label;
  state.no_label = no_label;
  state.cancel_label = cancel_label;
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kChoiceClass, kAppTitle, WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 420, 120, owner, nullptr, wc.hInstance, &state);
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

  RestoreOwnerWindow(owner, &state.owner_restored);
  if (!state.accepted) {
    return false;
  }
  if (result) {
    *result = state.result;
  }
  return true;
}

bool ShowDeleteDialog(HWND owner, const std::wstring& title, const std::wstring& message, bool* result) {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = DeleteDialogProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kDeleteClass;
  RegisterClassW(&wc);

  DeleteDialogState state;
  state.owner = owner;
  state.title = title;
  state.message = message;
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kDeleteClass, kAppTitle, WS_POPUP | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 280, 120, owner, nullptr, wc.hInstance, &state);
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

  RestoreOwnerWindow(owner, &state.owner_restored);
  if (!state.accepted) {
    return false;
  }
  if (result) {
    *result = state.result;
  }
  return true;
}

HRESULT CALLBACK TaskDialogThemeCallback(HWND hwnd, UINT msg, WPARAM, LPARAM, LONG_PTR ref_data) {
  if (msg == TDN_CREATED) {
    Theme::Current().ApplyToWindow(hwnd);
    Theme::Current().ApplyToChildren(hwnd);
    CenterWindowToOwner(hwnd, reinterpret_cast<HWND>(ref_data));
  }
  return S_OK;
}

bool ShowTaskDialog(HWND owner, const std::wstring& title, const std::wstring& message, TASKDIALOG_COMMON_BUTTON_FLAGS buttons, int* button, PCWSTR icon) {
  TASKDIALOGCONFIG config = {};
  config.cbSize = sizeof(config);
  config.hwndParent = owner;
  config.dwFlags = TDF_ALLOW_DIALOG_CANCELLATION;
  config.dwCommonButtons = buttons;
  config.pszWindowTitle = title.empty() ? kAppTitle : title.c_str();
  config.pszContent = message.c_str();
  config.pszMainIcon = icon;
  config.pfCallback = TaskDialogThemeCallback;
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

namespace {
bool ListViewItemSelected(HWND list, int item_index) {
  return item_index >= 0 && (ListView_GetItemState(list, item_index, LVIS_SELECTED) & LVIS_SELECTED) != 0;
}

void ApplyListViewThemeColors(NMLVCUSTOMDRAW* draw, bool selected, const Theme& theme) {
  draw->clrText = selected ? theme.SelectionTextColor() : theme.TextColor();
  draw->clrTextBk = selected ? theme.SelectionColor() : theme.PanelColor();
  if (selected) {
    draw->nmcd.uItemState &= ~CDIS_SELECTED;
  }
  draw->nmcd.uItemState &= ~(CDIS_FOCUS | CDIS_HOT);
}
} // namespace

void DrawListViewFocusBorder(HWND list, HDC hdc, int item_index, COLORREF color) {
  if (!list || !hdc || item_index < 0) {
    return;
  }
  RECT rect = {};
  if (!ListView_GetItemRect(list, item_index, &rect, LVIR_BOUNDS)) {
    return;
  }
  InflateRect(&rect, -1, -1);
  HPEN pen = GetCachedPen(color, 1);
  HGDIOBJ old_pen = SelectObject(hdc, pen);
  HGDIOBJ old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
  Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
  SelectObject(hdc, old_brush);
  SelectObject(hdc, old_pen);
}

LRESULT HandleThemedListViewCustomDraw(HWND list, NMLVCUSTOMDRAW* draw) {
  if (!list || !draw) {
    return CDRF_DODEFAULT;
  }
  const Theme& theme = Theme::Current();
  switch (draw->nmcd.dwDrawStage) {
  case CDDS_PREPAINT:
    return CDRF_NOTIFYITEMDRAW;
  case CDDS_ITEMPREPAINT: {
    int item_index = static_cast<int>(draw->nmcd.dwItemSpec);
    bool selected = ListViewItemSelected(list, item_index);
    ApplyListViewThemeColors(draw, selected, theme);
    return CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW | CDRF_NOTIFYPOSTPAINT;
  }
  case CDDS_ITEMPREPAINT | CDDS_SUBITEM: {
    int item_index = static_cast<int>(draw->nmcd.dwItemSpec);
    bool selected = ListViewItemSelected(list, item_index);
    ApplyListViewThemeColors(draw, selected, theme);
    return CDRF_NEWFONT;
  }
  case CDDS_ITEMPOSTPAINT: {
    int item_index = static_cast<int>(draw->nmcd.dwItemSpec);
    if (ListViewItemSelected(list, item_index)) {
      COLORREF border = (GetFocus() == list) ? theme.FocusColor() : theme.BorderColor();
      DrawListViewFocusBorder(list, draw->nmcd.hdc, item_index, border);
    }
    return CDRF_SKIPDEFAULT;
  }
  default:
    break;
  }
  return CDRF_DODEFAULT;
}

LOGFONTW DefaultUIFontLogFont() {
  LOGFONTW lf = {};
  HFONT stock = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
  if (stock) {
    GetObjectW(stock, sizeof(lf), &lf);
  }
  if (lf.lfHeight == 0) {
    lf.lfHeight = FontHeightFromPointSize(9);
    lf.lfWeight = FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
  }
  std::wstring default_face = ReadFontSubstitute(L"Segoe UI");
  if (default_face.empty()) {
    default_face = L"Segoe UI";
  }
  wcsncpy_s(lf.lfFaceName, default_face.c_str(), _TRUNCATE);

  FontSettings settings;
  if (LoadFontSettings(&settings) && !settings.use_default) {
    if (!settings.face.empty()) {
      wcsncpy_s(lf.lfFaceName, settings.face.c_str(), _TRUNCATE);
    }
    if (settings.size > 0) {
      lf.lfHeight = FontHeightFromPointSize(settings.size);
    }
    if (settings.weight > 0) {
      lf.lfWeight = settings.weight;
    }
    if (settings.italic_set) {
      lf.lfItalic = settings.italic ? TRUE : FALSE;
    }
  }
  return lf;
}

HFONT DefaultUIFont() {
  LOGFONTW lf = DefaultUIFontLogFont();
  HFONT font = CreateFontIndirectW(&lf);
  return font;
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
  if (ShowErrorDialog(owner, message)) {
    return;
  }
  if (!ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_ERROR_ICON)) {
    ShowErrorDialog(owner, message);
  }
}

void ShowWarning(HWND owner, const std::wstring& message) {
  if (Theme::UseDarkMode()) {
    if (ShowErrorDialog(owner, message)) {
      return;
    }
  }
  if (!ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_WARNING_ICON)) {
    ShowErrorDialog(owner, message);
  }
}

void ShowInfo(HWND owner, const std::wstring& message) {
  if (!ShowTaskDialog(owner, kAppTitle, message, TDCBF_OK_BUTTON, nullptr, TD_INFORMATION_ICON)) {
    ShowErrorDialog(owner, message);
  }
}

void ShowAbout(HWND owner) {
  if (ShowAboutDialog(owner)) {
    return;
  }
  ShowInfo(owner, L"\x00A9 Noverse (Nohuto) 2026\n"
                  L"Repository: https://github.com/nohuto/regkit\n"
                  L"Discord: https://discord.com/invite/E2ybG4j9jU\n"
                  L"Website: https://www.noverse.dev/\n"
                  L"Email: nohuto@duck.com");
}

bool ConfirmDelete(HWND owner, const std::wstring& title, const std::wstring& name) {
  bool result = false;
  std::wstring message;
  if (_wcsicmp(title.c_str(), L"Delete Key") == 0) {
    message = L"Delete key \"" + name + L"\"?";
  } else if (_wcsicmp(title.c_str(), L"Delete Value") == 0) {
    message = L"Delete value \"" + name + L"\"?";
  } else {
    message = L"Delete \"" + name + L"\"?";
  }
  if (ShowDeleteDialog(owner, title, message, &result)) {
    return result;
  }
  int clicked = 0;
  if (ShowTaskDialog(owner, title, message, TDCBF_YES_BUTTON | TDCBF_NO_BUTTON, &clicked, TD_WARNING_ICON)) {
    return clicked == IDYES;
  }
  return false;
}

int PromptYesNoCancel(HWND owner, const std::wstring& message, const std::wstring& title) {
  int clicked = 0;
  if (ShowTaskDialog(owner, title, message, TDCBF_YES_BUTTON | TDCBF_NO_BUTTON | TDCBF_CANCEL_BUTTON, &clicked, TD_WARNING_ICON)) {
    return clicked;
  }
  return IDCANCEL;
}

int PromptChoice(HWND owner, const std::wstring& message, const std::wstring& title, const std::wstring& yes_label, const std::wstring& no_label, const std::wstring& cancel_label) {
  int result = IDCANCEL;
  if (ShowChoiceDialog(owner, title, message, yes_label, no_label, cancel_label, &result)) {
    return result;
  }
  return IDCANCEL;
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
