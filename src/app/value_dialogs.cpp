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

#include "app/value_dialogs.h"

#include <algorithm>
#include <cerrno>
#include <commctrl.h>
#include <commdlg.h>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <limits>
#include <sstream>
#include <uxtheme.h>

#include "app/theme.h"
#include "app/ui_helpers.h"
#include "resource.h"

namespace regkit {

namespace {

struct TextDialogState {
  const wchar_t* title = nullptr;
  const wchar_t* label = nullptr;
  std::wstring text;
  std::wstring value_name;
  std::wstring value_type;
  bool show_details = false;
  HFONT ui_font = nullptr;
};

struct NumberDialogState {
  const wchar_t* title = nullptr;
  const wchar_t* label = nullptr;
  int base = 16;
  unsigned long long value = 0;
  std::wstring value_name;
  std::wstring value_type;
  bool show_details = false;
  HFONT ui_font = nullptr;
};

struct BinaryDialogState {
  std::wstring text;
  int group_bytes = 1;
  bool unicode = false;
  HFONT mono_font = nullptr;
  bool updating = false;
  std::wstring value_name;
  std::wstring value_type;
  bool show_details = false;
  HFONT ui_font = nullptr;
};

struct BinaryGroupState {
  int group_bytes = 1;
  bool unicode = false;
  bool updating = false;
};

struct TraceValueDialogState {
  std::wstring value_name;
  DWORD type = REG_SZ;
  std::vector<BYTE> data;
  bool accepted = false;
  int dword_base = 16;
  int qword_base = 16;
  BinaryGroupState binary;
  BinaryGroupState none;
  HFONT mono_font = nullptr;
  HFONT ui_font = nullptr;
};

struct CommentDialogState {
  std::wstring text;
  bool apply_all = false;
  HFONT ui_font = nullptr;
  bool accepted = false;
};

struct ExtendedValueDialogState {
  DWORD base_type = REG_SZ;
  std::wstring value_name;
  std::wstring value_type;
  std::wstring initial_text;
  std::vector<BYTE> initial_data;
  std::vector<BYTE> data;
  bool accepted = false;
  int number_base = 16;
  HFONT ui_font = nullptr;
};

bool ParseHexBytes(const std::wstring& text, std::vector<BYTE>* out);
std::wstring FormatBinaryPreview(const std::vector<BYTE>& data, int group_bytes, bool unicode);
std::wstring BinaryToHex(const std::vector<BYTE>& data);
std::wstring RegDataToString(const std::vector<BYTE>& data);
std::vector<BYTE> StringToRegData(const std::wstring& text);
std::vector<BYTE> TextToMultiSz(const std::wstring& text);
bool ParseNumberValue(const std::wstring& text, int base, unsigned long long* value);

constexpr wchar_t kAppTitle[] = L"RegKit";
constexpr UINT_PTR kMultilineEditSubclassId = 1;

LRESULT CALLBACK MultilineEditProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR) {
  switch (msg) {
  case WM_KEYDOWN: {
    if (wparam == VK_RETURN && (GetKeyState(VK_SHIFT) & 0x8000)) {
      LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
      if ((style & ES_MULTILINE) != 0 && (style & ES_READONLY) == 0) {
        SendMessageW(hwnd, EM_REPLACESEL, TRUE, reinterpret_cast<LPARAM>(L"\r\n"));
        return 0;
      }
    }
    break;
  }
  case WM_NCDESTROY:
    RemoveWindowSubclass(hwnd, MultilineEditProc, kMultilineEditSubclassId);
    break;
  default:
    break;
  }
  return DefSubclassProc(hwnd, msg, wparam, lparam);
}

void EnableShiftEnterForMultilineEdits(HWND dlg) {
  EnumChildWindows(
      dlg,
      [](HWND child, LPARAM) -> BOOL {
        wchar_t class_name[16] = {};
        if (!GetClassNameW(child, class_name, static_cast<int>(_countof(class_name)))) {
          return TRUE;
        }
        if (_wcsicmp(class_name, L"Edit") != 0) {
          return TRUE;
        }
        LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
        if ((style & ES_MULTILINE) == 0 || (style & ES_READONLY) != 0) {
          return TRUE;
        }
        SetWindowSubclass(child, MultilineEditProc, kMultilineEditSubclassId, 0);
        return TRUE;
      },
      0);
}

void MoveDialogControl(HWND dlg, int id, int dx, int dy) {
  HWND control = GetDlgItem(dlg, id);
  if (!control) {
    return;
  }
  RECT rect = {};
  GetWindowRect(control, &rect);
  MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&rect), 2);
  SetWindowPos(control, nullptr, rect.left + dx, rect.top + dy, 0, 0, SWP_NOZORDER | SWP_NOSIZE);
}

void CenterDialogToOwner(HWND dlg) {
  if (!dlg) {
    return;
  }
  RECT rect = {};
  if (!GetWindowRect(dlg, &rect)) {
    return;
  }
  int width = rect.right - rect.left;
  int height = rect.bottom - rect.top;
  HWND owner = GetWindow(dlg, GW_OWNER);
  RECT owner_rect = {};
  if (owner && GetWindowRect(owner, &owner_rect)) {
    int owner_w = owner_rect.right - owner_rect.left;
    int owner_h = owner_rect.bottom - owner_rect.top;
    int x = owner_rect.left + std::max(0, (owner_w - width) / 2);
    int y = owner_rect.top + std::max(0, (owner_h - height) / 2);
    SetWindowPos(dlg, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
    return;
  }
  RECT work_area = {};
  if (SystemParametersInfoW(SPI_GETWORKAREA, 0, &work_area, 0)) {
    int work_w = work_area.right - work_area.left;
    int work_h = work_area.bottom - work_area.top;
    int x = work_area.left + std::max(0, (work_w - width) / 2);
    int y = work_area.top + std::max(0, (work_h - height) / 2);
    SetWindowPos(dlg, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
  }
}

void ApplyDialogFonts(HWND hwnd, HFONT font) {
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

void ApplyThinEditBorder(HWND dlg, int id) {
  HWND edit = GetDlgItem(dlg, id);
  if (!edit) {
    return;
  }
  LONG_PTR ex_style = GetWindowLongPtrW(edit, GWL_EXSTYLE);
  if (ex_style & WS_EX_CLIENTEDGE) {
    ex_style &= ~WS_EX_CLIENTEDGE;
    SetWindowLongPtrW(edit, GWL_EXSTYLE, ex_style);
  }
  LONG_PTR style = GetWindowLongPtrW(edit, GWL_STYLE);
  if ((style & WS_BORDER) == 0) {
    style |= WS_BORDER;
    SetWindowLongPtrW(edit, GWL_STYLE, style);
  }
  SetWindowPos(edit, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
}

void ResizeDialogHeight(HWND dlg, int delta) {
  if (delta == 0) {
    return;
  }
  RECT rect = {};
  GetWindowRect(dlg, &rect);
  SetWindowPos(dlg, nullptr, 0, 0, rect.right - rect.left, rect.bottom - rect.top + delta, SWP_NOZORDER | SWP_NOMOVE);
}

void ConfigureValueDetails(HWND dlg, const std::wstring& name, const std::wstring& type, bool show, int hide_offset, const std::vector<int>& move_ids) {
  HWND name_label = GetDlgItem(dlg, IDC_VALUE_NAME_LABEL);
  HWND name_value = GetDlgItem(dlg, IDC_VALUE_NAME);
  HWND type_label = GetDlgItem(dlg, IDC_VALUE_TYPE_LABEL);
  HWND type_value = GetDlgItem(dlg, IDC_VALUE_TYPE);
  HWND bytes_label = GetDlgItem(dlg, IDC_VALUE_BYTES_LABEL);
  HWND bytes_value = GetDlgItem(dlg, IDC_VALUE_BYTES);

  if (show) {
    if (name_label) {
      ShowWindow(name_label, SW_SHOW);
    }
    if (name_value) {
      SetWindowTextW(name_value, name.c_str());
      ShowWindow(name_value, SW_SHOW);
    }
    if (type_label) {
      ShowWindow(type_label, SW_SHOW);
    }
    if (type_value) {
      SetWindowTextW(type_value, type.c_str());
      ShowWindow(type_value, SW_SHOW);
    }
    if (bytes_label) {
      ShowWindow(bytes_label, SW_SHOW);
    }
    if (bytes_value) {
      ShowWindow(bytes_value, SW_SHOW);
    }
    return;
  }

  if (name_label) {
    ShowWindow(name_label, SW_HIDE);
  }
  if (name_value) {
    ShowWindow(name_value, SW_HIDE);
  }
  if (type_label) {
    ShowWindow(type_label, SW_HIDE);
  }
  if (type_value) {
    ShowWindow(type_value, SW_HIDE);
  }
  if (bytes_label) {
    ShowWindow(bytes_label, SW_HIDE);
  }
  if (bytes_value) {
    ShowWindow(bytes_value, SW_HIDE);
  }
  if (hide_offset != 0) {
    for (int id : move_ids) {
      MoveDialogControl(dlg, id, 0, -hide_offset);
    }
    ResizeDialogHeight(dlg, -hide_offset);
  }
}

struct BinaryGroupIds {
  int edit_id = 0;
  int preview_id = 0;
  int format_byte_id = 0;
  int format_word_id = 0;
  int format_dword_id = 0;
  int format_qword_id = 0;
  int text_ansi_id = 0;
  int text_unicode_id = 0;
};

const BinaryGroupIds kBinaryIds = {IDC_REG_BINARY_EDIT, IDC_REG_BINARY_PREVIEW, IDC_REG_BINARY_FORMAT_BYTE, IDC_REG_BINARY_FORMAT_WORD, IDC_REG_BINARY_FORMAT_DWORD, IDC_REG_BINARY_FORMAT_QWORD, IDC_REG_BINARY_TEXT_ANSI, IDC_REG_BINARY_TEXT_UNICODE};
const BinaryGroupIds kNoneIds = {IDC_REG_NONE_EDIT, IDC_REG_NONE_PREVIEW, IDC_REG_NONE_FORMAT_BYTE, IDC_REG_NONE_FORMAT_WORD, IDC_REG_NONE_FORMAT_DWORD, IDC_REG_NONE_FORMAT_QWORD, IDC_REG_NONE_TEXT_ANSI, IDC_REG_NONE_TEXT_UNICODE};

struct TraceTypeEntry {
  DWORD type = REG_SZ;
  const wchar_t* label = nullptr;
};

const TraceTypeEntry kTraceTypes[] = {
    {REG_SZ, L"REG_SZ"}, {REG_EXPAND_SZ, L"REG_EXPAND_SZ"}, {REG_MULTI_SZ, L"REG_MULTI_SZ"}, {REG_LINK, L"REG_LINK"}, {REG_DWORD, L"REG_DWORD"}, {REG_DWORD_BIG_ENDIAN, L"REG_DWORD_BIG_ENDIAN"}, {REG_QWORD, L"REG_QWORD"}, {REG_BINARY, L"REG_BINARY"}, {REG_RESOURCE_LIST, L"REG_RESOURCE_LIST"}, {REG_FULL_RESOURCE_DESCRIPTOR, L"REG_FULL_RESOURCE_DESCRIPTOR"}, {REG_RESOURCE_REQUIREMENTS_LIST, L"REG_RESOURCE_REQUIREMENTS_LIST"}, {REG_NONE, L"REG_NONE"},
};

int TypeToComboIndex(DWORD type) {
  for (size_t i = 0; i < _countof(kTraceTypes); ++i) {
    if (kTraceTypes[i].type == type) {
      return static_cast<int>(i);
    }
  }
  return 0;
}

DWORD ComboIndexToType(int index) {
  if (index < 0 || index >= static_cast<int>(_countof(kTraceTypes))) {
    return REG_SZ;
  }
  return kTraceTypes[index].type;
}

void PopulateTraceTypeCombo(HWND dlg) {
  HWND combo = GetDlgItem(dlg, IDC_TYPE_COMBO);
  if (!combo) {
    return;
  }
  SendMessageW(combo, CB_RESETCONTENT, 0, 0);
  for (const auto& entry : kTraceTypes) {
    int idx = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(entry.label)));
    SendMessageW(combo, CB_SETITEMDATA, idx, static_cast<LPARAM>(entry.type));
  }
}

void SetGroupVisibility(HWND dlg, const int* ids, size_t count, bool visible) {
  if (!dlg || !ids) {
    return;
  }
  int cmd = visible ? SW_SHOW : SW_HIDE;
  for (size_t i = 0; i < count; ++i) {
    HWND ctrl = GetDlgItem(dlg, ids[i]);
    if (ctrl) {
      ShowWindow(ctrl, cmd);
    }
  }
}

const int kRegSzGroupIds[] = {IDC_GROUP_REG_SZ, IDC_REG_SZ_EDIT};
const int kRegExpandGroupIds[] = {IDC_GROUP_REG_EXPAND, IDC_REG_EXPAND_EDIT};
const int kRegMultiGroupIds[] = {IDC_GROUP_REG_MULTI, IDC_REG_MULTI_EDIT};
const int kRegDwordGroupIds[] = {IDC_GROUP_REG_DWORD, IDC_REG_DWORD_EDIT, IDC_REG_DWORD_BASE_GROUP, IDC_REG_DWORD_HEX, IDC_REG_DWORD_DEC, IDC_REG_DWORD_BIN};
const int kRegQwordGroupIds[] = {IDC_GROUP_REG_QWORD, IDC_REG_QWORD_EDIT, IDC_REG_QWORD_BASE_GROUP, IDC_REG_QWORD_HEX, IDC_REG_QWORD_DEC, IDC_REG_QWORD_BIN};
const int kRegBinaryGroupIds[] = {IDC_GROUP_REG_BINARY, IDC_REG_BINARY_LABEL_HEX, IDC_REG_BINARY_EDIT, IDC_REG_BINARY_LABEL_PREVIEW, IDC_REG_BINARY_PREVIEW, IDC_REG_BINARY_FORMAT_GROUP, IDC_REG_BINARY_FORMAT_BYTE, IDC_REG_BINARY_FORMAT_WORD, IDC_REG_BINARY_FORMAT_DWORD, IDC_REG_BINARY_FORMAT_QWORD, IDC_REG_BINARY_TEXT_GROUP, IDC_REG_BINARY_TEXT_ANSI, IDC_REG_BINARY_TEXT_UNICODE};
const int kRegNoneGroupIds[] = {IDC_GROUP_REG_NONE, IDC_REG_NONE_LABEL_HEX, IDC_REG_NONE_EDIT, IDC_REG_NONE_LABEL_PREVIEW, IDC_REG_NONE_PREVIEW, IDC_REG_NONE_FORMAT_GROUP, IDC_REG_NONE_FORMAT_BYTE, IDC_REG_NONE_FORMAT_WORD, IDC_REG_NONE_FORMAT_DWORD, IDC_REG_NONE_FORMAT_QWORD, IDC_REG_NONE_TEXT_GROUP, IDC_REG_NONE_TEXT_ANSI, IDC_REG_NONE_TEXT_UNICODE};

struct TraceTypeGroup {
  DWORD type = REG_SZ;
  const int* ids = nullptr;
  size_t count = 0;
};

const TraceTypeGroup kTraceTypeGroups[] = {
    {REG_SZ, kRegSzGroupIds, _countof(kRegSzGroupIds)}, {REG_EXPAND_SZ, kRegExpandGroupIds, _countof(kRegExpandGroupIds)}, {REG_MULTI_SZ, kRegMultiGroupIds, _countof(kRegMultiGroupIds)}, {REG_LINK, kRegSzGroupIds, _countof(kRegSzGroupIds)}, {REG_DWORD, kRegDwordGroupIds, _countof(kRegDwordGroupIds)}, {REG_DWORD_BIG_ENDIAN, kRegDwordGroupIds, _countof(kRegDwordGroupIds)}, {REG_QWORD, kRegQwordGroupIds, _countof(kRegQwordGroupIds)}, {REG_BINARY, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)}, {REG_RESOURCE_LIST, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)}, {REG_FULL_RESOURCE_DESCRIPTOR, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)}, {REG_RESOURCE_REQUIREMENTS_LIST, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)}, {REG_NONE, kRegNoneGroupIds, _countof(kRegNoneGroupIds)},
};

const wchar_t* TraceTypeLabel(DWORD type) {
  for (const auto& entry : kTraceTypes) {
    if (entry.type == type) {
      return entry.label;
    }
  }
  return L"REG_BINARY";
}

bool IsBinaryGroupType(DWORD type) {
  switch (type) {
  case REG_BINARY:
  case REG_RESOURCE_LIST:
  case REG_FULL_RESOURCE_DESCRIPTOR:
  case REG_RESOURCE_REQUIREMENTS_LIST:
    return true;
  default:
    return false;
  }
}

void ShowTraceTypeGroup(HWND dlg, DWORD type) {
  for (const auto& group : kTraceTypeGroups) {
    SetGroupVisibility(dlg, group.ids, group.count, false);
  }
  if (IsBinaryGroupType(type)) {
    SetGroupVisibility(dlg, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds), true);
    return;
  }
  for (const auto& group : kTraceTypeGroups) {
    if (group.type == type) {
      SetGroupVisibility(dlg, group.ids, group.count, true);
      break;
    }
  }
}

void UpdateTraceGroupLabels(HWND dlg, DWORD type) {
  if (!dlg) {
    return;
  }
  const wchar_t* sz_label = (type == REG_LINK) ? L"REG_LINK" : L"REG_SZ";
  const wchar_t* dword_label = (type == REG_DWORD_BIG_ENDIAN) ? L"REG_DWORD_BIG_ENDIAN" : L"REG_DWORD";
  const wchar_t* raw_label = IsBinaryGroupType(type) ? TraceTypeLabel(type) : L"REG_BINARY";
  SetDlgItemTextW(dlg, IDC_GROUP_REG_SZ, sz_label);
  SetDlgItemTextW(dlg, IDC_GROUP_REG_DWORD, dword_label);
  SetDlgItemTextW(dlg, IDC_GROUP_REG_BINARY, raw_label);
}

void SelectTraceType(HWND dlg, TraceValueDialogState* state, DWORD type) {
  if (!dlg || !state) {
    return;
  }
  state->type = type;
  HWND combo = GetDlgItem(dlg, IDC_TYPE_COMBO);
  if (combo) {
    int index = TypeToComboIndex(type);
    SendMessageW(combo, CB_SETCURSEL, index, 0);
  }
  ShowTraceTypeGroup(dlg, type);
  UpdateTraceGroupLabels(dlg, type);
}

DWORD ReadTraceType(HWND dlg, TraceValueDialogState* state) {
  if (!dlg) {
    return state ? state->type : REG_SZ;
  }
  HWND combo = GetDlgItem(dlg, IDC_TYPE_COMBO);
  if (!combo) {
    return state ? state->type : REG_SZ;
  }
  int index = static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0));
  if (index == CB_ERR) {
    return state ? state->type : REG_SZ;
  }
  return ComboIndexToType(index);
}

std::wstring ReadDialogText(HWND dlg, int id) {
  std::wstring text;
  if (!dlg) {
    return text;
  }
  int length = GetWindowTextLengthW(GetDlgItem(dlg, id));
  if (length <= 0) {
    return text;
  }
  text.resize(static_cast<size_t>(length));
  GetDlgItemTextW(dlg, id, text.data(), length + 1);
  return text;
}

void SetBinaryGroupSelection(HWND dlg, const BinaryGroupIds& ids, int control_id) {
  if (!dlg) {
    return;
  }
  CheckDlgButton(dlg, ids.format_byte_id, control_id == ids.format_byte_id ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, ids.format_word_id, control_id == ids.format_word_id ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, ids.format_dword_id, control_id == ids.format_dword_id ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, ids.format_qword_id, control_id == ids.format_qword_id ? BST_CHECKED : BST_UNCHECKED);
}

void SetBinaryTextSelection(HWND dlg, const BinaryGroupIds& ids, int control_id) {
  if (!dlg) {
    return;
  }
  CheckDlgButton(dlg, ids.text_ansi_id, control_id == ids.text_ansi_id ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, ids.text_unicode_id, control_id == ids.text_unicode_id ? BST_CHECKED : BST_UNCHECKED);
}

void UpdateBinaryPreviewEx(HWND dlg, BinaryGroupState* state, const BinaryGroupIds& ids) {
  if (!dlg || !state || state->updating) {
    return;
  }
  std::wstring text = ReadDialogText(dlg, ids.edit_id);
  std::vector<BYTE> parsed;
  if (!ParseHexBytes(text, &parsed)) {
    SetDlgItemTextW(dlg, ids.preview_id, L"Invalid hex input.");
    return;
  }
  std::wstring preview = FormatBinaryPreview(parsed, state->group_bytes, state->unicode);
  SetDlgItemTextW(dlg, ids.preview_id, preview.c_str());
}

std::wstring FormatNumberValue(unsigned long long value, int base) {
  wchar_t buffer[64] = {};
  if (base == 16) {
    swprintf_s(buffer, L"%llX", value);
    return buffer;
  }
  if (base == 10) {
    swprintf_s(buffer, L"%llu", value);
    return buffer;
  }
  if (base == 2) {
    if (value == 0) {
      return L"0";
    }
    std::wstring out;
    while (value > 0) {
      out.push_back((value & 1ULL) ? L'1' : L'0');
      value >>= 1;
    }
    std::reverse(out.begin(), out.end());
    return out;
  }
  swprintf_s(buffer, L"%llu", value);
  return buffer;
}

unsigned long long ReadUnsignedFromBytes(const std::vector<BYTE>& data, size_t bytes) {
  unsigned long long value = 0;
  if (bytes == 0 || data.size() < bytes) {
    return 0;
  }
  memcpy(&value, data.data(), bytes);
  return value;
}

unsigned long long ReadNumberWithFallback(HWND dlg, int edit_id, int base, unsigned long long fallback) {
  std::wstring text = ReadDialogText(dlg, edit_id);
  unsigned long long parsed = 0;
  if (ParseNumberValue(text, base, &parsed)) {
    return parsed;
  }
  return fallback;
}

bool ParseNumberValue(const std::wstring& text, int base, unsigned long long* value) {
  if (!value) {
    return false;
  }
  if (base != 2 && base != 10 && base != 16) {
    base = 10;
  }
  if (base == 2) {
    const wchar_t* start = text.c_str();
    while (*start && iswspace(*start)) {
      ++start;
    }
    if (start[0] == L'0' && (start[1] == L'b' || start[1] == L'B')) {
      start += 2;
    }
    unsigned long long parsed = 0;
    bool saw_digit = false;
    for (const wchar_t* ptr = start; *ptr; ++ptr) {
      if (*ptr == L'0' || *ptr == L'1') {
        saw_digit = true;
        if (parsed > (std::numeric_limits<unsigned long long>::max() >> 1)) {
          return false;
        }
        parsed = (parsed << 1) | static_cast<unsigned long long>(*ptr - L'0');
        continue;
      }
      if (iswspace(*ptr) || *ptr == L'_' || *ptr == L'\'') {
        continue;
      }
      return false;
    }
    if (!saw_digit) {
      return false;
    }
    *value = parsed;
    return true;
  }
  const wchar_t* start = text.c_str();
  while (*start && iswspace(*start)) {
    ++start;
  }
  if (*start == L'\0') {
    return false;
  }
  errno = 0;
  wchar_t* end = nullptr;
  unsigned long long parsed = wcstoull(start, &end, base);
  if (start == end) {
    return false;
  }
  if (errno == ERANGE) {
    return false;
  }
  while (*end && iswspace(*end)) {
    ++end;
  }
  if (*end != L'\0') {
    return false;
  }
  *value = parsed;
  return true;
}

unsigned long long ReadUnsignedFromBytesBigEndian(const std::vector<BYTE>& data, size_t bytes) {
  if (bytes == 0 || data.size() < bytes) {
    return 0;
  }
  unsigned long long value = 0;
  for (size_t i = 0; i < bytes; ++i) {
    value = (value << 8) | static_cast<unsigned long long>(data[i]);
  }
  return value;
}

void WriteUnsignedToBytesBigEndian(unsigned long long value, size_t bytes, std::vector<BYTE>* out) {
  if (!out || bytes == 0) {
    return;
  }
  out->assign(bytes, 0);
  for (size_t i = 0; i < bytes; ++i) {
    size_t index = bytes - 1 - i;
    (*out)[index] = static_cast<BYTE>(value & 0xFF);
    value >>= 8;
  }
}

std::wstring FormatByteCount(size_t bytes) {
  wchar_t buffer[64] = {};
  swprintf_s(buffer, L"%llu byte%s", static_cast<unsigned long long>(bytes), (bytes == 1) ? L"" : L"s");
  return buffer;
}

void SetBytesLabel(HWND dlg, size_t bytes) {
  std::wstring text = FormatByteCount(bytes);
  SetDlgItemTextW(dlg, IDC_VALUE_BYTES, text.c_str());
}

void SetBytesInvalid(HWND dlg) {
  SetDlgItemTextW(dlg, IDC_VALUE_BYTES, L"Invalid");
}

bool CountHexBytes(const std::wstring& text, size_t* out_bytes) {
  if (!out_bytes) {
    return false;
  }
  size_t count = 0;
  int nibble = -1;
  for (wchar_t ch : text) {
    if (iswxdigit(ch)) {
      if (nibble < 0) {
        nibble = 0;
      } else {
        ++count;
        nibble = -1;
      }
    }
  }
  if (nibble >= 0) {
    return false;
  }
  *out_bytes = count;
  return true;
}

size_t StringByteCount(const std::wstring& text) {
  return (text.size() + 1) * sizeof(wchar_t);
}

size_t MultiSzByteCount(const std::wstring& text) {
  size_t chars = 1;
  size_t current = 0;
  for (wchar_t ch : text) {
    if (ch == L'\r') {
      continue;
    }
    if (ch == L'\n') {
      chars += current + 1;
      current = 0;
      continue;
    }
    ++current;
  }
  chars += current + 1;
  return chars * sizeof(wchar_t);
}

void UpdateBytesFromHexControl(HWND dlg, int id) {
  if (!dlg) {
    return;
  }
  std::wstring text = ReadDialogText(dlg, id);
  size_t bytes = 0;
  if (!CountHexBytes(text, &bytes)) {
    SetBytesInvalid(dlg);
    return;
  }
  SetBytesLabel(dlg, bytes);
}

INT_PTR CALLBACK CustomValueDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<TraceValueDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<TraceValueDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dlg, L"Edit Value");
    PopulateTraceTypeCombo(dlg);
    if (state) {
      std::wstring name = state->value_name.empty() ? L"(Default)" : state->value_name;
      SetDlgItemTextW(dlg, IDC_VALUE_NAME, name.c_str());
      SelectTraceType(dlg, state, state->type);
    } else {
      TraceValueDialogState temp;
      SelectTraceType(dlg, &temp, REG_SZ);
    }

    CheckDlgButton(dlg, IDC_REG_DWORD_HEX, BST_CHECKED);
    CheckDlgButton(dlg, IDC_REG_DWORD_DEC, BST_UNCHECKED);
    CheckDlgButton(dlg, IDC_REG_DWORD_BIN, BST_UNCHECKED);
    CheckDlgButton(dlg, IDC_REG_QWORD_HEX, BST_CHECKED);
    CheckDlgButton(dlg, IDC_REG_QWORD_DEC, BST_UNCHECKED);
    CheckDlgButton(dlg, IDC_REG_QWORD_BIN, BST_UNCHECKED);
    if (state) {
      state->dword_base = 16;
      state->qword_base = 16;
      state->binary.group_bytes = 1;
      state->binary.unicode = false;
      state->none.group_bytes = 1;
      state->none.unicode = false;
    }
    SetBinaryGroupSelection(dlg, kBinaryIds, IDC_REG_BINARY_FORMAT_BYTE);
    SetBinaryTextSelection(dlg, kBinaryIds, IDC_REG_BINARY_TEXT_ANSI);
    SetBinaryGroupSelection(dlg, kNoneIds, IDC_REG_NONE_FORMAT_BYTE);
    SetBinaryTextSelection(dlg, kNoneIds, IDC_REG_NONE_TEXT_ANSI);

    HFONT font = CreateFontW(-12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_MODERN, L"Consolas");
    if (state) {
      state->mono_font = font;
    }
    if (font) {
      SendDlgItemMessageW(dlg, IDC_REG_BINARY_EDIT, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
      SendDlgItemMessageW(dlg, IDC_REG_BINARY_PREVIEW, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
      SendDlgItemMessageW(dlg, IDC_REG_NONE_EDIT, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
      SendDlgItemMessageW(dlg, IDC_REG_NONE_PREVIEW, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
    ApplyThinEditBorder(dlg, IDC_REG_SZ_EDIT);
    ApplyThinEditBorder(dlg, IDC_REG_EXPAND_EDIT);
    ApplyThinEditBorder(dlg, IDC_REG_MULTI_EDIT);
    ApplyThinEditBorder(dlg, IDC_REG_DWORD_EDIT);
    ApplyThinEditBorder(dlg, IDC_REG_QWORD_EDIT);
    ApplyThinEditBorder(dlg, IDC_REG_BINARY_EDIT);
    ApplyThinEditBorder(dlg, IDC_REG_BINARY_PREVIEW);
    ApplyThinEditBorder(dlg, IDC_REG_NONE_EDIT);
    ApplyThinEditBorder(dlg, IDC_REG_NONE_PREVIEW);
    HFONT ui_font = ui::DefaultUIFont();
    if (state) {
      state->ui_font = ui_font;
    }
    ApplyDialogFonts(dlg, ui_font);
    EnableShiftEnterForMultilineEdits(dlg);
    if (!state && ui_font) {
      DeleteObject(ui_font);
    }
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    UpdateBinaryPreviewEx(dlg, state ? &state->binary : nullptr, kBinaryIds);
    UpdateBinaryPreviewEx(dlg, state ? &state->none : nullptr, kNoneIds);
    CenterDialogToOwner(dlg);
    return TRUE;
  }
  case WM_DESTROY: {
    if (state && state->mono_font) {
      DeleteObject(state->mono_font);
      state->mono_font = nullptr;
    }
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return TRUE;
  }
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, dlg, CTLCOLOR_DLG));
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
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_COMMAND: {
    if (!state) {
      return TRUE;
    }
    int id = LOWORD(wparam);
    int code = HIWORD(wparam);
    if (code == CBN_SELCHANGE && id == IDC_TYPE_COMBO) {
      DWORD type = ReadTraceType(dlg, state);
      SelectTraceType(dlg, state, type);
      return TRUE;
    }

    if (code == BN_CLICKED) {
      switch (id) {
      case IDC_REG_DWORD_HEX:
      case IDC_REG_DWORD_DEC:
      case IDC_REG_DWORD_BIN:
        if (state) {
          unsigned long long fallback = ReadUnsignedFromBytes(state->data, sizeof(DWORD));
          unsigned long long value = ReadNumberWithFallback(dlg, IDC_REG_DWORD_EDIT, state->dword_base, fallback);
          if (id == IDC_REG_DWORD_HEX) {
            state->dword_base = 16;
          } else if (id == IDC_REG_DWORD_BIN) {
            state->dword_base = 2;
          } else {
            state->dword_base = 10;
          }
          CheckDlgButton(dlg, IDC_REG_DWORD_HEX, state->dword_base == 16 ? BST_CHECKED : BST_UNCHECKED);
          CheckDlgButton(dlg, IDC_REG_DWORD_DEC, state->dword_base == 10 ? BST_CHECKED : BST_UNCHECKED);
          CheckDlgButton(dlg, IDC_REG_DWORD_BIN, state->dword_base == 2 ? BST_CHECKED : BST_UNCHECKED);
          std::wstring formatted = FormatNumberValue(value, state->dword_base);
          SetDlgItemTextW(dlg, IDC_REG_DWORD_EDIT, formatted.c_str());
          SendDlgItemMessageW(dlg, IDC_REG_DWORD_EDIT, EM_SETSEL, 0, -1);
        }
        return TRUE;
      case IDC_REG_QWORD_HEX:
      case IDC_REG_QWORD_DEC:
      case IDC_REG_QWORD_BIN:
        if (state) {
          int old_base = state->qword_base;
          unsigned long long fallback = ReadUnsignedFromBytes(state->data, sizeof(unsigned long long));
          unsigned long long value = ReadNumberWithFallback(dlg, IDC_REG_QWORD_EDIT, old_base, fallback);
          if (id == IDC_REG_QWORD_HEX) {
            state->qword_base = 16;
          } else if (id == IDC_REG_QWORD_BIN) {
            state->qword_base = 2;
          } else {
            state->qword_base = 10;
          }
          CheckDlgButton(dlg, IDC_REG_QWORD_HEX, state->qword_base == 16 ? BST_CHECKED : BST_UNCHECKED);
          CheckDlgButton(dlg, IDC_REG_QWORD_DEC, state->qword_base == 10 ? BST_CHECKED : BST_UNCHECKED);
          CheckDlgButton(dlg, IDC_REG_QWORD_BIN, state->qword_base == 2 ? BST_CHECKED : BST_UNCHECKED);
          std::wstring formatted = FormatNumberValue(value, state->qword_base);
          SetDlgItemTextW(dlg, IDC_REG_QWORD_EDIT, formatted.c_str());
          SendDlgItemMessageW(dlg, IDC_REG_QWORD_EDIT, EM_SETSEL, 0, -1);
        }
        return TRUE;
      case IDC_REG_BINARY_FORMAT_BYTE:
        state->binary.group_bytes = 1;
        SetBinaryGroupSelection(dlg, kBinaryIds, id);
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        return TRUE;
      case IDC_REG_BINARY_FORMAT_WORD:
        state->binary.group_bytes = 2;
        SetBinaryGroupSelection(dlg, kBinaryIds, id);
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        return TRUE;
      case IDC_REG_BINARY_FORMAT_DWORD:
        state->binary.group_bytes = 4;
        SetBinaryGroupSelection(dlg, kBinaryIds, id);
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        return TRUE;
      case IDC_REG_BINARY_FORMAT_QWORD:
        state->binary.group_bytes = 8;
        SetBinaryGroupSelection(dlg, kBinaryIds, id);
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        return TRUE;
      case IDC_REG_BINARY_TEXT_ANSI:
        state->binary.unicode = false;
        SetBinaryTextSelection(dlg, kBinaryIds, id);
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        return TRUE;
      case IDC_REG_BINARY_TEXT_UNICODE:
        state->binary.unicode = true;
        SetBinaryTextSelection(dlg, kBinaryIds, id);
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        return TRUE;
      case IDC_REG_NONE_FORMAT_BYTE:
        state->none.group_bytes = 1;
        SetBinaryGroupSelection(dlg, kNoneIds, id);
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        return TRUE;
      case IDC_REG_NONE_FORMAT_WORD:
        state->none.group_bytes = 2;
        SetBinaryGroupSelection(dlg, kNoneIds, id);
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        return TRUE;
      case IDC_REG_NONE_FORMAT_DWORD:
        state->none.group_bytes = 4;
        SetBinaryGroupSelection(dlg, kNoneIds, id);
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        return TRUE;
      case IDC_REG_NONE_FORMAT_QWORD:
        state->none.group_bytes = 8;
        SetBinaryGroupSelection(dlg, kNoneIds, id);
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        return TRUE;
      case IDC_REG_NONE_TEXT_ANSI:
        state->none.unicode = false;
        SetBinaryTextSelection(dlg, kNoneIds, id);
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        return TRUE;
      case IDC_REG_NONE_TEXT_UNICODE:
        state->none.unicode = true;
        SetBinaryTextSelection(dlg, kNoneIds, id);
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        return TRUE;
      default:
        break;
      }
    }

    if (code == EN_CHANGE) {
      if (id == IDC_REG_BINARY_EDIT) {
        UpdateBinaryPreviewEx(dlg, &state->binary, kBinaryIds);
        return TRUE;
      }
      if (id == IDC_REG_NONE_EDIT) {
        UpdateBinaryPreviewEx(dlg, &state->none, kNoneIds);
        return TRUE;
      }
    }

    if (id == IDOK) {
      DWORD type = ReadTraceType(dlg, state);
      std::vector<BYTE> data;
      bool ok = true;
      switch (type) {
      case REG_SZ: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_SZ_EDIT);
        data = StringToRegData(text);
        break;
      }
      case REG_EXPAND_SZ: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_EXPAND_EDIT);
        data = StringToRegData(text);
        break;
      }
      case REG_LINK: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_SZ_EDIT);
        data = StringToRegData(text);
        break;
      }
      case REG_MULTI_SZ: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_MULTI_EDIT);
        data = TextToMultiSz(text);
        break;
      }
      case REG_DWORD: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_DWORD_EDIT);
        unsigned long long value = 0;
        if (!ParseNumberValue(text, state->dword_base, &value) || value > std::numeric_limits<DWORD>::max()) {
          ok = false;
        } else {
          data.resize(sizeof(DWORD));
          DWORD v32 = static_cast<DWORD>(value);
          memcpy(data.data(), &v32, sizeof(DWORD));
        }
        break;
      }
      case REG_DWORD_BIG_ENDIAN: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_DWORD_EDIT);
        unsigned long long value = 0;
        if (!ParseNumberValue(text, state->dword_base, &value) || value > std::numeric_limits<DWORD>::max()) {
          ok = false;
        } else {
          WriteUnsignedToBytesBigEndian(value, sizeof(DWORD), &data);
        }
        break;
      }
      case REG_QWORD: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_QWORD_EDIT);
        unsigned long long value = 0;
        if (!ParseNumberValue(text, state->qword_base, &value)) {
          ok = false;
        } else {
          data.resize(sizeof(unsigned long long));
          memcpy(data.data(), &value, sizeof(unsigned long long));
        }
        break;
      }
      case REG_BINARY: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_BINARY_EDIT);
        ok = ParseHexBytes(text, &data);
        break;
      }
      case REG_RESOURCE_LIST:
      case REG_FULL_RESOURCE_DESCRIPTOR:
      case REG_RESOURCE_REQUIREMENTS_LIST: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_BINARY_EDIT);
        ok = ParseHexBytes(text, &data);
        break;
      }
      case REG_NONE: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_NONE_EDIT);
        ok = ParseHexBytes(text, &data);
        break;
      }
      default:
        ok = false;
        break;
      }
      if (!ok) {
        ui::ShowError(dlg, L"Invalid value data.");
        return TRUE;
      }
      state->type = type;
      state->data = std::move(data);
      state->accepted = true;
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    if (id == IDCANCEL) {
      state->accepted = false;
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    }
    break;
  }
  default:
    break;
  }
  return FALSE;
}

struct ExportDialogState {
  std::wstring path;
  bool include_subkeys = true;
  bool open_after = false;
  bool accepted = false;
  HFONT ui_font = nullptr;
};

bool PromptExportPath(HWND owner, std::wstring* path) {
  if (!path) {
    return false;
  }
  wchar_t buffer[MAX_PATH] = {};
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0";
  ofn.lpstrFile = buffer;
  ofn.nMaxFile = static_cast<DWORD>(_countof(buffer));
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
  if (!GetSaveFileNameW(&ofn)) {
    return false;
  }
  *path = buffer;
  return true;
}

INT_PTR CALLBACK ExportDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<ExportDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<ExportDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dlg, kAppTitle);
    if (state) {
      SetDlgItemTextW(dlg, IDC_EXPORT_PATH, state->path.c_str());
      CheckDlgButton(dlg, IDC_EXPORT_RANGE_BRANCH, state->include_subkeys ? BST_CHECKED : BST_UNCHECKED);
      CheckDlgButton(dlg, IDC_EXPORT_RANGE_KEY, state->include_subkeys ? BST_UNCHECKED : BST_CHECKED);
      CheckDlgButton(dlg, IDC_EXPORT_OPEN_AFTER, state->open_after ? BST_CHECKED : BST_UNCHECKED);
    }
    ApplyThinEditBorder(dlg, IDC_EXPORT_PATH);
    HFONT ui_font = ui::DefaultUIFont();
    if (state) {
      state->ui_font = ui_font;
    }
    ApplyDialogFonts(dlg, ui_font);
    EnableShiftEnterForMultilineEdits(dlg);
    if (!state && ui_font) {
      DeleteObject(ui_font);
    }
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    CenterDialogToOwner(dlg);
    return TRUE;
  }
  case WM_DESTROY:
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return TRUE;
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, dlg, CTLCOLOR_DLG));
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
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_COMMAND: {
    if (!state) {
      return TRUE;
    }
    int id = LOWORD(wparam);
    if (id == IDC_EXPORT_BROWSE && HIWORD(wparam) == BN_CLICKED) {
      std::wstring path;
      if (PromptExportPath(dlg, &path)) {
        SetDlgItemTextW(dlg, IDC_EXPORT_PATH, path.c_str());
      }
      return TRUE;
    }
    if (id == IDOK) {
      std::wstring path = ReadDialogText(dlg, IDC_EXPORT_PATH);
      if (path.empty()) {
        ui::ShowError(dlg, L"Select a destination file.");
        return TRUE;
      }
      state->path = path;
      state->include_subkeys = (IsDlgButtonChecked(dlg, IDC_EXPORT_RANGE_BRANCH) == BST_CHECKED);
      state->open_after = (IsDlgButtonChecked(dlg, IDC_EXPORT_OPEN_AFTER) == BST_CHECKED);
      state->accepted = true;
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    if (id == IDCANCEL) {
      state->accepted = false;
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    }
    break;
  }
  default:
    break;
  }
  return FALSE;
}

INT_PTR CALLBACK CommentDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<CommentDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<CommentDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dlg, L"Edit Comment");
    if (state) {
      SetDlgItemTextW(dlg, IDC_EDIT, state->text.c_str());
      CheckDlgButton(dlg, IDC_COMMENT_ALL, state->apply_all ? BST_CHECKED : BST_UNCHECKED);
    }
    ApplyThinEditBorder(dlg, IDC_EDIT);
    HFONT ui_font = ui::DefaultUIFont();
    if (state) {
      state->ui_font = ui_font;
    }
    ApplyDialogFonts(dlg, ui_font);
    EnableShiftEnterForMultilineEdits(dlg);
    if (!state && ui_font) {
      DeleteObject(ui_font);
    }
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    CenterDialogToOwner(dlg);
    return TRUE;
  }
  case WM_DESTROY:
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return TRUE;
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, dlg, CTLCOLOR_DLG));
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
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_COMMAND: {
    if (!state) {
      return TRUE;
    }
    int id = LOWORD(wparam);
    if (id == IDOK) {
      state->text = ReadDialogText(dlg, IDC_EDIT);
      state->apply_all = (IsDlgButtonChecked(dlg, IDC_COMMENT_ALL) == BST_CHECKED);
      state->accepted = true;
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    if (id == IDCANCEL) {
      state->accepted = false;
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    }
    break;
  }
  default:
    break;
  }
  return FALSE;
}

INT_PTR CALLBACK TextDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  TextDialogState* state = reinterpret_cast<TextDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<TextDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    if (state && state->title && *state->title) {
      SetWindowTextW(dlg, state->title);
    } else {
      SetWindowTextW(dlg, L"Edit Value");
    }
    if (state->label) {
      SetDlgItemTextW(dlg, IDC_LABEL, state->label);
    }
    ConfigureValueDetails(dlg, state->value_name, state->value_type, state->show_details, 20, {IDC_LABEL, IDC_EDIT, IDOK, IDCANCEL, IDC_NOTE});
    SetDlgItemTextW(dlg, IDC_EDIT, state->text.c_str());
    ApplyThinEditBorder(dlg, IDC_EDIT);
    SendDlgItemMessageW(dlg, IDC_EDIT, EM_SETSEL, 0, -1);
    if (state->title && wcscmp(state->title, L"Connect to Remote Registry") == 0) {
      RECT client = {};
      GetClientRect(dlg, &client);
      HWND label = GetDlgItem(dlg, IDC_LABEL);
      int shift = 0;
      if (label) {
        RECT rect = {};
        GetWindowRect(label, &rect);
        MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&rect), 2);
        const int desired_top = 9;
        shift = desired_top - rect.top;
      }
      if (shift != 0) {
        MoveDialogControl(dlg, IDC_LABEL, 0, shift);
        MoveDialogControl(dlg, IDC_EDIT, 0, shift);
        MoveDialogControl(dlg, IDOK, 0, shift);
        MoveDialogControl(dlg, IDCANCEL, 0, shift);
        ResizeDialogHeight(dlg, shift);
      }
      if (label) {
        RECT rect = {};
        GetWindowRect(label, &rect);
        MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&rect), 2);
        int width = std::max(0, static_cast<int>(client.right - client.left) - static_cast<int>(rect.left) - 8);
        SetWindowPos(label, nullptr, rect.left, rect.top, width, rect.bottom - rect.top, SWP_NOZORDER);
      }
    }
    state->ui_font = ui::DefaultUIFont();
    ApplyDialogFonts(dlg, state->ui_font);
    EnableShiftEnterForMultilineEdits(dlg);
    CenterDialogToOwner(dlg);
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    return TRUE;
  }
  case WM_DESTROY:
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return TRUE;
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, dlg, CTLCOLOR_DLG));
  }
  case WM_CTLCOLORSTATIC: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_STATIC));
  }
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_EDIT));
  }
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_BTN));
  }
  case WM_COMMAND: {
    switch (LOWORD(wparam)) {
    case IDOK: {
      wchar_t buffer[4096] = {};
      GetDlgItemTextW(dlg, IDC_EDIT, buffer, static_cast<int>(_countof(buffer)));
      if (state) {
        state->text = buffer;
      }
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    case IDCANCEL:
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    default:
      break;
    }
    break;
  }
  default:
    break;
  }
  return FALSE;
}

INT_PTR CALLBACK NumberDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  NumberDialogState* state = reinterpret_cast<NumberDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<NumberDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    if (state && state->title && *state->title) {
      SetWindowTextW(dlg, state->title);
    } else {
      SetWindowTextW(dlg, L"Edit Value");
    }
    if (state->label) {
      SetDlgItemTextW(dlg, IDC_LABEL, state->label);
    }
    ConfigureValueDetails(dlg, state->value_name, state->value_type, state->show_details, 20, {IDC_LABEL, IDC_EDIT, IDC_BASE_GROUP, IDC_HEX, IDC_DEC, IDC_BIN, IDOK, IDCANCEL});
    std::wstring formatted = FormatNumberValue(state->value, state->base);
    SetDlgItemTextW(dlg, IDC_EDIT, formatted.c_str());
    ApplyThinEditBorder(dlg, IDC_EDIT);
    CheckDlgButton(dlg, IDC_HEX, state->base == 16 ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dlg, IDC_DEC, state->base == 10 ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(dlg, IDC_BIN, state->base == 2 ? BST_CHECKED : BST_UNCHECKED);
    SendDlgItemMessageW(dlg, IDC_EDIT, EM_SETSEL, 0, -1);
    state->ui_font = ui::DefaultUIFont();
    ApplyDialogFonts(dlg, state->ui_font);
    EnableShiftEnterForMultilineEdits(dlg);
    CenterDialogToOwner(dlg);
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    return TRUE;
  }
  case WM_DESTROY:
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return TRUE;
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, dlg, CTLCOLOR_DLG));
  }
  case WM_CTLCOLORSTATIC: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_STATIC));
  }
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_EDIT));
  }
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_BTN));
  }
  case WM_COMMAND: {
    switch (LOWORD(wparam)) {
    case IDC_HEX:
    case IDC_DEC:
    case IDC_BIN: {
      if (!state) {
        return TRUE;
      }
      unsigned long long value = ReadNumberWithFallback(dlg, IDC_EDIT, state->base, state->value);
      int new_base = 10;
      if (LOWORD(wparam) == IDC_HEX) {
        new_base = 16;
      } else if (LOWORD(wparam) == IDC_BIN) {
        new_base = 2;
      }
      state->base = new_base;
      CheckDlgButton(dlg, IDC_HEX, state->base == 16 ? BST_CHECKED : BST_UNCHECKED);
      CheckDlgButton(dlg, IDC_DEC, state->base == 10 ? BST_CHECKED : BST_UNCHECKED);
      CheckDlgButton(dlg, IDC_BIN, state->base == 2 ? BST_CHECKED : BST_UNCHECKED);
      state->value = value;
      std::wstring formatted = FormatNumberValue(state->value, state->base);
      SetDlgItemTextW(dlg, IDC_EDIT, formatted.c_str());
      SendDlgItemMessageW(dlg, IDC_EDIT, EM_SETSEL, 0, -1);
      return TRUE;
    }
    case IDOK: {
      wchar_t buffer[128] = {};
      GetDlgItemTextW(dlg, IDC_EDIT, buffer, static_cast<int>(_countof(buffer)));
      unsigned long long value = 0;
      int base = 10;
      if (IsDlgButtonChecked(dlg, IDC_HEX) == BST_CHECKED) {
        base = 16;
      } else if (IsDlgButtonChecked(dlg, IDC_BIN) == BST_CHECKED) {
        base = 2;
      }
      if (!ParseNumberValue(buffer, base, &value)) {
        ui::ShowError(dlg, L"Invalid number.");
        return TRUE;
      }
      if (state) {
        state->base = base;
        state->value = value;
      }
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    case IDCANCEL:
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    default:
      break;
    }
    break;
  }
  default:
    break;
  }
  return FALSE;
}

int FormatHexOffset(std::wstring* out, size_t offset, int width) {
  if (!out) {
    return 0;
  }
  wchar_t buffer[16] = {};
  if (width < 4) {
    width = 4;
  }
  if (width > 8) {
    width = 8;
  }
  swprintf_s(buffer, L"%0*X", width, static_cast<unsigned int>(offset));
  out->append(buffer);
  return width;
}

std::wstring FormatBinaryPreview(const std::vector<BYTE>& data, int group_bytes, bool unicode) {
  std::wstring out;
  if (group_bytes <= 0) {
    group_bytes = 1;
  }
  const int bytes_per_line = 16;
  size_t size = data.size();
  int offset_width = (size > 0xFFFF) ? 8 : 4;
  size_t index = 0;
  while (index < size) {
    FormatHexOffset(&out, index, offset_width);
    out.append(L"  ");

    for (int i = 0; i < bytes_per_line; ++i) {
      size_t pos = index + static_cast<size_t>(i);
      if (pos < size) {
        wchar_t hex[4] = {};
        swprintf_s(hex, L"%02X", static_cast<unsigned int>(data[pos]));
        out.append(hex);
      } else {
        out.append(L"  ");
      }
      out.push_back(L' ');
      if (i == 7) {
        out.push_back(L' ');
      }
      if (group_bytes > 1 && ((i + 1) % group_bytes == 0)) {
        out.push_back(L' ');
      }
    }

    out.push_back(L' ');
    if (unicode) {
      for (int i = 0; i < bytes_per_line; i += 2) {
        size_t pos = index + static_cast<size_t>(i);
        wchar_t ch = L'.';
        if (pos + 1 < size) {
          wchar_t value = static_cast<wchar_t>(data[pos] | (data[pos + 1] << 8));
          if (iswprint(value)) {
            ch = value;
          }
        }
        out.push_back(ch);
      }
    } else {
      for (int i = 0; i < bytes_per_line; ++i) {
        size_t pos = index + static_cast<size_t>(i);
        wchar_t ch = L' ';
        if (pos < size) {
          wchar_t value = static_cast<wchar_t>(data[pos]);
          ch = iswprint(value) ? value : L'.';
        }
        out.push_back(ch);
      }
    }
    out.append(L"\r\n");
    index += bytes_per_line;
  }
  if (out.empty()) {
    out = L"";
  }
  return out;
}

void SetBinaryGroup(HWND dlg, int control_id) {
  if (!dlg) {
    return;
  }
  CheckDlgButton(dlg, IDC_FORMAT_BYTE, control_id == IDC_FORMAT_BYTE ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, IDC_FORMAT_WORD, control_id == IDC_FORMAT_WORD ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, IDC_FORMAT_DWORD, control_id == IDC_FORMAT_DWORD ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, IDC_FORMAT_QWORD, control_id == IDC_FORMAT_QWORD ? BST_CHECKED : BST_UNCHECKED);
}

void SetBinaryTextMode(HWND dlg, int control_id) {
  if (!dlg) {
    return;
  }
  CheckDlgButton(dlg, IDC_TEXT_ANSI, control_id == IDC_TEXT_ANSI ? BST_CHECKED : BST_UNCHECKED);
  CheckDlgButton(dlg, IDC_TEXT_UNICODE, control_id == IDC_TEXT_UNICODE ? BST_CHECKED : BST_UNCHECKED);
}

void UpdateBinaryPreview(HWND dlg, BinaryDialogState* state) {
  if (!dlg || !state || state->updating) {
    return;
  }
  int length = GetWindowTextLengthW(GetDlgItem(dlg, IDC_EDIT));
  std::wstring text;
  text.resize(static_cast<size_t>(std::max(0, length)));
  if (length > 0) {
    GetDlgItemTextW(dlg, IDC_EDIT, text.data(), length + 1);
  }
  state->text = text;

  std::vector<BYTE> parsed;
  if (!ParseHexBytes(text, &parsed)) {
    SetDlgItemTextW(dlg, IDC_BINARY_PREVIEW, L"Invalid hex input.");
    return;
  }
  std::wstring preview = FormatBinaryPreview(parsed, state->group_bytes, state->unicode);
  SetDlgItemTextW(dlg, IDC_BINARY_PREVIEW, preview.c_str());
}

INT_PTR CALLBACK BinaryDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  BinaryDialogState* state = reinterpret_cast<BinaryDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<BinaryDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dlg, L"Edit Value");
    SetDlgItemTextW(dlg, IDC_LABEL, L"Hex bytes:");
    SetDlgItemTextW(dlg, IDC_NOTE, L"Preview:");
    ConfigureValueDetails(dlg, state->value_name, state->value_type, state->show_details, 20, {IDC_LABEL, IDC_EDIT, IDC_NOTE, IDC_BINARY_PREVIEW, IDC_FORMAT_GROUP, IDC_FORMAT_BYTE, IDC_FORMAT_WORD, IDC_FORMAT_DWORD, IDC_FORMAT_QWORD, IDC_TEXT_GROUP, IDC_TEXT_ANSI, IDC_TEXT_UNICODE, IDOK, IDCANCEL});
    SetDlgItemTextW(dlg, IDC_EDIT, state->text.c_str());
    ApplyThinEditBorder(dlg, IDC_EDIT);

    state->group_bytes = 1;
    state->unicode = false;
    SetBinaryGroup(dlg, IDC_FORMAT_BYTE);
    SetBinaryTextMode(dlg, IDC_TEXT_ANSI);

    HFONT font = CreateFontW(-12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_MODERN, L"Consolas");
    state->mono_font = font;
    if (font) {
      SendDlgItemMessageW(dlg, IDC_EDIT, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
      SendDlgItemMessageW(dlg, IDC_BINARY_PREVIEW, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
    ApplyThinEditBorder(dlg, IDC_EDIT);
    ApplyThinEditBorder(dlg, IDC_BINARY_PREVIEW);

    state->ui_font = ui::DefaultUIFont();
    ApplyDialogFonts(dlg, state->ui_font);
    EnableShiftEnterForMultilineEdits(dlg);
    CenterDialogToOwner(dlg);
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    UpdateBinaryPreview(dlg, state);
    UpdateBytesFromHexControl(dlg, IDC_EDIT);
    return TRUE;
  }
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, dlg, CTLCOLOR_DLG));
  }
  case WM_CTLCOLORSTATIC: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_STATIC));
  }
  case WM_CTLCOLOREDIT: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_EDIT));
  }
  case WM_CTLCOLORBTN: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    HWND target = reinterpret_cast<HWND>(lparam);
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, CTLCOLOR_BTN));
  }
  case WM_COMMAND: {
    switch (LOWORD(wparam)) {
    case IDC_EDIT:
      if (HIWORD(wparam) == EN_CHANGE) {
        UpdateBinaryPreview(dlg, state);
        UpdateBytesFromHexControl(dlg, IDC_EDIT);
      }
      break;
    case IDC_FORMAT_BYTE:
      state->group_bytes = 1;
      SetBinaryGroup(dlg, IDC_FORMAT_BYTE);
      UpdateBinaryPreview(dlg, state);
      return TRUE;
    case IDC_FORMAT_WORD:
      state->group_bytes = 2;
      SetBinaryGroup(dlg, IDC_FORMAT_WORD);
      UpdateBinaryPreview(dlg, state);
      return TRUE;
    case IDC_FORMAT_DWORD:
      state->group_bytes = 4;
      SetBinaryGroup(dlg, IDC_FORMAT_DWORD);
      UpdateBinaryPreview(dlg, state);
      return TRUE;
    case IDC_FORMAT_QWORD:
      state->group_bytes = 8;
      SetBinaryGroup(dlg, IDC_FORMAT_QWORD);
      UpdateBinaryPreview(dlg, state);
      return TRUE;
    case IDC_TEXT_ANSI:
      state->unicode = false;
      SetBinaryTextMode(dlg, IDC_TEXT_ANSI);
      UpdateBinaryPreview(dlg, state);
      return TRUE;
    case IDC_TEXT_UNICODE:
      state->unicode = true;
      SetBinaryTextMode(dlg, IDC_TEXT_UNICODE);
      UpdateBinaryPreview(dlg, state);
      return TRUE;
    case IDOK: {
      wchar_t buffer[8192] = {};
      GetDlgItemTextW(dlg, IDC_EDIT, buffer, static_cast<int>(_countof(buffer)));
      state->text = buffer;
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    case IDCANCEL:
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    default:
      break;
    }
    break;
  }
  case WM_DESTROY:
    if (state && state->mono_font) {
      DeleteObject(state->mono_font);
      state->mono_font = nullptr;
    }
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    break;
  default:
    break;
  }
  return FALSE;
}

INT_PTR CALLBACK ExtendedValueDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<ExtendedValueDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
  switch (msg) {
  case WM_INITDIALOG: {
    state = reinterpret_cast<ExtendedValueDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dlg, L"Edit Value");
    if (state) {
      SetDlgItemTextW(dlg, IDC_VALUE_NAME, state->value_name.c_str());
      SetDlgItemTextW(dlg, IDC_VALUE_TYPE, state->value_type.c_str());
      SetDlgItemTextW(dlg, IDC_EDIT, state->initial_text.c_str());
    }

    ApplyThinEditBorder(dlg, IDC_EDIT);

    if (state && (state->base_type == REG_DWORD || state->base_type == REG_DWORD_BIG_ENDIAN || state->base_type == REG_QWORD)) {
      CheckDlgButton(dlg, IDC_HEX, state->number_base == 16 ? BST_CHECKED : BST_UNCHECKED);
      CheckDlgButton(dlg, IDC_DEC, state->number_base == 10 ? BST_CHECKED : BST_UNCHECKED);
      CheckDlgButton(dlg, IDC_BIN, state->number_base == 2 ? BST_CHECKED : BST_UNCHECKED);
    }
    HFONT ui_font = ui::DefaultUIFont();
    if (state) {
      state->ui_font = ui_font;
    }
    ApplyDialogFonts(dlg, ui_font);
    EnableShiftEnterForMultilineEdits(dlg);
    if (!state && ui_font) {
      DeleteObject(ui_font);
    }
    Theme::Current().ApplyToWindow(dlg);
    Theme::Current().ApplyToChildren(dlg);
    CenterDialogToOwner(dlg);
    return TRUE;
  }
  case WM_DESTROY:
    if (state && state->ui_font) {
      DeleteObject(state->ui_font);
      state->ui_font = nullptr;
    }
    return TRUE;
  case WM_SETTINGCHANGE: {
    if (Theme::UpdateFromSystem()) {
      Theme::Current().ApplyToWindow(dlg);
      Theme::Current().ApplyToChildren(dlg);
      InvalidateRect(dlg, nullptr, TRUE);
    }
    return TRUE;
  }
  case WM_ERASEBKGND: {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(dlg, &rect);
    FillRect(hdc, &rect, Theme::Current().BackgroundBrush());
    return TRUE;
  }
  case WM_CTLCOLORDLG:
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
    } else if (msg == WM_CTLCOLORDLG) {
      type = CTLCOLOR_DLG;
    }
    return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, type));
  }
  case WM_COMMAND: {
    if (!state) {
      return TRUE;
    }
    int id = LOWORD(wparam);
    int code = HIWORD(wparam);

    if (code == BN_CLICKED) {
      switch (id) {
      case IDC_HEX:
      case IDC_DEC:
      case IDC_BIN:
        if (state) {
          unsigned long long fallback = 0;
          if (state->base_type == REG_DWORD) {
            fallback = ReadUnsignedFromBytes(state->initial_data, sizeof(DWORD));
          } else if (state->base_type == REG_DWORD_BIG_ENDIAN) {
            fallback = ReadUnsignedFromBytesBigEndian(state->initial_data, sizeof(DWORD));
          } else if (state->base_type == REG_QWORD) {
            fallback = ReadUnsignedFromBytes(state->initial_data, sizeof(unsigned long long));
          }
          unsigned long long value = ReadNumberWithFallback(dlg, IDC_EDIT, state->number_base, fallback);
          if (id == IDC_HEX) {
            state->number_base = 16;
          } else if (id == IDC_BIN) {
            state->number_base = 2;
          } else {
            state->number_base = 10;
          }
          CheckDlgButton(dlg, IDC_HEX, state->number_base == 16 ? BST_CHECKED : BST_UNCHECKED);
          CheckDlgButton(dlg, IDC_DEC, state->number_base == 10 ? BST_CHECKED : BST_UNCHECKED);
          CheckDlgButton(dlg, IDC_BIN, state->number_base == 2 ? BST_CHECKED : BST_UNCHECKED);
          std::wstring formatted = FormatNumberValue(value, state->number_base);
          SetDlgItemTextW(dlg, IDC_EDIT, formatted.c_str());
          SendDlgItemMessageW(dlg, IDC_EDIT, EM_SETSEL, 0, -1);
        }
        return TRUE;
      default:
        break;
      }
    }

    if (id == IDOK) {
      std::wstring base_text = ReadDialogText(dlg, IDC_EDIT);
      if (base_text == state->initial_text) {
        state->data = state->initial_data;
      } else {
        std::vector<BYTE> base_data;
        switch (state->base_type) {
        case REG_SZ:
        case REG_EXPAND_SZ:
        case REG_LINK:
          base_data = StringToRegData(base_text);
          break;
        case REG_MULTI_SZ:
          base_data = TextToMultiSz(base_text);
          break;
        case REG_DWORD:
        case REG_DWORD_BIG_ENDIAN:
        case REG_QWORD: {
          unsigned long long value = 0;
          if (!ParseNumberValue(base_text, state->number_base, &value)) {
            ui::ShowError(dlg, L"Invalid number.");
            return TRUE;
          }
          if ((state->base_type == REG_DWORD || state->base_type == REG_DWORD_BIG_ENDIAN) && value > std::numeric_limits<DWORD>::max()) {
            ui::ShowError(dlg, L"Number is out of range.");
            return TRUE;
          }
          if (state->base_type == REG_DWORD) {
            DWORD v32 = static_cast<DWORD>(value);
            base_data.resize(sizeof(DWORD));
            memcpy(base_data.data(), &v32, sizeof(DWORD));
          } else if (state->base_type == REG_DWORD_BIG_ENDIAN) {
            WriteUnsignedToBytesBigEndian(value, sizeof(DWORD), &base_data);
          } else {
            base_data.resize(sizeof(unsigned long long));
            memcpy(base_data.data(), &value, sizeof(unsigned long long));
          }
          break;
        }
        default:
          ui::ShowError(dlg, L"Invalid value data.");
          return TRUE;
        }
        state->data = std::move(base_data);
      }
      state->accepted = true;
      EndDialog(dlg, IDOK);
      return TRUE;
    }
    if (id == IDCANCEL) {
      state->accepted = false;
      EndDialog(dlg, IDCANCEL);
      return TRUE;
    }
    break;
  }
  default:
    break;
  }
  return FALSE;
}

std::wstring BinaryToHex(const std::vector<BYTE>& data) {
  if (data.empty()) {
    return L"";
  }
  std::wstringstream stream;
  stream << std::hex << std::uppercase;
  for (size_t i = 0; i < data.size(); ++i) {
    stream.width(2);
    stream.fill(L'0');
    stream << static_cast<int>(data[i]);
    if (i + 1 < data.size()) {
      stream << L' ';
    }
  }
  return stream.str();
}

bool ParseHexBytes(const std::wstring& text, std::vector<BYTE>* out) {
  if (!out) {
    return false;
  }
  out->clear();
  int nibble = -1;
  for (wchar_t ch : text) {
    if (iswxdigit(ch)) {
      int value = 0;
      if (ch >= L'0' && ch <= L'9') {
        value = ch - L'0';
      } else if (ch >= L'a' && ch <= L'f') {
        value = 10 + (ch - L'a');
      } else if (ch >= L'A' && ch <= L'F') {
        value = 10 + (ch - L'A');
      }
      if (nibble < 0) {
        nibble = value;
      } else {
        out->push_back(static_cast<BYTE>((nibble << 4) | value));
        nibble = -1;
      }
    }
  }
  return nibble < 0;
}

std::wstring MultiSzToText(const std::vector<BYTE>& data) {
  if (data.empty()) {
    return L"";
  }
  const wchar_t* ptr = reinterpret_cast<const wchar_t*>(data.data());
  size_t count = data.size() / sizeof(wchar_t);
  std::wstring text;
  size_t offset = 0;
  while (offset < count) {
    const wchar_t* current = ptr + offset;
    size_t len = wcsnlen_s(current, count - offset);
    if (len == 0) {
      break;
    }
    if (!text.empty()) {
      text.append(L"\r\n");
    }
    text.append(current, len);
    offset += len + 1;
  }
  return text;
}

std::vector<BYTE> TextToMultiSz(const std::wstring& text) {
  std::vector<BYTE> data;
  std::wstring normalized = text;
  normalized.erase(std::remove(normalized.begin(), normalized.end(), L'\r'), normalized.end());

  std::vector<std::wstring> parts;
  std::wstring current;
  for (wchar_t ch : normalized) {
    if (ch == L'\n') {
      parts.push_back(current);
      current.clear();
    } else {
      current.push_back(ch);
    }
  }
  parts.push_back(current);

  size_t total_chars = 1;
  for (const auto& part : parts) {
    total_chars += part.size() + 1;
  }
  data.resize(total_chars * sizeof(wchar_t));
  wchar_t* out = reinterpret_cast<wchar_t*>(data.data());
  size_t offset = 0;
  for (const auto& part : parts) {
    wcsncpy_s(out + offset, total_chars - offset, part.c_str(), part.size());
    offset += part.size();
    out[offset++] = L'\0';
  }
  out[offset++] = L'\0';
  data.resize(offset * sizeof(wchar_t));
  return data;
}

std::vector<BYTE> StringToRegData(const std::wstring& text) {
  std::vector<BYTE> data((text.size() + 1) * sizeof(wchar_t));
  memcpy(data.data(), text.c_str(), data.size());
  return data;
}

std::wstring RegDataToString(const std::vector<BYTE>& data) {
  if (data.empty()) {
    return L"";
  }
  size_t wchar_count = data.size() / sizeof(wchar_t);
  std::wstring text(reinterpret_cast<const wchar_t*>(data.data()), wchar_count);
  while (!text.empty() && text.back() == L'\0') {
    text.pop_back();
  }
  return text;
}

} // namespace

bool PromptForValueText(HWND owner, const std::wstring& value_name, const wchar_t* title, const wchar_t* label, std::wstring* text, const wchar_t* value_type) {
  if (!text) {
    return false;
  }
  TextDialogState state;
  state.title = title;
  state.label = label;
  state.text = *text;
  state.value_name = value_name;
  if (value_type) {
    state.value_type = value_type;
  }
  if (value_type && state.value_name.empty()) {
    state.value_name = L"(Default)";
  }
  state.show_details = value_type != nullptr;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_INPUT), owner, TextDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result == IDOK) {
    *text = state.text;
    return true;
  }
  return false;
}

bool PromptForBinary(HWND owner, const std::wstring& value_name, const std::vector<BYTE>& initial, std::vector<BYTE>* out, const wchar_t* value_type) {
  if (!out) {
    return false;
  }
  BinaryDialogState state;
  state.text = BinaryToHex(initial);
  state.value_name = value_name.empty() ? L"(Default)" : value_name;
  state.value_type = value_type ? value_type : L"REG_BINARY";
  state.show_details = true;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_BINARY), owner, BinaryDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result != IDOK) {
    return false;
  }
  std::vector<BYTE> parsed;
  if (!ParseHexBytes(state.text, &parsed)) {
    ui::ShowError(owner, L"Invalid hex input.");
    return false;
  }
  *out = std::move(parsed);
  return true;
}

bool PromptForNumber(HWND owner, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& initial, std::vector<BYTE>* out) {
  if (!out) {
    return false;
  }
  NumberDialogState state;
  state.title = L"Edit Value";
  state.label = L"Value:";
  state.base = 16;
  state.value_name = value_name.empty() ? L"(Default)" : value_name;
  if (type == REG_QWORD) {
    state.value_type = L"REG_QWORD";
  } else if (type == REG_DWORD_BIG_ENDIAN) {
    state.value_type = L"REG_DWORD_BIG_ENDIAN";
  } else {
    state.value_type = L"REG_DWORD";
  }
  state.show_details = true;
  if (type == REG_QWORD && initial.size() >= sizeof(unsigned long long)) {
    state.value = *reinterpret_cast<const unsigned long long*>(initial.data());
  } else if (type == REG_DWORD_BIG_ENDIAN && initial.size() >= sizeof(DWORD)) {
    state.value = ReadUnsignedFromBytesBigEndian(initial, sizeof(DWORD));
  } else if (type == REG_DWORD && initial.size() >= sizeof(DWORD)) {
    state.value = *reinterpret_cast<const DWORD*>(initial.data());
  }
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_NUMBER), owner, NumberDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result != IDOK) {
    return false;
  }
  if (type == REG_QWORD) {
    unsigned long long value = state.value;
    out->resize(sizeof(unsigned long long));
    memcpy(out->data(), &value, sizeof(unsigned long long));
  } else if (type == REG_DWORD_BIG_ENDIAN) {
    DWORD value = static_cast<DWORD>(state.value);
    WriteUnsignedToBytesBigEndian(value, sizeof(DWORD), out);
  } else {
    DWORD value = static_cast<DWORD>(state.value);
    out->resize(sizeof(DWORD));
    memcpy(out->data(), &value, sizeof(DWORD));
  }
  return true;
}

bool PromptForMultiString(HWND owner, const std::wstring& value_name, const std::vector<BYTE>& initial, std::vector<BYTE>* out) {
  if (!out) {
    return false;
  }
  TextDialogState state;
  state.title = L"Edit Value";
  state.label = L"Strings (one per line):";
  state.text = MultiSzToText(initial);
  state.value_name = value_name.empty() ? L"(Default)" : value_name;
  state.value_type = L"REG_MULTI_SZ";
  state.show_details = true;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_MULTI), owner, TextDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result != IDOK) {
    return false;
  }
  *out = TextToMultiSz(state.text);
  return true;
}

bool PromptForMultiLineText(HWND owner, const wchar_t* title, const wchar_t* label, std::wstring* text) {
  if (!text) {
    return false;
  }
  TextDialogState state;
  state.title = title;
  state.label = label;
  state.text = *text;
  state.show_details = false;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_MULTI_TEXT), owner, TextDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result == IDOK) {
    *text = state.text;
    return true;
  }
  return false;
}

bool PromptForComment(HWND owner, const std::wstring& initial, bool apply_all, std::wstring* out_text, bool* out_apply_all) {
  if (!out_text || !out_apply_all) {
    return false;
  }
  CommentDialogState state;
  state.text = initial;
  state.apply_all = apply_all;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_COMMENT), owner, CommentDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result == IDOK && state.accepted) {
    *out_text = state.text;
    *out_apply_all = state.apply_all;
    return true;
  }
  return false;
}

bool PromptForCustomValue(HWND owner, const std::wstring& value_name, DWORD* type, std::vector<BYTE>* out) {
  if (!type || !out) {
    return false;
  }
  TraceValueDialogState state;
  state.value_name = value_name;
  state.type = *type;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_CUSTOM_VALUE), owner, CustomValueDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result != IDOK || !state.accepted) {
    return false;
  }
  *type = state.type;
  *out = std::move(state.data);
  return true;
}

bool PromptForFlaggedValue(HWND owner, const std::wstring& value_name, DWORD base_type, const std::vector<BYTE>& initial, const std::wstring& value_type, std::vector<BYTE>* out) {
  if (!out) {
    return false;
  }
  ExtendedValueDialogState state;
  state.base_type = base_type;
  state.value_name = value_name.empty() ? L"(Default)" : value_name;
  state.value_type = value_type;
  state.initial_data = initial;
  state.number_base = 16;

  switch (base_type) {
  case REG_MULTI_SZ:
    state.initial_text = MultiSzToText(initial);
    break;
  case REG_DWORD: {
    DWORD value = 0;
    if (initial.size() >= sizeof(DWORD)) {
      value = *reinterpret_cast<const DWORD*>(initial.data());
    }
    state.initial_text = FormatNumberValue(value, state.number_base);
    break;
  }
  case REG_DWORD_BIG_ENDIAN: {
    DWORD value = 0;
    if (initial.size() >= sizeof(DWORD)) {
      value = static_cast<DWORD>(ReadUnsignedFromBytesBigEndian(initial, sizeof(DWORD)));
    }
    state.initial_text = FormatNumberValue(value, state.number_base);
    break;
  }
  case REG_QWORD: {
    unsigned long long value = 0;
    if (initial.size() >= sizeof(unsigned long long)) {
      value = *reinterpret_cast<const unsigned long long*>(initial.data());
    }
    state.initial_text = FormatNumberValue(value, state.number_base);
    break;
  }
  default:
    state.initial_text = RegDataToString(initial);
    break;
  }

  int dialog_id = IDD_INPUT_BINARY;
  if (base_type == REG_MULTI_SZ) {
    dialog_id = IDD_MULTI_BINARY;
  } else if (base_type == REG_DWORD || base_type == REG_DWORD_BIG_ENDIAN || base_type == REG_QWORD) {
    dialog_id = IDD_NUMBER_BINARY;
  }

  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(dialog_id), owner, ExtendedValueDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result != IDOK || !state.accepted) {
    return false;
  }
  *out = std::move(state.data);
  return true;
}

bool PromptForExportOptions(HWND owner, const std::wstring& default_path, ExportOptions* options) {
  if (!options) {
    return false;
  }
  ExportDialogState state;
  state.path = default_path;
  state.include_subkeys = options->include_subkeys;
  state.open_after = options->open_after;
  INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_EXPORT_OPTIONS), owner, ExportDialogProc, reinterpret_cast<LPARAM>(&state));
  if (result != IDOK || !state.accepted) {
    return false;
  }
  options->path = std::move(state.path);
  options->include_subkeys = state.include_subkeys;
  options->open_after = state.open_after;
  return true;
}

} // namespace regkit
