// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "editors/value_editor.h"
#include "win32/text_transform.h"

#include "editors/binary_text.h"
#include "appearance/dialog_layout.h"
#include "editors/dialog_support.h"

#include <algorithm>
#include <cerrno>
#include <commctrl.h>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <limits>

#include "appearance/feedback.h"
#include "registry/value_format.h"
#include "resource.h"

namespace regkit::editors {

namespace {

bool IsMultilineEdit(HWND dialog, int id) {
  HWND edit = GetDlgItem(dialog, id);
  return edit && (GetWindowLongPtrW(edit, GWL_STYLE) & ES_MULTILINE) != 0;
}
struct TextDialogState {
  const wchar_t* title = nullptr;
  const wchar_t* label = nullptr;
  std::wstring text;
  std::wstring value_name;
  std::wstring value_type;
  bool show_details = false;
  HFONT ui_font = nullptr;
  appearance::DialogResizer resizer;
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
  appearance::DialogResizer resizer;
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
  int initial_number_base = 16;
  HFONT ui_font = nullptr;
  appearance::DialogResizer resizer;
};

std::wstring RegDataToString(const std::vector<BYTE>& data);
bool ParseNumberValue(const std::wstring& text, int base, unsigned long long* value);

constexpr wchar_t kAppTitle[] = L"RegKit";
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

void ResizeDialogHeight(HWND dlg, int delta) {
  if (delta == 0) {
    return;
  }
  RECT rect = {};
  GetWindowRect(dlg, &rect);
  SetWindowPos(dlg, nullptr, 0, 0, rect.right - rect.left, rect.bottom - rect.top + delta, SWP_NOZORDER | SWP_NOMOVE);
}

void ConfigureReadOnlyNameField(HWND dlg, const std::wstring& name) {
  HWND name_value = GetDlgItem(dlg, IDC_VALUE_NAME);
  if (!name_value) {
    return;
  }
  SetWindowTextW(name_value, name.c_str());
  ShowWindow(name_value, SW_SHOW);
  SendMessageW(name_value, EM_SETREADONLY, TRUE, 0);
  LONG_PTR style = GetWindowLongPtrW(name_value, GWL_STYLE);
  if (style & WS_TABSTOP) {
    style &= ~WS_TABSTOP;
    SetWindowLongPtrW(name_value, GWL_STYLE, style);
  }
}

void ConfigureValueDetails(HWND dlg, const std::wstring& name, const std::wstring& type, bool show, bool show_type, int hide_offset, const std::vector<int>& move_ids) {
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
      ConfigureReadOnlyNameField(dlg, name);
    }
    if (type_label && show_type) {
      ShowWindow(type_label, SW_SHOW);
    } else if (type_label) {
      ShowWindow(type_label, SW_HIDE);
    }
    if (type_value && show_type) {
      SetWindowTextW(type_value, type.c_str());
      ShowWindow(type_value, SW_SHOW);
    } else if (type_value) {
      ShowWindow(type_value, SW_HIDE);
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
    {REG_SZ, L"REG_SZ"},
    {REG_EXPAND_SZ, L"REG_EXPAND_SZ"},
    {REG_MULTI_SZ, L"REG_MULTI_SZ"},
    {REG_LINK, L"REG_LINK"},
    {REG_DWORD, L"REG_DWORD"},
    {REG_DWORD_BIG_ENDIAN, L"REG_DWORD_BIG_ENDIAN"},
    {REG_QWORD, L"REG_QWORD"},
    {REG_BINARY, L"REG_BINARY"},
    {REG_RESOURCE_LIST, L"REG_RESOURCE_LIST"},
    {REG_FULL_RESOURCE_DESCRIPTOR, L"REG_FULL_RESOURCE_DESCRIPTOR"},
    {REG_RESOURCE_REQUIREMENTS_LIST, L"REG_RESOURCE_REQUIREMENTS_LIST"},
    {REG_NONE, L"REG_NONE"},
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
    {REG_SZ, kRegSzGroupIds, _countof(kRegSzGroupIds)},
    {REG_EXPAND_SZ, kRegExpandGroupIds, _countof(kRegExpandGroupIds)},
    {REG_MULTI_SZ, kRegMultiGroupIds, _countof(kRegMultiGroupIds)},
    {REG_LINK, kRegSzGroupIds, _countof(kRegSzGroupIds)},
    {REG_DWORD, kRegDwordGroupIds, _countof(kRegDwordGroupIds)},
    {REG_DWORD_BIG_ENDIAN, kRegDwordGroupIds, _countof(kRegDwordGroupIds)},
    {REG_QWORD, kRegQwordGroupIds, _countof(kRegQwordGroupIds)},
    {REG_BINARY, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_RESOURCE_LIST, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_FULL_RESOURCE_DESCRIPTOR, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_RESOURCE_REQUIREMENTS_LIST, kRegBinaryGroupIds, _countof(kRegBinaryGroupIds)},
    {REG_NONE, kRegNoneGroupIds, _countof(kRegNoneGroupIds)},
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
  return dlg ? util::WindowText(GetDlgItem(dlg, id)) : std::wstring();
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
  if (!value_format::ParseHex(text, &parsed)) {
    SetDlgItemTextW(dlg, ids.preview_id, L"Invalid hex input.");
    return;
  }
  std::wstring preview = binary_text::Preview(
      parsed, state->group_bytes, state->unicode);
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
  if (msg != WM_INITDIALOG && msg != WM_DESTROY) {
    INT_PTR themed = 0;
    if (dialog_support::HandleThemeMessage(
            dlg, msg, wparam, lparam, &themed)) {
      return themed;
    }
  }
  switch (msg) {
  case WM_SIZE:
    if (state) {
      state->resizer.Apply(dlg);
    }
    return TRUE;
  case WM_GETMINMAXINFO:
    if (state) {
      state->resizer.ClampMinSize(reinterpret_cast<MINMAXINFO*>(lparam));
    }
    return TRUE;
  case WM_INITDIALOG: {
    state = reinterpret_cast<TraceValueDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dlg, L"Edit Value");
    PopulateTraceTypeCombo(dlg);
    if (state) {
      std::wstring name = state->value_name.empty() ? L"(Default)" : state->value_name;
      ConfigureReadOnlyNameField(dlg, name);
      SelectTraceType(dlg, state, state->type);
    } else {
      ConfigureReadOnlyNameField(dlg, L"");
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

    if (!state) {
      return FALSE;
    }
    dialog_support::Initialize(
        dlg, &state->ui_font,
        {IDC_VALUE_NAME, IDC_REG_SZ_EDIT, IDC_REG_EXPAND_EDIT,
         IDC_REG_MULTI_EDIT, IDC_REG_DWORD_EDIT, IDC_REG_QWORD_EDIT,
         IDC_REG_BINARY_EDIT, IDC_REG_BINARY_PREVIEW, IDC_REG_NONE_EDIT,
         IDC_REG_NONE_PREVIEW});
    using namespace appearance;
    state->resizer.Attach(dlg, {
        {IDC_VALUE_NAME, kAnchorLeft | kAnchorTop | kAnchorRight},
        {IDC_GROUP_REG_SZ, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_REG_SZ_EDIT, kAnchorLeft | kAnchorTop | kAnchorRight},
        {IDC_GROUP_REG_EXPAND, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_REG_EXPAND_EDIT, kAnchorLeft | kAnchorTop | kAnchorRight},
        {IDC_GROUP_REG_MULTI, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_REG_MULTI_EDIT, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_GROUP_REG_DWORD, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_GROUP_REG_QWORD, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_GROUP_REG_BINARY, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_REG_BINARY_LABEL_HEX, kAnchorLeft | kAnchorTop},
        {IDC_REG_BINARY_EDIT, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_REG_BINARY_LABEL_PREVIEW, kAnchorLeft | kAnchorBottom},
        {IDC_REG_BINARY_PREVIEW, kAnchorLeft | kAnchorRight | kAnchorBottom},
        {IDC_REG_BINARY_FORMAT_GROUP, kAnchorLeft | kAnchorBottom},
        {IDC_REG_BINARY_FORMAT_BYTE, kAnchorLeft | kAnchorBottom},
        {IDC_REG_BINARY_FORMAT_WORD, kAnchorLeft | kAnchorBottom},
        {IDC_REG_BINARY_FORMAT_DWORD, kAnchorLeft | kAnchorBottom},
        {IDC_REG_BINARY_FORMAT_QWORD, kAnchorLeft | kAnchorBottom},
        {IDC_REG_BINARY_TEXT_GROUP, kAnchorRight | kAnchorBottom},
        {IDC_REG_BINARY_TEXT_ANSI, kAnchorRight | kAnchorBottom},
        {IDC_REG_BINARY_TEXT_UNICODE, kAnchorRight | kAnchorBottom},
        {IDC_GROUP_REG_NONE, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_REG_NONE_LABEL_HEX, kAnchorLeft | kAnchorTop},
        {IDC_REG_NONE_EDIT, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
        {IDC_REG_NONE_LABEL_PREVIEW, kAnchorLeft | kAnchorBottom},
        {IDC_REG_NONE_PREVIEW, kAnchorLeft | kAnchorRight | kAnchorBottom},
        {IDC_REG_NONE_FORMAT_GROUP, kAnchorLeft | kAnchorBottom},
        {IDC_REG_NONE_FORMAT_BYTE, kAnchorLeft | kAnchorBottom},
        {IDC_REG_NONE_FORMAT_WORD, kAnchorLeft | kAnchorBottom},
        {IDC_REG_NONE_FORMAT_DWORD, kAnchorLeft | kAnchorBottom},
        {IDC_REG_NONE_FORMAT_QWORD, kAnchorLeft | kAnchorBottom},
        {IDC_REG_NONE_TEXT_GROUP, kAnchorRight | kAnchorBottom},
        {IDC_REG_NONE_TEXT_ANSI, kAnchorRight | kAnchorBottom},
        {IDC_REG_NONE_TEXT_UNICODE, kAnchorRight | kAnchorBottom},
        {IDOK, kAnchorRight | kAnchorBottom},
        {IDCANCEL, kAnchorRight | kAnchorBottom},
    });
    state->mono_font = CreateFontW(
        -12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_MODERN,
        L"Consolas");
    if (state->mono_font) {
      SendDlgItemMessageW(dlg, IDC_REG_BINARY_EDIT, WM_SETFONT,
                          reinterpret_cast<WPARAM>(state->mono_font), TRUE);
      SendDlgItemMessageW(dlg, IDC_REG_BINARY_PREVIEW, WM_SETFONT,
                          reinterpret_cast<WPARAM>(state->mono_font), TRUE);
      SendDlgItemMessageW(dlg, IDC_REG_NONE_EDIT, WM_SETFONT,
                          reinterpret_cast<WPARAM>(state->mono_font), TRUE);
      SendDlgItemMessageW(dlg, IDC_REG_NONE_PREVIEW, WM_SETFONT,
                          reinterpret_cast<WPARAM>(state->mono_font), TRUE);
    }
    UpdateBinaryPreviewEx(dlg, state ? &state->binary : nullptr, kBinaryIds);
    UpdateBinaryPreviewEx(dlg, state ? &state->none : nullptr, kNoneIds);
    return TRUE;
  }
  case WM_DESTROY: {
    if (state) {
      dialog_support::ReleaseFont(&state->mono_font);
      dialog_support::ReleaseFont(&state->ui_font);
    }
    return TRUE;
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
        data = value_format::StringData(text);
        break;
      }
      case REG_EXPAND_SZ: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_EXPAND_EDIT);
        data = value_format::StringData(text);
        break;
      }
      case REG_LINK: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_SZ_EDIT);
        data = value_format::StringData(text);
        break;
      }
      case REG_MULTI_SZ: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_MULTI_EDIT);
        data = value_format::MultiStringData(text);
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
        ok = value_format::ParseHex(text, &data);
        break;
      }
      case REG_RESOURCE_LIST:
      case REG_FULL_RESOURCE_DESCRIPTOR:
      case REG_RESOURCE_REQUIREMENTS_LIST: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_BINARY_EDIT);
        ok = value_format::ParseHex(text, &data);
        break;
      }
      case REG_NONE: {
        std::wstring text = ReadDialogText(dlg, IDC_REG_NONE_EDIT);
        ok = value_format::ParseHex(text, &data);
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

INT_PTR CALLBACK TextDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  TextDialogState* state = reinterpret_cast<TextDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
  if (msg != WM_INITDIALOG && msg != WM_DESTROY) {
    INT_PTR themed = 0;
    if (dialog_support::HandleThemeMessage(
            dlg, msg, wparam, lparam, &themed)) {
      return themed;
    }
  }

  switch (msg) {
  case WM_SIZE:
    if (state) {
      state->resizer.Apply(dlg);
    }
    return TRUE;
  case WM_GETMINMAXINFO:
    if (state) {
      state->resizer.ClampMinSize(reinterpret_cast<MINMAXINFO*>(lparam));
    }
    return TRUE;
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
    ConfigureValueDetails(dlg, state->value_name, state->value_type, state->show_details, false, 20, {IDC_LABEL, IDC_EDIT, IDOK, IDCANCEL, IDC_NOTE});
    SetDlgItemTextW(dlg, IDC_EDIT, state->text.c_str());
    SendDlgItemMessageW(dlg, IDC_EDIT, EM_SETSEL, 0, -1);
    const bool is_remote_registry = (state->title && wcscmp(state->title, L"Connect to Remote Registry") == 0);
    if (is_remote_registry) {
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
    dialog_support::Initialize(
        dlg, &state->ui_font, {IDC_VALUE_NAME, IDC_EDIT});
    if (IsMultilineEdit(dlg, IDC_EDIT)) {
      using namespace appearance;
      state->resizer.Attach(dlg, {
          {IDC_VALUE_NAME, kAnchorLeft | kAnchorTop | kAnchorRight},
          {IDC_VALUE_TYPE, kAnchorLeft | kAnchorTop | kAnchorRight},
          {IDC_LABEL, kAnchorLeft | kAnchorTop | kAnchorRight},
          {IDC_EDIT, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
          {IDOK, kAnchorRight | kAnchorBottom},
          {IDCANCEL, kAnchorRight | kAnchorBottom},
      });
    }
    return TRUE;
  }
  case WM_DESTROY:
    if (state) {
      dialog_support::ReleaseFont(&state->ui_font);
    }
    return TRUE;
  case WM_COMMAND: {
    switch (LOWORD(wparam)) {
    case IDOK: {
      if (state) {
        state->text = dialog_support::ReadText(dlg, IDC_EDIT);
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

INT_PTR CALLBACK ExtendedValueDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<ExtendedValueDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
  if (msg != WM_INITDIALOG && msg != WM_DESTROY) {
    INT_PTR themed = 0;
    if (dialog_support::HandleThemeMessage(
            dlg, msg, wparam, lparam, &themed)) {
      return themed;
    }
  }
  switch (msg) {
  case WM_SIZE:
    if (state) {
      state->resizer.Apply(dlg);
    }
    return TRUE;
  case WM_GETMINMAXINFO:
    if (state) {
      state->resizer.ClampMinSize(reinterpret_cast<MINMAXINFO*>(lparam));
    }
    return TRUE;
  case WM_INITDIALOG: {
    state = reinterpret_cast<ExtendedValueDialogState*>(lparam);
    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SetWindowTextW(dlg, L"Edit Value");
    std::wstring value_name;
    std::wstring value_type;
    if (state) {
      value_name = state->value_name;
      value_type = state->value_type;
      SetDlgItemTextW(dlg, IDC_EDIT, state->initial_text.c_str());
    }
    ConfigureValueDetails(dlg, value_name, value_type, true, false, 20, {IDC_LABEL, IDC_EDIT, IDC_BASE_GROUP, IDC_HEX, IDC_DEC, IDC_BIN, IDOK, IDCANCEL});

    if (state && (state->base_type == REG_DWORD || state->base_type == REG_DWORD_BIG_ENDIAN || state->base_type == REG_QWORD)) {
      CheckDlgButton(dlg, IDC_HEX, state->number_base == 16 ? BST_CHECKED : BST_UNCHECKED);
      CheckDlgButton(dlg, IDC_DEC, state->number_base == 10 ? BST_CHECKED : BST_UNCHECKED);
      CheckDlgButton(dlg, IDC_BIN, state->number_base == 2 ? BST_CHECKED : BST_UNCHECKED);
    }
    if (!state) {
      return FALSE;
    }
    dialog_support::Initialize(
        dlg, &state->ui_font, {IDC_VALUE_NAME, IDC_EDIT});
    if (state->base_type == REG_MULTI_SZ) {
      dialog_support::AllowNewlines(dlg, IDC_EDIT);
    }
    if (IsMultilineEdit(dlg, IDC_EDIT)) {
      using namespace appearance;
      state->resizer.Attach(dlg, {
          {IDC_VALUE_NAME, kAnchorLeft | kAnchorTop | kAnchorRight},
          {IDC_VALUE_TYPE, kAnchorLeft | kAnchorTop | kAnchorRight},
          {IDC_LABEL, kAnchorLeft | kAnchorTop | kAnchorRight},
          {IDC_EDIT, kAnchorLeft | kAnchorTop | kAnchorRight | kAnchorBottom},
          {IDOK, kAnchorRight | kAnchorBottom},
          {IDCANCEL, kAnchorRight | kAnchorBottom},
      });
    }
    return TRUE;
  }
  case WM_DESTROY:
    if (state) {
      dialog_support::ReleaseFont(&state->ui_font);
    }
    return TRUE;
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
      const bool is_number = state->base_type == REG_DWORD ||
                             state->base_type == REG_DWORD_BIG_ENDIAN ||
                             state->base_type == REG_QWORD;
      const bool unchanged = base_text == state->initial_text &&
                             (!is_number || state->number_base == state->initial_number_base);
      if (unchanged) {
        state->data = state->initial_data;
      } else {
        std::vector<BYTE> base_data;
        switch (state->base_type) {
        case REG_SZ:
        case REG_EXPAND_SZ:
        case REG_LINK:
          base_data = value_format::StringData(base_text);
          break;
        case REG_MULTI_SZ:
          base_data = value_format::MultiStringData(base_text);
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

bool EditText(HWND owner, const TextRequest& request, TextResult* result) {
  if (!result) {
    return false;
  }
  TextDialogState state;
  state.title = request.title.c_str();
  state.label = request.label.c_str();
  state.text = request.text;
  state.value_name = request.value_name;
  state.value_type = request.value_type;
  if (request.show_value_details && state.value_name.empty()) {
    state.value_name = L"(Default)";
  }
  state.show_details = request.show_value_details;
  const int dialog_id = request.multiline ? IDD_MULTI_TEXT : IDD_INPUT;
  const INT_PTR dialog_result = DialogBoxParamW(
      GetModuleHandleW(nullptr), MAKEINTRESOURCEW(dialog_id), owner,
      TextDialogProc, reinterpret_cast<LPARAM>(&state));
  if (dialog_result != IDOK) {
    return false;
  }
  result->text = std::move(state.text);
  return true;
}

bool EditCustomValue(HWND owner, const CustomValueRequest& request,
                     CustomValueResult* result) {
  if (!result) {
    return false;
  }
  TraceValueDialogState state;
  state.value_name = request.value_name;
  state.type = request.type;
  const INT_PTR dialog_result = DialogBoxParamW(
      GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_CUSTOM_VALUE), owner,
      CustomValueDialogProc, reinterpret_cast<LPARAM>(&state));
  if (dialog_result != IDOK || !state.accepted) {
    return false;
  }
  result->type = state.type;
  result->data = std::move(state.data);
  return true;
}

bool EditFlaggedValue(HWND owner, const FlaggedValueRequest& request,
                      FlaggedValueResult* result) {
  if (!result) {
    return false;
  }
  ExtendedValueDialogState state;
  state.base_type = request.base_type;
  state.value_name =
      request.value_name.empty() ? L"(Default)" : request.value_name;
  state.value_type = request.value_type;
  state.initial_data.assign(request.data.begin(), request.data.end());
  state.number_base = 16;
  state.initial_number_base = state.number_base;

  switch (request.base_type) {
  case REG_MULTI_SZ:
    state.initial_text = value_format::MultiStringText(state.initial_data);
    break;
  case REG_DWORD: {
    DWORD value = 0;
    if (request.data.size() >= sizeof(DWORD)) {
      memcpy(&value, request.data.data(), sizeof(value));
    }
    state.initial_text = FormatNumberValue(value, state.number_base);
    break;
  }
  case REG_DWORD_BIG_ENDIAN: {
    DWORD value = 0;
    if (request.data.size() >= sizeof(DWORD)) {
      value = static_cast<DWORD>(
          ReadUnsignedFromBytesBigEndian(state.initial_data, sizeof(DWORD)));
    }
    state.initial_text = FormatNumberValue(value, state.number_base);
    break;
  }
  case REG_QWORD: {
    unsigned long long value = 0;
    if (request.data.size() >= sizeof(unsigned long long)) {
      memcpy(&value, request.data.data(), sizeof(value));
    }
    state.initial_text = FormatNumberValue(value, state.number_base);
    break;
  }
  default:
    state.initial_text = RegDataToString(state.initial_data);
    break;
  }

  int dialog_id = IDD_INPUT_BINARY;
  if (request.base_type == REG_MULTI_SZ) {
    dialog_id = IDD_MULTI_BINARY;
  } else if (request.base_type == REG_DWORD ||
             request.base_type == REG_DWORD_BIG_ENDIAN ||
             request.base_type == REG_QWORD) {
    dialog_id = IDD_NUMBER_BINARY;
  }

  const INT_PTR dialog_result = DialogBoxParamW(
      GetModuleHandleW(nullptr), MAKEINTRESOURCEW(dialog_id), owner,
      ExtendedValueDialogProc, reinterpret_cast<LPARAM>(&state));
  if (dialog_result != IDOK || !state.accepted) {
    return false;
  }
  result->data = std::move(state.data);
  return true;
}

} // namespace regkit::editors
