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

#include "app/app_window.h"

#include <algorithm>
#include <commctrl.h>
#include <commdlg.h>
#include <uxtheme.h>
#include <vsstyle.h>

#include "app/ui_helpers.h"
#include "app/value_dialogs.h"

namespace regkit {

namespace {

constexpr wchar_t kThemePresetClass[] = L"RegKitThemePresetsWindow";
constexpr wchar_t kThemePresetTitle[] = L"Theme Presets";

constexpr int kWindowWidth = 560;
constexpr int kWindowHeight = 360;
constexpr int kPadding = 12;
constexpr int kGap = 8;
constexpr int kButtonHeight = 22;
constexpr int kBottomButtonHeight = 24;
constexpr int kLeftPanelWidth = 190;
constexpr int kGroupBoxCaptionHeight = 18;
constexpr int kGroupBoxPadding = 10;
constexpr UINT_PTR kThemePresetHeaderSubclassId = 1;
constexpr UINT_PTR kThemePresetListViewSubclassId = 2;

enum ControlId {
  kPresetListId = 5001,
  kColorListId = 5002,
  kNewPresetId = 5003,
  kDuplicatePresetId = 5004,
  kRenamePresetId = 5005,
  kDeletePresetId = 5006,
  kImportPresetId = 5007,
  kExportPresetId = 5008,
  kEditColorId = 5009,
  kDarkCheckId = 5010,
  kTemplateComboId = 5011,
  kApplyTemplateId = 5012,
  kApplyId = 5013,
};

struct ColorField {
  const wchar_t* label = nullptr;
  COLORREF ThemeColors::* member = nullptr;
};

constexpr ColorField kColorFields[] = {
    {L"Background", &ThemeColors::background}, {L"Panel", &ThemeColors::panel}, {L"Surface", &ThemeColors::surface}, {L"Header", &ThemeColors::header}, {L"Border", &ThemeColors::border}, {L"Text", &ThemeColors::text}, {L"Muted Text", &ThemeColors::muted_text}, {L"Accent", &ThemeColors::accent}, {L"Selection", &ThemeColors::selection}, {L"Selection Text", &ThemeColors::selection_text}, {L"Hover", &ThemeColors::hover}, {L"Focus", &ThemeColors::focus},
};

struct ThemePresetWindowState {
  HWND hwnd = nullptr;
  HWND owner = nullptr;
  HWND presets_group = nullptr;
  HWND preset_list = nullptr;
  HWND colors_group = nullptr;
  HWND color_list = nullptr;
  HWND templates_group = nullptr;
  HWND new_btn = nullptr;
  HWND duplicate_btn = nullptr;
  HWND rename_btn = nullptr;
  HWND delete_btn = nullptr;
  HWND import_btn = nullptr;
  HWND export_btn = nullptr;
  HWND edit_color_btn = nullptr;
  HWND dark_check = nullptr;
  HWND template_combo = nullptr;
  HWND template_btn = nullptr;
  HWND apply_btn = nullptr;
  HWND ok_btn = nullptr;
  HWND cancel_btn = nullptr;
  HFONT font = nullptr;
  MainWindow* window = nullptr;
  std::vector<ThemePreset> presets;
  std::vector<ThemePreset> templates;
  std::wstring active_name;
  int selected_index = -1;
  int color_sort_column = -1;
  bool color_sort_ascending = true;
  COLORREF custom_colors[16] = {};
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

void ApplyFontRecursive(HWND hwnd, HFONT font) {
  if (!hwnd || !font) {
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
  if (!hwnd) {
    return;
  }
  RECT rect = {};
  if (owner && GetWindowRect(owner, &rect)) {
    int owner_w = rect.right - rect.left;
    int owner_h = rect.bottom - rect.top;
    int x = rect.left + std::max(0, (owner_w - width) / 2);
    int y = rect.top + std::max(0, (owner_h - height) / 2);
    SetWindowPos(hwnd, nullptr, x, y, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
    return;
  }
  SetWindowPos(hwnd, nullptr, CW_USEDEFAULT, CW_USEDEFAULT, width, height, SWP_NOZORDER | SWP_NOACTIVATE);
}

ThemePreset* CurrentPreset(ThemePresetWindowState* state) {
  if (!state) {
    return nullptr;
  }
  if (state->selected_index < 0 || static_cast<size_t>(state->selected_index) >= state->presets.size()) {
    return nullptr;
  }
  return &state->presets[static_cast<size_t>(state->selected_index)];
}

int FindPresetIndexByName(const std::vector<ThemePreset>& presets, const std::wstring& name) {
  for (size_t i = 0; i < presets.size(); ++i) {
    if (_wcsicmp(presets[i].name.c_str(), name.c_str()) == 0) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

int GetSelectedPresetIndex(HWND list) {
  if (!list) {
    return -1;
  }
  int row = ListView_GetNextItem(list, -1, LVNI_SELECTED);
  if (row < 0) {
    return -1;
  }
  LVITEMW item = {};
  item.mask = LVIF_PARAM;
  item.iItem = row;
  if (!ListView_GetItem(list, &item)) {
    return -1;
  }
  return static_cast<int>(item.lParam);
}

std::wstring MakeUniquePresetName(const std::vector<ThemePreset>& presets, const std::wstring& base_name) {
  std::wstring base = base_name.empty() ? L"Preset" : base_name;
  auto exists = [&](const std::wstring& name) -> bool { return std::any_of(presets.begin(), presets.end(), [&](const ThemePreset& preset) { return _wcsicmp(preset.name.c_str(), name.c_str()) == 0; }); };
  if (!exists(base)) {
    return base;
  }
  for (int i = 2; i < 1000; ++i) {
    std::wstring candidate = base + L" " + std::to_wstring(i);
    if (!exists(candidate)) {
      return candidate;
    }
  }
  return base + L" Copy";
}

bool PromptPresetName(HWND owner, const wchar_t* title, const std::wstring& initial, std::wstring* out_name) {
  if (!out_name) {
    return false;
  }
  std::wstring name = initial;
  if (!PromptForValueText(owner, L"", title, L"Preset name:", &name, nullptr)) {
    return false;
  }
  if (name.empty()) {
    ui::ShowError(owner, L"Preset name cannot be empty.");
    return false;
  }
  *out_name = name;
  return true;
}

bool PromptOpenThemeFile(HWND owner, std::wstring* path) {
  if (!path) {
    return false;
  }
  wchar_t buffer[MAX_PATH] = {};
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = L"RegKit Theme Presets (*.rktheme)\0*.rktheme\0All Files (*.*)\0*.*\0\0";
  ofn.lpstrFile = buffer;
  ofn.nMaxFile = static_cast<DWORD>(_countof(buffer));
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
  if (!GetOpenFileNameW(&ofn)) {
    return false;
  }
  *path = buffer;
  return true;
}

bool PromptSaveThemeFile(HWND owner, std::wstring* path) {
  if (!path) {
    return false;
  }
  wchar_t buffer[MAX_PATH] = {};
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = L"RegKit Theme Presets (*.rktheme)\0*.rktheme\0All Files (*.*)\0*.*\0\0";
  ofn.lpstrFile = buffer;
  ofn.nMaxFile = static_cast<DWORD>(_countof(buffer));
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
  if (!GetSaveFileNameW(&ofn)) {
    return false;
  }
  *path = buffer;
  return true;
}

bool ChooseColorFor(HWND owner, COLORREF* color, COLORREF* custom_colors) {
  if (!color) {
    return false;
  }
  CHOOSECOLORW cc = {};
  cc.lStructSize = sizeof(cc);
  cc.hwndOwner = owner;
  cc.rgbResult = *color;
  cc.lpCustColors = custom_colors;
  cc.Flags = CC_FULLOPEN | CC_RGBINIT;
  if (!ChooseColorW(&cc)) {
    return false;
  }
  *color = cc.rgbResult;
  return true;
}

int CompareTextInsensitive(const wchar_t* left, const wchar_t* right) {
  const wchar_t* safe_left = left ? left : L"";
  const wchar_t* safe_right = right ? right : L"";
  return _wcsicmp(safe_left, safe_right);
}

int CompareColorValue(COLORREF left, COLORREF right) {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
}

int GetListViewColumnSubItem(HWND list, int display_index) {
  if (!list || display_index < 0) {
    return -1;
  }
  LVCOLUMNW col = {};
  col.mask = LVCF_SUBITEM;
  if (!ListView_GetColumn(list, display_index, &col)) {
    return -1;
  }
  return col.iSubItem;
}

void UpdateListViewSort(HWND list, int column, bool ascending) {
  if (!list) {
    return;
  }
  HWND header = ListView_GetHeader(list);
  if (!header) {
    return;
  }
  int count = Header_GetItemCount(header);
  for (int i = 0; i < count; ++i) {
    HDITEMW item = {};
    item.mask = HDI_FORMAT;
    if (!Header_GetItem(header, i, &item)) {
      continue;
    }
    item.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
    if (column >= 0 && GetListViewColumnSubItem(list, i) == column) {
      item.fmt |= ascending ? HDF_SORTUP : HDF_SORTDOWN;
    }
    Header_SetItem(header, i, &item);
  }
}

LRESULT CALLBACK ThemePresetHeaderProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR) {
  if (message == WM_ERASEBKGND) {
    return 1;
  }
  if (message == WM_PAINT) {
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd, &ps);
    const Theme& theme = Theme::Current();
    RECT client = {};
    GetClientRect(hwnd, &client);
    FillRect(hdc, &client, theme.HeaderBrush());

    HFONT font = reinterpret_cast<HFONT>(SendMessageW(hwnd, WM_GETFONT, 0, 0));
    HFONT old_font = nullptr;
    if (font) {
      old_font = reinterpret_cast<HFONT>(SelectObject(hdc, font));
    }

    HTHEME header_theme = OpenThemeData(hwnd, VSCLASS_HEADER);
    SIZE arrow_size = {0, 0};
    if (header_theme) {
      GetThemePartSize(header_theme, hdc, HP_HEADERSORTARROW, HSAS_SORTEDUP, nullptr, TS_TRUE, &arrow_size);
    }
    if (arrow_size.cx <= 0 || arrow_size.cy <= 0) {
      arrow_size.cx = 8;
      arrow_size.cy = 8;
    }

    int count = Header_GetItemCount(hwnd);
    for (int i = 0; i < count; ++i) {
      RECT rect = {};
      if (!Header_GetItemRect(hwnd, i, &rect)) {
        continue;
      }

      wchar_t text[128] = {};
      HDITEMW item = {};
      item.mask = HDI_TEXT | HDI_FORMAT;
      item.pszText = text;
      item.cchTextMax = static_cast<int>(_countof(text));
      Header_GetItem(hwnd, i, &item);

      bool sorted_up = (item.fmt & HDF_SORTUP) != 0;
      bool sorted_down = (item.fmt & HDF_SORTDOWN) != 0;

      FillRect(hdc, &rect, theme.HeaderBrush());

      RECT text_rect = rect;
      text_rect.left += 8;
      text_rect.right -= 8;
      if (sorted_up || sorted_down) {
        text_rect.right -= arrow_size.cx + 6;
      }

      UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
      if (item.fmt & HDF_RIGHT) {
        format |= DT_RIGHT;
      } else if (item.fmt & HDF_CENTER) {
        format |= DT_CENTER;
      }

      SetBkMode(hdc, TRANSPARENT);
      SetTextColor(hdc, theme.TextColor());
      DrawTextW(hdc, text, -1, &text_rect, format);

      if ((sorted_up || sorted_down) && header_theme) {
        RECT arrow_rect = rect;
        arrow_rect.right -= 6;
        arrow_rect.left = arrow_rect.right - arrow_size.cx;
        arrow_rect.top = rect.top + (rect.bottom - rect.top - arrow_size.cy) / 2;
        arrow_rect.bottom = arrow_rect.top + arrow_size.cy;
        int arrow_state = sorted_up ? HSAS_SORTEDUP : HSAS_SORTEDDOWN;
        DrawThemeBackground(header_theme, hdc, HP_HEADERSORTARROW, arrow_state, &arrow_rect, nullptr);
      }
    }

    if (header_theme) {
      CloseThemeData(header_theme);
    }
    if (old_font) {
      SelectObject(hdc, old_font);
    }
    EndPaint(hwnd, &ps);
    return 0;
  }
  if (message == WM_THEMECHANGED) {
    InvalidateRect(hwnd, nullptr, TRUE);
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK ThemePresetListViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR) {
  if (message == WM_SETFOCUS || message == WM_KILLFOCUS) {
    SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
  }
  if (message == WM_UPDATEUISTATE) {
    LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
    SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
    return result;
  }
  if (message == WM_THEMECHANGED) {
    InvalidateRect(hwnd, nullptr, TRUE);
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

void SetupPresetListView(HWND list) {
  if (!list) {
    return;
  }
  DWORD ex_mask = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_BORDERSELECT | LVS_EX_TRACKSELECT | LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE | LVS_EX_UNDERLINEHOT;
  DWORD ex_style = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER;
  ListView_SetExtendedListViewStyleEx(list, ex_mask, ex_style);
  LVCOLUMNW col = {};
  col.mask = LVCF_WIDTH | LVCF_FMT;
  col.fmt = LVCFMT_LEFT;
  col.cx = 120;
  ListView_InsertColumn(list, 0, &col);
  SendMessageW(list, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
  if (!GetWindowSubclass(list, ThemePresetListViewProc, kThemePresetListViewSubclassId, nullptr)) {
    SetWindowSubclass(list, ThemePresetListViewProc, kThemePresetListViewSubclassId, 0);
  }
  Theme::Current().ApplyToListView(list);
}

void SetupColorListView(HWND list) {
  if (!list) {
    return;
  }
  DWORD ex_mask = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_BORDERSELECT | LVS_EX_TRACKSELECT | LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE | LVS_EX_UNDERLINEHOT;
  DWORD ex_style = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER;
  ListView_SetExtendedListViewStyleEx(list, ex_mask, ex_style);
  LVCOLUMNW col = {};
  col.mask = LVCF_TEXT | LVCF_WIDTH;
  col.cx = 150;
  col.pszText = const_cast<wchar_t*>(L"Color");
  ListView_InsertColumn(list, 0, &col);
  col.cx = 90;
  col.pszText = const_cast<wchar_t*>(L"Hex");
  ListView_InsertColumn(list, 1, &col);
  SendMessageW(list, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
  if (!GetWindowSubclass(list, ThemePresetListViewProc, kThemePresetListViewSubclassId, nullptr)) {
    SetWindowSubclass(list, ThemePresetListViewProc, kThemePresetListViewSubclassId, 0);
  }
  Theme::Current().ApplyToListView(list);
  HWND header = ListView_GetHeader(list);
  if (header && !GetWindowSubclass(header, ThemePresetHeaderProc, kThemePresetHeaderSubclassId, nullptr)) {
    SetWindowSubclass(header, ThemePresetHeaderProc, kThemePresetHeaderSubclassId, 0);
  }
}

struct ColorSortContext {
  const ThemePreset* preset = nullptr;
  int column = 0;
  bool ascending = true;
};

int CALLBACK CompareColorListItems(LPARAM left_param, LPARAM right_param, LPARAM sort_param) {
  auto* ctx = reinterpret_cast<ColorSortContext*>(sort_param);
  if (!ctx || !ctx->preset) {
    return 0;
  }
  int left_index = static_cast<int>(left_param);
  int right_index = static_cast<int>(right_param);
  if (left_index < 0 || left_index >= static_cast<int>(std::size(kColorFields)) || right_index < 0 || right_index >= static_cast<int>(std::size(kColorFields))) {
    return 0;
  }
  int result = 0;
  if (ctx->column == 0) {
    result = CompareTextInsensitive(kColorFields[left_index].label, kColorFields[right_index].label);
  } else if (ctx->column == 1) {
    COLORREF left = ctx->preset->colors.*(kColorFields[left_index].member);
    COLORREF right = ctx->preset->colors.*(kColorFields[right_index].member);
    result = CompareColorValue(left, right);
  }
  if (result == 0) {
    return 0;
  }
  return ctx->ascending ? result : -result;
}

int GetSelectedColorField(HWND list) {
  if (!list) {
    return -1;
  }
  int row = ListView_GetNextItem(list, -1, LVNI_SELECTED);
  if (row < 0) {
    return -1;
  }
  LVITEMW item = {};
  item.mask = LVIF_PARAM;
  item.iItem = row;
  if (!ListView_GetItem(list, &item)) {
    return -1;
  }
  return static_cast<int>(item.lParam);
}

void ReselectColorField(HWND list, int field_index) {
  if (!list || field_index < 0) {
    return;
  }
  LVFINDINFOW find = {};
  find.flags = LVFI_PARAM;
  find.lParam = static_cast<LPARAM>(field_index);
  int row = ListView_FindItem(list, -1, &find);
  if (row < 0) {
    return;
  }
  ListView_SetItemState(list, row, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
  ListView_EnsureVisible(list, row, FALSE);
}

void FillColorList(ThemePresetWindowState* state, const ThemePreset* preset) {
  if (!state || !state->color_list) {
    return;
  }
  HWND list = state->color_list;
  int selected_field = GetSelectedColorField(list);
  ListView_DeleteAllItems(list);
  if (!preset) {
    UpdateListViewSort(list, state->color_sort_column, state->color_sort_ascending);
    return;
  }
  for (size_t i = 0; i < std::size(kColorFields); ++i) {
    const auto& field = kColorFields[i];
    COLORREF color = preset->colors.*(field.member);
    std::wstring hex = FormatColorHex(color);
    LVITEMW item = {};
    item.mask = LVIF_TEXT | LVIF_PARAM;
    item.iItem = static_cast<int>(i);
    item.pszText = const_cast<wchar_t*>(field.label);
    item.lParam = static_cast<LPARAM>(i);
    int index = ListView_InsertItem(list, &item);
    if (index >= 0) {
      ListView_SetItemText(list, index, 1, const_cast<wchar_t*>(hex.c_str()));
    }
  }
  if (state->color_sort_column >= 0) {
    ColorSortContext ctx = {};
    ctx.preset = preset;
    ctx.column = state->color_sort_column;
    ctx.ascending = state->color_sort_ascending;
    ListView_SortItemsEx(list, CompareColorListItems, reinterpret_cast<LPARAM>(&ctx));
  }
  UpdateListViewSort(list, state->color_sort_column, state->color_sort_ascending);
  HWND header = ListView_GetHeader(list);
  if (header) {
    InvalidateRect(header, nullptr, TRUE);
  }
  if (selected_field >= 0) {
    ReselectColorField(list, selected_field);
  }
}

int RefreshPresetList(HWND list, const std::vector<ThemePreset>& presets, int selected_index) {
  if (!list) {
    return -1;
  }
  ListView_DeleteAllItems(list);
  for (size_t i = 0; i < presets.size(); ++i) {
    const auto& preset = presets[i];
    LVITEMW item = {};
    item.mask = LVIF_TEXT | LVIF_PARAM;
    item.iItem = static_cast<int>(i);
    item.pszText = const_cast<wchar_t*>(preset.name.c_str());
    item.lParam = static_cast<LPARAM>(i);
    ListView_InsertItem(list, &item);
  }
  int index = selected_index;
  if (index < 0 || index >= static_cast<int>(presets.size())) {
    index = presets.empty() ? -1 : 0;
  }
  if (index >= 0) {
    ListView_SetItemState(list, index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(list, index, FALSE);
  }
  return index;
}

void SyncSelection(ThemePresetWindowState* state) {
  if (!state) {
    return;
  }
  int preset_index = state->preset_list ? GetSelectedPresetIndex(state->preset_list) : -1;
  if (preset_index >= 0 && preset_index < static_cast<int>(state->presets.size())) {
    state->selected_index = preset_index;
  }
  ThemePreset* preset = CurrentPreset(state);
  FillColorList(state, preset);
  if (state->dark_check) {
    SendMessageW(state->dark_check, BM_SETCHECK, (preset && preset->is_dark) ? BST_CHECKED : BST_UNCHECKED, 0);
  }
}

ThemePreset BuildPresetFromTemplate(const ThemePresetWindowState* state, int template_index) {
  ThemePreset preset;
  if (!state || state->templates.empty()) {
    return preset;
  }
  if (template_index < 0 || template_index >= static_cast<int>(state->templates.size())) {
    return state->templates.front();
  }
  return state->templates[static_cast<size_t>(template_index)];
}

void ApplyCurrentTheme(HWND hwnd) {
  Theme::Current().ApplyToWindow(hwnd);
  Theme::Current().ApplyToChildren(hwnd);
  InvalidateRect(hwnd, nullptr, TRUE);
}

void RefreshThemeRendering(ThemePresetWindowState* state) {
  if (!state) {
    return;
  }
  if (state->preset_list) {
    Theme::Current().ApplyToListView(state->preset_list);
    state->selected_index = RefreshPresetList(state->preset_list, state->presets, state->selected_index);
    RedrawWindow(state->preset_list, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME);
  }
  if (state->color_list) {
    Theme::Current().ApplyToListView(state->color_list);
    RedrawWindow(state->color_list, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME);
  }
  RedrawWindow(state->hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
}

void LayoutControls(ThemePresetWindowState* state) {
  if (!state || !state->hwnd) {
    return;
  }
  RECT rc = {};
  GetClientRect(state->hwnd, &rc);
  int width = rc.right - rc.left;
  int height = rc.bottom - rc.top;

  int bottom_row_h = kBottomButtonHeight;
  int content_top = kPadding;
  int content_h = height - content_top - kPadding - bottom_row_h - kGap;
  if (content_h < 100) {
    content_h = 100;
  }

  int button_rows_h = kButtonHeight * 3 + kGap * 2;
  int left_x = kPadding;
  int left_y = content_top;
  int left_w = kLeftPanelWidth;
  int left_h = content_h;

  int right_x = left_x + left_w + kGap;
  int right_w = std::max(180, width - right_x - kPadding);
  int right_y = content_top;

  int template_group_h = kGroupBoxCaptionHeight + kGroupBoxPadding * 2 + kButtonHeight;
  if (template_group_h < 60) {
    template_group_h = 60;
  }
  int colors_group_h = content_h - template_group_h - kGap;
  if (colors_group_h < 100) {
    colors_group_h = 100;
    template_group_h = std::max(60, content_h - colors_group_h - kGap);
  }

  int left_inner_x = left_x + kGroupBoxPadding;
  int left_inner_y = left_y + kGroupBoxCaptionHeight;
  int left_inner_w = left_w - kGroupBoxPadding * 2;
  int left_inner_h = left_h - kGroupBoxCaptionHeight - kGroupBoxPadding;
  int list_h = left_inner_h - button_rows_h - kGap;
  if (list_h < 80) {
    list_h = 80;
  }
  int list_y = left_inner_y;

  int left_btn_w = (left_inner_w - kGap) / 2;
  int row1_y = list_y + list_h + kGap;
  int row2_y = row1_y + kButtonHeight + kGap;
  int row3_y = row2_y + kButtonHeight + kGap;

  if (state->presets_group) {
    SetWindowPos(state->presets_group, nullptr, left_x, left_y, left_w, left_h, SWP_NOZORDER);
  }
  if (state->preset_list) {
    SetWindowPos(state->preset_list, nullptr, left_inner_x, list_y, left_inner_w, list_h, SWP_NOZORDER);
    int col_width = std::max(60, left_inner_w - 6);
    ListView_SetColumnWidth(state->preset_list, 0, col_width);
  }
  if (state->new_btn) {
    SetWindowPos(state->new_btn, nullptr, left_inner_x, row1_y, left_btn_w, kButtonHeight, SWP_NOZORDER);
  }
  if (state->duplicate_btn) {
    SetWindowPos(state->duplicate_btn, nullptr, left_inner_x + left_btn_w + kGap, row1_y, left_btn_w, kButtonHeight, SWP_NOZORDER);
  }
  if (state->rename_btn) {
    SetWindowPos(state->rename_btn, nullptr, left_inner_x, row2_y, left_btn_w, kButtonHeight, SWP_NOZORDER);
  }
  if (state->delete_btn) {
    SetWindowPos(state->delete_btn, nullptr, left_inner_x + left_btn_w + kGap, row2_y, left_btn_w, kButtonHeight, SWP_NOZORDER);
  }
  if (state->import_btn) {
    SetWindowPos(state->import_btn, nullptr, left_inner_x, row3_y, left_btn_w, kButtonHeight, SWP_NOZORDER);
  }
  if (state->export_btn) {
    SetWindowPos(state->export_btn, nullptr, left_inner_x + left_btn_w + kGap, row3_y, left_btn_w, kButtonHeight, SWP_NOZORDER);
  }

  if (state->colors_group) {
    SetWindowPos(state->colors_group, nullptr, right_x, right_y, right_w, colors_group_h, SWP_NOZORDER);
  }
  int colors_inner_x = right_x + kGroupBoxPadding;
  int colors_inner_y = right_y + kGroupBoxCaptionHeight;
  int colors_inner_w = right_w - kGroupBoxPadding * 2;
  int colors_inner_h = colors_group_h - kGroupBoxCaptionHeight - kGroupBoxPadding;
  int edit_row_h = kButtonHeight;
  int color_list_h = colors_inner_h - edit_row_h - kGap;
  if (color_list_h < 80) {
    color_list_h = 80;
  }
  int color_list_y = colors_inner_y;
  if (state->color_list) {
    SetWindowPos(state->color_list, nullptr, colors_inner_x, color_list_y, colors_inner_w, color_list_h, SWP_NOZORDER);
  }

  int edit_row_y = color_list_y + color_list_h + kGap;
  int edit_btn_w = 90;
  if (state->edit_color_btn) {
    SetWindowPos(state->edit_color_btn, nullptr, colors_inner_x, edit_row_y, edit_btn_w, edit_row_h, SWP_NOZORDER);
  }
  if (state->dark_check) {
    SetWindowPos(state->dark_check, nullptr, colors_inner_x + edit_btn_w + kGap, edit_row_y, colors_inner_w - edit_btn_w - kGap, edit_row_h, SWP_NOZORDER);
  }

  int templates_group_y = right_y + colors_group_h + kGap;
  if (state->templates_group) {
    SetWindowPos(state->templates_group, nullptr, right_x, templates_group_y, right_w, template_group_h, SWP_NOZORDER);
  }
  int templates_inner_x = right_x + kGroupBoxPadding;
  int templates_inner_y = templates_group_y + kGroupBoxCaptionHeight;
  int templates_inner_w = right_w - kGroupBoxPadding * 2;
  int template_row_h = kButtonHeight;
  int template_row_y = templates_inner_y;
  int template_btn_w = 110;
  int combo_w = std::max(120, templates_inner_w - template_btn_w - kGap);
  if (state->template_combo) {
    SetWindowPos(state->template_combo, nullptr, templates_inner_x, template_row_y, combo_w, template_row_h, SWP_NOZORDER);
  }
  if (state->template_btn) {
    SetWindowPos(state->template_btn, nullptr, templates_inner_x + combo_w + kGap, template_row_y, template_btn_w, template_row_h, SWP_NOZORDER);
  }

  int bottom_y = height - kPadding - bottom_row_h;
  int btn_w = 70;
  int cancel_x = width - kPadding - btn_w;
  int ok_x = cancel_x - kGap - btn_w;
  int apply_x = ok_x - kGap - btn_w;

  if (state->apply_btn) {
    SetWindowPos(state->apply_btn, nullptr, apply_x, bottom_y, btn_w, bottom_row_h, SWP_NOZORDER);
  }
  if (state->ok_btn) {
    SetWindowPos(state->ok_btn, nullptr, ok_x, bottom_y, btn_w, bottom_row_h, SWP_NOZORDER);
  }
  if (state->cancel_btn) {
    SetWindowPos(state->cancel_btn, nullptr, cancel_x, bottom_y, btn_w, bottom_row_h, SWP_NOZORDER);
  }
}

void CreateControls(ThemePresetWindowState* state) {
  if (!state || !state->hwnd) {
    return;
  }
  HWND hwnd = state->hwnd;

  state->presets_group = CreateWindowExW(0, L"BUTTON", L"Presets", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);

  state->preset_list = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | LVS_NOCOLUMNHEADER | LVS_NOSORTHEADER, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kPresetListId)), nullptr, nullptr);

  state->new_btn = CreateWindowExW(0, L"BUTTON", L"New...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kNewPresetId)), nullptr, nullptr);
  state->duplicate_btn = CreateWindowExW(0, L"BUTTON", L"Duplicate", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kDuplicatePresetId)), nullptr, nullptr);
  state->rename_btn = CreateWindowExW(0, L"BUTTON", L"Rename...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRenamePresetId)), nullptr, nullptr);
  state->delete_btn = CreateWindowExW(0, L"BUTTON", L"Delete", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kDeletePresetId)), nullptr, nullptr);
  state->import_btn = CreateWindowExW(0, L"BUTTON", L"Import...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kImportPresetId)), nullptr, nullptr);
  state->export_btn = CreateWindowExW(0, L"BUTTON", L"Export...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kExportPresetId)), nullptr, nullptr);

  state->colors_group = CreateWindowExW(0, L"BUTTON", L"Colors", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);

  state->color_list = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kColorListId)), nullptr, nullptr);

  state->edit_color_btn = CreateWindowExW(0, L"BUTTON", L"Edit Color...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kEditColorId)), nullptr, nullptr);

  state->dark_check = CreateWindowExW(0, L"BUTTON", L"Treat as dark theme", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kDarkCheckId)), nullptr, nullptr);

  state->templates_group = CreateWindowExW(0, L"BUTTON", L"Templates", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 0, 0, 0, 0, hwnd, nullptr, nullptr, nullptr);

  state->template_combo = CreateWindowExW(0, WC_COMBOBOXW, L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTemplateComboId)), nullptr, nullptr);

  state->template_btn = CreateWindowExW(0, L"BUTTON", L"Apply Template", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kApplyTemplateId)), nullptr, nullptr);

  state->apply_btn = CreateWindowExW(0, L"BUTTON", L"Apply", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kApplyId)), nullptr, nullptr);
  state->ok_btn = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);
  state->cancel_btn = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);

  SetupPresetListView(state->preset_list);
  SetupColorListView(state->color_list);
}

void PopulateTemplates(ThemePresetWindowState* state) {
  if (!state || !state->template_combo) {
    return;
  }
  SendMessageW(state->template_combo, CB_RESETCONTENT, 0, 0);
  for (const auto& preset : state->templates) {
    SendMessageW(state->template_combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(preset.name.c_str()));
  }
  SendMessageW(state->template_combo, CB_SETCURSEL, 0, 0);
}

void PopulatePresets(ThemePresetWindowState* state) {
  if (!state || !state->preset_list) {
    return;
  }
  int active_index = FindPresetIndexByName(state->presets, state->active_name);
  state->selected_index = RefreshPresetList(state->preset_list, state->presets, active_index);
  SyncSelection(state);
}

void ApplySelectedPreset(ThemePresetWindowState* state, bool close_dialog) {
  if (!state || !state->window) {
    return;
  }
  SyncSelection(state);
  ThemePreset* preset = CurrentPreset(state);
  if (!preset) {
    return;
  }
  state->active_name = preset->name;
  state->window->UpdateThemePresets(state->presets, state->active_name, true);
  ApplyCurrentTheme(state->hwnd);
  RefreshThemeRendering(state);
  if (close_dialog) {
    RestoreOwnerWindow(state->owner, &state->owner_restored);
    DestroyWindow(state->hwnd);
  }
}

LRESULT CALLBACK ThemePresetWindowProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam) {
  auto* state = reinterpret_cast<ThemePresetWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
  case WM_NCCREATE: {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
    return TRUE;
  }
  case WM_CREATE: {
    state = reinterpret_cast<ThemePresetWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) {
      return -1;
    }
    state->hwnd = hwnd;
    SetWindowTextW(hwnd, kThemePresetTitle);
    CreateControls(state);
    state->font = ui::DefaultUIFont();
    ApplyFontRecursive(hwnd, state->font);
    PopulateTemplates(state);
    PopulatePresets(state);
    ApplyCurrentTheme(hwnd);
    LayoutControls(state);
    return 0;
  }
  case WM_SIZE:
    LayoutControls(state);
    return 0;
  case WM_SETTINGCHANGE:
    if (Theme::UpdateFromSystem()) {
      ApplyCurrentTheme(hwnd);
      RefreshThemeRendering(state);
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
    if (!state) {
      return 0;
    }
    int id = LOWORD(wparam);
    int code = HIWORD(wparam);
    switch (id) {
    case kNewPresetId: {
      std::wstring name;
      if (!PromptPresetName(hwnd, L"New Preset", L"", &name)) {
        return 0;
      }
      int sel = state->template_combo ? static_cast<int>(SendMessageW(state->template_combo, CB_GETCURSEL, 0, 0)) : -1;
      ThemePreset preset = BuildPresetFromTemplate(state, sel);
      preset.name = MakeUniquePresetName(state->presets, name);
      state->presets.push_back(preset);
      state->selected_index = static_cast<int>(state->presets.size() - 1);
      state->selected_index = RefreshPresetList(state->preset_list, state->presets, state->selected_index);
      SyncSelection(state);
      return 0;
    }
    case kDuplicatePresetId: {
      ThemePreset* preset = CurrentPreset(state);
      if (!preset) {
        return 0;
      }
      ThemePreset copy = *preset;
      copy.name = MakeUniquePresetName(state->presets, preset->name + L" Copy");
      state->presets.push_back(copy);
      state->selected_index = static_cast<int>(state->presets.size() - 1);
      state->selected_index = RefreshPresetList(state->preset_list, state->presets, state->selected_index);
      SyncSelection(state);
      return 0;
    }
    case kRenamePresetId: {
      ThemePreset* preset = CurrentPreset(state);
      if (!preset) {
        return 0;
      }
      std::wstring name;
      if (!PromptPresetName(hwnd, L"Rename Preset", preset->name, &name)) {
        return 0;
      }
      preset->name = MakeUniquePresetName(state->presets, name);
      state->selected_index = RefreshPresetList(state->preset_list, state->presets, state->selected_index);
      SyncSelection(state);
      return 0;
    }
    case kDeletePresetId: {
      if (state->presets.size() <= 1) {
        ui::ShowWarning(hwnd, L"At least one preset must remain.");
        return 0;
      }
      ThemePreset* preset = CurrentPreset(state);
      if (!preset) {
        return 0;
      }
      if (!ui::ConfirmDelete(hwnd, L"Delete Preset", preset->name)) {
        return 0;
      }
      state->presets.erase(state->presets.begin() + state->selected_index);
      if (state->selected_index >= static_cast<int>(state->presets.size())) {
        state->selected_index = static_cast<int>(state->presets.size() - 1);
      }
      state->selected_index = RefreshPresetList(state->preset_list, state->presets, state->selected_index);
      SyncSelection(state);
      return 0;
    }
    case kImportPresetId: {
      std::wstring path;
      if (!PromptOpenThemeFile(hwnd, &path)) {
        return 0;
      }
      std::vector<ThemePreset> imported;
      std::wstring error;
      if (!ThemePresetStore::ImportFromFile(path, &imported, &error)) {
        ui::ShowError(hwnd, error.empty() ? L"Failed to import theme presets." : error);
        return 0;
      }
      if (!imported.empty()) {
        for (auto& preset : imported) {
          preset.name = MakeUniquePresetName(state->presets, preset.name);
          state->presets.push_back(preset);
        }
        state->selected_index = static_cast<int>(state->presets.size() - 1);
        state->selected_index = RefreshPresetList(state->preset_list, state->presets, state->selected_index);
        SyncSelection(state);
      }
      return 0;
    }
    case kExportPresetId: {
      std::wstring path;
      if (!PromptSaveThemeFile(hwnd, &path)) {
        return 0;
      }
      std::wstring error;
      if (!ThemePresetStore::ExportToFile(path, state->presets, &error)) {
        ui::ShowError(hwnd, error.empty() ? L"Failed to export theme presets." : error);
      }
      return 0;
    }
    case kEditColorId: {
      ThemePreset* preset = CurrentPreset(state);
      if (!preset || !state->color_list) {
        return 0;
      }
      int row = ListView_GetNextItem(state->color_list, -1, LVNI_SELECTED);
      if (row < 0) {
        return 0;
      }
      LVITEMW item = {};
      item.mask = LVIF_PARAM;
      item.iItem = row;
      if (!ListView_GetItem(state->color_list, &item)) {
        return 0;
      }
      int field_index = static_cast<int>(item.lParam);
      if (field_index < 0 || field_index >= static_cast<int>(std::size(kColorFields))) {
        return 0;
      }
      COLORREF* color = &(preset->colors.*(kColorFields[field_index].member));
      if (ChooseColorFor(hwnd, color, state->custom_colors)) {
        FillColorList(state, preset);
      }
      return 0;
    }
    case kDarkCheckId: {
      ThemePreset* preset = CurrentPreset(state);
      if (!preset || !state->dark_check) {
        return 0;
      }
      preset->is_dark = (SendMessageW(state->dark_check, BM_GETCHECK, 0, 0) == BST_CHECKED);
      return 0;
    }
    case kApplyTemplateId: {
      ThemePreset* preset = CurrentPreset(state);
      if (!preset) {
        return 0;
      }
      int sel = state->template_combo ? static_cast<int>(SendMessageW(state->template_combo, CB_GETCURSEL, 0, 0)) : -1;
      ThemePreset tmpl = BuildPresetFromTemplate(state, sel);
      preset->colors = tmpl.colors;
      preset->is_dark = tmpl.is_dark;
      SyncSelection(state);
      return 0;
    }
    case kApplyId:
      ApplySelectedPreset(state, false);
      return 0;
    case IDOK:
      ApplySelectedPreset(state, true);
      return 0;
    case IDCANCEL:
      RestoreOwnerWindow(state->owner, &state->owner_restored);
      DestroyWindow(hwnd);
      return 0;
    default:
      break;
    }
    break;
  }
  case WM_NOTIFY: {
    auto* hdr = reinterpret_cast<NMHDR*>(lparam);
    if (!hdr || !state) {
      break;
    }
    if (hdr->hwndFrom == state->preset_list && hdr->code == LVN_ITEMCHANGED) {
      auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
      if (info && (info->uNewState & LVIS_SELECTED) && info->iItem >= 0) {
        SyncSelection(state);
      }
      return 0;
    }
    if (hdr->hwndFrom == state->preset_list && hdr->code == NM_CUSTOMDRAW) {
      return ui::HandleThemedListViewCustomDraw(state->preset_list, reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
    }
    if (hdr->hwndFrom == state->color_list && hdr->code == NM_DBLCLK) {
      SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(kEditColorId, 0), 0);
      return 0;
    }
    if (hdr->hwndFrom == state->color_list && hdr->code == LVN_COLUMNCLICK) {
      auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
      if (info) {
        if (state->color_sort_column == info->iSubItem) {
          state->color_sort_ascending = !state->color_sort_ascending;
        } else {
          state->color_sort_column = info->iSubItem;
          state->color_sort_ascending = true;
        }
        ThemePreset* preset = CurrentPreset(state);
        if (preset) {
          ColorSortContext ctx = {};
          ctx.preset = preset;
          ctx.column = state->color_sort_column;
          ctx.ascending = state->color_sort_ascending;
          ListView_SortItemsEx(state->color_list, CompareColorListItems, reinterpret_cast<LPARAM>(&ctx));
        }
        UpdateListViewSort(state->color_list, state->color_sort_column, state->color_sort_ascending);
        HWND header = ListView_GetHeader(state->color_list);
        if (header) {
          InvalidateRect(header, nullptr, TRUE);
        }
      }
      return 0;
    }
    if (hdr->hwndFrom == state->color_list && hdr->code == NM_CUSTOMDRAW) {
      return ui::HandleThemedListViewCustomDraw(state->color_list, reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
    }
    break;
  }
  case WM_CLOSE:
    RestoreOwnerWindow(state ? state->owner : nullptr, state ? &state->owner_restored : nullptr);
    DestroyWindow(hwnd);
    return 0;
  case WM_NCDESTROY:
    if (state) {
      if (state->font) {
        DeleteObject(state->font);
        state->font = nullptr;
      }
    }
    return 0;
  default:
    break;
  }
  return DefWindowProcW(hwnd, msg, wparam, lparam);
}

} // namespace

void MainWindow::ShowThemePresetsDialog() {
  WNDCLASSW wc = {};
  wc.lpfnWndProc = ThemePresetWindowProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = nullptr;
  wc.lpszClassName = kThemePresetClass;
  RegisterClassW(&wc);

  auto* state = new ThemePresetWindowState();
  state->window = this;
  state->owner = hwnd_;
  state->presets = theme_presets_;
  state->templates = ThemePresetStore::BuiltInPresets();
  state->active_name = active_theme_preset_;

  DWORD style = WS_POPUP | WS_CAPTION | WS_SYSMENU;
  DWORD ex_style = WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT;
  RECT rect = {0, 0, kWindowWidth, kWindowHeight};
  AdjustWindowRectEx(&rect, style, FALSE, ex_style);
  int width = rect.right - rect.left;
  int height = rect.bottom - rect.top;

  HWND hwnd = CreateWindowExW(ex_style, kThemePresetClass, kThemePresetTitle, style, CW_USEDEFAULT, CW_USEDEFAULT, width, height, hwnd_, nullptr, wc.hInstance, state);
  if (!hwnd) {
    delete state;
    return;
  }

  CenterToOwner(hwnd, hwnd_, width, height);
  ApplyCurrentTheme(hwnd);
  EnableWindow(hwnd_, FALSE);
  ShowWindow(hwnd, SW_SHOW);
  UpdateWindow(hwnd);

  MSG msg = {};
  while (IsWindow(hwnd) && GetMessageW(&msg, nullptr, 0, 0)) {
    if (!IsDialogMessageW(hwnd, &msg)) {
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  }
  RestoreOwnerWindow(hwnd_, &state->owner_restored);
  delete state;
}

} // namespace regkit
