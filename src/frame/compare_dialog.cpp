// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/compare_dialog.h"

#include "frame/window_detail.h"

#include <commctrl.h>
#include <shellapi.h>

#include "appearance/default_font.h"
#include "appearance/dialog_layout.h"
#include "appearance/feedback.h"
#include "appearance/theme.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "resource.h"
#include "win32/file_dialog.h"
#include "win32/window_metrics.h"

namespace regkit::command_detail
{

using namespace window_detail;

namespace
{


const wchar_t* CompareSourceLabel(CompareSourceType type)
{
    switch (type)
    {
    case CompareSourceType::kRegFile:
        return L"Reg File";
    case CompareSourceType::kNetwork:
        return L"Network Registry";
    case CompareSourceType::kOfflineHive:
        return L"Offline Hive";
    default:
        return L"Registry";
    }
}

CompareSourceType CompareSourceFromIndex(int index)
{
    switch (index)
    {
    case 1:
        return CompareSourceType::kRegFile;
    case 2:
        return CompareSourceType::kOfflineHive;
    case 3:
        return CompareSourceType::kNetwork;
    default:
        return CompareSourceType::kRegistry;
    }
}
struct CompareDialogState
{
    CompareDialogDefaults data;
    HFONT ui_font = nullptr;
};

struct EditBorderState
{
    bool hot = false;
    UINT dpi = 0;
    int x_edge = 1;
    int y_edge = 1;
    int x_scroll = 0;
    int y_scroll = 0;
};

using GetSystemMetricsForDpiFn = int(WINAPI*)(int, UINT);

int GetMetricForDpi(int index, UINT dpi)
{
    static GetSystemMetricsForDpiFn get_for_dpi = []() -> GetSystemMetricsForDpiFn {
        HMODULE user32 = GetModuleHandleW(L"user32.dll");
        if (!user32)
        {
            return nullptr;
        }
        return reinterpret_cast<GetSystemMetricsForDpiFn>(GetProcAddress(user32, "GetSystemMetricsForDpi"));
    }();
    if (get_for_dpi)
    {
        return get_for_dpi(index, dpi);
    }
    int value = GetSystemMetrics(index);
    return MulDiv(value, static_cast<int>(dpi), 96);
}

void UpdateEditBorderMetrics(HWND hwnd, EditBorderState* state, UINT dpi_override = 0)
{
    if (!state)
    {
        return;
    }
    UINT dpi = dpi_override ? dpi_override : (hwnd ? win32::DpiForWindow(hwnd) : 96);
    state->dpi = dpi;
    state->x_edge = GetMetricForDpi(SM_CXEDGE, dpi);
    state->y_edge = GetMetricForDpi(SM_CYEDGE, dpi);
    state->x_scroll = GetMetricForDpi(SM_CXVSCROLL, dpi);
    state->y_scroll = GetMetricForDpi(SM_CYVSCROLL, dpi);
    if (state->x_edge < 1)
    {
        state->x_edge = 1;
    }
    if (state->y_edge < 1)
    {
        state->y_edge = 1;
    }
}

LRESULT CALLBACK EditBorderSubclassProc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam, UINT_PTR id, DWORD_PTR data)
{
    auto* state = reinterpret_cast<EditBorderState*>(data);
    switch (msg)
    {
    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, EditBorderSubclassProc, id);
        delete state;
        break;
    case WM_NCCALCSIZE:
        {
            UpdateEditBorderMetrics(hwnd, state);
            int x_edge = state ? state->x_edge : 1;
            int y_edge = state ? state->y_edge : 1;
            if (wparam)
            {
                auto* params = reinterpret_cast<NCCALCSIZE_PARAMS*>(lparam);
                InflateRect(&params->rgrc[0], -x_edge, -y_edge);
                return 0;
            }
            auto* rect = reinterpret_cast<RECT*>(lparam);
            InflateRect(rect, -x_edge, -y_edge);
            return 0;
        }
    case WM_NCPAINT:
        {
            LRESULT result = DefSubclassProc(hwnd, msg, wparam, lparam);
            HDC hdc = GetWindowDC(hwnd);
            if (!hdc)
            {
                return result;
            }
            UpdateEditBorderMetrics(hwnd, state);
            RECT rect = {};
            GetClientRect(hwnd, &rect);
            if (state)
            {
                rect.right += 2 * state->x_edge;
                rect.bottom += 2 * state->y_edge;
                LONG_PTR style = GetWindowLongPtrW(hwnd, GWL_STYLE);
                if ((style & WS_VSCROLL) == WS_VSCROLL)
                {
                    rect.right += state->x_scroll;
                }
                if ((style & WS_HSCROLL) == WS_HSCROLL)
                {
                    rect.bottom += state->y_scroll;
                }
            }

            const Theme& theme = Theme::Current();
            RECT inner = rect;
            InflateRect(&inner, -1, -1);
            HPEN inner_pen = appearance::CachedPen(theme.BackgroundColor(), 1);
            HGDIOBJ old_pen = SelectObject(hdc, inner_pen);
            HGDIOBJ old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
            Rectangle(hdc, inner.left, inner.top, inner.right, inner.bottom);
            SelectObject(hdc, old_brush);
            SelectObject(hdc, old_pen);

            bool enabled = IsWindowEnabled(hwnd) != FALSE;
            COLORREF border = theme.BorderColor();
            if (enabled)
            {
                if (GetFocus() == hwnd)
                {
                    border = theme.FocusColor();
                }
                else if (state && state->hot)
                {
                    border = theme.HoverColor();
                }
            }
            HPEN pen = appearance::CachedPen(border, 1);
            old_pen = SelectObject(hdc, pen);
            old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
            Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
            SelectObject(hdc, old_brush);
            SelectObject(hdc, old_pen);
            ReleaseDC(hwnd, hdc);
            return 0;
        }
    case WM_MOUSEMOVE:
        {
            if (state && !state->hot)
            {
                state->hot = true;
                TRACKMOUSEEVENT tme = {};
                tme.cbSize = sizeof(tme);
                tme.dwFlags = TME_LEAVE;
                tme.hwndTrack = hwnd;
                TrackMouseEvent(&tme);
                SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
            }
            break;
        }
    case WM_MOUSELEAVE:
        if (state && state->hot)
        {
            state->hot = false;
            SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }
        break;
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
    case WM_ENABLE:
        RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME);
        break;
    case WM_DPICHANGED:
    case WM_DPICHANGED_AFTERPARENT:
        UpdateEditBorderMetrics(hwnd, state, (msg == WM_DPICHANGED) ? LOWORD(wparam) : 0);
        SetWindowPos(hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        break;
    default:
        break;
    }
    return DefSubclassProc(hwnd, msg, wparam, lparam);
}

void ApplyEditCustomBorder(HWND parent, int id)
{
    HWND ctrl = GetDlgItem(parent, id);
    if (!ctrl)
    {
        return;
    }
    LONG_PTR ex = GetWindowLongPtrW(ctrl, GWL_EXSTYLE);
    if (ex & WS_EX_CLIENTEDGE)
    {
        ex &= ~static_cast<LONG_PTR>(WS_EX_CLIENTEDGE);
        SetWindowLongPtrW(ctrl, GWL_EXSTYLE, ex);
    }
    LONG_PTR style = GetWindowLongPtrW(ctrl, GWL_STYLE);
    if (style & WS_BORDER)
    {
        style &= ~static_cast<LONG_PTR>(WS_BORDER);
        SetWindowLongPtrW(ctrl, GWL_STYLE, style);
    }
    if (!GetWindowSubclass(ctrl, EditBorderSubclassProc, 1, nullptr))
    {
        auto* state = new EditBorderState();
        if (!SetWindowSubclass(ctrl, EditBorderSubclassProc, 1, reinterpret_cast<DWORD_PTR>(state)))
        {
            delete state;
        }
    }
    SetWindowPos(ctrl, nullptr, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

std::vector<std::wstring> ExtractRegFileKeys(const regfile::Document& data)
{
    std::vector<std::wstring> keys = data.key_order;
    if (keys.empty())
    {
        keys.reserve(data.keys.size());
        for (const auto& entry : data.keys)
        {
            keys.push_back(entry.second.path);
        }
    }
    std::sort(keys.begin(), keys.end(), [](const std::wstring& a, const std::wstring& b) { return util::CompareInsensitive(a, b) < 0; });
    return keys;
}

void ApplyDialogFonts(HWND hwnd, HFONT font)
{
    if (!font)
    {
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
        reinterpret_cast<LPARAM>(font)
    );
}

HFONT CreateDefaultGuiFont()
{
    HFONT stock = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    if (!stock)
    {
        return ui::DefaultUIFont();
    }
    LOGFONTW lf = {};
    if (GetObjectW(stock, sizeof(lf), &lf) == 0)
    {
        return ui::DefaultUIFont();
    }
    HFONT font = CreateFontIndirectW(&lf);
    return font ? font : ui::DefaultUIFont();
}

int ControlHeight(HWND dlg, int id)
{
    HWND ctrl = GetDlgItem(dlg, id);
    if (!ctrl)
    {
        return 0;
    }
    RECT rect = {};
    if (!GetWindowRect(ctrl, &rect))
    {
        return 0;
    }
    int height = static_cast<int>(rect.bottom - rect.top);
    if (height < 0)
    {
        height = 0;
    }
    return height;
}

void SetComboHeights(HWND dlg, int id, int height)
{
    HWND ctrl = GetDlgItem(dlg, id);
    if (!ctrl || height <= 0)
    {
        return;
    }
    RECT rect = {};
    if (!GetWindowRect(ctrl, &rect))
    {
        return;
    }
    int window_height = rect.bottom - rect.top;
    if (window_height <= 0)
    {
        return;
    }
    int target = height;
    if (target > window_height)
    {
        target = window_height;
    }
    SendMessageW(ctrl, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), static_cast<LPARAM>(target));
    int new_total = window_height;
    POINT pt = {rect.left, rect.top};
    ScreenToClient(dlg, &pt);
    SetWindowPos(ctrl, nullptr, pt.x, pt.y, rect.right - rect.left, new_total, SWP_NOZORDER | SWP_NOACTIVATE);
}

void PopulateCombo(HWND combo, const std::vector<std::wstring>& items)
{
    if (!combo)
    {
        return;
    }
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    for (const auto& item : items)
    {
        SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item.c_str()));
    }
}

void SetComboSelection(HWND combo, const std::wstring& value)
{
    if (!combo)
    {
        return;
    }
    if (!value.empty())
    {
        int count = static_cast<int>(SendMessageW(combo, CB_GETCOUNT, 0, 0));
        for (int i = 0; i < count; ++i)
        {
            wchar_t buffer[256] = {};
            SendMessageW(combo, CB_GETLBTEXT, i, reinterpret_cast<LPARAM>(buffer));
            if (util::EqualsInsensitive(buffer, value))
            {
                SendMessageW(combo, CB_SETCURSEL, i, 0);
                return;
            }
        }
    }
    if (SendMessageW(combo, CB_GETCOUNT, 0, 0) > 0)
    {
        SendMessageW(combo, CB_SETCURSEL, 0, 0);
    }
}

std::wstring ReadComboText(HWND combo)
{
    return util::WindowText(combo);
}

void SetDialogText(HWND dlg, int id, const std::wstring& text)
{
    HWND ctrl = GetDlgItem(dlg, id);
    if (ctrl)
    {
        SetWindowTextW(ctrl, text.c_str());
    }
}

void ToggleCompareControls(HWND dlg, bool left, CompareSourceType type)
{
    int file_id = left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE;
    int label_id = left ? IDC_COMPARE_LEFT_FILE_LABEL : IDC_COMPARE_RIGHT_FILE_LABEL;
    int browse_id = left ? IDC_COMPARE_LEFT_BROWSE : IDC_COMPARE_RIGHT_BROWSE;
    const bool network = type == CompareSourceType::kNetwork;
    const bool local = type == CompareSourceType::kRegistry;
    EnableWindow(GetDlgItem(dlg, file_id), !local);
    EnableWindow(GetDlgItem(dlg, browse_id), !local);
    SetDialogText(dlg, label_id, network ? L"Computer:" : L"File:");
}

INT_PTR CALLBACK CompareDialogProc(HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
    auto* state = reinterpret_cast<CompareDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    switch (msg)
    {
    case WM_INITDIALOG:
        {
            state = reinterpret_cast<CompareDialogState*>(lparam);
            SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
            if (!state)
            {
                return TRUE;
            }
            state->ui_font = CreateDefaultGuiFont();
            ApplyDialogFonts(dlg, state->ui_font);
            Theme::Current().ApplyToWindow(dlg);
            Theme::Current().ApplyToChildren(dlg);
            ApplyEditCustomBorder(dlg, IDC_COMPARE_LEFT_FILE);
            ApplyEditCustomBorder(dlg, IDC_COMPARE_RIGHT_FILE);

            PopulateCombo(GetDlgItem(dlg, IDC_COMPARE_LEFT_SOURCE), {L"Registry", L"Reg File", L"Offline Hive", L"Network Registry"});
            PopulateCombo(GetDlgItem(dlg, IDC_COMPARE_RIGHT_SOURCE), {L"Registry", L"Reg File", L"Offline Hive", L"Network Registry"});

            SetComboSelection(GetDlgItem(dlg, IDC_COMPARE_LEFT_SOURCE), CompareSourceLabel(state->data.left.type));
            SetComboSelection(GetDlgItem(dlg, IDC_COMPARE_RIGHT_SOURCE), CompareSourceLabel(state->data.right.type));
            SetDialogText(dlg, IDC_COMPARE_LEFT_FILE, state->data.left.file_path);
            SetDialogText(dlg, IDC_COMPARE_RIGHT_FILE, state->data.right.file_path);
            SetDialogText(dlg, IDC_COMPARE_LEFT_KEY, state->data.left.key_path);
            SetDialogText(dlg, IDC_COMPARE_RIGHT_KEY, state->data.right.key_path);
            CheckDlgButton(dlg, IDC_COMPARE_LEFT_RECURSIVE, state->data.left.recursive ? BST_CHECKED : BST_UNCHECKED);
            CheckDlgButton(dlg, IDC_COMPARE_RIGHT_RECURSIVE, state->data.right.recursive ? BST_CHECKED : BST_UNCHECKED);
            int filter_id = IDC_COMPARE_SHOW_DIFFERENCES;
            if (state->data.filter == search::compare::RowFilter::kMatches)
            {
                filter_id = IDC_COMPARE_SHOW_MATCHING;
            }
            else if (state->data.filter == search::compare::RowFilter::kAll)
            {
                filter_id = IDC_COMPARE_SHOW_BOTH;
            }
            CheckRadioButton(dlg, IDC_COMPARE_SHOW_DIFFERENCES, IDC_COMPARE_SHOW_BOTH, filter_id);

            auto populate_file_keys = [&](bool left) {
                std::wstring file_path = util::DialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE);
                if (file_path.empty())
                {
                    return;
                }
                regfile::Document data;
                std::wstring error;
                if (!regfile::Load(file_path, &data, &error))
                {
                    return;
                }
                std::vector<std::wstring> keys = ExtractRegFileKeys(data);
                HWND combo = GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY);
                std::wstring current = ReadComboText(combo);
                PopulateCombo(combo, keys);
                if (!current.empty())
                {
                    SetComboSelection(combo, current);
                }
                else if (!keys.empty())
                {
                    SendMessageW(combo, CB_SETCURSEL, 0, 0);
                    SetDialogText(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY, keys.front());
                }
            };
            populate_file_keys(true);
            populate_file_keys(false);

            int edit_height = ControlHeight(dlg, IDC_COMPARE_LEFT_FILE);
            if (edit_height > 0)
            {
                SetComboHeights(dlg, IDC_COMPARE_LEFT_SOURCE, edit_height);
                SetComboHeights(dlg, IDC_COMPARE_LEFT_KEY, edit_height);
                SetComboHeights(dlg, IDC_COMPARE_RIGHT_SOURCE, edit_height);
                SetComboHeights(dlg, IDC_COMPARE_RIGHT_KEY, edit_height);
            }

            ToggleCompareControls(dlg, true, state->data.left.type);
            ToggleCompareControls(dlg, false, state->data.right.type);
            appearance::CenterWindow(dlg, GetWindow(dlg, GW_OWNER));
            return TRUE;
        }
    case WM_DESTROY:
        if (state && state->ui_font)
        {
            DeleteObject(state->ui_font);
            state->ui_font = nullptr;
        }
        return TRUE;
    case WM_SETTINGCHANGE:
        if (Theme::UpdateFromSystem())
        {
            Theme::Current().ApplyToWindow(dlg);
            Theme::Current().ApplyToChildren(dlg);
            InvalidateRect(dlg, nullptr, TRUE);
        }
        return TRUE;
    case WM_ERASEBKGND:
        {
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
    case WM_CTLCOLORBTN:
        {
            HDC hdc = reinterpret_cast<HDC>(wparam);
            HWND target = reinterpret_cast<HWND>(lparam);
            int type = CTLCOLOR_STATIC;
            if (msg == WM_CTLCOLOREDIT)
            {
                type = CTLCOLOR_EDIT;
            }
            else if (msg == WM_CTLCOLORLISTBOX)
            {
                type = CTLCOLOR_LISTBOX;
            }
            else if (msg == WM_CTLCOLORBTN)
            {
                type = CTLCOLOR_BTN;
            }
            else if (msg == WM_CTLCOLORDLG)
            {
                type = CTLCOLOR_DLG;
            }
            return reinterpret_cast<INT_PTR>(Theme::Current().ControlColor(hdc, target, type));
        }
    case WM_COMMAND:
        {
            if (!state)
            {
                return TRUE;
            }
            int id = LOWORD(wparam);
            int code = HIWORD(wparam);
            if (code == CBN_SELCHANGE && (id == IDC_COMPARE_LEFT_SOURCE || id == IDC_COMPARE_RIGHT_SOURCE))
            {
                bool left = id == IDC_COMPARE_LEFT_SOURCE;
                HWND combo = GetDlgItem(dlg, id);
                int sel = combo ? static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0)) : 0;
                CompareSourceType type = CompareSourceFromIndex(sel);
                ToggleCompareControls(dlg, left, type);
                SetDialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE, L"");
                SetDialogText(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY, L"");
                return TRUE;
            }
            if (code == BN_CLICKED && (id == IDC_COMPARE_LEFT_BROWSE || id == IDC_COMPARE_RIGHT_BROWSE))
            {
                bool left = id == IDC_COMPARE_LEFT_BROWSE;
                HWND browse_source = GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_SOURCE : IDC_COMPARE_RIGHT_SOURCE);
                const CompareSourceType browse_type = CompareSourceFromIndex(
                    browse_source ? static_cast<int>(SendMessageW(browse_source, CB_GETCURSEL, 0, 0)) : 0
                );
                std::wstring path;
                if (browse_type == CompareSourceType::kNetwork)
                {
                    if (FAILED(win32::ChooseComputer(dlg, &path)) || path.empty())
                    {
                        return TRUE;
                    }
                    SetDialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE, path);
                    return TRUE;
                }
                if (!ui::PromptOpenFile(dlg, browse_type == CompareSourceType::kOfflineHive ? L"Registry Hive Files\0*.*\0\0" : ui::kRegFileFilter, &path))
                {
                    return TRUE;
                }
                SetDialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE, path);
                if (browse_type == CompareSourceType::kOfflineHive)
                {
                    return TRUE;
                }
                regfile::Document data;
                std::wstring error;
                if (regfile::Load(path, &data, &error))
                {
                    std::vector<std::wstring> keys = ExtractRegFileKeys(data);
                    if (keys.empty())
                    {
                        ui::ShowError(dlg, L"No registry keys were found in the .reg file.");
                        return TRUE;
                    }
                    HWND combo = GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY);
                    PopulateCombo(combo, keys);
                    if (!keys.empty())
                    {
                        SendMessageW(combo, CB_SETCURSEL, 0, 0);
                        SetDialogText(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY, keys.front());
                    }
                }
                else if (!error.empty())
                {
                    ui::ShowError(dlg, error);
                }
                return TRUE;
            }
            if (id == IDOK)
            {
                CompareDialogResult result;
                auto read_side = [&](bool left, CompareDialogSelection* out) -> bool {
                    out->recursive = IsDlgButtonChecked(dlg, left ? IDC_COMPARE_LEFT_RECURSIVE : IDC_COMPARE_RIGHT_RECURSIVE) == BST_CHECKED;
                    HWND source_combo = GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_SOURCE : IDC_COMPARE_RIGHT_SOURCE);
                    int source_index = source_combo ? static_cast<int>(SendMessageW(source_combo, CB_GETCURSEL, 0, 0)) : 0;
                    out->type = CompareSourceFromIndex(source_index);
                    out->key_path =
                        TrimWhitespace(ReadComboText(GetDlgItem(dlg, left ? IDC_COMPARE_LEFT_KEY : IDC_COMPARE_RIGHT_KEY)));
                    if (out->type == CompareSourceType::kRegistry || out->type == CompareSourceType::kNetwork)
                    {
                        if (out->key_path.empty())
                        {
                            ui::ShowError(dlg, L"Registry key is required.");
                            return false;
                        }
                        if (out->type == CompareSourceType::kNetwork)
                        {
                            out->file_path = TrimWhitespace(
                                util::DialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE)
                            );
                            if (out->file_path.empty())
                            {
                                ui::ShowError(dlg, L"Computer name is required.");
                                return false;
                            }
                        }
                        return true;
                    }
                    out->file_path =
                        TrimWhitespace(util::DialogText(dlg, left ? IDC_COMPARE_LEFT_FILE : IDC_COMPARE_RIGHT_FILE));
                    if (out->file_path.empty())
                    {
                        ui::ShowError(dlg, out->type == CompareSourceType::kOfflineHive ? L"Hive file path is required." : L"Registry file path is required.");
                        return false;
                    }
                    if (out->type == CompareSourceType::kOfflineHive)
                    {
                        return true;
                    }
                    regfile::Document data;
                    std::wstring error;
                    if (!regfile::Load(out->file_path, &data, &error))
                    {
                        ui::ShowError(dlg, error.empty() ? L"Failed to read registry file." : error);
                        return false;
                    }
                    std::vector<std::wstring> keys = ExtractRegFileKeys(data);
                    if (keys.empty())
                    {
                        ui::ShowError(dlg, L"No registry keys were found in the .reg file.");
                        return false;
                    }
                    if (out->key_path.empty())
                    {
                        out->key_path = keys.front();
                    }
                    std::wstring key_lower = ToLower(out->key_path);
                    bool found = false;
                    for (const auto& key : keys)
                    {
                        if (util::EqualsInsensitive(key, out->key_path))
                        {
                            found = true;
                            break;
                        }
                        std::wstring key_check = ToLower(key);
                        if (StartsWithInsensitive(key_check, key_lower) || StartsWithInsensitive(key_lower, key_check))
                        {
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        ui::ShowError(dlg, L"The selected key path wasn't found in the .reg file.");
                        return false;
                    }
                    return true;
                };
                if (!read_side(true, &result.left))
                {
                    return TRUE;
                }
                if (!read_side(false, &result.right))
                {
                    return TRUE;
                }
                if (IsDlgButtonChecked(dlg, IDC_COMPARE_SHOW_MATCHING) == BST_CHECKED)
                {
                    result.filter = search::compare::RowFilter::kMatches;
                }
                else if (IsDlgButtonChecked(dlg, IDC_COMPARE_SHOW_BOTH) == BST_CHECKED)
                {
                    result.filter = search::compare::RowFilter::kAll;
                }
                state->data.left = result.left;
                state->data.right = result.right;
                state->data.filter = result.filter;
                EndDialog(dlg, IDOK);
                return TRUE;
            }
            if (id == IDCANCEL)
            {
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
} // namespace

bool ShowCompareDialog(HWND owner, const CompareDialogDefaults& defaults, CompareDialogResult* out)
{
    if (!out)
    {
        return false;
    }
    CompareDialogState state;
    state.data = defaults;
    INT_PTR result = DialogBoxParamW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_COMPARE), owner, CompareDialogProc, reinterpret_cast<LPARAM>(&state));
    if (result != IDOK)
    {
        return false;
    }
    out->left = state.data.left;
    out->right = state.data.right;
    out->filter = state.data.filter;
    return true;
}


} // namespace regkit::command_detail
