// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

namespace {

bool ValueNameExists(const RegistryNode& node, const std::wstring& name) {
  ValueEntry entry;
  return RegistryStore::QueryValue(node, name, &entry);
}

bool KeyNameExists(const RegistryNode& parent, const std::wstring& name) {
  KeyInfo info = {};
  return RegistryStore::QueryKeyInfo(MakeChildNode(parent, name), &info);
}

void ReportNameTaken(HWND owner, const wchar_t* message, const wchar_t* title,
                     const std::wstring& name) {
  ui::PromptKeyChoice(owner, message, name, title, L"", L"", L"OK");
}

constexpr int kDividerProbeWidth = 32;
constexpr int kDividerProbeHeight = 16;



COLORREF HeaderDividerColor(HWND header) {
  COLORREF color = Theme::Current().BorderColor();
  HTHEME theme = OpenThemeData(header, VSCLASS_HEADER);
  if (!theme) {
    return color;
  }
  HDC screen = GetDC(nullptr);
  if (HDC mem = CreateCompatibleDC(screen)) {
    if (HBITMAP bitmap =
            CreateCompatibleBitmap(screen, kDividerProbeWidth, kDividerProbeHeight)) {
      HGDIOBJ previous = SelectObject(mem, bitmap);
      RECT probe = {0, 0, kDividerProbeWidth, kDividerProbeHeight};
      if (SUCCEEDED(DrawThemeBackground(theme, mem, HP_HEADERITEM, HIS_NORMAL,
                                        &probe, nullptr))) {
        const COLORREF edge =
            GetPixel(mem, kDividerProbeWidth - 1, kDividerProbeHeight / 2);
        if (edge != CLR_INVALID) {
          color = edge;
        }
      }
      SelectObject(mem, previous);
      DeleteObject(bitmap);
    }
    DeleteDC(mem);
  }
  ReleaseDC(nullptr, screen);
  CloseThemeData(theme);
  return color;
}

void FormatCellFileTime(const FILETIME& filetime, wchar_t* buffer, int capacity) {
  buffer[0] = L'\0';
  if (filetime.dwLowDateTime == 0 && filetime.dwHighDateTime == 0) {
    return;
  }
  FILETIME local = {};
  SYSTEMTIME st = {};
  if (!FileTimeToLocalFileTime(&filetime, &local) ||
      !FileTimeToSystemTime(&local, &st)) {
    return;
  }
  swprintf_s(buffer, static_cast<size_t>(capacity), L"%d/%d/%d %d:%02d",
             st.wMonth, st.wDay, st.wYear, st.wHour, st.wMinute);
}

} // namespace


LRESULT MainWindow::Impl::HandleNotification(LPARAM lparam) {
  auto* header = reinterpret_cast<NMHDR*>(lparam);
  if (!header) {
    return 0;
  }
  if (header->code == TTN_GETDISPINFOW || header->code == TTN_NEEDTEXTW ||
      header->code == TTN_SHOW) {
    return HandleTooltipNotification(header, lparam);
  }
  if (header->hwndFrom == toolbar_.hwnd() || header->hwndFrom == value_grid_toolbar_ ||
      header->hwndFrom == search_grid_toolbar_) {
    return HandleToolbarNotification(header, lparam);
  }
  if (header->hwndFrom == tab_) {
    return HandleTabNotification(header, lparam);
  }
  if (header->hwndFrom == browse_.tree().hwnd()) {
    return HandleTreeNotification(header, lparam);
  }
  if (header->hwndFrom == browse_.values().hwnd()) {
    return HandleValueNotification(header, lparam);
  }
  if (header->hwndFrom == history_list_) {
    return HandleHistoryNotification(header, lparam);
  }
  if (header->hwndFrom == search_results_list_) {
    return HandleSearchNotification(header, lparam);
  }
  if (header->hwndFrom == ListView_GetHeader(browse_.values().hwnd()) ||
      header->hwndFrom == ListView_GetHeader(history_list_) ||
      header->hwndFrom == ListView_GetHeader(search_results_list_)) {
    return HandleHeaderNotification(header, lparam);
  }
  return 0;
}

bool MainWindow::Impl::ValueCellTooltipText(std::wstring* out) {
  HWND list = browse_.values().hwnd();
  if (!out || !list) {
    return false;
  }
  POINT pt = {};
  GetCursorPos(&pt);
  ScreenToClient(list, &pt);
  LVHITTESTINFO hit = {};
  hit.pt = pt;
  const int item = ListView_SubItemHitTest(list, &hit);
  if (item < 0) {
    return false;
  }
  ListRow* row = browse_.values().MutableRowAt(item);
  if (!row) {
    return false;
  }
  const int subitem = MappedSubItem(value_column_subitems_, hit.iSubItem);
  const std::wstring& text = ValueRowFieldText(*row, subitem);
  RECT cell = {};
  const bool measured =
      hit.iSubItem == 0
          ? ListView_GetItemRect(list, item, &cell, LVIR_LABEL) != FALSE
          : ListView_GetSubItemRect(list, item, hit.iSubItem, LVIR_BOUNDS, &cell) != FALSE;
  const int available = static_cast<int>(cell.right - cell.left) - kCellTooltipPadding;
  if (!measured || text.empty() || available <= 0 ||
      !CellTextIsClipped(list, text, available)) {
    return false;
  }
  size_t limit = std::min(text.size(), kValueTooltipTextLimit);
  size_t lines = 0;
  for (size_t i = 0; i < limit; ++i) {
    if (text[i] == L'\n' && ++lines == kValueTooltipLineLimit) {
      limit = i;
      break;
    }
  }
  out->assign(text, 0, limit);
  if (limit < text.size()) {
    out->append(L"...");
  }
  return true;
}

std::wstring MainWindow::Impl::SearchCellFieldText(const search::Result& result,
                                                  int subitem) const {
  switch (subitem) {
  case 0:
    return result.key_path;
  case 1:
    return std::wstring(search::DisplayName(result));
  case 2:
    return search::TypeText(result);
  case 3:
    return result.data_text;
  case 4:
    return search::IsKeyRow(result) ||
                   result.kind == search::ResultKind::kTraceValue
               ? std::wstring()
               : std::to_wstring(result.data_size);
  default:
    return std::wstring();
  }
}

bool MainWindow::Impl::SearchCellTooltipText(std::wstring* out) {
  HWND list = search_results_list_;
  if (!out || !list) {
    return false;
  }
  POINT pt = {};
  GetCursorPos(&pt);
  ScreenToClient(list, &pt);
  LVHITTESTINFO hit = {};
  hit.pt = pt;
  const int item = ListView_SubItemHitTest(list, &hit);
  if (item < 0) {
    return false;
  }
  const search::Result* result = SearchResultAt(item);
  if (!result) {
    return false;
  }
  const std::wstring text = SearchCellFieldText(
      *result, MappedSubItem(search_column_subitems_, hit.iSubItem));
  RECT cell = {};
  const bool measured =
      hit.iSubItem == 0
          ? ListView_GetItemRect(list, item, &cell, LVIR_LABEL) != FALSE
          : ListView_GetSubItemRect(list, item, hit.iSubItem, LVIR_BOUNDS, &cell) != FALSE;
  const int available = static_cast<int>(cell.right - cell.left) - kCellTooltipPadding;
  if (!measured || text.empty() || available <= 0 ||
      !CellTextIsClipped(list, text, available)) {
    return false;
  }
  out->assign(text, 0, std::min(text.size(), kValueTooltipTextLimit));
  if (out->size() < text.size()) {
    out->append(L"...");
  }
  return true;
}

bool MainWindow::Impl::ListCellTooltipText(std::wstring* out) {
  POINT pt = {};
  GetCursorPos(&pt);
  HWND under = WindowFromPoint(pt);
  if (under == search_results_list_) {
    return SearchCellTooltipText(out);
  }
  if (under == browse_.values().hwnd()) {
    return ValueCellTooltipText(out);
  }
  return false;
}

LRESULT MainWindow::Impl::HandleTooltipNotification(NMHDR* header, LPARAM lparam) {
  if (header->code == TTN_SHOW && header->hwndFrom == value_tooltip_) {
    RECT tip = {};
    POINT pt = {};
    MONITORINFO monitor = {};
    monitor.cbSize = sizeof(monitor);
    if (!GetWindowRect(value_tooltip_, &tip) || !GetCursorPos(&pt) ||
        !GetMonitorInfoW(MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST), &monitor)) {
      return 0;
    }
    const int width = tip.right - tip.left;
    const int height = tip.bottom - tip.top;
    int x = pt.x + kValueTooltipCursorGap;
    int y = pt.y + kValueTooltipCursorGap;
    if (x + width > monitor.rcWork.right) {
      x = std::max<LONG>(monitor.rcWork.left, pt.x - kValueTooltipCursorGap - width);
    }
    if (y + height > monitor.rcWork.bottom) {
      y = std::max<LONG>(monitor.rcWork.top, pt.y - kValueTooltipCursorGap - height);
    }
    SetWindowPos(value_tooltip_, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE);
    return TRUE;
  }
  if (header->code == TTN_GETDISPINFOW || header->code == TTN_NEEDTEXTW) {
    auto* info = reinterpret_cast<LPTOOLTIPTEXTW>(lparam);
    if (info && header->hwndFrom == value_tooltip_) {
      value_tooltip_text_.clear();
      if (ListCellTooltipText(&value_tooltip_text_)) {
        info->lpszText = value_tooltip_text_.data();
      } else {
        info->lpszText = const_cast<wchar_t*>(L"");
      }
      return 0;
    }
    if (info) {
      int command_id = static_cast<int>(info->hdr.idFrom);
      std::wstring tip = CommandTooltipText(command_id);
      if (!tip.empty()) {
        std::wstring shortcut = CommandShortcutText(command_id);
        if (!shortcut.empty()) {
          tip.append(L" (");
          tip.append(shortcut);
          tip.append(L")");
        }
        static std::wstring tip_storage;
        tip_storage = tip;
        info->lpszText = const_cast<wchar_t*>(tip_storage.c_str());
        return 0;
      }
    }
  }
  return 0;
}

HBRUSH MainWindow::Impl::ValueHeaderSurfaceBrush() const {
  const COLORREF color = browse_.values().hwnd()
                             ? ListView_GetBkColor(browse_.values().hwnd())
                             : CLR_NONE;
  if (color == CLR_NONE || color == CLR_DEFAULT) {
    return Theme::Current().PanelBrush();
  }
  return appearance::CachedBrush(color);
}

LRESULT MainWindow::Impl::HandleToolbarNotification(NMHDR* header, LPARAM lparam) {
  const bool is_grid_bar = header->hwndFrom == value_grid_toolbar_ ||
                           header->hwndFrom == search_grid_toolbar_;
  if ((header->hwndFrom == toolbar_.hwnd() || is_grid_bar) &&
      header->code == NM_CUSTOMDRAW) {
    auto* draw = reinterpret_cast<NMTBCUSTOMDRAW*>(lparam);
    if (!draw || (!is_grid_bar && !Theme::UseDarkMode())) {
      return CDRF_DODEFAULT;
    }
    HWND bar = header->hwndFrom;
    const Theme& theme = Theme::Current();
    switch (draw->nmcd.dwDrawStage) {
    case CDDS_PREPAINT:
      FillRect(draw->nmcd.hdc, &draw->nmcd.rc,
               is_grid_bar ? ValueHeaderSurfaceBrush() : theme.BackgroundBrush());
      return CDRF_NOTIFYITEMDRAW;
    case CDDS_ITEMPREPAINT: {
      bool is_separator = false;
      int command_id = static_cast<int>(draw->nmcd.dwItemSpec);
      int index = static_cast<int>(SendMessageW(bar, TB_COMMANDTOINDEX, command_id, 0));
      if (index >= 0) {
        TBBUTTON button = {};
        if (SendMessageW(bar, TB_GETBUTTON, index, reinterpret_cast<LPARAM>(&button))) {
          is_separator = (button.fsStyle & BTNS_SEP) != 0;
        }
      }
      if (is_separator) {
        return CDRF_DODEFAULT;
      }

      POINT cursor = {};
      GetCursorPos(&cursor);
      ScreenToClient(bar, &cursor);
      bool is_hovered = ((draw->nmcd.uItemState & CDIS_HOT) == CDIS_HOT) || PtInRect(&draw->nmcd.rc, cursor);

      draw->hbrMonoDither = theme.BackgroundBrush();
      draw->hbrLines = theme.BackgroundBrush();
      draw->hpenLines = appearance::CachedPen(theme.BorderColor(), 1);
      draw->clrText = theme.TextColor();
      draw->clrTextHighlight = theme.TextColor();
      draw->clrBtnFace = theme.BackgroundColor();
      draw->clrBtnHighlight = theme.SurfaceColor();
      draw->clrHighlightHotTrack = theme.HoverColor();
      draw->nStringBkMode = TRANSPARENT;
      draw->nHLStringBkMode = TRANSPARENT;

      if (is_hovered) {
        DrawToolbarButtonBackground(draw->nmcd.hdc, draw->nmcd.rc, theme.HoverColor(),
                                    is_grid_bar ? theme.HoverColor() : theme.BorderColor());
        draw->nmcd.uItemState &= ~(CDIS_HOT | CDIS_CHECKED);
      } else if ((draw->nmcd.uItemState & CDIS_CHECKED) == CDIS_CHECKED) {
        DrawToolbarButtonBackground(draw->nmcd.hdc, draw->nmcd.rc, theme.SurfaceColor(),
                                    is_grid_bar ? theme.SurfaceColor() : theme.BorderColor());
        draw->nmcd.uItemState &= ~CDIS_CHECKED;
      }

      LRESULT lr = TBCDRF_USECDCOLORS;
      if ((draw->nmcd.uItemState & CDIS_SELECTED) == CDIS_SELECTED) {
        lr |= TBCDRF_NOBACKGROUND;
      }
      return lr;
    }
    default:
      break;
    }
    return CDRF_DODEFAULT;
  }
  return 0;
}

LRESULT MainWindow::Impl::HandleTabNotification(NMHDR* header, LPARAM lparam) {
  (void)lparam;
  if (header->hwndFrom == tab_ && header->code == TCN_SELCHANGING) {
    if (!suppress_tab_change_ && tab_) {
      int current = TabCtrl_GetCurSel(tab_);
      CaptureRegistryTabState(current);
    }
    return 0;
  }
  if (header->hwndFrom == tab_ && header->code == TCN_SELCHANGE) {
    if (suppress_tab_change_) {
      ApplyViewVisibility();
      UpdateSearchResultsView();
      UpdateStatus();
      return 0;
    }
    int sel = TabCtrl_GetCurSel(tab_);
    ApplyTabSelection(sel);
    ApplyViewVisibility();
    UpdateSearchResultsView();
    UpdateStatus();
    return 0;
  }
  return 0;
}

LRESULT MainWindow::Impl::HandleTreeNotification(NMHDR* header, LPARAM lparam) {
  if (header->hwndFrom == browse_.tree().hwnd()) {
    if (header->code == TVN_ITEMEXPANDINGW) {
      browse_.tree().OnItemExpanding(reinterpret_cast<NMTREEVIEWW*>(lparam));
      return 0;
    }
    if (header->code == TVN_ITEMEXPANDEDW) {
      if (!jump_ui_batch_active_) {
        MarkTreeStateDirty();
      }
      return 0;
    }
    if (header->code == TVN_BEGINLABELEDITW) {
      if (read_only_) {
        return TRUE;
      }
      auto* disp = reinterpret_cast<NMTVDISPINFOW*>(lparam);
      if (!disp) {
        return TRUE;
      }
      RegistryNode* node = browse_.tree().NodeFromItem(disp->item.hItem);
      if (!node || node->subkey.empty()) {
        return TRUE;
      }
      HWND edit = TreeView_GetEditControl(browse_.tree().hwnd());
      if (edit) {
        Theme::Current().ApplyToWindow(edit);
        Theme::Current().ApplyToChildren(edit);
        const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
        SetWindowTheme(edit, theme_name, nullptr);
      }
      return FALSE;
    }
    if (header->code == TVN_ENDLABELEDITW) {
      if (read_only_) {
        return FALSE;
      }
      auto* disp = reinterpret_cast<NMTVDISPINFOW*>(lparam);
      if (!disp || !disp->item.pszText) {
        return FALSE;
      }
      RegistryNode* node = browse_.tree().NodeFromItem(disp->item.hItem);
      if (!node || node->subkey.empty()) {
        return FALSE;
      }
      std::wstring new_name = TrimWhitespace(disp->item.pszText);
      std::wstring old_name = LeafName(*node);
      if (new_name.empty() || _wcsicmp(new_name.c_str(), old_name.c_str()) == 0) {
        return FALSE;
      }
      RegistryNode rename_parent = *node;
      size_t rename_sep = rename_parent.subkey.rfind(L'\\');
      rename_parent.subkey = (rename_sep == std::wstring::npos)
                                 ? L""
                                 : rename_parent.subkey.substr(0, rename_sep);
      if (KeyNameExists(rename_parent, new_name)) {
        ReportNameTaken(hwnd_, L"A key with this name already exists:",
                        L"Rename key", new_name);
        return FALSE;
      }
      if (!RegistryStore::RenameKey(*node, new_name)) {
        ui::ShowError(hwnd_, L"Failed to rename key.");
        return FALSE;
      }
      UpdateLeafName(node, new_name);
      if (browse_.current_node() && SameNode(*browse_.current_node(), *node)) {
        UpdateAddressBar(browse_.current_node());
      }
      AppendHistoryEntry(L"Rename key", old_name, new_name);
      MarkOfflineDirty();
      RegistryNode parent = *node;
      if (!parent.subkey.empty()) {
        size_t pos = parent.subkey.rfind(L'\\');
        parent.subkey = (pos == std::wstring::npos) ? L"" : parent.subkey.substr(0, pos);
      }
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kRenameKey;
      op.node = parent;
      op.name = old_name;
      op.new_name = new_name;
      PushUndo(std::move(op));
      RefreshTreeSelection();
      UpdateValueListForNode(browse_.current_node());
      return TRUE;
    }
    if (header->code == TVN_SELCHANGEDW) {
      auto* info = reinterpret_cast<NMTREEVIEWW*>(lparam);
      RegistryNode* previous_node = browse_.current_node();
      RegistryNode* node = browse_.tree().OnSelectionChanged(info);
      if (startup_tree_restore_pending_ && !applying_startup_tree_restore_) {
        startup_tree_restore_pending_ = false;
        tree_state_restored_ = true;
      }
      browse_.set_current_node(node);
      if (previous_node && node && SameNode(*previous_node, *node)) {
        return 0;
      }
      if (!jump_ui_batch_active_) {
        ApplyTreeSelectionEffects(node);
      }
      return 0;
    }
    if (header->code == NM_CUSTOMDRAW) {
      if (!Theme::UseDarkMode()) {
        return CDRF_DODEFAULT;
      }
      auto* draw = reinterpret_cast<NMTVCUSTOMDRAW*>(lparam);
      if (!draw) {
        return CDRF_DODEFAULT;
      }
      switch (draw->nmcd.dwDrawStage) {
      case CDDS_PREPAINT:
        return CDRF_NOTIFYITEMDRAW;
      case CDDS_ITEMPREPAINT: {
        if (draw->nmcd.uItemState & CDIS_SELECTED) {
          return CDRF_DODEFAULT;
        }
        const Theme& theme = Theme::Current();
        bool hot = (draw->nmcd.uItemState & CDIS_HOT) != 0;
        if (hot) {
          draw->clrText = theme.TextColor();
          draw->clrTextBk = theme.HoverColor();
        } else {
          draw->clrText = theme.TextColor();
          draw->clrTextBk = theme.PanelColor();
        }
        return CDRF_NEWFONT;
      }
      default:
        break;
      }
    }
  }
  return 0;
}

LRESULT MainWindow::Impl::HandleHeaderNotification(NMHDR* header, LPARAM lparam) {
  HWND value_header = ListView_GetHeader(browse_.values().hwnd());
  HWND history_header = ListView_GetHeader(history_list_);
  HWND search_header = ListView_GetHeader(search_results_list_);

  if (header->hwndFrom == value_header && (header->code == HDN_ENDTRACKW || header->code == HDN_ENDTRACKA || header->code == HDN_ITEMCHANGEDW || header->code == HDN_ITEMCHANGEDA)) {
    auto* info = reinterpret_cast<NMHEADERW*>(lparam);
    if (info && info->iItem >= 0 && info->pitem && (info->pitem->mask & HDI_WIDTH)) {
      int subitem = GetListViewColumnSubItem(browse_.values().hwnd(), info->iItem);
      if (subitem >= 0 && static_cast<size_t>(subitem) < browse_.columns().widths.size()) {
        browse_.columns().widths[static_cast<size_t>(subitem)] = info->pitem->cxy;
        if (header->code == HDN_ENDTRACKW || header->code == HDN_ENDTRACKA) {
          SaveSettings();
        }
      }


      InvalidateListViewColumn(browse_.values().hwnd(), info->iItem);
      InvalidateListViewTail(browse_.values().hwnd());
    }
  }
  if (header->hwndFrom == history_header && (header->code == HDN_ENDTRACKW || header->code == HDN_ENDTRACKA || header->code == HDN_ITEMCHANGEDW || header->code == HDN_ITEMCHANGEDA)) {
    auto* info = reinterpret_cast<NMHEADERW*>(lparam);
    if (info && info->iItem >= 0 && info->pitem && (info->pitem->mask & HDI_WIDTH)) {
      int subitem = GetListViewColumnSubItem(history_list_, info->iItem);
      if (subitem >= 0 && static_cast<size_t>(subitem) < history_column_widths_.size()) {
        history_column_widths_[static_cast<size_t>(subitem)] = info->pitem->cxy;
      }
      InvalidateListViewColumn(history_list_, info->iItem);
      InvalidateListViewTail(history_list_);
    }
  }
  if (header->hwndFrom == search_header && (header->code == HDN_ENDTRACKW || header->code == HDN_ENDTRACKA || header->code == HDN_ITEMCHANGEDW || header->code == HDN_ITEMCHANGEDA)) {
    auto* info = reinterpret_cast<NMHEADERW*>(lparam);
    if (info && info->iItem >= 0 && info->pitem && (info->pitem->mask & HDI_WIDTH)) {
      int subitem = GetListViewColumnSubItem(search_results_list_, info->iItem);
      bool compare = IsCompareTabSelected();
      auto& widths = compare ? compare_column_widths_ : search_column_widths_;
      if (subitem >= 0 && static_cast<size_t>(subitem) < widths.size()) {
        widths[static_cast<size_t>(subitem)] = info->pitem->cxy;
      }
      InvalidateListViewColumn(search_results_list_, info->iItem);
      InvalidateListViewTail(search_results_list_);
    }
  }
  return 0;
}



bool MainWindow::Impl::PaintHeaderItem(HWND header, NMCUSTOMDRAW* draw) {
  HTHEME theme = OpenThemeData(header, VSCLASS_HEADER);
  if (!theme) {
    return false;
  }
  int state = HIS_NORMAL;
  if (draw->uItemState & CDIS_SELECTED) {
    state = HIS_PRESSED;
  } else if (draw->uItemState & CDIS_HOT) {
    state = HIS_HOT;
  }
  if (state == HIS_NORMAL) {
    if (grid_line_color_ == CLR_INVALID) {
      grid_line_color_ = HeaderDividerColor(header);
    }
    FillRect(draw->hdc, &draw->rc,
             appearance::CachedBrush(ListView_GetBkColor(GetParent(header))));
    RECT divider = {draw->rc.right - 1, draw->rc.top, draw->rc.right,
                    draw->rc.bottom};
    FillRect(draw->hdc, &divider, appearance::CachedBrush(grid_line_color_));
  } else {
    DrawThemeBackground(theme, draw->hdc, HP_HEADERITEM, state, &draw->rc, nullptr);
  }

  wchar_t text[128] = {};
  HDITEMW item = {};
  item.mask = HDI_TEXT | HDI_FORMAT;
  item.pszText = text;
  item.cchTextMax = static_cast<int>(_countof(text));
  Header_GetItem(header, static_cast<int>(draw->dwItemSpec), &item);

  if (item.fmt & (HDF_SORTUP | HDF_SORTDOWN)) {
    const int arrow_state = (item.fmt & HDF_SORTUP) ? HSAS_SORTEDUP : HSAS_SORTEDDOWN;
    SIZE size = {};
    if (SUCCEEDED(GetThemePartSize(theme, draw->hdc, HP_HEADERSORTARROW,
                                   arrow_state, nullptr, TS_TRUE, &size))) {
      RECT arrow = draw->rc;
      arrow.bottom = arrow.top + size.cy;
      DrawThemeBackground(theme, draw->hdc, HP_HEADERSORTARROW, arrow_state,
                          &arrow, nullptr);
    }
  }

  RECT text_rect = draw->rc;
  text_rect.left += kHeaderTextPadding;
  text_rect.right -= kHeaderTextPadding;
  UINT format = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
  if (item.fmt & HDF_RIGHT) {
    format |= DT_RIGHT;
  } else if (item.fmt & HDF_CENTER) {
    format |= DT_CENTER;
  }
  const int saved = SaveDC(draw->hdc);
  SetBkMode(draw->hdc, TRANSPARENT);
  SetTextColor(draw->hdc, Theme::Current().TextColor());
  DrawTextW(draw->hdc, text, -1, &text_rect, format);
  RestoreDC(draw->hdc, saved);
  CloseThemeData(theme);
  return true;
}

LRESULT MainWindow::Impl::HandleValueNotification(NMHDR* header, LPARAM lparam) {
  if (header->hwndFrom == browse_.values().hwnd() &&
      header->code == LVN_ODCACHEHINT) {
    auto* hint = reinterpret_cast<NMLVCACHEHINT*>(lparam);
    if (hint) {
      QueueValuePreviews(hint->iFrom, hint->iTo);
    }
    return 0;
  }
  if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_GETDISPINFOW) {
    auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
    ListRow* mutable_row = browse_.values().MutableRowAt(disp->item.iItem);
    const ListRow* row = mutable_row;
    if (!row) {
      if (disp->item.mask & LVIF_TEXT) {
        if (disp->item.pszText && disp->item.cchTextMax > 0) {
          disp->item.pszText[0] = L'\0';
        }
      }
      if (disp->item.mask & LVIF_IMAGE) {
        disp->item.iImage = 0;
      }
      return 0;
    }
    if (disp->item.mask & LVIF_TEXT) {
      const int subitem = MappedSubItem(value_column_subitems_, disp->item.iSubItem);
      if (subitem == kValueColData && mutable_row && !mutable_row->data_ready &&
          !value_preview_request_posted_) {
        value_preview_request_posted_ = true;
        if (!PostMessageW(hwnd_, frame::message_id::kValuePreviewRequest,
                          static_cast<WPARAM>(disp->item.iItem), 0)) {
          value_preview_request_posted_ = false;
        }
      }
      const std::wstring& text = ValueRowFieldText(*row, subitem);
      if (text.size() > kCellTextDrawLimit && disp->item.pszText &&
          disp->item.cchTextMax > 0) {
        lstrcpynW(disp->item.pszText, text.c_str(), disp->item.cchTextMax);
      } else {
        disp->item.pszText = const_cast<wchar_t*>(text.c_str());
      }
    }
    if (disp->item.mask & LVIF_IMAGE) {
      disp->item.iImage = row->image_index;
    }
    return 0;
  }
  if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_BEGINLABELEDITW) {
    if (read_only_) {
      return TRUE;
    }
    auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
    if (!disp) {
      return TRUE;
    }
    const ListRow* row = browse_.values().RowAt(disp->item.iItem);
    if (!row || row->extra.empty() || (row->kind != rowkind::kValue && row->kind != rowkind::kKey)) {
      return TRUE;
    }
    HWND edit = ListView_GetEditControl(browse_.values().hwnd());
    if (edit) {
      Theme::Current().ApplyToWindow(edit);
      Theme::Current().ApplyToChildren(edit);
      const wchar_t* theme_name = Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer";
      SetWindowTheme(edit, theme_name, nullptr);
    }
    return FALSE;
  }
  if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_ENDLABELEDITW) {
    if (read_only_) {
      return FALSE;
    }
    auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
    if (!disp || !disp->item.pszText || !browse_.current_node()) {
      return FALSE;
    }
    const ListRow* row = browse_.values().RowAt(disp->item.iItem);
    if (!row || row->extra.empty()) {
      return FALSE;
    }
    std::wstring new_name = TrimWhitespace(disp->item.pszText);
    std::wstring old_name = row->extra;
    if (new_name.empty() || _wcsicmp(new_name.c_str(), old_name.c_str()) == 0) {
      return FALSE;
    }
    if (row->kind == rowkind::kKey) {
      RegistryNode child = MakeChildNode(*browse_.current_node(), old_name);
      if (KeyNameExists(*browse_.current_node(), new_name)) {
        ReportNameTaken(hwnd_, L"A key with this name already exists:",
                        L"Rename key", new_name);
        return FALSE;
      }
      if (!RegistryStore::RenameKey(child, new_name)) {
        ui::ShowError(hwnd_, L"Failed to rename key.");
        return FALSE;
      }
      AppendHistoryEntry(L"Rename key " + old_name, old_name, new_name);
      MarkOfflineDirty();
      changes::UndoOperation op;
      op.type = changes::UndoOperation::Type::kRenameKey;
      op.node = *browse_.current_node();
      op.name = old_name;
      op.new_name = new_name;
      PushUndo(std::move(op));
      RefreshTreeSelection();
      UpdateValueListForNode(browse_.current_node());
      return TRUE;
    }
    if (ValueNameExists(*browse_.current_node(), new_name)) {
      ReportNameTaken(hwnd_, L"A value with this name already exists:",
                      L"Rename value", new_name);
      return FALSE;
    }
    if (!RegistryStore::RenameValue(*browse_.current_node(), old_name, new_name)) {
      ui::ShowError(hwnd_, L"Failed to rename value.");
      return FALSE;
    }
    AppendValueHistoryEntry(L"Rename value " + old_name, old_name, new_name,
                            *browse_.current_node(), new_name,
                            HistoryEntry::RevertKind::kNone);
    MarkOfflineDirty();
    changes::UndoOperation op;
    op.type = changes::UndoOperation::Type::kRenameValue;
    op.node = *browse_.current_node();
    op.name = old_name;
    op.new_name = new_name;
    PushUndo(std::move(op));
    if (EqualsInsensitive(old_name, appended_value_name_)) {
      ListRow* updated = browse_.values().MutableRowAt(disp->item.iItem);
      if (updated) {
        updated->name = new_name;
        updated->extra = new_name;
        browse_.values().InvalidateFilterCache(updated);
        ListView_RedrawItems(browse_.values().hwnd(), disp->item.iItem,
                             disp->item.iItem);
        browse_.SelectValue(new_name);
      }
      appended_value_name_.clear();
      return TRUE;
    }
    UpdateValueListForNode(browse_.current_node());
    retained_value_name_ = new_name;
    retained_value_key_path_ = registry_path::Build(*browse_.current_node());
    return TRUE;
  }
  if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_ITEMCHANGED) {
    auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
    if (!updating_value_list_ && info &&
        ((info->uOldState ^ info->uNewState) & LVIS_SELECTED) != 0) {
      UpdateStatus();
    }
    return 0;
  }
  if (header->hwndFrom == browse_.values().hwnd() && header->code == LVN_COLUMNCLICK) {
    auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
    if (info) {
      SortValueList(info->iSubItem, true);
    }
    return 0;
  }
  if (header->hwndFrom == browse_.values().hwnd() && (header->code == NM_DBLCLK || header->code == LVN_ITEMACTIVATE)) {
    auto* activate = reinterpret_cast<NMITEMACTIVATE*>(lparam);
    if (activate && activate->iItem >= 0 && browse_.current_node()) {
      const ListRow* row = browse_.values().RowAt(activate->iItem);
      bool fast_activate = false;
      if (header->code == LVN_ITEMACTIVATE) {
        if (!value_activate_from_key_) {
          return 0;
        }
        value_activate_from_key_ = false;
        fast_activate = true;
      }
      if (header->code == NM_DBLCLK) {
        fast_activate = true;
      }
      if (row && row->kind == rowkind::kKey) {
        if (fast_activate) {
          std::wstring path = registry_path::Build(*browse_.current_node());
          if (!row->extra.empty()) {
            path.append(L"\\");
            path.append(row->extra);
          }
          SelectTreePath(path);
        }
        return 0;
      }
      if (row && row->kind == rowkind::kValue) {
        if (activate->iSubItem == kValueColComment) {
          HandleMenuCommand(cmd::kEditModifyComment);
        } else {
          HandleMenuCommand(cmd::kEditModify);
        }
        return 0;
      }
    }
    return 0;
  }

  if (header->code == NM_CUSTOMDRAW) {
    auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);
    if (!draw) {
      return CDRF_DODEFAULT;
    }
    switch (draw->nmcd.dwDrawStage) {
    case CDDS_PREPAINT:
      return show_value_grid_ ? (CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT)
                              : CDRF_NOTIFYITEMDRAW;
    case CDDS_ITEMPREPAINT:
      draw->nmcd.uItemState &= ~CDIS_FOCUS;
      return show_value_grid_ ? CDRF_NOTIFYPOSTPAINT : CDRF_DODEFAULT;
    case CDDS_ITEMPOSTPAINT: {


      RECT row = {};
      if (ListView_GetItemRect(browse_.values().hwnd(),
                               static_cast<int>(draw->nmcd.dwItemSpec), &row,
                               LVIR_BOUNDS)) {
        PaintValueGridLines(browse_.values().hwnd(), draw->nmcd.hdc, row,
                            row.bottom - 1, row.bottom - row.top);
      }
      return CDRF_DODEFAULT;
    }
    case CDDS_POSTPAINT:
      PaintValueGridTail(browse_.values().hwnd(), draw->nmcd.hdc);
      return CDRF_DODEFAULT;
    default:
      return CDRF_DODEFAULT;
    }
  }
  return 0;
}



void MainWindow::Impl::PaintValueGridLines(HWND list, HDC hdc, const RECT& area,
                                           int first_line_y, int row_height) {
  HWND header = list ? ListView_GetHeader(list) : nullptr;
  if (!hdc || !header || area.top >= area.bottom) {
    return;
  }
  RECT header_rect = {};
  if (!GetWindowRect(header, &header_rect)) {
    return;
  }
  MapWindowPoints(nullptr, list, reinterpret_cast<POINT*>(&header_rect), 2);
  RECT client = {};
  GetClientRect(list, &client);
  if (grid_line_color_ == CLR_INVALID) {
    grid_line_color_ = HeaderDividerColor(header);
  }
  HBRUSH brush = appearance::CachedBrush(grid_line_color_);

  const int count = Header_GetItemCount(header);
  for (int i = 0; i < count; ++i) {
    RECT item = {};
    if (!Header_GetItemRect(header, i, &item)) {
      continue;
    }
    const int x = header_rect.left + item.right - 1;
    if (x < client.left || x >= client.right) {
      continue;
    }
    RECT line = {x, area.top, x + 1, area.bottom};
    FillRect(hdc, &line, brush);
  }

  if (row_height <= 0) {
    return;
  }
  for (int y = first_line_y; y < area.bottom; y += row_height) {
    if (y < area.top) {
      continue;
    }
    RECT line = {client.left, y, client.right, y + 1};
    FillRect(hdc, &line, brush);
  }
}


void MainWindow::Impl::PaintValueGridTail(HWND list, HDC hdc) {
  HWND header = list ? ListView_GetHeader(list) : nullptr;
  if (!hdc || !header) {
    return;
  }
  RECT client = {};
  GetClientRect(list, &client);
  RECT header_rect = {};
  if (!GetWindowRect(header, &header_rect)) {
    return;
  }
  MapWindowPoints(nullptr, list, reinterpret_cast<POINT*>(&header_rect), 2);

  RECT area = client;
  area.top = header_rect.bottom;
  int row_height = 0;
  const int count = ListView_GetItemCount(list);
  if (count > 0) {
    RECT last = {};
    if (ListView_GetItemRect(list, count - 1, &last, LVIR_BOUNDS)) {
      row_height = last.bottom - last.top;
      if (last.bottom > area.top) {
        area.top = last.bottom;
      }
    }
  }
  PaintValueGridLines(list, hdc, area, area.top + row_height - 1, row_height);
}

search::Result* MainWindow::Impl::SearchResultAt(int item) {
  const int tab_index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
  if (item < 0 || tab_index < 0 || static_cast<size_t>(tab_index) >= search_tabs_.size()) {
    return nullptr;
  }
  SearchTab& tab = search_tabs_[static_cast<size_t>(tab_index)];
  if (tab.is_compare) {
    return nullptr;
  }
  return static_cast<size_t>(item) < tab.results.size()
             ? &tab.results[static_cast<size_t>(item)]
             : nullptr;
}


std::wstring MainWindow::Impl::SearchRowKeyPath(int tab_index, int item) const {
  if (item < 0 || tab_index < 0 ||
      static_cast<size_t>(tab_index) >= search_tabs_.size()) {
    return std::wstring();
  }
  const SearchTab& tab = search_tabs_[static_cast<size_t>(tab_index)];
  const size_t row = static_cast<size_t>(item);
  if (tab.is_compare) {
    return row < tab.compare_rows.size() ? tab.compare_rows[row].key_path
                                         : std::wstring();
  }
  return row < tab.results.size() ? tab.results[row].key_path : std::wstring();
}

size_t MainWindow::Impl::SearchRowCount(int tab_index) const {
  if (tab_index < 0 || static_cast<size_t>(tab_index) >= search_tabs_.size()) {
    return 0;
  }
  const SearchTab& tab = search_tabs_[static_cast<size_t>(tab_index)];
  return tab.is_compare ? tab.compare_rows.size() : tab.results.size();
}

LRESULT MainWindow::Impl::HandleSearchListCustomDraw(NMLVCUSTOMDRAW* draw) {
  if (!draw || !search_results_list_) {
    return CDRF_DODEFAULT;
  }
  const int item = static_cast<int>(draw->nmcd.dwItemSpec);
  switch (draw->nmcd.dwDrawStage) {
  case CDDS_PREPAINT:
    return show_value_grid_ ? (CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT)
                            : CDRF_NOTIFYITEMDRAW;
  case CDDS_ITEMPREPAINT: {
    draw->nmcd.uItemState &= ~CDIS_FOCUS;
    LRESULT stage = show_value_grid_ ? CDRF_NOTIFYPOSTPAINT : CDRF_DODEFAULT;
    const search::Result* result = SearchResultAt(item);
    if (result && result->match_length > 0) {
      stage |= CDRF_NOTIFYSUBITEMDRAW;
    }
    return stage;
  }
  case CDDS_ITEMPOSTPAINT: {
    RECT row = {};
    if (show_value_grid_ &&
        ListView_GetItemRect(search_results_list_, item, &row, LVIR_BOUNDS)) {
      PaintValueGridLines(search_results_list_, draw->nmcd.hdc, row,
                          row.bottom - 1, row.bottom - row.top);
    }
    return CDRF_DODEFAULT;
  }
  case CDDS_POSTPAINT:
    if (show_value_grid_) {
      PaintValueGridTail(search_results_list_, draw->nmcd.hdc);
    }
    return CDRF_DODEFAULT;
  case CDDS_ITEMPREPAINT | CDDS_SUBITEM: {
    search::Result* result = SearchResultAt(item);
    if (!result ||
        SearchMatchSubItem(*result) != MappedSubItem(search_column_subitems_, draw->iSubItem)) {
      return CDRF_DODEFAULT;
    }
    return CDRF_NOTIFYPOSTPAINT;
  }
  case CDDS_ITEMPOSTPAINT | CDDS_SUBITEM: {
    const search::Result* result = SearchResultAt(item);
    const int subitem = MappedSubItem(search_column_subitems_, draw->iSubItem);
    if (!result || SearchMatchSubItem(*result) != subitem) {
      return CDRF_DODEFAULT;
    }
    RECT cell = {};
    if (draw->iSubItem == 0) {
      RECT row = {};
      if (!ListView_GetItemRect(search_results_list_, item, &cell, LVIR_LABEL) ||
          !ListView_GetItemRect(search_results_list_, item, &row, LVIR_BOUNDS)) {
        return CDRF_DODEFAULT;
      }
      cell.right = row.left + ListView_GetColumnWidth(search_results_list_, 0);
    } else {
      cell.top = draw->iSubItem;
      cell.left = LVIR_BOUNDS;
      if (!SendMessageW(search_results_list_, LVM_GETSUBITEMRECT, item,
                        reinterpret_cast<LPARAM>(&cell))) {
        return CDRF_DODEFAULT;
      }
      cell.left += kCellTextPadding;
    }
    cell.right -= kCellTextPadding;
    if (cell.left >= cell.right) {
      return CDRF_DODEFAULT;
    }
    std::wstring_view cell_text;
    switch (subitem) {
    case 0:
      cell_text = result->key_path;
      break;
    case 1:
      cell_text = search::DisplayName(*result);
      break;
    case 3:
      cell_text = result->data_text;
      break;
    default:
      return CDRF_DODEFAULT;
    }
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(search_results_list_, WM_GETFONT, 0, 0));
    HFONT old_font = font ? reinterpret_cast<HFONT>(SelectObject(draw->nmcd.hdc, font)) : nullptr;
    DrawSearchMatchOverlay(draw->nmcd.hdc, cell, cell_text,
                           static_cast<int>(result->match_start),
                           static_cast<int>(result->match_length));
    if (old_font) {
      SelectObject(draw->nmcd.hdc, old_font);
    }
    return CDRF_DODEFAULT;
  }
  default:
    return CDRF_DODEFAULT;
  }
}

LRESULT MainWindow::Impl::HandleHistoryNotification(NMHDR* header, LPARAM lparam) {
  if (header->hwndFrom == history_list_ && header->code == LVN_GETDISPINFOW) {
    auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
    const auto& entries = change_history_.entries();
    if (!disp || disp->item.iItem < 0 ||
        static_cast<size_t>(disp->item.iItem) >= entries.size()) {
      if (disp && (disp->item.mask & LVIF_TEXT) && disp->item.pszText &&
          disp->item.cchTextMax > 0) {
        disp->item.pszText[0] = L'\0';
      }
      return 0;
    }
    if (disp->item.mask & LVIF_TEXT) {
      const auto& entry = entries[static_cast<size_t>(disp->item.iItem)];
      const std::wstring* text = &entry.time_text;
      switch (disp->item.iSubItem) {
      case 1: text = &entry.action; break;
      case 2: text = &entry.old_data; break;
      case 3: text = &entry.new_data; break;
      default: break;
      }
      if (text->size() > kCellTextDrawLimit && disp->item.pszText &&
          disp->item.cchTextMax > 0) {
        lstrcpynW(disp->item.pszText, text->c_str(), disp->item.cchTextMax);
      } else {
        disp->item.pszText = const_cast<wchar_t*>(text->c_str());
      }
    }
    return 0;
  }
  if (header->hwndFrom == history_list_ && header->code == LVN_COLUMNCLICK) {
    auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
    if (info) {
      SortHistoryList(info->iSubItem, true);
    }
    return 0;
  }
  if (header->code == NM_CUSTOMDRAW) {
    auto* draw = reinterpret_cast<NMLVCUSTOMDRAW*>(lparam);
    if (!draw) {
      return CDRF_DODEFAULT;
    }
    if (draw->nmcd.dwDrawStage == CDDS_PREPAINT) {
      return CDRF_NOTIFYITEMDRAW;
    }
    if (draw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
      draw->nmcd.uItemState &= ~CDIS_FOCUS;
    }
    return CDRF_DODEFAULT;
  }
  return 0;
}

LRESULT MainWindow::Impl::HandleSearchNotification(NMHDR* header, LPARAM lparam) {
  if (header->hwndFrom == search_results_list_ &&
      header->code == LVN_ODCACHEHINT) {
    auto* hint = reinterpret_cast<NMLVCACHEHINT*>(lparam);
    if (hint) {
      QueueSearchPreviews(hint->iFrom, hint->iTo);
    }
    return 0;
  }
  if (header->hwndFrom == search_results_list_ && header->code == LVN_GETDISPINFOW) {
    auto* disp = reinterpret_cast<NMLVDISPINFOW*>(lparam);
    const int index = SearchIndexFromTab(TabCtrl_GetCurSel(tab_));
    SearchTab* tab = index >= 0 && static_cast<size_t>(index) < search_tabs_.size()
                         ? &search_tabs_[static_cast<size_t>(index)]
                         : nullptr;
    auto clear_cell = [&]() {
      if ((disp->item.mask & LVIF_TEXT) && disp->item.pszText &&
          disp->item.cchTextMax > 0) {
        disp->item.pszText[0] = L'\0';
      }
      if (disp->item.mask & LVIF_IMAGE) {
        disp->item.iImage = 0;
      }
    };
    if (!tab || disp->item.iItem < 0) {
      clear_cell();
      return 0;
    }
    const size_t row_index = static_cast<size_t>(disp->item.iItem);
    const int subitem = MappedSubItem(search_column_subitems_, disp->item.iSubItem);

    if (tab->is_compare) {
      if (row_index >= tab->compare_rows.size()) {
        clear_cell();
        return 0;
      }
      const search::compare::Row& row = tab->compare_rows[row_index];
      if (disp->item.mask & LVIF_TEXT) {
        if (disp->item.pszText && disp->item.cchTextMax > 0) {
          disp->item.pszText[0] = L'\0';
        }
        auto set_text = [&](const std::wstring& text) {
          if (text.size() > kCellTextDrawLimit && disp->item.pszText &&
              disp->item.cchTextMax > 0) {
            lstrcpynW(disp->item.pszText, text.c_str(), disp->item.cchTextMax);
          } else {
            disp->item.pszText = const_cast<wchar_t*>(text.c_str());
          }
        };
        switch (subitem) {
        case 0:
          set_text(row.key_path);
          break;
        case 1:
          if (row.is_key) {
            disp->item.pszText = const_cast<wchar_t*>(L"(Key)");
          } else if (row.value_name.empty()) {
            disp->item.pszText = const_cast<wchar_t*>(L"(Default)");
          } else {
            set_text(row.value_name);
          }
          break;
        case 2:
          set_text(row.first_text);
          break;
        case 3:
          set_text(row.second_text);
          break;
        default:
          break;
        }
      }
      if (disp->item.mask & LVIF_IMAGE) {
        disp->item.iImage = row.is_key ? kFolderIconIndex : kValueIconIndex;
      }
      return 0;
    }

    if (row_index >= tab->results.size()) {
      clear_cell();
      return 0;
    }
    search::Result& result = tab->results[row_index];

    if (subitem == 3 && result.data_state == search::DataState::kNotLoaded &&
        !search_preview_request_posted_) {
      search_preview_request_posted_ = true;
      if (!PostMessageW(hwnd_, frame::message_id::kSearchPreviewRequest, 0, 0)) {
        search_preview_request_posted_ = false;
      }
    }
    if (disp->item.mask & LVIF_TEXT) {
      wchar_t* buffer = disp->item.pszText;
      const int capacity = disp->item.cchTextMax;
      if (buffer && capacity > 0) {
        buffer[0] = L'\0';
      }
      auto set_text = [&](const std::wstring& text) {
        if (text.size() > kCellTextDrawLimit && buffer && capacity > 0) {
          lstrcpynW(buffer, text.c_str(), capacity);
        } else {
          disp->item.pszText = const_cast<wchar_t*>(text.c_str());
        }
      };
      switch (subitem) {
      case 0:
        set_text(result.key_path);
        break;
      case 1:
        if (search::IsKeyRow(result)) {
          disp->item.pszText = const_cast<wchar_t*>(L"");
        } else if (result.value_name.empty()) {
          disp->item.pszText = const_cast<wchar_t*>(L"(Default)");
        } else {
          disp->item.pszText = const_cast<wchar_t*>(result.value_name.c_str());
        }
        break;
      case 2:
        if (search::IsKeyRow(result)) {
          disp->item.pszText = const_cast<wchar_t*>(L"Key");
        } else if (result.kind == search::ResultKind::kTraceValue) {
          disp->item.pszText = const_cast<wchar_t*>(L"TRACE");
        } else if (buffer && capacity > 0) {
          lstrcpynW(buffer, value_format::TypeName(result.type).c_str(), capacity);
        }
        break;
      case 3:
        set_text(result.data_text);
        break;
      case 4:
        if (buffer && capacity > 0 && !search::IsKeyRow(result) &&
            result.kind != search::ResultKind::kTraceValue) {
          swprintf_s(buffer, static_cast<size_t>(capacity), L"%lu",
                     static_cast<unsigned long>(result.data_size));
        }
        break;
      case 5:
        if (buffer && capacity > 0) {
          FormatCellFileTime(result.modified, buffer, capacity);
        }
        break;
      default:
        break;
      }
    }
    if (disp->item.mask & LVIF_IMAGE) {
      if (search::IsKeyRow(result)) {
        disp->item.iImage = kFolderIconIndex;
      } else if (UseBinaryValueIcon(result.type)) {
        disp->item.iImage = kBinaryIconIndex;
      } else {
        disp->item.iImage = kValueIconIndex;
      }
    }
    return 0;
  }
  if (header->hwndFrom == search_results_list_ && header->code == LVN_COLUMNCLICK) {
    auto* info = reinterpret_cast<NMLISTVIEW*>(lparam);
    if (info) {
      SortSearchResults(info->iSubItem, true);
    }
    return 0;
  }
  if (header->hwndFrom == search_results_list_ && (header->code == NM_DBLCLK || header->code == LVN_ITEMACTIVATE)) {
    if (IsCompareTabSelected()) {
      return 0;
    }
    auto* activate = reinterpret_cast<NMITEMACTIVATE*>(lparam);
    if (activate && activate->iItem >= 0) {
      int sel = TabCtrl_GetCurSel(tab_);
      int index = SearchIndexFromTab(sel);
      if (index >= 0 &&
          static_cast<size_t>(activate->iItem) < SearchRowCount(index)) {
        const std::wstring activated_path =
            SearchRowKeyPath(index, activate->iItem);
        if (SearchResultOpensInNewTab()) {
          OpenLocalRegistryTab();
        } else {
          ActivateRegistryTab();
        }
        ApplyViewVisibility();
        UpdateStatus();
        SelectTreePath(activated_path);
        const search::Result* row = SearchResultAt(activate->iItem);
        if (row && !search::IsKeyRow(*row)) {
          SelectValueByName(row->value_name);
        }
      }
    }
    return 0;
  }
  if (header->code == NM_CUSTOMDRAW) {
    return HandleSearchListCustomDraw(reinterpret_cast<NMLVCUSTOMDRAW*>(lparam));
  }
  return 0;
}

bool MainWindow::Impl::OnCreate() {
  ChangeWindowMessageFilterEx(hwnd_, WM_COPYDATA, MSGFLT_ALLOW, nullptr);
  ui_font_ = CreateUIFont();
  icon_font_ = CreateIconFont(10);
  custom_font_ = DefaultLogFont();
  LoadSettings();
  if (theme_mode_ == ThemeMode::kCustom) {
    LoadThemePresets();
  }
  ApplySavedWindowPlacement();
  if (theme_mode_ != ThemeMode::kCustom ||
      !ApplyThemePresetByName(active_theme_preset_, false)) {
    Theme::SetMode(theme_mode_);
    ApplySystemTheme();
  }
  UpdateUIFont();
  BuildMenus();
  BuildAccelerators();

  toolbar_.Create(hwnd_, instance_, kToolbarId);

  std::vector<TBBUTTON> buttons;
  buttons.push_back({0, cmd::kRegistryLocal, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({1, cmd::kRegistryNetwork, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({2, cmd::kRegistryOffline, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({6, kToolbarSepGroup1, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0});
  buttons.push_back({3, cmd::kEditFind, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({4, cmd::kEditReplace, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({5, cmd::kFileExport, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({6, kToolbarSepGroup2, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0});
  buttons.push_back({6, cmd::kEditUndo, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({7, cmd::kEditRedo, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({8, cmd::kEditCopy, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({9, cmd::kEditPaste, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({10, cmd::kEditDelete, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({11, cmd::kViewRefresh, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({6, kToolbarSepGroup3, TBSTATE_ENABLED, BTNS_SEP, {0}, 0, 0});
  buttons.push_back({12, cmd::kNavBack, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({13, cmd::kNavForward, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  buttons.push_back({14, cmd::kNavUp, TBSTATE_ENABLED, BTNS_BUTTON, {0}, 0, 0});
  toolbar_.AddButtons(buttons);

  browse::CreateRequest browse_request;
  browse_request.parent = hwnd_;
  browse_request.instance = instance_;
  browse_request.address_id = kAddressEditId;
  browse_request.go_id = kAddressGoId;
  browse_request.filter_id = kFilterEditId;
  browse_request.tree_id = kTreeId;
  browse_request.values_id = kValueListId;
  browse_request.address_proc = AddressEditProc;
  browse_request.address_subclass_id = kAddressSubclassId;
  browse_request.filter_proc = FilterEditProc;
  browse_request.filter_subclass_id = kFilterSubclassId;
  browse_request.tree_proc = TreeViewProc;
  browse_request.tree_subclass_id = kTreeViewSubclassId;
  browse_request.values_proc = ListViewProc;
  browse_request.values_subclass_id = kListViewSubclassId;
  browse_request.callback_context = reinterpret_cast<DWORD_PTR>(this);
  if (!browse_.Create(browse_request)) {
    return false;
  }

  value_tooltip_ = CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr,
                                   WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
                                   CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                                   CW_USEDEFAULT, hwnd_, nullptr, instance_, nullptr);
  if (value_tooltip_) {
    TOOLINFOW info = {};
    info.cbSize = sizeof(info);
    info.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
    info.hwnd = hwnd_;
    info.uId = reinterpret_cast<UINT_PTR>(browse_.values().hwnd());
    info.lpszText = LPSTR_TEXTCALLBACKW;
    SendMessageW(value_tooltip_, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&info));
    SendMessageW(value_tooltip_, TTM_SETMAXTIPWIDTH, 0, kValueTooltipMaxWidth);
    AllowDarkModeForWindow(value_tooltip_, Theme::UseDarkMode());
    SetWindowTheme(value_tooltip_,
                   Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
  }

  tab_ = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TCS_TABS | TCS_FOCUSNEVER, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTabId)), instance_, nullptr);
  ApplyFont(tab_, ui_font_);
  TabCtrl_SetPadding(tab_, kTabTextPaddingX, kTabInsetY);
  SetWindowSubclass(tab_, TabProc, kTabSubclassId, reinterpret_cast<DWORD_PTR>(this));

  tree_header_ = CreateWindowExW(0, L"STATIC", L"Key Tree", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | SS_LEFT | SS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTreeHeaderId)), instance_, nullptr);
  tree_close_btn_ = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kTreeHeaderCloseId)), instance_, nullptr);
  SetWindowPos(tree_close_btn_, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

  browse_.tree().SetIconResolver([this](const RegistryNode& node) { return KeyIconIndex(node, nullptr, nullptr); });
  browse_.tree().SetVirtualChildProvider([this](const RegistryNode& node, const std::unordered_set<std::wstring>& existing_lower, std::vector<std::wstring>* out) { AppendTraceChildren(node, existing_lower, out); });
  search_results_list_ = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | WS_CLIPSIBLINGS | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kSearchResultsListId)), instance_, nullptr);
  LoadTabs();

  history_label_ = CreateWindowExW(0, L"STATIC", L"History", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | SS_LEFT | SS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryLabelId)), instance_, nullptr);
  history_close_btn_ = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | BS_OWNERDRAW, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryHeaderCloseId)), instance_, nullptr);
  SetWindowPos(history_close_btn_, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
  status_bar_ = CreateWindowExW(0, STATUSCLASSNAMEW, L"", WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kStatusBarId)), instance_, nullptr);
  if (status_bar_) {
    int parts[4] = {0, 0, 0, 0};
    SendMessageW(status_bar_, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));
  }
  search_progress_ = CreateWindowExW(0, PROGRESS_CLASSW, L"", WS_CHILD | PBS_MARQUEE, 0, 0, 0, 0, status_bar_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kSearchProgressId)), instance_, nullptr);
  if (search_progress_) {
    SendMessageW(search_progress_, PBM_SETMARQUEE, TRUE, 30);
    SendMessageW(search_progress_, PBM_SETRANGE32, 0, 1);
    ShowWindow(search_progress_, SW_HIDE);
  }
  history_list_ = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA, 0, 0, 0, 0, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kHistoryListId)), instance_, nullptr);

  DWORD ex_mask = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_BORDERSELECT | LVS_EX_TRACKSELECT | LVS_EX_ONECLICKACTIVATE | LVS_EX_TWOCLICKACTIVATE | LVS_EX_UNDERLINEHOT;
  DWORD ex_style = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER;
  ListView_SetExtendedListViewStyleEx(history_list_, ex_mask, ex_style);
  ListView_SetExtendedListViewStyleEx(search_results_list_, ex_mask, ex_style);
  if (value_tooltip_) {
    TOOLINFOW tip = {};
    tip.cbSize = sizeof(tip);
    tip.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
    tip.hwnd = hwnd_;
    tip.lpszText = LPSTR_TEXTCALLBACKW;
    tip.uId = reinterpret_cast<UINT_PTR>(search_results_list_);
    SendMessageW(value_tooltip_, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&tip));
    tip.uId = reinterpret_cast<UINT_PTR>(history_list_);
    SendMessageW(value_tooltip_, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&tip));
  }
  SendMessageW(search_results_list_, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
  SendMessageW(history_list_, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
  SetWindowSubclass(history_list_, ListViewProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this));
  SetWindowSubclass(search_results_list_, ListViewProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this));

  AttachBorder(browse_.values().hwnd());
  AttachBorder(tree_header_);
  AttachBorder(browse_.tree().hwnd());
  AttachBorder(history_label_);
  AttachBorder(history_list_);
  AttachBorder(search_results_list_);
  AttachBorder(browse_.address());
  AttachBorder(browse_.go_button());
  UpdateGoButtonState();
  AttachBorder(browse_.filter());

  ApplyUIFontToControls();

  CreateValueColumns();
  CreateHistoryColumns();
  CreateSearchColumns();
  ApplyThemeToChildren();
  SetValueGridEnabled(show_value_grid_, false);
  if (toolbar_.hwnd()) {
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditUndo, 0);
    SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditRedo, 0);
    if (read_only_) {
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditPaste, 0);
      SendMessageW(toolbar_.hwnd(), TB_SETSTATE, cmd::kEditDelete, 0);
    }
  }

  browse_.roots() = RegistryStore::DefaultRoots(show_extra_hives_);
  AppendRealRegistryRoot(&browse_.roots());
  browse_.tree().SetRegeditLayout(false);
  browse_.tree().SetRootLabel(TreeRootLabel());
  browse_.tree().PopulateRoots(browse_.roots());

  int initial_tab = tab_ ? TabCtrl_GetCurSel(tab_) : -1;
  if (initial_tab >= 0 && !IsSearchTabIndex(initial_tab) && !IsRegFileTabIndex(initial_tab)) {
    RestoreRegistryTabState(initial_tab);
  } else {
    SelectDefaultTreeItem();
  }
  StartValueListWorker();

  ApplyViewVisibility();
  ApplyAlwaysOnTop();
  UpdateStatus();
  return true;
}

void MainWindow::Impl::RunDeferredStartup() {
  if (deferred_startup_complete_) {
    return;
  }
  deferred_startup_complete_ = true;
  const bool has_external_jump = !queued_external_jump_target_.empty();
  bool use_global_tree_state = (!has_external_jump && save_tree_state_);
  if (use_global_tree_state && tab_) {
    int active_tab = TabCtrl_GetCurSel(tab_);
    if (active_tab >= 0 && static_cast<size_t>(active_tab) < tabs_.size()) {
      const TabEntry& entry = tabs_[static_cast<size_t>(active_tab)];
      if (entry.kind == TabEntry::Kind::kRegistry && (!entry.selected_path.empty() || !entry.expanded_paths.empty())) {
        use_global_tree_state = false;
      }
    }
  }
  startup_tree_restore_pending_ = use_global_tree_state;

  EnableAddressAutoComplete();
  ReloadThemeIcons();

  UpdateSearchResultsView();
  if (has_external_jump) {
    tree_state_restored_ = true;
  } else if (!use_global_tree_state) {
    tree_state_restored_ = true;
  }
  StartStartupCacheLoad(use_global_tree_state);
  ApplyQueuedExternalJump();
  StartTreeStateWorker();
  MarkTreeStateDirty();

  BuildMenus();
  UpdateStatus();
  if (auto_check_updates_) {
    CheckForUpdates(true);
  }
}

void MainWindow::Impl::StartStartupCacheLoad(bool include_tree_state) {
  StopStartupCacheLoad();
  const bool load_tree_state = include_tree_state && save_tree_state_;
  int history_max_rows = history_max_rows_;
  int history_sort_column = history_sort_column_;
  bool history_sort_ascending = history_sort_ascending_;
  const HWND hwnd = hwnd_;
  startup_cache_session_.Start(
      [this, load_tree_state, history_max_rows, history_sort_column,
       history_sort_ascending, hwnd](
          uint64_t generation, const std::atomic_bool& cancel) {
        auto payload = std::make_unique<StartupCachePayload>();
        payload->generation = generation;

        std::wstring comments_path = CommentsPath();
        std::wstring comments_content;
        if (!comments_path.empty() &&
            util::ReadTextFile(
                comments_path, &comments_content, nullptr,
                static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
          changes::CommentDocument comments =
              changes::ParseComments(comments_content);
          payload->value_comments = std::move(comments.value_entries);
          payload->name_comments = std::move(comments.name_entries);
          payload->comments_loaded = true;
        }
        if (cancel.load()) {
          return;
        }

        std::wstring history_path = HistoryCachePath();
        std::wstring history_content;
        if (!history_path.empty() &&
            util::ReadTextFile(
                history_path, &history_content, nullptr,
                static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
          changes::ChangeHistory history;
          history.Replace(
              std::move(changes::ParseHistory(history_content).entries),
              static_cast<size_t>(history_max_rows));
          history.Sort(history_sort_column, history_sort_ascending);
          payload->history_entries = std::move(history.entries());
          payload->history_loaded = true;
        } else {
          payload->history_loaded = true;
        }
        if (cancel.load()) {
          return;
        }

        if (load_tree_state) {
          std::wstring tree_path = TreeStatePath();
          std::wstring tree_content;
          if (!tree_path.empty() &&
              util::ReadTextFile(
                  tree_path, &tree_content, nullptr,
                  static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
            workspace::TreeState state =
                workspace::ParseTreeState(tree_content);
            payload->tree_selected_path = std::move(state.selected_path);
            payload->tree_expanded_paths = std::move(state.expanded_paths);
          }
          payload->tree_state_loaded = true;
        }

        if (cancel.load()) {
          return;
        }
        if (hwnd && IsWindow(hwnd) &&
            PostMessageW(hwnd, frame::message_id::kStartupCacheReady,
                         static_cast<WPARAM>(generation),
                         reinterpret_cast<LPARAM>(payload.get()))) {
          ReleasePostedPayload(payload);
        }
      });
}

void MainWindow::Impl::StopStartupCacheLoad() {
  startup_cache_session_.CancelAndJoin();
}

void MainWindow::Impl::ApplyStartupCachePayload(StartupCachePayload* payload) {
  if (!payload) {
    return;
  }
  std::unique_ptr<StartupCachePayload> owned(payload);
  if (!startup_cache_session_.IsCurrent(owned->generation)) {
    return;
  }
  startup_cache_session_.Join();

  if (owned->comments_loaded) {
    std::unordered_map<std::wstring, changes::CommentEntry>
        merged_value_comments;
    std::unordered_map<std::wstring, changes::CommentEntry>
        merged_name_comments;
    for (auto& entry : owned->value_comments) {
      merged_value_comments[changes::ValueComments::ValueKey(entry.path, entry.name, entry.type)] = std::move(entry);
    }
    for (auto& entry : owned->name_comments) {
      merged_name_comments[changes::ValueComments::NameKey(entry.name, entry.type)] = std::move(entry);
    }
    for (auto& pair : value_comments_.value_entries()) {
      merged_value_comments[pair.first] = std::move(pair.second);
    }
    for (auto& pair : value_comments_.name_entries()) {
      merged_name_comments[pair.first] = std::move(pair.second);
    }
    value_comments_.value_entries() = std::move(merged_value_comments);
    value_comments_.name_entries() = std::move(merged_name_comments);
    RefreshValueListComments();
    UpdateSearchResultsView();
  }

  if (owned->history_loaded) {
    std::vector<HistoryEntry> pending_session_entries;
    if (!history_loaded_ && !change_history_.entries().empty()) {
      pending_session_entries = change_history_.entries();
    }
    if (!change_history_.entries().empty()) {
      owned->history_entries.insert(owned->history_entries.end(), change_history_.entries().begin(), change_history_.entries().end());
    }
    change_history_.Replace(std::move(owned->history_entries),
                            static_cast<size_t>(history_max_rows_));
    change_history_.Sort(history_sort_column_, history_sort_ascending_);
    history_loaded_ = true;
    for (const auto& entry : pending_session_entries) {
      AppendHistoryCache(entry);
    }
    RebuildHistoryList();
  }

  if (owned->tree_state_loaded) {
    saved_tree_state_.selected_path = std::move(owned->tree_selected_path);
    saved_tree_state_.expanded_paths =
        std::move(owned->tree_expanded_paths);
    if (startup_tree_restore_pending_ && !tree_state_restored_) {
      applying_startup_tree_restore_ = true;
      RestoreTreeState();
      applying_startup_tree_restore_ = false;
    } else if (!tree_state_restored_) {
      tree_state_restored_ = true;
    }
    startup_tree_restore_pending_ = false;
  }
}

void MainWindow::Impl::OnDestroy() {
  if (hwnd_) {
    RemovePropW(hwnd_, kRegKitWindowProperty);
  }
  EndJumpUiBatch();
  StopStartupCacheLoad();
  StopReplace();
  StopTraceParseSessions();
  StopDefaultParseSessions();
  StopRegFileParseSessions();
  StopTraceLoadWorker();
  StopDefaultLoadWorker();
  StopValueListWorker();
  StopTreeStateWorker();
  CancelSearch();
  DiscardWorkerMessages();
  for (auto& entry : tabs_) {
    if (entry.kind == TabEntry::Kind::kRegFile) {
      ReleaseRegFileRoots(&entry);
    }
  }
  if (clear_tabs_on_exit_) {
    ClearTabsCache();
  } else if (save_tabs_ && !SaveTabs()) {
    ui::ShowError(hwnd_, L"The open tabs couldn't be saved for the next session.");
  }
  ClearHistoryItems(false);
  if (clear_history_on_exit_) {
    std::wstring history_path = HistoryCachePath();
    if (!history_path.empty()) {
      DeleteFileW(history_path.c_str());
    }
  }
  UnloadOfflineRegistry(nullptr);
  ReleaseRemoteRegistry();
  if (ui_font_ && ui_font_owned_) {
    DeleteObject(ui_font_);
  }
  ui_font_ = nullptr;
  ui_font_owned_ = false;
  if (icon_font_) {
    DeleteObject(icon_font_);
    icon_font_ = nullptr;
  }
  if (tree_images_) {
    ImageList_Destroy(tree_images_);
    tree_images_ = nullptr;
  }
  if (list_images_) {
    ImageList_Destroy(list_images_);
    list_images_ = nullptr;
  }
  if (address_go_icon_) {
    DestroyIcon(address_go_icon_);
    address_go_icon_ = nullptr;
  }
  if (value_grid_image_list_) {
    ImageList_Destroy(value_grid_image_list_);
    value_grid_image_list_ = nullptr;
  }
  if (address_autocomplete_) {
    address_autocomplete_->Release();
    address_autocomplete_ = nullptr;
  }
  if (address_autocomplete_source_) {
    address_autocomplete_source_->Release();
    address_autocomplete_source_ = nullptr;
  }
  if (accelerators_) {
    DestroyAcceleratorTable(accelerators_);
    accelerators_ = nullptr;
  }
  menu_items_.clear();
}

void MainWindow::Impl::DiscardWorkerMessages() {
  if (!hwnd_) {
    return;
  }
  MSG message = {};
  const UINT payload_messages[] = {
      frame::message_id::kTraceLoadReady, frame::message_id::kDefaultLoadReady,
      frame::message_id::kStartupCacheReady, frame::message_id::kRegFileLoadReady,
      frame::message_id::kTraceParseBatch, frame::message_id::kDefaultParseBatch,
      frame::message_id::kValueListReady, frame::message_id::kReplaceReady,
      frame::message_id::kValuePreviewReady};
  for (const UINT id : payload_messages) {
    while (PeekMessageW(&message, hwnd_, id, id, PM_REMOVE)) {
      switch (id) {
      case frame::message_id::kTraceLoadReady:
        delete reinterpret_cast<TraceLoadPayload*>(message.lParam);
        break;
      case frame::message_id::kDefaultLoadReady:
        delete reinterpret_cast<DefaultLoadPayload*>(message.lParam);
        break;
      case frame::message_id::kStartupCacheReady:
        delete reinterpret_cast<StartupCachePayload*>(message.lParam);
        break;
      case frame::message_id::kRegFileLoadReady:
        delete reinterpret_cast<RegFileParsePayload*>(message.lParam);
        break;
      case frame::message_id::kTraceParseBatch:
        delete reinterpret_cast<TraceParseBatch*>(message.lParam);
        break;
      case frame::message_id::kDefaultParseBatch:
        delete reinterpret_cast<DefaultParseBatch*>(message.lParam);
        break;
      case frame::message_id::kValueListReady:
        delete reinterpret_cast<ValueListPayload*>(message.lParam);
        break;
      case frame::message_id::kValuePreviewReady:
        delete reinterpret_cast<ValuePreviewPayload*>(message.lParam);
        break;
      case frame::message_id::kReplaceReady:
        delete reinterpret_cast<ReplacePayload*>(message.lParam);
        break;
      default:
        break;
      }
    }
  }
}

} // namespace regkit
