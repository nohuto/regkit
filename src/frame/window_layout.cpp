// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

void MainWindow::Impl::ComputeSplitterLimits(int* min_width, int* max_width) const {
  if (!min_width || !max_width || !hwnd_) {
    return;
  }
  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  int width = rect.right - rect.left;
  int available_width = std::max(0, width);
  int max_tree = std::max(kMinTreeWidth, available_width - kMinValueListWidth - kSplitterWidth);
  *min_width = kMinTreeWidth;
  *max_width = max_tree;
}

void MainWindow::Impl::ComputeHistorySplitterLimits(int* min_height, int* max_height) const {
  if (!min_height || !max_height || !hwnd_) {
    return;
  }
  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  int height = rect.bottom - rect.top;

  UINT dpi = win32::DpiForWindow(hwnd_);
  const int address_height = CalcEditHeight(browse_.address(), ui_font_, util::ScaleForDpi(18, dpi));
  const int tabs_height = std::max(20, tab_height_);
  int status_height = 0;
  if (status_bar_ && show_status_bar_) {
    RECT sb_rect = {};
    GetWindowRect(status_bar_, &sb_rect);
    status_height = sb_rect.bottom - sb_rect.top;
    if (status_height <= 0) {
      status_height = 20;
    }
  }

  int y = kMainVerticalGap;
  if (show_toolbar_) {
    SendMessageW(toolbar_.hwnd(), TB_AUTOSIZE, 0, 0);
    RECT tb_rect = {};
    GetWindowRect(toolbar_.hwnd(), &tb_rect);
    y += tb_rect.bottom - tb_rect.top;
  }
  if (show_address_bar_) {
    y += address_height + kMainVerticalGap;
  }

  const bool show_search = IsSearchTabSelected();
  const bool show_value = show_value_ && !show_search;
  const bool show_tabs = show_tab_control_ && tab_;
  const bool show_filter = show_value && show_filter_bar_ && browse_.filter();
  if (show_tabs || show_filter) {
    y += tabs_height + kMainVerticalGap;
  }

  int status_top = height - status_height;
  int content_total_height = std::max(0, status_top - y);
  int max_history = std::max(kMinHistoryHeight, content_total_height - kHistoryMaxPadding);
  *min_height = kMinHistoryHeight;
  *max_height = max_history;
}

void MainWindow::Impl::InitDragLayout() {
  if (!hwnd_) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd_, &client);
  drag_client_width_ = client.right - client.left;
  drag_client_height_ = client.bottom - client.top;
  drag_content_left_ = 0;
  drag_content_right_ = drag_client_width_;

  drag_content_top_ = splitter_rect_.top;
  if (drag_content_top_ <= 0) {
    RECT rect = {};
    HWND target = browse_.values().hwnd() ? browse_.values().hwnd() : search_results_list_;
    if (target && GetWindowRect(target, &rect)) {
      MapWindowPoints(nullptr, hwnd_, reinterpret_cast<POINT*>(&rect), 2);
      drag_content_top_ = rect.top;
    }
  }
  if (drag_content_top_ <= 0) {
    drag_content_top_ = 0;
  }

  drag_status_top_ = drag_client_height_;
  if (show_status_bar_ && status_bar_) {
    RECT rect = {};
    if (GetWindowRect(status_bar_, &rect)) {
      MapWindowPoints(nullptr, hwnd_, reinterpret_cast<POINT*>(&rect), 2);
      drag_status_top_ = rect.top;
    }
  }

  const UINT dpi = win32::DpiForWindow(hwnd_);
  drag_tree_header_height_ = util::ScaleForDpi(kPanelHeaderHeight, dpi);
  drag_history_label_height_ = drag_tree_header_height_;
  drag_layout_valid_ = true;
}

void MainWindow::Impl::ApplyDragLayout() {
  if (!hwnd_) {
    return;
  }
  RECT client = {};
  GetClientRect(hwnd_, &client);
  int width = client.right - client.left;
  int height = client.bottom - client.top;
  if (!drag_layout_valid_ || width != drag_client_width_ || height != drag_client_height_) {
    InitDragLayout();
  }

  auto get_panel_rect = [&](HWND header, HWND body, RECT* rect) {
    if (!rect) {
      return false;
    }
    RECT header_rect = {};
    RECT body_rect = {};
    bool has_header = GetChildRectInParent(hwnd_, header, &header_rect);
    bool has_body = GetChildRectInParent(hwnd_, body, &body_rect);
    if (!has_header && !has_body) {
      return false;
    }
    if (!has_body) {
      *rect = header_rect;
      return true;
    }
    if (!has_header) {
      *rect = body_rect;
      return true;
    }
    UnionRect(rect, &header_rect, &body_rect);
    return true;
  };

  RECT old_tree_panel_rect = {};
  RECT old_history_panel_rect = {};
  RECT old_value_rect = {};
  const RECT old_splitter_rect = splitter_rect_;
  const RECT old_history_splitter_rect = history_splitter_rect_;
  const bool had_old_tree_panel = get_panel_rect(tree_header_, browse_.tree().hwnd(), &old_tree_panel_rect);
  const bool had_old_history_panel = get_panel_rect(history_label_, history_list_, &old_history_panel_rect);
  const bool had_old_value = GetChildRectInParent(hwnd_, browse_.values().hwnd(), &old_value_rect);

  const bool show_search = IsSearchTabSelected();
  const bool show_tree = show_tree_ && !show_search;
  const bool show_history = show_history_ && !show_search;
  const bool show_value = show_value_ && !show_search;

  int content_left = drag_content_left_;
  int content_right = drag_content_right_;
  int y = drag_content_top_;
  int status_top = drag_status_top_;
  int content_total_height = std::max(0, status_top - y);
  int min_history = kMinHistoryHeight;
  int max_history = std::max(min_history, content_total_height - kHistoryMaxPadding);
  int history_height = show_history ? ClampValue(history_height_, min_history, max_history) : 0;
  if (show_history) {
    history_height_ = history_height;
  }
  int history_top = status_top - history_height;

  int history_splitter_height = show_history ? kHistorySplitterHeight : 0;
  int history_gap = show_history ? kHistoryGap : 0;
  int splitter_bottom = show_history ? (history_top - history_gap) : history_top;
  int splitter_top = show_history ? (splitter_bottom - history_splitter_height) : history_top;
  if (show_history) {
    history_splitter_rect_.left = content_left;
    history_splitter_rect_.right = content_right;
    history_splitter_rect_.top = splitter_top;
    history_splitter_rect_.bottom = splitter_bottom;
  } else {
    history_splitter_rect_ = {};
  }
  int content_bottom = show_history ? splitter_top : status_top;
  int content_height = std::max(0, content_bottom - y);

  int available_width = content_right - content_left;
  int min_tree = kMinTreeWidth;
  int min_list = kMinValueListWidth;
  int max_tree = std::max(min_tree, available_width - min_list - kSplitterWidth);
  int tree_width = show_tree ? ClampValue(tree_width_, min_tree, max_tree) : 0;
  if (show_tree) {
    tree_width_ = tree_width;
  }

  int tree_header_height = drag_tree_header_height_;
  int history_label_height = drag_history_label_height_;
  const UINT dpi = win32::DpiForWindow(hwnd_);
  const int close_size = util::ScaleForDpi(kPanelCloseSize, dpi);
  const int close_inset = util::ScaleForDpi(kPanelCloseInset, dpi);
  int list_x = show_tree ? (content_left + tree_width + kSplitterWidth) : content_left;
  int list_width = content_right - list_x;
  int tree_content_height = std::max(0, content_height - (show_tree ? tree_header_height : 0));

  struct PanelPlacement {
    HWND target;
    int x;
    int y;
    int width;
    int height;
  };
  PanelPlacement placements[7] = {};
  int placement_count = 0;
  auto defer = [&](HWND target, int x, int y_pos, int w, int h) {
    if (target && placement_count < static_cast<int>(_countof(placements))) {
      placements[placement_count++] = {target, x, y_pos, w, h};
    }
  };

  if (show_history) {
    int history_width = content_right - content_left;
    defer(history_label_, content_left, history_top, history_width, history_label_height);
    defer(history_close_btn_, content_left + history_width - close_inset - close_size,
          history_top + (history_label_height - close_size) / 2, close_size, close_size);
    defer(history_list_, content_left,
          history_top + history_label_height - kPanelBorderOverlap,
          history_width,
          history_height - history_label_height + kPanelBorderOverlap);
  }

  if (show_tree) {
    defer(tree_header_, content_left, y, tree_width, tree_header_height);
    defer(tree_close_btn_, content_left + tree_width - close_inset - close_size,
          y + (tree_header_height - close_size) / 2, close_size, close_size);
    defer(browse_.tree().hwnd(), content_left,
          y + tree_header_height - kPanelBorderOverlap,
          tree_width, tree_content_height + kPanelBorderOverlap);
    splitter_rect_.left = content_left + tree_width;
    splitter_rect_.right = splitter_rect_.left + kSplitterWidth;
    splitter_rect_.top = y;
    splitter_rect_.bottom = y + content_height;
  } else {
    splitter_rect_ = {};
  }

  if (show_search) {
    defer(search_results_list_, content_left, y, content_right - content_left, content_height);
  } else if (show_value) {
    defer(browse_.values().hwnd(), list_x, y, list_width, content_height);
  }

  const UINT placement_flags = SWP_NOZORDER | SWP_NOACTIVATE;
  HDWP hdwp = BeginDeferWindowPos(placement_count);
  for (int i = 0; hdwp && i < placement_count; ++i) {
    const PanelPlacement& p = placements[i];
    hdwp = DeferWindowPos(hdwp, p.target, nullptr, p.x, p.y, p.width, p.height, placement_flags);
  }
  if (hdwp) {
    EndDeferWindowPos(hdwp);
  } else {
    for (int i = 0; i < placement_count; ++i) {
      const PanelPlacement& p = placements[i];
      SetWindowPos(p.target, nullptr, p.x, p.y, p.width, p.height, placement_flags);
    }
  }

  RECT new_tree_panel_rect = {};
  RECT new_history_panel_rect = {};
  RECT new_value_rect = {};
  bool has_new_tree_panel = get_panel_rect(tree_header_, browse_.tree().hwnd(), &new_tree_panel_rect);
  bool has_new_history_panel = get_panel_rect(history_label_, history_list_, &new_history_panel_rect);
  bool has_new_value = GetChildRectInParent(hwnd_, browse_.values().hwnd(), &new_value_rect);

  RECT dirty_layout = {};
  bool has_dirty_layout = false;
  auto extend_dirty = [&](const RECT& rect, bool has_rect) {
    if (!has_rect) {
      return;
    }
    if (!has_dirty_layout) {
      dirty_layout = rect;
      has_dirty_layout = true;
      return;
    }
    RECT combined = {};
    UnionRect(&combined, &dirty_layout, &rect);
    dirty_layout = combined;
  };
  extend_dirty(old_tree_panel_rect, had_old_tree_panel);
  extend_dirty(new_tree_panel_rect, has_new_tree_panel);
  extend_dirty(old_history_panel_rect, had_old_history_panel);
  extend_dirty(new_history_panel_rect, has_new_history_panel);
  extend_dirty(old_value_rect, had_old_value);
  extend_dirty(new_value_rect, has_new_value);
  extend_dirty(old_splitter_rect, old_splitter_rect.right > old_splitter_rect.left);
  extend_dirty(splitter_rect_, splitter_rect_.right > splitter_rect_.left);
  extend_dirty(old_history_splitter_rect, old_history_splitter_rect.bottom > old_history_splitter_rect.top);
  extend_dirty(history_splitter_rect_, history_splitter_rect_.bottom > history_splitter_rect_.top);

  LayoutValueGridToolbar();
  if (has_dirty_layout) {
    RedrawWindow(hwnd_, &dirty_layout, nullptr,
                 RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN | RDW_UPDATENOW);
  }
}

void MainWindow::Impl::BeginSplitterDrag() {
  splitter_dragging_ = true;
  ComputeSplitterLimits(&splitter_min_width_, &splitter_max_width_);
  drag_layout_valid_ = false;
  SetCapture(hwnd_);
}

void MainWindow::Impl::BeginHistorySplitterDrag() {
  history_splitter_dragging_ = true;
  ComputeHistorySplitterLimits(&history_splitter_min_height_, &history_splitter_max_height_);
  drag_layout_valid_ = false;
  SetCapture(hwnd_);
}

void MainWindow::Impl::UpdateSplitterTrack(int client_x) {
  if (!splitter_dragging_) {
    return;
  }
  int desired = splitter_start_width_ + (client_x - splitter_start_x_);
  desired = ClampValue(desired, splitter_min_width_, splitter_max_width_);
  if (desired == tree_width_) {
    return;
  }
  tree_width_ = desired;
  ApplyDragLayout();
}

void MainWindow::Impl::UpdateHistorySplitterTrack(int client_y) {
  if (!history_splitter_dragging_) {
    return;
  }
  int desired = history_splitter_start_height_ - (client_y - history_splitter_start_y_);
  desired = ClampValue(desired, history_splitter_min_height_, history_splitter_max_height_);
  if (desired == history_height_) {
    return;
  }
  history_height_ = desired;
  ApplyDragLayout();
}

void MainWindow::Impl::EndSplitterDrag() {
  if (!splitter_dragging_) {
    return;
  }
  splitter_dragging_ = false;
  if (GetCapture() == hwnd_) {
    ReleaseCapture();
  }
  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  LayoutControls(rect.right, rect.bottom);
}

void MainWindow::Impl::EndHistorySplitterDrag() {
  if (!history_splitter_dragging_) {
    return;
  }
  history_splitter_dragging_ = false;
  if (GetCapture() == hwnd_) {
    ReleaseCapture();
  }
  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  LayoutControls(rect.right, rect.bottom);
}

void MainWindow::Impl::ApplyViewVisibility() {
  bool show_search = IsSearchTabSelected();
  bool show_tree = show_tree_ && !show_search;
  bool show_value = show_value_ && !show_search;
  bool show_history = show_history_ && !show_search;
  ShowWindow(toolbar_.hwnd(), show_toolbar_ ? SW_SHOW : SW_HIDE);
  ShowWindow(browse_.address(), show_address_bar_ ? SW_SHOW : SW_HIDE);
  ShowWindow(browse_.go_button(), show_address_bar_ ? SW_SHOW : SW_HIDE);
  ShowWindow(tab_, show_tab_control_ ? SW_SHOW : SW_HIDE);
  ShowWindow(browse_.filter(), (show_value && show_filter_bar_) ? SW_SHOW : SW_HIDE);
  ShowWindow(tree_header_, show_tree ? SW_SHOW : SW_HIDE);
  ShowWindow(tree_close_btn_, show_tree ? SW_SHOW : SW_HIDE);
  ShowWindow(browse_.tree().hwnd(), show_tree ? SW_SHOW : SW_HIDE);
  ShowWindow(browse_.values().hwnd(), show_value ? SW_SHOW : SW_HIDE);
  if (value_grid_toolbar_) {
    ShowWindow(value_grid_toolbar_, show_value ? SW_SHOW : SW_HIDE);
  }
  ShowWindow(history_label_, show_history ? SW_SHOW : SW_HIDE);
  ShowWindow(history_close_btn_, show_history ? SW_SHOW : SW_HIDE);
  ShowWindow(history_list_, show_history ? SW_SHOW : SW_HIDE);
  ShowWindow(search_results_list_, show_search ? SW_SHOW : SW_HIDE);
  if (show_search && search_results_list_) {
    LONG_PTR style = GetWindowLongPtrW(search_results_list_, GWL_STYLE);
    if (style & LVS_SINGLESEL) {
      SetWindowLongPtrW(search_results_list_, GWL_STYLE, style & ~LVS_SINGLESEL);
      SetWindowPos(search_results_list_, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    }
  }
  ShowWindow(status_bar_, show_status_bar_ ? SW_SHOW : SW_HIDE);
  if (search_progress_) {
    bool show_progress = show_status_bar_ && show_search && search_running_ && !IsCompareTabSelected();
    ShowWindow(search_progress_, show_progress ? SW_SHOW : SW_HIDE);
  }

  RECT rect = {};
  GetClientRect(hwnd_, &rect);
  LayoutControls(rect.right, rect.bottom);
}

void MainWindow::Impl::ApplyTabSelection(int index) {
  if (index < 0 || static_cast<size_t>(index) >= tabs_.size()) {
    return;
  }
  const TabEntry& entry = tabs_[static_cast<size_t>(index)];
  if (entry.kind == TabEntry::Kind::kRegistry) {
    switch (entry.registry_mode) {
    case RegistryMode::kLocal: {
      SwitchToLocalRegistry();
      break;
    }
    case RegistryMode::kOffline:
      if (!entry.offline_path.empty()) {
        LoadOfflineRegistryFromPath(entry.offline_path, false);
      }
      break;
    case RegistryMode::kRemote:
      if (!entry.remote_machine.empty()) {
        remote_machine_ = entry.remote_machine;
      }
      if (registry_mode_ != RegistryMode::kRemote) {
        SwitchToRemoteRegistry();
      }
      break;
    }
    RestoreRegistryTabState(index);
  } else if (entry.kind == TabEntry::Kind::kSearch) {
    EnsureSearchTabResultsLoaded(entry.search_index);
  } else if (entry.kind == TabEntry::Kind::kRegFile) {
    SyncRegFileTabSelection();
  }
}

void MainWindow::Impl::ResetHiveListCache() {
  hive_list_loaded_ = false;
  hive_list_.clear();
  hive_roots_.reset();
}

void MainWindow::Impl::EnsureHiveListLoaded() {
  if (hive_list_loaded_) {
    return;
  }
  hive_list_loaded_ = true;
  hive_list_.clear();
  auto hive_roots = std::make_shared<std::unordered_set<std::wstring>>();
  hive_roots_ = hive_roots;

  HKEY hklm = nullptr;
  for (const auto& root : browse_.roots()) {
    if (EqualsInsensitive(root.display_name, L"HKEY_LOCAL_MACHINE")) {
      hklm = root.root;
      break;
    }
  }
  if (!hklm) {
    return;
  }
  util::UniqueHKey hive_key;
  if (RegOpenKeyExW(hklm, L"SYSTEM\\CurrentControlSet\\Control\\hivelist", 0, KEY_READ, hive_key.put()) != ERROR_SUCCESS) {
    return;
  }

  DWORD value_count = 0;
  DWORD max_name_len = 0;
  DWORD max_data_len = 0;
  if (RegQueryInfoKeyW(hive_key.get(), nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, &value_count, &max_name_len, &max_data_len, nullptr, nullptr) != ERROR_SUCCESS) {
    return;
  }

  std::vector<wchar_t> name_buffer(max_name_len + 1, L'\0');
  std::vector<BYTE> data_buffer(max_data_len > 0 ? max_data_len : 1);

  for (DWORD i = 0; i < value_count; ++i) {
    DWORD name_len = static_cast<DWORD>(name_buffer.size());
    DWORD data_len = static_cast<DWORD>(data_buffer.size());
    DWORD type = 0;
    LONG result = RegEnumValueW(hive_key.get(), i, name_buffer.data(), &name_len, nullptr, &type, data_buffer.data(), &data_len);
    if (result != ERROR_SUCCESS || name_len == 0 || data_len == 0) {
      continue;
    }
    if (type != REG_SZ && type != REG_EXPAND_SZ) {
      continue;
    }
    std::wstring name(name_buffer.data(), name_len);
    std::wstring data(reinterpret_cast<wchar_t*>(data_buffer.data()), data_len / sizeof(wchar_t));
    while (!data.empty() && data.back() == L'\0') {
      data.pop_back();
    }
    if (data.empty()) {
      continue;
    }
    data = NormalizeHiveFilePath(data);
    if (data.empty()) {
      continue;
    }
    std::wstring name_lower = ToLower(name);
    hive_list_.emplace(name_lower, std::move(data));
    hive_roots->insert(std::move(name_lower));
  }
}

std::wstring MainWindow::Impl::LookupHivePath(const RegistryNode& node, bool* is_root) {
  if (is_root) {
    *is_root = false;
  }
  EnsureHiveListLoaded();
  if (hive_list_.empty()) {
    return L"";
  }
  std::wstring nt_path = registry_path::BuildNative(node);
  if (nt_path.empty() && !node.root_name.empty()) {
    auto equals_root = [&](const wchar_t* name) -> bool { return EqualsInsensitive(node.root_name, name); };
    if (equals_root(L"REGISTRY")) {
      nt_path = L"\\REGISTRY";
    } else if (equals_root(L"HKLM") || equals_root(L"HKEY_LOCAL_MACHINE")) {
      nt_path = L"\\REGISTRY\\MACHINE";
    } else if (equals_root(L"HKU") || equals_root(L"HKEY_USERS")) {
      nt_path = L"\\REGISTRY\\USER";
    } else if (equals_root(L"HKCU") || equals_root(L"HKEY_CURRENT_USER")) {
      std::wstring sid = util::GetCurrentUserSidString();
      if (!sid.empty()) {
        nt_path = L"\\REGISTRY\\USER\\" + sid;
      }
    } else if (equals_root(L"HKCC") || equals_root(L"HKEY_CURRENT_CONFIG")) {
      nt_path = L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Hardware\\Profiles\\Current";
    } else if (equals_root(L"HKCR") || equals_root(L"HKEY_CLASSES_ROOT")) {
      nt_path = L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes";
    }
    if (!nt_path.empty() && !node.subkey.empty()) {
      nt_path += L"\\" + node.subkey;
    }
  }
  if (nt_path.empty()) {
    return L"";
  }
  std::wstring nt_lower = ToLower(nt_path);
  size_t best_len = 0;
  std::wstring best_path;
  for (const auto& entry : hive_list_) {
    const std::wstring& hive_key = entry.first;
    if (nt_lower.size() < hive_key.size()) {
      continue;
    }
    if (nt_lower.compare(0, hive_key.size(), hive_key) != 0) {
      continue;
    }
    if (nt_lower.size() > hive_key.size() && nt_lower[hive_key.size()] != L'\\') {
      continue;
    }
    if (hive_key.size() > best_len) {
      best_len = hive_key.size();
      best_path = entry.second;
    }
  }
  if (best_len > 0 && is_root) {
    *is_root = nt_lower.size() == best_len;
  }
  return best_path;
}

int MainWindow::Impl::KeyIconIndex(const RegistryNode& node, bool* is_link, bool* is_hive_root) {
  if (is_link) {
    *is_link = false;
  }
  if (is_hive_root) {
    *is_hive_root = false;
  }
  if (node.simulated) {
    return kFolderSimIconIndex;
  }
  std::wstring link_target;
  if (RegistryStore::QuerySymbolicLinkTarget(node, &link_target)) {
    if (is_link) {
      *is_link = true;
    }
    return kSymlinkIconIndex;
  }
  bool hive_root = false;
  std::wstring hive_path = LookupHivePath(node, &hive_root);
  if (!hive_path.empty() && hive_root && node.subkey.empty()) {
    if (node.root == HKEY_CURRENT_USER || EqualsInsensitive(node.root_name, L"HKEY_CURRENT_USER")) {
      hive_root = false;
    }
  }
  if (!hive_path.empty() && hive_root) {
    if (is_hive_root) {
      *is_hive_root = true;
    }
    return kDatabaseIconIndex;
  }
  return kFolderIconIndex;
}

std::wstring MainWindow::Impl::ResolveIconDir(bool use_light) const {
  if (IsIconSetName(icon_set_, kIconSetPhosphor)) {
    return L"";
  }
  if (IsIconSetName(icon_set_, kIconSetCustom)) {
    std::wstring root = util::JoinPath(util::GetAppDataFolder(), L"icons");
    if (root.empty()) {
      return L"";
    }
    std::wstring dark_dir = util::JoinPath(root, L"dark");
    std::wstring light_dir = util::JoinPath(root, L"light");
    if (IsDirectoryPath(dark_dir) && IsDirectoryPath(light_dir)) {
      return use_light ? light_dir : dark_dir;
    }
    return IsDirectoryPath(root) ? root : L"";
  }
  if (!IsKnownIconSetName(icon_set_)) {
    return L"";
  }
  std::wstring base = AssetsIconsRoot();
  if (base.empty()) {
    return L"";
  }
  std::wstring dir = util::JoinPath(util::JoinPath(base, icon_set_), use_light ? L"light" : L"dark");
  return IsDirectoryPath(dir) ? dir : L"";
}

std::wstring MainWindow::Impl::ResolveIconPath(const wchar_t* filename) const {
  if (!filename || !*filename || icon_dir_.empty()) {
    return L"";
  }
  return util::JoinPath(icon_dir_, filename);
}

bool MainWindow::Impl::ShouldUseLightIcons() const {
  switch (theme_mode_) {
  case ThemeMode::kDark:
    return true;
  case ThemeMode::kLight:
    return false;
  case ThemeMode::kSystem:
    return Theme::IsSystemDarkMode();
  case ThemeMode::kCustom:
  default:
    return Theme::UseDarkMode();
  }
}

void MainWindow::Impl::ApplyValueGridToolbarIcon() {
  if (!value_grid_toolbar_) {
    return;
  }
  const UINT dpi = win32::DpiForWindow(value_grid_toolbar_);
  const int size = util::ScaleForDpi(kToolbarGlyphSize, dpi);
  HICON icon = LoadThemeIcon(L"grid.ico", IDI_ICON_LIGHT_GRID, IDI_ICON_DARK_GRID,
                             kToolbarGlyphSize, dpi);
  HIMAGELIST images = ImageList_Create(size, size, ILC_COLOR32, 1, 1);
  if (!images) {
    if (icon) {
      DestroyIcon(icon);
    }
    return;
  }
  ImageList_SetBkColor(images, CLR_NONE);
  util::ImageListAddOrBlank(images, icon, size);
  if (icon) {
    DestroyIcon(icon);
  }
  SendMessageW(value_grid_toolbar_, TB_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(images));
  if (value_grid_image_list_) {
    ImageList_Destroy(value_grid_image_list_);
  }
  value_grid_image_list_ = images;
  InvalidateRect(value_grid_toolbar_, nullptr, TRUE);
}

void MainWindow::Impl::ApplyValueGridToolbarTheme() {
  if (!value_grid_toolbar_) {
    return;
  }
  Theme::Current().ApplyToToolbar(value_grid_toolbar_);
  HWND tooltip =
      reinterpret_cast<HWND>(SendMessageW(value_grid_toolbar_, TB_GETTOOLTIPS, 0, 0));
  if (tooltip) {
    AllowDarkModeForWindow(tooltip, Theme::UseDarkMode());
    SetWindowTheme(tooltip, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer",
                   nullptr);
  }
}

int MainWindow::Impl::ValueGridToggleWidth(HWND header) const {
  RECT client = {};
  if (!header || !GetClientRect(header, &client)) {
    return 0;
  }
  const int width = util::ScaleForDpi(kValueGridButtonWidth, win32::DpiForWindow(header));
  return std::min<int>(client.right - client.left, width);
}

void MainWindow::Impl::LayoutValueGridToolbar() {
  HWND list = browse_.values().hwnd();
  HWND header = list ? ListView_GetHeader(list) : nullptr;
  if (!value_grid_toolbar_ || !header || !IsWindowVisible(list)) {
    if (value_grid_toolbar_) {
      ShowWindow(value_grid_toolbar_, SW_HIDE);
    }
    return;
  }
  RECT header_rect = {};
  if (!GetWindowRect(header, &header_rect)) {
    return;
  }
  MapWindowPoints(nullptr, hwnd_, reinterpret_cast<POINT*>(&header_rect), 2);
  const int width = ValueGridToggleWidth(header);
  const int height = header_rect.bottom - header_rect.top;
  const int left = header_rect.right - width;
  RECT current = {};
  if (GetChildRectInParent(hwnd_, value_grid_toolbar_, &current) && current.left == left &&
      current.top == header_rect.top && current.right - current.left == width &&
      current.bottom - current.top == height && IsWindowVisible(value_grid_toolbar_)) {
    return;
  }
  SendMessageW(value_grid_toolbar_, TB_SETBUTTONSIZE, 0, MAKELPARAM(width, height));
  SetWindowPos(value_grid_toolbar_, HWND_TOP, left, header_rect.top, width, height,
               SWP_NOACTIVATE | SWP_SHOWWINDOW);
}

void MainWindow::Impl::SetValueGridEnabled(bool enabled, bool persist) {
  show_value_grid_ = enabled;
  if (HWND list = browse_.values().hwnd()) {
    RedrawWindow(list, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
  }
  if (value_grid_toolbar_) {
    SendMessageW(value_grid_toolbar_, TB_CHECKBUTTON, kValueGridButtonId,
                 MAKELPARAM(enabled ? TRUE : FALSE, 0));
  }
  if (persist) {
    SaveSettings();
  }
}

void MainWindow::Impl::EnsureValueGridToolbar() {
  HWND list = browse_.values().hwnd();
  HWND header = list ? ListView_GetHeader(list) : nullptr;
  if (!header) {
    return;
  }
  if (value_grid_toolbar_ && IsWindow(value_grid_toolbar_)) {
    LayoutValueGridToolbar();
    return;
  }
  value_grid_toolbar_ = CreateWindowExW(
      0, TOOLBARCLASSNAMEW, L"Value-list grid",
      WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS |
          CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE,
      0, 0, 0, 0, hwnd_,
      reinterpret_cast<HMENU>(static_cast<INT_PTR>(kValueGridButtonId)), instance_, nullptr);
  if (!value_grid_toolbar_) {
    return;
  }
  SendMessageW(value_grid_toolbar_, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
  SendMessageW(value_grid_toolbar_, TB_SETMAXTEXTROWS, 0, 0);
  SendMessageW(value_grid_toolbar_, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DOUBLEBUFFER);
  ApplyValueGridToolbarIcon();
  const LRESULT string_index = SendMessageW(value_grid_toolbar_, TB_ADDSTRINGW, 0,
                                            reinterpret_cast<LPARAM>(L"Grid lines"));
  TBBUTTON button = {};
  button.iBitmap = 0;
  button.idCommand = kValueGridButtonId;
  button.fsState = TBSTATE_ENABLED;
  button.fsStyle = BTNS_CHECK;
  button.iString = static_cast<INT_PTR>(string_index);
  SendMessageW(value_grid_toolbar_, TB_ADDBUTTONSW, 1, reinterpret_cast<LPARAM>(&button));
  ApplyValueGridToolbarTheme();
  SendMessageW(value_grid_toolbar_, TB_CHECKBUTTON, kValueGridButtonId,
               MAKELPARAM(show_value_grid_ ? TRUE : FALSE, 0));
  LayoutValueGridToolbar();
}

HICON MainWindow::Impl::LoadThemeIcon(const wchar_t* filename, int light_id, int dark_id, int size, UINT dpi) const {
  std::wstring path = ResolveIconPath(filename);
  HICON icon = nullptr;
  if (!path.empty()) {
    icon = util::LoadIconFromFile(path, size, dpi);
  }
  if (!icon) {
    icon = util::LoadIconResource(ShouldUseLightIcons() ? light_id : dark_id, size, dpi);
  }
  return icon;
}

ToolbarIcon MainWindow::Impl::MakeToolbarIcon(const wchar_t* filename, int light_id, int dark_id, bool use_light) const {
  ToolbarIcon icon;
  icon.resource_id = use_light ? light_id : dark_id;
  icon.path = ResolveIconPath(filename);
  return icon;
}

void MainWindow::Impl::ReloadThemeIcons() {
  UINT dpi = win32::DpiForWindow(hwnd_);
  bool use_light = ShouldUseLightIcons();
  icon_dir_ = ResolveIconDir(use_light);
  auto set_redraw = [](HWND hwnd, bool enable) {
    if (!hwnd) {
      return;
    }
    SendMessageW(hwnd, WM_SETREDRAW, enable ? TRUE : FALSE, 0);
  };
  set_redraw(toolbar_.hwnd(), false);
  set_redraw(browse_.tree().hwnd(), false);
  set_redraw(browse_.values().hwnd(), false);
  set_redraw(search_results_list_, false);
  set_redraw(browse_.go_button(), false);

  toolbar_.LoadIcons(
      {
          MakeToolbarIcon(L"local-registry.ico", IDI_ICON_LIGHT_LOCAL_REGISTRY, IDI_ICON_DARK_LOCAL_REGISTRY, use_light),
          MakeToolbarIcon(L"remote-registry.ico", IDI_ICON_LIGHT_REMOTE_REGISTRY, IDI_ICON_DARK_REMOTE_REGISTRY, use_light),
          MakeToolbarIcon(L"offline-registry.ico", IDI_ICON_LIGHT_OFFLINE_REGISTRY, IDI_ICON_DARK_OFFLINE_REGISTRY, use_light),
          MakeToolbarIcon(L"search.ico", IDI_ICON_LIGHT_SEARCH, IDI_ICON_DARK_SEARCH, use_light),
          MakeToolbarIcon(L"replace.ico", IDI_ICON_LIGHT_REPLACE, IDI_ICON_DARK_REPLACE, use_light),
          MakeToolbarIcon(L"export.ico", IDI_ICON_LIGHT_EXPORT, IDI_ICON_DARK_EXPORT, use_light),
          MakeToolbarIcon(L"undo.ico", IDI_ICON_LIGHT_UNDO, IDI_ICON_DARK_UNDO, use_light),
          MakeToolbarIcon(L"redo.ico", IDI_ICON_LIGHT_REDO, IDI_ICON_DARK_REDO, use_light),
          MakeToolbarIcon(L"copy.ico", IDI_ICON_LIGHT_COPY, IDI_ICON_DARK_COPY, use_light),
          MakeToolbarIcon(L"paste.ico", IDI_ICON_LIGHT_PASTE, IDI_ICON_DARK_PASTE, use_light),
          MakeToolbarIcon(L"delete.ico", IDI_ICON_LIGHT_DELETE, IDI_ICON_DARK_DELETE, use_light),
          MakeToolbarIcon(L"refresh.ico", IDI_ICON_LIGHT_REFRESH, IDI_ICON_DARK_REFRESH, use_light),
          MakeToolbarIcon(L"back.ico", IDI_ICON_LIGHT_BACK, IDI_ICON_DARK_BACK, use_light),
          MakeToolbarIcon(L"forward.ico", IDI_ICON_LIGHT_FORWARD, IDI_ICON_DARK_FORWARD, use_light),
          MakeToolbarIcon(L"up.ico", IDI_ICON_LIGHT_UP, IDI_ICON_DARK_UP, use_light),
      },
      kToolbarIconSize, kToolbarGlyphSize);

  BuildImageLists();
  if (browse_.tree().hwnd()) {
    browse_.tree().SetImageList(tree_images_);
  }
  if (browse_.values().hwnd()) {
    browse_.values().SetImageList(list_images_);
  }
  if (search_results_list_) {
    ListView_SetImageList(search_results_list_, list_images_, LVSIL_SMALL);
  }

  if (address_go_icon_) {
    DestroyIcon(address_go_icon_);
    address_go_icon_ = nullptr;
  }
  address_go_icon_ = LoadThemeIcon(L"forward.ico", IDI_ICON_LIGHT_FORWARD, IDI_ICON_DARK_FORWARD, kToolbarGlyphSize, dpi);
  ApplyValueGridToolbarIcon();
  LayoutValueGridToolbar();

  set_redraw(toolbar_.hwnd(), true);
  set_redraw(browse_.tree().hwnd(), true);
  set_redraw(browse_.values().hwnd(), true);
  set_redraw(search_results_list_, true);
  set_redraw(browse_.go_button(), true);
  if (toolbar_.hwnd()) {
    InvalidateRect(toolbar_.hwnd(), nullptr, TRUE);
  }
  if (browse_.tree().hwnd()) {
    InvalidateRect(browse_.tree().hwnd(), nullptr, TRUE);
  }
  if (browse_.values().hwnd()) {
    InvalidateRect(browse_.values().hwnd(), nullptr, TRUE);
  }
  if (search_results_list_) {
    InvalidateRect(search_results_list_, nullptr, TRUE);
  }
  if (browse_.go_button()) {
    InvalidateRect(browse_.go_button(), nullptr, TRUE);
  }
}

void MainWindow::Impl::LayoutControls(int width, int height) {
  if (width <= 0 || height <= 0) {
    return;
  }

  const int padding = 8;
  const int splitter_width = kSplitterWidth;
  UINT dpi = win32::DpiForWindow(hwnd_);
  const int address_height = CalcEditHeight(browse_.address(), ui_font_, util::ScaleForDpi(18, dpi));
  const int address_btn_width = std::max(util::ScaleForDpi(18, dpi), address_height);
  const int tabs_height = std::max(20, tab_height_);
  const int filter_height = address_height;
  const int filter_min_width = 160;
  const int filter_max_width = 260;
  const int filter_gap = 6;
  const int tree_header_height = util::ScaleForDpi(kPanelHeaderHeight, dpi);
  const int history_label_height = tree_header_height;
  const int close_size = util::ScaleForDpi(kPanelCloseSize, dpi);
  const int close_inset = util::ScaleForDpi(kPanelCloseInset, dpi);
  int status_height = 0;
  if (status_bar_ && show_status_bar_) {
    RECT sb_rect = {};
    GetWindowRect(status_bar_, &sb_rect);
    status_height = sb_rect.bottom - sb_rect.top;
    if (status_height <= 0) {
      status_height = 20;
    }
  }
  const bool show_search = IsSearchTabSelected();
  const bool show_tree = show_tree_ && !show_search;
  const bool show_history = show_history_ && !show_search;
  const bool show_value = show_value_ && !show_search;

  int y = kMainVerticalGap;

  const bool dragging_splitter = splitter_dragging_ || history_splitter_dragging_;
  auto place = [&](HWND hwnd, int x, int y_pos, int w, int h) {
    if (!hwnd) {
      return;
    }
    UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;
    if (!dragging_splitter) {
      flags |= SWP_NOREDRAW;
    }
    SetWindowPos(hwnd, nullptr, x, y_pos, w, h, flags);
  };

  if (show_toolbar_) {
    SendMessageW(toolbar_.hwnd(), TB_AUTOSIZE, 0, 0);
    SIZE ideal = {};
    SendMessageW(toolbar_.hwnd(), TB_GETMAXSIZE, 0, reinterpret_cast<LPARAM>(&ideal));
    RECT tb_rect = {};
    GetWindowRect(toolbar_.hwnd(), &tb_rect);
    int toolbar_height = tb_rect.bottom - tb_rect.top;
    if (toolbar_height <= 0) {
      toolbar_height = ideal.cy;
    }
    const int toolbar_area_width = std::max(0, width - padding * 2);
    place(toolbar_.hwnd(), padding, y, toolbar_area_width, toolbar_height);
    y += toolbar_height;
  }
  if (show_address_bar_) {
    int address_width = width - padding * 2 - address_btn_width - 2;
    if (address_width < 120) {
      address_width = 120;
    }
    place(browse_.address(), padding, y, address_width, address_height);
    place(browse_.go_button(), padding + address_width, y, address_btn_width, address_height);
    SetEditMargins(browse_.address(), 6, 6);
    SetEditVerticalRect(browse_.address(), ui_font_, 2, 6, 6);
    y += address_height + kMainVerticalGap;
  }

  int tabs_width = width - padding * 2;
  bool show_tabs = show_tab_control_ && tab_;
  bool show_filter = show_value && show_filter_bar_ && browse_.filter();
  bool show_tab_row = show_tabs || show_filter;
  if (show_tab_row) {
    if (show_tabs && show_filter) {
      int available = std::max(0, tabs_width);
      int min_needed = kTabMinWidth + filter_min_width + filter_gap;
      if (available >= min_needed) {
        int target_width = ClampValue(available / 4, filter_min_width, filter_max_width);
        int filter_width = std::min(target_width, std::max(filter_min_width, available - kTabMinWidth - filter_gap));
        tabs_width = std::max(kTabMinWidth, available - filter_width - filter_gap);
        int filter_y = y + std::max(0, (tabs_height - filter_height) / 2);
        place(tab_, padding, y, tabs_width, tabs_height);
        place(browse_.filter(), padding + tabs_width + filter_gap, filter_y, filter_width, filter_height);
        SetEditMargins(browse_.filter(), 6, 6);
        SetEditVerticalRect(browse_.filter(), ui_font_, 2, 6, 6);
        ShowWindow(browse_.filter(), SW_SHOW);
      } else {
        show_filter = false;
      }
    }
    if (show_tabs && !show_filter) {
      place(tab_, padding, y, tabs_width, tabs_height);
      if (browse_.filter()) {
        ShowWindow(browse_.filter(), SW_HIDE);
      }
    } else if (!show_tabs && show_filter) {
      int available = std::max(0, tabs_width);
      int filter_width = ClampValue(available, filter_min_width, filter_max_width);
      int filter_y = y + std::max(0, (tabs_height - filter_height) / 2);
      int filter_x = padding + std::max(0, tabs_width - filter_width);
      place(browse_.filter(), filter_x, filter_y, filter_width, filter_height);
      SetEditMargins(browse_.filter(), 6, 6);
      SetEditVerticalRect(browse_.filter(), ui_font_, 2, 6, 6);
      ShowWindow(browse_.filter(), SW_SHOW);
    }
    y += tabs_height + kMainVerticalGap;
  } else {
    if (tab_) {
      ShowWindow(tab_, SW_HIDE);
    }
    if (browse_.filter()) {
      ShowWindow(browse_.filter(), SW_HIDE);
    }
  }

  int status_top = height - status_height;
  int content_left = 0;
  int content_right = width;
  if (show_status_bar_ && status_bar_) {
    place(status_bar_, content_left, status_top, content_right - content_left, status_height);
    SendMessageW(status_bar_, WM_SIZE, 0, 0);
  }

  int history_splitter_height = show_history ? kHistorySplitterHeight : 0;
  int history_gap = show_history ? kHistoryGap : 0;
  int content_total_height = std::max(0, status_top - y);
  int min_history = kMinHistoryHeight;
  int max_history = std::max(min_history, content_total_height - kHistoryMaxPadding);
  int history_height = show_history ? ClampValue(history_height_, min_history, max_history) : 0;
  if (show_history) {
    history_height_ = history_height;
  }
  int history_top = status_top - history_height;
  if (show_history) {
    int history_width = content_right - content_left;
    place(history_label_, content_left, history_top, history_width, history_label_height);
    place(history_close_btn_, content_left + history_width - close_inset - close_size,
          history_top + (history_label_height - close_size) / 2, close_size, close_size);
    place(history_list_, content_left,
          history_top + history_label_height - kPanelBorderOverlap,
          history_width,
          history_height - history_label_height + kPanelBorderOverlap);
  }

  int splitter_bottom = show_history ? (history_top - history_gap) : history_top;
  int splitter_top = show_history ? (splitter_bottom - history_splitter_height) : history_top;
  if (show_history) {
    history_splitter_rect_.left = content_left;
    history_splitter_rect_.right = content_right;
    history_splitter_rect_.top = splitter_top;
    history_splitter_rect_.bottom = splitter_bottom;
  } else {
    history_splitter_rect_ = {};
  }
  int content_bottom = show_history ? splitter_top : status_top;
  int available_width = content_right - content_left;
  int min_tree = kMinTreeWidth;
  int min_list = kMinValueListWidth;
  int max_tree = std::max(min_tree, available_width - min_list - splitter_width);
  int tree_width = show_tree ? std::min(tree_width_, max_tree) : 0;
  tree_width = show_tree ? std::max(tree_width, min_tree) : 0;
  int list_x = show_tree ? (content_left + tree_width + splitter_width) : content_left;
  int list_width = content_right - list_x;
  int content_height = std::max(0, content_bottom - y);
  int tree_content_height = std::max(0, content_height - (show_tree ? tree_header_height : 0));
  if (show_tree) {
    place(tree_header_, content_left, y, tree_width, tree_header_height);
    place(tree_close_btn_, content_left + tree_width - close_inset - close_size,
          y + (tree_header_height - close_size) / 2, close_size, close_size);
    place(browse_.tree().hwnd(), content_left,
          y + tree_header_height - kPanelBorderOverlap,
          tree_width, tree_content_height + kPanelBorderOverlap);
    splitter_rect_.left = content_left + tree_width;
    splitter_rect_.right = splitter_rect_.left + splitter_width;
    splitter_rect_.top = y;
    splitter_rect_.bottom = y + content_height;
  } else {
    splitter_rect_ = {};
  }
  if (show_search) {
    place(search_results_list_, content_left, y, content_right - content_left, content_height);
  } else {
    place(browse_.values().hwnd(), list_x, y, list_width, content_height);
  }
  LayoutValueGridToolbar();

  UpdateStatus();
  if (!dragging_splitter) {
    RedrawWindow(hwnd_, nullptr, nullptr,
                 RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_ERASE | RDW_UPDATENOW);
  }
  drag_layout_valid_ = false;
}

} // namespace regkit
