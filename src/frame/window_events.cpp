// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

namespace regkit {
using namespace window_detail;

MainWindow::Impl::~Impl() = default;

bool MainWindow::Impl::Create(HINSTANCE instance) {
  instance_ = instance;
  last_search_.criteria.search_keys = false;

  WNDCLASSEXW wc = {};
  wc.cbSize = sizeof(wc);
  wc.lpfnWndProc = MainWindow::Impl::WndProc;
  wc.hInstance = instance;
  wc.lpszClassName = kRegeditWindowClassName;
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
  wc.hIconSm = LoadIconW(instance, MAKEINTRESOURCEW(IDI_APPICON));
  wc.hbrBackground = nullptr;

  RegisterClassExW(&wc);

  std::wstring title = L"RegKit V";
  title.append(REGKIT_VERSION_STR_W);
  if (util::IsProcessTrustedInstaller()) {
    title.append(L" - [TrustedInstaller]");
  } else if (util::IsProcessSystem()) {
    title.append(L" - [SYSTEM]");
  } else if (util::IsProcessElevated()) {
    title.append(L" - [Administrator]");
  }
  hwnd_ = CreateWindowExW(0, wc.lpszClassName, title.c_str(), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, CW_USEDEFAULT, CW_USEDEFAULT, 1200, 800, nullptr, nullptr, instance, this);
  if (hwnd_) {
    SetPropW(hwnd_, kRegKitWindowProperty, reinterpret_cast<HANDLE>(static_cast<INT_PTR>(1)));
    DragAcceptFiles(hwnd_, TRUE);
  }
  return hwnd_ != nullptr;
}

void MainWindow::Impl::ActivateRegeditCompatibilityMode() {
  bool was_enabled = regedit_compatibility_mode_;
  regedit_compatibility_mode_ = true;
  EnsureRegeditCompatControls();
  if (!was_enabled) {
    RefreshVisibleRegistryTreeLayout(true);
  }
  FocusAddressBarForExternalJump(true);
}

void MainWindow::Impl::Show(int cmd_show) {
  int show_cmd = cmd_show;
  if (window_placement_loaded_ && window_width_ > 0 && window_height_ > 0) {
    show_cmd = window_maximized_ ? SW_MAXIMIZE : SW_SHOWNORMAL;
  } else if (window_placement_loaded_ && window_maximized_) {
    show_cmd = SW_MAXIMIZE;
  }
  ShowWindow(hwnd_, show_cmd);
  UpdateWindow(hwnd_);
  if (regedit_compatibility_mode_) {
    EnsureRegeditCompatControls();
    FocusAddressBarForExternalJump(true);
  }
  PostMessageW(hwnd_, frame::message_id::kDeferredStartup, 0, 0);
  PostMessageW(hwnd_, frame::message_id::kLoadTraces, 0, 0);
  PostMessageW(hwnd_, frame::message_id::kLoadDefaults, 0, 0);
}

void MainWindow::Impl::QueueExternalJump(const std::wstring& target) {
  queued_external_jump_target_ = target;
}

void MainWindow::Impl::CreateRegeditCompatControls() {
  if (regedit_compat_edit_ || regedit_compat_tree_ || regedit_compat_list_) {
    return;
  }
  DWORD base_style = WS_CHILD | WS_TABSTOP;
  regedit_compat_edit_ = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", base_style | ES_AUTOHSCROLL, -2000, -2000, 8, 8, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRegeditCompatEditId)), instance_, nullptr);
  regedit_compat_tree_ = CreateWindowExW(WS_EX_CLIENTEDGE, WC_TREEVIEWW, L"", base_style | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS, -2000, -2000, 8, 8, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRegeditCompatTreeId)), instance_, nullptr);
  regedit_compat_list_ = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"", base_style | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SINGLESEL, -2000, -2000, 8, 8, hwnd_, reinterpret_cast<HMENU>(static_cast<INT_PTR>(kRegeditCompatListId)), instance_, nullptr);
  if (regedit_compat_edit_) {
    SetWindowSubclass(regedit_compat_edit_, AddressEditProc, kAddressSubclassId, reinterpret_cast<DWORD_PTR>(this));
  }
  if (regedit_compat_list_) {
    LVCOLUMNW col = {};
    col.mask = LVCF_TEXT | LVCF_WIDTH;
    col.cx = 240;
    col.pszText = const_cast<wchar_t*>(L"Name");
    ListView_InsertColumn(regedit_compat_list_, 0, &col);
  }
  PopulateRegeditCompatTree();
}

void MainWindow::Impl::EnsureRegeditCompatControls() {
  if (!hwnd_ || (regedit_compat_edit_ && regedit_compat_tree_ && regedit_compat_list_)) {
    return;
  }
  CreateRegeditCompatControls();
  ApplyFont(regedit_compat_edit_, ui_font_);
  ApplyFont(regedit_compat_tree_, ui_font_);
  ApplyFont(regedit_compat_list_, ui_font_);
  const Theme& theme = Theme::Current();
  theme.ApplyToTreeView(regedit_compat_tree_);
  theme.ApplyToListView(regedit_compat_list_);
  if (regedit_compat_edit_) {
    SetWindowTheme(regedit_compat_edit_, Theme::UseDarkMode() ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    SetEditMargins(regedit_compat_edit_, 6, 6);
    SetEditVerticalRect(regedit_compat_edit_, ui_font_, 2, 6, 6);
  }
}

void MainWindow::Impl::PopulateRegeditCompatTree() {
  if (!regedit_compat_tree_) {
    return;
  }
  TreeView_DeleteAllItems(regedit_compat_tree_);
  regedit_compat_nodes_.clear();

  auto add_node = [&](HTREEITEM parent, const std::wstring& text, const std::wstring& path, bool has_placeholder) {
    auto node = std::make_unique<RegeditCompatTreeNode>();
    node->path = path;
    RegeditCompatTreeNode* raw = node.get();
    regedit_compat_nodes_.push_back(std::move(node));
    TVINSERTSTRUCTW insert = {};
    insert.hParent = parent;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM;
    insert.item.pszText = const_cast<wchar_t*>(text.c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(raw);
    HTREEITEM item = TreeView_InsertItem(regedit_compat_tree_, &insert);
    if (item && has_placeholder) {
      TVINSERTSTRUCTW placeholder = {};
      placeholder.hParent = item;
      placeholder.hInsertAfter = TVI_LAST;
      placeholder.item.mask = TVIF_TEXT;
      placeholder.item.pszText = const_cast<wchar_t*>(L"");
      TreeView_InsertItem(regedit_compat_tree_, &placeholder);
    }
    return item;
  };

  HTREEITEM computer = add_node(TVI_ROOT, L"Computer", L"Computer", true);
  static const wchar_t* kCompatHives[] = {L"HKEY_CLASSES_ROOT", L"HKEY_CURRENT_USER", L"HKEY_LOCAL_MACHINE", L"HKEY_USERS", L"HKEY_CURRENT_CONFIG"};
  for (const wchar_t* hive : kCompatHives) {
    add_node(computer, hive, hive, true);
  }
  TreeView_Expand(regedit_compat_tree_, computer, TVE_EXPAND);
  regedit_compat_selected_key_path_.clear();
  if (regedit_compat_edit_) {
    SetWindowTextW(regedit_compat_edit_, L"");
  }
}

void MainWindow::Impl::PopulateRegeditCompatChildren(HTREEITEM item) {
  if (!regedit_compat_tree_ || !item) {
    return;
  }
  TVITEMW tvi = {};
  tvi.mask = TVIF_PARAM;
  tvi.hItem = item;
  if (!TreeView_GetItem(regedit_compat_tree_, &tvi)) {
    return;
  }
  auto* node = reinterpret_cast<RegeditCompatTreeNode*>(tvi.lParam);
  if (!node || node->populated) {
    return;
  }
  node->populated = true;

  while (HTREEITEM child = TreeView_GetChild(regedit_compat_tree_, item)) {
    TreeView_DeleteItem(regedit_compat_tree_, child);
  }

  if (EqualsInsensitive(node->path, L"Computer")) {
    static const wchar_t* kCompatHives[] = {L"HKEY_CLASSES_ROOT", L"HKEY_CURRENT_USER", L"HKEY_LOCAL_MACHINE", L"HKEY_USERS", L"HKEY_CURRENT_CONFIG"};
    for (const wchar_t* hive : kCompatHives) {
      auto child_node = std::make_unique<RegeditCompatTreeNode>();
      child_node->path = hive;
      RegeditCompatTreeNode* raw = child_node.get();
      regedit_compat_nodes_.push_back(std::move(child_node));
      TVINSERTSTRUCTW insert = {};
      insert.hParent = item;
      insert.hInsertAfter = TVI_LAST;
      insert.item.mask = TVIF_TEXT | TVIF_PARAM;
      insert.item.pszText = const_cast<wchar_t*>(hive);
      insert.item.lParam = reinterpret_cast<LPARAM>(raw);
      HTREEITEM child_item = TreeView_InsertItem(regedit_compat_tree_, &insert);
      if (child_item) {
        TVINSERTSTRUCTW placeholder = {};
        placeholder.hParent = child_item;
        placeholder.hInsertAfter = TVI_LAST;
        placeholder.item.mask = TVIF_TEXT;
        placeholder.item.pszText = const_cast<wchar_t*>(L"");
        TreeView_InsertItem(regedit_compat_tree_, &placeholder);
      }
    }
    return;
  }

  RegistryNode registry_node;
  if (!ResolvePathToNode(node->path, &registry_node)) {
    return;
  }
  auto subkeys = RegistryStore::EnumSubKeyNames(registry_node, true);
  for (const auto& subkey : subkeys) {
    auto child_node = std::make_unique<RegeditCompatTreeNode>();
    child_node->path = node->path + L"\\" + subkey;
    RegeditCompatTreeNode* raw = child_node.get();
    regedit_compat_nodes_.push_back(std::move(child_node));
    TVINSERTSTRUCTW insert = {};
    insert.hParent = item;
    insert.hInsertAfter = TVI_LAST;
    insert.item.mask = TVIF_TEXT | TVIF_PARAM;
    insert.item.pszText = const_cast<wchar_t*>(subkey.c_str());
    insert.item.lParam = reinterpret_cast<LPARAM>(raw);
    HTREEITEM child_item = TreeView_InsertItem(regedit_compat_tree_, &insert);
    if (child_item) {
      TVINSERTSTRUCTW placeholder = {};
      placeholder.hParent = child_item;
      placeholder.hInsertAfter = TVI_LAST;
      placeholder.item.mask = TVIF_TEXT;
      placeholder.item.pszText = const_cast<wchar_t*>(L"");
      TreeView_InsertItem(regedit_compat_tree_, &placeholder);
    }
  }
}

void MainWindow::Impl::UpdateRegeditCompatList(const std::wstring& key_path) {
  if (!regedit_compat_list_) {
    return;
  }
  regedit_compat_list_names_.clear();
  ListView_DeleteAllItems(regedit_compat_list_);

  RegistryNode node;
  if (!ResolvePathToNode(key_path, &node)) {
    return;
  }
  RegistryStore::KeyEnumResult enum_result;
  bool names_reserved = false;
  int index = 0;
  RegistryStore::EnumKeyStreaming(
      node, true, false, false, &enum_result,
      [&](const ValueInfo& value, const BYTE*, DWORD) {
        if (!names_reserved) {
          if (enum_result.info_valid) {
            regedit_compat_list_names_.reserve(enum_result.info.value_count);
          }
          names_reserved = true;
        }
        std::wstring display_name = value.name.empty() ? L"(Default)" : value.name;
        regedit_compat_list_names_.push_back(value.name);
        LVITEMW item = {};
        item.mask = LVIF_TEXT;
        item.iItem = index++;
        item.pszText = display_name.data();
        ListView_InsertItem(regedit_compat_list_, &item);
        return true;
      },
      {});
}

HTREEITEM MainWindow::Impl::FindRegeditCompatItemByPath(const std::wstring& key_path) {
  if (!regedit_compat_tree_) {
    return nullptr;
  }
  HTREEITEM current = TreeView_GetRoot(regedit_compat_tree_);
  if (!current) {
    return nullptr;
  }
  std::vector<std::wstring> parts = registry_path::Split(key_path);
  if (!parts.empty() && EqualsInsensitive(parts.front(), L"Computer")) {
    parts.erase(parts.begin());
  }
  if (parts.empty()) {
    return current;
  }
  for (const auto& part : parts) {
    PopulateRegeditCompatChildren(current);
    TreeView_Expand(regedit_compat_tree_, current, TVE_EXPAND);
    HTREEITEM child = FindChildByText(regedit_compat_tree_, current, part);
    if (!child) {
      return nullptr;
    }
    current = child;
  }
  return current;
}

void MainWindow::Impl::SyncRegeditCompatControls(const std::wstring& key_path, const std::wstring& value_name) {
  if (!regedit_compat_edit_ || !regedit_compat_tree_ || !regedit_compat_list_) {
    return;
  }
  syncing_regedit_compat_controls_ = true;
  auto clear_sync = [&]() {
    syncing_regedit_compat_controls_ = false;
  };

  regedit_compat_selected_key_path_ = key_path;
  if (key_path.empty()) {
    SetWindowTextW(regedit_compat_edit_, L"Computer");
    HTREEITEM root = TreeView_GetRoot(regedit_compat_tree_);
    if (root) {
      TreeView_SelectItem(regedit_compat_tree_, root);
    }
    regedit_compat_list_names_.clear();
    ListView_DeleteAllItems(regedit_compat_list_);
    clear_sync();
    return;
  }

  SetWindowTextW(regedit_compat_edit_, key_path.c_str());
  HTREEITEM item = FindRegeditCompatItemByPath(key_path);
  if (item) {
    TreeView_SelectItem(regedit_compat_tree_, item);
    TreeView_EnsureVisible(regedit_compat_tree_, item);
  }
  UpdateRegeditCompatList(key_path);
  if (!value_name.empty()) {
    for (size_t i = 0; i < regedit_compat_list_names_.size(); ++i) {
      if (regedit_compat_list_names_[i] == value_name) {
        ListView_SetItemState(regedit_compat_list_, static_cast<int>(i), LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(regedit_compat_list_, static_cast<int>(i), FALSE);
        break;
      }
    }
  }
  clear_sync();
}

void MainWindow::Impl::HandleRegeditCompatTreeSelection(NMTREEVIEWW* info) {
  if (syncing_regedit_compat_controls_) {
    return;
  }
  if (!info) {
    return;
  }
  auto* node = reinterpret_cast<RegeditCompatTreeNode*>(info->itemNew.lParam);
  if (!node) {
    return;
  }
  if (EqualsInsensitive(node->path, L"Computer")) {
    regedit_compat_selected_key_path_.clear();
    regedit_compat_pending_key_path_.clear();
    regedit_compat_pending_value_name_.clear();
    KillTimer(hwnd_, kRegeditCompatApplyTimerId);
    if (regedit_compat_edit_) {
      SetWindowTextW(regedit_compat_edit_, L"Computer");
    }
    return;
  }
  regedit_compat_selected_key_path_ = node->path;
  regedit_compat_pending_value_name_.clear();
  QueueRegeditCompatNavigation(node->path);
  if (regedit_compat_edit_) {
    SetWindowTextW(regedit_compat_edit_, node->path.c_str());
  }
  UpdateRegeditCompatList(node->path);
}

void MainWindow::Impl::HandleRegeditCompatListSelection(NMLISTVIEW* info) {
  if (syncing_regedit_compat_controls_) {
    return;
  }
  if (!info || info->iItem < 0 || (info->uNewState & LVIS_SELECTED) == 0) {
    return;
  }
  if (static_cast<size_t>(info->iItem) >= regedit_compat_list_names_.size()) {
    return;
  }
  regedit_compat_pending_value_name_ = regedit_compat_list_names_[static_cast<size_t>(info->iItem)];
  if (!regedit_compat_selected_key_path_.empty()) {
    QueueRegeditCompatNavigation(regedit_compat_selected_key_path_);
  }
}

void MainWindow::Impl::QueueRegeditCompatNavigation(const std::wstring& key_path) {
  if (!hwnd_) {
    return;
  }
  regedit_compat_pending_key_path_ = key_path;
  KillTimer(hwnd_, kRegeditCompatApplyTimerId);
  SetTimer(hwnd_, kRegeditCompatApplyTimerId, kRegeditCompatApplyDelayMs, nullptr);
}

void MainWindow::Impl::BeginJumpUiBatch() {
  if (jump_ui_batch_active_) {
    return;
  }
  jump_ui_batch_active_ = true;
  if (browse_.tree().hwnd()) {
    SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, FALSE, 0);
  }
  if (browse_.values().hwnd()) {
    SendMessageW(browse_.values().hwnd(), WM_SETREDRAW, FALSE, 0);
  }
}

void MainWindow::Impl::EndJumpUiBatch() {
  if (!jump_ui_batch_active_) {
    return;
  }
  jump_ui_batch_active_ = false;
  if (browse_.tree().hwnd()) {
    SendMessageW(browse_.tree().hwnd(), WM_SETREDRAW, TRUE, 0);
    InvalidateRect(browse_.tree().hwnd(), nullptr, TRUE);
  }
  if (browse_.values().hwnd()) {
    SendMessageW(browse_.values().hwnd(), WM_SETREDRAW, TRUE, 0);
    InvalidateRect(browse_.values().hwnd(), nullptr, TRUE);
  }
}

void MainWindow::Impl::ApplyTreeSelectionEffects(RegistryNode* node) {
  UpdateAddressBar(node);
  UpdateValueListForNode(node);
  MarkTreeStateDirty();
}

void MainWindow::Impl::ApplyPendingRegeditCompatNavigation() {
  std::wstring key_path = std::move(regedit_compat_pending_key_path_);
  std::wstring value_name = std::move(regedit_compat_pending_value_name_);
  regedit_compat_pending_key_path_.clear();
  regedit_compat_pending_value_name_.clear();
  if (key_path.empty()) {
    return;
  }
  std::wstring resolved_key_path;
  std::wstring resolved_value_name;
  if (!ResolveExternalJumpTarget(key_path, &resolved_key_path, &resolved_value_name)) {
    return;
  }
  if (value_name.empty()) {
    value_name = resolved_value_name;
  }
  if (!NavigateToResolvedExternalJump(resolved_key_path, value_name, false)) {
    return;
  }
  if (value_name.empty()) {
    return;
  }
  if (!SelectValueByName(value_name) && browse_.current_node()) {
    pending_external_value_key_path_ = resolved_key_path;
    pending_external_value_name_ = value_name;
    if (!value_list_loading_) {
      UpdateValueListForNode(browse_.current_node());
    }
  }
}

void MainWindow::Impl::FocusAddressBarForExternalJump(bool defer_if_needed) {
  if (!hwnd_) {
    return;
  }
  HWND target = nullptr;
  if (regedit_compatibility_mode_) {
    if (regedit_compat_edit_) {
      target = regedit_compat_edit_;
    } else if (browse_.tree().hwnd()) {
      target = browse_.tree().hwnd();
    }
  } else if (browse_.address()) {
    target = browse_.address();
  }
  if (!target) {
    return;
  }
  if (IsWindowVisible(hwnd_) && !IsIconic(hwnd_)) {
    SetFocus(target);
    if (target == browse_.address() || target == regedit_compat_edit_) {
      SendMessageW(target, EM_SETSEL, 0, -1);
    }
    return;
  }
  if (defer_if_needed) {
    PostMessageW(hwnd_, frame::message_id::kFocusAddressBar, 0, 0);
  }
}

bool MainWindow::Impl::TranslateAccelerator(const MSG& msg) {
  if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN) {
    const bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
    HWND focus = GetFocus();
    auto is_text_input = [](HWND hwnd) -> bool {
      if (!hwnd) {
        return false;
      }
      wchar_t cls[64] = {};
      GetClassNameW(hwnd, cls, static_cast<int>(_countof(cls)));
      if (_wcsicmp(cls, L"Edit") == 0) {
        return true;
      }
      if (_wcsicmp(cls, L"RichEdit20W") == 0 || _wcsicmp(cls, L"RichEdit20A") == 0) {
        return true;
      }
      if (_wcsicmp(cls, L"ComboBox") == 0 || _wcsicmp(cls, L"ComboBoxEx32") == 0) {
        return true;
      }
      HWND parent = GetParent(hwnd);
      if (parent) {
        GetClassNameW(parent, cls, static_cast<int>(_countof(cls)));
        if (_wcsicmp(cls, L"ComboBox") == 0 || _wcsicmp(cls, L"ComboBoxEx32") == 0) {
          return true;
        }
      }
      return false;
    };
    const bool focus_edit = is_text_input(focus);

    if ((alt && msg.wParam == 'D') || (ctrl && !alt && msg.wParam == 'L')) {
      if (browse_.address()) {
        SetFocus(browse_.address());
        SendMessageW(browse_.address(), EM_SETSEL, 0, -1);
        return true;
      }
    }

    if (ctrl && !alt) {
      if (shift && msg.wParam == 'C' && !focus_edit) {
        HandleMenuCommand(cmd::kEditCopyKey);
        return true;
      }
      switch (msg.wParam) {
      case 'A':
        if (SelectAllInFocusedList()) {
          return true;
        }
        if (focus_edit && focus) {
          SendMessageW(focus, EM_SETSEL, 0, -1);
          return true;
        }
        break;
      case 'C':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditCopy);
          return true;
        }
        return false;
      case 'V':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditPaste);
          return true;
        }
        return false;
      case 'X':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditDelete);
          return true;
        }
        return false;
      case 'Z':
        if (!focus_edit) {
          HandleMenuCommand(cmd::kEditUndo);
          return true;
        }
        return false;
      case 'Y':
        if (focus_edit && focus) {
          SendMessageW(focus, EM_REDO, 0, 0);
          return true;
        }
        HandleMenuCommand(cmd::kEditRedo);
        return true;
      case 'F':
        HandleMenuCommand(cmd::kEditFind);
        return true;
      case 'G':
        HandleMenuCommand(cmd::kEditGoTo);
        return true;
      case 'H':
        HandleMenuCommand(cmd::kEditReplace);
        return true;
      case 'S':
        HandleMenuCommand(cmd::kFileSave);
        return true;
      case 'E':
        HandleMenuCommand(cmd::kFileExport);
        return true;
      case 'N':
        OpenLocalRegistryTab();
        return true;
      }
    }

    if (!ctrl && !alt) {
      if (msg.wParam == VK_DELETE && !focus_edit) {
        HandleMenuCommand(cmd::kEditDelete);
        return true;
      }
      if (msg.wParam == VK_F2 && !focus_edit) {
        HandleMenuCommand(cmd::kEditRename);
        return true;
      }
      if (msg.wParam == VK_F5) {
        HandleMenuCommand(cmd::kViewRefresh);
        return true;
      }
    }

    if (focus_edit) {
      if (msg.wParam == VK_DELETE || msg.wParam == VK_BACK) {
        return false;
      }
      if (ctrl && !alt) {
        switch (msg.wParam) {
        case 'C':
        case 'V':
        case 'X':
        case 'Z':
        case 'Y':
          return false;
        default:
          break;
        }
      }
    }
  }
  if (accelerators_) {
    return ::TranslateAcceleratorW(hwnd_, accelerators_, const_cast<MSG*>(&msg)) != 0;
  }
  return false;
}

LRESULT CALLBACK MainWindow::Impl::WndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
  if (message == WM_NCCREATE) {
    auto* create = reinterpret_cast<CREATESTRUCTW*>(lparam);
    auto* self = static_cast<MainWindow::Impl*>(create->lpCreateParams);
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    self->hwnd_ = hwnd;
  }

  auto* self = reinterpret_cast<MainWindow::Impl*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  if (!self) {
    return DefWindowProcW(hwnd, message, wparam, lparam);
  }
  return self->HandleMessage(message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::AddressEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
  if (message == WM_SETTEXT) {
    const wchar_t* text = reinterpret_cast<const wchar_t*>(lparam);
    if (self && hwnd == self->regedit_compat_edit_ && text && *text) {
      if (self->syncing_regedit_compat_controls_) {
        return DefSubclassProc(hwnd, message, wparam, lparam);
      }
      if (_wcsicmp(text, L"Computer") == 0) {
        self->regedit_compat_selected_key_path_.clear();
        self->regedit_compat_pending_key_path_.clear();
        self->regedit_compat_pending_value_name_.clear();
        KillTimer(self->hwnd_, kRegeditCompatApplyTimerId);
      } else {
        self->regedit_compat_selected_key_path_ = text;
        self->regedit_compat_pending_value_name_.clear();
        self->QueueRegeditCompatNavigation(text);
      }
    }
  }
  if (message == WM_KEYDOWN && wparam == VK_RETURN) {
    if (self && self->hwnd_) {
      if (hwnd == self->regedit_compat_edit_) {
        wchar_t buffer[2048] = {};
        GetWindowTextW(hwnd, buffer, static_cast<int>(_countof(buffer)));
        self->regedit_compat_selected_key_path_ = buffer;
        self->regedit_compat_pending_value_name_.clear();
        self->regedit_compat_pending_key_path_ = buffer;
        KillTimer(self->hwnd_, kRegeditCompatApplyTimerId);
        self->ApplyPendingRegeditCompatNavigation();
      } else {
        SendMessageW(self->hwnd_, frame::message_id::kAddressEnter, 0, 0);
      }
    }
    return 0;
  }
  if (message == WM_CHAR && wparam == VK_RETURN) {
    return 0;
  }
  if (message == WM_SETFOCUS) {
    LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
    SendMessageW(hwnd, EM_SETSEL, 0, -1);
    return result;
  }
  if (message == WM_KEYUP) {
    LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
    auto* window = reinterpret_cast<MainWindow::Impl*>(ref_data);
    if (window) {
      window->ApplyAutoCompleteTheme();
    }
    return result;
  }
  if (message == WM_LBUTTONDOWN) {
    if (GetFocus() != hwnd) {
      LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
      SendMessageW(hwnd, EM_SETSEL, 0, -1);
      return result;
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::FilterEditProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR /*ref_data*/) {
  if (message == WM_KEYDOWN && wparam == VK_RETURN) {
    return 0;
  }
  if (message == WM_CHAR && wparam == VK_RETURN) {
    return 0;
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::TabProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
  if (!self) {
    return DefSubclassProc(hwnd, message, wparam, lparam);
  }

  switch (message) {
  case WM_ERASEBKGND:
    return 1;
  case WM_MOUSEMOVE: {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    self->UpdateTabHotState(hwnd, pt);
    if (!self->tab_mouse_tracking_) {
      TRACKMOUSEEVENT tme = {};
      tme.cbSize = sizeof(tme);
      tme.dwFlags = TME_LEAVE;
      tme.hwndTrack = hwnd;
      TrackMouseEvent(&tme);
      self->tab_mouse_tracking_ = true;
    }
    return 0;
  }
  case WM_MOUSELEAVE:
    self->tab_mouse_tracking_ = false;
    if (self->tab_hot_index_ != -1 || self->tab_close_hot_index_ != -1) {
      self->tab_hot_index_ = -1;
      self->tab_close_hot_index_ = -1;
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    return 0;
  case WM_LBUTTONDOWN: {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    TCHITTESTINFO hit = {};
    hit.pt = pt;
    int index = TabCtrl_HitTest(hwnd, &hit);
    RECT close_rect = {};
    if (self->GetTabCloseRect(index, &close_rect) && PtInRect(&close_rect, pt)) {
      self->tab_close_down_index_ = index;
      SetCapture(hwnd);
      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }
    if (self->tab_close_down_index_ != -1) {
      self->tab_close_down_index_ = -1;
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    break;
  }
  case WM_LBUTTONUP: {
    if (self->tab_close_down_index_ >= 0) {
      POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
      int close_index = self->tab_close_down_index_;
      self->tab_close_down_index_ = -1;
      ReleaseCapture();
      RECT close_rect = {};
      if (self->GetTabCloseRect(close_index, &close_rect) && PtInRect(&close_rect, pt)) {
        self->CloseTab(close_index);
        self->tab_hot_index_ = -1;
        self->tab_close_hot_index_ = -1;
      }
      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }
    break;
  }
  case WM_CAPTURECHANGED:
    if (self->tab_close_down_index_ >= 0) {
      self->tab_close_down_index_ = -1;
      InvalidateRect(hwnd, nullptr, FALSE);
    }
    break;
  case WM_PAINT: {
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd, &ps);
    self->PaintTabControl(hwnd, hdc);
    EndPaint(hwnd, &ps);
    return 0;
  }
  default:
    break;
  }

  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::ListViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
  if (message == WM_LBUTTONDOWN && self && hwnd == self->browse_.values().hwnd()) {
    POINT pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
    LVHITTESTINFO hit = {};
    hit.pt = pt;
    int index = ListView_HitTest(hwnd, &hit);
    DWORD now = GetTickCount();
    if (index >= 0 && index == self->last_value_click_index_) {
      self->last_value_click_delta_ = now - self->last_value_click_time_;
      self->last_value_click_delta_valid_ = true;
    } else {
      self->last_value_click_delta_valid_ = false;
    }
    self->last_value_click_time_ = now;
    self->last_value_click_index_ = index;
  }
  if (message == WM_KEYDOWN && self && hwnd == self->browse_.values().hwnd()) {
    if (wparam == VK_RETURN) {
      self->value_activate_from_key_ = true;
      self->last_value_click_delta_valid_ = false;
    }
  }
  if (message == WM_CHAR && self && hwnd == self->browse_.values().hwnd()) {
    wchar_t ch = static_cast<wchar_t>(wparam);
    if (ch == L'\b' || (iswprint(ch) && ch != L'\r' && ch != L'\n' && ch != L'\t')) {
      self->HandleTypeToSelectList(ch);
      return 0;
    }
  }
  if (message == WM_SETFOCUS || message == WM_KILLFOCUS) {
    SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
    if (self && hwnd == self->history_list_) {
      ListView_SetItemState(hwnd, -1, 0, LVIS_FOCUSED);
    }
  }
  if (message == WM_UPDATEUISTATE) {
    LRESULT result = DefSubclassProc(hwnd, message, wparam, lparam);
    SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
    if (self && hwnd == self->history_list_) {
      ListView_SetItemState(hwnd, -1, 0, LVIS_FOCUSED);
    }
    return result;
  }
  if (message == WM_ERASEBKGND) {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().PanelBrush());
    return 1;
  }
  if (message == WM_CTLCOLOREDIT) {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    SetTextColor(hdc, Theme::Current().TextColor());
    SetBkColor(hdc, Theme::Current().PanelColor());
    return reinterpret_cast<LRESULT>(Theme::Current().PanelBrush());
  }
  if (message == WM_PRINTCLIENT) {
    HDC hdc = reinterpret_cast<HDC>(wparam);
    RECT rect = {};
    GetClientRect(hwnd, &rect);
    FillRect(hdc, &rect, Theme::Current().PanelBrush());
  }
  if (message == WM_THEMECHANGED) {
    if (self) {
      InvalidateRect(hwnd, nullptr, TRUE);
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::TreeViewProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
  if (message == WM_KEYDOWN && self && hwnd == self->browse_.tree().hwnd()) {
    switch (wparam) {
    case VK_RIGHT:
      self->browse_.set_tree_type_select_descend(true);
      break;
    case VK_LEFT:
    case VK_UP:
    case VK_DOWN:
    case VK_HOME:
    case VK_END:
    case VK_PRIOR:
    case VK_NEXT:
      self->browse_.set_tree_type_select_descend(false);
      break;
    default:
      break;
    }
  }
  if (message == WM_CHAR && self && hwnd == self->browse_.tree().hwnd()) {
    wchar_t ch = static_cast<wchar_t>(wparam);
    if (ch == L'\b' || (iswprint(ch) && ch != L'\r' && ch != L'\n' && ch != L'\t')) {
      self->HandleTypeToSelectTree(ch);
      return 0;
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

LRESULT CALLBACK MainWindow::Impl::HeaderProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam, UINT_PTR, DWORD_PTR ref_data) {
  auto* self = reinterpret_cast<MainWindow::Impl*>(ref_data);
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

    HFONT old_font = nullptr;
    if (self && self->ui_font_) {
      old_font = reinterpret_cast<HFONT>(SelectObject(hdc, self->ui_font_));
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
  if (message == WM_CONTEXTMENU) {
    if (self) {
      HWND value_header = ListView_GetHeader(self->browse_.values().hwnd());
      HWND history_header = ListView_GetHeader(self->history_list_);
      HWND search_header = ListView_GetHeader(self->search_results_list_);
      if (hwnd == value_header) {
        POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        if (screen_pt.x == -1 && screen_pt.y == -1) {
          RECT rect = {};
          GetWindowRect(hwnd, &rect);
          screen_pt.x = rect.left + 12;
          screen_pt.y = rect.bottom - 4;
        }
        self->ShowValueHeaderMenu(screen_pt);
        return 0;
      }
      if (hwnd == history_header) {
        POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        if (screen_pt.x == -1 && screen_pt.y == -1) {
          RECT rect = {};
          GetWindowRect(hwnd, &rect);
          screen_pt.x = rect.left + 12;
          screen_pt.y = rect.bottom - 4;
        }
        self->ShowHistoryHeaderMenu(screen_pt);
        return 0;
      }
      if (hwnd == search_header) {
        POINT screen_pt = {GET_X_LPARAM(lparam), GET_Y_LPARAM(lparam)};
        if (screen_pt.x == -1 && screen_pt.y == -1) {
          RECT rect = {};
          GetWindowRect(hwnd, &rect);
          screen_pt.x = rect.left + 12;
          screen_pt.y = rect.bottom - 4;
        }
        self->ShowSearchHeaderMenu(screen_pt);
        return 0;
      }
    }
  }
  return DefSubclassProc(hwnd, message, wparam, lparam);
}

} // namespace regkit
