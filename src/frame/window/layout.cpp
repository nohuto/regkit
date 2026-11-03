// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

#include "registry/key_algorithms.h"

#include <win32ui/dialog/dialog_metrics.h>
#include <win32ui/dialog/edit.h>
#include <win32ui/gdi/fonts.h>
#include <win32ui/gdi/icon_loader.h>
#include <win32ui/handle_owner.h>
#include <win32ui/window.h>
#include <win32ui/window_metrics.h>

namespace regkit
{
using namespace window_detail;

// single line controls get taller when the UI font needs it
int MainWindow::Impl::TextRowHeight(int nominal, int padding) const
{
    const UINT dpi = win32ui::DpiForWindow(hwnd_);
    return std::max(win32ui::Scale(nominal, dpi), static_cast<int>(win32ui::FontMetrics(ui_font_).tmHeight) + win32ui::Scale(padding, dpi));
}

void MainWindow::Impl::LayoutContent(bool dragging)
{
    const bool show_search = IsSearchTabSelected();
    const bool show_tree = settings_.show_tree && !show_search;
    const bool show_history = settings_.show_history && !settings_.simple_mode && !show_search;
    const UINT dpi = win32ui::DpiForWindow(hwnd_);
    const int header_height = TextRowHeight(kPanelHeaderHeight);
    const int close_size = win32ui::Scale(kPanelCloseSize, dpi);
    const int close_inset = win32ui::Scale(kPanelCloseInset, dpi);
    const int left = content_rect_.left;
    const int right = content_rect_.right;
    const int width = right - left;
    const int top = content_rect_.top;
    const int bottom = content_rect_.bottom;

    struct Placement
    {
        HWND hwnd;
        RECT rect;
    };
    Placement placements[9] = {};
    int count = 0;
    auto place = [&](HWND hwnd, int x, int y, int w, int h) {
        if (hwnd)
        {
            placements[count++] = {hwnd, {x, y, x + w, y + h}};
        }
    };
    RECT dirty = {};

    // saved sizes use DIPs, narrow windows shrink shown sizes without changing saved values
    int history_bar_top = bottom;
    history_shown_height_ = 0;
    if (show_history)
    {
        const int minimum = win32ui::Scale(kMinHistoryHeight, dpi);
        const int history_height = ClampValue(win32ui::Scale(settings_.history_height, dpi), minimum, std::max(minimum, bottom - top - win32ui::Scale(kHistoryMaxPadding, dpi)));
        const int history_top = bottom - history_height;
        const int bar_height = win32ui::Scale(kHistorySplitterHeight, dpi);
        place(history_label_, left, history_top, width, header_height);
        place(history_close_btn_, right - close_inset - close_size, history_top + (header_height - close_size) / 2, close_size, close_size);
        place(history_list_, left, history_top + header_height - kPanelBorderOverlap, width, history_height - header_height + kPanelBorderOverlap);
        history_bar_top = history_top - win32ui::Scale(kHistoryGap, dpi) - bar_height;
        place(history_splitter_, left, history_bar_top, width, bar_height);
        history_shown_height_ = history_height;
    }
    const int content_height = std::max(0, history_bar_top - top);

    int list_x = left;
    tree_shown_width_ = 0;
    if (show_tree)
    {
        const int minimum = win32ui::Scale(kMinTreeWidth, dpi);
        const int bar_width = win32ui::Scale(kSplitterWidth, dpi);
        const int tree_width = ClampValue(win32ui::Scale(settings_.tree_width, dpi), minimum, std::max(minimum, width - win32ui::Scale(kMinValueListWidth, dpi) - bar_width));
        place(tree_header_, left, top, tree_width, header_height);
        place(tree_close_btn_, left + tree_width - close_inset - close_size, top + (header_height - close_size) / 2, close_size, close_size);
        place(browse_.tree().hwnd(), left, top + header_height - kPanelBorderOverlap, tree_width, std::max(0, content_height - header_height) + kPanelBorderOverlap);
        place(tree_splitter_, left + tree_width, top, bar_width, content_height);
        list_x = left + tree_width + bar_width;
        tree_shown_width_ = tree_width;
    }
    if (show_search)
    {
        place(search_results_list_, left, top, width, content_height);
    }
    else
    {
        place(browse_.values().hwnd(), list_x, top, right - list_x, content_height);
    }

    // drag redraws only old/new control areas, otherwise caller repaints everything
    const UINT flags = SWP_NOZORDER | SWP_NOACTIVATE | (dragging ? 0 : SWP_NOREDRAW);
    HDWP batch = BeginDeferWindowPos(count);
    for (int i = 0; i < count; ++i)
    {
        const Placement& p = placements[i];
        RECT old_rect = {};
        if (dragging && win32ui::ChildRect(hwnd_, p.hwnd, &old_rect))
        {
            UnionRect(&dirty, &dirty, &old_rect);
            UnionRect(&dirty, &dirty, &p.rect);
        }
        if (batch)
        {
            batch = DeferWindowPos(batch, p.hwnd, nullptr, p.rect.left, p.rect.top, p.rect.right - p.rect.left, p.rect.bottom - p.rect.top, flags);
        }
    }
    // failed batches are abandoned, reposition every control
    if (!batch || !EndDeferWindowPos(batch))
    {
        for (int i = 0; i < count; ++i)
        {
            const Placement& p = placements[i];
            SetWindowPos(p.hwnd, nullptr, p.rect.left, p.rect.top, p.rect.right - p.rect.left, p.rect.bottom - p.rect.top, flags);
        }
    }
    ShowWindow(tree_splitter_, show_tree ? SW_SHOWNA : SW_HIDE);
    ShowWindow(history_splitter_, show_history ? SW_SHOWNA : SW_HIDE);
    if (dragging)
    {
        RedrawWindow(hwnd_, &dirty, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_UPDATENOW);
    }
}

// clamp splitter drag like LayoutContent and store in DIPs
void MainWindow::Impl::ResizePane(HWND splitter, int size, bool done)
{
    const UINT dpi = win32ui::DpiForWindow(hwnd_);
    const bool tree = splitter == tree_splitter_;
    // click without movement keeps shown size, narrow windows can show less than saved size
    const bool moved = size != (tree ? tree_shown_width_ : history_shown_height_);
    if (moved && tree)
    {
        const int minimum = win32ui::Scale(kMinTreeWidth, dpi);
        const int maximum = content_rect_.right - content_rect_.left - win32ui::Scale(kMinValueListWidth, dpi) - win32ui::Scale(kSplitterWidth, dpi);
        settings_.tree_width = win32ui::Unscale(ClampValue(size, minimum, std::max(minimum, maximum)), dpi);
    }
    else if (moved)
    {
        const int minimum = win32ui::Scale(kMinHistoryHeight, dpi);
        const int maximum = content_rect_.bottom - content_rect_.top - win32ui::Scale(kHistoryMaxPadding, dpi);
        settings_.history_height = win32ui::Unscale(ClampValue(size, minimum, std::max(minimum, maximum)), dpi);
    }
    splitter_dragging_ = !done;
    if (!done)
    {
        LayoutContent(true);
        return;
    }
    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    LayoutControls(rect.right, rect.bottom);
}

void MainWindow::Impl::ApplyViewVisibility()
{
    FlushTreeRedraw();
    bool show_search = IsSearchTabSelected();
    bool show_tree = settings_.show_tree && !show_search;
    bool show_value = show_value_ && !show_search;
    bool show_history = settings_.show_history && !settings_.simple_mode && !show_search;
    ShowWindow(toolbar_, settings_.show_toolbar && !settings_.simple_mode ? SW_SHOW : SW_HIDE);
    ShowWindow(browse_.address(), settings_.show_address_bar ? SW_SHOW : SW_HIDE);
    ShowWindow(browse_.go_button(), settings_.show_address_bar ? SW_SHOW : SW_HIDE);
    ShowWindow(tab_, settings_.show_tab_control ? SW_SHOW : SW_HIDE);
    ShowWindow(browse_.filter(), (show_value && settings_.show_filter_bar) ? SW_SHOW : SW_HIDE);
    ShowWindow(filter_clear_btn_, (show_value && settings_.show_filter_bar) ? SW_SHOW : SW_HIDE);
    ShowWindow(tree_header_, show_tree ? SW_SHOW : SW_HIDE);
    ShowWindow(tree_close_btn_, show_tree ? SW_SHOW : SW_HIDE);
    ShowWindow(browse_.tree().hwnd(), show_tree ? SW_SHOW : SW_HIDE);
    ShowWindow(browse_.values().hwnd(), show_value ? SW_SHOW : SW_HIDE);
    ShowWindow(history_label_, show_history ? SW_SHOW : SW_HIDE);
    ShowWindow(history_close_btn_, show_history ? SW_SHOW : SW_HIDE);
    ShowWindow(history_list_, show_history ? SW_SHOW : SW_HIDE);
    ShowWindow(search_results_list_, show_search ? SW_SHOW : SW_HIDE);
    if (show_search && search_results_list_)
    {
        LONG_PTR style = GetWindowLongPtrW(search_results_list_, GWL_STYLE);
        if (style & LVS_SINGLESEL)
        {
            SetWindowLongPtrW(search_results_list_, GWL_STYLE, style & ~LVS_SINGLESEL);
            SetWindowPos(search_results_list_, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
        }
    }
    ShowWindow(status_bar_, settings_.show_status_bar ? SW_SHOW : SW_HIDE);

    RECT rect = {};
    GetClientRect(hwnd_, &rect);
    LayoutControls(rect.right, rect.bottom);
}

void MainWindow::Impl::ApplyTabSelection(int index)
{
    if (index < 0 || static_cast<size_t>(index) >= tabs_.size())
    {
        return;
    }
    const TabEntry& entry = tabs_[static_cast<size_t>(index)];
    if (entry.kind == TabEntry::Kind::kRegistry)
    {
        // remote/offline tabs without a session connect when shown, including restored or disconnected tabs
        const bool connect = !entry.session && entry.registry_mode != RegistryMode::kLocal;
        if (!ActivateTabTree(index) && !connect)
        {
            ResumeTabTree(index);
            return;
        }
        SuspendTreeRedraw();
        bool shown = false;
        if (entry.session)
        {
            ShowSession(std::shared_ptr<RegistrySession>(entry.session));
            shown = true;
        }
        else if (entry.registry_mode == RegistryMode::kLocal && entry.registry_view)
        {
            ShowSession(LocalViewSession(entry.registry_view));
            shown = true;
        }
        else if (entry.registry_mode == RegistryMode::kRemote && !entry.remote_machine.empty())
        {
            shown = ConnectRemoteRegistry(std::wstring(entry.remote_machine), false);
        }
        else if (entry.registry_mode == RegistryMode::kOffline && !entry.offline_path.empty())
        {
            shown = LoadOfflineRegistryFromPath(std::wstring(entry.offline_path), false);
        }
        if (!shown && entry.registry_mode != RegistryMode::kLocal)
        {
            // disconnected, empty tree under its own label
            auto disconnected = std::make_shared<RegistrySession>();
            disconnected->mode = entry.registry_mode;
            disconnected->remote_machine = entry.remote_machine;
            ShowSession(disconnected);
            tabs_[static_cast<size_t>(index)].session = nullptr;
        }
        else if (!shown)
        {
            ShowSession(local_session_);
        }
        RestoreRegistryTabState(index);
    }
    else if (entry.kind == TabEntry::Kind::kSearch)
    {
        EnsureSearchTabResultsLoaded(entry.search_index);
    }
    else if (entry.kind == TabEntry::Kind::kRegFile)
    {
        SyncRegFileTabSelection();
    }
}

constexpr DWORD kMaxHiveValueBytes = 64 * 1024;

void MainWindow::Impl::ResetHiveListCache()
{
    hive_list_loaded_ = false;
    hive_list_.clear();
    hive_roots_.reset();
}

void MainWindow::Impl::EnsureHiveListLoaded()
{
    if (hive_list_loaded_)
    {
        return;
    }
    hive_list_loaded_ = true;
    hive_list_.clear();
    auto hive_roots = std::make_shared<std::unordered_set<std::wstring>>();
    hive_roots_ = hive_roots;

    HKEY hklm = nullptr;
    for (const auto& root : browse_.roots())
    {
        if (EqualsInsensitive(root.display_name, L"HKEY_LOCAL_MACHINE"))
        {
            hklm = root.root;
            break;
        }
    }
    if (!hklm)
    {
        return;
    }
    // cache mounted hive paths for file lookup and root icons
    win32ui::UniqueHKey hive_key;
    if (RegOpenKeyExW(hklm, L"SYSTEM\\CurrentControlSet\\Control\\hivelist", 0, KEY_READ, hive_key.put()) !=
        ERROR_SUCCESS)
    {
        return;
    }
    registry_backend::EnumerateKey(
        registry_backend::RegistryKeyHandle(std::move(hive_key)),
        true,
        true,
        false,
        nullptr,
        [&](const ValueInfo& value, const BYTE* data, DWORD size) {
            if (!data || size == 0 || value.name.empty() || (value.type != REG_SZ && value.type != REG_EXPAND_SZ))
            {
                return true;
            }
            std::wstring path(reinterpret_cast<const wchar_t*>(data), size / sizeof(wchar_t));
            while (!path.empty() && path.back() == L'\0')
            {
                path.pop_back();
            }
            path = NormalizeHiveFilePath(path);
            if (path.empty())
            {
                return true;
            }
            std::wstring name_lower = ToLower(value.name);
            hive_list_.emplace(name_lower, std::move(path));
            hive_roots->insert(std::move(name_lower));
            return true;
        },
        nullptr,
        kMaxHiveValueBytes,
        nullptr
    );
}

std::wstring MainWindow::Impl::LookupHivePath(const RegistryNode& node, bool* is_root)
{
    std::wstring nt_path = registry_path::BuildNative(node);
    if (nt_path.empty() && !node.root_name.empty())
    {
        // nodes without predefined handles still identify their root
        RegistryNode named = node;
        named.root = registry_path::RootFromName(node.root_name);
        nt_path = named.root == HKEY_CLASSES_ROOT ? registry_path::JoinSubkey(L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes", node.subkey) : registry_path::BuildNative(named);
    }
    return LookupNativeHivePath(nt_path, is_root);
}

std::wstring MainWindow::Impl::LookupNativeHivePath(const std::wstring& nt_path, bool* is_root)
{
    if (is_root)
    {
        *is_root = false;
    }
    EnsureHiveListLoaded();
    if (hive_list_.empty() || nt_path.empty())
    {
        return L"";
    }
    std::wstring nt_lower = ToLower(nt_path);
    size_t best_len = 0;
    std::wstring best_path;
    for (const auto& entry : hive_list_)
    {
        const std::wstring& hive_key = entry.first;
        if (nt_lower.size() < hive_key.size())
        {
            continue;
        }
        if (nt_lower.compare(0, hive_key.size(), hive_key) != 0)
        {
            continue;
        }
        if (nt_lower.size() > hive_key.size() && nt_lower[hive_key.size()] != L'\\')
        {
            continue;
        }
        if (hive_key.size() > best_len)
        {
            best_len = hive_key.size();
            best_path = entry.second;
        }
    }
    if (best_len > 0 && is_root)
    {
        *is_root = nt_lower.size() == best_len;
    }
    return best_path;
}

int MainWindow::Impl::KeyIconIndex(const RegistryNode& node, bool* is_link, bool* is_hive_root)
{
    if (is_link)
    {
        *is_link = false;
    }
    if (is_hive_root)
    {
        *is_hive_root = false;
    }
    if (node.simulated)
    {
        return kFolderSimIconIndex;
    }
    const KeyInspection inspection = RegistryStore::InspectKey(node, false);
    const bool denied = inspection.denied;
    if (inspection.link)
    {
        if (is_link)
        {
            *is_link = true;
        }
        return kSymlinkIconIndex;
    }
    bool hive_root = false;
    std::wstring hive_path = LookupHivePath(node, &hive_root);
    if (!hive_path.empty() && hive_root && node.subkey.empty())
    {
        if (node.root == HKEY_CURRENT_USER || EqualsInsensitive(node.root_name, L"HKEY_CURRENT_USER"))
        {
            // treat HKCU as an alias
            hive_root = false;
        }
    }
    if (!hive_path.empty() && hive_root)
    {
        if (is_hive_root)
        {
            *is_hive_root = true;
        }
        return denied ? kDatabaseDeniedIconIndex : kDatabaseIconIndex;
    }
    return denied ? kFolderDeniedIconIndex : inspection.is_volatile ? kFolderVolatileIconIndex
                                                                    : kFolderIconIndex;
}

std::wstring MainWindow::Impl::ResolveIconDir(bool use_light) const
{
    if (util::EqualsInsensitive(settings_.icon_set, kIconSetCustom))
    {
        // JoinPath of an empty folder would give a path relative to the current directory
        const std::wstring folder = util::GetAppDataFolder();
        if (folder.empty())
        {
            return L"";
        }
        std::wstring root = util::JoinPath(folder, L"icons");
        std::wstring dark_dir = util::JoinPath(root, L"dark");
        std::wstring light_dir = util::JoinPath(root, L"light");
        if (util::IsDirectory(dark_dir.c_str()) && util::IsDirectory(light_dir.c_str()))
        {
            return use_light ? light_dir : dark_dir;
        }
        return util::IsDirectory(root.c_str()) ? root : L"";
    }
    const std::wstring base = AssetsIconsRoot();
    if (!util::EqualsInsensitive(settings_.icon_set, kIconSetClassic) || base.empty())
    {
        return L"";
    }
    const std::wstring dir = util::JoinPath(base, kIconSetClassic);
    return util::IsDirectory(dir.c_str()) ? dir : L"";
}

std::wstring MainWindow::Impl::ResolveIconPath(const wchar_t* filename) const
{
    if (!filename || !*filename || icon_dir_.empty())
    {
        return L"";
    }
    return util::JoinPath(icon_dir_, filename);
}

void MainWindow::Impl::ApplyGridToolbarIcons()
{
    for (HWND list : {browse_.values().hwnd(), history_list_, search_results_list_})
    {
        if (list)
        {
            win32ui::RefreshListView(list);
        }
    }
}

void MainWindow::Impl::SetValueGridEnabled(bool enabled, bool persist)
{
    settings_.show_value_grid = enabled;
    win32ui::SetAllListGridLines(enabled);
    if (persist)
    {
        SaveSettings();
    }
}

HICON MainWindow::Impl::ToolbarIconFor(int command, UINT dpi) const
{
    struct IconFile
    {
        int command;
        const wchar_t* file;
        int resource_id;
    };
    static constexpr IconFile kIcons[] = {
        {cmd::kRegistryLocal, L"local-registry.ico", IDI_ICON_LOCAL_REGISTRY},
        {cmd::kRegistryNetwork, L"remote-registry.ico", IDI_ICON_REMOTE_REGISTRY},
        {cmd::kRegistryOffline, L"offline-registry.ico", IDI_ICON_OFFLINE_REGISTRY},
        {cmd::kEditFind, L"search.ico", IDI_ICON_SEARCH},
        {cmd::kEditReplace, L"replace.ico", IDI_ICON_REPLACE},
        {cmd::kEditUndo, L"undo.ico", IDI_ICON_UNDO},
        {cmd::kEditRedo, L"redo.ico", IDI_ICON_REDO},
        {cmd::kEditCopy, L"copy.ico", IDI_ICON_COPY},
        {cmd::kEditPaste, L"paste.ico", IDI_ICON_PASTE},
        {cmd::kEditDelete, L"delete.ico", IDI_ICON_DELETE},
        {cmd::kViewRefresh, L"refresh.ico", IDI_ICON_REFRESH},
        {cmd::kNavBack, L"back.ico", IDI_ICON_BACK},
        {cmd::kNavForward, L"forward.ico", IDI_ICON_FORWARD},
        {cmd::kNavUp, L"up.ico", IDI_ICON_UP},
    };
    const auto found = std::ranges::find(kIcons, command, &IconFile::command);
    if (found == std::end(kIcons))
    {
        return nullptr;
    }
    const std::wstring path = ResolveIconPath(found->file);
    HICON icon = path.empty() ? nullptr : win32ui::LoadIconFromFile(path, dpi);
    return icon ? icon : win32ui::LoadIconResource(found->resource_id + (win32ui::ThemeLightGlyphs() ? IDI_ICON_LIGHT_OFFSET : 0), dpi);
}

void MainWindow::Impl::ReloadThemeIcons()
{
    UINT dpi = win32ui::DpiForWindow(hwnd_);
    bool use_light = win32ui::ThemeLightGlyphs();
    icon_dir_ = ResolveIconDir(use_light);
    // pause redraw while replacing image lists, toolbar manages its own redraw, hidden windows stay hidden
    const win32ui::RedrawScope tree_redraw(browse_.tree().hwnd(), RDW_INVALIDATE | RDW_ERASE);
    const win32ui::RedrawScope values_redraw(browse_.values().hwnd(), RDW_INVALIDATE | RDW_ERASE);
    const win32ui::RedrawScope results_redraw(search_results_list_, RDW_INVALIDATE | RDW_ERASE);
    const win32ui::RedrawScope go_redraw(browse_.go_button(), RDW_INVALIDATE | RDW_ERASE);

    win32ui::RefreshToolbar(toolbar_);

    BuildImageLists();
    for (const auto& tree : browse_.trees())
    {
        tree->SetImageList(tree_images_);
    }
    if (browse_.values().hwnd())
    {
        browse_.values().SetImageList(list_images_);
    }
    if (search_results_list_)
    {
        ListView_SetImageList(search_results_list_, list_images_, LVSIL_SMALL);
    }

    if (address_go_icon_)
    {
        DestroyIcon(address_go_icon_);
        address_go_icon_ = nullptr;
    }
    address_go_icon_ = win32ui::LoadIconResource(use_light ? IDI_ICON_LIGHT_GO : IDI_ICON_DARK_GO, dpi);
    ApplyGridToolbarIcons();
}

void MainWindow::Impl::LayoutControls(int width, int height)
{
    if (width <= 0 || height <= 0)
    {
        return;
    }

    UINT dpi = win32ui::DpiForWindow(hwnd_);
    const int padding = win32ui::Scale(8, dpi);
    const int gap = win32ui::Scale(kMainVerticalGap, dpi);
    const int address_height = TextRowHeight(win32ui::metrics::kControlHeight);
    const int address_btn_width = std::max(win32ui::Scale(18, dpi), address_height);
    const int tabs_height = std::max(win32ui::Scale(20, dpi), tab_height_);
    const int filter_height = address_height;
    const int filter_min_width = win32ui::Scale(160, dpi);
    const int filter_max_width = win32ui::Scale(260, dpi);
    const int filter_gap = win32ui::Scale(6, dpi);
    const int tab_min_width = win32ui::Scale(win32ui::kTabMinWidthDip, dpi);
    int status_height = 0;
    if (status_bar_ && settings_.show_status_bar)
    {
        RECT sb_rect = {};
        GetWindowRect(status_bar_, &sb_rect);
        status_height = sb_rect.bottom - sb_rect.top;
        if (status_height <= 0)
        {
            status_height = win32ui::Scale(20, dpi);
        }
    }
    const bool show_search = IsSearchTabSelected();
    const bool show_value = show_value_ && !show_search;

    int y = gap;

    const bool dragging_splitter = splitter_dragging_;
    auto place = [&](HWND hwnd, int x, int y_pos, int w, int h) {
        if (!hwnd)
        {
            return;
        }
        UINT flags = SWP_NOZORDER | SWP_NOACTIVATE;
        if (!dragging_splitter)
        {
            flags |= SWP_NOREDRAW;
        }
        SetWindowPos(hwnd, nullptr, x, y_pos, w, h, flags);
    };

    if (settings_.show_toolbar && !settings_.simple_mode)
    {
        const int toolbar_height = win32ui::ToolbarHeight(toolbar_);
        place(toolbar_, 0, 0, width, toolbar_height);
        y = toolbar_height;
    }
    if (settings_.show_address_bar)
    {
        const int address_width = std::max(width - padding * 2 - address_btn_width - win32ui::Scale(2, dpi), win32ui::Scale(120, dpi));
        place(browse_.address(), padding, y, address_width, address_height);
        place(browse_.go_button(), padding + address_width, y, address_btn_width, address_height);
        win32ui::SetEditPadding(browse_.address(), 6, 2);
        y += address_height + gap;
    }

    int tabs_width = width - padding * 2;
    bool show_tabs = settings_.show_tab_control && tab_;
    bool show_filter = show_value && settings_.show_filter_bar && browse_.filter();
    bool show_tab_row = show_tabs || show_filter;
    // tabs wrap beside the filter up to a quarter of the window, then scroll in one row
    int rows_height = tabs_height;
    const int max_tabs_height = std::max(tabs_height, height / 4);
    if (show_tab_row)
    {
        if (show_tabs && show_filter)
        {
            int available = std::max(0, tabs_width);
            int min_needed = tab_min_width + filter_min_width + filter_gap;
            if (available >= min_needed)
            {
                int target_width = ClampValue(available / 4, filter_min_width, filter_max_width);
                int filter_width =
                    std::min(target_width, std::max(filter_min_width, available - tab_min_width - filter_gap));
                tabs_width = std::max(tab_min_width, available - filter_width - filter_gap);
                int filter_y = y + std::max(0, (tabs_height - filter_height) / 2);
                int edit_width = std::max(filter_min_width / 2, filter_width - address_btn_width);
                rows_height = win32ui::FitTabStrip(tab_, tabs_width, max_tabs_height);
                place(tab_, padding, y, tabs_width, rows_height);
                place(browse_.filter(), padding + tabs_width + filter_gap, filter_y, edit_width, filter_height);
                place(filter_clear_btn_, padding + tabs_width + filter_gap + edit_width, filter_y, address_btn_width, filter_height);
                win32ui::SetEditPadding(browse_.filter(), 6, 2);
                ShowWindow(browse_.filter(), SW_SHOW);
                ShowWindow(filter_clear_btn_, SW_SHOW);
            }
            else
            {
                // keep tabs usable when there isnt space for both controls
                show_filter = false;
            }
        }
        if (show_tabs && !show_filter)
        {
            rows_height = win32ui::FitTabStrip(tab_, tabs_width, max_tabs_height);
            place(tab_, padding, y, tabs_width, rows_height);
            if (browse_.filter())
            {
                ShowWindow(browse_.filter(), SW_HIDE);
            }
            ShowWindow(filter_clear_btn_, SW_HIDE);
        }
        else if (!show_tabs && show_filter)
        {
            int available = std::max(0, tabs_width);
            int filter_width = ClampValue(available, filter_min_width, filter_max_width);
            int filter_y = y + std::max(0, (tabs_height - filter_height) / 2);
            int filter_x = padding + std::max(0, tabs_width - filter_width);
            int edit_width = std::max(filter_min_width / 2, filter_width - address_btn_width);
            place(browse_.filter(), filter_x, filter_y, edit_width, filter_height);
            place(filter_clear_btn_, filter_x + edit_width, filter_y, address_btn_width, filter_height);
            win32ui::SetEditPadding(browse_.filter(), 6, 2);
            ShowWindow(browse_.filter(), SW_SHOW);
            ShowWindow(filter_clear_btn_, SW_SHOW);
        }
        y += rows_height + gap;
    }
    else
    {
        if (tab_)
        {
            ShowWindow(tab_, SW_HIDE);
        }
        if (browse_.filter())
        {
            ShowWindow(browse_.filter(), SW_HIDE);
        }
        ShowWindow(filter_clear_btn_, SW_HIDE);
    }

    const int status_top = height - status_height;
    if (settings_.show_status_bar && status_bar_)
    {
        place(status_bar_, 0, status_top, width, status_height);
        SendMessageW(status_bar_, WM_SIZE, 0, 0);
    }

    content_rect_ = {0, y, width, status_top};
    LayoutContent(dragging_splitter);

    UpdateStatus();
    if (!dragging_splitter)
    {
        RedrawWindow(hwnd_, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_ERASE | RDW_UPDATENOW);
    }
}

}
