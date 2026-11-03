// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/commands/command_detail.h"
#include "frame/window_impl.h"

#include "dialogs/bitfield_definition.h"

#include "frame/tools/research_links.h"
#include "frame/window/shortcut_bindings.h"
#include "win32/shell_integration.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/system/process_token.h>

#include <filesystem>
#include <optional>

namespace regkit
{
using namespace command_detail;

namespace
{

struct CacheAvailability
{
    bool any = false;
    bool tabs = false;
    bool history = false;
    bool search_history = false;
    bool dialog_state = false;
    bool tree_state = false;
    bool temporary = false;
};

bool HasCachePattern(const wchar_t* name, const wchar_t* prefix, const wchar_t* suffix)
{
    if (!name || !prefix || !suffix)
    {
        return false;
    }
    return wcslen(name) >= wcslen(prefix) + wcslen(suffix) && util::StartsWithInsensitive(name, prefix) &&
           util::EndsWithInsensitive(name, suffix);
}

CacheAvailability InspectCacheFiles(const std::wstring& folder)
{
    CacheAvailability available;
    if (folder.empty())
    {
        return available;
    }
    util::ForEachFile(util::JoinPath(folder, L"*").c_str(), [&](const WIN32_FIND_DATAW& data) {
        if (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            return true;
        }
        available.any = true;
        const wchar_t* name = data.cFileName;
        available.tabs = available.tabs || util::EqualsInsensitive(name, L"tabs.ini") ||
                         util::EqualsInsensitive(name, L"session.ini") || HasCachePattern(name, L"search_", L".tsv") ||
                         HasCachePattern(name, L"compare_", L".tsv");
        available.history = available.history || util::EqualsInsensitive(name, L"history.tsv");
        available.search_history = available.search_history || util::EqualsInsensitive(name, L"search_history.txt");
        available.dialog_state = available.dialog_state || util::EqualsInsensitive(name, L"dialog_state.ini");
        available.tree_state = available.tree_state || util::EqualsInsensitive(name, L"tree_state.ini");
        available.temporary = available.temporary || HasCachePattern(name, L"export_", L".reg");
        return true;
    });
    return available;
}

constexpr int kSimpleModeHidden[] = {
    cmd::kFileImportComments,
    cmd::kFileExportComments,
    cmd::kFileSaveOfflineHive,
    cmd::kNewVolatileKey,
    cmd::kNewSymbolicLink,
    cmd::kEditChangeType,
    cmd::kEditBits,
    cmd::kEditDecodeValue,
    cmd::kEditModifyComment,
    cmd::kEditResetDefault,
    cmd::kEditCopyKeyPathAbbrev,
    cmd::kEditCopyKeyPathRegEdit,
    cmd::kEditCopyKeyPathRegFile,
    cmd::kEditCopyKeyPathPowerShell,
    cmd::kEditCopyKeyPathPowerShellProvider,
    cmd::kEditCopyKeyPathEscaped,
    cmd::kEditCopyKeyPathNative,
    cmd::kEditCopyKeyPathNativeResolved,
    cmd::kEditInvertSelection,
    cmd::kRegistryOtherView,
    cmd::kEditGoToVirtualStore,
    cmd::kEditGoToGlobalKey,
    cmd::kEditKeyInfo,
    cmd::kEditOpenSourceHive,
    cmd::kOptionsHiveFileDir,
    cmd::kTreeExpandAll,
    cmd::kViewAutoRefresh,
    cmd::kViewToolbar,
    cmd::kViewHistory,
    cmd::kViewKeysInList,
    cmd::kViewGridLines,
    cmd::kViewSimulatedKeys,
    cmd::kViewExtraHives,
    cmd::kOptionsIconSetClassic,
    cmd::kOptionsIconSetPhosphor,
    cmd::kOptionsIconSetCustom,
    cmd::kFavoritesEdit,
    cmd::kFavoritesImport,
    cmd::kFavoritesImportRegEdit,
    cmd::kFavoritesExport,
};

// context menu items that follow selection like menu bar items
constexpr int kSelectionCommands[] = {
    cmd::kEditPermissions,
    cmd::kEditKeyInfo,
    cmd::kEditModifyComment,
    cmd::kEditModify,
    cmd::kEditModifyBinary,
    cmd::kEditCopyValueName,
    cmd::kEditChangeType,
    cmd::kEditDecodeValue,
    cmd::kEditBits,
    cmd::kEditCopyValueData,
    cmd::kFileLoadHive,
    cmd::kFileUnloadHive,
    cmd::kOptionsHiveFileDir,
    cmd::kNewVolatileKey,
    cmd::kEditCopyKeyPathNative,
    cmd::kEditCopyKeyPathNativeResolved,
};

}

std::wstring MainWindow::Impl::CommandTooltipText(int command_id) const
{
    switch (command_id)
    {
    case cmd::kRegistryLocal:
        return util::Tr(L"Local Registry");
    case cmd::kRegistryNetwork:
        return util::Tr(L"Remote Registry");
    case cmd::kRegistryOffline:
        return util::Tr(L"Offline Registry");
    case cmd::kEditFind:
        return util::Tr(L"Find");
    case cmd::kEditReplace:
        return util::Tr(L"Replace");
    case cmd::kFileSave:
        return util::Tr(L"Save");
    case cmd::kFileExport:
        return util::Tr(L"Export");
    case cmd::kEditUndo:
        return util::Tr(L"Undo");
    case cmd::kEditRedo:
        return util::Tr(L"Redo");
    case cmd::kEditCopy:
        return util::Tr(L"Copy");
    case cmd::kEditPaste:
        return util::Tr(L"Paste");
    case cmd::kEditDelete:
        return util::Tr(L"Delete");
    case cmd::kViewRefresh:
        return util::Tr(L"Refresh");
    case cmd::kNavBack:
        return util::Tr(L"Back");
    case cmd::kNavForward:
        return util::Tr(L"Forward");
    case cmd::kNavUp:
        return util::Tr(L"Up");
    default:
        return L"";
    }
}

bool MainWindow::Impl::EnsureWritable()
{
    if (!settings_.read_only)
    {
        return true;
    }
    ui::ShowWarning(hwnd_, util::Tr(L"Read only mode is enabled."));
    return false;
}

bool MainWindow::Impl::ShowsRegFile() const
{
    return std::any_of(session_->roots.begin(), session_->roots.end(), [](const RegistryRootEntry& root) { return RegistryStore::IsVirtualRoot(root.root); });
}

void MainWindow::Impl::FillBitfieldMenu(HMENU menu)
{
    if (GetMenuItemCount(menu) != 1 || GetMenuItemID(menu, 0) != cmd::kToolsBitfieldDefinitions)
    {
        return;
    }
    const std::vector<editors::bitfield::DefinitionFile>& files = editors::bitfield::BundledFiles();
    if (!files.empty())
    {
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }
    for (size_t i = 0; i < files.size() && i <= static_cast<size_t>(cmd::kToolsBitfieldFileMax - cmd::kToolsBitfieldFileBase); ++i)
    {
        std::wstring label = files[i].name.empty() ? util::FileBaseName(files[i].path) : files[i].name;
        label.append(L"   (").append(std::to_wstring(files[i].definitions.size())).append(L")");
        AppendMenuW(menu, MF_STRING, cmd::kToolsBitfieldFileBase + i, MenuName(label).c_str());
    }
}

void MainWindow::Impl::BuildMenus()
{
    if (deferred_startup_complete_ && !bundled_defaults_loaded_)
    {
        RefreshBundledDefaultsCache();
    }
    HMENU menu = CreateMenu();
    const auto append_menu = frame::AppendCommand;
    auto append_popup = [&](HMENU target, HMENU popup, const wchar_t* text) {
        AppendMenuW(target, MF_POPUP, reinterpret_cast<UINT_PTR>(popup), text);
    };
    auto separator = [](HMENU target) { AppendMenuW(target, MF_SEPARATOR, 0, nullptr); };

    HMENU file_menu = CreatePopupMenu();
    append_menu(file_menu, MF_STRING, cmd::kFileOpenRegFile, util::Tr(L"Open .reg &File..."));
    append_menu(file_menu, MF_STRING, cmd::kFileSave, util::Tr(L"&Save"));
    separator(file_menu);
    append_menu(file_menu, MF_STRING, cmd::kFileImport, util::Tr(L"&Import..."));
    append_menu(file_menu, MF_STRING, cmd::kFileExport, util::Tr(L"&Export..."));
    separator(file_menu);
    append_menu(file_menu, MF_STRING, cmd::kFileImportComments, util::Tr(L"I&mport Comments..."));
    append_menu(file_menu, MF_STRING, cmd::kFileExportComments, util::Tr(L"Export &Comments..."));
    separator(file_menu);
    frame::AppendRadioCommand(file_menu, cmd::kRegistryLocal, util::Tr(L"&Local Registry"));
    if (win32::HasAlternateView())
    {
        frame::AppendRadioCommand(file_menu, cmd::kRegistryLocal32, util::Tr(L"Local Registry (&32-bit)"));
    }
    frame::AppendRadioCommand(file_menu, cmd::kRegistryNetwork, util::Tr(L"&Remote Registry..."));
    frame::AppendRadioCommand(file_menu, cmd::kRegistryOffline, util::Tr(L"&Offline Registry..."));
    append_menu(file_menu, MF_STRING, cmd::kFileSaveOfflineHive, util::Tr(L"S&ave Offline Hive..."));
    separator(file_menu);
    append_menu(file_menu, MF_STRING, cmd::kFileLoadHive, util::Tr(L"Load &Hive..."));
    append_menu(file_menu, MF_STRING, cmd::kFileUnloadHive, util::Tr(L"&Unload Hive..."));
    separator(file_menu);
    append_menu(file_menu, MF_STRING, cmd::kFileExit, util::Tr(L"E&xit"));
    append_popup(menu, file_menu, util::Tr(L"&File"));

    HMENU edit_menu = CreatePopupMenu();
    append_menu(edit_menu, MF_STRING, cmd::kEditUndo, util::Tr(L"&Undo"));
    append_menu(edit_menu, MF_STRING, cmd::kEditRedo, util::Tr(L"&Redo"));
    separator(edit_menu);
    append_popup(edit_menu, BuildNewMenu(), util::Tr(L"Ne&w"));
    separator(edit_menu);
    append_menu(edit_menu, MF_STRING, cmd::kEditModify, util::Tr(L"&Modify..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditModifyBinary, util::Tr(L"Modify &Binary Data..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditChangeType, util::Tr(L"C&hange Data Type..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditBits, util::Tr(L"Edit Bit&s..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditDecodeValue, util::Tr(L"Decode &Value..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditModifyComment, util::Tr(L"Modif&y Comment..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditResetDefault, util::Tr(L"Reset &to Default"));
    separator(edit_menu);
    append_menu(edit_menu, MF_STRING, cmd::kEditCopy, util::Tr(L"&Copy"));
    HMENU copy_other_menu = CreatePopupMenu();
    append_menu(copy_other_menu, MF_STRING, cmd::kEditCopyKey, util::Tr(L"Key &Name"));
    append_menu(copy_other_menu, MF_STRING, cmd::kEditCopyKeyPath, util::Tr(L"Key &Path"));
    append_popup(copy_other_menu, BuildCopyKeyPathMenu(), util::Tr(L"&Key Path As"));
    separator(copy_other_menu);
    append_menu(copy_other_menu, MF_STRING, cmd::kEditCopyValueName, util::Tr(L"&Value Name"));
    append_menu(copy_other_menu, MF_STRING, cmd::kEditCopyValueData, util::Tr(L"Value &Data"));
    append_popup(edit_menu, copy_other_menu, util::Tr(L"Copy &Other"));
    append_menu(edit_menu, MF_STRING, cmd::kEditPaste, util::Tr(L"&Paste"));
    append_menu(edit_menu, MF_STRING, cmd::kEditRename, util::Tr(L"Re&name"));
    append_menu(edit_menu, MF_STRING, cmd::kEditDelete, util::Tr(L"&Delete"));
    separator(edit_menu);
    append_menu(edit_menu, MF_STRING, cmd::kViewSelectAll, util::Tr(L"Select &All"));
    append_menu(edit_menu, MF_STRING, cmd::kEditInvertSelection, util::Tr(L"Inv&ert Selection"));
    separator(edit_menu);
    append_menu(edit_menu, MF_STRING, cmd::kEditFind, util::Tr(L"&Find..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditReplace, util::Tr(L"Rep&lace..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditGoTo, util::Tr(L"&Go To..."));
    if (win32::HasAlternateView())
    {
        append_menu(edit_menu, MF_STRING, cmd::kRegistryOtherView, util::Tr(L"&Go to Other View"));
    }
    append_menu(edit_menu, MF_STRING, cmd::kEditGoToVirtualStore, util::Tr(L"&Go to Virtual Store Copy"));
    append_menu(edit_menu, MF_STRING, cmd::kEditGoToGlobalKey, util::Tr(L"&Go to Global Key"));
    separator(edit_menu);
    append_menu(edit_menu, MF_STRING, cmd::kEditPermissions, util::Tr(L"Perm&issions..."));
    append_menu(edit_menu, MF_STRING, cmd::kEditKeyInfo, util::Tr(L"&Key Information..."));
    append_popup(menu, edit_menu, util::Tr(L"&Edit"));

    HMENU view_menu = CreatePopupMenu();
    append_menu(view_menu, MF_STRING, cmd::kViewRefresh, util::Tr(L"&Refresh"));
    append_menu(view_menu, MF_STRING, cmd::kViewAutoRefresh, util::Tr(L"&Auto Refresh"));
    separator(view_menu);
    AppendMenuW(view_menu, MF_STRING, cmd::kViewToolbar, util::Tr(L"T&oolbar"));
    AppendMenuW(view_menu, MF_STRING, cmd::kViewAddressBar, util::Tr(L"A&ddress Bar"));
    AppendMenuW(view_menu, MF_STRING, cmd::kViewFilterBar, util::Tr(L"Filter &Bar"));
    AppendMenuW(view_menu, MF_STRING, cmd::kViewTabControl, util::Tr(L"&Tabs"));
    AppendMenuW(view_menu, MF_STRING, cmd::kViewStatusBar, util::Tr(L"&Status Bar"));
    separator(view_menu);
    AppendMenuW(view_menu, MF_STRING, cmd::kViewKeyTree, util::Tr(L"&Key Tree"));
    append_menu(view_menu, MF_STRING, cmd::kViewHistory, util::Tr(L"&History"));
    separator(view_menu);
    AppendMenuW(view_menu, MF_STRING, cmd::kViewKeysInList, util::Tr(L"Keys &in List"));
    AppendMenuW(view_menu, MF_STRING, cmd::kViewGridLines, util::Tr(L"&Grid Lines"));
    AppendMenuW(view_menu, MF_STRING, cmd::kViewSimulatedKeys, util::Tr(L"Sim&ulated Keys"));
    AppendMenuW(view_menu, MF_STRING, cmd::kViewExtraHives, util::Tr(L"Sho&w Extra Root Keys"));
    separator(view_menu);

    HMENU options_menu = CreatePopupMenu();
    // simple mode hides Options and moves appearance settings to View
    HMENU look_menu = settings_.simple_mode ? view_menu : options_menu;
    HMENU theme_menu = CreatePopupMenu();
    frame::AppendRadioCommand(theme_menu, cmd::kOptionsThemeSystem, util::Tr(L"&System"));
    frame::AppendRadioCommand(theme_menu, cmd::kOptionsThemeLight, util::Tr(L"&Light"));
    frame::AppendRadioCommand(theme_menu, cmd::kOptionsThemeDark, util::Tr(L"&Dark"));
    frame::AppendRadioCommand(theme_menu, cmd::kOptionsThemeCustom, util::Tr(L"C&ustom"));
    separator(theme_menu);
    AppendMenuW(theme_menu, MF_STRING, cmd::kOptionsThemePresets, util::Tr(L"&Theme Presets..."));
    append_popup(look_menu, theme_menu, util::Tr(L"Th&eme"));
    HMENU icon_menu = CreatePopupMenu();
    frame::AppendRadioCommand(icon_menu, cmd::kOptionsIconSetClassic, util::Tr(L"&Classic"));
    frame::AppendRadioCommand(icon_menu, cmd::kOptionsIconSetPhosphor, util::Tr(L"&Phosphor"));
    frame::AppendRadioCommand(icon_menu, cmd::kOptionsIconSetCustom, util::Tr(L"C&ustom"));
    append_popup(look_menu, icon_menu, util::Tr(L"Ico&ns"));
    HMENU language_menu = CreatePopupMenu();
    language_packs_ = util::InstalledLanguages();
    language_packs_.insert(language_packs_.begin(), {L"en", L"English"});
    frame::AppendRadioCommand(language_menu, cmd::kOptionsLanguageAuto, util::Tr(L"&Automatic"));
    separator(language_menu);
    for (size_t i = 0; i < language_packs_.size() && cmd::kOptionsLanguageBase + static_cast<int>(i) <= cmd::kOptionsLanguageMax; ++i)
    {
        frame::AppendRadioCommand(language_menu, cmd::kOptionsLanguageBase + static_cast<int>(i), MenuName(language_packs_[i].name).c_str());
    }
    append_popup(look_menu, language_menu, util::Tr(L"&Language"));
    AppendMenuW(look_menu, MF_STRING, cmd::kViewFont, util::Tr(L"&Font..."));
    separator(look_menu);
    AppendMenuW(view_menu, MF_STRING, cmd::kViewSimpleMode, util::Tr(L"Simple &Mode"));
    append_popup(menu, view_menu, util::Tr(L"&View"));
    AppendMenuW(options_menu, MF_STRING, cmd::kOptionsReadOnly, util::Tr(L"Read Only &Mode"));
    AppendMenuW(options_menu, MF_STRING, cmd::kOptionsBackupRestore, util::Tr(L"Use &Backup/Restore Privileges"));
    separator(options_menu);
    HMENU run_as_menu = CreatePopupMenu();
    append_menu(run_as_menu, MF_STRING, cmd::kFileRestart, util::Tr(L"&Restart"));
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsRestartUser, util::Tr(L"Restart as &User"));
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsRestartAdmin, util::Tr(L"R&estart as Admin"));
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsRestartSystem, util::Tr(L"Restart as &SYSTEM"));
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsRestartTrustedInstaller, util::Tr(L"Restart as &TrustedInstaller"));
    separator(run_as_menu);
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsAlwaysRunAdmin, util::Tr(L"A&lways Run as Admin"));
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsAlwaysRunSystem, util::Tr(L"&Always Run as SYSTEM"));
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsAlwaysRunTrustedInstaller, util::Tr(L"Al&ways Run as TrustedInstaller"));
    separator(run_as_menu);
    AppendMenuW(run_as_menu, MF_STRING, cmd::kOptionsHkcuFollowsUser, util::Tr(L"&HKCU Follows Signed-In User"));
    append_popup(options_menu, run_as_menu, util::Tr(L"R&un As"));
    separator(options_menu);
    AppendMenuW(options_menu, MF_STRING, cmd::kOptionsReplaceRegEdit, util::Tr(L"&Replace RegEdit"));
    AppendMenuW(options_menu, MF_STRING, cmd::kOptionsEditContextMenu, util::Tr(L"&Add \"Edit\" Context Menu"));
    AppendMenuW(options_menu, MF_STRING, cmd::kOptionsSingleInstance, util::Tr(L"Single &Instance"));
    AppendMenuW(options_menu, MF_STRING, cmd::kOptionsAutoComplete, util::Tr(L"Autocomplete &Key Paths"));
    HMENU save_tabs_menu = CreatePopupMenu();
    AppendMenuW(save_tabs_menu, MF_STRING, cmd::kOptionsSaveTabs, util::Tr(L"&All Tabs"));
    separator(save_tabs_menu);
    AppendMenuW(save_tabs_menu, MF_STRING, cmd::kOptionsSaveTabsLocal, util::Tr(L"&Local Registry Tabs"));
    AppendMenuW(save_tabs_menu, MF_STRING, cmd::kOptionsSaveTabsOffline, util::Tr(L"&Offline Hive Tabs"));
    AppendMenuW(save_tabs_menu, MF_STRING, cmd::kOptionsSaveTabsRemote, util::Tr(L"&Remote Registry Tabs"));
    separator(save_tabs_menu);
    AppendMenuW(save_tabs_menu, MF_STRING, cmd::kOptionsSaveTabsSearch, util::Tr(L"&Find Results"));
    AppendMenuW(save_tabs_menu, MF_STRING, cmd::kOptionsSaveTabsCompare, util::Tr(L"&Comparison Results"));
    AppendMenuW(save_tabs_menu, MF_STRING, cmd::kOptionsSaveTabsRegFile, util::Tr(L".reg File &Tabs"));
    separator(options_menu);
    append_popup(options_menu, save_tabs_menu, util::Tr(L"Save &Tabs"));
    AppendMenuW(options_menu, MF_STRING, cmd::kViewSaveTreeState, util::Tr(L"Save &Previous Tree State"));
    append_menu(options_menu, MF_STRING, cmd::kFileClearHistoryOnExit, util::Tr(L"Clear &History on Exit"));
    append_menu(options_menu, MF_STRING, cmd::kFileClearTabsOnExit, util::Tr(L"Clear Tabs &on Exit"));
    separator(options_menu);
    HMENU clear_cache_menu = CreatePopupMenu();
    AppendMenuW(clear_cache_menu, MF_STRING, cmd::kFileClearCacheAll, util::Tr(L"Clear &All"));
    separator(clear_cache_menu);
    AppendMenuW(clear_cache_menu, MF_STRING, cmd::kFileClearCacheTabs, util::Tr(L"Tab &Sessions"));
    AppendMenuW(clear_cache_menu, MF_STRING, cmd::kFileClearCacheHistory, util::Tr(L"&Change History"));
    AppendMenuW(clear_cache_menu, MF_STRING, cmd::kFileClearCacheSearchHistory, util::Tr(L"Search &History"));
    AppendMenuW(clear_cache_menu, MF_STRING, cmd::kFileClearCacheDialogState, util::Tr(L"&Dialog Options"));
    AppendMenuW(clear_cache_menu, MF_STRING, cmd::kFileClearCacheTreeState, util::Tr(L"&Tree State"));
    AppendMenuW(clear_cache_menu, MF_STRING, cmd::kFileClearCacheTemporary, util::Tr(L"Temporary &Files"));
    append_popup(options_menu, clear_cache_menu, util::Tr(L"&Clear Caches"));
    AppendMenuW(options_menu, MF_STRING, cmd::kOptionsResetSettings, util::Tr(L"Reset &Settings..."));
    HMENU favorites_menu = CreatePopupMenu();
    AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesAdd, util::Tr(L"&Add to Favorites..."));
    AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesRemove, util::Tr(L"&Remove Favorite"));
    AppendMenuW(favorites_menu, MF_STRING, cmd::kFavoritesEdit, util::Tr(L"Edit &Favorites..."));
    separator(favorites_menu);
    append_menu(favorites_menu, MF_STRING, cmd::kFavoritesImport, util::Tr(L"&Import Favorites..."));
    append_menu(favorites_menu, MF_STRING, cmd::kFavoritesImportRegEdit, util::Tr(L"I&mport RegEdit Favorites"));
    append_menu(favorites_menu, MF_STRING, cmd::kFavoritesExport, util::Tr(L"&Export Favorites..."));
    favorites_menu_ = favorites_menu;
    regedit_favorites_.clear();
    append_popup(menu, favorites_menu, util::Tr(L"F&avorites"));
    append_popup(menu, options_menu, util::Tr(L"&Options"));

    HMENU tools_menu = CreatePopupMenu();
    append_menu(tools_menu, MF_STRING, cmd::kOptionsCompareRegistries, util::Tr(L"&Compare Registries..."));
    append_menu(tools_menu, MF_STRING, cmd::kToolsConvertFile, util::Tr(L"Convert &File..."));
    append_menu(tools_menu, MF_STRING, cmd::kToolsKeyHandles, util::Tr(L"&Key Handles..."));
    HMENU bitfield_menu = CreatePopupMenu();
    AppendMenuW(bitfield_menu, MF_STRING, cmd::kToolsBitfieldDefinitions, util::Tr(L"&New Definition File..."));
    append_popup(tools_menu, bitfield_menu, util::Tr(L"&Bit Definitions"));
    append_popup(menu, tools_menu, util::Tr(L"&Tools"));

    HMENU window_menu = CreatePopupMenu();
    append_menu(window_menu, MF_STRING, cmd::kWindowNew, util::Tr(L"&New Window"));
    separator(window_menu);
    append_menu(window_menu, MF_STRING, cmd::kTabClose, util::Tr(L"&Close Tab"));
    AppendMenuW(window_menu, MF_STRING, cmd::kWindowClose, util::Tr(L"Close &Window"));
    separator(window_menu);
    AppendMenuW(window_menu, MF_STRING, cmd::kWindowAlwaysOnTop, util::Tr(L"&Always on Top"));
    append_popup(menu, window_menu, util::Tr(L"&Window"));

    HMENU trace_menu = CreatePopupMenu();
    append_menu(trace_menu, MF_STRING, cmd::kTraceLoad23H2, util::Tr(L"2&3H2"));
    append_menu(trace_menu, MF_STRING, cmd::kTraceLoad24H2, util::Tr(L"2&4H2"));
    append_menu(trace_menu, MF_STRING, cmd::kTraceLoad25H2, util::Tr(L"&25H2"));
    separator(trace_menu);
    append_menu(trace_menu, MF_STRING, cmd::kTraceLoadCustom, util::Tr(L"&Open Trace File..."));
    separator(trace_menu);
    append_menu(trace_menu, MF_STRING, cmd::kTraceEditActive, util::Tr(L"&Edit Active Traces..."));
    append_menu(trace_menu, MF_STRING, cmd::kTraceClear, util::Tr(L"Clear &Active Traces"));
    separator(trace_menu);
    append_menu(trace_menu, MF_STRING, cmd::kTraceEditRecent, util::Tr(L"Edit &Recent Traces..."));
    append_menu(trace_menu, MF_STRING, cmd::kTraceClearRecent, util::Tr(L"&Clear Recent Traces"));
    separator(trace_menu);
    append_menu(trace_menu, MF_STRING, cmd::kTraceGuide, util::Tr(L"&Guide"));

    HMENU default_menu = CreatePopupMenu();
    HMENU bundled_menu = nullptr;
    std::wstring bundled_group;
    for (size_t i = 0; i < bundled_defaults_.size(); ++i)
    {
        const auto& entry = bundled_defaults_[i];
        if (!bundled_menu || entry.group != bundled_group)
        {
            bundled_group = entry.group;
            bundled_menu = CreatePopupMenu();
            append_popup(default_menu, bundled_menu, MenuName(bundled_group).c_str());
        }
        append_menu(bundled_menu, MF_STRING, cmd::kDefaultBundledBase + static_cast<int>(i), MenuName(entry.label).c_str());
    }
    separator(default_menu);
    append_menu(default_menu, MF_STRING, cmd::kDefaultLoadCustom, util::Tr(L"&Open Default File..."));
    separator(default_menu);
    append_menu(default_menu, MF_STRING, cmd::kDefaultEditActive, util::Tr(L"Edit &Active Defaults..."));
    append_menu(default_menu, MF_STRING, cmd::kDefaultClear, util::Tr(L"&Clear Active Defaults"));
    separator(default_menu);
    append_menu(default_menu, MF_STRING, cmd::kDefaultEditRecent, util::Tr(L"Edit Recent &Defaults..."));
    append_menu(default_menu, MF_STRING, cmd::kDefaultClearRecent, util::Tr(L"Clear &Recent Defaults"));
    separator(default_menu);
    append_menu(default_menu, MF_STRING, cmd::kDefaultResetEnable, util::Tr(L"&Enable Context Menu (risky)"));

    HMENU help_menu = CreatePopupMenu();
    append_menu(help_menu, MF_STRING, cmd::kHelpContents, util::Tr(L"D&ocumentation"));
    AppendMenuW(help_menu, MF_STRING, cmd::kHelpDiscord, L"&Discord");
    separator(help_menu);
    AppendMenuW(help_menu, MF_STRING, cmd::kHelpCheckUpdates, util::Tr(L"Check &for Updates"));
    AppendMenuW(help_menu, MF_STRING, cmd::kHelpAutoCheckUpdates, util::Tr(L"&Check for Updates Automatically"));
    separator(help_menu);
    AppendMenuW(help_menu, MF_STRING, cmd::kHelpAbout, util::Tr(L"&About RegKit"));

    HMENU research_menu = CreatePopupMenu();
    int research_command = cmd::kResearchItemBase;
    for (const auto& link : frame::ResearchLinks())
    {
        AppendMenuW(research_menu, MF_STRING, research_command++, link.name);
        if (link.separator_after)
        {
            separator(research_menu);
        }
    }

    append_popup(menu, trace_menu, util::Tr(L"T&race"));
    append_popup(menu, default_menu, util::Tr(L"Defa&ult"));
    append_popup(menu, research_menu, util::Tr(L"Re&search"));
    append_popup(menu, help_menu, util::Tr(L"&Help"));
    if (settings_.simple_mode)
    {
        for (int command : {cmd::kOptionsReadOnly, cmd::kToolsKeyHandles, cmd::kWindowNew, cmd::kTraceGuide, cmd::kDefaultLoadCustom, cmd::kResearchItemBase})
        {
            const int position = win32ui::SubMenuWithCommand(menu, static_cast<UINT>(command));
            if (position >= 0)
            {
                DeleteMenu(menu, static_cast<UINT>(position), MF_BYPOSITION);
            }
        }
        ApplySimpleMode(menu);
    }
    trace_menu_ = win32ui::MenuContains(menu, trace_menu) ? trace_menu : nullptr;
    default_menu_ = win32ui::MenuContains(menu, default_menu) ? default_menu : nullptr;

    menu_bar_.Attach(hwnd_, menu);
    menu_bar_.ApplyTheme();
}

void MainWindow::Impl::ApplySimpleMode(HMENU menu) const
{
    if (settings_.simple_mode)
    {
        win32ui::RemoveMenuCommands(menu, kSimpleModeHidden);
    }
}

struct MainWindow::Impl::CommandState
{
    Impl* window = nullptr;
    // one lookup per popup, shared by all items
    std::optional<CacheAvailability> cache;
    std::optional<bool> settings_changed;
    std::optional<bool> comment_targets;
    std::optional<std::vector<DefaultValueChoice>> reset_choices;
    bool row_known = false;
    const ListRow* row = nullptr;
};

bool MainWindow::Impl::CommandEnabled(int command, CommandState* state)
{
    const RegistryNode* node = browse_.current_node();
    const bool writable = !settings_.read_only;
    const RegistryMode mode = session_->mode;
    const HWND values = browse_.values().hwnd();
    auto selected_value = [&]() -> const ListRow* {
        if (!state->row_known)
        {
            const int index = values && ListView_GetSelectedCount(values) == 1 ? ListView_GetNextItem(values, -1, LVNI_SELECTED) : -1;
            const ListRow* row = index >= 0 ? browse_.values().RowAt(index) : nullptr;
            state->row = row && row->kind == rowkind::kValue && !row->simulated ? row : nullptr;
            state->row_known = true;
        }
        return state->row;
    };
    auto focused_value = [&]() { return GetFocus() == values ? selected_value() : nullptr; };
    auto cache = [&]() -> const CacheAvailability& {
        if (!state->cache)
        {
            state->cache = InspectCacheFiles(util::GetCacheFolder());
        }
        return *state->cache;
    };
    auto any_path = [](const std::vector<std::wstring>& paths, int limit) {
        return std::any_of(paths.begin(), paths.begin() + std::min(static_cast<ptrdiff_t>(paths.size()), static_cast<ptrdiff_t>(limit)), [](const std::wstring& path) { return !path.empty(); });
    };
    switch (command)
    {
    case cmd::kEditPermissions:
    case cmd::kEditKeyInfo:
        return node != nullptr;
    case cmd::kEditModifyComment:
        if (!state->comment_targets)
        {
            state->comment_targets = !CommentTargets().empty();
        }
        return *state->comment_targets;
    case cmd::kEditModify:
    case cmd::kEditModifyBinary:
    case cmd::kEditCopyValueName:
        return selected_value() != nullptr;
    case cmd::kEditChangeType:
    case cmd::kEditDecodeValue:
    case cmd::kEditBits:
    case cmd::kEditCopyValueData:
        return selected_value() && !selected_value()->missing;
    case cmd::kEditResetDefault:
        if (!state->reset_choices)
        {
            state->reset_choices = SelectedValueDefaultChoices();
        }
        return !state->reset_choices->empty();
    case cmd::kFileLoadHive:
        return writable && mode != RegistryMode::kRemote;
    case cmd::kFileUnloadHive:
        return writable && mode != RegistryMode::kRemote && node && IsMountedHive(node->root, node->subkey);
    case cmd::kOptionsHiveFileDir:
        return !ResolveSelectedHiveFilePath().empty();
    // .reg files and hives have no volatile keys
    case cmd::kNewVolatileKey:
        return writable && mode != RegistryMode::kOffline && !ShowsRegFile();
    // only the local registry has native paths
    case cmd::kEditCopyKeyPathNative:
    case cmd::kEditCopyKeyPathNativeResolved:
        return mode == RegistryMode::kLocal && !ShowsRegFile();
    case cmd::kEditUndo:
        return writable && session_->undo.CanUndo();
    case cmd::kEditRedo:
        return writable && session_->undo.CanRedo();
    case cmd::kEditPaste:
    case cmd::kEditReplace:
    case cmd::kNewKey:
    case cmd::kNewString:
    case cmd::kNewExpandString:
    case cmd::kNewBinary:
    case cmd::kNewDword:
    case cmd::kNewQword:
    case cmd::kNewMultiString:
    case cmd::kNewSymbolicLink:
        return writable;
    // reg.exe imports into the local registry only
    case cmd::kFileImport:
        return writable && mode == RegistryMode::kLocal && !ShowsRegFile();
    case cmd::kEditRename:
        return writable && !(focused_value() && (focused_value()->extra.empty() || focused_value()->missing));
    case cmd::kEditDelete:
        return writable && !(focused_value() && focused_value()->missing);
    case cmd::kFileSave:
        {
            const int tab_index = tab_ ? TabCtrl_GetCurSel(tab_) : -1;
            const TabEntry* tab = tab_index >= 0 && static_cast<size_t>(tab_index) < tabs_.size() ? &tabs_[static_cast<size_t>(tab_index)] : nullptr;
            return writable && tab && (tab->kind == TabEntry::Kind::kRegFile ? tab->reg_file_dirty : tab->kind == TabEntry::Kind::kRegistry && tab->session && tab->session->offline_dirty);
        }
    case cmd::kRegistryOtherView:
        return mode == RegistryMode::kLocal && node != nullptr;
    case cmd::kEditGoToVirtualStore:
        return !VirtualStoreTarget().empty();
    case cmd::kEditGoToGlobalKey:
        return !GlobalKeyTarget().empty();
    case cmd::kFileSaveOfflineHive:
        return mode == RegistryMode::kOffline && !session_->offline_mount.empty();
    case cmd::kViewSimulatedKeys:
        return HasActiveTraces();
    case cmd::kViewExtraHives:
        return mode == RegistryMode::kLocal;
    case cmd::kWindowNew:
        return !settings_.single_instance;
    case cmd::kTraceClear:
        return !active_traces_.empty();
    case cmd::kDefaultClear:
        return !active_defaults_.empty();
    case cmd::kTraceEditRecent:
    case cmd::kTraceClearRecent:
        return any_path(recent_trace_paths_.items(), cmd::kTraceRecentMax - cmd::kTraceRecentBase + 1);
    case cmd::kDefaultEditRecent:
    case cmd::kDefaultClearRecent:
        return any_path(recent_default_paths_.items(), cmd::kDefaultRecentMax - cmd::kDefaultRecentBase + 1);
    case cmd::kFileClearCacheAll:
        return cache().any;
    case cmd::kFileClearCacheTabs:
        return cache().tabs;
    case cmd::kFileClearCacheHistory:
        return cache().history;
    case cmd::kFileClearCacheSearchHistory:
        return cache().search_history;
    case cmd::kFileClearCacheDialogState:
        return cache().dialog_state;
    case cmd::kFileClearCacheTreeState:
        return cache().tree_state;
    case cmd::kFileClearCacheTemporary:
        return cache().temporary;
    case cmd::kOptionsResetSettings:
        if (!state->settings_changed)
        {
            const workspace::Settings settings = CurrentSettings();
            state->settings_changed = workspace::SerializeSettings(settings) != workspace::SerializeSettings(workspace::DefaultOptions(settings));
        }
        return *state->settings_changed;
    case cmd::kOptionsBackupRestore:
        return util::IsProcessPrivileged();
    case cmd::kOptionsRestartUser:
        return (win32ui::IsProcessElevated() || win32ui::IsProcessSystem() || win32ui::IsProcessTrustedInstaller()) && util::IsUacEnabled();
    case cmd::kOptionsRestartAdmin:
        return !win32ui::IsProcessElevated() || win32ui::IsProcessSystem() || win32ui::IsProcessTrustedInstaller();
    case cmd::kOptionsRestartSystem:
        return !win32ui::IsProcessSystem();
    case cmd::kOptionsRestartTrustedInstaller:
        return !win32ui::IsProcessTrustedInstaller();
    case cmd::kOptionsReplaceRegEdit:
        return win32ui::IsProcessElevated() || win32ui::IsProcessSystem() || win32ui::IsProcessTrustedInstaller();
    case cmd::kOptionsEditContextMenu:
        return !win32ui::IsProcessSystem() && !win32ui::IsProcessTrustedInstaller();
    case cmd::kNavBack:
        return browse_.navigation().back;
    case cmd::kNavForward:
        return browse_.navigation().forward;
    case cmd::kNavUp:
        return browse_.navigation().up;
    default:
        return true;
    }
}

bool MainWindow::Impl::CommandChecked(int command) const
{
    const RegistryMode mode = session_->mode;
    auto trace_active = [&](auto match) { return std::any_of(active_traces_.begin(), active_traces_.end(), match); };
    auto label_active = [&](const wchar_t* label) { return trace_active([label](const auto& trace) { return util::EqualsInsensitive(trace.label, label); }); };
    auto default_active = [&](const std::wstring& path) {
        return std::any_of(active_defaults_.begin(), active_defaults_.end(), [&](const auto& defaults) { return EqualsInsensitive(defaults.source_path, path); });
    };
    for (const PaneToggle& toggle : kPaneToggles)
    {
        if (toggle.command == command)
        {
            return settings_.*toggle.shown;
        }
    }
    if (command >= cmd::kOptionsLanguageBase && command <= cmd::kOptionsLanguageMax)
    {
        const size_t index = static_cast<size_t>(command - cmd::kOptionsLanguageBase);
        return index < language_packs_.size() && settings_.language == language_packs_[index].code;
    }
    if (command >= cmd::kTraceRecentBase && command <= cmd::kTraceRecentMax)
    {
        const size_t index = static_cast<size_t>(command - cmd::kTraceRecentBase);
        return index < recent_trace_paths_.items().size() && trace_active([&](const auto& trace) { return EqualsInsensitive(trace.source_path, recent_trace_paths_.items()[index]); });
    }
    if (command >= cmd::kDefaultBundledBase && command <= cmd::kDefaultBundledMax)
    {
        const size_t index = static_cast<size_t>(command - cmd::kDefaultBundledBase);
        return index < bundled_defaults_.size() && default_active(bundled_defaults_[index].path);
    }
    if (command >= cmd::kDefaultRecentBase && command <= cmd::kDefaultRecentMax)
    {
        const size_t index = static_cast<size_t>(command - cmd::kDefaultRecentBase);
        return index < recent_default_paths_.items().size() && default_active(recent_default_paths_.items()[index]);
    }
    const int kinds = settings_.save_tab_kinds;
    switch (command)
    {
    case cmd::kViewGridLines:
        return settings_.show_value_grid;
    case cmd::kRegistryLocal:
        return mode == RegistryMode::kLocal && !session_->view;
    case cmd::kRegistryLocal32:
        return mode == RegistryMode::kLocal && session_->view;
    case cmd::kRegistryNetwork:
        return mode == RegistryMode::kRemote;
    case cmd::kRegistryOffline:
        return mode == RegistryMode::kOffline;
    case cmd::kViewAutoRefresh:
        return settings_.auto_refresh;
    case cmd::kViewKeysInList:
        return settings_.show_keys_in_list;
    case cmd::kViewSimulatedKeys:
        return settings_.show_simulated_keys;
    case cmd::kViewExtraHives:
        return settings_.show_extra_hives;
    case cmd::kViewSimpleMode:
        return settings_.simple_mode;
    case cmd::kOptionsThemeSystem:
        return theme_mode_ == win32ui::ThemeMode::system;
    case cmd::kOptionsThemeLight:
        return theme_mode_ == win32ui::ThemeMode::light;
    case cmd::kOptionsThemeDark:
        return theme_mode_ == win32ui::ThemeMode::dark;
    case cmd::kOptionsThemeCustom:
        return theme_mode_ == win32ui::ThemeMode::custom;
    case cmd::kOptionsIconSetClassic:
        return util::EqualsInsensitive(settings_.icon_set, kIconSetClassic);
    case cmd::kOptionsIconSetPhosphor:
        return util::EqualsInsensitive(settings_.icon_set, kIconSetPhosphor);
    case cmd::kOptionsIconSetCustom:
        return util::EqualsInsensitive(settings_.icon_set, kIconSetCustom);
    case cmd::kOptionsLanguageAuto:
        return settings_.language.empty();
    case cmd::kOptionsReadOnly:
        return settings_.read_only;
    case cmd::kOptionsBackupRestore:
        return backup_privileges_ != nullptr;
    case cmd::kOptionsAlwaysRunAdmin:
        return settings_.always_run_as_admin;
    case cmd::kOptionsAlwaysRunSystem:
        return settings_.always_run_as_system;
    case cmd::kOptionsAlwaysRunTrustedInstaller:
        return settings_.always_run_as_trustedinstaller;
    case cmd::kOptionsHkcuFollowsUser:
        return settings_.hkcu_follows_shell_user;
    case cmd::kOptionsReplaceRegEdit:
        return win32::IsRegEditReplacementRegistered(util::ModulePath());
    case cmd::kOptionsEditContextMenu:
        return win32::IsEditMenuRegistered(util::ModulePath());
    case cmd::kOptionsSingleInstance:
        return settings_.single_instance;
    case cmd::kOptionsAutoComplete:
        return settings_.autocomplete;
    case cmd::kOptionsSaveTabs:
        return (kinds & workspace::kSaveTabsAll) == workspace::kSaveTabsAll;
    case cmd::kOptionsSaveTabsLocal:
        return (kinds & workspace::kSaveTabsLocal) != 0;
    case cmd::kOptionsSaveTabsOffline:
        return (kinds & workspace::kSaveTabsOffline) != 0;
    case cmd::kOptionsSaveTabsRemote:
        return (kinds & workspace::kSaveTabsRemote) != 0;
    case cmd::kOptionsSaveTabsSearch:
        return (kinds & workspace::kSaveTabsSearch) != 0;
    case cmd::kOptionsSaveTabsCompare:
        return (kinds & workspace::kSaveTabsCompare) != 0;
    case cmd::kOptionsSaveTabsRegFile:
        return (kinds & workspace::kSaveTabsRegFile) != 0;
    case cmd::kViewSaveTreeState:
        return settings_.save_tree_state;
    case cmd::kFileClearHistoryOnExit:
        return settings_.clear_history_on_exit;
    case cmd::kFileClearTabsOnExit:
        return settings_.clear_tabs_on_exit;
    case cmd::kWindowAlwaysOnTop:
        return settings_.always_on_top;
    case cmd::kHelpAutoCheckUpdates:
        return settings_.auto_check_updates;
    case cmd::kTraceLoad23H2:
        return label_active(L"23H2");
    case cmd::kTraceLoad24H2:
        return label_active(L"24H2");
    case cmd::kTraceLoad25H2:
        return label_active(L"25H2");
    case cmd::kDefaultResetEnable:
        return settings_.default_reset_enabled;
    default:
        return false;
    }
}

void MainWindow::Impl::UpdateMenuState(HMENU menu)
{
    CommandState state{this};
    FillBitfieldMenu(menu);
    RefreshResetDefaultMenu(menu, &state);
    if (win32ui::MenuContains(GetMenu(hwnd_), menu))
    {
        FillMenuLists(menu);
        const win32ui::MenuHooks hooks = {
            &state,
            [](void* context, int command) {
                auto* state = static_cast<CommandState*>(context);
                return state->window->CommandEnabled(command, state);
            },
            [](void* context, int command) { return static_cast<CommandState*>(context)->window->CommandChecked(command); },
        };
        win32ui::UpdateMenuPopup(menu, hooks);
        const int save_tabs_position = win32ui::SubMenuWithCommand(menu, cmd::kOptionsSaveTabs);
        if (save_tabs_position >= 0)
        {
            CheckMenuItem(menu, static_cast<UINT>(save_tabs_position), MF_BYPOSITION | (settings_.save_tab_kinds != 0 ? MF_CHECKED : MF_UNCHECKED));
        }
        return;
    }
    // most context menu state stays fixed, these commands follow selection like the menu bar
    CheckMenuItem(menu, cmd::kViewGridLines, MF_BYCOMMAND | (settings_.show_value_grid ? MF_CHECKED : MF_UNCHECKED));
    for (const int command : kSelectionCommands)
    {
        if (GetMenuState(menu, static_cast<UINT>(command), MF_BYCOMMAND) != static_cast<UINT>(-1))
        {
            EnableMenuItem(menu, static_cast<UINT>(command), MF_BYCOMMAND | (CommandEnabled(command, &state) ? MF_ENABLED : MF_GRAYED));
        }
    }
}

void MainWindow::Impl::UpdateToolbarStates()
{
    if (!toolbar_)
    {
        return;
    }
    CommandState state{this};
    const int count = static_cast<int>(SendMessageW(toolbar_, TB_BUTTONCOUNT, 0, 0));
    for (int index = 0; index < count; ++index)
    {
        TBBUTTON button = {};
        if (SendMessageW(toolbar_, TB_GETBUTTON, static_cast<WPARAM>(index), reinterpret_cast<LPARAM>(&button)) && button.idCommand != 0)
        {
            win32ui::EnableToolbarButton(toolbar_, button.idCommand, CommandEnabled(button.idCommand, &state));
        }
    }
}

void MainWindow::Impl::RefreshFavoritesCache()
{
    favorites_cache_.clear();
    FavoritesStore::Load(&favorites_cache_);
    favorites_loaded_ = true;
}

void MainWindow::Impl::FillMenuLists(HMENU menu)
{
    if (menu == trace_menu_)
    {
        std::vector<std::wstring> labels;
        for (const std::wstring& path : recent_trace_paths_.items())
        {
            const std::wstring name = util::FileName(path);
            labels.push_back(path.empty() ? std::wstring() : name.empty() ? std::wstring(util::Tr(L"Trace"))
                                                                          : MenuName(name));
        }
        win32ui::ReplaceMenuRange(menu, cmd::kTraceRecentBase, cmd::kTraceRecentMax, cmd::kTraceLoadCustom, labels);
    }
    else if (menu == default_menu_)
    {
        std::vector<std::wstring> labels;
        for (const std::wstring& path : recent_default_paths_.items())
        {
            std::wstring name = util::FileBaseName(path);
            const std::wstring build = defaults::ShortLabel(std::wstring(), path);
            if (!build.empty())
            {
                name = build + L" - " + name;
            }
            labels.push_back(path.empty() ? std::wstring() : name.empty() ? std::wstring(util::Tr(L"Default"))
                                                                          : MenuName(name));
        }
        win32ui::ReplaceMenuRange(menu, cmd::kDefaultRecentBase, cmd::kDefaultRecentMax, cmd::kDefaultLoadCustom, labels);
    }
    else if (menu == favorites_menu_)
    {
        if (!favorites_loaded_)
        {
            RefreshFavoritesCache();
        }
        std::vector<std::wstring> paths;
        const size_t shown = std::min(favorites_cache_.size(), static_cast<size_t>(cmd::kFavoritesItemMax - cmd::kFavoritesItemBase + 1));
        for (size_t i = 0; i < shown; ++i)
        {
            paths.push_back(MenuName(favorites_cache_[i]));
        }
        win32ui::ReplaceMenuRange(menu, cmd::kFavoritesItemBase, cmd::kFavoritesItemMax, cmd::kRegEditFavoriteBase, paths);
        regedit_favorites_.clear();
        FavoritesStore::LoadRegEdit(&regedit_favorites_);
        regedit_favorites_.resize(std::min(regedit_favorites_.size(), static_cast<size_t>(cmd::kRegEditFavoriteMax - cmd::kRegEditFavoriteBase + 1)));
        std::vector<std::wstring> names;
        for (const auto& favorite : regedit_favorites_)
        {
            names.push_back(MenuName(favorite.name));
        }
        win32ui::ReplaceMenuRange(menu, cmd::kRegEditFavoriteBase, cmd::kRegEditFavoriteMax, 0, names);
    }
}

void MainWindow::Impl::RefreshBundledDefaultsCache()
{
    bundled_defaults_.clear();
    bundled_defaults_loaded_ = true;

    std::wstring module_dir = util::ModuleDirectory();
    if (module_dir.empty())
    {
        return;
    }
    std::wstring assets = util::JoinPath(module_dir, L"assets");
    std::wstring defaults_dir = util::JoinPath(assets, L"defaults");
    std::error_code error;
    const std::filesystem::path root(defaults_dir);
    for (const auto& folder :
         std::filesystem::directory_iterator(root, std::filesystem::directory_options::skip_permission_denied, error))
    {
        if (!folder.is_directory(error))
        {
            error.clear();
            continue;
        }
        for (const auto& file : std::filesystem::directory_iterator(
                 folder.path(),
                 std::filesystem::directory_options::skip_permission_denied,
                 error
             ))
        {
            if (!file.is_regular_file(error) || !util::EqualsInsensitive(file.path().extension().c_str(), L".reg"))
            {
                error.clear();
                continue;
            }
            BundledDefault entry;
            entry.group = folder.path().filename().wstring();
            entry.label = file.path().stem().wstring();
            entry.path = file.path().wstring();
            bundled_defaults_.push_back(std::move(entry));
        }
        error.clear();
    }

    std::sort(bundled_defaults_.begin(), bundled_defaults_.end(), [](const BundledDefault& left, const BundledDefault& right) {
        const int group = util::CompareInsensitive(left.group, right.group);
        return group != 0 ? group < 0 : util::CompareInsensitive(left.label, right.label) < 0;
    });
    size_t bundled_limit =
        std::min(bundled_defaults_.size(), static_cast<size_t>(cmd::kDefaultBundledMax - cmd::kDefaultBundledBase + 1));
    if (bundled_defaults_.size() > bundled_limit)
    {
        bundled_defaults_.resize(bundled_limit);
    }
}

bool MainWindow::Impl::HandleMenuCommand(int command_id)
{
    return HandleDynamicCommand(command_id) || HandleFileCommand(command_id) || HandleViewCommand(command_id) ||
           HandleTraceDefaultCommand(command_id) || HandleWorkspaceAppearanceCommand(command_id) ||
           HandleNavigateClipboardCommand(command_id) || HandleMutationCommand(command_id) || HandleToolsCommand(command_id);
}

std::vector<MainWindow::Impl::DefaultValueChoice> MainWindow::Impl::SelectedValueDefaultChoices() const
{
    if (!settings_.default_reset_enabled || settings_.read_only || active_defaults_.empty())
    {
        return {};
    }
    const std::vector<const ListRow*> rows = SelectedListRows(browse_.values());
    if (rows.size() != 1 || rows.front()->kind != rowkind::kValue || rows.front()->simulated)
    {
        return {};
    }
    return CollectDefaultChoices(rows.front()->extra);
}

HMENU MainWindow::Impl::BuildResetDefaultMenu(const std::vector<DefaultValueChoice>& choices) const
{
    HMENU menu = CreatePopupMenu();
    const int limit = cmd::kResetDefaultMax - cmd::kResetDefaultBase + 1;
    for (size_t i = 0; i < choices.size() && static_cast<int>(i) < limit; ++i)
    {
        std::wstring text = choices[i].label.empty() ? std::wstring(util::Tr(L"Default")) : choices[i].label;
        if (!choices[i].present)
        {
            text.append(L" ").append(util::Tr(L"(Missing)"));
        }
        AppendMenuW(menu, MF_STRING, cmd::kResetDefaultBase + static_cast<int>(i), MenuName(text).c_str());
    }
    return menu;
}

void MainWindow::Impl::RefreshResetDefaultMenu(HMENU menu, CommandState* state)
{
    const int count = GetMenuItemCount(menu);
    for (int position = 0; position < count; ++position)
    {
        MENUITEMINFOW item = {};
        item.cbSize = sizeof(item);
        item.fMask = MIIM_ID;
        if (!GetMenuItemInfoW(menu, position, TRUE, &item) || item.wID != static_cast<UINT>(cmd::kEditResetDefault))
        {
            continue;
        }
        if (!state->reset_choices)
        {
            state->reset_choices = SelectedValueDefaultChoices();
        }
        const std::vector<DefaultValueChoice>& choices = *state->reset_choices;
        item.fMask = MIIM_STATE | MIIM_SUBMENU;
        item.hSubMenu = choices.size() < 2 ? nullptr : BuildResetDefaultMenu(choices);
        item.fState = choices.empty() ? MFS_GRAYED : MFS_ENABLED;
        SetMenuItemInfoW(menu, position, TRUE, &item);
        return;
    }
}

}
