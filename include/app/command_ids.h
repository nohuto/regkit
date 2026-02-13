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

#pragma once

namespace regkit::cmd {

constexpr int kFileExit = 2000;
constexpr int kFileImport = 2001;
constexpr int kFileExport = 2002;
constexpr int kFileLoadHive = 2003;
constexpr int kFileUnloadHive = 2004;
constexpr int kFileClearTabsOnExit = 2005;
constexpr int kFileClearHistoryOnExit = 2006;
constexpr int kFileSaveOfflineHive = 2007;
constexpr int kFileImportComments = 2008;
constexpr int kFileExportComments = 2009;
constexpr int kFileSave = 2017;

constexpr int kNewKey = 2010;
constexpr int kNewString = 2011;
constexpr int kNewBinary = 2012;
constexpr int kNewDword = 2013;
constexpr int kNewQword = 2014;
constexpr int kNewMultiString = 2015;
constexpr int kNewExpandString = 2016;

constexpr int kEditModify = 2100;
constexpr int kEditModifyBinary = 2101;
constexpr int kEditModifyComment = 2115;
constexpr int kEditRename = 2102;
constexpr int kEditDelete = 2103;
constexpr int kEditCopyKey = 2104;
constexpr int kEditFind = 2105;
constexpr int kEditCopy = 2107;
constexpr int kEditPaste = 2108;
constexpr int kEditReplace = 2109;
constexpr int kEditUndo = 2110;
constexpr int kEditRedo = 2111;
constexpr int kEditCopyKeyPath = 2112;
constexpr int kEditGoTo = 2113;
constexpr int kEditPermissions = 2114;
constexpr int kEditCopyKeyPathAbbrev = 2130;
constexpr int kEditCopyKeyPathRegedit = 2131;
constexpr int kEditCopyKeyPathRegFile = 2132;
constexpr int kEditCopyKeyPathPowerShell = 2133;
constexpr int kEditCopyKeyPathPowerShellProvider = 2134;
constexpr int kEditCopyKeyPathEscaped = 2135;
constexpr int kEditCopyValueName = 2136;
constexpr int kEditCopyValueData = 2137;
constexpr int kEditInvertSelection = 2138;

constexpr int kRegistryLocal = 2120;
constexpr int kRegistryNetwork = 2121;
constexpr int kRegistryOffline = 2122;

constexpr int kNavBack = 2150;
constexpr int kNavForward = 2151;
constexpr int kNavUp = 2152;
constexpr int kTreeToggleExpand = 2153;

constexpr int kViewRefresh = 2200;
constexpr int kViewToolbar = 2201;
constexpr int kViewKeyTree = 2202;
constexpr int kViewHistory = 2203;
constexpr int kViewFont = 2204;
constexpr int kViewStatusBar = 2205;
constexpr int kViewSelectAll = 2206;
constexpr int kViewKeysInList = 2207;
constexpr int kViewExtraHives = 2208;
constexpr int kViewAddressBar = 2209;
constexpr int kViewSaveTreeState = 2210;
constexpr int kViewFilterBar = 2211;
constexpr int kViewTabControl = 2212;
constexpr int kViewSimulatedKeys = 2213;

constexpr int kFavoritesAdd = 2300;
constexpr int kFavoritesRemove = 2301;
constexpr int kFavoritesEdit = 2302;
constexpr int kFavoritesExport = 2303;
constexpr int kFavoritesImport = 2304;
constexpr int kFavoritesImportRegedit = 2305;
constexpr int kFavoritesItemBase = 2600;
constexpr int kFavoritesItemMax = 2699;

constexpr int kWindowNew = 2400;
constexpr int kWindowClose = 2401;
constexpr int kWindowAlwaysOnTop = 2402;

constexpr int kOptionsThemeSystem = 2450;
constexpr int kOptionsThemeLight = 2451;
constexpr int kOptionsThemeDark = 2452;
constexpr int kOptionsReplaceRegedit = 2453;
constexpr int kOptionsSingleInstance = 2454;
constexpr int kOptionsHiveFileDir = 2455;
constexpr int kOptionsRestartAdmin = 2456;
constexpr int kOptionsAlwaysRunAdmin = 2457;
constexpr int kCreateSimulatedKey = 2458;
constexpr int kOptionsSaveTabs = 2459;
constexpr int kOptionsReadOnly = 2460;
constexpr int kOptionsThemeCustom = 2461;
constexpr int kOptionsCompareRegistries = 2462;
constexpr int kOptionsRestartSystem = 2463;
constexpr int kOptionsAlwaysRunSystem = 2464;
constexpr int kOptionsOpenDefaultRegedit = 2465;
constexpr int kOptionsRestartTrustedInstaller = 2466;
constexpr int kOptionsAlwaysRunTrustedInstaller = 2467;
constexpr int kOptionsThemePresets = 2468;
constexpr int kOptionsIconSetDefault = 2469;
constexpr int kOptionsIconSetFluentUi = 2470;
constexpr int kOptionsIconSetMaterialSymbols = 2472;
constexpr int kOptionsIconSetCustom = 2473;
constexpr int kOptionsIconSetTabler = 2474;

constexpr int kHelpAbout = 2500;
constexpr int kHelpContents = 2501;

constexpr int kTraceLoad23H2 = 2700;
constexpr int kTraceLoad24H2 = 2701;
constexpr int kTraceLoad25H2 = 2702;
constexpr int kTraceLoadCustom = 2703;
constexpr int kTraceClear = 2704;
constexpr int kTraceGuide = 2705;
constexpr int kTraceEditRecent = 2706;
constexpr int kTraceEditActive = 2707;
constexpr int kTraceRecentBase = 2710;
constexpr int kTraceRecentMax = 2719;

constexpr int kDefaultLoadCustom = 2720;
constexpr int kDefaultClear = 2721;
constexpr int kDefaultEditRecent = 2722;
constexpr int kDefaultEditActive = 2723;
constexpr int kDefaultRecentBase = 2730;
constexpr int kDefaultRecentMax = 2739;
constexpr int kDefaultBundledBase = 2740;
constexpr int kDefaultBundledMax = 2759;

constexpr int kResearchRecordsTable = 2799;
constexpr int kResearchDxgKernel = 2800;
constexpr int kResearchSessionManager = 2801;
constexpr int kResearchPower = 2802;
constexpr int kResearchDwm = 2803;
constexpr int kResearchUsb = 2804;
constexpr int kResearchBcd = 2805;
constexpr int kResearchIntelNic = 2806;
constexpr int kResearchMmcss = 2807;
constexpr int kResearchStorNvme = 2808;
constexpr int kResearchMisc = 2809;

constexpr int kHeaderSizeToFit = 4000;
constexpr int kHeaderSizeAll = 4001;
constexpr int kHeaderToggleBase = 4100;

} // namespace regkit::cmd

namespace regkit::rowkind {

constexpr LPARAM kKey = 1;
constexpr LPARAM kValue = 2;

} // namespace regkit::rowkind
