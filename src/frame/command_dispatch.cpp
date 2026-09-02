// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/command_dispatch.h"

#include "frame/command_ids.h"

namespace regkit::frame {

CommandArea ClassifyCommand(int command_id) noexcept {
  if ((command_id >= cmd::kFavoritesItemBase &&
       command_id <= cmd::kFavoritesItemMax) ||
      (command_id >= cmd::kTraceRecentBase &&
       command_id <= cmd::kTraceRecentMax) ||
      (command_id >= cmd::kDefaultRecentBase &&
       command_id <= cmd::kDefaultRecentMax) ||
      (command_id >= cmd::kDefaultBundledBase &&
       command_id <= cmd::kDefaultBundledMax)) {
    return CommandArea::kDynamic;
  }
  if (command_id >= cmd::kResearchItemBase &&
      command_id <= cmd::kResearchItemMax) {
    return CommandArea::kWorkspaceAppearance;
  }
  if ((command_id >= cmd::kFileExit &&
       command_id <= cmd::kFileExportComments) ||
      command_id == cmd::kFileSave ||
      command_id == cmd::kFileOpenRegFile) {
    return CommandArea::kFile;
  }
  if (command_id >= cmd::kNewKey &&
      command_id <= cmd::kNewExpandString) {
    return CommandArea::kMutation;
  }
  switch (command_id) {
  case cmd::kTraceEditRecent:
  case cmd::kDefaultEditRecent:
    return CommandArea::kTraceDefaults;
  case cmd::kEditModify:
  case cmd::kEditModifyBinary:
  case cmd::kEditModifyComment:
  case cmd::kEditRename:
  case cmd::kEditDelete:
  case cmd::kCreateSimulatedKey:
    return CommandArea::kMutation;
  case cmd::kEditInvertSelection:
  case cmd::kTreeToggleExpand:
  case cmd::kTreeExpandAll:
  case cmd::kOptionsSaveTabs:
  case cmd::kOptionsReadOnly:
  case cmd::kOptionsCompareRegistries:
    return CommandArea::kView;
  default:
    break;
  }
  if ((command_id >= cmd::kEditCopyKey &&
       command_id <= cmd::kEditCopyValueData) ||
      (command_id >= cmd::kRegistryLocal &&
       command_id <= cmd::kRegistryOffline) ||
      (command_id >= cmd::kNavBack && command_id <= cmd::kTreeExpandAll)) {
    return CommandArea::kNavigateClipboard;
  }
  if (command_id >= cmd::kViewRefresh &&
      command_id <= cmd::kViewSimulatedKeys) {
    return CommandArea::kView;
  }
  if ((command_id >= cmd::kTraceLoad23H2 &&
       command_id <= cmd::kTraceEditActive) ||
      (command_id >= cmd::kDefaultLoadCustom &&
       command_id <= cmd::kDefaultEditActive)) {
    return CommandArea::kTraceDefaults;
  }
  if ((command_id >= cmd::kFavoritesAdd &&
       command_id <= cmd::kFavoritesImportRegedit) ||
      (command_id >= cmd::kWindowNew &&
       command_id <= cmd::kWindowAlwaysOnTop) ||
      (command_id >= cmd::kOptionsThemeSystem &&
       command_id <= cmd::kHistoryRevert) ||
      (command_id >= cmd::kHelpAbout &&
       command_id <= cmd::kHelpContents)) {
    return CommandArea::kWorkspaceAppearance;
  }
  return CommandArea::kUnknown;
}

bool DispatchCommand(int command_id, const CommandContext& context) {
  CommandHandler handler = nullptr;
  switch (ClassifyCommand(command_id)) {
  case CommandArea::kDynamic:
    handler = context.dynamic;
    break;
  case CommandArea::kFile:
    handler = context.file;
    break;
  case CommandArea::kView:
    handler = context.view;
    break;
  case CommandArea::kTraceDefaults:
    handler = context.trace_defaults;
    break;
  case CommandArea::kWorkspaceAppearance:
    handler = context.workspace_appearance;
    break;
  case CommandArea::kNavigateClipboard:
    handler = context.navigate_clipboard;
    break;
  case CommandArea::kMutation:
    handler = context.mutation;
    break;
  case CommandArea::kUnknown:
    break;
  }
  return handler && handler(context.context, command_id);
}

} // namespace regkit::frame
