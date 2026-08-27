// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

namespace regkit::frame {

enum class CommandArea {
  kUnknown,
  kDynamic,
  kFile,
  kView,
  kTraceDefaults,
  kWorkspaceAppearance,
  kNavigateClipboard,
  kMutation,
};

using CommandHandler = bool (*)(void* context, int command_id);

struct CommandContext {
  void* context = nullptr;
  CommandHandler dynamic = nullptr;
  CommandHandler file = nullptr;
  CommandHandler view = nullptr;
  CommandHandler trace_defaults = nullptr;
  CommandHandler workspace_appearance = nullptr;
  CommandHandler navigate_clipboard = nullptr;
  CommandHandler mutation = nullptr;
};

CommandArea ClassifyCommand(int command_id) noexcept;
bool DispatchCommand(int command_id, const CommandContext& context);

} // namespace regkit::frame
