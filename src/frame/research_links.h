// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <span>

namespace regkit::frame {

struct ResearchLink {
  const wchar_t* name;
  const wchar_t* url;
  bool separator_after = false;
};

std::span<const ResearchLink> ResearchLinks() noexcept;
const ResearchLink* ResearchLinkForCommand(int command_id) noexcept;

} // namespace regkit::frame
