// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/research_links.h"

#include "frame/command_ids.h"

#include <array>
#include <cstddef>

namespace regkit::frame {
namespace {

constexpr std::array<ResearchLink, 22> kLinks = {{
    {L"Capture Table", L"https://noverse.dev/docs/regkit/overview/#capture-table", true},
    {L"MMCSS Values", L"https://noverse.dev/docs/win-config/system/mmcss-values/#registry-values"},
    {L"DWM Values", L"https://noverse.dev/docs/win-config/system/dwm-values/#registry-values"},
    {L"Session Manager Values", L"https://noverse.dev/docs/win-config/system/kernel-values/"},
    {L"DXG Kernel Values", L"https://www.noverse.dev/docs/win-config/system/dxg-kernel-values/#registry-values"},
    {L"BCD Edits", L"https://www.noverse.dev/docs/win-config/system/bcd-edits/#registry-values"},
    {L"Accessibility Values", L"https://noverse.dev/docs/win-config/system/disable-accessibility-features/#systemsettings-captures"},
    {L"Explorer Values", L"https://noverse.dev/docs/win-config/visibility/explorer-options/#explorer-captures"},
    {L"Taskbar Values", L"https://noverse.dev/docs/win-config/visibility/taskbar-settings/#systemsettings-captures"},
    {L"UserPreferenceMask Bitmask", L"https://noverse.dev/docs/win-config/visibility/minimal-visual-effects/#userpreferencesmask"},
    {L"DiagnosticDataSettings Values", L"https://noverse.dev/docs/win-config/privacy/disable-general-telemetry/#diagnosticdatasettings-values"},
    {L"SubscribedContent IDs", L"https://noverse.dev/docs/win-config/privacy/disable-suggestions-tips-tricks/#subscribedcontent-ids"},
    {L"Mouse Values", L"https://noverse.dev/docs/win-config/peripheral/mouse-values/#rawmousethrottle-details"},
    {L"USBFLAGS Values", L"https://www.noverse.dev/docs/win-config/peripheral/usbflags-values/#registry-values"},
    {L"USB Values", L"https://www.noverse.dev/docs/win-config/peripheral/usb-values/#registry-values"},
    {L"USBHUB Values", L"https://www.noverse.dev/docs/win-config/peripheral/usbhub-values/#registry-values"},
    {L"Audio Values", L"https://noverse.dev/docs/win-config/peripheral/audio-values/#registry-values"},
    {L"StorNVMe Values", L"https://www.noverse.dev/docs/win-config/peripheral/stornvme-values/#registry-values"},
    {L"StorPort Values", L"https://noverse.dev/docs/win-config/peripheral/storport-values/#registry-values"},
    {L"PnP Device Values", L"https://www.noverse.dev/docs/win-config/power/pnp-device-values/#registry-values"},
    {L"Power Values", L"https://www.noverse.dev/docs/win-config/power/power-values/#registry-values"},
    {L"Windows Security Values", L"https://noverse.dev/docs/win-config/security/windows-defender/#windows-security-captures"},
}};

static_assert(kLinks.size() ==
              static_cast<std::size_t>(cmd::kResearchItemMax -
                                       cmd::kResearchItemBase + 1));

} // namespace

std::span<const ResearchLink> ResearchLinks() noexcept {
  return kLinks;
}

const ResearchLink* ResearchLinkForCommand(int command_id) noexcept {
  if (command_id < cmd::kResearchItemBase ||
      command_id > cmd::kResearchItemMax) {
    return nullptr;
  }
  return &kLinks[static_cast<std::size_t>(command_id -
                                          cmd::kResearchItemBase)];
}

} // namespace regkit::frame
