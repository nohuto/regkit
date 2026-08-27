// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_impl.h"

#include "appearance/preset_editor.h"
#include "editors/value_editor.h"

namespace regkit {

void MainWindow::Impl::ShowThemePresetsDialog() {
  if (theme_presets_.empty()) {
    LoadThemePresets();
  }

  appearance::ShowThemePresetEditor(
      hwnd_, theme_presets_, active_theme_preset_,
      [](void* context, const std::vector<ThemePreset>& presets, const std::wstring& active_name) {
        static_cast<MainWindow::Impl*>(context)->UpdateThemePresets(presets, active_name, true);
      },
      [](void*, HWND owner, const wchar_t* title, const std::wstring& initial, std::wstring* name) {
        editors::TextRequest request;
        request.title = title;
        request.label = L"Preset name:";
        request.text = initial;
        editors::TextResult result;
        if (!editors::EditText(owner, request, &result)) {
          return false;
        }
        *name = std::move(result.text);
        return true;
      },
      this);
}

} // namespace regkit
