// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <memory>
#include <string>

namespace regkit {

class MainWindow {
public:
  MainWindow();
  ~MainWindow();

  MainWindow(const MainWindow&) = delete;
  MainWindow& operator=(const MainWindow&) = delete;

  bool Create(HINSTANCE instance);
  void Show(int command);
  bool OpenRegFileTab(const std::wstring& path);
  bool TranslateAccelerator(const MSG& message);
  void QueueExternalJump(const std::wstring& target);

private:
  friend class MainWindowBenchmarks;
  friend class MainWindowCharacterization;
  friend class RegistryAddressEnum;

  class Impl;
  std::unique_ptr<Impl> impl_;
};

} // namespace regkit
