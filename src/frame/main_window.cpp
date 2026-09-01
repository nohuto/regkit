// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/main_window.h"

#include "frame/window_impl.h"

namespace regkit {

MainWindow::MainWindow() : impl_(std::make_unique<Impl>()) {}

MainWindow::~MainWindow() = default;

bool MainWindow::Create(HINSTANCE instance) {
  return impl_->Create(instance);
}

void MainWindow::Show(int command) { impl_->Show(command); }

bool MainWindow::OpenRegFileTab(const std::wstring& path) {
  return impl_->OpenRegFileTab(path);
}

bool MainWindow::TranslateAccelerator(const MSG& message) {
  return impl_->TranslateAccelerator(message);
}

void MainWindow::QueueExternalJump(const std::wstring& target) {
  impl_->QueueExternalJump(target);
}

} // namespace regkit
