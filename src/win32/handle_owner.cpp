// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/handle_owner.h"

namespace util {

ComInit::ComInit(DWORD flags) noexcept : hr_(CoInitializeEx(nullptr, flags)) {}

ComInit::~ComInit() {
  if (SUCCEEDED(hr_)) {
    CoUninitialize();
  }
}

bool ComInit::ok() const noexcept {
  return SUCCEEDED(hr_);
}

UniqueHKey::UniqueHKey(HKEY key) noexcept : key_(key) {}

UniqueHKey::~UniqueHKey() {
  reset();
}

UniqueHKey::UniqueHKey(UniqueHKey&& other) noexcept : key_(other.key_) {
  other.key_ = nullptr;
}

UniqueHKey& UniqueHKey::operator=(UniqueHKey&& other) noexcept {
  if (this != &other) {
    reset();
    key_ = other.key_;
    other.key_ = nullptr;
  }
  return *this;
}

HKEY UniqueHKey::get() const noexcept {
  return key_;
}

HKEY* UniqueHKey::put() noexcept {
  reset();
  return &key_;
}

HKEY UniqueHKey::release() noexcept {
  HKEY key = key_;
  key_ = nullptr;
  return key;
}

void UniqueHKey::reset(HKEY key) noexcept {
  if (key_) {
    RegCloseKey(key_);
  }
  key_ = key;
}

} // namespace util
