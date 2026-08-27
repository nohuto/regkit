// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <objbase.h>
#include <windows.h>

namespace util {

class ComInit {
public:
  explicit ComInit(DWORD flags = COINIT_APARTMENTTHREADED) noexcept;
  ~ComInit();
  ComInit(const ComInit&) = delete;
  ComInit& operator=(const ComInit&) = delete;

  bool ok() const noexcept;

private:
  HRESULT hr_;
};

class UniqueHKey {
public:
  UniqueHKey() noexcept = default;
  explicit UniqueHKey(HKEY key) noexcept;
  ~UniqueHKey();
  UniqueHKey(UniqueHKey&& other) noexcept;
  UniqueHKey& operator=(UniqueHKey&& other) noexcept;
  UniqueHKey(const UniqueHKey&) = delete;
  UniqueHKey& operator=(const UniqueHKey&) = delete;

  HKEY get() const noexcept;
  HKEY* put() noexcept;
  HKEY release() noexcept;
  void reset(HKEY key = nullptr) noexcept;

private:
  HKEY key_ = nullptr;
};

template <typename T>
class UniqueGdiObject {
public:
  UniqueGdiObject() noexcept = default;
  explicit UniqueGdiObject(T handle) noexcept : handle_(handle) {}
  ~UniqueGdiObject() { reset(); }
  UniqueGdiObject(const UniqueGdiObject&) = delete;
  UniqueGdiObject& operator=(const UniqueGdiObject&) = delete;
  UniqueGdiObject(UniqueGdiObject&& other) noexcept : handle_(other.handle_) {
    other.handle_ = nullptr;
  }
  UniqueGdiObject& operator=(UniqueGdiObject&& other) noexcept {
    if (this != &other) {
      reset();
      handle_ = other.handle_;
      other.handle_ = nullptr;
    }
    return *this;
  }

  T get() const noexcept { return handle_; }
  T* put() noexcept {
    reset();
    return &handle_;
  }
  T release() noexcept {
    T handle = handle_;
    handle_ = nullptr;
    return handle;
  }
  void reset(T handle = nullptr) noexcept {
    if (handle_) {
      DeleteObject(handle_);
    }
    handle_ = handle;
  }
  explicit operator bool() const noexcept { return handle_ != nullptr; }

private:
  T handle_ = nullptr;
};

} // namespace util
