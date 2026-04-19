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

#include <windows.h>
#include <objbase.h>

#include <string>

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

template <typename T> class UniqueGdiObject {
public:
  UniqueGdiObject() noexcept = default;
  explicit UniqueGdiObject(T handle) noexcept : handle_(handle) {}
  ~UniqueGdiObject() { reset(); }
  UniqueGdiObject(const UniqueGdiObject&) = delete;
  UniqueGdiObject& operator=(const UniqueGdiObject&) = delete;
  UniqueGdiObject(UniqueGdiObject&& other) noexcept : handle_(other.handle_) { other.handle_ = nullptr; }
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
    T temp = handle_;
    handle_ = nullptr;
    return temp;
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

std::wstring GetModuleDirectory();
std::wstring JoinPath(const std::wstring& left, const std::wstring& right);
std::string WideToUtf8(const std::wstring& text);
std::wstring Utf8ToWide(const std::string& text);
std::wstring ToHex(const BYTE* data, size_t size, size_t max_bytes);
std::wstring GetCurrentUserSidString();
std::wstring GetAppDataFolder();
bool IsProcessElevated();
bool IsProcessSystem();
bool IsProcessTrustedInstaller();
bool LaunchProcessAsSystem(const std::wstring& command_line, const std::wstring& work_dir, DWORD* error_code = nullptr);
bool LaunchProcessAsTrustedInstaller(const std::wstring& command_line, const std::wstring& work_dir, DWORD* error_code = nullptr);

} // namespace util
