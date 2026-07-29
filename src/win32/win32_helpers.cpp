// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "../../include/win32/win32_helpers.h"

#include <cwctype>
#include <iterator>

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

std::wstring FormatWin32Error(DWORD code) {
  if (code == ERROR_SUCCESS) {
    return {};
  }
  wchar_t buffer[512] = {};
  DWORD length = FormatMessageW(
      FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, code,
      0, buffer, static_cast<DWORD>(std::size(buffer)), nullptr);
  if (length == 0) {
    return L"Unknown error.";
  }
  while (length > 0 &&
         (buffer[length - 1] == L'\r' || buffer[length - 1] == L'\n')) {
    buffer[--length] = L'\0';
  }
  return buffer;
}

std::wstring ToLower(const std::wstring& text) {
  std::wstring result = text;
  for (wchar_t& character : result) {
    character = towlower(character);
  }
  return result;
}

std::wstring TrimWhitespace(const std::wstring& text) {
  size_t first = 0;
  while (first < text.size() && iswspace(text[first])) {
    ++first;
  }
  size_t last = text.size();
  while (last > first && iswspace(text[last - 1])) {
    --last;
  }
  return text.substr(first, last - first);
}

std::wstring ExpandEnvironmentStringsDynamic(const std::wstring& text) {
  if (text.empty()) {
    return {};
  }
  const DWORD needed = ExpandEnvironmentStringsW(text.c_str(), nullptr, 0);
  if (needed == 0) {
    return {};
  }
  std::wstring expanded(needed, L'\0');
  const DWORD written =
      ExpandEnvironmentStringsW(text.c_str(), expanded.data(), needed);
  if (written == 0 || written > needed) {
    return {};
  }
  while (!expanded.empty() && expanded.back() == L'\0') {
    expanded.pop_back();
  }
  return expanded;
}

std::wstring ToHex(const BYTE* data, size_t size, size_t max_bytes) {
  if (!data || size == 0) {
    return {};
  }
  const size_t count =
      max_bytes == 0 || max_bytes > size ? size : max_bytes;
  static constexpr wchar_t digits[] = L"0123456789abcdef";
  std::wstring output;
  output.reserve(count * 3 + (count < size ? 4 : 0));
  for (size_t index = 0; index < count; ++index) {
    output.push_back(digits[data[index] >> 4]);
    output.push_back(digits[data[index] & 0x0F]);
    if (index + 1 < count) {
      output.push_back(L' ');
    }
  }
  if (count < size) {
    output += L" ...";
  }
  return output;
}

} // namespace util
