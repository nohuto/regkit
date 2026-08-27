// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/file_text.h"

#include <limits>

namespace util {

std::string WideToUtf8(const std::wstring& text) {
  if (text.empty()) {
    return {};
  }
  const int size = WideCharToMultiByte(
      CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0,
      nullptr, nullptr);
  if (size <= 0) {
    return {};
  }
  std::string output(static_cast<size_t>(size), '\0');
  WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                      output.data(), size, nullptr, nullptr);
  return output;
}

std::wstring Utf8ToWide(std::string_view text) {
  if (text.empty()) {
    return {};
  }
  const int size = MultiByteToWideChar(
      CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
  if (size <= 0) {
    return {};
  }
  std::wstring output(static_cast<size_t>(size), L'\0');
  MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                      output.data(), size);
  return output;
}

bool ReadFileBytes(const std::wstring& path, std::vector<BYTE>* output,
                   uint64_t max_bytes, DWORD share_mode) {
  if (!output) {
    return false;
  }
  output->clear();
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, share_mode, nullptr,
                            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  LARGE_INTEGER size = {};
  const bool valid =
      GetFileSizeEx(file, &size) && size.QuadPart > 0 &&
      static_cast<uint64_t>(size.QuadPart) <= max_bytes &&
      size.QuadPart <= static_cast<LONGLONG>(std::numeric_limits<DWORD>::max());
  if (!valid) {
    CloseHandle(file);
    return false;
  }
  output->resize(static_cast<size_t>(size.QuadPart));
  DWORD read = 0;
  const BOOL result = ReadFile(file, output->data(),
                               static_cast<DWORD>(output->size()), &read,
                               nullptr);
  CloseHandle(file);
  if (!result || read != output->size()) {
    output->clear();
    return false;
  }
  return true;
}

bool ReadTextFile(const std::wstring& path, std::wstring* output, bool* utf16,
                  uint64_t max_bytes, DWORD share_mode) {
  if (!output) {
    return false;
  }
  output->clear();
  if (utf16) {
    *utf16 = false;
  }
  std::vector<BYTE> bytes;
  if (!ReadFileBytes(path, &bytes, max_bytes, share_mode)) {
    return false;
  }
  if (bytes.size() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) {
    output->assign(reinterpret_cast<const wchar_t*>(bytes.data() + 2),
                   (bytes.size() - 2) / sizeof(wchar_t));
    if (utf16) {
      *utf16 = true;
    }
    return !output->empty();
  }
  size_t offset = 0;
  if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB &&
      bytes[2] == 0xBF) {
    offset = 3;
  }
  *output = Utf8ToWide(std::string_view(
      reinterpret_cast<const char*>(bytes.data() + offset),
      bytes.size() - offset));
  return !output->empty();
}

bool WriteTextFile(const std::wstring& path, const std::wstring& text,
                   bool utf16) {
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                            nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL,
                            nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  DWORD written = 0;
  if (utf16) {
    constexpr BYTE bom[] = {0xFF, 0xFE};
    if (!WriteFile(file, bom, sizeof(bom), &written, nullptr) ||
        written != sizeof(bom)) {
      CloseHandle(file);
      return false;
    }
    if (text.empty()) {
      CloseHandle(file);
      return true;
    }
    const DWORD byte_count =
        static_cast<DWORD>(text.size() * sizeof(wchar_t));
    const BOOL result =
        WriteFile(file, text.data(), byte_count, &written, nullptr);
    CloseHandle(file);
    return result && written == byte_count;
  }
  const std::string utf8 = WideToUtf8(text);
  const DWORD byte_count = static_cast<DWORD>(utf8.size());
  const BOOL result =
      byte_count == 0 ||
      WriteFile(file, utf8.data(), byte_count, &written, nullptr);
  CloseHandle(file);
  return result && written == byte_count;
}

} // namespace util
