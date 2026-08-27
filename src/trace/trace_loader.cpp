// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "trace/trace_loader.h"

#include "win32/file_text.h"

#include <limits>
#include <string_view>

namespace regkit::trace {

namespace {

bool Read(const std::wstring& path, std::vector<BYTE>* bytes,
          std::wstring* error, const std::atomic_bool* cancel) {
  if (cancel && cancel->load()) {
    return false;
  }
  if (!util::ReadFileBytes(
          path, bytes,
          static_cast<uint64_t>(std::numeric_limits<int>::max()))) {
    if (error) {
      *error = L"Failed to read trace file.";
    }
    return false;
  }
  return true;
}

} // namespace

bool LoadEntries(const std::wstring& path,
                 const Normalizers& normalizers,
                 const EntryCallback& callback,
                 std::wstring* error,
                 const std::atomic_bool* cancel) {
  std::vector<BYTE> bytes;
  if (!Read(path, &bytes, error, cancel)) {
    return false;
  }
  const std::string_view buffer(
      reinterpret_cast<const char*>(bytes.data()), bytes.size());
  return ParseEntries(buffer, normalizers, callback, error, cancel);
}

bool Load(const std::wstring& label, const std::wstring& path,
          const Normalizers& normalizers, Data* data,
          std::wstring* error, const std::atomic_bool* cancel) {
  std::vector<BYTE> bytes;
  if (!Read(path, &bytes, error, cancel)) {
    return false;
  }
  const std::string_view buffer(
      reinterpret_cast<const char*>(bytes.data()), bytes.size());
  return Parse(label, path, buffer, normalizers, data, error, cancel);
}

} // namespace regkit::trace
