// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/system/file_io.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include <cstdint>
#include <string>

namespace regkit::util
{

inline constexpr uint64_t kMaxStateFileBytes = 4ull * 1024ull * 1024ull;
inline constexpr uint64_t kMaxCommentFileBytes = 16ull * 1024ull * 1024ull;
inline constexpr uint64_t kMaxHistoryFileBytes = 64ull * 1024ull * 1024ull;

using win32ui::RandomFileSuffix;

bool ReadStateFile(const std::wstring& path, std::wstring* output, std::wstring* error = nullptr, uint64_t max_bytes = kMaxStateFileBytes);
std::wstring BadLineReason(size_t line);

}
