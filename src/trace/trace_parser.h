// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "trace/trace_data.h"

#include <atomic>
#include <functional>
#include <string>
#include <string_view>

namespace regkit::trace {

struct Normalizers {
  std::function<std::wstring(const std::wstring&)> key;
  std::function<std::wstring(const std::wstring&)> display;
};

using EntryCallback = std::function<bool(Entry&& entry)>;

bool ParseEntries(std::string_view buffer,
                  const Normalizers& normalizers,
                  const EntryCallback& callback, std::wstring* error,
                  const std::atomic_bool* cancel = nullptr);

bool Parse(const std::wstring& label, const std::wstring& source,
           std::string_view buffer, const Normalizers& normalizers,
           Data* data, std::wstring* error,
           const std::atomic_bool* cancel = nullptr);

} // namespace regkit::trace
