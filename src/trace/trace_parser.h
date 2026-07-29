// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

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
