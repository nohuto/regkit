// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "defaults/default_data.h"

#include <atomic>
#include <functional>
#include <vector>

namespace regkit::defaults {

using NormalizePath =
    std::function<std::wstring(const std::wstring& path)>;

bool Load(const std::wstring& path, const NormalizePath& normalize,
          Data* data, std::vector<Entry>* entries,
          std::wstring* error,
          const std::atomic_bool* cancel = nullptr);

} // namespace regkit::defaults
