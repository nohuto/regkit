// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "trace/trace_parser.h"

namespace regkit::trace {

bool LoadEntries(const std::wstring& path,
                 const Normalizers& normalizers,
                 const EntryCallback& callback,
                 std::wstring* error,
                 const std::atomic_bool* cancel = nullptr);

bool Load(const std::wstring& label, const std::wstring& path,
          const Normalizers& normalizers, Data* data,
          std::wstring* error,
          const std::atomic_bool* cancel = nullptr);

} // namespace regkit::trace
