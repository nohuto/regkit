// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

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
