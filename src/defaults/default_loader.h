// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "defaults/default_data.h"

#include <atomic>
#include <functional>
#include <vector>

namespace regkit::defaults
{

using NormalizePath = std::function<std::wstring(const std::wstring& path)>;

bool Load(const std::wstring& path, const NormalizePath& normalize, Data* data, std::vector<Entry>* entries, std::wstring* error, const std::atomic_bool* cancel = nullptr);
std::wstring ShortLabel(const std::wstring& label, const std::wstring& source_path);

}
