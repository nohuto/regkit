// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "search/search.h"

#include <string>
#include <vector>

namespace regkit::search
{

bool ParseResults(const std::wstring& content, std::vector<Result>* out, std::wstring* error = nullptr);
std::wstring SerializeResults(const std::vector<Result>& results);
bool LoadResults(const std::wstring& path, std::vector<Result>* results, std::wstring* error = nullptr);
bool SaveResults(const std::wstring& path, const std::vector<Result>& results);

}
