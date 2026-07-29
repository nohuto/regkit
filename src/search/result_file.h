// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "search/search.h"

#include <string>
#include <vector>

namespace regkit::search {

std::vector<Result> ParseResults(const std::wstring& content);
std::wstring SerializeResults(const std::vector<Result>& results);
bool LoadResults(const std::wstring& path,
                 std::vector<Result>* results);
bool SaveResults(const std::wstring& path,
                 const std::vector<Result>& results);

} // namespace regkit::search
