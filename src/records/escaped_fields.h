// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <string>
#include <vector>

namespace regkit::record_fields {

std::wstring Escape(const std::wstring& text);
std::wstring Unescape(const std::wstring& text);
std::vector<std::wstring> Split(const std::wstring& line);
std::vector<std::wstring> Lines(const std::wstring& content);

} // namespace regkit::record_fields
