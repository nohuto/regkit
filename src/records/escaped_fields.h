// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>
#include <vector>

namespace regkit::record_fields {

std::wstring Escape(const std::wstring& text);
std::wstring Unescape(const std::wstring& text);
std::vector<std::wstring> Split(const std::wstring& line);
std::vector<std::wstring> Lines(const std::wstring& content);

} // namespace regkit::record_fields
