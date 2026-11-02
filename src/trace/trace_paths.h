// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "trace/trace_parser.h"

#include <string>

namespace regkit::trace
{

std::wstring NormalizeKeyPathBasic(const std::wstring& text);
std::wstring NormalizeKeyPath(const std::wstring& text);
std::wstring NormalizeSelectionPath(const std::wstring& text);
Normalizers PathNormalizers();
std::wstring MapControlSetToCurrent(const std::wstring& path);

}
