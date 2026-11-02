// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "ui/dialog_support.h"
#include "win32/text_transform.h"

#include <string>

namespace regkit::editors::dialog_support
{

bool Matches(const std::wstring& text, const std::wstring& filter)
{
    return filter.empty() || util::ContainsInsensitive(text, filter);
}

}
