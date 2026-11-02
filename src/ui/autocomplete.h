// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/dialog/autocomplete.h>

#include <string>
#include <vector>

namespace regkit::appearance
{

void SetKeySuggest(win32ui::AutoCompleteSuggest suggest);
std::vector<std::wstring> SuggestKeys(const std::wstring& text);

}
