// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "ui/autocomplete.h"

#include <utility>

namespace regkit::appearance
{
namespace
{

win32ui::AutoCompleteSuggest& KeySuggest()
{
    static win32ui::AutoCompleteSuggest suggest;
    return suggest;
}

}

void SetKeySuggest(win32ui::AutoCompleteSuggest suggest)
{
    KeySuggest() = std::move(suggest);
}

std::vector<std::wstring> SuggestKeys(const std::wstring& text)
{
    return KeySuggest() ? KeySuggest()(text) : std::vector<std::wstring>();
}

}
