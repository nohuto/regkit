// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/text/language_pack.h>
#include <win32ui/text/text.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include <string>
#include <string_view>
#include <vector>

namespace regkit::util
{

using win32ui::LanguagePack;

bool LoadLanguage(const std::wstring& code);
std::vector<LanguagePack> InstalledLanguages();
inline const wchar_t* Tr(const wchar_t* text)
{
    return win32ui::Text(text);
}
std::wstring TrDetail(const wchar_t* headline, std::wstring_view detail);
std::wstring TrLabel(const wchar_t* label, std::wstring_view value);
constexpr const wchar_t* TrNoop(const wchar_t* text)
{
    return text;
}

}
