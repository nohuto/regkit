// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/translation.h"

#include "win32/shell_paths.h"

namespace regkit::util
{
namespace
{

std::wstring LanguageFolder()
{
    return JoinPath(ModuleDirectory(), L"assets\\lang");
}

}

bool LoadLanguage(const std::wstring& code)
{
    return win32ui::LoadLanguage(LanguageFolder().c_str(), code) == ERROR_SUCCESS;
}

std::vector<LanguagePack> InstalledLanguages()
{
    return win32ui::InstalledLanguages(LanguageFolder().c_str());
}

std::wstring TrDetail(const wchar_t* headline, std::wstring_view detail)
{
    return std::wstring(Tr(headline)).append(L"\n").append(detail);
}

std::wstring TrLabel(const wchar_t* label, std::wstring_view value)
{
    return std::wstring(Tr(label)).append(L": ").append(value);
}

}
