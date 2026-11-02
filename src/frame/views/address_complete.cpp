// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"
#include "frame/window_impl.h"

#include "ui/autocomplete.h"

#include <win32ui/dialog/autocomplete.h>

namespace regkit
{
using namespace window_detail;

void MainWindow::Impl::EnableAddressAutoComplete()
{
    if (!address_autocomplete_)
    {
        appearance::SetKeySuggest([this](const std::wstring& text) { return BuildAddressSuggestions(text); });
        address_autocomplete_ = win32ui::AttachAutoComplete(browse_.address(), appearance::SuggestKeys);
    }
}

std::vector<std::wstring> MainWindow::Impl::BuildAddressSuggestions(const std::wstring& text) const
{
    const wchar_t separator = text.find(L'\\') == std::wstring::npos && text.find(L'/') != std::wstring::npos ? L'/' : L'\\';
    const size_t sep = text.find_last_of(separator);
    std::vector<std::wstring> items;
    if (sep == std::wstring::npos)
    {
        std::unordered_set<std::wstring> seen;
        auto add = [&](const std::wstring& root) {
            if (StartsWithInsensitive(root, text) && seen.insert(ToLower(root)).second)
            {
                items.push_back(root);
            }
        };
        for (const auto& root : browse_.roots())
        {
            add(root.path_name);
        }
        for (const wchar_t* alias : {L"HKCR", L"HKCU", L"HKLM", L"HKU", L"HKCC", L"HKEY_CLASSES_ROOT", L"HKEY_CURRENT_USER", L"HKEY_LOCAL_MACHINE", L"HKEY_USERS", L"HKEY_CURRENT_CONFIG"})
        {
            add(alias);
        }
        return items;
    }
    const std::wstring parent = text.substr(0, sep + 1);
    const std::wstring_view partial = std::wstring_view(text).substr(sep + 1);
    RegistryNode node;
    if (!ResolvePathToNode(NormalizeRegistryPath(parent), &node))
    {
        return items;
    }
    // popup filters after separators, it needs every child
    for (const auto& raw_name : RegistryStore::EnumSubKeyNames(node, true))
    {
        std::wstring name = registry_path::DisplayName(raw_name);
        if (StartsWithInsensitive(name, partial))
        {
            items.push_back(parent + name);
        }
    }
    return items;
}

}
