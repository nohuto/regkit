// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "regfile/reg_exe_syntax.h"

#include "registry/value_format.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <algorithm>

namespace regkit::reg_exe
{
namespace
{

struct ValueType
{
    const wchar_t* name;
    DWORD type;
};

constexpr ValueType kTypes[] = {
    {L"REG_SZ", REG_SZ},
    {L"REG_MULTI_SZ", REG_MULTI_SZ},
    {L"REG_EXPAND_SZ", REG_EXPAND_SZ},
    {L"REG_DWORD", REG_DWORD},
    {L"REG_DWORD_LITTLE_ENDIAN", REG_DWORD_LITTLE_ENDIAN},
    {L"REG_DWORD_BIG_ENDIAN", REG_DWORD_BIG_ENDIAN},
    {L"REG_QWORD", REG_QWORD},
    {L"REG_QWORD_LITTLE_ENDIAN", REG_QWORD_LITTLE_ENDIAN},
    {L"REG_BINARY", REG_BINARY},
    {L"REG_NONE", REG_NONE},
    {L"REG_LINK", REG_LINK},
};

bool Fail(std::wstring* error, const std::wstring& message)
{
    if (error)
    {
        *error = message;
    }
    return false;
}

}

bool IsSwitch(std::wstring_view text, std::wstring_view name)
{
    return !text.empty() && (text[0] == L'/' || text[0] == L'-') && util::EqualsInsensitive(text.substr(1), name);
}

bool ParseOptions(const std::vector<std::wstring>& args, size_t first, Options* options, std::vector<std::wstring>* positional, Verb verb, std::wstring* error)
{
    auto is_switch = [](const std::wstring& text) { return !text.empty() && (text[0] == L'/' || text[0] == L'-'); };
    int selectors = 0;
    int views = 0;
    int outputs = 0;
    for (size_t i = first; i < args.size(); ++i)
    {
        const std::wstring& arg = args[i];
        auto next = [&](std::wstring* out) -> bool {
            if (i + 1 >= args.size())
            {
                return Fail(error, util::TrLabel(L"Missing argument", arg));
            }
            *out = args[++i];
            return true;
        };
        if (is_switch(arg))
        {
            std::wstring name = util::ToLower(std::wstring_view(arg).substr(1));
            if (std::find(options->switches.begin(), options->switches.end(), name) != options->switches.end() || ((name == L"v" || name == L"ve" || name == L"va") && ++selectors > 1) ||
                (name.starts_with(L"reg:") && ++views > 1) || ((name == L"oa" || name == L"od" || name == L"os" || name == L"on") && ++outputs > 1))
            {
                return Fail(error, util::TrLabel(L"Invalid option", arg));
            }
            options->switches.push_back(std::move(name));
        }
        if (IsSwitch(arg, L"v") && verb == Verb::kQuery && (i + 1 >= args.size() || is_switch(args[i + 1])))
        {
            options->value_names = true;
        }
        else if (IsSwitch(arg, L"v"))
        {
            if (!next(&options->value_name))
                return false;
            options->has_value = true;
        }
        else if (IsSwitch(arg, L"ve"))
        {
            options->default_value = true;
            options->has_value = true;
        }
        else if (IsSwitch(arg, L"va"))
        {
            options->all_values = true;
        }
        else if (IsSwitch(arg, L"t"))
        {
            if (!next(&options->type_text))
                return false;
        }
        else if (IsSwitch(arg, L"d") && verb == Verb::kQuery)
        {
            options->data_only = true;
        }
        else if (IsSwitch(arg, L"d"))
        {
            if (!next(&options->data))
                return false;
            options->has_data = true;
        }
        else if (IsSwitch(arg, L"s") && verb == Verb::kAdd)
        {
            if (!next(&options->separator))
                return false;
        }
        else if (IsSwitch(arg, L"s"))
        {
            options->recurse = true;
        }
        else if (IsSwitch(arg, L"se") && verb == Verb::kQuery)
        {
            if (!next(&options->separator))
                return false;
        }
        else if (IsSwitch(arg, L"f") && verb == Verb::kQuery)
        {
            if (!next(&options->find))
                return false;
            options->has_find = true;
        }
        else if (IsSwitch(arg, L"f") || IsSwitch(arg, L"y"))
        {
            options->force = true;
        }
        else if (verb == Verb::kQuery && IsSwitch(arg, L"k"))
        {
            options->keys_only = true;
        }
        else if (verb == Verb::kQuery && IsSwitch(arg, L"c"))
        {
            options->case_sensitive = true;
        }
        else if (verb == Verb::kQuery && IsSwitch(arg, L"e"))
        {
            options->exact = true;
        }
        else if (verb == Verb::kQuery && IsSwitch(arg, L"z"))
        {
            options->verbose = true;
        }
        else if (IsSwitch(arg, L"oa") || IsSwitch(arg, L"od") || IsSwitch(arg, L"os") || IsSwitch(arg, L"on"))
        {
            options->compare_output = static_cast<wchar_t>(towlower(arg[2]));
        }
        else if (IsSwitch(arg, L"reg:32"))
        {
            options->view = KEY_WOW64_32KEY;
        }
        else if (IsSwitch(arg, L"reg:64"))
        {
            options->view = KEY_WOW64_64KEY;
        }
        else if (!arg.empty() && (arg[0] == L'/' || arg[0] == L'-'))
        {
            return Fail(error, util::TrLabel(L"Invalid option", arg));
        }
        else
        {
            positional->push_back(arg);
        }
    }
    DWORD type = REG_SZ;
    if (verb == Verb::kAdd && std::find(options->switches.begin(), options->switches.end(), L"s") != options->switches.end() &&
        (options->type_text.empty() || ParseType(options->type_text, &type)) && type != REG_MULTI_SZ)
    {
        return Fail(error, util::TrLabel(L"Invalid option", L"/s"));
    }
    return true;
}

bool ParseType(std::wstring_view text, DWORD* type)
{
    for (const ValueType& entry : kTypes)
    {
        if (util::EqualsInsensitive(text, entry.name))
        {
            *type = entry.type;
            return true;
        }
    }
    // reg.exe also accepts signed 32 bit decimal type numbers
    const bool negative = text.starts_with(L'-');
    const std::wstring_view digits = text.substr(negative || text.starts_with(L'+') ? 1 : 0);
    long long number = 0;
    for (wchar_t digit : digits)
    {
        if (digit < L'0' || digit > L'9' || (number = number * 10 + (digit - L'0')) > 0x80000000ll)
        {
            return false;
        }
    }
    if (digits.empty() || number > (negative ? 0x80000000ll : 0x7FFFFFFFll))
    {
        return false;
    }
    *type = static_cast<DWORD>(negative ? -number : number);
    return true;
}

std::wstring TypeName(DWORD type)
{
    for (const ValueType& entry : kTypes)
    {
        if (entry.type == type)
        {
            return entry.name;
        }
    }
    return value_format::TypeName(type);
}

bool BuildData(DWORD type, std::wstring_view text, std::wstring_view separator, std::vector<BYTE>* data, std::wstring* error)
{
    data->clear();
    switch (type)
    {
    default: // REG_NONE and unknown types use text, like reg.exe
        *data = value_format::StringData(text);
        return true;
    case REG_MULTI_SZ:
        {
            std::vector<std::wstring> items;
            for (size_t start = 0; !text.empty();)
            {
                const size_t end = separator.empty() ? std::wstring_view::npos : text.find(separator, start);
                items.emplace_back(text.substr(start, end == std::wstring_view::npos ? std::wstring_view::npos : end - start));
                if (end == std::wstring_view::npos)
                {
                    break;
                }
                start = end + separator.size();
            }
            if (!items.empty() && items.back().empty())
            {
                items.pop_back();
            }
            for (const std::wstring& item : items)
            {
                if (item.empty() && items.size() > 1)
                {
                    return Fail(error, util::TrLabel(L"Invalid multi-string data", text));
                }
            }
            *data = value_format::MultiStringData(items);
            return true;
        }
    case REG_DWORD:
    case REG_DWORD_BIG_ENDIAN:
    case REG_QWORD:
        {
            std::wstring number = util::TrimWhitespace(text);
            if (!number.empty() && number.front() == L'+')
            {
                number.erase(0, 1);
            }
            const std::optional<uint64_t> value = win32ui::ParseUnsigned(number, 10, true);
            if (!value)
            {
                return Fail(error, util::TrLabel(L"Invalid numeric data", text));
            }
            if (type != REG_QWORD && *value > 0xFFFFFFFFull)
            {
                return Fail(error, util::TrLabel(L"Numeric data out of range for a DWORD", text));
            }
            // reg.exe stores REG_DWORD_BIG_ENDIAN as little endian, like REG_DWORD
            *data = value_format::UnsignedBytes(*value, type == REG_QWORD ? sizeof(ULONGLONG) : sizeof(DWORD));
            return true;
        }
    case REG_BINARY:
        {
            std::wstring digits(text);
            for (wchar_t character : digits)
            {
                if (util::HexDigitValue(character) < 0)
                {
                    return Fail(error, util::TrLabel(L"Invalid binary data", text));
                }
            }
            if (digits.size() % 2 != 0)
            {
                digits.insert(digits.begin(), L'0');
            }
            return win32ui::ParseHexBytes(digits, data) || Fail(error, util::TrLabel(L"Invalid binary data", text));
        }
    }
}

}
