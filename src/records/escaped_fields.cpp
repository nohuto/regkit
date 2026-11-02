// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "records/escaped_fields.h"

#include <win32ui/system/strings.h>

namespace regkit::record_fields
{

namespace
{

bool IsHigh(wchar_t character)
{
    return character >= 0xD800 && character <= 0xDBFF;
}

bool IsLow(wchar_t character)
{
    return character >= 0xDC00 && character <= 0xDFFF;
}

// escape lone surrogates as \uXXXX to preserve invalid UTF-16 registry names in UTF-8 files
void AppendEscaped(std::wstring* output, std::wstring_view text)
{
    size_t start = 0;
    for (size_t index = 0; index < text.size(); ++index)
    {
        const wchar_t character = text[index];
        if (character >= L' ' && character != L'\\' && (character < 0xD800 || character > 0xDFFF))
        {
            continue;
        }
        const wchar_t code = character == L'\\' ? L'\\' : character == L'\t' ? L't'
                                                      : character == L'\r'   ? L'r'
                                                      : character == L'\n'   ? L'n'
                                                                             : 0;
        const bool lone = (IsHigh(character) && (index + 1 == text.size() || !IsLow(text[index + 1]))) ||
                          (IsLow(character) && (index == 0 || !IsHigh(text[index - 1])));
        if (!code && !lone)
        {
            continue;
        }
        output->append(text.substr(start, index - start));
        if (code)
        {
            output->append({L'\\', code});
        }
        else
        {
            static constexpr wchar_t kHex[] = L"0123456789ABCDEF";
            output->append({L'\\', L'u', kHex[character >> 12], kHex[(character >> 8) & 15], kHex[(character >> 4) & 15], kHex[character & 15]});
        }
        start = index + 1;
    }
    output->append(text.substr(start));
}

void AppendUnescaped(std::wstring* output, std::wstring_view text)
{
    size_t start = 0;
    for (size_t slash = text.find(L'\\'); slash != std::wstring_view::npos && slash + 1 < text.size();
         slash = text.find(L'\\', start))
    {
        const wchar_t next = text[slash + 1];
        const wchar_t decoded = next == L't'    ? L'\t'
                                : next == L'r'  ? L'\r'
                                : next == L'n'  ? L'\n'
                                : next == L'\\' ? L'\\'
                                                : 0;
        if (!decoded)
        {
            int value = next == L'u' && slash + 6 <= text.size() ? 0 : -1;
            for (size_t index = slash + 2; index < slash + 6 && value >= 0; ++index)
            {
                const int digit = win32ui::HexDigitValue(text[index]);
                value = digit < 0 ? -1 : value * 16 + digit;
            }
            const bool surrogate = IsHigh(static_cast<wchar_t>(value)) || IsLow(static_cast<wchar_t>(value));
            output->append(text.substr(start, slash + (surrogate ? 0 : 1) - start));
            if (surrogate)
            {
                output->push_back(static_cast<wchar_t>(value));
            }
            start = slash + (surrogate ? 6 : 1);
            continue;
        }
        output->append(text.substr(start, slash - start));
        output->push_back(decoded);
        start = slash + 2;
    }
    output->append(text.substr(start));
}

template <typename Fields>
void AppendFields(std::wstring* output, const Fields& fields)
{
    bool first = true;
    for (std::wstring_view field : fields)
    {
        if (!first)
        {
            output->push_back(L'\t');
        }
        AppendEscaped(output, field);
        first = false;
    }
    output->push_back(L'\n');
}

}

std::wstring Escape(std::wstring_view text)
{
    std::wstring escaped;
    escaped.reserve(text.size());
    AppendEscaped(&escaped, text);
    return escaped;
}

std::wstring Unescape(std::wstring_view text)
{
    std::wstring unescaped;
    unescaped.reserve(text.size());
    AppendUnescaped(&unescaped, text);
    return unescaped;
}

void AppendRecord(std::wstring* output, std::initializer_list<std::wstring_view> fields)
{
    AppendFields(output, fields);
}

void AppendRecord(std::wstring* output, std::span<const std::wstring> fields)
{
    AppendFields(output, fields);
}

std::vector<std::wstring> DecodeRecord(std::wstring_view line)
{
    std::vector<std::wstring> fields;
    size_t start = 0;
    for (;;)
    {
        const size_t separator = line.find(L'\t', start);
        std::wstring& field = fields.emplace_back();
        AppendUnescaped(&field, line.substr(start, separator == std::wstring_view::npos ? std::wstring_view::npos : separator - start));
        if (separator == std::wstring_view::npos)
        {
            return fields;
        }
        start = separator + 1;
    }
}

bool ParseUnsigned(std::wstring_view text, uint64_t maximum, uint64_t* value)
{
    if (text.empty())
    {
        return false;
    }
    uint64_t result = 0;
    for (const wchar_t character : text)
    {
        if (character < L'0' || character > L'9')
        {
            return false;
        }
        const uint64_t digit = static_cast<uint64_t>(character - L'0');
        if (digit > maximum || result > (maximum - digit) / 10)
        {
            return false;
        }
        result = result * 10 + digit;
    }
    *value = result;
    return true;
}

}
