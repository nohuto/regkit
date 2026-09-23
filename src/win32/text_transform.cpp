// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "win32/text_transform.h"

#include <algorithm>
#include <cwctype>

namespace util
{

std::wstring FormatLocalTime(const SYSTEMTIME& time, bool with_seconds)
{
    wchar_t date[80] = {};
    wchar_t clock[80] = {};
    if (!GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &time, nullptr, date, static_cast<int>(std::size(date)), nullptr) ||
        !GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, with_seconds ? 0 : TIME_NOSECONDS, &time, nullptr, clock, static_cast<int>(std::size(clock))))
    {
        return L"";
    }
    return std::wstring(date) + L' ' + clock;
}

std::wstring WindowText(HWND window)
{
    const int length = window ? GetWindowTextLengthW(window) : 0;
    if (length <= 0)
    {
        return {};
    }
    std::wstring text(static_cast<size_t>(length), L'\0');
    const int copied = GetWindowTextW(window, text.data(), length + 1);
    text.resize(copied > 0 ? static_cast<size_t>(copied) : 0);
    return text;
}

std::wstring DialogText(HWND dialog, int control_id)
{
    return dialog ? WindowText(GetDlgItem(dialog, control_id)) : std::wstring();
}

bool ParseUnsignedNumber(std::wstring_view text, int base, unsigned long long* value)
{
    if (!value)
    {
        return false;
    }
    if (base != 2 && base != 10 && base != 16)
    {
        base = 10;
    }
    std::wstring digits(TrimWhitespace(text));
    if (digits.empty() || digits.front() == L'-' || digits.front() == L'+')
    {
        return false;
    }
    if (digits.size() > 2 && digits[0] == L'0' && (digits[1] == L'x' || digits[1] == L'X'))
    {
        base = 16;
        digits.erase(0, 2);
    }
    else if (base == 2 && digits.size() > 2 && digits[0] == L'0' && (digits[1] == L'b' || digits[1] == L'B'))
    {
        digits.erase(0, 2);
    }
    if (digits.empty())
    {
        return false;
    }
    if (base == 2)
    {
        unsigned long long parsed = 0;
        bool saw_digit = false;
        for (const wchar_t character : digits)
        {
            if (character == L'_' || character == L'\'' || iswspace(character))
            {
                continue;
            }
            if (character != L'0' && character != L'1')
            {
                return false;
            }
            if (parsed > (std::numeric_limits<unsigned long long>::max() >> 1))
            {
                return false;
            }
            parsed = (parsed << 1) | static_cast<unsigned long long>(character - L'0');
            saw_digit = true;
        }
        if (!saw_digit)
        {
            return false;
        }
        *value = parsed;
        return true;
    }
    errno = 0;
    wchar_t* stop = nullptr;
    const unsigned long long parsed = wcstoull(digits.c_str(), &stop, base);
    if (!stop || *stop != L'\0' || errno == ERANGE)
    {
        return false;
    }
    *value = parsed;
    return true;
}

std::wstring ToLower(std::wstring_view text)
{
    std::wstring result(text);
    if (!result.empty())
    {
        CharLowerBuffW(result.data(), static_cast<DWORD>(result.size()));
    }
    return result;
}

std::wstring TrimWhitespace(std::wstring_view text)
{
    size_t first = 0;
    while (first < text.size() && iswspace(text[first]))
    {
        ++first;
    }
    size_t last = text.size();
    while (last > first && iswspace(text[last - 1]))
    {
        --last;
    }
    return std::wstring(text.substr(first, last - first));
}

bool IsBlank(std::wstring_view text)
{
    return std::all_of(text.begin(), text.end(), [](wchar_t character) { return iswspace(character) != 0; });
}

std::wstring ExpandEnvironmentStringsDynamic(const std::wstring& text)
{
    if (text.empty())
    {
        return {};
    }
    const DWORD needed = ExpandEnvironmentStringsW(text.c_str(), nullptr, 0);
    if (needed == 0)
    {
        return {};
    }
    std::wstring expanded(needed, L'\0');
    const DWORD written = ExpandEnvironmentStringsW(text.c_str(), expanded.data(), needed);
    if (written == 0 || written > needed)
    {
        return {};
    }
    while (!expanded.empty() && expanded.back() == L'\0')
    {
        expanded.pop_back();
    }
    return expanded;
}

std::wstring ToHex(std::span<const BYTE> data, wchar_t separator, bool uppercase, size_t max_bytes)
{
    const wchar_t* digits = uppercase ? L"0123456789ABCDEF" : L"0123456789abcdef";
    const size_t count = max_bytes == 0 ? data.size() : std::min(max_bytes, data.size());
    std::wstring output;
    output.reserve(count * 3 + 4);
    for (size_t index = 0; index < count; ++index)
    {
        if (index != 0 && separator)
        {
            output.push_back(separator);
        }
        output.push_back(digits[data[index] >> 4]);
        output.push_back(digits[data[index] & 0x0F]);
    }
    if (count < data.size())
    {
        output += L" ...";
    }
    return output;
}

int CompareInsensitive(std::wstring_view left, std::wstring_view right)
{
    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE) -
           CSTR_EQUAL;
}

int CompareListText(std::wstring_view left, std::wstring_view right)
{
    if (left.empty() || right.empty())
    {
        return static_cast<int>(left.empty()) - static_cast<int>(right.empty());
    }
    return CompareInsensitive(left, right);
}

bool EqualsInsensitive(std::wstring_view left, std::wstring_view right)
{
    return left.size() == right.size() && CompareInsensitive(left, right) == 0;
}

bool StartsWithInsensitive(std::wstring_view text, std::wstring_view prefix)
{
    return prefix.size() <= text.size() && EqualsInsensitive(text.substr(0, prefix.size()), prefix);
}

bool EndsWithInsensitive(std::wstring_view text, std::wstring_view suffix)
{
    return suffix.size() <= text.size() && EqualsInsensitive(text.substr(text.size() - suffix.size()), suffix);
}

int FindInsensitive(std::wstring_view text, std::wstring_view needle)
{
    if (needle.empty() || needle.size() > text.size())
    {
        return -1;
    }
    return FindStringOrdinal(FIND_FROMSTART, text.data(), static_cast<int>(text.size()), needle.data(), static_cast<int>(needle.size()), TRUE);
}

bool ContainsInsensitive(std::wstring_view text, std::wstring_view needle)
{
    return FindInsensitive(text, needle) >= 0;
}

bool ParseBool(std::wstring_view text)
{
    return text == L"1" || EqualsInsensitive(text, L"true") || EqualsInsensitive(text, L"yes");
}

std::vector<std::wstring> SplitLines(std::wstring_view text)
{
    std::vector<std::wstring> lines;
    size_t start = 0;
    while (start < text.size())
    {
        size_t end = text.find_first_of(L"\r\n", start);
        if (end == std::wstring_view::npos)
        {
            end = text.size();
        }
        std::wstring line = TrimWhitespace(text.substr(start, end - start));
        if (!line.empty())
        {
            lines.push_back(std::move(line));
        }
        start = end + 1;
    }
    return lines;
}

std::wstring JoinLines(const std::vector<std::wstring>& lines)
{
    std::wstring text;
    for (const std::wstring& line : lines)
    {
        if (line.empty())
        {
            continue;
        }
        if (!text.empty())
        {
            text.append(L"\r\n");
        }
        text.append(line);
    }
    return text;
}

} // namespace util
