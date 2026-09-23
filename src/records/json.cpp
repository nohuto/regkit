// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "records/json.h"

#include <cwchar>

namespace regkit::json
{

namespace
{

bool Digit(wchar_t character)
{
    return character >= L'0' && character <= L'9';
}

int HexValue(wchar_t character)
{
    if (Digit(character))
    {
        return character - L'0';
    }
    if (character >= L'a' && character <= L'f')
    {
        return character - L'a' + 10;
    }
    if (character >= L'A' && character <= L'F')
    {
        return character - L'A' + 10;
    }
    return -1;
}

bool Hex4(const wchar_t** ptr, const wchar_t* end, unsigned* value)
{
    if (end - *ptr < 4)
    {
        return false;
    }
    unsigned result = 0;
    for (int index = 0; index < 4; ++index)
    {
        const int digit = HexValue((*ptr)[index]);
        if (digit < 0)
        {
            return false;
        }
        result = (result << 4) | static_cast<unsigned>(digit);
    }
    *ptr += 4;
    *value = result;
    return true;
}

} // namespace

bool Reader::Fail(const wchar_t* message)
{
    if (error_ && error_->empty())
    {
        *error_ = message;
    }
    ptr_ = end_;
    return false;
}

void Reader::SkipSpace()
{
    while (ptr_ < end_)
    {
        if (*ptr_ == L' ' || *ptr_ == L'\t' || *ptr_ == L'\r' || *ptr_ == L'\n')
        {
            ++ptr_;
            continue;
        }
        if (*ptr_ != L'/' || ptr_ + 1 >= end_)
        {
            return;
        }
        if (ptr_[1] == L'/')
        {
            ptr_ += 2;
            while (ptr_ < end_ && *ptr_ != L'\n')
            {
                ++ptr_;
            }
            continue;
        }
        if (ptr_[1] != L'*')
        {
            return;
        }
        ptr_ += 2;
        while (ptr_ + 1 < end_ && (*ptr_ != L'*' || ptr_[1] != L'/'))
        {
            ++ptr_;
        }
        ptr_ = ptr_ + 1 < end_ ? ptr_ + 2 : end_;
    }
}

wchar_t Reader::Next()
{
    SkipSpace();
    return ptr_ < end_ ? *ptr_ : L'\0';
}

bool Reader::Expect(wchar_t character)
{
    if (Next() != character)
    {
        return Fail(L"The file isn't valid JSON.");
    }
    ++ptr_;
    return true;
}

bool Reader::Open(wchar_t bracket)
{
    if (++depth_ > max_depth_)
    {
        return Fail(L"The file is nested too deeply.");
    }
    return Expect(bracket);
}

bool Reader::Empty(wchar_t bracket)
{
    return Next() == bracket;
}

bool Reader::Close(wchar_t bracket)
{
    --depth_;
    return Expect(bracket);
}

bool Reader::Comma()
{
    if (Next() != L',')
    {
        return false;
    }
    ++ptr_;
    return true;
}

bool Reader::Word(std::wstring_view word)
{
    if (static_cast<size_t>(end_ - ptr_) < word.size() || std::wstring_view(ptr_, word.size()) != word)
    {
        return Fail(L"The file isn't valid JSON.");
    }
    ptr_ += word.size();
    return true;
}

bool Reader::String(std::wstring* out, size_t limit)
{
    if (!Expect(L'"'))
    {
        return false;
    }
    out->clear();
    for (;;)
    {
        const wchar_t* run = ptr_;
        while (ptr_ < end_ && *ptr_ != L'"' && *ptr_ != L'\\' && *ptr_ >= 0x20)
        {
            ++ptr_;
        }
        out->append(run, ptr_);
        if (out->size() > limit)
        {
            return Fail(L"A text member is longer than the format allows.");
        }
        if (ptr_ == end_)
        {
            return Fail(L"The file ends inside a string.");
        }
        const wchar_t character = *ptr_++;
        if (character == L'"')
        {
            return true;
        }
        if (character != L'\\')
        {
            return Fail(L"A string contains an unescaped control character.");
        }
        if (ptr_ == end_)
        {
            return Fail(L"The file ends inside a string.");
        }
        const wchar_t escape = *ptr_++;
        switch (escape)
        {
        case L'"':
        case L'\\':
        case L'/':
            out->push_back(escape);
            break;
        case L'b':
            out->push_back(L'\b');
            break;
        case L'f':
            out->push_back(L'\f');
            break;
        case L'n':
            out->push_back(L'\n');
            break;
        case L'r':
            out->push_back(L'\r');
            break;
        case L't':
            out->push_back(L'\t');
            break;
        case L'u':
            {
                unsigned first = 0;
                if (!Hex4(&ptr_, end_, &first))
                {
                    return Fail(L"A string contains a malformed escape.");
                }
                if (first >= 0xDC00 && first <= 0xDFFF)
                {
                    return Fail(L"A string contains a lone surrogate.");
                }
                out->push_back(static_cast<wchar_t>(first));
                if (first < 0xD800 || first > 0xDBFF)
                {
                    break;
                }
                unsigned second = 0;
                if (end_ - ptr_ < 2 || ptr_[0] != L'\\' || ptr_[1] != L'u' || (ptr_ += 2, !Hex4(&ptr_, end_, &second)) ||
                    second < 0xDC00 || second > 0xDFFF)
                {
                    return Fail(L"A string contains a lone surrogate.");
                }
                out->push_back(static_cast<wchar_t>(second));
                break;
            }
        default:
            return Fail(L"A string contains an unsupported escape.");
        }
    }
}

bool Reader::Unsigned(uint64_t* out)
{
    if (!Digit(Next()))
    {
        return Fail(L"An unsigned number was expected.");
    }
    if (*ptr_ == L'0' && ptr_ + 1 < end_ && Digit(ptr_[1]))
    {
        return Fail(L"A number has a leading zero.");
    }
    uint64_t value = 0;
    for (; ptr_ < end_ && Digit(*ptr_); ++ptr_)
    {
        const uint64_t digit = static_cast<uint64_t>(*ptr_ - L'0');
        if (value > (UINT64_MAX - digit) / 10)
        {
            return Fail(L"A number is out of range.");
        }
        value = value * 10 + digit;
    }
    if (ptr_ < end_ && std::wcschr(L".eE-+", *ptr_))
    {
        return Fail(L"Only unsigned integers are supported.");
    }
    *out = value;
    return true;
}

bool Reader::Skip()
{
    switch (Next())
    {
    case L'{':
        return Object([this](const std::wstring&) { return Skip(); });
    case L'[':
        return Array([this] { return Skip(); });
    case L'"':
        {
            std::wstring ignored;
            return String(&ignored);
        }
    case L't':
        return Word(L"true");
    case L'f':
        return Word(L"false");
    case L'n':
        return Word(L"null");
    default:
        break;
    }
    if (ptr_ < end_ && *ptr_ == L'-')
    {
        ++ptr_;
    }
    const wchar_t* digits = ptr_;
    while (ptr_ < end_ && Digit(*ptr_))
    {
        ++ptr_;
    }
    if (ptr_ == digits || (*digits == L'0' && ptr_ - digits > 1))
    {
        return Fail(L"The file isn't valid JSON.");
    }
    if (ptr_ < end_ && *ptr_ == L'.')
    {
        const wchar_t* fraction = ++ptr_;
        while (ptr_ < end_ && Digit(*ptr_))
        {
            ++ptr_;
        }
        if (ptr_ == fraction)
        {
            return Fail(L"The file isn't valid JSON.");
        }
    }
    if (ptr_ < end_ && (*ptr_ == L'e' || *ptr_ == L'E'))
    {
        ++ptr_;
        if (ptr_ < end_ && (*ptr_ == L'+' || *ptr_ == L'-'))
        {
            ++ptr_;
        }
        const wchar_t* exponent = ptr_;
        while (ptr_ < end_ && Digit(*ptr_))
        {
            ++ptr_;
        }
        if (ptr_ == exponent)
        {
            return Fail(L"The file isn't valid JSON.");
        }
    }
    return true;
}

bool Reader::End()
{
    SkipSpace();
    return ptr_ == end_ || Fail(L"The file contains trailing content.");
}

void AppendString(std::wstring* out, std::wstring_view text)
{
    out->push_back(L'"');
    size_t start = 0;
    for (size_t index = 0; index < text.size(); ++index)
    {
        const wchar_t character = text[index];
        const bool high = character >= 0xD800 && character <= 0xDBFF;
        const bool paired = high && index + 1 < text.size() && text[index + 1] >= 0xDC00 && text[index + 1] <= 0xDFFF;
        if (paired)
        {
            ++index;
            continue;
        }
        if (character >= 0x20 && character != L'"' && character != L'\\' && character != 0x7F && !high &&
            (character < 0xDC00 || character > 0xDFFF))
        {
            continue;
        }
        out->append(text.substr(start, index - start));
        start = index + 1;
        switch (character)
        {
        case L'"':
            out->append(L"\\\"");
            break;
        case L'\\':
            out->append(L"\\\\");
            break;
        case L'\b':
            out->append(L"\\b");
            break;
        case L'\f':
            out->append(L"\\f");
            break;
        case L'\n':
            out->append(L"\\n");
            break;
        case L'\r':
            out->append(L"\\r");
            break;
        case L'\t':
            out->append(L"\\t");
            break;
        default:
            {
                wchar_t buffer[8] = {};
                swprintf_s(buffer, L"\\u%04X", static_cast<unsigned>(character));
                out->append(buffer);
                break;
            }
        }
    }
    out->append(text.substr(start));
    out->push_back(L'"');
}

} // namespace regkit::json
