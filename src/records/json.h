// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <cstdint>
#include <string>
#include <string_view>

namespace regkit::json
{

class Reader
{
  public:
    Reader(std::wstring_view text, std::wstring* error, int max_depth = 16)
        : ptr_(text.data()), end_(text.data() + text.size()), error_(error), max_depth_(max_depth)
    {
    }

    template <typename Member>
    bool Object(Member&& member)
    {
        if (!Open(L'{'))
        {
            return false;
        }
        if (!Empty(L'}'))
        {
            std::wstring key;
            do
            {
                if (!String(&key, 256) || !Expect(L':') || !member(static_cast<const std::wstring&>(key)))
                {
                    return false;
                }
            } while (Comma() && !Empty(L'}'));
        }
        return Close(L'}');
    }

    template <typename Element>
    bool Array(Element&& element)
    {
        if (!Open(L'['))
        {
            return false;
        }
        if (!Empty(L']'))
        {
            do
            {
                if (!element())
                {
                    return false;
                }
            } while (Comma() && !Empty(L']'));
        }
        return Close(L']');
    }

    bool String(std::wstring* out, size_t limit = SIZE_MAX);
    bool Unsigned(uint64_t* out);
    bool Skip();
    bool End();
    bool Fail(const wchar_t* message);
    wchar_t Next();

  private:
    bool Open(wchar_t bracket);
    bool Empty(wchar_t bracket);
    bool Close(wchar_t bracket);
    bool Comma();
    bool Expect(wchar_t character);
    bool Word(std::wstring_view word);
    void SkipSpace();

    const wchar_t* ptr_ = nullptr;
    const wchar_t* end_ = nullptr;
    std::wstring* error_ = nullptr;
    int max_depth_ = 16;
    int depth_ = 0;
};

void AppendString(std::wstring* out, std::wstring_view text);

} // namespace regkit::json
