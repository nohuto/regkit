// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <string>

#include "search/search.h"

namespace regkit::search
{

struct ReplaceOptions
{
    std::wstring find_text;
    std::wstring replace_text;
    std::wstring start_key;
    bool recursive = true;
    bool match_case = false;
    bool match_whole = false;
    bool use_regex = false;
    bool replace_keys = false;
    bool replace_values = true;
    bool replace_data = true;
    bool number_decimal = true;
    bool number_hex = false;
};

class Replacer
{
  public:
    explicit Replacer(const ReplaceOptions& options)
        : matcher_({options.find_text, options.match_case, options.match_whole, options.use_regex}), replacement_(options.replace_text)
    {
    }

    bool valid() const noexcept
    {
        return matcher_.valid();
    }
    const regex::Error& error() const noexcept
    {
        return matcher_.error();
    }
    regex::Status Replace(const std::wstring& text, std::wstring* result) const
    {
        return matcher_.Replace(text, replacement_, result);
    }

  private:
    Matcher matcher_;
    std::wstring replacement_;
};

}
