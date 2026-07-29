// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "search/replace.h"

#include <windows.h>

#include <utility>

namespace regkit::search {

Replacer::Replacer(const ReplaceOptions& options)
    : query_(options.find_text), replacement_(options.replace_text),
      use_regex_(options.use_regex), match_case_(options.match_case),
      match_whole_(options.match_whole), valid_(!query_.empty()) {
  if (!valid_ || !use_regex_) {
    return;
  }
  try {
    auto flags = std::regex_constants::ECMAScript;
    if (!match_case_) {
      flags |= std::regex_constants::icase;
    }
    regex_ = std::wregex(query_, flags);
  } catch (const std::regex_error&) {
    valid_ = false;
  }
}

bool Replacer::valid() const noexcept {
  return valid_;
}

bool Replacer::Replace(const std::wstring& text,
                       std::wstring* result) const {
  if (!result || !valid_) {
    return false;
  }
  if (use_regex_) {
    const bool matched =
        match_whole_ ? std::regex_match(text, regex_)
                     : std::regex_search(text, regex_);
    if (!matched) {
      return false;
    }
    *result = std::regex_replace(text, regex_, replacement_);
    return true;
  }

  if (match_whole_) {
    const bool matched =
        match_case_
            ? text == query_
            : CompareStringOrdinal(
                  text.c_str(), static_cast<int>(text.size()),
                  query_.c_str(), static_cast<int>(query_.size()),
                  TRUE) == CSTR_EQUAL;
    if (!matched) {
      return false;
    }
    *result = replacement_;
    return true;
  }

  if (match_case_) {
    size_t position = text.find(query_);
    if (position == std::wstring::npos) {
      return false;
    }
    std::wstring replaced;
    size_t cursor = 0;
    while (position != std::wstring::npos) {
      replaced.append(text, cursor, position - cursor);
      replaced.append(replacement_);
      cursor = position + query_.size();
      position = text.find(query_, cursor);
    }
    replaced.append(text, cursor, std::wstring::npos);
    *result = std::move(replaced);
    return true;
  }

  size_t cursor = 0;
  std::wstring replaced;
  bool matched = false;
  while (cursor < text.size()) {
    const int position = FindStringOrdinal(
        FIND_FROMSTART, text.c_str() + cursor,
        static_cast<int>(text.size() - cursor), query_.c_str(),
        static_cast<int>(query_.size()), TRUE);
    if (position < 0) {
      break;
    }
    const size_t match = cursor + static_cast<size_t>(position);
    replaced.append(text, cursor, match - cursor);
    replaced.append(replacement_);
    cursor = match + query_.size();
    matched = true;
  }
  if (!matched) {
    return false;
  }
  replaced.append(text, cursor, std::wstring::npos);
  *result = std::move(replaced);
  return true;
}

} // namespace regkit::search
