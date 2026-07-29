// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <regex>
#include <string>

namespace regkit::search {

struct ReplaceOptions {
  std::wstring find_text;
  std::wstring replace_text;
  std::wstring start_key;
  bool recursive = true;
  bool match_case = false;
  bool match_whole = false;
  bool use_regex = false;
};

class Replacer {
public:
  explicit Replacer(const ReplaceOptions& options);

  bool valid() const noexcept;
  bool Replace(const std::wstring& text, std::wstring* result) const;

private:
  std::wstring query_;
  std::wstring replacement_;
  std::wregex regex_;
  bool use_regex_ = false;
  bool match_case_ = false;
  bool match_whole_ = false;
  bool valid_ = true;
};

} // namespace regkit::search
