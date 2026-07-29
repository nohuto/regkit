// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace regkit::workspace {

class RecentItems {
public:
  explicit RecentItems(size_t maximum = 16) noexcept;

  void Add(const std::wstring& path);
  void Replace(std::vector<std::wstring> paths);
  void Normalize();

  const std::vector<std::wstring>& items() const noexcept;
  std::vector<std::wstring>& items() noexcept;

private:
  size_t maximum_;
  std::vector<std::wstring> items_;
};

} // namespace regkit::workspace
