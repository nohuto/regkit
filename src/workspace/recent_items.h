// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

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
