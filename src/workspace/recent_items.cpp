// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "workspace/recent_items.h"

#include "win32/win32_helpers.h"

#include <algorithm>

namespace regkit::workspace {

RecentItems::RecentItems(size_t maximum) noexcept : maximum_(maximum) {}

void RecentItems::Add(const std::wstring& path) {
  std::wstring cleaned = util::TrimWhitespace(path);
  if (cleaned.empty()) {
    return;
  }
  const auto existing =
      std::find_if(items_.begin(), items_.end(),
                   [&](const std::wstring& item) {
                     return _wcsicmp(item.c_str(), cleaned.c_str()) == 0;
                   });
  if (existing != items_.end()) {
    items_.erase(existing);
  }
  items_.insert(items_.begin(), std::move(cleaned));
  if (items_.size() > maximum_) {
    items_.resize(maximum_);
  }
}

void RecentItems::Replace(std::vector<std::wstring> paths) {
  items_ = std::move(paths);
  Normalize();
}

void RecentItems::Normalize() {
  std::vector<std::wstring> normalized;
  normalized.reserve(std::min(items_.size(), maximum_));
  for (const std::wstring& item : items_) {
    std::wstring cleaned = util::TrimWhitespace(item);
    if (cleaned.empty()) {
      continue;
    }
    const bool duplicate =
        std::any_of(normalized.begin(), normalized.end(),
                    [&](const std::wstring& existing) {
                      return _wcsicmp(existing.c_str(), cleaned.c_str()) == 0;
                    });
    if (!duplicate) {
      normalized.push_back(std::move(cleaned));
      if (normalized.size() == maximum_) {
        break;
      }
    }
  }
  items_.swap(normalized);
}

const std::vector<std::wstring>& RecentItems::items() const noexcept {
  return items_;
}

std::vector<std::wstring>& RecentItems::items() noexcept {
  return items_;
}

} // namespace regkit::workspace
