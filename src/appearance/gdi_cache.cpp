// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "appearance/gdi_cache.h"

#include <array>
#include <type_traits>

namespace regkit::appearance {
namespace {

template <typename Handle, size_t Capacity>
class ObjectCache {
public:
  ~ObjectCache() {
    for (const Entry& entry : entries_) {
      if (entry.handle) {
        DeleteObject(entry.handle);
      }
    }
  }

  Handle Brush(COLORREF color)
    requires std::is_same_v<Handle, HBRUSH>
  {
    return FindOrCreate(color, 0, [color] { return CreateSolidBrush(color); });
  }

  Handle Pen(COLORREF color, int width)
    requires std::is_same_v<Handle, HPEN>
  {
    return FindOrCreate(color, width, [color, width] { return CreatePen(PS_SOLID, width, color); });
  }

private:
  struct Entry {
    COLORREF color = CLR_INVALID;
    int width = 0;
    Handle handle = nullptr;
  };

  template <typename Factory>
  Handle FindOrCreate(COLORREF color, int width, Factory&& factory) {
    for (const Entry& entry : entries_) {
      if (entry.handle && entry.color == color && entry.width == width) {
        return entry.handle;
      }
    }

    for (Entry& entry : entries_) {
      if (!entry.handle) {
        entry = {color, width, factory()};
        return entry.handle;
      }
    }

    Entry& entry = entries_[next_];
    DeleteObject(entry.handle);
    entry = {color, width, factory()};
    next_ = (next_ + 1) % Capacity;
    return entry.handle;
  }

  std::array<Entry, Capacity> entries_{};
  size_t next_ = 0;
};

ObjectCache<HBRUSH, 16> g_brushes;
ObjectCache<HPEN, 16> g_pens;

} // namespace

HBRUSH CachedBrush(COLORREF color) {
  return g_brushes.Brush(color);
}

HPEN CachedPen(COLORREF color, int width) {
  return g_pens.Pen(color, width);
}

} // namespace regkit::appearance
