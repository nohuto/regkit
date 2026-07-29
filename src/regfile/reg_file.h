// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "registry/registry_value.h"
#include "win32/windows_config.h"

#include <windows.h>

#include <atomic>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace regkit::regfile {

using Value = RegistryValue;

struct Key {
  std::wstring path;
  std::unordered_map<std::wstring, Value> values;
};

struct Document {
  std::vector<std::wstring> key_order;
  std::unordered_map<std::wstring, Key> keys;
};

class Writer {
public:
  Writer();

  void AppendKey(std::wstring_view path,
                 std::vector<const Value*> values);
  std::wstring Finish() &&;

private:
  std::wstring output_;
};

bool Parse(std::wstring_view content, Document* output,
           const std::atomic_bool* cancel = nullptr, bool* cancelled = nullptr);
bool Load(const std::wstring& path, Document* output, std::wstring* error,
          const std::atomic_bool* cancel = nullptr, bool* cancelled = nullptr);
std::wstring Serialize(const Document& document);

} // namespace regkit::regfile
