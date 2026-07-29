// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <string>
#include <unordered_map>
#include <vector>

namespace regkit::changes {

struct CommentEntry {
  std::wstring path;
  std::wstring name;
  DWORD type = 0;
  std::wstring text;
};

struct CommentDocument {
  static constexpr int kCurrentVersion = 1;
  int source_version = 1;
  std::vector<CommentEntry> value_entries;
  std::vector<CommentEntry> name_entries;
};

class ValueComments {
public:
  bool Load(const std::wstring& path);
  bool Save(const std::wstring& path) const;
  bool Import(const std::wstring& path);
  bool Export(const std::wstring& path) const;
  void Clear();
  void Merge(const CommentDocument& document);

  const std::unordered_map<std::wstring, CommentEntry>&
  value_entries() const noexcept;
  std::unordered_map<std::wstring, CommentEntry>& value_entries() noexcept;
  const std::unordered_map<std::wstring, CommentEntry>&
  name_entries() const noexcept;
  std::unordered_map<std::wstring, CommentEntry>& name_entries() noexcept;

  static std::wstring ValueKey(const std::wstring& path,
                               const std::wstring& name, DWORD type);
  static std::wstring NameKey(const std::wstring& name, DWORD type);

private:
  std::unordered_map<std::wstring, CommentEntry> value_entries_;
  std::unordered_map<std::wstring, CommentEntry> name_entries_;
};

CommentDocument ParseComments(const std::wstring& content);
std::wstring SerializeComments(const ValueComments& comments);

} // namespace regkit::changes
