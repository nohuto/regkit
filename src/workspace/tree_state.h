// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#pragma once

#include <string>
#include <vector>

namespace regkit::workspace {

struct TreeState {
  static constexpr int kCurrentVersion = 1;
  int source_version = 1;
  std::wstring selected_path;
  std::vector<std::wstring> expanded_paths;

  void Clear();
  void Normalize();
};

TreeState ParseTreeState(const std::wstring& content);
std::wstring SerializeTreeState(const TreeState& state);
bool LoadTreeState(const std::wstring& path, TreeState* state);
bool SaveTreeState(const std::wstring& path, const TreeState& state);

} // namespace regkit::workspace
