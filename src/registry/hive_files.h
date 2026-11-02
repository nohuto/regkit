// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "registry/registry_store.h"

#include <string>
#include <vector>

namespace regkit
{

struct OfflineHiveCandidate
{
    std::wstring path;
    std::wstring label;
};

bool IsHiveFile(const std::wstring& path);
std::wstring NormalizeHiveFilePath(const std::wstring& raw_path);
void CollectOfflineHivesInFolder(const std::wstring& folder, std::vector<OfflineHiveCandidate>* out);
std::wstring ResolveOfflineRootName(const std::wstring& path, bool is_dir, const RegistryNode* current_node);

}
