// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/system/paths.h>
#include <win32ui/windows_config.h>

#include <string>

namespace regkit::util
{

std::wstring GetAppDataFolder();
std::wstring GetCacheFolder();

using win32ui::EnsureFileExtension;
using win32ui::FileBaseName;
using win32ui::FileName;
using win32ui::ForEachFile;
using win32ui::HasFileExtension;
using win32ui::IsDirectory;
using win32ui::IsFile;
using win32ui::IsMissing;
using win32ui::JoinPath;
using win32ui::ModuleDirectory;
using win32ui::ModulePath;
using win32ui::TrimTrailingSeparators;

}
