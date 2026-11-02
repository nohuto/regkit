// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <atomic>
#include <string>
#include <vector>

namespace regkit::win32
{

struct KeyHandle
{
    ULONG_PTR process_id = 0;
    ULONG_PTR handle = 0;
    ULONG_PTR object = 0;
    ACCESS_MASK access = 0;
    ULONG attributes = 0;
    std::wstring process;
    std::wstring name;
};

struct KeyHandleSnapshot
{
    std::vector<KeyHandle> handles;
    size_t inaccessible_processes = 0;
    DWORD error = ERROR_SUCCESS;
};

KeyHandleSnapshot SnapshotKeyHandles(const std::atomic_bool& cancel);
DWORD CloseKeyHandles(const std::vector<KeyHandle>& targets, size_t* closed);

}
