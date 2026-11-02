// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/handle_owner.h>
#include <win32ui/windows_config.h>

#include <atomic>
#include <cstdint>
#include <mutex>
#include <string>
#include <thread>

namespace regkit::work
{

class KeyWatcher
{
  public:
    KeyWatcher() = default;
    ~KeyWatcher();
    KeyWatcher(const KeyWatcher&) = delete;
    KeyWatcher& operator=(const KeyWatcher&) = delete;

    void Watch(HWND window, UINT message, HKEY root, const std::wstring& subkey, REGSAM view);
    void Rearm();
    void Stop();
    uint64_t generation() const noexcept
    {
        return generation_.load();
    }

  private:
    void Run();

    std::thread thread_;
    std::mutex mutex_;
    HWND window_ = nullptr;
    UINT message_ = 0;
    HKEY root_ = nullptr;
    std::wstring subkey_;
    REGSAM view_ = 0;
    std::atomic<uint64_t> generation_{0};
    win32ui::UniqueHandle stop_;
    win32ui::UniqueHandle retarget_;
    win32ui::UniqueHandle rearm_;
};

}
