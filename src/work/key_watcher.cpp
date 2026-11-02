// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "work/key_watcher.h"

#include "win32/registry_native.h"
#include "win32/registry_view.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/worker.h>

namespace regkit::work
{

KeyWatcher::~KeyWatcher()
{
    if (thread_.joinable())
    {
        SetEvent(stop_.get());
        thread_.join();
    }
}

void KeyWatcher::Watch(HWND window, UINT message, HKEY root, const std::wstring& subkey, REGSAM view)
{
    {
        std::lock_guard lock(mutex_);
        if (thread_.joinable() && root_ == root && subkey_ == subkey && view_ == view)
        {
            return;
        }
        window_ = window;
        message_ = message;
        root_ = root;
        subkey_ = subkey;
        view_ = view;
        generation_.fetch_add(1);
    }
    if (!thread_.joinable())
    {
        stop_.reset(CreateEventW(nullptr, TRUE, FALSE, nullptr));
        retarget_.reset(CreateEventW(nullptr, FALSE, TRUE, nullptr));
        rearm_.reset(CreateEventW(nullptr, FALSE, FALSE, nullptr));
        if (!stop_ || !retarget_ || !rearm_)
        {
            return;
        }
        // without a thread the next Watch tries again
        try
        {
            thread_ = std::thread([this] { Run(); });
        }
        catch (const std::exception&)
        {
            return;
        }
        win32ui::NameThread(thread_.native_handle(), L"KeyWatcherThread");
        return;
    }
    SetEvent(retarget_.get());
}

void KeyWatcher::Rearm()
{
    if (rearm_)
    {
        SetEvent(rearm_.get());
    }
}

void KeyWatcher::Stop()
{
    if (thread_.joinable())
    {
        Watch(nullptr, 0, nullptr, {}, 0);
    }
}

void KeyWatcher::Run()
{
    // registrations belong to this thread, keep it alive as long as the window
    win32ui::UniqueHandle change(CreateEventW(nullptr, TRUE, FALSE, nullptr));
    win32ui::UniqueHKey key;
    HWND window = nullptr;
    UINT message = 0;
    uint64_t generation = 0;
    bool armed = false;
    const HANDLE events[] = {stop_.get(), retarget_.get(), change.get(), rearm_.get()};
    constexpr DWORD kFilter = REG_NOTIFY_CHANGE_NAME | REG_NOTIFY_CHANGE_ATTRIBUTES | REG_NOTIFY_CHANGE_LAST_SET | REG_NOTIFY_CHANGE_SECURITY;
    while (change)
    {
        const DWORD signaled = WaitForMultipleObjects(4, events, FALSE, INFINITE);
        if (signaled == WAIT_OBJECT_0 + 1)
        {
            HKEY root = nullptr;
            std::wstring subkey;
            REGSAM view = 0;
            {
                std::lock_guard lock(mutex_);
                window = window_;
                message = message_;
                root = root_;
                subkey = subkey_;
                view = view_;
                generation = generation_.load();
            }
            // closing the old key signals its registration
            key.reset();
            armed = false;
            ResetEvent(change.get());
            if (root)
            {
                util::OpenRegistryPath(root, subkey, KEY_NOTIFY | view, false, &key);
            }
        }
        else if (signaled == WAIT_OBJECT_0 + 2)
        {
            ResetEvent(change.get());
            PostMessageW(window, message, static_cast<WPARAM>(generation), 0);
            armed = false;
            continue;
        }
        else if (signaled != WAIT_OBJECT_0 + 3)
        {
            break;
        }
        if (key && !armed)
        {
            armed = RegNotifyChangeKeyValue(key.get(), FALSE, kFilter, change.get(), TRUE) == ERROR_SUCCESS;
            if (!armed)
            {
                key.reset();
            }
        }
    }
}

}
