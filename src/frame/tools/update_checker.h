// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/system/worker.h>
#include <win32ui/windows_config.h>

#include <windows.h>

#include <functional>
#include <string>

namespace regkit::frame
{

struct UpdateCheckPayload : win32ui::PostedPayload
{
    bool silent = false;
    bool failed = false;
    bool installed = false;
    bool setup_exited = false;
    std::wstring version;
    std::wstring download_url;
    std::wstring file_name;
    std::string sha256;
    std::wstring download_path;
    std::wstring error;
};

class UpdateChecker
{
  public:
    using StatusCallback = std::function<void(const std::wstring&)>;

    void Attach(HWND owner, StatusCallback status);
    void Check(bool silent);
    void Apply(UpdateCheckPayload* payload);
    void Cancel();

    bool SetupPending() const
    {
        return !setup_path_.empty();
    }
    bool SetupRunning() const
    {
        return setup_running_;
    }
    bool StartSetup();
    void DropSetup();

  private:
    void Download(const UpdateCheckPayload& release);
    void SetStatus(const std::wstring& text) const;

    HWND owner_ = nullptr;
    StatusCallback status_;
    win32ui::Session session_;
    std::wstring setup_path_;
    std::string setup_sha256_;
    bool setup_running_ = false;
};

}
