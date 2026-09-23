// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include "win32/windows_config.h"

#include <windows.h>

#include <functional>
#include <string>

#include "work/session.h"

namespace regkit::frame
{

struct UpdateCheckPayload : work::MoveOnly
{
    bool silent = false;
    bool failed = false;
    std::wstring version;
    std::wstring download_url;
    std::string sha256;
    std::wstring setup_path;
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

  private:
    void Download(const UpdateCheckPayload& release);
    void SetStatus(const std::wstring& text) const;

    HWND owner_ = nullptr;
    StatusCallback status_;
    bool running_ = false;
    work::Session session_;
};

} // namespace regkit::frame
