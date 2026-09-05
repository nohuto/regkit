// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/window_detail.h"

#include <winhttp.h>

namespace regkit {
using namespace window_detail;

namespace {

constexpr wchar_t kReleasesPage[] = L"https://github.com/nohuto/regkit/releases";
constexpr wchar_t kApiHost[] = L"api.github.com";
constexpr wchar_t kApiPath[] = L"/repos/nohuto/regkit/releases/latest";

std::string HttpGet(const wchar_t* host, const wchar_t* path) {
  std::string body;
  HINTERNET session =
      WinHttpOpen(L"RegKit", WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
                  WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
  if (!session) {
    session = WinHttpOpen(L"RegKit", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                          WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
  }
  if (!session) {
    return body;
  }
  HINTERNET connect =
      WinHttpConnect(session, host, INTERNET_DEFAULT_HTTPS_PORT, 0);
  if (connect) {
    HINTERNET request = WinHttpOpenRequest(
        connect, L"GET", path, nullptr, WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
    if (request) {
      const wchar_t headers[] =
          L"Accept: application/vnd.github+json\r\n"
          L"X-GitHub-Api-Version: 2022-11-28\r\n";
      if (WinHttpSendRequest(request, headers, static_cast<DWORD>(-1),
                             WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
          WinHttpReceiveResponse(request, nullptr)) {
        DWORD status = 0;
        DWORD status_size = sizeof(status);
        WinHttpQueryHeaders(
            request, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            WINHTTP_HEADER_NAME_BY_INDEX, &status, &status_size,
            WINHTTP_NO_HEADER_INDEX);
        if (status == 200) {
          DWORD available = 0;
          while (WinHttpQueryDataAvailable(request, &available) &&
                 available > 0) {
            const size_t offset = body.size();
            body.resize(offset + available);
            DWORD read = 0;
            if (!WinHttpReadData(request, body.data() + offset, available,
                                 &read)) {
              break;
            }
            body.resize(offset + read);
          }
        }
      }
      WinHttpCloseHandle(request);
    }
    WinHttpCloseHandle(connect);
  }
  WinHttpCloseHandle(session);
  return body;
}

std::string JsonString(const std::string& json, const char* key,
                       size_t from = 0) {
  const std::string needle = std::string("\"") + key + "\"";
  size_t pos = json.find(needle, from);
  if (pos == std::string::npos) {
    return {};
  }
  pos = json.find(':', pos + needle.size());
  if (pos == std::string::npos) {
    return {};
  }
  pos = json.find('"', pos);
  if (pos == std::string::npos) {
    return {};
  }
  const size_t end = json.find('"', pos + 1);
  if (end == std::string::npos) {
    return {};
  }
  return json.substr(pos + 1, end - pos - 1);
}

std::vector<int> VersionParts(const std::wstring& text) {
  std::vector<int> parts;
  int value = 0;
  bool digits = false;
  for (wchar_t ch : text) {
    if (ch >= L'0' && ch <= L'9') {
      value = value * 10 + (ch - L'0');
      digits = true;
      continue;
    }
    if (digits) {
      parts.push_back(value);
      value = 0;
      digits = false;
    }
    if (ch != L'.' && !parts.empty()) {
      break;
    }
  }
  if (digits) {
    parts.push_back(value);
  }
  return parts;
}

bool IsNewerVersion(const std::wstring& candidate, const std::wstring& current) {
  const std::vector<int> left = VersionParts(candidate);
  const std::vector<int> right = VersionParts(current);
  const size_t count = std::max(left.size(), right.size());
  for (size_t i = 0; i < count; ++i) {
    const int a = i < left.size() ? left[i] : 0;
    const int b = i < right.size() ? right[i] : 0;
    if (a != b) {
      return a > b;
    }
  }
  return false;
}

bool AssetMatchesArchitecture(const std::string& name) {
  std::string lower = name;
  for (char& ch : lower) {
    ch = static_cast<char>(tolower(static_cast<unsigned char>(ch)));
  }
  if (lower.find(".exe") == std::string::npos &&
      lower.find(".msi") == std::string::npos) {
    return false;
  }
  const bool wants_64 = sizeof(void*) == 8;
  const bool has_64 = lower.find("x64") != std::string::npos ||
                      lower.find("amd64") != std::string::npos ||
                      lower.find("win64") != std::string::npos ||
                      lower.find("64-bit") != std::string::npos;
  const bool has_32 = lower.find("x86") != std::string::npos ||
                      lower.find("win32") != std::string::npos ||
                      lower.find("32-bit") != std::string::npos;
  if (wants_64) {
    return has_64 || (!has_32 && !has_64);
  }
  return has_32 || (!has_32 && !has_64);
}

std::wstring PickAssetUrl(const std::string& json) {
  std::wstring fallback;
  size_t pos = json.find("\"assets\"");
  if (pos == std::string::npos) {
    return fallback;
  }
  while (true) {
    const size_t name_pos = json.find("\"name\"", pos);
    if (name_pos == std::string::npos) {
      break;
    }
    const std::string name = JsonString(json, "name", name_pos);
    const std::string url = JsonString(json, "browser_download_url", name_pos);
    if (url.empty()) {
      break;
    }
    if (AssetMatchesArchitecture(name)) {
      return util::Utf8ToWide(url);
    }
    if (fallback.empty()) {
      fallback = util::Utf8ToWide(url);
    }
    pos = json.find("browser_download_url", name_pos);
    if (pos == std::string::npos) {
      break;
    }
    pos += 20;
  }
  return fallback;
}

} // namespace

void MainWindow::Impl::CheckForUpdates(bool silent) {
  if (update_check_running_) {
    return;
  }
  update_check_running_ = true;
  HWND owner = hwnd_;
  std::thread([owner, silent] {
    const std::string json = HttpGet(kApiHost, kApiPath);
    auto payload = std::make_unique<UpdateCheckPayload>();
    payload->silent = silent;
    if (json.empty()) {
      payload->failed = true;
    } else {
      payload->version = util::Utf8ToWide(JsonString(json, "tag_name"));
      payload->download_url = PickAssetUrl(json);
    }
    if (PostMessageW(owner, frame::message_id::kUpdateCheckReady, 0,
                     reinterpret_cast<LPARAM>(payload.get()))) {
      ReleasePostedPayload(payload);
    }
  }).detach();
}

void MainWindow::Impl::ApplyUpdateCheckResult(UpdateCheckPayload* payload) {
  update_check_running_ = false;
  if (!payload) {
    return;
  }
  if (payload->failed) {
    if (!payload->silent) {
      ui::ShowError(hwnd_, L"Failed to contact the update server.");
    }
    return;
  }
  if (!IsNewerVersion(payload->version, REGKIT_VERSION_STR_W)) {
    if (!payload->silent) {
      ui::ShowInfo(hwnd_, L"RegKit is up to date.");
    }
    return;
  }
  std::wstring message = L"RegKit ";
  message += payload->version;
  message += L" is available. You are running ";
  message += REGKIT_VERSION_STR_W;
  message += L".\n\nDownload it now?";
  if (ui::PromptChoice(hwnd_, message, L"Update available", L"Download",
                       L"", L"Close") == IDYES) {
    const std::wstring target = payload->download_url.empty()
                                    ? std::wstring(kReleasesPage)
                                    : payload->download_url;
    win32::ShellOpen(hwnd_, target.c_str());
  }
}

} // namespace regkit
