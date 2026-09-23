// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/update_checker.h"

#include "frame/message_ids.h"
#include "records/json.h"
#include "win32/file_dialog.h"
#include "win32/file_text.h"
#include "win32/handle_owner.h"
#include "win32/system_error.h"
#include "appearance/feedback.h"
#include "resource.h"

#include <bcrypt.h>
#include <winhttp.h>

#include <algorithm>
#include <array>
#include <stdlib.h>

namespace regkit::frame
{

namespace
{

constexpr wchar_t kReleasesPage[] = L"https://github.com/nohuto/regkit/releases";
constexpr wchar_t kLatestReleaseUrl[] = L"https://api.github.com/repos/nohuto/regkit/releases/latest";
constexpr const wchar_t* kSetupSuffix = sizeof(void*) == 8 ? L"-x64.exe" : L"-x86.exe";
constexpr size_t kMaxReleaseJsonBytes = 4ull * 1024ull * 1024ull;
constexpr size_t kMaxSetupBytes = 192ull * 1024ull * 1024ull;

bool IsAllowedReleaseHost(const std::wstring& host)
{
    constexpr const wchar_t* kHosts[] = {L"github.com", L"api.github.com", L"objects.githubusercontent.com", L"release-assets.githubusercontent.com"};
    return std::any_of(std::begin(kHosts), std::end(kHosts), [&](const wchar_t* allowed) { return util::EqualsInsensitive(host, allowed); });
}

std::wstring ErrorText(DWORD code)
{
    if (code < WINHTTP_ERROR_BASE || code > WINHTTP_ERROR_LAST)
    {
        return util::FormatWin32Error(code);
    }
    wchar_t text[512] = {};
    DWORD length =
        FormatMessageW(FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_IGNORE_INSERTS, GetModuleHandleW(L"winhttp.dll"), code, 0, text, static_cast<DWORD>(_countof(text)), nullptr);
    while (length > 0 && iswspace(text[length - 1]))
    {
        --length;
    }
    return length > 0 ? std::wstring(text, length) : L"WinHTTP error " + std::to_wstring(code) + L".";
}

std::wstring HttpGet(const std::wstring& url, const std::atomic_bool& cancel, std::string* body, size_t max_bytes)
{
    using Handle = std::unique_ptr<void, decltype(&WinHttpCloseHandle)>;
    URL_COMPONENTS parts = {sizeof(parts)};
    parts.dwHostNameLength = static_cast<DWORD>(-1);
    parts.dwUrlPathLength = static_cast<DWORD>(-1);
    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &parts))
    {
        return ErrorText(GetLastError());
    }
    if (parts.nScheme != INTERNET_SCHEME_HTTPS ||
        !IsAllowedReleaseHost(std::wstring(parts.lpszHostName, parts.dwHostNameLength)))
    {
        return L"The release location isn't a trusted RegKit download address.";
    }
    Handle session(
        WinHttpOpen(L"RegKit", WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0),
        WinHttpCloseHandle
    );
    if (!session)
    {
        session.reset(WinHttpOpen(L"RegKit", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
    }
    Handle connect(session ? WinHttpConnect(session.get(), std::wstring(parts.lpszHostName, parts.dwHostNameLength).c_str(), parts.nPort, 0) : nullptr, WinHttpCloseHandle);
    Handle request(connect ? WinHttpOpenRequest(connect.get(), L"GET", std::wstring(parts.lpszUrlPath, parts.dwUrlPathLength).c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE) : nullptr, WinHttpCloseHandle);
    DWORD status = 0;
    DWORD status_size = sizeof(status);
    if (!request || !WinHttpSetTimeouts(request.get(), 5000, 10000, 10000, 30000) ||
        !WinHttpSendRequest(request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(request.get(), nullptr) ||
        !WinHttpQueryHeaders(request.get(), WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &status_size, WINHTTP_NO_HEADER_INDEX))
    {
        return ErrorText(GetLastError());
    }
    if (status != HTTP_STATUS_OK)
    {
        return L"The server returned HTTP " + std::to_wstring(status) + L".";
    }
    char buffer[64 * 1024];
    DWORD read = 0;
    while (!cancel.load())
    {
        if (!WinHttpReadData(request.get(), buffer, sizeof(buffer), &read))
        {
            return ErrorText(GetLastError());
        }
        if (read == 0)
        {
            return {};
        }
        if (body->size() + read > max_bytes)
        {
            return L"The download is larger than RegKit accepts.";
        }
        body->append(buffer, read);
    }
    return ErrorText(ERROR_CANCELLED);
}

bool ReadRelease(const std::string& body, std::wstring* version, std::wstring* url, std::wstring* digest)
{
    const std::wstring text = util::Utf8ToWide(body);
    json::Reader reader(text, nullptr);
    const std::wstring_view suffix(kSetupSuffix);
    return reader.Object([&](const std::wstring& member) {
        if (member == L"tag_name")
        {
            return reader.String(version, 64);
        }
        if (member != L"assets")
        {
            return reader.Skip();
        }
        return reader.Array([&] {
            std::wstring name;
            std::wstring link;
            std::wstring hash;
            const bool read = reader.Object([&](const std::wstring& field) {
                std::wstring* out = field == L"name"                   ? &name
                                    : field == L"browser_download_url" ? &link
                                    : field == L"digest"               ? &hash
                                                                       : nullptr;
                return out && reader.Next() == L'"' ? reader.String(out) : reader.Skip();
            });
            if (read && url->empty() && name.starts_with(L"RegKit-Setup-") && name.size() > suffix.size() &&
                name.ends_with(suffix))
            {
                *url = std::move(link);
                *digest = std::move(hash);
            }
            return read;
        });
    }) && reader.End();
}

std::array<int, 4> VersionParts(const std::wstring& text)
{
    std::array<int, 4> parts = {};
    const size_t start = text.find_first_of(L"0123456789");
    if (start != std::wstring::npos)
    {
        swscanf_s(text.c_str() + start, L"%d.%d.%d.%d", &parts[0], &parts[1], &parts[2], &parts[3]);
    }
    return parts;
}

std::string Sha256(const std::string& data)
{
    BCRYPT_ALG_HANDLE algorithm = nullptr;
    BCRYPT_HASH_HANDLE hash = nullptr;
    UCHAR digest[32] = {};
    const bool hashed = BCRYPT_SUCCESS(BCryptOpenAlgorithmProvider(&algorithm, BCRYPT_SHA256_ALGORITHM, nullptr, 0)) &&
                        BCRYPT_SUCCESS(BCryptCreateHash(algorithm, &hash, nullptr, 0, nullptr, 0, 0)) &&
                        BCRYPT_SUCCESS(BCryptHashData(hash, reinterpret_cast<PUCHAR>(const_cast<char*>(data.data())), static_cast<ULONG>(data.size()), 0)) &&
                        BCRYPT_SUCCESS(BCryptFinishHash(hash, digest, sizeof(digest), 0));
    if (hash)
    {
        BCryptDestroyHash(hash);
    }
    if (algorithm)
    {
        BCryptCloseAlgorithmProvider(algorithm, 0);
    }
    std::string hex;
    for (size_t i = 0; hashed && i < sizeof(digest); ++i)
    {
        char pair[3] = {};
        sprintf_s(pair, "%02x", digest[i]);
        hex += pair;
    }
    return hex;
}

std::wstring RandomName(const wchar_t* prefix)
{
    // RandomFileSuffix returns a dotted suffix, which a name doesn't want
    return prefix + util::RandomFileSuffix(L"").substr(1);
}

std::wstring PrivateTempDirectory(std::wstring* directory)
{
    const HMODULE kernel = GetModuleHandleW(L"kernel32.dll");
    using GetTempPath2Fn = DWORD(WINAPI*)(DWORD, LPWSTR);
    const auto temp_path2 =
        kernel ? reinterpret_cast<GetTempPath2Fn>(GetProcAddress(kernel, "GetTempPath2W")) : nullptr;
    wchar_t temp[MAX_PATH + 1] = {};
    const DWORD length = temp_path2 ? temp_path2(static_cast<DWORD>(_countof(temp)), temp)
                                    : GetTempPathW(static_cast<DWORD>(_countof(temp)), temp);
    if (length == 0)
    {
        return util::FormatWin32Error(GetLastError());
    }
    for (int attempt = 0; attempt < 8; ++attempt)
    {
        const std::wstring candidate = util::JoinPath(temp, RandomName(L"RegKit-"));
        if (CreateDirectoryW(candidate.c_str(), nullptr))
        {
            *directory = candidate;
            return {};
        }
        if (GetLastError() != ERROR_ALREADY_EXISTS)
        {
            return util::FormatWin32Error(GetLastError());
        }
    }
    return util::FormatWin32Error(ERROR_ALREADY_EXISTS);
}

std::wstring SaveSetup(const std::wstring& url, const std::string& sha256, const std::atomic_bool& cancel, std::wstring* path)
{
    if (sha256.size() != 64 || sha256.find_first_not_of("0123456789abcdefABCDEF") != std::string::npos)
    {
        return L"The release didn't publish a checksum for this download.";
    }
    std::string data;
    const std::wstring error = HttpGet(url, cancel, &data, kMaxSetupBytes);
    if (!error.empty())
    {
        return error;
    }
    if (data.empty() || _stricmp(Sha256(data).c_str(), sha256.c_str()) != 0)
    {
        return L"The downloaded file doesn't match the release checksum.";
    }
    std::wstring directory;
    const std::wstring directory_error = PrivateTempDirectory(&directory);
    if (!directory_error.empty())
    {
        return directory_error;
    }
    *path = util::JoinPath(directory, RandomName(L"RegKit-Setup-") + L".exe");
    const util::UniqueHandle file(
        CreateFileW(path->c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr)
    );
    DWORD written = 0;
    const bool saved = file && WriteFile(file.get(), data.data(), static_cast<DWORD>(data.size()), &written, nullptr) &&
                       written == data.size();
    const DWORD code = saved ? ERROR_SUCCESS : GetLastError();
    if (!saved)
    {
        DeleteFileW(path->c_str());
        RemoveDirectoryW(directory.c_str());
        path->clear();
        return util::FormatWin32Error(code != ERROR_SUCCESS ? code : ERROR_WRITE_FAULT);
    }
    return {};
}

util::UniqueHandle OpenVerifiedSetup(const std::wstring& path, const std::string& sha256)
{
    util::UniqueHandle file(CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
    LARGE_INTEGER size = {};
    if (!file || !GetFileSizeEx(file.get(), &size) || size.QuadPart <= 0 ||
        static_cast<uint64_t>(size.QuadPart) > kMaxSetupBytes)
    {
        return {};
    }
    std::string data(static_cast<size_t>(size.QuadPart), '\0');
    DWORD read = 0;
    if (!ReadFile(file.get(), data.data(), static_cast<DWORD>(data.size()), &read, nullptr) || read != data.size() ||
        _stricmp(Sha256(data).c_str(), sha256.c_str()) != 0)
    {
        return {};
    }
    return file;
}

} // namespace

void UpdateChecker::Attach(HWND owner, StatusCallback status)
{
    owner_ = owner;
    status_ = std::move(status);
}

void UpdateChecker::SetStatus(const std::wstring& text) const
{
    if (status_)
    {
        status_(text);
    }
}

void UpdateChecker::Cancel()
{
    session_.CancelAndJoin();
    running_ = false;
}

void UpdateChecker::Check(bool silent)
{
    if (running_)
    {
        return;
    }
    running_ = true;
    HWND owner = owner_;
    session_.Start([owner, silent](uint64_t, std::atomic_bool& cancel) {
        auto payload = std::make_unique<UpdateCheckPayload>();
        payload->silent = silent;
        std::string json;
        std::wstring error = HttpGet(kLatestReleaseUrl, cancel, &json, kMaxReleaseJsonBytes);
        std::wstring digest;
        if (error.empty() &&
            (!ReadRelease(json, &payload->version, &payload->download_url, &digest) || payload->version.empty()))
        {
            error = L"The response didn't contain a release.";
        }
        payload->sha256 = digest.starts_with(L"sha256:") ? util::WideToUtf8(digest.substr(7)) : std::string();
        if (cancel.load())
        {
            return;
        }
        if (!error.empty())
        {
            payload->failed = true;
            payload->error = L"Failed to reach the update server.\n" + error;
        }
        if (PostMessageW(owner, frame::message_id::kUpdateCheckReady, 0, reinterpret_cast<LPARAM>(payload.get())))
        {
            (void)payload.release();
        }
    });
}

void UpdateChecker::Download(const UpdateCheckPayload& release)
{
    if (running_)
    {
        return;
    }
    running_ = true;
    SetStatus(L"Downloading RegKit " + release.version + L"...");
    HWND owner = owner_;
    session_.Start(
        [owner, url = release.download_url, sha256 = release.sha256](uint64_t, std::atomic_bool& cancel) {
            auto payload = std::make_unique<UpdateCheckPayload>();
            payload->sha256 = sha256;
            const std::wstring error = SaveSetup(url, sha256, cancel, &payload->setup_path);
            if (cancel.load())
            {
                return;
            }
            if (!error.empty())
            {
                payload->failed = true;
                payload->setup_path.clear();
                payload->error = L"The update couldn't be downloaded.\n" + error;
            }
            if (PostMessageW(owner, frame::message_id::kUpdateCheckReady, 0, reinterpret_cast<LPARAM>(payload.get())))
            {
                (void)payload.release();
            }
        }
    );
}

void UpdateChecker::Apply(UpdateCheckPayload* payload)
{
    running_ = false;
    SetStatus(std::wstring());
    if (!payload)
    {
        return;
    }
    if (payload->failed)
    {
        if (!payload->silent)
        {
            ui::ShowError(owner_, payload->error);
        }
        return;
    }
    if (!payload->setup_path.empty())
    {
        const util::UniqueHandle verified = OpenVerifiedSetup(payload->setup_path, payload->sha256);
        if (!verified)
        {
            ui::ShowError(owner_, L"The downloaded setup changed after RegKit verified it and wasn't started.");
            return;
        }
        const HRESULT hr = win32::ShellOpen(owner_, payload->setup_path.c_str());
        if (SUCCEEDED(hr))
        {
            PostMessageW(owner_, WM_CLOSE, 0, 0);
        }
        else if (!win32::DialogCancelled(hr))
        {
            ui::ShowError(owner_, L"The setup couldn't be started.\n" + win32::FormatDialogError(hr));
        }
        return;
    }
    if (VersionParts(payload->version) <= VersionParts(REGKIT_VERSION_STR_W))
    {
        if (!payload->silent)
        {
            ui::ShowInfo(owner_, L"RegKit is up to date.");
        }
        return;
    }
    const std::wstring message =
        L"RegKit " + payload->version + L" is available. You are running " REGKIT_VERSION_STR_W L".\n\n" +
        (payload->download_url.empty() ? L"No setup was found for this build. Open the releases page?"
                                       : L"Download and install it now? RegKit closes when the setup starts.");
    if (ui::PromptChoice(owner_, message, L"Update available", payload->download_url.empty() ? L"Open" : L"Install", L"", L"Close", {70, 70, 70}) != IDYES)
    {
        return;
    }
    if (payload->download_url.empty())
    {
        win32::ShellOpen(owner_, kReleasesPage);
    }
    else
    {
        Download(*payload);
    }
}

} // namespace regkit::frame
