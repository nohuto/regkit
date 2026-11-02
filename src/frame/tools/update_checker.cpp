// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "frame/tools/update_checker.h"

#include "frame/window/message_ids.h"
#include "win32/file_dialog.h"
#include "win32/file_text.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"
#include "ui/feedback.h"
#include "win32/translation.h"

#include <win32ui/handle_owner.h>
#include <win32ui/json/reader.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system/file_io.h>
#include <win32ui/system/hash.h>
#include <win32ui/system/http.h>
#include <win32ui/system/paths.h>
#include <win32ui/system/registry_value.h>
#include <win32ui/system/shell_open.h>
#include <win32ui/system/strings.h>

#include <shlobj.h>

#include <algorithm>
#include <array>
#include <stdlib.h>
#include <utility>

namespace regkit::frame
{

namespace
{

constexpr wchar_t kReleasesPage[] = L"https://github.com/nohuto/regkit/releases";
constexpr wchar_t kLatestReleaseUrl[] = L"https://api.github.com/repos/nohuto/regkit/releases/latest";
constexpr const wchar_t* kSetupSuffix = sizeof(void*) == 8 ? L"-x64.exe" : L"-x86.exe";
constexpr const wchar_t* kPortableSuffix = sizeof(void*) == 8 ? L"-x64.zip" : L"-x86.zip";
constexpr wchar_t kUninstallKey[] = L"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\4678f42c-c6a2-4df9-bc2a-dddbd2613045_is1";
constexpr wchar_t kSetupArguments[] = L"/VERYSILENT /NORESTART /RELAUNCH=1 ";
constexpr DWORD kSetupCancelledBeforeInstall = 2;
constexpr size_t kMaxReleaseJsonBytes = 4ull * 1024ull * 1024ull;
constexpr size_t kMaxSetupBytes = 192ull * 1024ull * 1024ull;

const wchar_t* InstallScope()
{
    const std::wstring directory = util::ModuleDirectory();
    for (const auto& [root, scope] : {std::pair{HKEY_LOCAL_MACHINE, L"/ALLUSERS"}, std::pair{HKEY_CURRENT_USER, L"/CURRENTUSER"}})
    {
        std::wstring location;
        if (win32ui::ReadRegistryString(root, kUninstallKey, L"InstallLocation", &location) != ERROR_SUCCESS)
        {
            continue;
        }
        std::wstring_view path(location);
        while (path.size() > 3 && path.back() == L'\\')
        {
            path.remove_suffix(1);
        }
        if (util::EqualsInsensitive(path, directory))
        {
            return scope;
        }
    }
    return nullptr;
}

std::wstring ReleasePage(const std::wstring& tag)
{
    const bool plain = !tag.empty() && tag.find_first_not_of(L"0123456789.-_abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ") == std::wstring::npos;
    return plain ? std::wstring(kReleasesPage) + L"/tag/" + tag : kReleasesPage;
}

// https on the GitHub hosts only, also after redirects
bool IsReleaseUrl(std::wstring_view host, bool secure)
{
    constexpr const wchar_t* kHosts[] = {L"github.com", L"api.github.com", L"objects.githubusercontent.com", L"release-assets.githubusercontent.com"};
    return secure && std::any_of(std::begin(kHosts), std::end(kHosts), [&](const wchar_t* allowed) { return util::EqualsInsensitive(host, allowed); });
}

std::wstring HttpGet(const std::wstring& url, const std::atomic_bool& cancel, std::string* body, size_t max_bytes)
{
    const DWORD error = win32ui::HttpGet(url, L"RegKit", IsReleaseUrl, max_bytes, cancel, body);
    if (error == ERROR_ACCESS_DENIED)
    {
        return util::Tr(L"The release location isn't a trusted RegKit download address.");
    }
    if (error == ERROR_FILE_TOO_LARGE)
    {
        return util::Tr(L"The download is larger than RegKit accepts.");
    }
    if (HRESULT_FACILITY(error) == FACILITY_HTTP)
    {
        return L"HTTP " + std::to_wstring(HRESULT_CODE(error));
    }
    return win32ui::FormatWin32Error(error);
}

bool ReadRelease(const std::string& body, UpdateCheckPayload* release, std::wstring* digest)
{
    const std::wstring text = win32ui::Utf8ToWide(body).value_or(std::wstring());
    win32ui::JsonReader reader(text);
    const std::wstring_view prefix = release->installed ? L"RegKit-Setup-" : L"RegKit-Portable-";
    const std::wstring_view suffix = release->installed ? kSetupSuffix : kPortableSuffix;
    return reader.Object([&](const std::wstring& member) {
        if (member == L"tag_name")
        {
            return reader.String(&release->version, 64);
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
            // portable name becomes the saved file name
            if (read && release->download_url.empty() && name.starts_with(prefix) && name.size() > prefix.size() + suffix.size() &&
                name.ends_with(suffix) && name.find_first_of(L"\\/:") == std::wstring::npos)
            {
                release->download_url = std::move(link);
                release->file_name = std::move(name);
                *digest = std::move(hash);
            }
            return read;
        });
    }) && reader.End();
}

std::array<int, 3> VersionParts(const std::wstring& text)
{
    std::array<int, 3> parts = {};
    const size_t start = text.find_first_of(L"0123456789");
    if (start != std::wstring::npos)
    {
        swscanf_s(text.c_str() + start, L"%d.%d.%d", &parts[0], &parts[1], &parts[2]);
    }
    return parts;
}

DWORD CheckSha256(std::string_view data, const std::string& sha256)
{
    win32ui::Sha256Digest digest = {};
    if (const DWORD error = win32ui::Sha256(data, &digest))
    {
        return error;
    }
    return util::EqualsInsensitive(win32ui::HexBytes(digest, 0), std::wstring(sha256.begin(), sha256.end())) ? ERROR_SUCCESS : ERROR_INVALID_DATA;
}

void DeleteSetup(const std::wstring& path)
{
    DeleteFileW(path.c_str());
    RemoveDirectoryW(path.substr(0, path.find_last_of(L'\\')).c_str());
}

// setup uses a private folder, portable ZIP goes to Downloads
std::wstring SaveDownload(const std::wstring& url, const std::string& sha256, const std::wstring& file_name, bool installed, const std::atomic_bool& cancel, std::wstring* path)
{
    if (sha256.size() != 64 || sha256.find_first_not_of("0123456789abcdefABCDEF") != std::string::npos)
    {
        return util::Tr(L"The release didn't publish a checksum for this download.");
    }
    std::string data;
    const std::wstring error = HttpGet(url, cancel, &data, kMaxSetupBytes);
    if (!error.empty())
    {
        return error;
    }
    const DWORD check = data.empty() ? ERROR_INVALID_DATA : CheckSha256(data, sha256);
    if (check != ERROR_SUCCESS)
    {
        return check == ERROR_INVALID_DATA ? util::Tr(L"The downloaded file doesn't match the release checksum.") : win32ui::FormatWin32Error(check);
    }
    std::wstring directory;
    DWORD directory_error = ERROR_SUCCESS;
    if (installed)
    {
        directory_error = win32ui::CreateTempFolder(L"RegKit-", &directory);
    }
    else
    {
        directory = win32ui::KnownFolderPath(FOLDERID_Downloads);
        directory_error = directory.empty() ? ERROR_PATH_NOT_FOUND : ERROR_SUCCESS;
    }
    if (directory_error != ERROR_SUCCESS)
    {
        return win32ui::FormatWin32Error(directory_error);
    }
    win32ui::UniqueHandle file;
    for (int copy = 0; copy < 100 && !file; ++copy)
    {
        std::wstring name = installed ? L"RegKit-Setup-" + win32ui::RandomFileSuffix(L".exe").substr(1) : file_name;
        if (copy > 0)
        {
            name.insert(name.size() - 4, L" (" + std::to_wstring(copy) + L")");
        }
        *path = util::JoinPath(directory, name);
        file.reset(CreateFileW(path->c_str(), GENERIC_WRITE, 0, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (!file && GetLastError() != ERROR_FILE_EXISTS)
        {
            break;
        }
    }
    DWORD written = 0;
    const bool saved = file && WriteFile(file.get(), data.data(), static_cast<DWORD>(data.size()), &written, nullptr) &&
                       written == data.size();
    const DWORD code = saved ? ERROR_SUCCESS : GetLastError();
    const bool created = static_cast<bool>(file);
    file.reset();
    if (!saved)
    {
        if (created)
        {
            DeleteFileW(path->c_str());
        }
        if (installed)
        {
            RemoveDirectoryW(directory.c_str());
        }
        path->clear();
        return win32ui::FormatWin32Error(code != ERROR_SUCCESS ? code : ERROR_WRITE_FAULT);
    }
    return {};
}

win32ui::UniqueHandle OpenVerifiedSetup(const std::wstring& path, const std::string& sha256)
{
    win32ui::UniqueHandle file(CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_FLAG_OPEN_REPARSE_POINT, nullptr));
    LARGE_INTEGER size = {};
    if (!file || !GetFileSizeEx(file.get(), &size) || size.QuadPart <= 0 ||
        static_cast<uint64_t>(size.QuadPart) > kMaxSetupBytes)
    {
        return {};
    }
    std::string data(static_cast<size_t>(size.QuadPart), '\0');
    DWORD read = 0;
    if (!ReadFile(file.get(), data.data(), static_cast<DWORD>(data.size()), &read, nullptr) || read != data.size() ||
        CheckSha256(data, sha256) != ERROR_SUCCESS)
    {
        return {};
    }
    return file;
}

}

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
}

void UpdateChecker::Check(bool silent)
{
    HWND owner = owner_;
    session_.StartIfIdle(L"UpdateCheckThread", [owner, silent](uint64_t, std::atomic_bool& cancel) {
        auto payload = std::make_unique<UpdateCheckPayload>();
        payload->silent = silent;
        payload->installed = InstallScope() != nullptr;
        std::wstring error;
        // failures still post a payload so the UI shows the error
        try
        {
            std::string json;
            error = HttpGet(kLatestReleaseUrl, cancel, &json, kMaxReleaseJsonBytes);
            std::wstring digest;
            if (error.empty() && (!ReadRelease(json, payload.get(), &digest) || payload->version.empty()))
            {
                error = util::Tr(L"The response didn't contain a release.");
            }
            payload->sha256 = digest.starts_with(L"sha256:") ? win32ui::WideToMultiByte(digest.substr(7), CP_UTF8) : std::string();
        }
        catch (const std::exception&)
        {
            error = win32ui::FormatWin32Error(ERROR_NOT_ENOUGH_MEMORY);
        }
        if (cancel.load())
        {
            return;
        }
        if (!error.empty())
        {
            payload->failed = true;
            payload->error = util::TrDetail(L"Failed to reach the update server.", error);
        }
        win32ui::PostPayload(owner, frame::message_id::kUpdateCheckReady, 0, payload);
    });
}

void UpdateChecker::Download(const UpdateCheckPayload& release)
{
    HWND owner = owner_;
    const bool started = session_.StartIfIdle(
        L"UpdateDownloadThread",
        [owner, url = release.download_url, sha256 = release.sha256, file_name = release.file_name, installed = release.installed](uint64_t, std::atomic_bool& cancel) {
            auto payload = std::make_unique<UpdateCheckPayload>();
            payload->sha256 = sha256;
            payload->installed = installed;
            std::wstring error;
            try
            {
                error = SaveDownload(url, sha256, file_name, installed, cancel, &payload->download_path);
            }
            catch (const std::exception&)
            {
                error = win32ui::FormatWin32Error(ERROR_NOT_ENOUGH_MEMORY);
            }
            if (cancel.load())
            {
                return;
            }
            if (!error.empty())
            {
                payload->failed = true;
                payload->download_path.clear();
                payload->error = util::TrDetail(L"The update couldn't be downloaded.", error);
            }
            win32ui::PostPayload(owner, frame::message_id::kUpdateCheckReady, 0, payload);
        }
    );
    if (started)
    {
        SetStatus(util::TrLabel(L"Downloading", L"RegKit " + release.version));
    }
}

void UpdateChecker::Apply(UpdateCheckPayload* payload)
{
    SetStatus(std::wstring());
    if (!payload)
    {
        return;
    }
    if (payload->setup_exited)
    {
        setup_running_ = false;
        EnableWindow(owner_, TRUE);
        if (!payload->error.empty())
        {
            ui::ShowError(owner_, payload->error);
        }
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
    if (!payload->download_path.empty())
    {
        if (!payload->installed)
        {
            ui::RevealFile(owner_, payload->download_path);
            return;
        }
        DropSetup();
        setup_path_ = std::move(payload->download_path);
        setup_sha256_ = std::move(payload->sha256);
        PostMessageW(owner_, WM_CLOSE, 0, 0);
        return;
    }
    if (VersionParts(payload->version) <= VersionParts(REGKIT_VERSION_TEXT))
    {
        if (!payload->silent)
        {
            ui::ShowInfo(owner_, util::Tr(L"RegKit is up to date."));
        }
        return;
    }
    const bool found = !payload->download_url.empty();
    const std::wstring message =
        L"RegKit " + payload->version + L" (" + util::TrLabel(L"installed", REGKIT_VERSION_TEXT) + L")\n\n" +
        (!found               ? util::Tr(L"No download was found for this build. Open the release page?")
         : payload->installed ? util::Tr(L"Download and install it now? RegKit restarts once the update is installed.")
                              : util::Tr(L"Download the new portable version?"));
    const std::wstring download = !found ? L"" : payload->installed ? util::Tr(L"Update")
                                                                    : util::Tr(L"Download Portable");
    const std::wstring release = found && payload->installed ? L"" : util::Tr(L"Open Release");
    const int choice = ui::PromptChoice(owner_, message, util::Tr(L"Update Available"), download, release, util::Tr(L"Close"), {70, 70, 70});
    if (choice == IDYES)
    {
        Download(*payload);
    }
    else if (choice == IDNO)
    {
        ui::ReportFileDialogResult(owner_, win32ui::ShellOpen(owner_, ReleasePage(payload->version).c_str()));
    }
}

bool UpdateChecker::StartSetup()
{
    const std::wstring path = std::exchange(setup_path_, {});
    const std::string sha256 = std::exchange(setup_sha256_, {});
    const wchar_t* scope = InstallScope();
    // silent setup reuses prior choices and relaunches RegKit with saved session
    const std::wstring arguments = std::wstring(kSetupArguments) + (scope ? scope : L"");
    SHELLEXECUTEINFOW info = {sizeof(info)};
    info.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC | SEE_MASK_FLAG_NO_UI;
    info.hwnd = owner_;
    info.lpFile = path.c_str();
    info.lpParameters = arguments.c_str();
    info.nShow = SW_SHOWNORMAL;
    bool started = false;
    {
        const win32ui::UniqueHandle verified = OpenVerifiedSetup(path, sha256);
        if (!verified)
        {
            ui::ShowError(owner_, util::Tr(L"The downloaded setup changed after RegKit verified it and wasn't started."));
        }
        else if (!ShellExecuteExW(&info) || !info.hProcess)
        {
            ui::ShowError(owner_, util::TrDetail(L"The setup couldn't be started.", win32ui::FormatWin32Error(GetLastError())));
        }
        else
        {
            started = true;
        }
    }
    if (!started)
    {
        DeleteSetup(path);
        return false;
    }
    setup_running_ = true;
    const HANDLE process = info.hProcess;
    const uint64_t watcher = session_.Start(L"UpdateSetupThread", [owner = owner_, process](uint64_t, std::atomic_bool& cancel) {
        const win32ui::UniqueHandle owned(process);
        while (WaitForSingleObject(process, 200) == WAIT_TIMEOUT)
        {
            if (cancel.load())
            {
                return;
            }
        }
        DWORD code = 0;
        GetExitCodeProcess(process, &code);
        auto payload = std::make_unique<UpdateCheckPayload>();
        payload->setup_exited = true;
        if (code != kSetupCancelledBeforeInstall)
        {
            payload->error = util::TrDetail(L"The update couldn't be installed.", util::TrLabel(L"Exit code", std::to_wstring(code)));
        }
        win32ui::PostPayload(owner, frame::message_id::kUpdateCheckReady, 0, payload);
    });
    if (!watcher)
    {
        // nothing would report the exit of setup, keep the window usable (setup_running_ stays for its WM_ENDSESSION)
        CloseHandle(process);
        return false;
    }
    return true;
}

void UpdateChecker::DropSetup()
{
    if (!setup_path_.empty())
    {
        DeleteSetup(std::exchange(setup_path_, {}));
        setup_sha256_.clear();
    }
}

}
