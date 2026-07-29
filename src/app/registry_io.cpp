// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#include "../../include/app/registry_io.h"

#include <cwchar>
#include <cwctype>
#include <unordered_set>
#include <vector>

#include <commdlg.h>
#include <shellapi.h>
#include <shlobj.h>

#include "../../include/app/value_dialogs.h"
#include "../../include/win32/win32_helpers.h"
#include "regfile/reg_file.h"
#include "registry/registry_path.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/shell_paths.h"

namespace regkit {

namespace {

using util::FormatWin32Error;
using util::ToLower;
using util::TrimWhitespace;

std::wstring GetRegExePath() {
  wchar_t system_dir[MAX_PATH] = {};
  UINT len = GetSystemDirectoryW(system_dir, _countof(system_dir));
  if (len == 0 || len >= _countof(system_dir)) {
    return L"reg.exe";
  }
  return util::JoinPath(system_dir, L"reg.exe");
}

bool RunRegCommand(const std::wstring& args, DWORD* exit_code, std::wstring* error) {
  std::wstring reg = GetRegExePath();
  std::wstring cmdline = L"\"" + reg + L"\" " + args;
  SECURITY_ATTRIBUTES security = {};
  security.nLength = sizeof(security);
  security.bInheritHandle = TRUE;
  HANDLE read_pipe = nullptr;
  HANDLE write_pipe = nullptr;
  bool capture_output = CreatePipe(&read_pipe, &write_pipe, &security, 0) != FALSE;
  if (capture_output) {
    SetHandleInformation(read_pipe, HANDLE_FLAG_INHERIT, 0);
  }
  STARTUPINFOW si = {};
  si.cb = sizeof(si);
  si.dwFlags = STARTF_USESHOWWINDOW;
  si.wShowWindow = SW_HIDE;
  if (capture_output) {
    si.dwFlags |= STARTF_USESTDHANDLES;
    si.hStdInput = GetStdHandle(STD_INPUT_HANDLE);
    si.hStdOutput = write_pipe;
    si.hStdError = write_pipe;
  }
  PROCESS_INFORMATION pi = {};
  DWORD flags = CREATE_NO_WINDOW;
  if (!CreateProcessW(nullptr, cmdline.data(), nullptr, nullptr, capture_output ? TRUE : FALSE, flags, nullptr, nullptr, &si, &pi)) {
    if (read_pipe) {
      CloseHandle(read_pipe);
    }
    if (write_pipe) {
      CloseHandle(write_pipe);
    }
    if (error) {
      *error = FormatWin32Error(GetLastError());
    }
    return false;
  }
  if (write_pipe) {
    CloseHandle(write_pipe);
    write_pipe = nullptr;
  }
  std::string output;
  if (capture_output) {
    char buffer[4096] = {};
    DWORD read = 0;
    while (ReadFile(read_pipe, buffer, static_cast<DWORD>(sizeof(buffer)), &read, nullptr) && read > 0) {
      output.append(buffer, buffer + read);
    }
    CloseHandle(read_pipe);
    read_pipe = nullptr;
  }
  WaitForSingleObject(pi.hProcess, INFINITE);
  DWORD code = 0;
  GetExitCodeProcess(pi.hProcess, &code);
  CloseHandle(pi.hProcess);
  CloseHandle(pi.hThread);
  if (exit_code) {
    *exit_code = code;
  }
  if (code != 0 && error) {
    std::wstring detail;
    if (!output.empty()) {
      int chars = MultiByteToWideChar(CP_OEMCP, 0, output.data(), static_cast<int>(output.size()), nullptr, 0);
      if (chars > 0) {
        detail.resize(static_cast<size_t>(chars));
        MultiByteToWideChar(CP_OEMCP, 0, output.data(), static_cast<int>(output.size()), detail.data(), chars);
        detail = TrimWhitespace(detail);
      }
    }
    if (detail.empty()) {
      detail = L"reg.exe exited with code " + std::to_wstring(code) + L".";
    }
    *error = detail;
  }
  return code == 0;
}

std::wstring NormalizeExportKeyPath(const std::wstring& key_path, std::wstring* error) {
  const std::wstring path =
      registry_path::Normalize(key_path, util::GetCurrentUserSidString());
  const size_t split = path.find(L'\\');
  const std::wstring_view root(path.data(),
                               split == std::wstring::npos ? path.size()
                                                          : split);
  const bool standard =
      registry_path::Equals(root, L"HKEY_LOCAL_MACHINE") ||
      registry_path::Equals(root, L"HKEY_CURRENT_USER") ||
      registry_path::Equals(root, L"HKEY_CLASSES_ROOT") ||
      registry_path::Equals(root, L"HKEY_USERS") ||
      registry_path::Equals(root, L"HKEY_CURRENT_CONFIG");
  if (!standard) {
    if (error) {
      *error = L"Export supports standard hives only.";
    }
    return {};
  }
  return path;
}

bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  std::wstring buffer(32768, L'\0');
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer.data();
  ofn.nMaxFile = static_cast<DWORD>(buffer.size());
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
  if (!GetOpenFileNameW(&ofn)) {
    return false;
  }
  *path = buffer.c_str();
  return true;
}

std::wstring SanitizeFileName(const std::wstring& name) {
  std::wstring out;
  out.reserve(name.size());
  for (wchar_t ch : name) {
    if (ch < 32 || ch == L'<' || ch == L'>' || ch == L':' || ch == L'"' || ch == L'/' || ch == L'\\' || ch == L'|' || ch == L'?' || ch == L'*') {
      out.push_back(L'_');
    } else {
      out.push_back(ch);
    }
  }
  while (!out.empty() && (out.back() == L' ' || out.back() == L'.')) {
    out.pop_back();
  }
  while (!out.empty() && out.front() == L' ') {
    out.erase(out.begin());
  }
  if (out.empty()) {
    return L"RegistryExport";
  }
  return out;
}

bool PromptSaveRegFile(HWND owner, const std::wstring& default_name, std::wstring* path) {
  if (!path) {
    return false;
  }
  std::wstring name = default_name;
  if (name.empty()) {
    name = L"RegistryExport.reg";
  }
  std::wstring buffer(32768, L'\0');
  wcsncpy_s(buffer.data(), buffer.size(), name.c_str(), _TRUNCATE);
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0\0";
  ofn.lpstrFile = buffer.data();
  ofn.nMaxFile = static_cast<DWORD>(buffer.size());
  ofn.lpstrDefExt = L"reg";
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;
  if (!GetSaveFileNameW(&ofn)) {
    return false;
  }
  *path = buffer.c_str();
  return true;
}

std::wstring EnsureRegExtension(std::wstring path);

std::wstring ExportDefaultNameFromKeyPath(const std::wstring& key_path) {
  std::wstring trimmed = key_path;
  while (!trimmed.empty() && (trimmed.back() == L'\\' || trimmed.back() == L'/')) {
    trimmed.pop_back();
  }
  if (trimmed.empty()) {
    return L"RegistryExport.reg";
  }
  size_t slash = trimmed.find_last_of(L"\\/");
  std::wstring leaf = (slash == std::wstring::npos) ? trimmed : trimmed.substr(slash + 1);
  if (leaf.empty()) {
    leaf = L"RegistryExport";
  }
  std::wstring file_name = EnsureRegExtension(SanitizeFileName(leaf));
  PWSTR desktop = nullptr;
  if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Desktop, 0, nullptr, &desktop)) && desktop) {
    std::wstring path = util::JoinPath(desktop, file_name);
    CoTaskMemFree(desktop);
    return path;
  }
  if (desktop) {
    CoTaskMemFree(desktop);
  }
  return file_name;
}

bool FilterExportedRegFile(const std::wstring& source, const std::wstring& target, std::wstring* error) {
  std::wstring content;
  bool utf16 = false;
  if (!util::ReadTextFile(source, &content, &utf16)) {
    if (error) {
      *error = L"Failed to read exported registry file.";
    }
    return false;
  }

  std::wstring output;
  output.reserve(content.size());
  bool in_section = false;
  bool wrote_section = false;

  size_t start = 0;
  while (start < content.size()) {
    size_t end = content.find(L'\n', start);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(start, end - start);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;

    if (!line.empty() && line.front() == L'[' && line.back() == L']') {
      if (wrote_section) {
        break;
      }
      wrote_section = true;
      in_section = true;
      output.append(line);
      output.append(L"\r\n");
      continue;
    }

    if (!wrote_section || in_section) {
      output.append(line);
      output.append(L"\r\n");
    }
  }

  if (!util::WriteTextFile(target, output, utf16)) {
    if (error) {
      *error = L"Failed to write exported registry file.";
    }
    return false;
  }
  return true;
}

bool AppendRegContent(const std::wstring& content,
                      regfile::Document* output,
                      std::wstring* error) {
  if (!output) {
    return false;
  }
  regfile::Document parsed;
  if (!regfile::Parse(content, &parsed)) {
    if (error) {
      *error = L"Failed to parse exported registry data.";
    }
    return false;
  }
  for (const auto& path : parsed.key_order) {
    const std::wstring lower = ToLower(path);
    auto source = parsed.keys.find(lower);
    if (source == parsed.keys.end()) {
      continue;
    }
    auto [target, inserted] =
        output->keys.try_emplace(lower, std::move(source->second));
    if (inserted) {
      output->key_order.push_back(path);
      continue;
    }
    for (auto& value : source->second.values) {
      target->second.values[value.first] = std::move(value.second);
    }
  }
  return true;
}

bool FilterRegFileValues(const std::wstring& content,
                         const std::vector<std::wstring>& values,
                         regfile::Document* output, std::wstring* error) {
  if (!output || !regfile::Parse(content, output) ||
      output->key_order.empty()) {
    if (error) {
      *error = L"Failed to parse exported registry data.";
    }
    return false;
  }
  std::unordered_set<std::wstring> wanted;
  wanted.reserve(values.size());
  for (const auto& value : values) {
    wanted.insert(ToLower(value));
  }
  const std::wstring first_path = output->key_order.front();
  const std::wstring first_lower = ToLower(first_path);
  auto first_key = output->keys.find(first_lower);
  if (first_key == output->keys.end()) {
    return false;
  }
  for (auto value = first_key->second.values.begin();
       value != first_key->second.values.end();) {
    if (wanted.find(value->first) == wanted.end()) {
      value = first_key->second.values.erase(value);
    } else {
      ++value;
    }
  }
  if (first_key->second.values.empty()) {
    if (error) {
      *error = L"No selected values were found in the export.";
    }
    return false;
  }
  regfile::Key selected = std::move(first_key->second);
  output->keys.clear();
  output->key_order.assign(1, first_path);
  output->keys.emplace(first_lower, std::move(selected));
  return true;
}

bool ExportKeyToContent(const std::wstring& key_path, bool include_subkeys, std::wstring* content, bool* utf16, std::wstring* error) {
  if (!content) {
    return false;
  }
  std::wstring normalized = NormalizeExportKeyPath(key_path, error);
  if (normalized.empty()) {
    return false;
  }

  wchar_t temp_dir[MAX_PATH] = {};
  if (!GetTempPathW(_countof(temp_dir), temp_dir)) {
    if (error) {
      *error = L"Failed to locate temp path.";
    }
    return false;
  }
  wchar_t temp_file[MAX_PATH] = {};
  if (!GetTempFileNameW(temp_dir, L"reg", 0, temp_file)) {
    if (error) {
      *error = L"Failed to create temp file.";
    }
    return false;
  }
  std::wstring temp_path = temp_file;

  std::wstring args = L"export \"" + normalized + L"\" \"" + temp_path + L"\" /y";
  if (!RunRegCommand(args, nullptr, error)) {
    DeleteFileW(temp_path.c_str());
    return false;
  }

  std::wstring read_path = temp_path;
  std::wstring filtered_path;
  if (!include_subkeys) {
    wchar_t filtered_file[MAX_PATH] = {};
    if (!GetTempFileNameW(temp_dir, L"reg", 0, filtered_file)) {
      DeleteFileW(temp_path.c_str());
      if (error) {
        *error = L"Failed to create temp file.";
      }
      return false;
    }
    filtered_path = filtered_file;
    if (!FilterExportedRegFile(temp_path, filtered_path, error)) {
      DeleteFileW(temp_path.c_str());
      DeleteFileW(filtered_path.c_str());
      return false;
    }
    read_path = filtered_path;
  }

  bool is_utf16 = false;
  if (!util::ReadTextFile(read_path, content, &is_utf16)) {
    if (error) {
      *error = L"Failed to read exported registry file.";
    }
    DeleteFileW(temp_path.c_str());
    if (!filtered_path.empty()) {
      DeleteFileW(filtered_path.c_str());
    }
    return false;
  }

  DeleteFileW(temp_path.c_str());
  if (!filtered_path.empty()) {
    DeleteFileW(filtered_path.c_str());
  }
  if (utf16) {
    *utf16 = is_utf16;
  }
  return true;
}

std::wstring EnsureRegExtension(std::wstring path) {
  if (path.empty()) {
    return path;
  }
  size_t slash = path.find_last_of(L"\\/");
  size_t dot = path.find_last_of(L'.');
  if (dot == std::wstring::npos || (slash != std::wstring::npos && dot < slash)) {
    path.append(L".reg");
  }
  return path;
}

} // namespace

bool ImportRegFileFromPath(const std::wstring& path, std::wstring* error) {
  if (path.empty()) {
    return false;
  }
  std::wstring args = L"import \"" + path + L"\"";
  return RunRegCommand(args, nullptr, error);
}

bool ExportRegFile(HWND owner, const std::wstring& key_path, std::wstring* error) {
  ExportOptions options;
  if (!PromptForExportOptions(owner, ExportDefaultNameFromKeyPath(key_path), &options)) {
    return false;
  }
  options.path = EnsureRegExtension(options.path);

  std::wstring normalized = NormalizeExportKeyPath(key_path, error);
  if (normalized.empty()) {
    return false;
  }

  std::wstring target_path = options.path;
  std::wstring temp_path;
  if (!options.include_subkeys) {
    wchar_t temp_dir[MAX_PATH] = {};
    if (!GetTempPathW(_countof(temp_dir), temp_dir)) {
      if (error) {
        *error = L"Failed to locate temp path.";
      }
      return false;
    }
    wchar_t temp_file[MAX_PATH] = {};
    if (!GetTempFileNameW(temp_dir, L"reg", 0, temp_file)) {
      if (error) {
        *error = L"Failed to create temp file.";
      }
      return false;
    }
    temp_path = temp_file;
    target_path = temp_path;
  }

  std::wstring args = L"export \"" + normalized + L"\" \"" + target_path + L"\" /y";
  if (!RunRegCommand(args, nullptr, error)) {
    if (!temp_path.empty()) {
      DeleteFileW(temp_path.c_str());
    }
    return false;
  }

  if (!temp_path.empty()) {
    if (!FilterExportedRegFile(temp_path, options.path, error)) {
      DeleteFileW(temp_path.c_str());
      return false;
    }
    DeleteFileW(temp_path.c_str());
  }

  if (options.open_after) {
    ShellExecuteW(owner, L"open", options.path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  }
  return true;
}

bool ExportRegFileSelection(HWND owner, const std::wstring& base_key_path, const std::vector<std::wstring>& value_names, const std::vector<std::wstring>& subkey_names, std::wstring* error) {
  if (value_names.empty() && subkey_names.empty()) {
    if (error) {
      *error = L"No data to export.";
    }
    return false;
  }
  std::wstring first_name;
  if (!value_names.empty()) {
    first_name = value_names.front();
    if (first_name.empty()) {
      first_name = L"Default";
    }
  } else if (!subkey_names.empty()) {
    first_name = subkey_names.front();
  }
  std::wstring default_name = SanitizeFileName(first_name);
  std::wstring path;
  if (!PromptSaveRegFile(owner, EnsureRegExtension(default_name), &path)) {
    return false;
  }
  path = EnsureRegExtension(path);

  regfile::Document output_document;
  bool output_utf16 = false;
  bool output_utf16_set = false;

  auto append_content = [&](const std::wstring& content, bool utf16) -> bool {
    if (!output_utf16_set) {
      output_utf16 = utf16;
      output_utf16_set = true;
    }
    return AppendRegContent(content, &output_document, error);
  };

  if (!value_names.empty()) {
    std::wstring content;
    bool utf16 = false;
    if (!ExportKeyToContent(base_key_path, false, &content, &utf16, error)) {
      return false;
    }
    regfile::Document filtered;
    if (!FilterRegFileValues(content, value_names, &filtered, error)) {
      return false;
    }
    if (!output_utf16_set) {
      output_utf16 = utf16;
      output_utf16_set = true;
    }
    for (const auto& key_path : filtered.key_order) {
      const std::wstring lower = ToLower(key_path);
      auto key = filtered.keys.find(lower);
      if (key != filtered.keys.end()) {
        output_document.key_order.push_back(key_path);
        output_document.keys.emplace(lower, std::move(key->second));
      }
    }
  }

  for (const auto& subkey : subkey_names) {
    if (subkey.empty()) {
      continue;
    }
    std::wstring key_path = base_key_path;
    if (!key_path.empty()) {
      key_path += L"\\";
    }
    key_path += subkey;
    std::wstring content;
    bool utf16 = false;
    if (!ExportKeyToContent(key_path, true, &content, &utf16, error)) {
      return false;
    }
    append_content(content, utf16);
  }

  if (output_document.key_order.empty()) {
    if (error) {
      *error = L"No data to export.";
    }
    return false;
  }

  if (!util::WriteTextFile(path, regfile::Serialize(output_document),
                           output_utf16)) {
    if (error) {
      *error = L"Failed to write exported registry file.";
    }
    return false;
  }
  return true;
}

bool LoadHive(HWND owner, HKEY root, std::wstring* error) {
  std::wstring file_path;
  if (!PromptOpenFile(owner, L"Hive Files (*.*)\0*.*\0", &file_path)) {
    return false;
  }
  std::wstring mount_name;
  if (!PromptForValueText(owner, L"", L"Load Hive", L"Key name:", &mount_name)) {
    return false;
  }
  if (mount_name.empty()) {
    if (error) {
      *error = L"Key name is required.";
    }
    return false;
  }
  LONG result = RegLoadKeyW(root, mount_name.c_str(), file_path.c_str());
  if (result != ERROR_SUCCESS) {
    if (error) {
      *error = FormatWin32Error(result);
    }
    return false;
  }
  return true;
}

bool UnloadHive(HWND owner, HKEY root, const std::wstring& subkey, std::wstring* error) {
  std::wstring target = subkey;
  if (target.empty()) {
    if (!PromptForValueText(owner, L"", L"Unload Hive", L"Key name:", &target)) {
      return false;
    }
  }
  if (target.empty()) {
    if (error) {
      *error = L"Key name is required.";
    }
    return false;
  }
  LONG result = RegUnLoadKeyW(root, target.c_str());
  if (result != ERROR_SUCCESS) {
    if (error) {
      *error = FormatWin32Error(result);
    }
    return false;
  }
  return true;
}

} // namespace regkit
