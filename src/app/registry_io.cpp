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

#include "app/registry_io.h"

#include <cwchar>
#include <cwctype>
#include <unordered_set>
#include <vector>

#include <commdlg.h>
#include <shellapi.h>

#include "app/value_dialogs.h"
#include "win32/win32_helpers.h"

namespace regkit {

namespace {

std::wstring FormatWin32Error(DWORD code) {
  if (code == 0) {
    return L"";
  }
  wchar_t buffer[512] = {};
  DWORD len = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, code, 0, buffer, static_cast<DWORD>(_countof(buffer)), nullptr);
  if (len == 0) {
    return L"Unknown error.";
  }
  return buffer;
}

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
  STARTUPINFOW si = {};
  si.cb = sizeof(si);
  si.dwFlags = STARTF_USESHOWWINDOW;
  si.wShowWindow = SW_HIDE;
  PROCESS_INFORMATION pi = {};
  DWORD flags = CREATE_NO_WINDOW;
  if (!CreateProcessW(nullptr, cmdline.data(), nullptr, nullptr, FALSE, flags, nullptr, nullptr, &si, &pi)) {
    if (error) {
      *error = FormatWin32Error(GetLastError());
    }
    return false;
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
    *error = L"reg.exe exited with code " + std::to_wstring(code) + L".";
  }
  return code == 0;
}

bool EqualsInsensitive(const std::wstring& left, const std::wstring& right) {
  return _wcsicmp(left.c_str(), right.c_str()) == 0;
}

std::wstring NormalizeExportKeyPath(const std::wstring& key_path, std::wstring* error) {
  std::wstring path = key_path;
  if (path.empty()) {
    return path;
  }
  for (auto& ch : path) {
    if (ch == L'/') {
      ch = L'\\';
    }
  }
  while (!path.empty() && path.front() == L'\\') {
    path.erase(path.begin());
  }

  if (path.size() >= 9 && _wcsnicmp(path.c_str(), L"REGISTRY\\", 9) == 0) {
    path.erase(0, 9);
  }
  if (path.size() >= 9 && _wcsnicmp(path.c_str(), L"Registry\\", 9) == 0) {
    path.erase(0, 9);
  }

  size_t slash = path.find(L'\\');
  std::wstring root = (slash == std::wstring::npos) ? path : path.substr(0, slash);
  std::wstring rest = (slash == std::wstring::npos) ? L"" : path.substr(slash + 1);

  auto map_root = [&](const wchar_t* name, const wchar_t* mapped) -> bool {
    if (_wcsicmp(root.c_str(), name) != 0) {
      return false;
    }
    root = mapped;
    return true;
  };

  map_root(L"HKLM", L"HKEY_LOCAL_MACHINE") || map_root(L"HKEY_LOCAL_MACHINE", L"HKEY_LOCAL_MACHINE") || map_root(L"HKCU", L"HKEY_CURRENT_USER") || map_root(L"HKEY_CURRENT_USER", L"HKEY_CURRENT_USER") || map_root(L"HKCR", L"HKEY_CLASSES_ROOT") || map_root(L"HKEY_CLASSES_ROOT", L"HKEY_CLASSES_ROOT") || map_root(L"HKU", L"HKEY_USERS") || map_root(L"HKEY_USERS", L"HKEY_USERS") || map_root(L"HKCC", L"HKEY_CURRENT_CONFIG") || map_root(L"HKEY_CURRENT_CONFIG", L"HKEY_CURRENT_CONFIG");

  if (EqualsInsensitive(root, L"Machine")) {
    root = L"HKEY_LOCAL_MACHINE";
  } else if (EqualsInsensitive(root, L"User") || EqualsInsensitive(root, L"Users")) {
    root = L"HKEY_USERS";
  }

  if (root.empty() || root == L"REGISTRY") {
    if (error) {
      *error = L"Export supports standard hives only.";
    }
    return L"";
  }
  if (!rest.empty()) {
    return root + L"\\" + rest;
  }
  return root;
}

bool PromptOpenFile(HWND owner, const wchar_t* filter, std::wstring* path) {
  if (!path) {
    return false;
  }
  wchar_t buffer[MAX_PATH] = {};
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrFile = buffer;
  ofn.nMaxFile = static_cast<DWORD>(_countof(buffer));
  ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
  if (!GetOpenFileNameW(&ofn)) {
    return false;
  }
  *path = buffer;
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
  if (name.size() >= MAX_PATH) {
    name.resize(MAX_PATH - 1);
  }
  std::wstring buffer(MAX_PATH, L'\0');
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

bool ReadFileBytes(const std::wstring& path, std::vector<BYTE>* out) {
  if (!out) {
    return false;
  }
  out->clear();
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0) {
    CloseHandle(file);
    return false;
  }
  if (size.QuadPart > static_cast<LONGLONG>(64 * 1024 * 1024)) {
    CloseHandle(file);
    return false;
  }
  out->resize(static_cast<size_t>(size.QuadPart));
  DWORD read = 0;
  BOOL ok = ReadFile(file, out->data(), static_cast<DWORD>(out->size()), &read, nullptr);
  CloseHandle(file);
  if (!ok || read != out->size()) {
    out->clear();
    return false;
  }
  return true;
}

bool ReadRegFileText(const std::wstring& path, std::wstring* out, bool* utf16) {
  if (!out) {
    return false;
  }
  std::vector<BYTE> bytes;
  if (!ReadFileBytes(path, &bytes)) {
    return false;
  }
  bool is_utf16 = false;
  if (bytes.size() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) {
    is_utf16 = true;
    size_t wchar_count = (bytes.size() - 2) / sizeof(wchar_t);
    out->assign(reinterpret_cast<const wchar_t*>(bytes.data() + 2), wchar_count);
  } else {
    if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) {
      std::string utf8(reinterpret_cast<const char*>(bytes.data() + 3), bytes.size() - 3);
      *out = util::Utf8ToWide(utf8);
    } else {
      std::string utf8(reinterpret_cast<const char*>(bytes.data()), bytes.size());
      *out = util::Utf8ToWide(utf8);
    }
  }
  if (utf16) {
    *utf16 = is_utf16;
  }
  return !out->empty();
}

bool WriteRegFileText(const std::wstring& path, const std::wstring& text, bool utf16) {
  HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  DWORD written = 0;
  if (utf16) {
    const BYTE bom[2] = {0xFF, 0xFE};
    WriteFile(file, bom, sizeof(bom), &written, nullptr);
    const BYTE* data = reinterpret_cast<const BYTE*>(text.data());
    DWORD bytes = static_cast<DWORD>(text.size() * sizeof(wchar_t));
    BOOL ok = WriteFile(file, data, bytes, &written, nullptr);
    CloseHandle(file);
    return ok != 0;
  }
  std::string utf8 = util::WideToUtf8(text);
  BOOL ok = WriteFile(file, utf8.data(), static_cast<DWORD>(utf8.size()), &written, nullptr);
  CloseHandle(file);
  return ok != 0;
}

bool FilterExportedRegFile(const std::wstring& source, const std::wstring& target, std::wstring* error) {
  std::wstring content;
  bool utf16 = false;
  if (!ReadRegFileText(source, &content, &utf16)) {
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

  if (!WriteRegFileText(target, output, utf16)) {
    if (error) {
      *error = L"Failed to write exported registry file.";
    }
    return false;
  }
  return true;
}

bool FindFirstSection(const std::wstring& content, size_t* section_pos) {
  if (section_pos) {
    *section_pos = std::wstring::npos;
  }
  size_t pos = 0;
  while (pos < content.size()) {
    size_t end = content.find(L'\n', pos);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(pos, end - pos);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    if (!line.empty() && line.front() == L'[' && line.back() == L']') {
      if (section_pos) {
        *section_pos = pos;
      }
      return true;
    }
    pos = end + 1;
  }
  return false;
}

bool AppendRegContent(const std::wstring& content, std::wstring* output) {
  if (!output) {
    return false;
  }
  if (output->empty()) {
    output->append(content);
    return true;
  }
  size_t section_pos = std::wstring::npos;
  if (!FindFirstSection(content, &section_pos) || section_pos == std::wstring::npos) {
    return true;
  }
  output->append(content.substr(section_pos));
  return true;
}

bool ParseValueLine(const std::wstring& line, std::wstring* value_name) {
  if (!value_name) {
    return false;
  }
  value_name->clear();
  if (line.empty()) {
    return false;
  }
  if (line[0] == L'@') {
    if (line.size() >= 2 && line[1] == L'=') {
      return true;
    }
    return false;
  }
  if (line[0] != L'"') {
    return false;
  }
  bool escape = false;
  for (size_t i = 1; i < line.size(); ++i) {
    wchar_t ch = line[i];
    if (escape) {
      value_name->push_back(ch);
      escape = false;
      continue;
    }
    if (ch == L'\\') {
      escape = true;
      continue;
    }
    if (ch == L'"') {
      return true;
    }
    value_name->push_back(ch);
  }
  value_name->clear();
  return false;
}

std::wstring ToLower(const std::wstring& text) {
  std::wstring out;
  out.reserve(text.size());
  for (wchar_t ch : text) {
    out.push_back(static_cast<wchar_t>(towlower(ch)));
  }
  return out;
}

bool FilterRegFileValues(const std::wstring& content, const std::vector<std::wstring>& values, std::wstring* output, std::wstring* error) {
  if (!output) {
    return false;
  }

  std::unordered_set<std::wstring> wanted;
  wanted.reserve(values.size());
  for (const auto& value : values) {
    wanted.insert(ToLower(value));
  }

  std::wstring out;
  out.reserve(content.size());
  bool in_section = false;
  bool kept_value = false;
  bool keep_continuation = false;
  bool skip_continuation = false;

  size_t pos = 0;
  while (pos < content.size()) {
    size_t end = content.find(L'\n', pos);
    if (end == std::wstring::npos) {
      end = content.size();
    }
    std::wstring line = content.substr(pos, end - pos);
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    pos = end + 1;

    if (!line.empty() && line.front() == L'[' && line.back() == L']') {
      if (in_section) {
        break;
      }
      in_section = true;
      out.append(line);
      out.append(L"\r\n");
      continue;
    }

    if (!in_section) {
      out.append(line);
      out.append(L"\r\n");
      continue;
    }

    if (keep_continuation) {
      out.append(line);
      out.append(L"\r\n");
      if (line.empty() || line.back() != L'\\') {
        keep_continuation = false;
      }
      continue;
    }
    if (skip_continuation) {
      if (line.empty() || line.back() != L'\\') {
        skip_continuation = false;
      }
      continue;
    }

    std::wstring value_name;
    if (!ParseValueLine(line, &value_name)) {
      continue;
    }
    std::wstring lower = ToLower(value_name);
    if (wanted.find(lower) == wanted.end()) {
      skip_continuation = !line.empty() && line.back() == L'\\';
      continue;
    }
    out.append(line);
    out.append(L"\r\n");
    kept_value = true;
    if (!line.empty() && line.back() == L'\\') {
      keep_continuation = true;
    }
  }

  if (!kept_value) {
    if (error) {
      *error = L"No selected values were found in the export.";
    }
    return false;
  }
  *output = std::move(out);
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
  if (!ReadRegFileText(read_path, content, &is_utf16)) {
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

bool ImportRegFile(HWND owner, std::wstring* error) {
  std::wstring path;
  if (!PromptOpenFile(owner, L"Registry Files (*.reg)\0*.reg\0All Files (*.*)\0*.*\0", &path)) {
    return false;
  }
  return ImportRegFileFromPath(path, error);
}

bool ImportRegFileFromPath(const std::wstring& path, std::wstring* error) {
  if (path.empty()) {
    return false;
  }
  std::wstring args = L"import \"" + path + L"\"";
  return RunRegCommand(args, nullptr, error);
}

bool ExportRegFile(HWND owner, const std::wstring& key_path, std::wstring* error) {
  ExportOptions options;
  if (!PromptForExportOptions(owner, L"", &options)) {
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

  std::wstring output;
  bool output_utf16 = false;
  bool output_utf16_set = false;

  auto append_content = [&](const std::wstring& content, bool utf16) -> bool {
    if (!output_utf16_set) {
      output_utf16 = utf16;
      output_utf16_set = true;
    }
    return AppendRegContent(content, &output);
  };

  if (!value_names.empty()) {
    std::wstring content;
    bool utf16 = false;
    if (!ExportKeyToContent(base_key_path, false, &content, &utf16, error)) {
      return false;
    }
    std::wstring filtered;
    if (!FilterRegFileValues(content, value_names, &filtered, error)) {
      return false;
    }
    append_content(filtered, utf16);
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

  if (output.empty()) {
    if (error) {
      *error = L"No data to export.";
    }
    return false;
  }

  if (!WriteRegFileText(path, output, output_utf16)) {
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
