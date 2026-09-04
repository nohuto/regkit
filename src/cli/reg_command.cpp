// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "cli/reg_command.h"

#include "regfile/reg_file.h"
#include "regfile/registry_transfer.h"
#include "registry/value_format.h"
#include "win32/file_text.h"
#include "win32/registry_view.h"

#include <algorithm>
#include <cstdio>
#include <cwctype>

namespace regkit::cli {

namespace {

constexpr int kOk = 0;
constexpr int kFailed = 1;

bool g_console_ready = false;
HANDLE g_out = nullptr;
HANDLE g_err = nullptr;

void EnsureConsole() {
  if (g_console_ready) {
    return;
  }
  g_console_ready = true;
  g_out = GetStdHandle(STD_OUTPUT_HANDLE);
  g_err = GetStdHandle(STD_ERROR_HANDLE);
  const bool have_out = g_out && g_out != INVALID_HANDLE_VALUE;
  if (!have_out && !AttachConsole(ATTACH_PARENT_PROCESS)) {
    AllocConsole();
  } else if (!have_out) {
  }
  if (!have_out) {
    g_out = GetStdHandle(STD_OUTPUT_HANDLE);
    g_err = GetStdHandle(STD_ERROR_HANDLE);
  }
  if (!g_err || g_err == INVALID_HANDLE_VALUE) {
    g_err = g_out;
  }
}

void WriteTo(HANDLE handle, const std::wstring& text) {
  if (!handle || handle == INVALID_HANDLE_VALUE || text.empty()) {
    return;
  }
  DWORD written = 0;
  if (GetFileType(handle) == FILE_TYPE_CHAR) {
    WriteConsoleW(handle, text.c_str(), static_cast<DWORD>(text.size()),
                  &written, nullptr);
    return;
  }
  const std::string utf8 = util::WideToUtf8(text);
  WriteFile(handle, utf8.data(), static_cast<DWORD>(utf8.size()), &written,
            nullptr);
}

void Print(const std::wstring& text) {
  EnsureConsole();
  WriteTo(g_out, text + L"\r\n");
}

void PrintError(const std::wstring& text) {
  EnsureConsole();
  const bool prefixed = text.rfind(L"ERROR: ", 0) == 0;
  WriteTo(g_err, (prefixed ? text : L"ERROR: " + text) + L"\r\n");
}

std::wstring SystemMessage(LONG status) {
  wchar_t* buffer = nullptr;
  const DWORD length = FormatMessageW(
      FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
          FORMAT_MESSAGE_IGNORE_INSERTS,
      nullptr, static_cast<DWORD>(status), 0,
      reinterpret_cast<wchar_t*>(&buffer), 0, nullptr);
  std::wstring text = (length && buffer) ? buffer : L"The operation failed.";
  if (buffer) {
    LocalFree(buffer);
  }
  while (!text.empty() && (text.back() == L'\r' || text.back() == L'\n')) {
    text.pop_back();
  }
  return text;
}

int Fail(LONG status) {
  if (status == ERROR_FILE_NOT_FOUND) {
    PrintError(L"The system was unable to find the specified registry key or value.");
  } else {
    PrintError(SystemMessage(status));
  }
  return kFailed;
}

bool EqualsInsensitive(std::wstring_view left, const wchar_t* right) {
  return left.size() == wcslen(right) &&
         _wcsnicmp(left.data(), right, left.size()) == 0;
}

bool IsSwitch(const std::wstring& text, const wchar_t* name) {
  if (text.empty() || (text[0] != L'/' && text[0] != L'-')) {
    return false;
  }
  return EqualsInsensitive(std::wstring_view(text).substr(1), name);
}

struct KeyRef {
  HKEY root = nullptr;
  std::wstring subkey;
  std::wstring display;
};

HKEY RootFromName(std::wstring_view name) {
  struct Entry {
    const wchar_t* full;
    const wchar_t* shortcut;
    HKEY root;
  };
  static const Entry kRoots[] = {
      {L"HKEY_LOCAL_MACHINE", L"HKLM", HKEY_LOCAL_MACHINE},
      {L"HKEY_CURRENT_USER", L"HKCU", HKEY_CURRENT_USER},
      {L"HKEY_CLASSES_ROOT", L"HKCR", HKEY_CLASSES_ROOT},
      {L"HKEY_USERS", L"HKU", HKEY_USERS},
      {L"HKEY_CURRENT_CONFIG", L"HKCC", HKEY_CURRENT_CONFIG},
      {L"HKEY_PERFORMANCE_DATA", L"HKPD", HKEY_PERFORMANCE_DATA},
  };
  for (const Entry& entry : kRoots) {
    if (EqualsInsensitive(name, entry.full) ||
        EqualsInsensitive(name, entry.shortcut)) {
      return entry.root;
    }
  }
  return nullptr;
}

const wchar_t* RootName(HKEY root) {
  if (root == HKEY_LOCAL_MACHINE) return L"HKEY_LOCAL_MACHINE";
  if (root == HKEY_CURRENT_USER) return L"HKEY_CURRENT_USER";
  if (root == HKEY_CLASSES_ROOT) return L"HKEY_CLASSES_ROOT";
  if (root == HKEY_USERS) return L"HKEY_USERS";
  if (root == HKEY_CURRENT_CONFIG) return L"HKEY_CURRENT_CONFIG";
  if (root == HKEY_PERFORMANCE_DATA) return L"HKEY_PERFORMANCE_DATA";
  return L"HKEY";
}

bool ParseKey(const std::wstring& text, KeyRef* key) {
  std::wstring_view view = text;
  while (!view.empty() && view.front() == L'\\') {
    view.remove_prefix(1);
  }
  const size_t split = view.find(L'\\');
  const std::wstring_view root_name =
      split == std::wstring_view::npos ? view : view.substr(0, split);
  key->root = RootFromName(root_name);
  if (!key->root) {
    PrintError(L"Invalid key name: " + text);
    return false;
  }
  key->subkey = split == std::wstring_view::npos
                    ? std::wstring()
                    : std::wstring(view.substr(split + 1));
  key->display = RootName(key->root);
  if (!key->subkey.empty()) {
    key->display += L'\\';
    key->display += key->subkey;
  }
  return true;
}

struct ValueType {
  const wchar_t* name;
  DWORD type;
};

const ValueType kTypes[] = {
    {L"REG_SZ", REG_SZ},
    {L"REG_MULTI_SZ", REG_MULTI_SZ},
    {L"REG_EXPAND_SZ", REG_EXPAND_SZ},
    {L"REG_DWORD", REG_DWORD},
    {L"REG_DWORD_LITTLE_ENDIAN", REG_DWORD_LITTLE_ENDIAN},
    {L"REG_DWORD_BIG_ENDIAN", REG_DWORD_BIG_ENDIAN},
    {L"REG_QWORD", REG_QWORD},
    {L"REG_QWORD_LITTLE_ENDIAN", REG_QWORD_LITTLE_ENDIAN},
    {L"REG_BINARY", REG_BINARY},
    {L"REG_NONE", REG_NONE},
    {L"REG_LINK", REG_LINK},
};

bool ParseType(const std::wstring& text, DWORD* type) {
  for (const ValueType& entry : kTypes) {
    if (EqualsInsensitive(text, entry.name)) {
      *type = entry.type;
      return true;
    }
  }
  PrintError(L"Invalid type: " + text);
  return false;
}

bool BuildData(DWORD type, const std::wstring& text,
               const std::wstring& separator, std::vector<BYTE>* data) {
  data->clear();
  switch (type) {
  case REG_SZ:
  case REG_EXPAND_SZ:
  case REG_LINK: {
    const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    data->resize(bytes);
    memcpy(data->data(), text.c_str(), bytes);
    return true;
  }
  case REG_MULTI_SZ: {
    std::vector<std::wstring> items;
    const size_t step = separator.empty() ? 1 : separator.size();
    size_t start = 0;
    while (start <= text.size()) {
      const size_t end = separator.empty() ? std::wstring::npos
                                           : text.find(separator, start);
      const std::wstring item =
          text.substr(start, end == std::wstring::npos ? std::wstring::npos
                                                       : end - start);
      if (!item.empty()) {
        items.push_back(item);
      }
      if (end == std::wstring::npos) {
        break;
      }
      start = end + step;
    }
    *data = value_format::MultiStringData(items);
    return true;
  }
  case REG_DWORD:
  case REG_DWORD_BIG_ENDIAN:
  case REG_QWORD: {
    wchar_t* stop = nullptr;
    const int base = (text.rfind(L"0x", 0) == 0 || text.rfind(L"0X", 0) == 0) ? 16 : 10;
    const unsigned long long value = wcstoull(text.c_str(), &stop, base);
    if (stop == text.c_str()) {
      PrintError(L"Invalid numeric data: " + text);
      return false;
    }
    if (type == REG_QWORD) {
      data->resize(sizeof(unsigned long long));
      memcpy(data->data(), &value, sizeof(value));
      return true;
    }
    DWORD narrow = static_cast<DWORD>(value);
    if (type == REG_DWORD_BIG_ENDIAN) {
      narrow = _byteswap_ulong(narrow);
    }
    data->resize(sizeof(DWORD));
    memcpy(data->data(), &narrow, sizeof(narrow));
    return true;
  }
  case REG_BINARY:
  case REG_NONE:
  default:
    if (!value_format::ParseHex(text, data)) {
      PrintError(L"Invalid binary data: " + text);
      return false;
    }
    return true;
  }
}

std::wstring TypeName(DWORD type) {
  for (const ValueType& entry : kTypes) {
    if (entry.type == type) {
      return entry.name;
    }
  }
  return value_format::TypeName(type);
}

std::wstring FormatData(DWORD type, const BYTE* data, DWORD size) {
  switch (type) {
  case REG_SZ:
  case REG_EXPAND_SZ:
  case REG_LINK: {
    std::wstring text(reinterpret_cast<const wchar_t*>(data),
                      size / sizeof(wchar_t));
    while (!text.empty() && text.back() == L'\0') {
      text.pop_back();
    }
    return text;
  }
  case REG_MULTI_SZ: {
    std::wstring joined;
    for (const auto& item :
         value_format::MultiStringItems(std::vector<BYTE>(data, data + size))) {
      if (!joined.empty()) {
        joined += L"\\0";
      }
      joined += item;
    }
    return joined;
  }
  case REG_DWORD:
  case REG_DWORD_BIG_ENDIAN: {
    DWORD value = 0;
    if (size >= sizeof(value)) {
      memcpy(&value, data, sizeof(value));
    }
    if (type == REG_DWORD_BIG_ENDIAN) {
      value = _byteswap_ulong(value);
    }
    wchar_t buffer[24] = {};
    swprintf_s(buffer, L"0x%x", value);
    return buffer;
  }
  case REG_QWORD: {
    unsigned long long value = 0;
    if (size >= sizeof(value)) {
      memcpy(&value, data, sizeof(value));
    }
    wchar_t buffer[32] = {};
    swprintf_s(buffer, L"0x%llx", value);
    return buffer;
  }
  default: {
    std::wstring hex;
    hex.reserve(static_cast<size_t>(size) * 2);
    static const wchar_t digits[] = L"0123456789ABCDEF";
    for (DWORD i = 0; i < size; ++i) {
      hex.push_back(digits[data[i] >> 4]);
      hex.push_back(digits[data[i] & 0x0F]);
    }
    return hex;
  }
  }
}

struct Options {
  std::wstring value_name;
  std::wstring data;
  std::wstring type_text;
  std::wstring file;
  std::wstring separator = L"\\0";
  bool has_value = false;
  bool default_value = false;
  bool all_values = false;
  bool recurse = false;
  bool force = false;
  bool has_data = false;
  REGSAM view = win32::kDefaultRegistryView;
};

bool ParseOptions(const std::vector<std::wstring>& args, size_t first,
                  Options* options, std::vector<std::wstring>* positional,
                  bool separator_switch = false) {
  for (size_t i = first; i < args.size(); ++i) {
    const std::wstring& arg = args[i];
    auto next = [&](std::wstring* out) -> bool {
      if (i + 1 >= args.size()) {
        PrintError(L"Missing argument for " + arg);
        return false;
      }
      *out = args[++i];
      return true;
    };
    if (IsSwitch(arg, L"v")) {
      if (!next(&options->value_name)) return false;
      options->has_value = true;
    } else if (IsSwitch(arg, L"ve")) {
      options->default_value = true;
      options->has_value = true;
    } else if (IsSwitch(arg, L"va")) {
      options->all_values = true;
    } else if (IsSwitch(arg, L"t")) {
      if (!next(&options->type_text)) return false;
    } else if (IsSwitch(arg, L"d")) {
      if (!next(&options->data)) return false;
      options->has_data = true;
    } else if (IsSwitch(arg, L"s")) {
      if (separator_switch) {
        std::wstring separator;
        if (!next(&separator)) return false;
        options->separator = separator;
      } else {
        options->recurse = true;
      }
    } else if (IsSwitch(arg, L"f") || IsSwitch(arg, L"y")) {
      options->force = true;
    } else if (IsSwitch(arg, L"reg:32")) {
      options->view = KEY_WOW64_32KEY;
    } else if (IsSwitch(arg, L"reg:64")) {
      options->view = KEY_WOW64_64KEY;
    } else if (!arg.empty() && (arg[0] == L'/' || arg[0] == L'-')) {
      continue;
    } else {
      positional->push_back(arg);
    }
  }
  return true;
}

void CollectValues(HKEY root, const std::wstring& subkey, REGSAM view,
                   std::vector<RegistryValue>* values) {
  HKEY handle = nullptr;
  if (RegOpenKeyExW(root, subkey.c_str(), 0, KEY_READ | view, &handle) !=
      ERROR_SUCCESS) {
    return;
  }
  wchar_t name[16384] = {};
  std::vector<BYTE> data(4096);
  DWORD index = 0;
  while (true) {
    DWORD length = static_cast<DWORD>(std::size(name));
    DWORD size = static_cast<DWORD>(data.size());
    DWORD type = 0;
    const LONG status = RegEnumValueW(handle, index, name, &length, nullptr,
                                      &type, data.data(), &size);
    if (status == ERROR_MORE_DATA) {
      data.resize(size ? size : data.size() * 2);
      continue;
    }
    if (status != ERROR_SUCCESS) {
      break;
    }
    ++index;
    RegistryValue value;
    value.name.assign(name, length);
    value.type = type;
    value.data.assign(data.begin(), data.begin() + size);
    values->push_back(std::move(value));
  }
  RegCloseKey(handle);
}

LONG OpenKey(const KeyRef& key, REGSAM access, REGSAM view, HKEY* handle) {
  return RegOpenKeyExW(key.root, key.subkey.c_str(), 0, access | view, handle);
}

int QueryKey(const KeyRef& key, const Options& options, bool recurse,
             bool* matched) {
  HKEY handle = nullptr;
  LONG status = OpenKey(key, KEY_READ, options.view, &handle);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }

  Print(key.display);
  struct Entry {
    std::wstring name;
    DWORD type = 0;
    std::vector<BYTE> data;
  };
  std::vector<Entry> entries;
  DWORD index = 0;
  wchar_t name[16384] = {};
  DWORD name_length = 0;
  DWORD type = 0;
  std::vector<BYTE> data(4096);
  DWORD size = 0;
  bool printed = false;
  while (true) {
    name_length = static_cast<DWORD>(std::size(name));
    size = static_cast<DWORD>(data.size());
    status = RegEnumValueW(handle, index, name, &name_length, nullptr, &type,
                           data.data(), &size);
    if (status == ERROR_MORE_DATA) {
      data.resize(size ? size : data.size() * 2);
      continue;
    }
    if (status != ERROR_SUCCESS) {
      break;
    }
    ++index;
    const std::wstring value_name(name, name_length);
    if (options.has_value) {
      if (options.default_value ? !value_name.empty()
                                : _wcsicmp(value_name.c_str(),
                                           options.value_name.c_str()) != 0) {
        continue;
      }
    }
    Entry entry;
    entry.name = value_name;
    entry.type = type;
    entry.data.assign(data.begin(), data.begin() + size);
    entries.push_back(std::move(entry));
  }
  std::stable_partition(entries.begin(), entries.end(),
                        [](const Entry& entry) { return entry.name.empty(); });
  for (const Entry& entry : entries) {
    Print(L"    " + (entry.name.empty() ? std::wstring(L"(Default)") : entry.name) +
          L"    " + TypeName(entry.type) + L"    " +
          FormatData(entry.type, entry.data.data(),
                     static_cast<DWORD>(entry.data.size())));
    printed = true;
  }
  if (printed && matched) {
    *matched = true;
  }
  if (!options.has_value || printed) {
    Print(L"");
  }

  std::vector<std::wstring> children;
  index = 0;
  while (true) {
    name_length = static_cast<DWORD>(std::size(name));
    status = RegEnumKeyExW(handle, index, name, &name_length, nullptr, nullptr,
                           nullptr, nullptr);
    if (status != ERROR_SUCCESS) {
      break;
    }
    ++index;
    children.emplace_back(name, name_length);
  }
  if (!options.has_value && !recurse) {
    for (const auto& child : children) {
      Print(key.display + L'\\' + child);
    }
    if (!children.empty()) {
      Print(L"");
    }
  }
  RegCloseKey(handle);

  if (recurse) {
    for (const auto& child : children) {
      KeyRef sub = key;
      sub.subkey = key.subkey.empty() ? child : key.subkey + L'\\' + child;
      sub.display = key.display + L'\\' + child;
      QueryKey(sub, options, true, matched);
    }
  }
  return kOk;
}

int CmdQuery(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional)) {
    return kFailed;
  }
  if (positional.empty()) {
    PrintError(L"reg query requires a key name.");
    return kFailed;
  }
  if (positional.size() > 1) {
    PrintError(L"Invalid syntax.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(positional[0], &key)) {
    return kFailed;
  }
  bool matched = false;
  const int result = QueryKey(key, options, options.recurse, &matched);
  if (result == kOk && options.has_value && !matched) {
    return Fail(ERROR_FILE_NOT_FOUND);
  }
  return result;
}

int CmdAdd(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional, true)) {
    return kFailed;
  }
  if (positional.empty()) {
    PrintError(L"reg add requires a key name.");
    return kFailed;
  }
  if (positional.size() > 1) {
    PrintError(L"Invalid syntax.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(positional[0], &key)) {
    return kFailed;
  }

  HKEY handle = nullptr;
  DWORD disposition = 0;
  LONG status = RegCreateKeyExW(key.root, key.subkey.c_str(), 0, nullptr,
                                REG_OPTION_NON_VOLATILE,
                                KEY_WRITE | KEY_QUERY_VALUE | options.view,
                                nullptr, &handle, &disposition);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }

  int result = kOk;
  if (options.has_value) {
    const std::wstring name = options.default_value ? std::wstring() : options.value_name;
    DWORD type = REG_SZ;
    if (!options.type_text.empty() && !ParseType(options.type_text, &type)) {
      RegCloseKey(handle);
      return kFailed;
    }
    if (!options.force) {
      DWORD existing = 0;
      if (RegQueryValueExW(handle, name.c_str(), nullptr, &existing, nullptr,
                           nullptr) == ERROR_SUCCESS) {
        Print(L"Value " + (name.empty() ? std::wstring(L"(Default)") : name) +
              L" already exists. Use /f to overwrite.");
        RegCloseKey(handle);
        return kFailed;
      }
    }
    std::vector<BYTE> data;
    if (!BuildData(type, options.has_data ? options.data : std::wstring(),
                   options.separator, &data)) {
      RegCloseKey(handle);
      return kFailed;
    }
    status = RegSetValueExW(handle, name.c_str(), 0, type, data.data(),
                            static_cast<DWORD>(data.size()));
    if (status != ERROR_SUCCESS) {
      result = Fail(status);
    }
  }
  RegCloseKey(handle);
  if (result == kOk) {
    Print(L"The operation completed successfully.");
  }
  return result;
}

LONG DeleteTree(HKEY root, const std::wstring& subkey, REGSAM view) {
  HKEY handle = nullptr;
  LONG status = RegOpenKeyExW(root, subkey.c_str(), 0,
                              KEY_READ | KEY_WRITE | view, &handle);
  if (status != ERROR_SUCCESS) {
    return status;
  }
  std::vector<std::wstring> children;
  wchar_t name[512] = {};
  DWORD index = 0;
  while (true) {
    DWORD length = static_cast<DWORD>(std::size(name));
    if (RegEnumKeyExW(handle, index, name, &length, nullptr, nullptr, nullptr,
                      nullptr) != ERROR_SUCCESS) {
      break;
    }
    ++index;
    children.emplace_back(name, length);
  }
  RegCloseKey(handle);
  for (const auto& child : children) {
    DeleteTree(root, subkey + L'\\' + child, view);
  }
  return RegDeleteKeyExW(root, subkey.c_str(), view, 0);
}

int CmdDelete(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional)) {
    return kFailed;
  }
  if (positional.empty()) {
    PrintError(L"reg delete requires a key name.");
    return kFailed;
  }
  if (positional.size() > 1) {
    PrintError(L"Invalid syntax.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(positional[0], &key)) {
    return kFailed;
  }

  if (options.has_value || options.all_values) {
    HKEY handle = nullptr;
    LONG status = OpenKey(key, KEY_WRITE, options.view, &handle);
    if (status != ERROR_SUCCESS) {
      return Fail(status);
    }
    if (options.all_values) {
      wchar_t name[16384] = {};
      while (true) {
        DWORD length = static_cast<DWORD>(std::size(name));
        if (RegEnumValueW(handle, 0, name, &length, nullptr, nullptr, nullptr,
                          nullptr) != ERROR_SUCCESS) {
          break;
        }
        if (RegDeleteValueW(handle, name) != ERROR_SUCCESS) {
          break;
        }
      }
    } else {
      const std::wstring name =
          options.default_value ? std::wstring() : options.value_name;
      status = RegDeleteValueW(handle, name.c_str());
      if (status != ERROR_SUCCESS) {
        RegCloseKey(handle);
        return Fail(status);
      }
    }
    RegCloseKey(handle);
    Print(L"The operation completed successfully.");
    return kOk;
  }

  if (key.subkey.empty()) {
    PrintError(L"Refusing to delete a registry root.");
    return kFailed;
  }
  const LONG status = DeleteTree(key.root, key.subkey, options.view);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

LONG CopyTree(const KeyRef& from, const KeyRef& to, REGSAM view, bool recurse) {
  HKEY source = nullptr;
  LONG status = RegOpenKeyExW(from.root, from.subkey.c_str(), 0,
                              KEY_READ | view, &source);
  if (status != ERROR_SUCCESS) {
    return status;
  }
  HKEY target = nullptr;
  status = RegCreateKeyExW(to.root, to.subkey.c_str(), 0, nullptr,
                           REG_OPTION_NON_VOLATILE, KEY_WRITE | view, nullptr,
                           &target, nullptr);
  if (status != ERROR_SUCCESS) {
    RegCloseKey(source);
    return status;
  }

  wchar_t name[16384] = {};
  std::vector<BYTE> data(4096);
  DWORD index = 0;
  while (true) {
    DWORD length = static_cast<DWORD>(std::size(name));
    DWORD size = static_cast<DWORD>(data.size());
    DWORD type = 0;
    status = RegEnumValueW(source, index, name, &length, nullptr, &type,
                           data.data(), &size);
    if (status == ERROR_MORE_DATA) {
      data.resize(size ? size : data.size() * 2);
      continue;
    }
    if (status != ERROR_SUCCESS) {
      break;
    }
    ++index;
    RegSetValueExW(target, name, 0, type, data.data(), size);
  }

  std::vector<std::wstring> children;
  index = 0;
  while (recurse) {
    DWORD length = static_cast<DWORD>(std::size(name));
    if (RegEnumKeyExW(source, index, name, &length, nullptr, nullptr, nullptr,
                      nullptr) != ERROR_SUCCESS) {
      break;
    }
    ++index;
    children.emplace_back(name, length);
  }
  RegCloseKey(source);
  RegCloseKey(target);

  for (const auto& child : children) {
    KeyRef child_from = from;
    KeyRef child_to = to;
    child_from.subkey = from.subkey + L'\\' + child;
    child_to.subkey = to.subkey + L'\\' + child;
    CopyTree(child_from, child_to, view, true);
  }
  return ERROR_SUCCESS;
}

int CmdCopy(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional)) {
    return kFailed;
  }
  if (positional.size() != 2) {
    PrintError(positional.size() < 2 ? L"reg copy requires a source and a destination key."
                                     : L"Invalid syntax.");
    return kFailed;
  }
  KeyRef from;
  KeyRef to;
  if (!ParseKey(positional[0], &from) || !ParseKey(positional[1], &to)) {
    return kFailed;
  }
  const LONG status = CopyTree(from, to, options.view, options.recurse);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

bool ExportKeyToFile(const KeyRef& key, const std::wstring& path,
                     REGSAM view, std::wstring* error) {
  regfile::Writer writer;
  std::vector<KeyRef> pending{key};
  bool any = false;
  while (!pending.empty()) {
    const KeyRef current = pending.back();
    pending.pop_back();
    HKEY handle = nullptr;
    if (RegOpenKeyExW(current.root, current.subkey.c_str(), 0, KEY_READ | view,
                      &handle) != ERROR_SUCCESS) {
      continue;
    }
    any = true;

    std::vector<RegistryValue> values;
    CollectValues(current.root, current.subkey, view, &values);
    std::stable_partition(values.begin(), values.end(),
                          [](const RegistryValue& value) { return value.name.empty(); });
    std::vector<const regfile::Value*> pointers;
    pointers.reserve(values.size());
    for (const RegistryValue& value : values) {
      pointers.push_back(&value);
    }
    writer.AppendKey(current.display, std::move(pointers), false);

    wchar_t name[512] = {};
    DWORD index = 0;
    std::vector<std::wstring> children;
    while (true) {
      DWORD length = static_cast<DWORD>(std::size(name));
      if (RegEnumKeyExW(handle, index, name, &length, nullptr, nullptr, nullptr,
                        nullptr) != ERROR_SUCCESS) {
        break;
      }
      ++index;
      children.emplace_back(name, length);
    }
    RegCloseKey(handle);
    for (auto it = children.rbegin(); it != children.rend(); ++it) {
      KeyRef child = current;
      child.subkey = current.subkey.empty() ? *it : current.subkey + L'\\' + *it;
      child.display = current.display + L'\\' + *it;
      pending.push_back(child);
    }
  }
  if (!any) {
    if (error) {
      *error = L"The key does not exist: " + key.display;
    }
    return false;
  }
  if (!util::WriteTextFile(path, std::move(writer).Finish(), true)) {
    if (error) {
      *error = L"Failed to write " + path;
    }
    return false;
  }
  return true;
}

int CmdExport(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional)) {
    return kFailed;
  }
  if (positional.size() != 2) {
    PrintError(positional.size() < 2 ? L"reg export requires a key name and a file name."
                                     : L"Invalid syntax.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(positional[0], &key)) {
    return kFailed;
  }
  if (!options.force && GetFileAttributesW(positional[1].c_str()) != INVALID_FILE_ATTRIBUTES) {
    PrintError(positional[1] + L" already exists. Use /y to overwrite.");
    return kFailed;
  }
  std::wstring error;
  if (!ExportKeyToFile(key, positional[1], options.view, &error)) {
    PrintError(error.empty() ? L"Export failed." : error);
    return kFailed;
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

int CmdImport(const std::vector<std::wstring>& args) {
  if (args.size() < 2) {
    PrintError(L"reg import requires a file name.");
    return kFailed;
  }
  std::wstring error;
  if (!ImportRegFileFromPath(args[1], &error)) {
    PrintError(error.empty() ? L"Import failed." : error);
    return kFailed;
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

bool EnablePrivilege(const wchar_t* name) {
  HANDLE token = nullptr;
  if (!OpenProcessToken(GetCurrentProcess(),
                        TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &token)) {
    return false;
  }
  TOKEN_PRIVILEGES privileges = {};
  privileges.PrivilegeCount = 1;
  privileges.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
  bool ok = LookupPrivilegeValueW(nullptr, name,
                                  &privileges.Privileges[0].Luid) != FALSE;
  if (ok) {
    ok = AdjustTokenPrivileges(token, FALSE, &privileges, 0, nullptr, nullptr) &&
         GetLastError() == ERROR_SUCCESS;
  }
  CloseHandle(token);
  return ok;
}

int CmdSave(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional)) {
    return kFailed;
  }
  if (positional.size() < 2) {
    PrintError(L"reg save requires a key name and a file name.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(positional[0], &key)) {
    return kFailed;
  }
  if (!options.force && GetFileAttributesW(positional[1].c_str()) != INVALID_FILE_ATTRIBUTES) {
    PrintError(positional[1] + L" already exists. Use /y to overwrite.");
    return kFailed;
  }
  EnablePrivilege(SE_BACKUP_NAME);
  HKEY handle = nullptr;
  LONG status = OpenKey(key, KEY_READ, options.view, &handle);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  DeleteFileW(positional[1].c_str());
  status = RegSaveKeyW(handle, positional[1].c_str(), nullptr);
  RegCloseKey(handle);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

int CmdRestore(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional)) {
    return kFailed;
  }
  if (positional.size() < 2) {
    PrintError(L"reg restore requires a key name and a file name.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(positional[0], &key)) {
    return kFailed;
  }
  if (!EnablePrivilege(SE_RESTORE_NAME) ||
      !EnablePrivilege(SE_BACKUP_NAME)) {
    return Fail(ERROR_PRIVILEGE_NOT_HELD);
  }
  HKEY handle = nullptr;
  LONG status = OpenKey(key, KEY_WRITE, options.view, &handle);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  status = RegRestoreKeyW(handle, positional[1].c_str(), REG_FORCE_RESTORE);
  RegCloseKey(handle);
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

int CmdLoad(const std::vector<std::wstring>& args) {
  if (args.size() < 3) {
    PrintError(L"reg load requires a key name and a file name.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(args[1], &key)) {
    return kFailed;
  }
  EnablePrivilege(SE_RESTORE_NAME);
  EnablePrivilege(SE_BACKUP_NAME);
  const LONG status = RegLoadKeyW(key.root, key.subkey.c_str(), args[2].c_str());
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

int CmdUnload(const std::vector<std::wstring>& args) {
  if (args.size() < 2) {
    PrintError(L"reg unload requires a key name.");
    return kFailed;
  }
  KeyRef key;
  if (!ParseKey(args[1], &key)) {
    return kFailed;
  }
  EnablePrivilege(SE_RESTORE_NAME);
  EnablePrivilege(SE_BACKUP_NAME);
  const LONG status = RegUnLoadKeyW(key.root, key.subkey.c_str());
  if (status != ERROR_SUCCESS) {
    return Fail(status);
  }
  Print(L"The operation completed successfully.");
  return kOk;
}

int CmdCompare(const std::vector<std::wstring>& args) {
  Options options;
  std::vector<std::wstring> positional;
  if (!ParseOptions(args, 1, &options, &positional)) {
    return kFailed;
  }
  if (positional.size() < 2) {
    PrintError(L"reg compare requires two key names.");
    return kFailed;
  }
  KeyRef left;
  KeyRef right;
  if (!ParseKey(positional[0], &left) || !ParseKey(positional[1], &right)) {
    return kFailed;
  }

  std::vector<RegistryValue> left_values;
  std::vector<RegistryValue> right_values;
  CollectValues(left.root, left.subkey, options.view, &left_values);
  CollectValues(right.root, right.subkey, options.view, &right_values);

  auto find = [](const std::vector<RegistryValue>& list,
                 const std::wstring& name) -> const RegistryValue* {
    for (const RegistryValue& value : list) {
      if (_wcsicmp(value.name.c_str(), name.c_str()) == 0) {
        return &value;
      }
    }
    return nullptr;
  };

  bool identical = true;
  for (const RegistryValue& value : left_values) {
    const RegistryValue* other = find(right_values, value.name);
    if (!other) {
      Print(L"< " + left.display + L"    " + value.name);
      identical = false;
    } else if (other->type != value.type || other->data != value.data) {
      Print(L"< " + left.display + L"    " + value.name + L"    " +
            FormatData(value.type, value.data.data(),
                       static_cast<DWORD>(value.data.size())));
      Print(L"> " + right.display + L"    " + other->name + L"    " +
            FormatData(other->type, other->data.data(),
                       static_cast<DWORD>(other->data.size())));
      identical = false;
    }
  }
  for (const RegistryValue& value : right_values) {
    if (!find(left_values, value.name)) {
      Print(L"> " + right.display + L"    " + value.name);
      identical = false;
    }
  }
  Print(identical ? L"Result Compared: Identical"
                  : L"Result Compared: Different");
  return identical ? kOk : 2;
}

void PrintUsage() {
  Print(
      L"RegKit command line\n"
      L"\n"
      L"regedit compatible:\n"
      L"  regkit file.reg                 import a .reg file (asks first)\n"
      L"  regkit /s file.reg              import without prompting\n"
      L"  regkit /e file.reg [key]        export a key (or everything)\n"
      L"  regkit /a file.reg [key]        same as /e, kept for compatibility\n"
      L"  regkit /c /m /l:file /r:file    accepted and ignored (legacy)\n"
      L"\n"
      L"reg.exe compatible (the leading \"reg\" is optional):\n"
      L"  add <key> [/v name | /ve] [/t type] [/s sep] [/d data] [/f]\n"
      L"  delete <key> [/v name | /ve | /va] [/f]\n"
      L"  query <key> [/v name | /ve] [/s]\n"
      L"  copy <src> <dst> [/s] [/f]\n"
      L"  export <key> <file.reg> [/y]\n"
      L"  import <file.reg>\n"
      L"  save <key> <file.hiv> [/y]      restore <key> <file.hiv>\n"
      L"  load <key> <file.hiv>           unload <key>\n"
      L"  compare <key1> <key2> [/s]\n"
      L"  /reg:32 | /reg:64               pick the registry view\n"
      L"\n"
      L"RegKit additions:\n"
      L"  regkit <key>                    open the window at that key\n"
      L"  regkit --goto <key>             same, explicit form\n"
      L"  regkit --restart-system         relaunch as SYSTEM\n"
      L"  regkit --restart-ti             relaunch as TrustedInstaller\n"
      L"  regkit --help                   show this text\n"
      L"\n"
      L"Key names accept HKLM, HKCU, HKCR, HKU, HKCC and their full forms.\n"
      L"reg flags is not implemented, every other verb above is.");
}

int RunVerb(const std::wstring& verb, const std::vector<std::wstring>& args) {
  if (EqualsInsensitive(verb, L"query")) return CmdQuery(args);
  if (EqualsInsensitive(verb, L"add")) return CmdAdd(args);
  if (EqualsInsensitive(verb, L"delete")) return CmdDelete(args);
  if (EqualsInsensitive(verb, L"copy")) return CmdCopy(args);
  if (EqualsInsensitive(verb, L"export")) return CmdExport(args);
  if (EqualsInsensitive(verb, L"import")) return CmdImport(args);
  if (EqualsInsensitive(verb, L"save")) return CmdSave(args);
  if (EqualsInsensitive(verb, L"restore")) return CmdRestore(args);
  if (EqualsInsensitive(verb, L"load")) return CmdLoad(args);
  if (EqualsInsensitive(verb, L"unload")) return CmdUnload(args);
  if (EqualsInsensitive(verb, L"compare")) return CmdCompare(args);
  if (EqualsInsensitive(verb, L"flags")) {
    PrintError(L"reg flags is not implemented.");
    return kFailed;
  }
  return -1;
}

} // namespace

bool Execute(const std::vector<std::wstring>& args, int* exit_code) {
  if (args.empty() || !exit_code) {
    return false;
  }

  for (const std::wstring& arg : args) {
    if (IsSwitch(arg, L"?") || IsSwitch(arg, L"h") || IsSwitch(arg, L"-help") ||
        IsSwitch(arg, L"help")) {
      PrintUsage();
      *exit_code = kOk;
      return true;
    }
  }

  std::vector<std::wstring> verb_args = args;
  if (EqualsInsensitive(verb_args[0], L"reg")) {
    verb_args.erase(verb_args.begin());
  }
  if (!verb_args.empty()) {
    const int result = RunVerb(verb_args[0], verb_args);
    if (result >= 0) {
      *exit_code = result;
      return true;
    }
  }

  for (size_t i = 0; i < args.size(); ++i) {
    if (IsSwitch(args[i], L"s") && i + 1 < args.size()) {
      std::wstring error;
      if (!ImportRegFileFromPath(args[i + 1], &error)) {
        PrintError(error.empty() ? L"Import failed." : error);
        *exit_code = kFailed;
      } else {
        *exit_code = kOk;
      }
      return true;
    }
    if ((IsSwitch(args[i], L"e") || IsSwitch(args[i], L"a")) &&
        i + 1 < args.size()) {
      const std::wstring name = (i + 2 < args.size()) ? args[i + 2] : std::wstring();
      std::wstring error;
      bool ok = true;
      if (name.empty()) {
        static const wchar_t* kAll[] = {L"HKEY_CLASSES_ROOT", L"HKEY_CURRENT_USER",
                                        L"HKEY_LOCAL_MACHINE", L"HKEY_USERS",
                                        L"HKEY_CURRENT_CONFIG"};
        PrintError(L"Exporting the whole registry is not supported, name a key.");
        (void)kAll;
        ok = false;
      } else {
        KeyRef key;
        ok = ParseKey(name, &key) && ExportKeyToFile(key, args[i + 1], 0, &error);
      }
      if (!ok) {
        PrintError(error.empty() ? L"Export failed." : error);
        *exit_code = kFailed;
      } else {
        *exit_code = kOk;
      }
      return true;
    }
  }
  return false;
}

} // namespace regkit::cli
