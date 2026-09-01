// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "appearance/default_font.h"

#include "win32/registry_view.h"

#include "appearance/font_metrics.h"
#include "win32/file_text.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"

#include <limits>
#include <string>
#include <vector>

namespace regkit::ui {

namespace {

using util::TrimWhitespace;

struct FontSettings {
  bool use_default = true;
  std::wstring face;
  int size = 0;
  int weight = 0;
  bool italic = false;
  bool italic_set = false;
};

bool ParseBool(const std::wstring& value) {
  return (_wcsicmp(value.c_str(), L"1") == 0 || _wcsicmp(value.c_str(), L"true") == 0 || _wcsicmp(value.c_str(), L"yes") == 0);
}

bool LoadFontSettings(FontSettings* out) {
  if (!out) {
    return false;
  }
  std::wstring folder = util::GetAppDataFolder();
  if (folder.empty()) {
    return false;
  }
  std::wstring path = util::JoinPath(folder, L"settings.ini");
  HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }
  LARGE_INTEGER size = {};
  if (!GetFileSizeEx(file, &size) || size.QuadPart <= 0 || size.QuadPart > static_cast<LONGLONG>(std::numeric_limits<int>::max())) {
    CloseHandle(file);
    return false;
  }
  std::string buffer(static_cast<size_t>(size.QuadPart), '\0');
  DWORD read = 0;
  bool ok = ReadFile(file, buffer.data(), static_cast<DWORD>(buffer.size()), &read, nullptr) != 0;
  CloseHandle(file);
  if (!ok || read == 0) {
    return false;
  }
  buffer.resize(read);
  if (buffer.size() >= 3 && static_cast<unsigned char>(buffer[0]) == 0xEF && static_cast<unsigned char>(buffer[1]) == 0xBB && static_cast<unsigned char>(buffer[2]) == 0xBF) {
    buffer.erase(0, 3);
  }
  std::wstring content = util::Utf8ToWide(buffer);
  if (content.empty()) {
    return false;
  }

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
    if (line.empty()) {
      continue;
    }
    size_t sep = line.find(L'=');
    if (sep == std::wstring::npos) {
      continue;
    }
    std::wstring key = TrimWhitespace(line.substr(0, sep));
    std::wstring value = TrimWhitespace(line.substr(sep + 1));
    if (_wcsicmp(key.c_str(), L"font_use_default") == 0) {
      out->use_default = ParseBool(value);
    } else if (_wcsicmp(key.c_str(), L"font_face") == 0) {
      out->face = value;
    } else if (_wcsicmp(key.c_str(), L"font_size") == 0) {
      int size_value = _wtoi(value.c_str());
      if (size_value > 0) {
        out->size = size_value;
      }
    } else if (_wcsicmp(key.c_str(), L"font_weight") == 0) {
      int weight_value = _wtoi(value.c_str());
      if (weight_value > 0) {
        out->weight = weight_value;
      }
    } else if (_wcsicmp(key.c_str(), L"font_italic") == 0) {
      out->italic = ParseBool(value);
      out->italic_set = true;
    }
  }
  return true;
}

std::wstring ReadFontSubstitute(const wchar_t* value_name) {
  if (!value_name || !*value_name) {
    return L"";
  }
  const wchar_t* subkey = L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\FontSubstitutes";
  auto query = [&](REGSAM sam) -> std::wstring {
    HKEY key = nullptr;
    LONG result = RegOpenKeyExW(HKEY_LOCAL_MACHINE, subkey, 0, sam, &key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    DWORD type = 0;
    DWORD bytes = 0;
    result = RegQueryValueExW(key, value_name, nullptr, &type, nullptr, &bytes);
    if (result != ERROR_SUCCESS || bytes == 0 || (type != REG_SZ && type != REG_EXPAND_SZ)) {
      RegCloseKey(key);
      return L"";
    }
    std::vector<wchar_t> buffer(bytes / sizeof(wchar_t) + 1, L'\0');
    result = RegQueryValueExW(key, value_name, nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &bytes);
    RegCloseKey(key);
    if (result != ERROR_SUCCESS) {
      return L"";
    }
    std::wstring value(buffer.data());
    while (!value.empty() && value.back() == L'\0') {
      value.pop_back();
    }
    if (value.empty()) {
      return L"";
    }
    if (type == REG_EXPAND_SZ) {
      std::wstring expanded = util::ExpandEnvironmentStringsDynamic(value);
      if (!expanded.empty()) {
        value = std::move(expanded);
      }
    }
    return value;
  };

  std::wstring value = query(KEY_READ | win32::kDefaultRegistryView);
  if (!value.empty()) {
    return value;
  }
  return query(KEY_READ);
}

} // namespace

LOGFONTW DefaultUIFontLogFont() {
  LOGFONTW lf = {};
  HFONT stock = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
  if (stock) {
    GetObjectW(stock, sizeof(lf), &lf);
  }
  if (lf.lfHeight == 0) {
    lf.lfHeight = appearance::FontHeight(9);
    lf.lfWeight = FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
  }
  std::wstring default_face = ReadFontSubstitute(L"Segoe UI");
  if (default_face.empty()) {
    default_face = L"Segoe UI";
  }
  wcsncpy_s(lf.lfFaceName, default_face.c_str(), _TRUNCATE);

  FontSettings settings;
  if (LoadFontSettings(&settings) && !settings.use_default) {
    if (!settings.face.empty()) {
      wcsncpy_s(lf.lfFaceName, settings.face.c_str(), _TRUNCATE);
    }
    if (settings.size > 0) {
      lf.lfHeight = appearance::FontHeight(settings.size);
    }
    if (settings.weight > 0) {
      lf.lfWeight = settings.weight;
    }
    if (settings.italic_set) {
      lf.lfItalic = settings.italic ? TRUE : FALSE;
    }
  }
  return lf;
}

HFONT DefaultUIFont() {
  LOGFONTW lf = DefaultUIFontLogFont();
  HFONT font = CreateFontIndirectW(&lf);
  return font;
}

} // namespace regkit::ui
