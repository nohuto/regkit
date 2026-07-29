// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "registry/value_format.h"

#include "win32/win32_helpers.h"

#include <shlwapi.h>

#include <cstring>
#include <cwctype>

namespace regkit::value_format {

DWORD NormalizeType(DWORD type) {
  const DWORD base = type & 0xFFFF;
  switch (base) {
  case REG_NONE:
  case REG_SZ:
  case REG_EXPAND_SZ:
  case REG_MULTI_SZ:
  case REG_DWORD:
  case REG_QWORD:
  case REG_BINARY:
  case REG_RESOURCE_LIST:
  case REG_FULL_RESOURCE_DESCRIPTOR:
  case REG_RESOURCE_REQUIREMENTS_LIST:
  case REG_LINK:
  case REG_DWORD_BIG_ENDIAN:
    return base;
  default:
    return type;
  }
}

std::wstring TypeName(DWORD type) {
  const DWORD base = NormalizeType(type);
  const bool has_flags = base != type;
  const wchar_t* label = nullptr;
  switch (base) {
  case REG_NONE: label = L"REG_NONE"; break;
  case REG_SZ: label = L"REG_SZ"; break;
  case REG_EXPAND_SZ: label = L"REG_EXPAND_SZ"; break;
  case REG_MULTI_SZ: label = L"REG_MULTI_SZ"; break;
  case REG_DWORD: label = L"REG_DWORD"; break;
  case REG_QWORD: label = L"REG_QWORD"; break;
  case REG_BINARY: label = L"REG_BINARY"; break;
  case REG_RESOURCE_LIST: label = L"REG_RESOURCE_LIST"; break;
  case REG_FULL_RESOURCE_DESCRIPTOR:
    label = L"REG_FULL_RESOURCE_DESCRIPTOR";
    break;
  case REG_RESOURCE_REQUIREMENTS_LIST:
    label = L"REG_RESOURCE_REQUIREMENTS_LIST";
    break;
  case REG_LINK: label = L"REG_LINK"; break;
  case REG_DWORD_BIG_ENDIAN: label = L"REG_DWORD_BIG_ENDIAN"; break;
  default: break;
  }
  wchar_t buffer[64] = {};
  if (!label) {
    swprintf_s(buffer, L"REG_UNKNOWN (0x%X)", type);
    return buffer;
  }
  if (has_flags) {
    swprintf_s(buffer, L"%s (0x%X)", label, type);
    return buffer;
  }
  return label;
}

std::wstring Data(DWORD type, const BYTE* data, DWORD size) {
  if (!data || size == 0) {
    return {};
  }
  switch (NormalizeType(type)) {
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
    const wchar_t* current = reinterpret_cast<const wchar_t*>(data);
    size_t remaining = size / sizeof(wchar_t);
    while (remaining > 0 && *current) {
      const size_t length = wcsnlen_s(current, remaining);
      if (length == remaining) {
        break;
      }
      if (!joined.empty()) {
        joined += L"; ";
      }
      joined.append(current, length);
      current += length + 1;
      remaining -= length + 1;
    }
    return joined;
  }
  case REG_DWORD:
    if (size >= sizeof(DWORD)) {
      DWORD value = 0;
      std::memcpy(&value, data, sizeof(value));
      wchar_t buffer[32] = {};
      swprintf_s(buffer, L"0x%08X (%u)", value, value);
      return buffer;
    }
    break;
  case REG_QWORD:
    if (size >= sizeof(unsigned long long)) {
      unsigned long long value = 0;
      std::memcpy(&value, data, sizeof(value));
      wchar_t buffer[48] = {};
      swprintf_s(buffer, L"0x%016llX (%llu)", value, value);
      return buffer;
    }
    break;
  default:
    break;
  }
  return util::ToHex(data, size, 32);
}

std::wstring DisplayData(DWORD type, const BYTE* data, DWORD size) {
  const DWORD base_type = NormalizeType(type);
  std::wstring value = Data(type, data, size);
  if (value.empty()) {
    return value;
  }
  if ((base_type == REG_SZ || base_type == REG_EXPAND_SZ) &&
      value.front() == L'@') {
    std::wstring resolved(1024, L'\0');
    HRESULT result =
        SHLoadIndirectString(value.c_str(), resolved.data(),
                             static_cast<UINT>(resolved.size()), nullptr);
    if (result == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER)) {
      resolved.assign(4096, L'\0');
      result = SHLoadIndirectString(value.c_str(), resolved.data(),
                                    static_cast<UINT>(resolved.size()), nullptr);
    }
    if (SUCCEEDED(result)) {
      while (!resolved.empty() && resolved.back() == L'\0') {
        resolved.pop_back();
      }
      if (!resolved.empty()) {
        return resolved;
      }
    }
  }
  if (base_type == REG_EXPAND_SZ) {
    std::wstring expanded = util::ExpandEnvironmentStringsDynamic(value);
    if (!expanded.empty() && expanded != value) {
      return expanded;
    }
  }
  return value;
}

bool ParseHex(std::wstring_view text, std::vector<BYTE>* output) {
  if (!output) {
    return false;
  }
  output->clear();
  int high = -1;
  for (wchar_t character : text) {
    if (!iswxdigit(character)) {
      continue;
    }
    int value = character <= L'9'
                    ? character - L'0'
                    : 10 + (towlower(character) - L'a');
    if (high < 0) {
      high = value;
    } else {
      output->push_back(static_cast<BYTE>((high << 4) | value));
      high = -1;
    }
  }
  return high < 0;
}

std::vector<BYTE> StringData(std::wstring_view text) {
  std::vector<BYTE> data((text.size() + 1) * sizeof(wchar_t));
  std::memcpy(data.data(), text.data(), text.size() * sizeof(wchar_t));
  return data;
}

bool DecodeString(std::span<const BYTE> data, std::wstring* output) {
  if (!output || data.size() % sizeof(wchar_t) != 0) {
    return false;
  }
  output->assign(reinterpret_cast<const wchar_t*>(data.data()),
                 data.size() / sizeof(wchar_t));
  while (!output->empty() && output->back() == L'\0') {
    output->pop_back();
  }
  return output->find(L'\0') == std::wstring::npos;
}

std::vector<std::wstring> MultiStringItems(const std::vector<BYTE>& data) {
  std::vector<std::wstring> items;
  if (data.size() % sizeof(wchar_t) != 0) {
    return items;
  }
  const wchar_t* current = reinterpret_cast<const wchar_t*>(data.data());
  size_t remaining = data.size() / sizeof(wchar_t);
  while (remaining > 0 && *current) {
    const size_t length = wcsnlen_s(current, remaining);
    if (length == remaining) {
      break;
    }
    items.emplace_back(current, length);
    current += length + 1;
    remaining -= length + 1;
  }
  return items;
}

std::vector<BYTE> MultiStringData(const std::vector<std::wstring>& items) {
  size_t characters = 1;
  for (const auto& item : items) {
    characters += item.size() + 1;
  }
  std::vector<BYTE> data(characters * sizeof(wchar_t), 0);
  wchar_t* output = reinterpret_cast<wchar_t*>(data.data());
  for (const auto& item : items) {
    std::memcpy(output, item.data(), item.size() * sizeof(wchar_t));
    output += item.size() + 1;
  }
  return data;
}

std::wstring MultiStringText(const std::vector<BYTE>& data) {
  std::wstring text;
  for (const auto& item : MultiStringItems(data)) {
    if (!text.empty()) {
      text += L"\r\n";
    }
    text += item;
  }
  return text;
}

std::vector<BYTE> MultiStringData(std::wstring_view lines) {
  std::vector<std::wstring> items;
  size_t start = 0;
  while (start <= lines.size()) {
    size_t end = lines.find_first_of(L"\r\n", start);
    if (end == std::wstring_view::npos) {
      end = lines.size();
    }
    items.emplace_back(lines.substr(start, end - start));
    if (end == lines.size()) {
      break;
    }
    start = end + 1;
    if (lines[end] == L'\r' && start < lines.size() &&
        lines[start] == L'\n') {
      ++start;
    }
  }
  return MultiStringData(items);
}

} // namespace regkit::value_format
