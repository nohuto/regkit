// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "regfile/reg_file.h"

#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/file_text.h"

#include <algorithm>
#include <cstring>
#include <cwctype>

namespace regkit::regfile {
namespace {

std::wstring Trim(std::wstring_view text) {
  size_t first = 0;
  while (first < text.size() && iswspace(text[first])) {
    ++first;
  }
  size_t last = text.size();
  while (last > first && iswspace(text[last - 1])) {
    --last;
  }
  return std::wstring(text.substr(first, last - first));
}

std::wstring Lower(std::wstring_view text) {
  std::wstring result(text);
  for (wchar_t& character : result) {
    character = towlower(character);
  }
  return result;
}

bool ParseQuoted(std::wstring_view text, std::wstring* output) {
  if (!output || text.empty() || text.front() != L'"') {
    return false;
  }
  output->clear();
  bool escaped = false;
  for (size_t index = 1; index < text.size(); ++index) {
    const wchar_t character = text[index];
    if (escaped) {
      switch (character) {
      case L'n':
        output->push_back(L'\n');
        break;
      case L'r':
        output->push_back(L'\r');
        break;
      case L't':
        output->push_back(L'\t');
        break;
      case L'0':
        output->push_back(L'\0');
        break;
      default:
        output->push_back(character);
        break;
      }
      escaped = false;
    } else if (character == L'\\') {
      escaped = true;
    } else if (character == L'"') {
      return true;
    } else {
      output->push_back(character);
    }
  }
  return false;
}

DWORD TypeFromCode(unsigned long code) {
  switch (code) {
  case 0x0:
    return REG_NONE;
  case 0x1:
    return REG_SZ;
  case 0x2:
    return REG_EXPAND_SZ;
  case 0x3:
    return REG_BINARY;
  case 0x4:
    return REG_DWORD;
  case 0x5:
    return REG_DWORD_BIG_ENDIAN;
  case 0x6:
    return REG_LINK;
  case 0x7:
    return REG_MULTI_SZ;
  case 0x8:
    return REG_RESOURCE_LIST;
  case 0x9:
    return REG_FULL_RESOURCE_DESCRIPTOR;
  case 0xA:
    return REG_RESOURCE_REQUIREMENTS_LIST;
  case 0xB:
    return REG_QWORD;
  default:
    return REG_BINARY;
  }
}

DWORD TypeCode(DWORD type) {
  return value_format::NormalizeType(type);
}

std::wstring Escape(std::wstring_view text) {
  std::wstring output;
  output.reserve(text.size());
  for (wchar_t character : text) {
    switch (character) {
    case L'\\':
      output += L"\\\\";
      break;
    case L'"':
      output += L"\\\"";
      break;
    case L'\n':
      output += L"\\n";
      break;
    case L'\r':
      output += L"\\r";
      break;
    case L'\t':
      output += L"\\t";
      break;
    case L'\0':
      output += L"\\0";
      break;
    default:
      output.push_back(character);
      break;
    }
  }
  return output;
}

std::wstring Hex(std::span<const BYTE> data) {
  std::wstring output;
  output.reserve(data.size() * 3);
  for (size_t index = 0; index < data.size(); ++index) {
    if (index != 0) {
      output.push_back(L',');
    }
    wchar_t byte[3] = {};
    swprintf_s(byte, L"%02x", data[index]);
    output += byte;
  }
  return output;
}

std::wstring SerializeValue(const Value& value) {
  const DWORD type = value_format::NormalizeType(value.type);
  if (type == REG_SZ) {
    std::wstring text;
    if (value_format::DecodeString(value.data, &text)) {
      return L"\"" + Escape(text) + L"\"";
    }
  }
  if (type == REG_DWORD && value.data.size() >= sizeof(DWORD)) {
    DWORD number = 0;
    std::memcpy(&number, value.data.data(), sizeof(number));
    wchar_t output[16] = {};
    swprintf_s(output, L"dword:%08x", number);
    return output;
  }
  if (type == REG_BINARY && value.type == REG_BINARY) {
    return L"hex:" + Hex(value.data);
  }
  wchar_t code[16] = {};
  swprintf_s(code, L"%x", TypeCode(value.type));
  return L"hex(" + std::wstring(code) + L"):" + Hex(value.data);
}

} // namespace

Writer::Writer()
    : output_(L"Windows Registry Editor Version 5.00\r\n\r\n") {}

void Writer::AppendKey(std::wstring_view path,
                       std::vector<const Value*> values) {
  if (output_.size() >
      std::wstring_view(L"Windows Registry Editor Version 5.00\r\n\r\n")
          .size()) {
    output_ += L"\r\n";
  }
  output_.push_back(L'[');
  output_.append(path);
  output_ += L"]\r\n";
  std::sort(values.begin(), values.end(),
            [](const Value* left, const Value* right) {
              if (left->name.empty() != right->name.empty()) {
                return left->name.empty();
              }
              return _wcsicmp(left->name.c_str(), right->name.c_str()) < 0;
            });
  for (const Value* value : values) {
    if (value->name.empty()) {
      output_ += L"@=";
    } else {
      output_.push_back(L'"');
      output_ += Escape(value->name);
      output_ += L"\"=";
    }
    output_ += SerializeValue(*value);
    output_ += L"\r\n";
  }
}

std::wstring Writer::Finish() && {
  return std::move(output_);
}

bool Parse(std::wstring_view content, Document* output,
           const std::atomic_bool* cancel, bool* cancelled) {
  if (!output) {
    return false;
  }
  output->keys.clear();
  output->key_order.clear();
  if (cancelled) {
    *cancelled = false;
  }
  auto stopped = [&] {
    const bool value = cancel && cancel->load();
    if (value && cancelled) {
      *cancelled = true;
    }
    return value;
  };

  std::vector<std::wstring> lines;
  std::wstring continued;
  size_t start = 0;
  while (start < content.size()) {
    if (stopped()) {
      return false;
    }
    size_t end = content.find(L'\n', start);
    if (end == std::wstring_view::npos) {
      end = content.size();
    }
    std::wstring line(content.substr(start, end - start));
    if (!line.empty() && line.back() == L'\r') {
      line.pop_back();
    }
    start = end + 1;
    continued += line;
    while (!continued.empty() &&
           (continued.back() == L' ' || continued.back() == L'\t')) {
      continued.pop_back();
    }
    if (!continued.empty() && continued.back() == L'\\') {
      continued.pop_back();
      continue;
    }
    lines.push_back(std::move(continued));
    continued.clear();
  }
  if (!continued.empty()) {
    lines.push_back(std::move(continued));
  }

  Key* current_key = nullptr;
  for (const auto& raw : lines) {
    if (stopped()) {
      return false;
    }
    const std::wstring line = Trim(raw);
    if (line.empty() || line.front() == L';' ||
        registry_path::StartsWith(line, L"Windows Registry Editor") ||
        registry_path::StartsWith(line, L"REGEDIT4")) {
      continue;
    }
    if (line.front() == L'[' && line.back() == L']') {
      std::wstring path = Trim(
          std::wstring_view(line).substr(1, line.size() - 2));
      if (!path.empty() && path.front() == L'-') {
        current_key = nullptr;
        continue;
      }
      const std::wstring lower = Lower(path);
      auto [iterator, inserted] =
          output->keys.try_emplace(lower, Key{path, {}});
      if (inserted) {
        output->key_order.push_back(path);
      }
      current_key = &iterator->second;
      continue;
    }
    if (!current_key) {
      continue;
    }
    const size_t equals = line.find(L'=');
    if (equals == std::wstring::npos) {
      continue;
    }
    const std::wstring name_text = Trim(
        std::wstring_view(line).substr(0, equals));
    const std::wstring data_text = Trim(
        std::wstring_view(line).substr(equals + 1));
    if (name_text.empty() || data_text.empty() || data_text == L"-") {
      continue;
    }

    Value value;
    if (name_text == L"@") {
      value.name.clear();
    } else if (!ParseQuoted(name_text, &value.name)) {
      continue;
    }

    if (data_text.front() == L'"') {
      std::wstring text;
      if (!ParseQuoted(data_text, &text)) {
        continue;
      }
      value.type = REG_SZ;
      value.data = value_format::StringData(text);
    } else if (registry_path::StartsWith(data_text, L"dword:")) {
      const std::wstring number_text = Trim(
          std::wstring_view(data_text).substr(6));
      if (number_text.empty()) {
        continue;
      }
      const DWORD number =
          static_cast<DWORD>(wcstoul(number_text.c_str(), nullptr, 16));
      value.type = REG_DWORD;
      value.data.resize(sizeof(number));
      std::memcpy(value.data.data(), &number, sizeof(number));
    } else if (registry_path::StartsWith(data_text, L"hex")) {
      const size_t colon = data_text.find(L':');
      if (colon == std::wstring::npos) {
        continue;
      }
      value.type = REG_BINARY;
      const size_t open = data_text.find(L'(');
      const size_t close = data_text.find(L')');
      if (open != std::wstring::npos && close > open) {
        const std::wstring code =
            data_text.substr(open + 1, close - open - 1);
        value.type = TypeFromCode(wcstoul(code.c_str(), nullptr, 16));
      }
      if (!value_format::ParseHex(
              std::wstring_view(data_text).substr(colon + 1), &value.data)) {
        continue;
      }
    } else {
      continue;
    }
    current_key->values[Lower(value.name)] = std::move(value);
  }
  return true;
}

bool Load(const std::wstring& path, Document* output, std::wstring* error,
          const std::atomic_bool* cancel, bool* cancelled) {
  std::wstring content;
  if (!util::ReadTextFile(path, &content, nullptr, 32ull * 1024ull * 1024ull)) {
    if (error) {
      *error = L"Failed to read registry file.";
    }
    return false;
  }
  return Parse(content, output, cancel, cancelled);
}

std::wstring Serialize(const Document& document) {
  Writer writer;
  for (const auto& ordered_path : document.key_order) {
    auto key = document.keys.find(Lower(ordered_path));
    if (key == document.keys.end()) {
      continue;
    }
    std::vector<const Value*> values;
    values.reserve(key->second.values.size());
    for (const auto& entry : key->second.values) {
      values.push_back(&entry.second);
    }
    writer.AppendKey(key->second.path, std::move(values));
  }
  return std::move(writer).Finish();
}

} // namespace regkit::regfile
