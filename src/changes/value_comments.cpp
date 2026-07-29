// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "changes/value_comments.h"

#include "records/escaped_fields.h"
#include "win32/file_text.h"
#include "win32/win32_helpers.h"

#include <cwctype>

namespace regkit::changes {
namespace {

bool HasText(const std::wstring& text) {
  for (wchar_t character : text) {
    if (!iswspace(character)) {
      return true;
    }
  }
  return false;
}

} // namespace

std::wstring ValueComments::ValueKey(const std::wstring& path,
                                     const std::wstring& name, DWORD type) {
  std::wstring key = util::ToLower(path);
  key.push_back(L'\t');
  key.append(util::ToLower(name));
  key.push_back(L'\t');
  key.append(std::to_wstring(type));
  return key;
}

std::wstring ValueComments::NameKey(const std::wstring& name, DWORD type) {
  std::wstring key = util::ToLower(name);
  key.push_back(L'\t');
  key.append(std::to_wstring(type));
  return key;
}

bool ValueComments::Load(const std::wstring& path) {
  Clear();
  std::wstring content;
  if (!util::ReadTextFile(path, &content)) {
    return false;
  }
  Merge(ParseComments(content));
  return true;
}

bool ValueComments::Save(const std::wstring& path) const {
  return !path.empty() &&
         util::WriteTextFile(path, SerializeComments(*this), false);
}

bool ValueComments::Import(const std::wstring& path) {
  return Load(path);
}

bool ValueComments::Export(const std::wstring& path) const {
  return Save(path);
}

void ValueComments::Clear() {
  value_entries_.clear();
  name_entries_.clear();
}

void ValueComments::Merge(const CommentDocument& document) {
  for (const CommentEntry& entry : document.value_entries) {
    value_entries_[ValueKey(entry.path, entry.name, entry.type)] = entry;
  }
  for (const CommentEntry& entry : document.name_entries) {
    name_entries_[NameKey(entry.name, entry.type)] = entry;
  }
}

const std::unordered_map<std::wstring, CommentEntry>&
ValueComments::value_entries() const noexcept {
  return value_entries_;
}

std::unordered_map<std::wstring, CommentEntry>&
ValueComments::value_entries() noexcept {
  return value_entries_;
}

const std::unordered_map<std::wstring, CommentEntry>&
ValueComments::name_entries() const noexcept {
  return name_entries_;
}

std::unordered_map<std::wstring, CommentEntry>&
ValueComments::name_entries() noexcept {
  return name_entries_;
}

CommentDocument ParseComments(const std::wstring& content) {
  CommentDocument document;
  for (const std::wstring& line : record_fields::Lines(content)) {
    if (line.empty()) {
      continue;
    }
    const auto fields = record_fields::Split(line);
    if (fields.size() < 5) {
      continue;
    }
    CommentEntry entry;
    entry.path = record_fields::Unescape(fields[1]);
    entry.name = record_fields::Unescape(fields[2]);
    try {
      entry.type = static_cast<DWORD>(std::stoul(fields[3]));
    } catch (...) {
      continue;
    }
    entry.text = record_fields::Unescape(fields[4]);
    if (!HasText(entry.text)) {
      continue;
    }
    if (_wcsicmp(fields[0].c_str(), L"value") == 0) {
      document.value_entries.push_back(std::move(entry));
    } else if (_wcsicmp(fields[0].c_str(), L"name") == 0) {
      document.name_entries.push_back(std::move(entry));
    }
  }
  return document;
}

std::wstring SerializeComments(const ValueComments& comments) {
  std::wstring content;
  for (const auto& pair : comments.value_entries()) {
    const CommentEntry& entry = pair.second;
    if (!HasText(entry.text)) {
      continue;
    }
    content.append(L"value\t");
    content.append(record_fields::Escape(entry.path));
    content.push_back(L'\t');
    content.append(record_fields::Escape(entry.name));
    content.push_back(L'\t');
    content.append(std::to_wstring(entry.type));
    content.push_back(L'\t');
    content.append(record_fields::Escape(entry.text));
    content.push_back(L'\n');
  }
  for (const auto& pair : comments.name_entries()) {
    const CommentEntry& entry = pair.second;
    if (!HasText(entry.text)) {
      continue;
    }
    content.append(L"name\t\t");
    content.append(record_fields::Escape(entry.name));
    content.push_back(L'\t');
    content.append(std::to_wstring(entry.type));
    content.push_back(L'\t');
    content.append(record_fields::Escape(entry.text));
    content.push_back(L'\n');
  }
  return content;
}

} // namespace regkit::changes
