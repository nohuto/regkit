// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_path.h"

#include "registry/registry_store.h"
#include "win32/process_rights.h"

#include <algorithm>
#include <cwctype>

namespace regkit::registry_path {
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

std::wstring CanonicalRoot(std::wstring_view root) {
  if (Equals(root, L"HKCR") || Equals(root, L"HKEY_CLASSES_ROOT")) {
    return L"HKEY_CLASSES_ROOT";
  }
  if (Equals(root, L"HKCU") || Equals(root, L"HKEY_CURRENT_USER")) {
    return L"HKEY_CURRENT_USER";
  }
  if (Equals(root, L"HKLM") || Equals(root, L"HKEY_LOCAL_MACHINE") ||
      Equals(root, L"MACHINE")) {
    return L"HKEY_LOCAL_MACHINE";
  }
  if (Equals(root, L"HKU") || Equals(root, L"HKEY_USERS") ||
      Equals(root, L"USER") || Equals(root, L"USERS")) {
    return L"HKEY_USERS";
  }
  if (Equals(root, L"HKCC") || Equals(root, L"HKEY_CURRENT_CONFIG")) {
    return L"HKEY_CURRENT_CONFIG";
  }
  if (Equals(root, L"REGISTRY")) {
    return L"REGISTRY";
  }
  return {};
}

std::wstring AbbreviatedRoot(std::wstring_view root) {
  if (Equals(root, L"HKEY_CLASSES_ROOT")) {
    return L"HKCR";
  }
  if (Equals(root, L"HKEY_CURRENT_USER")) {
    return L"HKCU";
  }
  if (Equals(root, L"HKEY_LOCAL_MACHINE")) {
    return L"HKLM";
  }
  if (Equals(root, L"HKEY_USERS")) {
    return L"HKU";
  }
  if (Equals(root, L"HKEY_CURRENT_CONFIG")) {
    return L"HKCC";
  }
  return std::wstring(root);
}

std::wstring Join(std::wstring_view root, std::wstring_view rest) {
  if (rest.empty()) {
    return std::wstring(root);
  }
  std::wstring result(root);
  result.push_back(L'\\');
  result.append(rest);
  return result;
}

bool HasComponentPrefix(std::wstring_view path, std::wstring_view prefix) {
  return StartsWith(path, prefix) &&
         (path.size() == prefix.size() || path[prefix.size()] == L'\\');
}

std::wstring JoinRange(const std::vector<std::wstring>& parts, size_t first,
                       size_t last) {
  size_t characters = 0;
  for (size_t index = first; index < last; ++index) {
    characters += parts[index].size() + 1;
  }
  std::wstring result;
  result.reserve(characters);
  for (size_t index = first; index < last; ++index) {
    if (parts[index].empty()) {
      continue;
    }
    if (!result.empty()) {
      result.push_back(L'\\');
    }
    result += parts[index];
  }
  return result;
}

} // namespace

bool Equals(std::wstring_view left, std::wstring_view right) {
  return left.size() == right.size() &&
         _wcsnicmp(left.data(), right.data(), left.size()) == 0;
}

bool StartsWith(std::wstring_view text, std::wstring_view prefix) {
  return text.size() >= prefix.size() &&
         _wcsnicmp(text.data(), prefix.data(), prefix.size()) == 0;
}

std::wstring RootName(HKEY root) {
  std::wstring virtual_name;
  if (RegistryStore::GetVirtualRootName(root, &virtual_name)) {
    return virtual_name;
  }
  if (root == HKEY_CLASSES_ROOT) {
    return L"HKEY_CLASSES_ROOT";
  }
  if (root == HKEY_CURRENT_USER) {
    return L"HKEY_CURRENT_USER";
  }
  if (root == HKEY_LOCAL_MACHINE) {
    return L"HKEY_LOCAL_MACHINE";
  }
  if (root == HKEY_USERS) {
    return L"HKEY_USERS";
  }
  if (root == HKEY_CURRENT_CONFIG) {
    return L"HKEY_CURRENT_CONFIG";
  }
  if (root == HKEY_PERFORMANCE_DATA) {
    return L"HKEY_PERFORMANCE_DATA";
  }
  if (root == HKEY_PERFORMANCE_TEXT) {
    return L"HKEY_PERFORMANCE_TEXT";
  }
  if (root == HKEY_PERFORMANCE_NLSTEXT) {
    return L"HKEY_PERFORMANCE_NLSTEXT";
  }
  return {};
}

std::wstring Build(const RegistryNode& node) {
  const std::wstring root =
      node.root_name.empty() ? RootName(node.root) : node.root_name;
  return Join(root, node.subkey);
}

std::wstring BuildNative(const RegistryNode& node) {
  if (RegistryStore::IsVirtualRoot(node.root)) {
    return {};
  }
  if (Equals(node.root_name, L"REGISTRY")) {
    return Join(L"\\REGISTRY", node.subkey);
  }

  std::wstring root;
  if (node.root == HKEY_LOCAL_MACHINE) {
    root = L"\\REGISTRY\\MACHINE";
  } else if (node.root == HKEY_USERS) {
    root = L"\\REGISTRY\\USER";
  } else if (node.root == HKEY_CURRENT_USER) {
    const std::wstring sid = util::GetCurrentUserSidString();
    if (sid.empty()) {
      return {};
    }
    root = L"\\REGISTRY\\USER\\" + sid;
  } else if (node.root == HKEY_CURRENT_CONFIG) {
    root =
        L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Hardware\\Profiles\\Current";
  } else if (node.root == HKEY_CLASSES_ROOT) {
    root = L"\\REGISTRY\\MACHINE\\SOFTWARE\\Classes";
  } else {
    return {};
  }
  return Join(root, node.subkey);
}

std::vector<std::wstring> Split(std::wstring_view path) {
  std::vector<std::wstring> parts;
  size_t start = 0;
  while (start < path.size()) {
    while (start < path.size() &&
           (path[start] == L'\\' || path[start] == L'/')) {
      ++start;
    }
    if (start == path.size()) {
      break;
    }
    size_t end = path.find_first_of(L"\\/", start);
    if (end == std::wstring_view::npos) {
      end = path.size();
    }
    parts.emplace_back(path.substr(start, end - start));
    start = end + 1;
  }
  return parts;
}

std::wstring Join(const std::vector<std::wstring>& parts, size_t first_part) {
  return JoinRange(parts, std::min(first_part, parts.size()), parts.size());
}

std::wstring JoinPrefix(const std::vector<std::wstring>& parts,
                        size_t part_count) {
  return JoinRange(parts, 0, std::min(part_count, parts.size()));
}

std::wstring Parent(std::wstring_view path) {
  while (!path.empty() && (path.back() == L'\\' || path.back() == L'/')) {
    path.remove_suffix(1);
  }
  const size_t split = path.find_last_of(L"\\/");
  return split == std::wstring_view::npos ? std::wstring{}
                                          : std::wstring(path.substr(0, split));
}

std::wstring Leaf(std::wstring_view path) {
  while (!path.empty() && (path.back() == L'\\' || path.back() == L'/')) {
    path.remove_suffix(1);
  }
  const size_t split = path.find_last_of(L"\\/");
  return std::wstring(path.substr(split == std::wstring_view::npos ? 0
                                                                   : split + 1));
}

std::wstring Clean(std::wstring_view input) {
  std::wstring path = Trim(input);
  if (path.size() >= 2 &&
      ((path.front() == L'[' && path.back() == L']') ||
       (path.front() == L'"' && path.back() == L'"') ||
       (path.front() == L'\'' && path.back() == L'\''))) {
    path = Trim(std::wstring_view(path).substr(1, path.size() - 2));
  }
  if (!path.empty() && path.front() == L'-') {
    path = Trim(std::wstring_view(path).substr(1));
  }
  for (wchar_t& character : path) {
    if (character == L'/') {
      character = L'\\';
    }
  }

  std::wstring collapsed;
  collapsed.reserve(path.size());
  for (wchar_t character : path) {
    if (character != L'\\' || collapsed.empty() ||
        collapsed.back() != L'\\') {
      collapsed.push_back(character);
    }
  }
  path = std::move(collapsed);

  if (StartsWith(path, L"Registry::")) {
    path.erase(0, 10);
  }
  while (!path.empty() && path.front() == L'\\') {
    path.erase(path.begin());
  }
  if (StartsWith(path, L"My Computer\\")) {
    path.erase(0, 12);
  } else if (StartsWith(path, L"Computer\\")) {
    path.erase(0, 9);
  }
  return path;
}

std::wstring Normalize(std::wstring_view input,
                       std::wstring_view current_user_sid) {
  std::wstring path = Clean(input);

  if (StartsWith(path, L"REGISTRY\\")) {
    std::wstring native = path.substr(9);
    auto native_rest = [&](std::wstring_view prefix) {
      std::wstring_view rest(native);
      rest.remove_prefix(prefix.size());
      while (!rest.empty() && rest.front() == L'\\') {
        rest.remove_prefix(1);
      }
      return rest;
    };
    constexpr std::wstring_view classes =
        L"MACHINE\\SOFTWARE\\Classes";
    if (HasComponentPrefix(native, classes)) {
      return Join(L"HKEY_CLASSES_ROOT", native_rest(classes));
    }
    constexpr std::wstring_view current_config =
        L"MACHINE\\SYSTEM\\CurrentControlSet\\Hardware Profiles\\Current";
    if (HasComponentPrefix(native, current_config)) {
      return Join(L"HKEY_CURRENT_CONFIG", native_rest(current_config));
    }
    if (HasComponentPrefix(native, L"MACHINE")) {
      std::wstring_view rest(native);
      rest.remove_prefix(7);
      while (!rest.empty() && rest.front() == L'\\') {
        rest.remove_prefix(1);
      }
      return Join(L"HKEY_LOCAL_MACHINE", rest);
    }
    if (HasComponentPrefix(native, L"USER")) {
      std::wstring_view rest(native);
      rest.remove_prefix(4);
      while (!rest.empty() && rest.front() == L'\\') {
        rest.remove_prefix(1);
      }
      if (!current_user_sid.empty() &&
          HasComponentPrefix(rest, current_user_sid)) {
        rest.remove_prefix(current_user_sid.size());
        while (!rest.empty() && rest.front() == L'\\') {
          rest.remove_prefix(1);
        }
        return Join(L"HKEY_CURRENT_USER", rest);
      }
      return Join(L"HKEY_USERS", rest);
    }
    const size_t native_split = native.find(L'\\');
    const std::wstring_view native_root =
        native_split == std::wstring::npos
            ? std::wstring_view(native)
            : std::wstring_view(native).substr(0, native_split);
    const std::wstring canonical = CanonicalRoot(native_root);
    if (!canonical.empty() && !Equals(canonical, L"REGISTRY")) {
      const std::wstring_view rest =
          native_split == std::wstring::npos
              ? std::wstring_view{}
              : std::wstring_view(native).substr(native_split + 1);
      return Join(canonical, rest);
    }
  }

  const size_t split = path.find_first_of(L":\\");
  const std::wstring_view root(path.data(),
                               split == std::wstring::npos ? path.size()
                                                           : split);
  std::wstring_view rest =
      split == std::wstring::npos
          ? std::wstring_view{}
          : std::wstring_view(path).substr(split + (path[split] == L':' ? 1 : 0));
  while (!rest.empty() && rest.front() == L'\\') {
    rest.remove_prefix(1);
  }
  const std::wstring canonical = CanonicalRoot(root);
  return canonical.empty() ? path : Join(canonical, rest);
}

std::wstring Format(std::wstring_view path, Style style,
                    std::wstring_view tree_root) {
  const size_t split = path.find(L'\\');
  const std::wstring_view root =
      split == std::wstring_view::npos ? path : path.substr(0, split);
  const std::wstring_view rest =
      split == std::wstring_view::npos ? std::wstring_view{}
                                       : path.substr(split + 1);
  switch (style) {
  case Style::kAbbreviated:
    return Join(AbbreviatedRoot(root), rest);
  case Style::kRegeditAddress:
    return Join(tree_root.empty() ? L"Computer" : tree_root, path);
  case Style::kRegFileHeader:
    return L"[" + std::wstring(path) + L"]";
  case Style::kPowerShellDrive: {
    std::wstring result = AbbreviatedRoot(root) + L":";
    if (!rest.empty()) {
      result += L"\\" + std::wstring(rest);
    }
    return result;
  }
  case Style::kPowerShellProvider:
    return L"Registry::" + std::wstring(path);
  case Style::kEscaped: {
    std::wstring result;
    result.reserve(path.size() * 2);
    for (wchar_t character : path) {
      if (character == L'\\') {
        result.push_back(L'\\');
      }
      result.push_back(character);
    }
    return result;
  }
  case Style::kFull:
    return std::wstring(path);
  }
  return std::wstring(path);
}

bool ParseRoot(std::wstring_view input, RegistryNode* node) {
  if (!node) {
    return false;
  }
  const std::wstring normalized = Normalize(input);
  const size_t split = normalized.find(L'\\');
  const std::wstring root = normalized.substr(0, split);
  const std::wstring rest =
      split == std::wstring::npos ? L"" : normalized.substr(split + 1);
  node->subkey = rest;
  node->root_name = root;
  if (Equals(root, L"HKEY_CLASSES_ROOT")) {
    node->root = HKEY_CLASSES_ROOT;
  } else if (Equals(root, L"HKEY_CURRENT_USER")) {
    node->root = HKEY_CURRENT_USER;
  } else if (Equals(root, L"HKEY_LOCAL_MACHINE")) {
    node->root = HKEY_LOCAL_MACHINE;
  } else if (Equals(root, L"HKEY_USERS")) {
    node->root = HKEY_USERS;
  } else if (Equals(root, L"HKEY_CURRENT_CONFIG")) {
    node->root = HKEY_CURRENT_CONFIG;
  } else if (Equals(root, L"REGISTRY")) {
    node->root = nullptr;
  } else {
    return false;
  }
  return true;
}

} // namespace regkit::registry_path
