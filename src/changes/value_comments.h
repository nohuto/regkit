// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#pragma once

#include <win32ui/windows_config.h>

#include <windows.h>

#include <cstdint>
#include <optional>
#include <string>
#include <unordered_map>
#include <vector>

namespace regkit::changes
{

enum class CommentKeyScope
{
    kAny,
    kExact,
    kRecursive,
};

enum class CommentSource
{
    kNone,
    kUser,
    kDefault,
};

struct CommentRule
{
    std::wstring name;
    std::optional<DWORD> type;
    std::optional<uint64_t> data_size;
    CommentKeyScope key_scope = CommentKeyScope::kAny;
    std::wstring key_path;
    std::wstring text;
    bool key = false;
};

struct CommentTarget
{
    std::wstring path;
    std::wstring name;
    DWORD type = 0;
    uint64_t data_size = 0;
    bool key = false;
};

struct ResolvedComment
{
    std::wstring text;
    CommentSource source = CommentSource::kNone;
    CommentRule rule;
};

class ValueComments
{
  public:
    bool Load(const std::wstring& path);
    bool Save(const std::wstring& path) const;
    void Clear();
    void Merge(const std::vector<CommentRule>& rules);
    const CommentRule* Match(const CommentTarget& target) const;
    void Set(CommentRule rule);
    void Erase(const CommentRule& rule);
    const std::vector<CommentRule>& rules() const noexcept;
    bool HasKeyRules() const noexcept
    {
        return !key_index_.empty();
    }

  private:
    void Reindex();

    std::vector<CommentRule> rules_;
    std::unordered_map<std::wstring, std::vector<size_t>> index_;
    std::unordered_map<std::wstring, size_t> key_index_;
};

bool ParseComments(const std::wstring& content, std::vector<CommentRule>* out, std::wstring* error = nullptr);
bool ValidateCatalog(const std::vector<CommentRule>& rules);
std::wstring SerializeComments(const ValueComments& comments);
CommentRule ValueRule(const CommentTarget& target);
ResolvedComment ResolveComment(const ValueComments& user, const ValueComments& defaults, const CommentTarget& target);

}
