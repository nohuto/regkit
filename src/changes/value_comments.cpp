// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "changes/value_comments.h"

#include "registry/value_format.h"
#include "win32/file_text.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <win32ui/json/reader.h>
#include <win32ui/json/writer.h>

#include <algorithm>
#include <tuple>
#include <unordered_set>

namespace regkit::changes
{

namespace
{

constexpr wchar_t kFormat[] = L"regkit-comments";
constexpr size_t kMaxTextLength = 65536;
constexpr size_t kMaxPathLength = 32767;

std::wstring NormalizeKeyPath(std::wstring path)
{
    while (!path.empty() && path.back() == L'\\')
    {
        path.pop_back();
    }
    return path;
}

std::wstring ConditionKey(const CommentRule& rule)
{
    return (rule.key ? L"k\t" : L"v\t") + std::to_wstring(rule.name.size()) + L':' + util::ToLower(rule.name) + L'\t' +
           (rule.type ? std::to_wstring(*rule.type) : L"*") + L'\t' +
           (rule.data_size ? std::to_wstring(*rule.data_size) : L"*") + L'\t' +
           std::to_wstring(static_cast<int>(rule.key_scope)) + L'\t' + util::ToLower(NormalizeKeyPath(rule.key_path));
}

bool KeyMatches(const CommentRule& rule, const std::wstring& path)
{
    switch (rule.key_scope)
    {
    case CommentKeyScope::kExact:
        return util::EqualsInsensitive(path, rule.key_path);
    case CommentKeyScope::kRecursive:
        return util::StartsWithInsensitive(path, rule.key_path) &&
               (path.size() == rule.key_path.size() || path[rule.key_path.size()] == L'\\');
    default:
        return true;
    }
}

bool Matches(const CommentRule& rule, const CommentTarget& target)
{
    return (!rule.type || *rule.type == target.type) && (!rule.data_size || *rule.data_size == target.data_size) &&
           KeyMatches(rule, target.path);
}

auto Specificity(const CommentRule& rule)
{
    const int scope = rule.key_scope == CommentKeyScope::kExact       ? 3
                      : rule.key_scope == CommentKeyScope::kRecursive ? 2
                                                                      : 1;
    const size_t depth = rule.key_scope == CommentKeyScope::kRecursive ? rule.key_path.size() : 0;
    return std::tuple(scope, depth, static_cast<int>(rule.type.has_value()) + static_cast<int>(rule.data_size.has_value()));
}

bool ReadType(win32ui::JsonReader& reader, std::optional<DWORD>* type)
{
    if (reader.Next() == L'"')
    {
        std::wstring name;
        if (!reader.String(&name, 64))
        {
            return false;
        }
        for (DWORD candidate = REG_NONE; candidate <= REG_QWORD; ++candidate)
        {
            if (util::EqualsInsensitive(name, value_format::TypeName(candidate)))
            {
                *type = candidate;
                return true;
            }
        }
        return reader.Fail(util::Tr(L"A comment has an unknown registry type."));
    }
    uint64_t value = 0;
    if (!reader.Unsigned(&value))
    {
        return false;
    }
    if (value > MAXDWORD)
    {
        return reader.Fail(util::Tr(L"A comment type is out of range."));
    }
    *type = static_cast<DWORD>(value);
    return true;
}

bool ReadRule(win32ui::JsonReader& reader, CommentRule* rule)
{
    unsigned seen = 0;
    const auto claim = [&](unsigned bit) {
        const bool first = (seen & bit) == 0;
        seen |= bit;
        return first || reader.Fail(util::Tr(L"A comment contains a duplicate member."));
    };
    const bool read = reader.Object([&](const std::wstring& member) {
        if (member == L"name")
        {
            return claim(1) && reader.String(&rule->name, kMaxPathLength);
        }
        if (member == L"text")
        {
            return claim(2) && reader.String(&rule->text, kMaxTextLength);
        }
        if (member == L"type")
        {
            return claim(4) && ReadType(reader, &rule->type);
        }
        if (member == L"size")
        {
            uint64_t size = 0;
            return claim(8) && reader.Unsigned(&size) && (rule->data_size = size, true);
        }
        if (member == L"key" || member == L"tree")
        {
            rule->key_scope = member == L"key" ? CommentKeyScope::kExact : CommentKeyScope::kRecursive;
            return claim(16) && reader.String(&rule->key_path, kMaxPathLength) &&
                   (!(rule->key_path = NormalizeKeyPath(std::move(rule->key_path))).empty() ||
                    reader.Fail(util::Tr(L"A comment has an empty key.")));
        }
        return reader.Fail(util::Tr(L"A comment contains an unknown member."));
    });
    if (!read || ((seen & 2) == 0 && !reader.Fail(util::Tr(L"A comment is missing its text."))))
    {
        return false;
    }
    rule->key = (seen & 1) == 0;
    return !rule->key || ((seen & 12) == 0 && rule->key_scope == CommentKeyScope::kExact) ||
           reader.Fail(util::Tr(L"A key comment needs a key and can't have a name, type, size or tree."));
}

void AppendMember(std::wstring* out, const wchar_t* name, std::wstring_view value, bool quote)
{
    out->append(out->back() == L'{' ? L"\n      \"" : L",\n      \"").append(name).append(L"\": ");
    if (quote)
    {
        win32ui::AppendJsonString(out, value);
    }
    else
    {
        out->append(value);
    }
}

}

bool ValueComments::Load(const std::wstring& path)
{
    std::wstring content;
    std::vector<CommentRule> rules;
    if (!util::ReadStateFile(path, &content, nullptr, util::kMaxCommentFileBytes) || (!content.empty() && !ParseComments(content, &rules)))
    {
        return false;
    }
    Clear();
    Merge(rules);
    return true;
}

bool ValueComments::Save(const std::wstring& path) const
{
    return !path.empty() && win32ui::WriteTextFileAtomic(path.c_str(), SerializeComments(*this)) == ERROR_SUCCESS;
}

void ValueComments::Clear()
{
    rules_.clear();
    Reindex();
}

void ValueComments::Merge(const std::vector<CommentRule>& rules)
{
    rules_.insert(rules_.end(), rules.begin(), rules.end());
    std::unordered_set<std::wstring> seen;
    std::vector<CommentRule> unique;
    for (auto rule = rules_.rbegin(); rule != rules_.rend(); ++rule)
    {
        rule->key_path = NormalizeKeyPath(std::move(rule->key_path));
        if (seen.insert(ConditionKey(*rule)).second)
        {
            unique.push_back(std::move(*rule));
        }
    }
    rules_.assign(std::make_move_iterator(unique.rbegin()), std::make_move_iterator(unique.rend()));
    Reindex();
}

const CommentRule* ValueComments::Match(const CommentTarget& target) const
{
    if (target.key)
    {
        const auto match = key_index_.find(util::ToLower(NormalizeKeyPath(target.path)));
        return match == key_index_.end() ? nullptr : &rules_[match->second];
    }
    const auto candidates = index_.find(util::ToLower(target.name));
    if (candidates == index_.end())
    {
        return nullptr;
    }
    const CommentRule* best = nullptr;
    for (const size_t index : candidates->second)
    {
        const CommentRule& rule = rules_[index];
        if (Matches(rule, target) && (!best || Specificity(rule) >= Specificity(*best)))
        {
            best = &rule;
        }
    }
    return best;
}

void ValueComments::Set(CommentRule rule)
{
    Merge({std::move(rule)});
}

void ValueComments::Erase(const CommentRule& rule)
{
    const std::wstring key = ConditionKey(rule);
    std::erase_if(rules_, [&](const CommentRule& existing) { return ConditionKey(existing) == key; });
    Reindex();
}

const std::vector<CommentRule>& ValueComments::rules() const noexcept
{
    return rules_;
}

void ValueComments::Reindex()
{
    index_.clear();
    key_index_.clear();
    for (size_t index = 0; index < rules_.size(); ++index)
    {
        if (rules_[index].key)
        {
            key_index_[util::ToLower(rules_[index].key_path)] = index;
        }
        else
        {
            index_[util::ToLower(rules_[index].name)].push_back(index);
        }
    }
}

bool ParseComments(const std::wstring& content, std::vector<CommentRule>* out, std::wstring* error)
{
    std::vector<CommentRule> rules;
    win32ui::JsonReader reader(content, win32ui::JsonErrorText(error), 4);
    bool format = false;
    const bool read = reader.Object([&](const std::wstring& member) {
        if (member == L"format")
        {
            std::wstring value;
            format = reader.String(&value, 64) && value == kFormat;
            return format || reader.Fail(util::Tr(L"The file isn't a RegKit comments file."));
        }
        if (member == L"comments")
        {
            return reader.Array([&] { return ReadRule(reader, &rules.emplace_back()); });
        }
        return reader.Fail(util::Tr(L"The file contains an unknown member."));
    });
    if (!read || !reader.End() || (!format && !reader.Fail(util::Tr(L"The file isn't a RegKit comments file."))))
    {
        return false;
    }
    for (CommentRule& rule : rules)
    {
        if (util::IsBlank(rule.text))
        {
            rule.text.clear();
        }
    }
    *out = std::move(rules);
    return true;
}

bool ValidateCatalog(const std::vector<CommentRule>& rules)
{
    for (size_t left = 0; left < rules.size(); ++left)
    {
        const CommentRule& rule = rules[left];
        if (util::IsBlank(rule.text))
        {
            return false;
        }
        for (size_t right = left + 1; right < rules.size(); ++right)
        {
            const CommentRule& other = rules[right];
            if (rule.key == other.key && util::EqualsInsensitive(rule.name, other.name) &&
                Specificity(rule) == Specificity(other) && (!rule.type || !other.type || rule.type == other.type) &&
                (!rule.data_size || !other.data_size || rule.data_size == other.data_size) &&
                (rule.key_scope == CommentKeyScope::kAny || util::EqualsInsensitive(rule.key_path, other.key_path)))
            {
                return false;
            }
        }
    }
    return true;
}

std::wstring SerializeComments(const ValueComments& comments)
{
    std::wstring out = L"{\n  \"format\": \"regkit-comments\",\n  \"comments\": [";
    bool first = true;
    for (const CommentRule& rule : comments.rules())
    {
        out.append(first ? L"\n    {" : L",\n    {");
        if (!rule.key)
        {
            AppendMember(&out, L"name", rule.name, true);
        }
        if (rule.type)
        {
            const bool named = *rule.type <= REG_QWORD;
            AppendMember(&out, L"type", named ? value_format::TypeName(*rule.type) : std::to_wstring(*rule.type), named);
        }
        if (rule.data_size)
        {
            AppendMember(&out, L"size", std::to_wstring(*rule.data_size), false);
        }
        if (rule.key_scope != CommentKeyScope::kAny)
        {
            AppendMember(&out, rule.key_scope == CommentKeyScope::kExact ? L"key" : L"tree", rule.key_path, true);
        }
        AppendMember(&out, L"text", rule.text, true);
        out.append(L"\n    }");
        first = false;
    }
    out.append(first ? L"]\n}\n" : L"\n  ]\n}\n");
    return out;
}

CommentRule ValueRule(const CommentTarget& target)
{
    CommentRule rule;
    rule.key = target.key;
    rule.name = target.key ? std::wstring() : target.name;
    rule.type = target.key ? std::nullopt : std::optional<DWORD>(target.type);
    rule.key_scope = CommentKeyScope::kExact;
    rule.key_path = NormalizeKeyPath(target.path);
    return rule;
}

ResolvedComment ResolveComment(const ValueComments& user, const ValueComments& defaults, const CommentTarget& target)
{
    for (const auto& [comments, source] :
         {std::pair{&user, CommentSource::kUser}, std::pair{&defaults, CommentSource::kDefault}})
    {
        if (const CommentRule* rule = comments->Match(target))
        {
            return {rule->text, source, *rule};
        }
    }
    return {};
}

}
