// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "regfile/script_convert.h"
#include "registry/value_format.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <algorithm>
#include <climits>
#include <cwctype>
#include <initializer_list>
#include <unordered_map>

namespace regkit::regfile
{
namespace
{

constexpr wchar_t kNativeHelper[] =
    L"Add-Type -Namespace RegKit -Name Native -MemberDefinition '[DllImport(\"advapi32.dll\", CharSet = CharSet.Unicode)] "
    L"static extern int RegSetValueExW(Microsoft.Win32.SafeHandles.SafeRegistryHandle key, string name, int reserved, int type, byte[] data, int size); "
    L"public static void SetValue(Microsoft.Win32.RegistryKey key, string name, int type, byte[] data) { int status = "
    L"RegSetValueExW(key.Handle, name, 0, type, data, data.Length); if (status != 0) throw new System.ComponentModel.Win32Exception(status); }'";

struct RootProperty
{
    const wchar_t* root;
    const wchar_t* property;
};

constexpr RootProperty kRoots[] = {
    {L"HKEY_CLASSES_ROOT", L"ClassesRoot"},
    {L"HKEY_CURRENT_USER", L"CurrentUser"},
    {L"HKEY_LOCAL_MACHINE", L"LocalMachine"},
    {L"HKEY_USERS", L"Users"},
    {L"HKEY_CURRENT_CONFIG", L"CurrentConfig"},
};

struct ParseError
{
    size_t line;
    std::wstring message;
};

bool IsSingleQuote(wchar_t character)
{
    return character == L'\'' || (character >= 0x2018 && character <= 0x201B);
}

bool IsDoubleQuote(wchar_t character)
{
    return character == L'"' || (character >= 0x201C && character <= 0x201E);
}

bool IsIdentifier(wchar_t character)
{
    return iswalnum(character) || character == L'_';
}

struct Token
{
    enum class Kind
    {
        kEnd,
        kNewline,
        kString,
        kNumber,
        kVariable,
        kParameter,
        kType,
        kWord,
        kPunct,
    };
    Kind kind = Kind::kEnd;
    std::wstring text;
    std::vector<std::pair<bool, std::wstring>> parts;
    long long number = 0;
    bool wide = false;
    size_t line = 0;
};

using Kind = Token::Kind;

std::vector<std::pair<bool, std::wstring>> ExpandableParts(std::wstring_view body, size_t line)
{
    std::vector<std::pair<bool, std::wstring>> parts(1);
    for (size_t index = 0; index < body.size(); ++index)
    {
        const wchar_t character = body[index];
        if (character == L'`' && index + 1 < body.size())
        {
            const wchar_t next = body[++index];
            const std::wstring_view codes = L"0abefnrtv";
            const std::wstring_view values(L"\0\a\b\x1b\f\n\r\t\v", 9);
            const size_t code = codes.find(next);
            parts.back().second.push_back(code == std::wstring_view::npos ? next : values[code]);
        }
        else if (character == L'$' && index + 1 < body.size() && body[index + 1] == L'(')
        {
            throw ParseError{line, util::Tr(L"Subexpressions aren't supported.")};
        }
        else if (character == L'$' && index + 1 < body.size() && (IsIdentifier(body[index + 1]) || body[index + 1] == L'{'))
        {
            size_t end = index + 1;
            std::wstring name;
            if (body[end] == L'{')
            {
                const size_t close = body.find(L'}', end);
                if (close == std::wstring_view::npos)
                {
                    throw ParseError{line, util::Tr(L"Unterminated variable name.")};
                }
                name = body.substr(end + 1, close - end - 1);
                end = close + 1;
            }
            else
            {
                while (end < body.size() && (IsIdentifier(body[end]) || (body[end] == L':' && end + 1 < body.size() && IsIdentifier(body[end + 1]))))
                {
                    name.push_back(body[end++]);
                }
            }
            parts.push_back({true, name});
            parts.push_back({});
            index = end - 1;
        }
        else
        {
            parts.back().second.push_back(character);
        }
    }
    return parts;
}

std::vector<Token> Tokenize(std::wstring_view text)
{
    std::vector<Token> tokens;
    size_t index = 0;
    size_t line = 1;
    int depth = 0;
    size_t last_end = std::wstring_view::npos;
    bool member = false;
    auto at = [&](size_t position) { return position < text.size() ? text[position] : L'\0'; };
    auto fail = [&](const std::wstring& message) { throw ParseError{line, message}; };
    auto push = [&](Kind kind, std::wstring value) -> Token& {
        Token token;
        token.kind = kind;
        token.text = std::move(value);
        token.line = line;
        tokens.push_back(std::move(token));
        return tokens.back();
    };
    auto count_lines = [&](size_t from, size_t to) {
        for (size_t position = from; position < to; ++position)
        {
            line += text[position] == L'\n';
        }
    };
    while (index < text.size())
    {
        const wchar_t character = text[index];
        const size_t start = index;
        const bool adjacent = last_end == index;
        const bool was_member = member;
        member = false;
        if (character == L' ' || character == L'\t' || character == L'\r' || character == 0xFEFF)
        {
            ++index;
            continue;
        }
        if (character == L'`' && (at(index + 1) == L'\n' || (at(index + 1) == L'\r' && at(index + 2) == L'\n')))
        {
            index += at(index + 1) == L'\r' ? 3 : 2;
            ++line;
            continue;
        }
        if (character == L'\n' || character == L';')
        {
            const bool joins = !tokens.empty() && tokens.back().kind == Kind::kPunct && (tokens.back().text == L"|" || tokens.back().text == L",");
            if (depth == 0 && !(character == L'\n' && joins))
            {
                push(Kind::kNewline, {});
            }
            line += character == L'\n';
            ++index;
            continue;
        }
        if (character == L'#')
        {
            while (index < text.size() && text[index] != L'\n')
            {
                ++index;
            }
            continue;
        }
        if (character == L'<' && at(index + 1) == L'#')
        {
            const size_t close = text.find(L"#>", index + 2);
            if (close == std::wstring_view::npos)
            {
                fail(util::Tr(L"Unterminated block comment."));
            }
            count_lines(index, close);
            index = close + 2;
            continue;
        }
        if (character == L'@' && (IsSingleQuote(at(index + 1)) || IsDoubleQuote(at(index + 1))) && (at(index + 2) == L'\n' || at(index + 2) == L'\r'))
        {
            const bool literal = IsSingleQuote(at(index + 1));
            const size_t body = text.find(L'\n', index) + 1;
            size_t close = body;
            while (true)
            {
                close = text.find(L'\n', close);
                if (close == std::wstring_view::npos)
                {
                    fail(util::Tr(L"Unterminated here-string."));
                }
                ++close;
                if ((literal ? IsSingleQuote(at(close)) : IsDoubleQuote(at(close))) && at(close + 1) == L'@')
                {
                    break;
                }
            }
            size_t end = close - 1;
            end -= end > body && text[end - 1] == L'\r';
            const std::wstring_view content = end > body ? text.substr(body, end - body) : std::wstring_view{};
            Token& token = push(Kind::kString, std::wstring(content));
            if (!literal)
            {
                token.parts = ExpandableParts(content, line);
            }
            count_lines(index, close);
            index = close + 2;
        }
        else if (IsSingleQuote(character))
        {
            std::wstring value;
            for (++index;; ++index)
            {
                if (index >= text.size())
                {
                    fail(util::Tr(L"Unterminated string."));
                }
                if (IsSingleQuote(text[index]) && !IsSingleQuote(at(index + 1)))
                {
                    break;
                }
                index += IsSingleQuote(text[index]);
                value.push_back(text[index]);
            }
            Token& token = push(Kind::kString, value);
            count_lines(start, ++index);
            token.line = line;
        }
        else if (IsDoubleQuote(character))
        {
            size_t end = index + 1;
            for (; end < text.size(); ++end)
            {
                if (text[end] == L'`')
                {
                    ++end;
                }
                else if (IsDoubleQuote(text[end]) && IsDoubleQuote(at(end + 1)))
                {
                    ++end;
                }
                else if (IsDoubleQuote(text[end]))
                {
                    break;
                }
            }
            if (end >= text.size())
            {
                fail(util::Tr(L"Unterminated string."));
            }
            std::wstring body;
            for (size_t position = index + 1; position < end; ++position)
            {
                if ((text[position] == L'`' || IsDoubleQuote(text[position])) && position + 1 < end)
                {
                    body.push_back(L'`');
                    ++position;
                }
                body.push_back(text[position]);
            }
            Token& token = push(Kind::kString, {});
            token.parts = ExpandableParts(body, line);
            count_lines(index, end);
            index = end + 1;
        }
        else if (character == L'$')
        {
            ++index;
            std::wstring name;
            if (at(index) == L'{')
            {
                const size_t close = text.find(L'}', index);
                if (close == std::wstring_view::npos)
                {
                    fail(util::Tr(L"Unterminated variable name."));
                }
                name = text.substr(index + 1, close - index - 1);
                index = close + 1;
            }
            else
            {
                for (; IsIdentifier(at(index)) || (at(index) == L':' && IsIdentifier(at(index + 1))); ++index)
                {
                    name.push_back(text[index]);
                }
            }
            push(Kind::kVariable, name);
        }
        else if (character == L'[')
        {
            int nesting = 0;
            size_t end = index;
            for (; end < text.size(); ++end)
            {
                nesting += text[end] == L'[' ? 1 : text[end] == L']' ? -1
                                                                     : 0;
                if (nesting == 0)
                {
                    break;
                }
            }
            if (end >= text.size())
            {
                fail(util::Tr(L"Unterminated type name."));
            }
            push(Kind::kType, util::TrimWhitespace(text.substr(index + 1, end - index - 1)));
            index = end + 1;
        }
        else if (character == L'@' && at(index + 1) == L'(')
        {
            push(Kind::kPunct, L"@(");
            ++depth;
            index += 2;
        }
        else if (character == L':' && at(index + 1) == L':')
        {
            push(Kind::kPunct, L"::");
            index += 2;
            member = true;
        }
        else if (character == L'.' && adjacent && !tokens.empty())
        {
            push(Kind::kPunct, L".");
            ++index;
            member = true;
        }
        else if (std::wstring_view(L"(){},=|!").find(character) != std::wstring_view::npos)
        {
            depth += character == L'(' ? 1 : character == L')' ? -1
                                                               : 0;
            push(Kind::kPunct, std::wstring(1, character));
            ++index;
        }
        else if (character == L'-' && iswalpha(at(index + 1)))
        {
            std::wstring name;
            for (++index; IsIdentifier(at(index)); ++index)
            {
                name.push_back(text[index]);
            }
            if (at(index) == L':')
            {
                name.push_back(text[index++]);
            }
            push(Kind::kParameter, name);
        }
        else if (!was_member && (iswdigit(character) || (character == L'-' && iswdigit(at(index + 1)))))
        {
            const bool negative = character == L'-';
            index += negative;
            const bool hex = at(index) == L'0' && (at(index + 1) == L'x' || at(index + 1) == L'X');
            index += hex ? 2 : 0;
            const size_t digits = index;
            while (hex ? util::HexDigitValue(at(index)) >= 0 : iswdigit(at(index)) != 0)
            {
                ++index;
            }
            const bool suffix = at(index) == L'l' || at(index) == L'L';
            index += suffix;
            const std::optional<uint64_t> parsed = index == digits || IsIdentifier(at(index)) ? std::nullopt : win32ui::ParseUnsigned(text.substr(digits, index - digits - suffix), hex ? 16 : 10);
            if (!parsed)
            {
                fail(util::TrLabel(L"Unsupported number", text.substr(start, index - start + 1)));
            }
            const unsigned long long value = parsed.value_or(0);
            Token& token = push(Kind::kNumber, std::wstring(text.substr(start, index - start)));
            token.wide = suffix || (hex ? value > 0xFFFFFFFFull : value > 0x7FFFFFFFull);
            token.number = hex && !token.wide ? static_cast<int>(static_cast<unsigned int>(value)) : static_cast<long long>(value);
            token.number = negative ? -token.number : token.number;
        }
        else
        {
            std::wstring word;
            for (; index < text.size(); ++index)
            {
                const wchar_t next = text[index];
                if (iswspace(next) || std::wstring_view(L"(){},;|=\"'`$").find(next) != std::wstring_view::npos || (was_member && !IsIdentifier(next)) ||
                    (next == L':' && at(index + 1) == L':') || IsSingleQuote(next) || IsDoubleQuote(next))
                {
                    break;
                }
                word.push_back(next);
            }
            if (word.empty())
            {
                fail(util::TrLabel(L"Unexpected character", std::wstring(1, character)));
            }
            push(Kind::kWord, word);
        }
        last_end = index;
    }
    push(Kind::kNewline, {});
    push(Kind::kEnd, {});
    return tokens;
}

struct Item
{
    enum class Type
    {
        kNull,
        kString,
        kNumber,
        kList,
        kKey,
    };
    Type type = Type::kNull;
    std::wstring text;
    long long number = 0;
    bool wide = false;
    std::wstring cast;
    std::vector<Item> items;
};

Item StringItem(std::wstring text)
{
    Item item;
    item.type = Item::Type::kString;
    item.text = std::move(text);
    return item;
}

Item NumberItem(long long number, bool wide)
{
    Item item;
    item.type = Item::Type::kNumber;
    item.number = number;
    item.wide = wide;
    return item;
}

Item KeyItem(std::wstring path)
{
    Item item;
    item.type = Item::Type::kKey;
    item.text = std::move(path);
    return item;
}

std::vector<Item> Elements(const Item& item)
{
    return item.type == Item::Type::kList ? item.items : std::vector<Item>{item};
}

std::wstring PropertyName(std::wstring name)
{
    return util::EqualsInsensitive(name, L"(default)") ? std::wstring() : name;
}

struct Argument
{
    std::wstring name;
    Item value;
};

class Interpreter
{
  public:
    Interpreter(std::vector<Token> tokens, std::vector<Operation>* output)
        : tokens_(std::move(tokens)), output_(output)
    {
    }

    void Run()
    {
        while (!At(Kind::kEnd))
        {
            if (At(Kind::kNewline))
            {
                ++position_;
                continue;
            }
            if (Statement())
            {
                return;
            }
            if (!At(Kind::kNewline) && !At(Kind::kEnd))
            {
                Fail(util::TrLabel(L"Unexpected token", Peek().text));
            }
        }
    }

  private:
    enum class Condition
    {
        kNone,
        kMissing,
        kExists,
    };

    const Token& Peek(size_t ahead = 0) const
    {
        return tokens_[std::min(position_ + ahead, tokens_.size() - 1)];
    }

    bool At(Kind kind, std::wstring_view text = {}, size_t ahead = 0) const
    {
        const Token& token = Peek(ahead);
        return token.kind == kind && (text.empty() || util::EqualsInsensitive(token.text, text));
    }

    const Token& Take()
    {
        const Token& token = Peek();
        position_ = std::min(position_ + 1, tokens_.size() - 1);
        return token;
    }

    [[noreturn]] void Fail(const std::wstring& message) const
    {
        throw ParseError{Peek().line, message};
    }

    void Expect(std::wstring_view punct)
    {
        if (!At(Kind::kPunct, punct))
        {
            Fail(util::TrLabel(L"Expected", punct));
        }
        Take();
    }

    bool Statement()
    {
        if (At(Kind::kWord, L"exit") || At(Kind::kWord, L"return"))
        {
            return true;
        }
        if (At(Kind::kWord, L"if"))
        {
            IfStatement();
            return false;
        }
        if (At(Kind::kVariable) && At(Kind::kPunct, L"=", 1))
        {
            const std::wstring name = util::ToLower(Take().text);
            Take();
            Item value = At(Kind::kWord) ? Command() : Expression(false);
            if (name != L"null")
            {
                variables_[name] = std::move(value);
            }
            return false;
        }
        if (At(Kind::kWord))
        {
            Command();
            return false;
        }
        Expression(false);
        Pipeline();
        return false;
    }

    void IfStatement()
    {
        if (condition_ != Condition::kNone)
        {
            Fail(util::Tr(L"Only if conditions on Test-Path are supported."));
        }
        Take();
        Expect(L"(");
        const bool negate = At(Kind::kPunct, L"!") || At(Kind::kParameter, L"not");
        if (negate)
        {
            Take();
        }
        const bool parenthesized = At(Kind::kPunct, L"(");
        if (parenthesized)
        {
            Take();
        }
        if (!At(Kind::kWord, L"Test-Path"))
        {
            Fail(util::Tr(L"Only if conditions on Test-Path are supported."));
        }
        Take();
        const std::vector<Argument> arguments = Arguments();
        for (const Argument& argument : arguments)
        {
            static constexpr std::wstring_view kKnown[] = {L"", L"path", L"literalpath", L"lp", L"pspath", L"erroraction", L"ea"};
            if (std::find(std::begin(kKnown), std::end(kKnown), argument.name) == std::end(kKnown))
            {
                Fail(util::TrLabel(L"Unsupported parameter", L"-" + argument.name));
            }
        }
        size_t shift = 0;
        const std::vector<std::wstring> tested = TargetKeys(arguments, false, &shift);
        if (tested.size() != 1)
        {
            Fail(util::Tr(L"Only if conditions on Test-Path are supported."));
        }
        tested_ = tested.front();
        if (parenthesized)
        {
            Expect(L")");
        }
        Expect(L")");
        while (At(Kind::kNewline))
        {
            Take();
        }
        Expect(L"{");
        condition_ = negate ? Condition::kMissing : Condition::kExists;
        while (!At(Kind::kPunct, L"}"))
        {
            if (At(Kind::kEnd))
            {
                Fail(util::Tr(L"Expected '}'."));
            }
            if (At(Kind::kNewline))
            {
                Take();
                continue;
            }
            if (Statement())
            {
                Fail(util::TrLabel(L"Unexpected token", Peek().text));
            }
        }
        condition_ = Condition::kNone;
        Take();
        if (At(Kind::kWord, L"else") || At(Kind::kWord, L"elseif"))
        {
            Fail(util::Tr(L"else blocks aren't supported."));
        }
    }

    void Pipeline()
    {
        while (At(Kind::kPunct, L"|"))
        {
            Take();
            if (!At(Kind::kWord, L"Out-Null"))
            {
                Fail(util::Tr(L"Only | Out-Null pipelines are supported."));
            }
            Take();
        }
    }

    std::vector<Argument> Arguments()
    {
        static constexpr std::wstring_view kSwitches[] = {L"force", L"recurse", L"confirm", L"passthru", L"verbose", L"whatif"};
        std::vector<Argument> arguments;
        while (!At(Kind::kNewline) && !At(Kind::kEnd) && !At(Kind::kPunct, L"|") && !At(Kind::kPunct, L")") && !At(Kind::kPunct, L"}"))
        {
            Argument argument;
            if (At(Kind::kParameter))
            {
                argument.name = util::ToLower(Take().text);
                const bool colon = argument.name.back() == L':';
                if (colon)
                {
                    argument.name.pop_back();
                }
                const bool is_switch = std::find(std::begin(kSwitches), std::end(kSwitches), argument.name) != std::end(kSwitches);
                if (argument.name == L"whatif")
                {
                    Fail(util::Tr(L"-WhatIf makes no changes and can't be converted."));
                }
                if (colon)
                {
                    argument.value = Unary(true);
                }
                else if (!is_switch)
                {
                    argument.value = Expression(true);
                }
                if (is_switch && argument.name != L"force" && argument.name != L"recurse")
                {
                    continue;
                }
            }
            else
            {
                argument.value = Expression(true);
            }
            arguments.push_back(std::move(argument));
        }
        return arguments;
    }

    Item Command()
    {
        const std::wstring name = util::ToLower(Take().text);
        std::vector<Argument> arguments = Arguments();
        Pipeline();
        Execute(name, arguments);
        return {};
    }

    Item Expression(bool argument_mode)
    {
        Item first = Unary(argument_mode);
        if (!At(Kind::kPunct, L","))
        {
            return first;
        }
        Item list;
        list.type = Item::Type::kList;
        list.items.push_back(std::move(first));
        while (At(Kind::kPunct, L","))
        {
            Take();
            list.items.push_back(Unary(argument_mode));
        }
        return list;
    }

    Item Unary(bool argument_mode)
    {
        if (++depth_ > 256)
        {
            Fail(util::TrLabel(L"Unexpected token", Peek().text));
        }
        Item item;
        if (At(Kind::kType) && !At(Kind::kPunct, L"::", 1))
        {
            const std::wstring type = Take().text;
            item = Cast(type, Unary(argument_mode));
        }
        else
        {
            item = Postfix(Primary(argument_mode));
        }
        --depth_;
        return item;
    }

    Item Primary(bool argument_mode)
    {
        const Token& token = Take();
        switch (token.kind)
        {
        case Kind::kString:
            return token.parts.empty() ? StringItem(token.text) : Interpolate(token.parts);
        case Kind::kNumber:
            return NumberItem(token.number, token.wide);
        case Kind::kVariable:
            {
                const std::wstring name = util::ToLower(token.text);
                if (name == L"true" || name == L"false")
                {
                    return NumberItem(name == L"true", false);
                }
                if (name == L"null")
                {
                    return {};
                }
                const auto variable = variables_.find(name);
                if (variable == variables_.end())
                {
                    Fail(util::TrLabel(L"Variable can't be resolved", L"$" + token.text));
                }
                return variable->second;
            }
        case Kind::kWord:
            if (argument_mode)
            {
                return StringItem(token.text);
            }
            break;
        case Kind::kType:
            return StaticMember(token.text);
        case Kind::kPunct:
            if (token.text == L"(" || token.text == L"@(")
            {
                Item value;
                if (At(Kind::kWord))
                {
                    value = Command();
                }
                else if (!At(Kind::kPunct, L")"))
                {
                    value = Expression(false);
                }
                Expect(L")");
                if (token.text == L"@(" && value.type != Item::Type::kList)
                {
                    Item list;
                    list.type = Item::Type::kList;
                    if (value.type != Item::Type::kNull)
                    {
                        list.items.push_back(std::move(value));
                    }
                    return list;
                }
                return value;
            }
            break;
        default:
            break;
        }
        Fail(util::TrLabel(L"Unexpected token", token.text));
    }

    Item Interpolate(const std::vector<std::pair<bool, std::wstring>>& parts)
    {
        std::wstring text;
        for (const auto& [variable, value] : parts)
        {
            if (!variable)
            {
                text += value;
                continue;
            }
            const auto found = variables_.find(util::ToLower(value));
            if (found == variables_.end() || (found->second.type != Item::Type::kString && found->second.type != Item::Type::kNumber))
            {
                Fail(util::TrLabel(L"Variable can't be resolved", L"$" + value));
            }
            text += found->second.type == Item::Type::kString ? found->second.text : std::to_wstring(found->second.number);
        }
        return StringItem(text);
    }

    std::vector<Item> CallArguments()
    {
        std::vector<Item> arguments;
        Expect(L"(");
        while (!At(Kind::kPunct, L")"))
        {
            arguments.push_back(Unary(false));
            if (!At(Kind::kPunct, L")"))
            {
                Expect(L",");
            }
        }
        Take();
        return arguments;
    }

    std::wstring TakeMember()
    {
        Take();
        if (!At(Kind::kWord))
        {
            Fail(util::Tr(L"Expected a member name."));
        }
        return util::ToLower(Take().text);
    }

    Item StaticMember(std::wstring type)
    {
        type = util::ToLower(type);
        if (!At(Kind::kPunct, L"::"))
        {
            Fail(util::TrLabel(L"Unexpected type", L"[" + type + L"]"));
        }
        const std::wstring member = TakeMember();
        if (type == L"microsoft.win32.registry")
        {
            for (const RootProperty& root : kRoots)
            {
                if (util::EqualsInsensitive(member, root.property))
                {
                    return KeyItem(root.root);
                }
            }
        }
        else if (type == L"microsoft.win32.registryvaluekind" || type == L"microsoft.win32.registryhive" || type == L"microsoft.win32.registryview")
        {
            return StringItem(member);
        }
        else if (type == L"microsoft.win32.registrykey" && member == L"openbasekey")
        {
            const std::vector<Item> arguments = CallArguments();
            for (const RootProperty& root : kRoots)
            {
                if (!arguments.empty() && util::EqualsInsensitive(arguments[0].text, root.property) && (arguments.size() == 1 || util::EqualsInsensitive(arguments[1].text, L"Default")))
                {
                    return KeyItem(root.root);
                }
            }
        }
        else if (type == L"regkit.native" && member == L"setvalue")
        {
            const std::vector<Item> arguments = CallArguments();
            if (arguments.size() != 4 || arguments[0].type != Item::Type::kKey || arguments[2].type != Item::Type::kNumber)
            {
                Fail(util::Tr(L"Invalid RegKit.Native call."));
            }
            Operation operation{Operation::Kind::kValue, arguments[0].text};
            operation.value.name = Text(arguments[1]);
            operation.value.type = static_cast<DWORD>(arguments[2].number);
            operation.value.data = Bytes(arguments[3]);
            output_->push_back(std::move(operation));
            return {};
        }
        Fail(util::TrLabel(L"Unsupported member", L"[" + type + L"]::" + member));
    }

    Item Postfix(Item item)
    {
        while (At(Kind::kPunct, L"."))
        {
            const std::wstring member = TakeMember();
            if (item.type != Item::Type::kKey || !At(Kind::kPunct, L"("))
            {
                Fail(util::TrLabel(L"Unsupported member", L"." + member));
            }
            const std::vector<Item> arguments = CallArguments();
            const std::wstring sub = arguments.empty() ? std::wstring() : Text(arguments[0]);
            const std::wstring path = sub.empty() ? item.text : item.text + L"\\" + sub;
            if (member == L"createsubkey" || member == L"opensubkey")
            {
                if (member == L"createsubkey")
                {
                    output_->push_back({Operation::Kind::kKey, path});
                }
                item = KeyItem(path);
            }
            else if (member == L"setvalue" && (arguments.size() == 2 || arguments.size() == 3))
            {
                Set(item.text, sub, arguments[1], arguments.size() == 3 ? Text(arguments[2]) : std::wstring());
                item = {};
            }
            else if (member == L"deletevalue")
            {
                Operation operation{Operation::Kind::kRemoveValue, item.text};
                operation.value.name = sub;
                output_->push_back(std::move(operation));
                item = {};
            }
            else if (member == L"deletesubkeytree" || member == L"deletesubkey")
            {
                output_->push_back({Operation::Kind::kRemoveKey, path});
                item = {};
            }
            else if (member == L"close" || member == L"dispose" || member == L"flush")
            {
                item = {};
            }
            else
            {
                Fail(util::TrLabel(L"Unsupported member", L"." + member + L"()"));
            }
        }
        return item;
    }

    std::wstring Text(const Item& item) const
    {
        if (item.type == Item::Type::kNumber)
        {
            return std::to_wstring(item.number);
        }
        if (item.type != Item::Type::kString)
        {
            Fail(util::Tr(L"Expected text."));
        }
        return item.text;
    }

    std::vector<BYTE> Bytes(const Item& item) const
    {
        std::vector<BYTE> bytes;
        for (const Item& element : Elements(item))
        {
            if (element.type != Item::Type::kNumber || element.number < 0 || element.number > 0xFF)
            {
                Fail(util::Tr(L"Byte values must be numbers from 0 to 255."));
            }
            bytes.push_back(static_cast<BYTE>(element.number));
        }
        return bytes;
    }

    Item Cast(std::wstring type, Item value) const
    {
        type = util::ToLower(type);
        if (util::StartsWithInsensitive(type, L"system."))
        {
            type.erase(0, 7);
        }
        if (type == L"void")
        {
            return {};
        }
        if (type == L"byte[]" || type == L"string[]")
        {
            Item list;
            list.type = Item::Type::kList;
            list.cast = type;
            for (const Item& element : Elements(value))
            {
                if (type == L"string[]")
                {
                    list.items.push_back(StringItem(Text(element)));
                    continue;
                }
                const std::vector<BYTE> bytes = Bytes(element);
                if (bytes.size() != 1)
                {
                    Fail(util::Tr(L"Byte values must be numbers from 0 to 255."));
                }
                list.items.push_back(NumberItem(bytes[0], false));
            }
            return list;
        }
        if (type == L"string")
        {
            return StringItem(Text(value));
        }
        if (value.type == Item::Type::kNumber && (type == L"int" || type == L"int32"))
        {
            if (value.number < INT_MIN || value.number > INT_MAX)
            {
                Fail(util::Tr(L"The number doesn't fit an [int]."));
            }
            return NumberItem(value.number, false);
        }
        if (value.type == Item::Type::kNumber && (type == L"long" || type == L"int64" || type == L"uint32" || type == L"uint64" || type == L"ulong"))
        {
            return NumberItem(value.number, true);
        }
        if (type == L"microsoft.win32.registryvaluekind")
        {
            return value;
        }
        Fail(util::TrLabel(L"Unsupported cast", L"[" + type + L"]"));
    }

    void Set(const std::wstring& path, const std::wstring& name, const Item& data, std::wstring kind)
    {
        kind = util::ToLower(kind);
        if (kind.empty())
        {
            kind = data.type == Item::Type::kString   ? L"string"
                   : data.type == Item::Type::kNumber ? (data.wide ? L"qword" : L"dword")
                   : data.cast == L"byte[]"           ? L"binary"
                   : data.cast == L"string[]"         ? L"multistring"
                                                      : L"";
        }
        Operation operation{Operation::Kind::kValue, path};
        operation.value.name = name;
        Value& value = operation.value;
        if (kind == L"string" || kind == L"expandstring")
        {
            value.type = kind == L"string" ? REG_SZ : REG_EXPAND_SZ;
            value.data = value_format::StringData(Text(data));
        }
        else if (kind == L"dword" || kind == L"qword")
        {
            const bool qword = kind == L"qword";
            unsigned long long number = static_cast<unsigned long long>(data.number);
            if (data.type == Item::Type::kString)
            {
                const std::optional<uint64_t> parsed = win32ui::ParseUnsigned(data.text, 10, true);
                if (!parsed)
                {
                    Fail(util::TrLabel(L"Invalid number", data.text));
                }
                number = parsed.value_or(number);
            }
            if (data.type != Item::Type::kString && data.type != Item::Type::kNumber)
            {
                Fail(util::Tr(L"Expected a number."));
            }
            if (!qword && (data.type == Item::Type::kNumber ? data.number < INT_MIN || data.number > 0xFFFFFFFFll : number > 0xFFFFFFFFull))
            {
                Fail(util::Tr(L"The number doesn't fit a DWORD."));
            }
            value.type = qword ? REG_QWORD : REG_DWORD;
            value.data = value_format::UnsignedBytes(number, qword ? sizeof(ULONGLONG) : sizeof(DWORD));
        }
        else if (kind == L"binary" || kind == L"unknown" || kind == L"none")
        {
            value.type = kind == L"none" ? REG_NONE : REG_BINARY;
            value.data = Bytes(data);
        }
        else if (kind == L"multistring")
        {
            std::vector<std::wstring> items;
            for (const Item& element : Elements(data))
            {
                items.push_back(Text(element));
            }
            value.type = REG_MULTI_SZ;
            value.data = value_format::MultiStringData(items);
        }
        else
        {
            Fail(kind.empty() ? std::wstring(util::Tr(L"The value type can't be inferred, add -Type.")) : util::TrLabel(L"Unsupported value type", kind));
        }
        output_->push_back(std::move(operation));
    }

    std::vector<std::wstring> Paths(const Item& item, bool literal, bool wildcards_literal) const
    {
        std::vector<std::wstring> paths;
        for (const Item& element : Elements(item))
        {
            std::wstring text = Text(element);
            if (util::StartsWithInsensitive(text, L"Microsoft.PowerShell.Core\\"))
            {
                text.erase(0, 26);
            }
            std::wstring root;
            if (util::StartsWithInsensitive(text, L"Registry::"))
            {
                text.erase(0, 10);
            }
            else if (const size_t colon = text.find(L':'); colon != std::wstring::npos)
            {
                const std::wstring drive = util::ToLower(text.substr(0, colon));
                const auto mapped = drives_.find(drive);
                root = drive == L"hklm" ? L"HKEY_LOCAL_MACHINE" : drive == L"hkcu"      ? L"HKEY_CURRENT_USER"
                                                              : mapped != drives_.end() ? mapped->second
                                                                                        : L"";
                if (root.empty())
                {
                    Fail(util::TrLabel(L"Not a registry path", text));
                }
                text.erase(0, colon + 1);
                text = root + (text.empty() || text.front() == L'\\' ? L"" : L"\\") + text;
            }
            if (!literal)
            {
                std::replace(text.begin(), text.end(), L'/', L'\\');
                if (!wildcards_literal && text.find_first_of(L"*?[") != std::wstring::npos)
                {
                    Fail(util::Tr(L"Wildcard paths can't be converted, use -LiteralPath."));
                }
            }
            std::wstring path;
            if (!NormalizeKeyPath(text, &path))
            {
                Fail(util::TrLabel(L"Not a registry path", text));
            }
            paths.push_back(std::move(path));
        }
        return paths;
    }

    // keys from -LiteralPath, -Path or first positional argument, shift 1 for named arguments
    std::vector<std::wstring> TargetKeys(const std::vector<Argument>& arguments, bool wildcards_literal, size_t* shift) const
    {
        const Item* literal = Find(arguments, {L"literalpath", L"lp", L"pspath"}, SIZE_MAX);
        const Item* named = literal ? literal : Find(arguments, {L"path"}, SIZE_MAX);
        const Item* path = named ? named : Find(arguments, {}, 0);
        if (!path)
        {
            Fail(util::Tr(L"A -Path is required."));
        }
        *shift = named ? 1 : 0;
        return Paths(*path, literal != nullptr, wildcards_literal);
    }

    void Execute(const std::wstring& name, const std::vector<Argument>& arguments)
    {
        static constexpr std::wstring_view kIgnored[] = {L"out-null", L"write-host", L"write-output", L"write-verbose", L"write-warning", L"set-strictmode", L"add-type", L"clear-host", L"start-sleep", L"test-path"};
        if (std::find(std::begin(kIgnored), std::end(kIgnored), name) != std::end(kIgnored))
        {
            return;
        }
        const bool new_item = name == L"new-item" || name == L"ni";
        const bool set_property = name == L"set-itemproperty" || name == L"sp" || name == L"new-itemproperty";
        const bool remove_property = name == L"remove-itemproperty" || name == L"rp";
        const bool remove_item = name == L"remove-item" || name == L"ri" || name == L"del" || name == L"rm" || name == L"rd" || name == L"rmdir" || name == L"erase";
        if (name == L"new-psdrive" || name == L"ndr")
        {
            const Item* drive = Find(arguments, {L"name"}, 0);
            const Item* provider = Find(arguments, {L"psprovider"}, 1);
            const Item* root = Find(arguments, {L"root"}, 2);
            std::wstring path;
            if (!drive || !provider || !root || !util::EqualsInsensitive(Text(*provider), L"Registry") ||
                !NormalizeKeyPath(util::StartsWithInsensitive(Text(*root), L"Registry::") ? Text(*root).substr(10) : Text(*root), &path))
            {
                Fail(util::Tr(L"Only registry drives are supported."));
            }
            drives_[util::ToLower(Text(*drive))] = path;
            return;
        }
        if (!new_item && !set_property && !remove_property && !remove_item)
        {
            Fail(util::TrLabel(L"Unsupported command", name));
        }
        if ((condition_ == Condition::kMissing && !new_item) || (condition_ == Condition::kExists && !remove_property && !remove_item))
        {
            Fail(util::Tr(L"Only New-Item can follow if (-not (Test-Path)), and only removals can follow if (Test-Path)."));
        }
        size_t shift = 0;
        const std::vector<std::wstring> keys = TargetKeys(arguments, new_item, &shift);
        const Item* force_value = Find(arguments, {L"force"}, SIZE_MAX);
        if (force_value && force_value->type != Item::Type::kNull && force_value->type != Item::Type::kNumber)
        {
            Fail(util::TrLabel(L"Unsupported parameter", L"-force"));
        }
        const bool force = force_value && (force_value->type == Item::Type::kNull || force_value->number != 0);
        for (const Argument& argument : arguments)
        {
            static constexpr std::wstring_view kKnown[] = {L"", L"path", L"literalpath", L"lp", L"pspath", L"name", L"value", L"type", L"propertytype", L"force", L"recurse", L"erroraction", L"ea", L"itemtype"};
            if (std::find(std::begin(kKnown), std::end(kKnown), argument.name) == std::end(kKnown))
            {
                Fail(util::TrLabel(L"Unsupported parameter", L"-" + argument.name));
            }
        }
        const Item* leaf = new_item ? Find(arguments, {L"name"}, 1 - shift) : nullptr;
        const Item* default_value = new_item ? Find(arguments, {L"value"}, SIZE_MAX) : nullptr;
        for (const std::wstring& key : keys)
        {
            const std::wstring full = leaf ? key + L"\\" + Text(*leaf) : key;
            if (condition_ != Condition::kNone && (!util::EqualsInsensitive(full, tested_) || default_value))
            {
                Fail(util::Tr(L"Only New-Item can follow if (-not (Test-Path)), and only removals can follow if (Test-Path)."));
            }
            if (new_item)
            {
                if (force && condition_ == Condition::kNone)
                {
                    output_->push_back({Operation::Kind::kRemoveKey, full});
                }
                output_->push_back({Operation::Kind::kKey, full});
                if (default_value)
                {
                    Set(full, std::wstring(), *default_value, L"string");
                }
            }
            else if (set_property)
            {
                const Item* property = Find(arguments, {L"name"}, 1 - shift);
                const Item* value = Find(arguments, {L"value"}, 2 - shift);
                const Item* type = Find(arguments, {L"type", L"propertytype"}, SIZE_MAX);
                if (!property || !value)
                {
                    Fail(util::Tr(L"-Name and -Value are required."));
                }
                Set(key, PropertyName(Text(*property)), *value, type ? Text(*type) : std::wstring());
            }
            else if (remove_property)
            {
                const Item* property = Find(arguments, {L"name"}, 1 - shift);
                if (!property)
                {
                    Fail(util::Tr(L"-Name is required."));
                }
                for (const Item& element : Elements(*property))
                {
                    Operation operation{Operation::Kind::kRemoveValue, key};
                    operation.value.name = PropertyName(Text(element));
                    if (operation.value.name.find_first_of(L"*?[") != std::wstring::npos)
                    {
                        Fail(util::Tr(L"Wildcard value names can't be converted."));
                    }
                    output_->push_back(std::move(operation));
                }
            }
            else
            {
                output_->push_back({Operation::Kind::kRemoveKey, key});
            }
        }
    }

    static const Item* Find(const std::vector<Argument>& arguments, std::initializer_list<std::wstring_view> names, size_t position)
    {
        size_t positional = 0;
        for (const Argument& argument : arguments)
        {
            if (argument.name.empty() ? positional++ == position : std::find(names.begin(), names.end(), argument.name) != names.end())
            {
                return &argument.value;
            }
        }
        return nullptr;
    }

    std::vector<Token> tokens_;
    size_t position_ = 0;
    std::vector<Operation>* output_;
    std::unordered_map<std::wstring, Item> variables_;
    std::unordered_map<std::wstring, std::wstring> drives_;
    Condition condition_ = Condition::kNone;
    std::wstring tested_;
    int depth_ = 0;
};

std::wstring Hex(unsigned long long value, int width)
{
    wchar_t text[24] = {};
    swprintf_s(text, L"0x%0*llx", width, value);
    return text;
}

// prefer single quotes, otherwise double quotes with backtick escapes
bool Literal(std::wstring_view text, std::wstring* output)
{
    const bool plain = std::none_of(text.begin(), text.end(), [](wchar_t character) { return character < 0x20; });
    output->assign(1, plain ? L'\'' : L'"');
    for (wchar_t character : text)
    {
        if (plain)
        {
            if (IsSingleQuote(character))
            {
                output->push_back(character);
            }
            output->push_back(character);
            continue;
        }
        const std::wstring_view values(L"\0\t\r\n", 4);
        const size_t code = values.find(character);
        if (code != std::wstring_view::npos)
        {
            output->push_back(L'`');
            output->push_back(L"0trn"[code]);
            continue;
        }
        if (character < 0x20)
        {
            return false;
        }
        if (character == L'`' || character == L'$' || IsDoubleQuote(character))
        {
            output->push_back(L'`');
        }
        output->push_back(character);
    }
    output->push_back(plain ? L'\'' : L'"');
    return true;
}

std::wstring ByteArray(const std::vector<BYTE>& data)
{
    if (data.empty())
    {
        return L"[byte[]]@()";
    }
    std::wstring text = L"[byte[]](";
    for (size_t index = 0; index < data.size(); ++index)
    {
        text.append(index == 0 ? L"" : L",").append(Hex(data[index], 2));
    }
    return text + L")";
}

bool DotNetValue(const Value& value, std::wstring* expression, std::wstring* kind)
{
    std::wstring text;
    switch (value.type)
    {
    case REG_SZ:
    case REG_EXPAND_SZ:
        *kind = value.type == REG_SZ ? L"String" : L"ExpandString";
        return value_format::DecodeString(value.data, &text) && value_format::StringData(text) == value.data && Literal(text, expression);
    case REG_MULTI_SZ:
        {
            const std::vector<std::wstring> items = MultiStringItems(value.data);
            if (value_format::MultiStringData(items) != value.data)
            {
                return false;
            }
            *kind = L"MultiString";
            *expression = L"[string[]]@(";
            for (size_t index = 0; index < items.size(); ++index)
            {
                if (!Literal(items[index], &text))
                {
                    return false;
                }
                expression->append(index == 0 ? L"" : L", ").append(text);
            }
            expression->push_back(L')');
            return true;
        }
    case REG_DWORD:
    case REG_QWORD:
        {
            const size_t width = value.type == REG_DWORD ? sizeof(DWORD) : sizeof(ULONGLONG);
            *kind = value.type == REG_DWORD ? L"DWord" : L"QWord";
            *expression = Hex(value_format::ReadUnsigned(value.data, width), static_cast<int>(width * 2)) + (value.type == REG_QWORD ? L"L" : L"");
            return value.data.size() == width;
        }
    case REG_BINARY:
    case REG_NONE:
        *kind = value.type == REG_BINARY ? L"Binary" : L"None";
        *expression = ByteArray(value.data);
        return true;
    default:
        return false;
    }
}

}

bool ParsePowerShell(std::wstring_view content, std::vector<Operation>* output, std::wstring* error)
{
    output->clear();
    try
    {
        Interpreter(Tokenize(content), output).Run();
        return true;
    }
    catch (const ParseError& failure)
    {
        *error = util::Tr(L"Line") + std::wstring(L" ") + std::to_wstring(failure.line) + L": " + failure.message;
        return false;
    }
}

std::wstring RenderPowerShell(const std::vector<Operation>& operations, bool admin_check, std::vector<std::wstring>* skipped)
{
    std::wstring body;
    bool native = false;
    const std::wstring* current = nullptr;
    for (const Operation& operation : operations)
    {
        const size_t split = operation.path.find(L'\\');
        const std::wstring_view root_name = std::wstring_view(operation.path).substr(0, split);
        const std::wstring sub = split == std::wstring::npos ? std::wstring() : operation.path.substr(split + 1);
        std::wstring root = L"[Microsoft.Win32.Registry]::";
        for (const RootProperty& entry : kRoots)
        {
            root += root_name == entry.root ? entry.property : L"";
        }
        std::wstring sub_literal;
        std::wstring name;
        if (!Literal(sub, &sub_literal) || !Literal(operation.value.name, &name))
        {
            skipped->push_back(Describe(operation, util::Tr(L"contains control characters")));
            continue;
        }
        if (operation.kind == Operation::Kind::kRemoveKey)
        {
            if (sub.empty())
            {
                skipped->push_back(Describe(operation, util::Tr(L"a root key can't be deleted")));
                continue;
            }
            body += root + L".DeleteSubKeyTree(" + sub_literal + L", $false)\r\n";
            current = nullptr;
            continue;
        }
        if (!current || *current != operation.path)
        {
            body += L"$key = " + root + (sub.empty() ? L"" : L".CreateSubKey(" + sub_literal + L")") + L"\r\n";
            current = &operation.path;
        }
        if (operation.kind == Operation::Kind::kRemoveValue)
        {
            body += L"$key.DeleteValue(" + name + L", $false)\r\n";
        }
        else if (operation.kind == Operation::Kind::kValue)
        {
            std::wstring expression;
            std::wstring kind;
            if (DotNetValue(operation.value, &expression, &kind))
            {
                body += L"$key.SetValue(" + name + L", " + expression + L", '" + kind + L"')\r\n";
            }
            else
            {
                native = true;
                body += L"[RegKit.Native]::SetValue($key, " + name + L", " + Hex(operation.value.type, 1) + L", " + ByteArray(operation.value.data) + L")\r\n";
            }
        }
    }
    return std::wstring(admin_check && NeedsAdmin(operations) ? L"#Requires -RunAsAdministrator\r\n" : L"") + L"$ErrorActionPreference = 'Stop'\r\n" +
           (native ? kNativeHelper : L"") + (native ? L"\r\n" : L"") + body;
}

}
