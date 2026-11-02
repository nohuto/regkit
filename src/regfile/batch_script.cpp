// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "regfile/reg_exe_syntax.h"
#include "regfile/script_convert.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/text_transform.h"
#include "win32/translation.h"

#include <algorithm>
#include <cwctype>
#include <unordered_map>

namespace regkit::regfile
{
namespace
{

constexpr size_t kMaxCommandLine = 8191;
constexpr std::wstring_view kDefaultSeparator = L"\\0";

// MSVCRT argument quoting rules used by reg.exe
std::vector<std::wstring> SplitArguments(std::wstring_view text)
{
    std::vector<std::wstring> args;
    size_t index = 0;
    auto blank = [&](size_t at) { return text[at] == L' ' || text[at] == L'\t'; };
    while (true)
    {
        while (index < text.size() && blank(index))
        {
            ++index;
        }
        if (index >= text.size())
        {
            return args;
        }
        std::wstring arg;
        bool quoted = false;
        while (index < text.size() && (quoted || !blank(index)))
        {
            size_t slashes = 0;
            while (index < text.size() && text[index] == L'\\')
            {
                ++slashes;
                ++index;
            }
            if (index < text.size() && text[index] == L'"')
            {
                arg.append(slashes / 2, L'\\');
                if (slashes % 2 != 0 || (quoted && index + 1 < text.size() && text[index + 1] == L'"'))
                {
                    arg.push_back(L'"');
                    index += slashes % 2 != 0 ? 1 : 2;
                    continue;
                }
                quoted = !quoted;
                ++index;
                continue;
            }
            arg.append(slashes, L'\\');
            if (index < text.size() && (quoted || !blank(index)))
            {
                arg.push_back(text[index++]);
            }
        }
        args.push_back(std::move(arg));
    }
}

std::wstring Quote(std::wstring_view text)
{
    std::wstring output = L"\"";
    size_t slashes = 0;
    for (wchar_t character : text)
    {
        if (character == L'\\')
        {
            ++slashes;
            continue;
        }
        output.append(character == L'"' ? slashes * 2 : slashes, L'\\');
        slashes = 0;
        if (character == L'"' || character == L'%')
        {
            output.push_back(character);
        }
        output.push_back(character);
    }
    output.append(slashes * 2, L'\\');
    output.push_back(L'"');
    return output;
}

bool HasLineBreak(std::wstring_view text)
{
    return text.find_first_of(std::wstring_view(L"\r\n\0", 3)) != std::wstring_view::npos;
}

bool IsPadded(std::wstring_view text)
{
    return !text.empty() && (iswspace(text.front()) || iswspace(text.back()));
}

// reg.exe rejects key names at the 255-character registry limit
bool HasLongName(std::wstring_view path)
{
    for (size_t start = 0; start <= path.size();)
    {
        const size_t end = std::min(path.find(L'\\', start), path.size());
        if (end - start > 254)
        {
            return true;
        }
        start = end + 1;
    }
    return false;
}

bool BatchData(const Value& value, std::wstring* type, std::wstring* text, std::wstring* separator)
{
    switch (value.type)
    {
    case REG_SZ:
    case REG_EXPAND_SZ:
    case REG_NONE:
        if (!value_format::DecodeString(value.data, text))
        {
            return false;
        }
        break;
    case REG_MULTI_SZ:
        {
            const std::vector<std::wstring> items = MultiStringItems(value.data);
            auto unused = [&](std::wstring_view candidate) {
                return std::none_of(items.begin(), items.end(), [&](const std::wstring& item) { return item.find(candidate) != std::wstring::npos; });
            };
            *separator = kDefaultSeparator;
            for (const wchar_t* candidate = L"|;,~#@$*+="; !unused(*separator) && *candidate; ++candidate)
            {
                *separator = std::wstring(1, *candidate);
            }
            for (size_t index = 0; index < items.size(); ++index)
            {
                text->append(index == 0 ? L"" : *separator).append(items[index]);
            }
            if (items.size() == 1 && items.front().empty())
            {
                *text = *separator;
            }
            break;
        }
    case REG_DWORD:
    case REG_DWORD_BIG_ENDIAN:
    case REG_QWORD:
        {
            const size_t width = value.type == REG_QWORD ? sizeof(ULONGLONG) : sizeof(DWORD);
            if (value.data.size() != width)
            {
                return false;
            }
            wchar_t number[24] = {};
            swprintf_s(number, L"0x%llx", value_format::ReadUnsigned(value.data, width));
            *text = number;
            break;
        }
    case REG_BINARY:
        *text = win32ui::HexBytes(value.data, L'\0');
        break;
    default:
        return false;
    }
    *type = reg_exe::TypeName(value.type);
    std::vector<BYTE> check;
    return !HasLineBreak(*text) && reg_exe::BuildData(value.type, *text, *separator, &check, nullptr) && check == value.data;
}

bool ApplyReg(const std::vector<std::wstring>& args, std::vector<Operation>* output, std::wstring* message)
{
    const bool add = !args.empty() && util::EqualsInsensitive(args[0], L"add");
    if (!add && (args.empty() || !util::EqualsInsensitive(args[0], L"delete")))
    {
        *message = util::TrLabel(L"Unsupported reg command", L"reg " + (args.empty() ? std::wstring() : args[0]));
        return false;
    }
    reg_exe::Options options;
    std::vector<std::wstring> positional;
    if (!reg_exe::ParseOptions(args, 1, &options, &positional, add ? reg_exe::Verb::kAdd : reg_exe::Verb::kOther, message))
    {
        return false;
    }
    Operation operation{Operation::Kind::kValue};
    if (positional.size() != 1 || !NormalizeKeyPath(positional[0], &operation.path))
    {
        *message = util::Tr(L"Invalid key name.");
        return false;
    }
    if (options.view == KEY_WOW64_32KEY)
    {
        *message = util::Tr(L"/reg:32 targets the 32-bit view, which can't be converted.");
        return false;
    }
    if (options.all_values)
    {
        *message = util::Tr(L"/va isn't supported.");
        return false;
    }
    operation.value.name = options.default_value ? std::wstring() : options.value_name;
    if (!add)
    {
        operation.kind = options.has_value ? Operation::Kind::kRemoveValue : Operation::Kind::kRemoveKey;
        output->push_back(std::move(operation));
        return true;
    }
    operation.value.type = REG_SZ;
    if (!options.type_text.empty() && !reg_exe::ParseType(options.type_text, &operation.value.type))
    {
        *message = util::TrLabel(L"Invalid type", options.type_text);
        return false;
    }
    if (!reg_exe::BuildData(operation.value.type, options.data, options.separator, &operation.value.data, message))
    {
        return false;
    }
    if (!options.has_value)
    {
        output->push_back({Operation::Kind::kKey, operation.path});
    }
    output->push_back(std::move(operation));
    return true;
}

struct Command
{
    std::wstring text;
    std::wstring separator;
};

std::vector<std::wstring> RegArguments(const Command& command)
{
    std::vector<std::wstring> args = SplitArguments(command.text);
    if (!args.empty())
    {
        args.erase(args.begin());
    }
    return args;
}

bool IsEnsureKey(const std::vector<Command>& commands, size_t first)
{
    if (commands.size() != first + 3 || commands[first + 1].separator != L"||" || commands[first + 2].separator != L"&&")
    {
        return false;
    }
    const std::vector<std::wstring> query = RegArguments(commands[first]);
    const std::vector<std::wstring> add = RegArguments(commands[first + 1]);
    const std::vector<std::wstring> remove = RegArguments(commands[first + 2]);
    return query.size() == 2 && add.size() == 3 && remove.size() == 4 && util::EqualsInsensitive(add[0], L"add") &&
           add[1] == query[1] && reg_exe::IsSwitch(add[2], L"f") && util::EqualsInsensitive(remove[0], L"delete") &&
           remove[1] == query[1] && reg_exe::IsSwitch(remove[2], L"ve") && reg_exe::IsSwitch(remove[3], L"f");
}

bool SplitCommands(std::wstring_view line, std::vector<Command>* commands, std::wstring* message)
{
    commands->assign(1, {});
    bool quoted = false;
    for (size_t index = 0; index < line.size(); ++index)
    {
        const wchar_t character = line[index];
        std::wstring& text = commands->back().text;
        if (character == L'"' || quoted)
        {
            quoted = quoted != (character == L'"');
        }
        else if (character == L'^')
        {
            if (index + 1 < line.size())
            {
                text.push_back(line[++index]);
            }
            continue;
        }
        else if (character == L'&' || character == L'|')
        {
            std::wstring separator(1, character);
            if (index + 1 < line.size() && line[index + 1] == character)
            {
                separator.push_back(line[++index]);
            }
            if (separator == L"|")
            {
                *message = util::Tr(L"Pipes aren't supported.");
                return false;
            }
            commands->push_back({{}, separator});
            continue;
        }
        else if (character == L'(' || character == L')')
        {
            continue;
        }
        else if (character == L'>' || character == L'<')
        {
            if (!text.empty() && iswdigit(text.back()) && (text.size() == 1 || text[text.size() - 2] == L' '))
            {
                text.pop_back();
            }
            size_t next = index + 1;
            if (next < line.size() && line[next] == L'>')
            {
                ++next;
            }
            if (next < line.size() && line[next] == L'&')
            {
                next += 2;
            }
            else
            {
                while (next < line.size() && line[next] == L' ')
                {
                    ++next;
                }
                while (next < line.size() && std::wstring_view(L" &|<>()").find(line[next]) == std::wstring_view::npos)
                {
                    ++next;
                }
            }
            index = next - 1;
            continue;
        }
        text.push_back(character);
    }
    return true;
}

bool ExpandPercents(std::wstring_view line, const std::unordered_map<std::wstring, std::wstring>& variables, std::wstring* output, std::wstring* message)
{
    for (size_t index = 0; index < line.size(); ++index)
    {
        if (line[index] != L'%')
        {
            output->push_back(line[index]);
            continue;
        }
        if (index + 1 < line.size() && line[index + 1] == L'%')
        {
            output->push_back(L'%');
            ++index;
            continue;
        }
        if (index + 1 < line.size() && (iswdigit(line[index + 1]) || line[index + 1] == L'~' || line[index + 1] == L'*'))
        {
            *message = util::Tr(L"Batch arguments can't be resolved.");
            return false;
        }
        const size_t close = line.find(L'%', index + 1);
        if (close == std::wstring_view::npos)
        {
            continue;
        }
        const std::wstring name(line.substr(index + 1, close - index - 1));
        const auto variable = variables.find(util::ToLower(name));
        if (variable == variables.end())
        {
            *message = util::TrLabel(L"Variable can't be resolved", L"%" + name + L"%");
            return false;
        }
        output->append(variable->second);
        index = close;
    }
    return true;
}

bool IsWord(std::wstring_view text, std::wstring_view word)
{
    return util::StartsWithInsensitive(text, word) && (text.size() == word.size() || text[word.size()] == L' ' || text[word.size()] == L'\t');
}

}

bool ParseBatch(std::wstring_view content, std::vector<Operation>* output, std::wstring* error)
{
    output->clear();
    std::vector<std::wstring> lines;
    for (size_t start = 0; start <= content.size();)
    {
        size_t end = content.find(L'\n', start);
        end = end == std::wstring_view::npos ? content.size() : end;
        std::wstring line(content.substr(start, end - start));
        if (!line.empty() && line.back() == L'\r')
        {
            line.pop_back();
        }
        lines.push_back(std::move(line));
        start = end + 1;
    }

    std::unordered_map<std::wstring, std::wstring> variables;
    bool delayed = false;
    std::vector<std::pair<std::unordered_map<std::wstring, std::wstring>, bool>> scopes;
    for (size_t index = 0; index < lines.size(); ++index)
    {
        const size_t number = index + 1;
        auto fail = [&](const std::wstring& message) {
            *error = util::Tr(L"Line") + std::wstring(L" ") + std::to_wstring(number) + L": " + message;
            return false;
        };
        std::wstring line = lines[index];
        while (!line.empty() && line.back() == L'^' && index + 1 < lines.size())
        {
            line.pop_back();
            line += lines[++index];
        }
        std::wstring text = util::TrimWhitespace(line);
        text.erase(0, text.find_first_not_of(L"@ \t"));
        if (text.empty() || text.front() == L':' || IsWord(text, L"rem"))
        {
            continue;
        }
        std::wstring expanded;
        std::wstring message;
        std::vector<Command> commands;
        if (!ExpandPercents(text, variables, &expanded, &message) || !SplitCommands(expanded, &commands, &message))
        {
            return fail(message);
        }
        for (size_t at = 0; at < commands.size(); ++at)
        {
            if (commands[at].separator == L"||")
            {
                break;
            }
            std::wstring command = util::TrimWhitespace(commands[at].text);
            command.erase(0, command.find_first_not_of(L"@ \t"));
            const size_t space = command.find_first_of(L" \t");
            const std::wstring verb = util::ToLower(command.substr(0, space));
            const std::wstring rest = space == std::wstring::npos ? std::wstring() : command.substr(space + 1);
            const bool reg = verb == L"reg" || verb == L"reg.exe" || util::EndsWithInsensitive(verb, L"\\reg.exe");
            // delayed expansion rewrites every ! at run time
            if (delayed && (reg || verb == L"set") && command.find(L'!') != std::wstring::npos)
            {
                return fail(util::TrLabel(L"Unsupported command", L"setlocal EnableDelayedExpansion"));
            }
            if (reg)
            {
                const std::vector<std::wstring> args = SplitArguments(rest);
                if (!args.empty() && util::EqualsInsensitive(args[0], L"query"))
                {
                    std::wstring path;
                    if (!IsEnsureKey(commands, at) || !NormalizeKeyPath(args.size() > 1 ? args[1] : std::wstring(), &path))
                    {
                        return fail(util::Tr(L"reg query is only supported in the form RegKit writes for empty keys."));
                    }
                    output->push_back({Operation::Kind::kKey, std::move(path)});
                    break;
                }
                if (!ApplyReg(args, output, &message))
                {
                    return fail(message);
                }
            }
            else if (verb == L"set")
            {
                std::wstring assignment = util::TrimWhitespace(rest);
                if (!assignment.empty() && assignment.front() == L'/')
                {
                    return fail(util::Tr(L"set /a and set /p aren't supported."));
                }
                if (!assignment.empty() && assignment.front() == L'"')
                {
                    assignment = assignment.substr(1, assignment.rfind(L'"') - 1);
                }
                const size_t equals = assignment.find(L'=');
                if (equals != std::wstring::npos && equals != 0)
                {
                    variables[util::ToLower(assignment.substr(0, equals))] = assignment.substr(equals + 1);
                }
            }
            else if (verb == L"setlocal")
            {
                scopes.emplace_back(variables, delayed);
                delayed = util::ContainsInsensitive(rest, L"enabledelayedexpansion") || (delayed && !util::ContainsInsensitive(rest, L"disabledelayedexpansion"));
            }
            else if (verb == L"endlocal")
            {
                if (!scopes.empty())
                {
                    variables = std::move(scopes.back().first);
                    delayed = scopes.back().second;
                    scopes.pop_back();
                }
            }
            else if (verb == L"exit")
            {
                return true;
            }
            else if (!util::StartsWithInsensitive(verb, L"echo") && verb != L"title" && verb != L"cls" && verb != L"color" &&
                     verb != L"pause" && verb != L"chcp" && verb != L"timeout" &&
                     !(verb == L"net" && util::EqualsInsensitive(util::TrimWhitespace(rest), L"session")) && !verb.empty())
            {
                return fail(util::TrLabel(L"Unsupported command", verb));
            }
        }
    }
    return true;
}

std::wstring RenderBatch(const std::vector<Operation>& operations, bool admin_check, std::vector<std::wstring>* skipped)
{
    std::wstring body;
    // defer key creation to the following value line, emit it if that line is skipped
    std::wstring pending_key;
    for (size_t index = 0; index < operations.size(); ++index)
    {
        const Operation& operation = operations[index];
        const bool value = operation.kind == Operation::Kind::kValue || operation.kind == Operation::Kind::kRemoveValue;
        const std::wstring key = Quote(registry_path::Format(operation.path, registry_path::Style::kAbbreviated));
        const std::wstring name = operation.value.name.empty() ? L"/ve" : L"/v " + Quote(operation.value.name);
        std::wstring line;
        std::wstring reason;
        if (HasLineBreak(operation.path) || (value && HasLineBreak(operation.value.name)))
        {
            reason = util::Tr(L"contains a line break");
        }
        else if (IsPadded(operation.path) || (value && IsPadded(operation.value.name)))
        {
            reason = util::Tr(L"reg.exe trims leading and trailing spaces");
        }
        else if (HasLongName(operation.path))
        {
            reason = util::Tr(L"Invalid key name.");
        }
        else if (operation.kind == Operation::Kind::kKey)
        {
            const Operation* next = index + 1 < operations.size() ? &operations[index + 1] : nullptr;
            line = L"reg query " + key + L" >nul 2>&1 || (reg add " + key + L" /f >nul && reg delete " + key + L" /ve /f >nul)";
            if (next && next->kind != Operation::Kind::kRemoveKey && next->kind != Operation::Kind::kRemoveValue && IsUnderKey(next->path, operation.path))
            {
                pending_key = line + L"\r\n";
                continue;
            }
        }
        else if (operation.kind == Operation::Kind::kRemoveKey)
        {
            line = L"reg delete " + key + L" /f >nul 2>&1";
            reason = operation.path.find(L'\\') == std::wstring::npos ? util::Tr(L"a root key can't be deleted") : L"";
        }
        else if (operation.kind == Operation::Kind::kRemoveValue)
        {
            line = L"reg delete " + key + L" " + name + L" /f >nul 2>&1";
        }
        else
        {
            std::wstring type;
            std::wstring data;
            std::wstring separator(kDefaultSeparator);
            if (BatchData(operation.value, &type, &data, &separator))
            {
                line = L"reg add " + key + L" " + name + L" /t " + type +
                       (separator == kDefaultSeparator ? L"" : L" /s " + Quote(separator)) + L" /d " + Quote(data) + L" /f >nul";
            }
            else
            {
                reason = util::TrLabel(L"reg.exe can't write this data type", reg_exe::TypeName(operation.value.type));
            }
        }
        if (reason.empty() && line.size() > kMaxCommandLine)
        {
            reason = util::Tr(L"too long for a batch line");
        }
        if (reason.empty() && !line.empty() && !WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, line.data(), static_cast<int>(line.size()), nullptr, 0, nullptr, nullptr))
        {
            reason = util::Tr(L"The text isn't valid UTF-16.");
        }
        if (!reason.empty())
        {
            skipped->push_back(Describe(operation, reason));
            body.append(pending_key);
            pending_key.clear();
            continue;
        }
        pending_key.clear();
        body.append(line).append(L"\r\n");
    }
    const bool ascii = std::all_of(body.begin(), body.end(), [](wchar_t character) { return character < 0x80; });
    return L"@echo off\r\n" + std::wstring(ascii ? L"" : L"chcp 65001 >nul\r\n") +
           (admin_check && NeedsAdmin(operations) ? L"net session >nul 2>&1 || (echo Run this script as administrator & exit /b 1)\r\n" : L"") + body;
}

}
