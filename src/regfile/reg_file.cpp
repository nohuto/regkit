// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "regfile/reg_file.h"
#include "win32/text_transform.h"

#include "registry/key_algorithms.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "regfile/script_convert.h"
#include "win32/file_text.h"
#include "win32/translation.h"

#include <win32ui/system/error_text.h>

#include <algorithm>
#include <cstring>

namespace regkit::regfile
{
namespace
{

bool ParseQuoted(std::wstring_view text, std::wstring* output)
{
    if (!output || text.empty() || text.front() != L'"')
    {
        return false;
    }
    output->clear();
    bool escaped = false;
    for (size_t index = 1; index < text.size(); ++index)
    {
        const wchar_t character = text[index];
        if (escaped)
        {
            switch (character)
            {
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
        }
        else if (character == L'\\')
        {
            escaped = true;
        }
        else if (character == L'"')
        {
            for (size_t rest = index + 1; rest < text.size(); ++rest)
            {
                if (text[rest] != L' ' && text[rest] != L'\t')
                {
                    return false;
                }
            }
            return true;
        }
        else
        {
            output->push_back(character);
        }
    }
    return false;
}

size_t FindAssignment(std::wstring_view text)
{
    if (text.empty() || text.front() != L'"')
    {
        return text.find(L'=');
    }
    bool escaped = false;
    for (size_t index = 1; index < text.size(); ++index)
    {
        const wchar_t character = text[index];
        if (escaped)
        {
            escaped = false;
        }
        else if (character == L'\\')
        {
            escaped = true;
        }
        else if (character == L'"')
        {
            return text.find(L'=', index + 1);
        }
    }
    return std::wstring_view::npos;
}

std::wstring Escape(std::wstring_view text)
{
    std::wstring output;
    output.reserve(text.size());
    for (wchar_t character : text)
    {
        if (character == L'\\' || character == L'"')
        {
            output.push_back(L'\\');
        }
        output.push_back(character);
    }
    return output;
}

void AppendHexLine(std::wstring* output, std::wstring_view prefix, const std::vector<BYTE>& data)
{
    constexpr size_t kHexLineLimit = 77;
    constexpr wchar_t kDigits[] = L"0123456789abcdef";
    output->append(prefix);
    size_t line = prefix.size();
    for (size_t index = 0; index < data.size(); ++index)
    {
        output->push_back(kDigits[data[index] >> 4]);
        output->push_back(kDigits[data[index] & 0x0F]);
        if (index + 1 == data.size())
        {
            break;
        }
        output->push_back(L',');
        line += 3;
        if (line >= kHexLineLimit)
        {
            output->append(L"\\\r\n  ");
            line = 2;
        }
    }
    output->append(L"\r\n");
}

std::wstring ValueNamePrefix(const std::wstring& name)
{
    return name.empty() ? L"@=" : L"\"" + Escape(name) + L"\"=";
}

void AppendValue(std::wstring* output, const Value& value)
{
    const std::wstring name = ValueNamePrefix(value.name);
    std::wstring text;
    if (value.type == REG_SZ && value_format::DecodeString(value.data, &text) && value_format::StringData(text) == value.data &&
        text.find_first_of(L"\r\n") == std::wstring::npos)
    {
        output->append(name).append(L"\"").append(Escape(text)).append(L"\"\r\n");
        return;
    }
    if (value.type == REG_DWORD && value.data.size() == sizeof(DWORD))
    {
        DWORD number = 0;
        std::memcpy(&number, value.data.data(), sizeof(number));
        wchar_t dword[24] = {};
        swprintf_s(dword, L"dword:%08x\r\n", number);
        output->append(name).append(dword);
        return;
    }
    wchar_t code[16] = {};
    swprintf_s(code, L"hex(%x):", value.type);
    AppendHexLine(output, name + (value.type == REG_BINARY ? L"hex:" : code), value.data);
}

constexpr std::wstring_view kRegFileHeader = L"Windows Registry Editor Version 5.00\r\n\r\n";

}

Writer::Writer()
    : output_(kRegFileHeader)
{
}

void Writer::AppendKeyHeader(std::wstring_view prefix, std::wstring_view path)
{
    if (output_.size() > kRegFileHeader.size())
    {
        output_ += L"\r\n";
    }
    output_.append(prefix).append(path).append(L"]\r\n");
}

void Writer::AppendRemovedKey(std::wstring_view path)
{
    AppendKeyHeader(L"[-", path);
}

void Writer::AppendValue(const Value& value)
{
    regfile::AppendValue(&output_, value);
}

void Writer::AppendRemovedValue(std::wstring_view name)
{
    output_.append(ValueNamePrefix(std::wstring(name))).append(L"-\r\n");
}

void Writer::AppendKey(std::wstring_view path, std::vector<const Value*> values, bool sorted)
{
    AppendKeyHeader(L"[", path);
    if (sorted)
    {
        std::sort(values.begin(), values.end(), [](const Value* left, const Value* right) {
            return left->name.empty() != right->name.empty() ? left->name.empty()
                                                             : util::CompareInsensitive(left->name, right->name) < 0;
        });
    }
    for (const Value* value : values)
    {
        regfile::AppendValue(&output_, *value);
    }
}

std::wstring Writer::Finish() &&
{
    if (output_.size() > kRegFileHeader.size())
    {
        output_ += L"\r\n";
    }
    return std::move(output_);
}

namespace
{

// depth first export order, visit receives display path and contents of each readable key
template <typename Visit>
LONG VisitRegistryTree(HKEY root, const std::wstring& subkey, const std::wstring& display_path, REGSAM view, bool recurse, std::vector<std::wstring>* skipped, std::vector<std::wstring>* lost_volatility, Visit visit)
{
    if (subkey.find(L'\0') != std::wstring::npos)
    {
        return ERROR_INVALID_NAME;
    }
    struct Pending
    {
        std::wstring subkey;
        std::wstring display;
        bool volatile_parent = false;
    };
    std::vector<Pending> pending{{subkey, display_path}};
    LONG top_status = ERROR_SUCCESS;
    for (bool top = true; !pending.empty(); top = false)
    {
        const Pending current = std::move(pending.back());
        pending.pop_back();
        registry_backend::KeyContents contents;
        bool is_volatile = current.volatile_parent;
        const LONG status = registry_backend::ReadKeyContents(root, current.subkey, view, true, &contents, lost_volatility && !is_volatile ? &is_volatile : nullptr);
        if (top)
        {
            top_status = status;
        }
        if (status != ERROR_SUCCESS)
        {
            if (!top && skipped)
            {
                skipped->push_back(Describe({Operation::Kind::kKey, current.display}, win32ui::FormatWin32Error(static_cast<DWORD>(status))));
            }
            continue;
        }
        if (is_volatile && !current.volatile_parent)
        {
            lost_volatility->push_back(Describe({Operation::Kind::kKey, current.display}, util::Tr(L"only the volatile flag of it and its subkeys")));
        }
        SkipNullNames(current.display, &contents.values, &contents.subkeys, skipped);
        visit(current.display, contents);
        for (auto child = contents.subkeys.rbegin(); recurse && child != contents.subkeys.rend(); ++child)
        {
            pending.push_back({registry_path::JoinSubkey(current.subkey, *child), current.display + L"\\" + *child, is_volatile});
        }
    }
    return top_status;
}

}

void SkipNullNames(const std::wstring& display_path, std::vector<Value>* values, std::vector<std::wstring>* subkeys, std::vector<std::wstring>* skipped)
{
    // reg files and reg.exe cannot address names with embedded NULs
    const auto unaddressable = [&](const Operation& operation, const std::wstring& name) {
        if (name.find(L'\0') == std::wstring::npos)
        {
            return false;
        }
        if (skipped)
        {
            skipped->push_back(Describe(operation, util::Tr(L"the name contains a null character")));
        }
        return true;
    };
    std::erase_if(*values, [&](const Value& value) { return unaddressable({Operation::Kind::kValue, display_path, value}, value.name); });
    if (subkeys)
    {
        std::erase_if(*subkeys, [&](const std::wstring& name) { return unaddressable({Operation::Kind::kKey, display_path + L"\\" + name}, name); });
    }
}

LONG AppendRegistryTree(Writer* writer, HKEY root, const std::wstring& subkey, const std::wstring& display_path, REGSAM view, bool recurse, std::vector<std::wstring>* skipped, std::vector<std::wstring>* lost_volatility)
{
    return VisitRegistryTree(root, subkey, display_path, view, recurse, skipped, lost_volatility, [&](const std::wstring& display, const registry_backend::KeyContents& contents) {
        std::vector<const Value*> values;
        values.reserve(contents.values.size());
        for (const Value& value : contents.values)
        {
            values.push_back(&value);
        }
        writer->AppendKey(display, std::move(values), false);
    });
}

LONG ReadRegistryOperations(HKEY root, const std::wstring& subkey, const std::wstring& display_path, REGSAM view, bool recurse, std::vector<Operation>* output, std::vector<std::wstring>* skipped, std::vector<std::wstring>* lost_volatility)
{
    return VisitRegistryTree(root, subkey, display_path, view, recurse, skipped, lost_volatility, [&](const std::wstring& display, registry_backend::KeyContents& contents) {
        output->push_back({Operation::Kind::kKey, display});
        for (Value& value : contents.values)
        {
            output->push_back({Operation::Kind::kValue, display, std::move(value)});
        }
    });
}

namespace
{

// REGEDIT4 stores hex string data as ANSI bytes
std::vector<BYTE> AnsiToUtf16(const std::vector<BYTE>& data)
{
    const std::wstring text = win32ui::MultiByteToWide(std::string_view(reinterpret_cast<const char*>(data.data()), data.size()), CP_ACP);
    std::vector<BYTE> output(text.size() * sizeof(wchar_t));
    std::memcpy(output.data(), text.data(), output.size());
    return output;
}

bool ParseValueData(const std::wstring& data_text, bool ansi, Value* value)
{
    if (data_text.front() == L'"')
    {
        std::wstring text;
        if (!ParseQuoted(data_text, &text))
        {
            return false;
        }
        value->type = REG_SZ;
        value->data = value_format::StringData(text);
        return true;
    }
    if (util::StartsWithInsensitive(data_text, L"dword:"))
    {
        const std::optional<uint64_t> number = win32ui::ParseUnsigned(std::wstring_view(data_text).substr(6), 16);
        if (!number || *number > 0xFFFFFFFFull)
        {
            return false;
        }
        value->type = REG_DWORD;
        value->data = value_format::UnsignedBytes(*number, sizeof(DWORD));
        return true;
    }
    if (!util::StartsWithInsensitive(data_text, L"hex"))
    {
        return false;
    }
    const size_t colon = data_text.find(L':');
    if (colon == std::wstring::npos)
    {
        return false;
    }
    value->type = REG_BINARY;
    const size_t open = data_text.find(L'(');
    const size_t close = data_text.find(L')');
    if (open != std::wstring::npos && close != std::wstring::npos && close > open && close < colon)
    {
        const std::optional<uint64_t> type = win32ui::ParseUnsigned(std::wstring_view(data_text).substr(open + 1, close - open - 1), 16);
        if (!type || *type > 0xFFFFFFFFull)
        {
            return false;
        }
        value->type = static_cast<DWORD>(*type);
    }
    else if (colon != 3)
    {
        return false;
    }
    if (!win32ui::ParseHexBytes(std::wstring_view(data_text).substr(colon + 1), &value->data))
    {
        return false;
    }
    if (ansi && (value->type == REG_SZ || value->type == REG_EXPAND_SZ || value->type == REG_MULTI_SZ))
    {
        value->data = AnsiToUtf16(value->data);
    }
    return true;
}

}

bool ParseOperations(std::wstring_view content, std::vector<Operation>* output, const std::atomic_bool* cancel, bool* cancelled, std::wstring* error)
{
    if (!output)
    {
        return false;
    }
    output->clear();
    if (cancelled)
    {
        *cancelled = false;
    }
    auto stopped = [&] {
        const bool value = cancel && cancel->load();
        if (value && cancelled)
        {
            *cancelled = true;
        }
        return value;
    };
    auto fail = [&](const std::wstring& line) {
        if (error)
        {
            std::wstring shown = line.size() > 80 ? line.substr(0, 80) + L"..." : line;
            *error = util::TrDetail(L"The file contains an entry RegKit can't parse.", shown);
        }
        return false;
    };

    bool ansi = false;
    bool in_key = false;
    std::wstring current_path;
    auto handle = [&](const std::wstring& raw) {
        const std::wstring line = util::TrimWhitespace(raw);
        if (line.empty() || line.front() == L';' || util::StartsWithInsensitive(line, L"Windows Registry Editor"))
        {
            return true;
        }
        if (util::StartsWithInsensitive(line, L"REGEDIT4"))
        {
            ansi = true;
            return true;
        }
        if (line.front() == L'[' && line.back() == L']')
        {
            std::wstring path = util::TrimWhitespace(std::wstring_view(line).substr(1, line.size() - 2));
            const bool removed = !path.empty() && path.front() == L'-';
            if (removed)
            {
                path = util::TrimWhitespace(std::wstring_view(path).substr(1));
            }
            if (path.empty())
            {
                return fail(line);
            }
            in_key = !removed;
            current_path = path;
            output->push_back({removed ? Operation::Kind::kRemoveKey : Operation::Kind::kKey, std::move(path)});
            return true;
        }
        if (!in_key)
        {
            return fail(line);
        }
        const size_t equals = FindAssignment(line);
        if (equals == std::wstring::npos)
        {
            return fail(line);
        }
        const std::wstring name_text = util::TrimWhitespace(std::wstring_view(line).substr(0, equals));
        const std::wstring data_text = util::TrimWhitespace(std::wstring_view(line).substr(equals + 1));
        if (name_text.empty() || data_text.empty())
        {
            return fail(line);
        }
        Operation operation{Operation::Kind::kValue, current_path};
        if (name_text != L"@" && !ParseQuoted(name_text, &operation.value.name))
        {
            return fail(line);
        }
        if (data_text == L"-")
        {
            operation.kind = Operation::Kind::kRemoveValue;
        }
        else if (!ParseValueData(data_text, ansi, &operation.value))
        {
            return fail(line);
        }
        output->push_back(std::move(operation));
        return true;
    };

    std::wstring logical;
    bool continuing = false;
    bool quoted = false;
    bool escaped = false;
    for (size_t start = 0; start < content.size();)
    {
        if (stopped())
        {
            return false;
        }
        size_t end = content.find(L'\n', start);
        const bool has_newline = end != std::wstring_view::npos;
        if (!has_newline)
        {
            end = content.size();
        }
        std::wstring_view physical = content.substr(start, end - start);
        const bool crlf = !physical.empty() && physical.back() == L'\r';
        if (crlf)
        {
            physical.remove_suffix(1);
        }
        start = end + 1;
        if (continuing)
        {
            const size_t first = physical.find_first_not_of(L" \t");
            physical.remove_prefix(first == std::wstring_view::npos ? physical.size() : first);
        }
        const size_t scanned = logical.size();
        logical.append(physical);
        const size_t first = logical.find_first_not_of(L" \t");
        if (quoted || (first != std::wstring::npos && logical[first] != L'[' && logical[first] != L';'))
        {
            for (size_t index = scanned; index < logical.size(); ++index)
            {
                if (escaped)
                {
                    escaped = false;
                }
                else if (quoted && logical[index] == L'\\')
                {
                    escaped = true;
                }
                else if (logical[index] == L'"')
                {
                    quoted = !quoted;
                }
            }
        }
        // RegEdit writes string LF as CRLF, other exporters wrap long strings with a backslash and two spaces
        if (quoted && has_newline)
        {
            if (!logical.empty() && logical.back() == L'\\' && content.substr(start, 2) == L"  ")
            {
                logical.pop_back();
                escaped = !escaped;
                start += 2;
                continuing = false;
                continue;
            }
            logical.push_back(L'\n');
            continuing = false;
            continue;
        }
        while (!logical.empty() && (logical.back() == L' ' || logical.back() == L'\t'))
        {
            logical.pop_back();
        }
        continuing = !quoted && !logical.empty() && logical.back() == L'\\';
        if (continuing)
        {
            logical.pop_back();
            continue;
        }
        if (!handle(logical))
        {
            return false;
        }
        logical.clear();
        quoted = false;
        escaped = false;
    }
    return logical.empty() || handle(logical);
}

bool Parse(std::wstring_view content, Document* output, const std::atomic_bool* cancel, bool* cancelled, std::wstring* error)
{
    if (!output)
    {
        return false;
    }
    output->keys.clear();
    output->key_order.clear();
    std::vector<Operation> operations;
    if (!ParseOperations(content, &operations, cancel, cancelled, error))
    {
        return false;
    }
    Key* current = nullptr;
    for (Operation& operation : operations)
    {
        if (operation.kind == Operation::Kind::kKey || operation.kind == Operation::Kind::kRemoveKey)
        {
            const std::wstring lower = util::ToLower(operation.path);
            auto [iterator, inserted] = output->keys.try_emplace(lower, Key{operation.path, {}});
            if (inserted)
            {
                output->key_order.push_back(std::move(operation.path));
            }
            current = &iterator->second;
            if (operation.kind == Operation::Kind::kKey)
            {
                current->recreated = current->removed;
                continue;
            }
            // deletions drop all content written to that key and descendants so far
            current->values.clear();
            current->removed_values.clear();
            current->removed = true;
            current->recreated = false;
            const std::wstring prefix = lower + L"\\";
            if (std::erase_if(output->keys, [&](const auto& entry) { return entry.first.starts_with(prefix); }))
            {
                std::erase_if(output->key_order, [&](const std::wstring& path) { return !output->keys.contains(util::ToLower(path)); });
            }
        }
        else if (operation.kind == Operation::Kind::kRemoveValue)
        {
            current->values.erase(util::ToLower(operation.value.name));
            current->removed_values.push_back(std::move(operation.value.name));
        }
        else
        {
            current->values[util::ToLower(operation.value.name)] = std::move(operation.value);
        }
    }
    return true;
}

bool ReadText(const std::wstring& path, std::wstring* content, std::wstring* error)
{
    // parsing peaks at 3 to 8 times file size, 32 bit processes run out above about 600 MiB
    constexpr uint64_t kMaxBytes = sizeof(void*) == 8 ? 1024ull << 20 : 256ull << 20;
    const DWORD code = win32ui::ReadTextFile(path.c_str(), content, kMaxBytes);
    if (code == ERROR_SUCCESS && !content->empty())
    {
        return true;
    }
    if (error)
    {
        *error = code == ERROR_FILE_TOO_LARGE ? util::TrDetail(L"The file is larger than RegKit can load.", path)
                 : code == ERROR_SUCCESS      ? util::TrDetail(L"The file has no registry changes.", path)
                                              : util::TrDetail(L"Failed to read registry file.", win32ui::FileReadErrorText(code) + L"\r\n" + path);
    }
    return false;
}

bool Load(const std::wstring& path, Document* output, std::wstring* error, const std::atomic_bool* cancel, bool* cancelled)
{
    std::wstring content;
    return ReadText(path, &content, error) && Parse(content, output, cancel, cancelled, error);
}

std::wstring RenderReg(const std::vector<Operation>& operations)
{
    Writer writer;
    const std::wstring* current = nullptr;
    for (const Operation& operation : operations)
    {
        if (operation.kind == Operation::Kind::kRemoveKey)
        {
            writer.AppendRemovedKey(operation.path);
            current = nullptr;
            continue;
        }
        if (!current || *current != operation.path)
        {
            writer.AppendKey(operation.path, {});
            current = &operation.path;
        }
        if (operation.kind == Operation::Kind::kValue)
        {
            writer.AppendValue(operation.value);
        }
        else if (operation.kind == Operation::Kind::kRemoveValue)
        {
            writer.AppendRemovedValue(operation.value.name);
        }
    }
    return std::move(writer).Finish();
}

bool LoadVirtualRoots(const std::wstring& path, std::vector<VirtualRoot>* roots, std::wstring* error, const std::atomic_bool* cancel, bool* cancelled)
{
    roots->clear();
    Document document;
    if (!Load(path, &document, error, cancel, cancelled))
    {
        return false;
    }
    std::unordered_map<std::wstring, VirtualRegistryData*> by_root;
    for (const auto& source_path : document.key_order)
    {
        if (cancel && cancel->load())
        {
            if (cancelled)
            {
                *cancelled = true;
            }
            return false;
        }
        const std::wstring normalized = registry_path::Normalize(source_path);
        const std::wstring key_path = normalized.empty() ? source_path : normalized;
        const size_t slash = key_path.find(L'\\');
        const std::wstring root_name = key_path.substr(0, slash);
        const auto source = document.keys.find(util::ToLower(source_path));
        if (root_name.empty() || source == document.keys.end())
        {
            continue;
        }
        VirtualRegistryData*& data = by_root[util::ToLower(root_name)];
        if (!data)
        {
            VirtualRoot& root = roots->emplace_back(VirtualRoot{root_name, std::make_shared<VirtualRegistryData>()});
            root.data->root_name = root_name;
            root.data->root = std::make_unique<VirtualRegistryKey>();
            root.data->root->name = root_name;
            data = root.data.get();
        }
        VirtualRegistryKey* target = data->root.get();
        for (const std::wstring& part : registry_path::Split(slash == std::wstring::npos ? L"" : key_path.substr(slash + 1)))
        {
            auto& child = target->children[util::ToLower(part)];
            if (!child)
            {
                child = std::make_unique<VirtualRegistryKey>();
                child->name = part;
            }
            target = child.get();
        }
        const Key& key = source->second;
        target->values.insert(key.values.begin(), key.values.end());
        target->removed_values.insert(target->removed_values.end(), key.removed_values.begin(), key.removed_values.end());
        target->removed |= key.removed;
        target->recreated |= key.recreated;
    }
    return true;
}

}
