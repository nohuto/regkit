// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "cli/reg_command.h"

#include "regfile/reg_exe_syntax.h"
#include "regfile/reg_file.h"
#include "regfile/script_convert.h"
#include "regfile/registry_transfer.h"
#include "registry/key_algorithms.h"
#include "registry/registry_path.h"
#include "registry/value_format.h"
#include "win32/file_text.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/registry_view.h"
#include "win32/shell_paths.h"
#include "win32/text_transform.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/console.h>
#include <win32ui/system/error_text.h>

#include <algorithm>
#include <cstdio>
#include <memory>
#include <numeric>
#include <optional>
#include <utility>

namespace regkit::cli
{

namespace
{

constexpr int kOk = 0;
constexpr int kFailed = 1;

void Print(const std::wstring& text)
{
    win32ui::WriteConsoleText(win32ui::ConsoleStream::output, text + L"\r\n");
}

void PrintError(const std::wstring& text)
{
    const bool prefixed = text.rfind(L"ERROR: ", 0) == 0;
    win32ui::WriteConsoleText(win32ui::ConsoleStream::error, (prefixed ? text : L"ERROR: " + text) + L"\r\n");
}

void PrintSuccess()
{
    Print(L"The operation completed successfully.");
    Print(L"");
}

int Fail(LONG status)
{
    PrintError(status == ERROR_FILE_NOT_FOUND ? L"The system was unable to find the specified registry key or value." : win32ui::FormatWin32Error(static_cast<DWORD>(status)));
    return kFailed;
}

struct KeyRef
{
    HKEY root = nullptr;
    std::wstring subkey;
    std::wstring path;
    std::wstring display;
    std::shared_ptr<win32ui::UniqueHKey> connection;
    // root becomes the connection on a remote key
    HKEY hive = nullptr;
    std::wstring machine;
};

// [\\machine\]ROOT\subkey, remote roots are HKLM and HKU, save, restore, load, unload and flags take local keys only
bool ParseKey(const std::wstring& text, KeyRef* key, bool local_only = false)
{
    std::wstring_view view = text;
    std::wstring machine;
    if (view.starts_with(L"\\\\"))
    {
        view.remove_prefix(2);
        const size_t end = view.find(L'\\');
        machine = std::wstring(view.substr(0, end));
        view = end == std::wstring_view::npos ? std::wstring_view() : view.substr(end + 1);
    }
    while (!view.empty() && view.front() == L'\\')
    {
        view.remove_prefix(1);
    }
    const size_t split = view.find(L'\\');
    const std::wstring_view root_name = split == std::wstring_view::npos ? view : view.substr(0, split);
    const HKEY root = registry_path::RootFromName(root_name);
    if (!root || (local_only && !machine.empty() && machine != L"."))
    {
        PrintError(L"Invalid key name: " + text);
        return false;
    }
    key->root = root;
    key->hive = root;
    wchar_t local[MAX_COMPUTERNAME_LENGTH + 1] = {};
    DWORD local_size = static_cast<DWORD>(_countof(local));
    key->machine = !machine.empty() && machine != L"." && !util::EqualsInsensitive(machine, L"localhost") ? machine : GetComputerNameW(local, &local_size) ? std::wstring(local, local_size)
                                                                                                                                                           : std::wstring();
    key->subkey = split == std::wstring_view::npos ? std::wstring() : std::wstring(view.substr(split + 1));
    key->path = registry_path::RootName(root);
    if (!key->subkey.empty())
    {
        key->path += L'\\';
        key->path += key->subkey;
    }
    key->display = key->path;
    key->connection.reset();
    if (!machine.empty() && machine != L".")
    {
        if (root != HKEY_LOCAL_MACHINE && root != HKEY_USERS)
        {
            PrintError(L"A remote machine was specified, the root key must be HKLM or HKU.");
            return false;
        }
        key->connection = std::make_shared<win32ui::UniqueHKey>();
        const LONG status = RegConnectRegistryW((L"\\\\" + machine).c_str(), root, key->connection->put());
        if (status != ERROR_SUCCESS)
        {
            Fail(status);
            return false;
        }
        key->root = key->connection->get();
        key->display = L"\\\\" + machine + L"\\" + key->path;
    }
    return true;
}

using reg_exe::IsSwitch;
using reg_exe::Options;

// reg.exe displays unknown types as REG_NONE
std::wstring TypeName(DWORD type)
{
    return type > REG_QWORD || type == REG_RESOURCE_REQUIREMENTS_LIST ? L"REG_NONE" : reg_exe::TypeName(type);
}

bool ParseOptions(const std::vector<std::wstring>& args, size_t first, Options* options, std::vector<std::wstring>* positional, std::initializer_list<std::wstring_view> allowed, reg_exe::Verb verb = reg_exe::Verb::kOther)
{
    std::wstring error;
    if (!reg_exe::ParseOptions(args, first, options, positional, verb, &error))
    {
        PrintError(error);
        return false;
    }
    for (const std::wstring& name : options->switches)
    {
        if (std::find(allowed.begin(), allowed.end(), name) == allowed.end())
        {
            PrintError(L"Invalid option: /" + name);
            return false;
        }
    }
    return true;
}

std::wstring FormatData(DWORD type, const BYTE* data, DWORD size, const std::wstring& separator = L"\\0")
{
    switch (type)
    {
    case REG_SZ:
    case REG_EXPAND_SZ:
        {
            std::wstring text(reinterpret_cast<const wchar_t*>(data), size / sizeof(wchar_t));
            while (!text.empty() && text.back() == L'\0')
            {
                text.pop_back();
            }
            return text;
        }
    case REG_MULTI_SZ:
        {
            std::wstring joined;
            for (const auto& item : value_format::MultiStringItems({data, size}))
            {
                if (!joined.empty())
                {
                    joined += separator;
                }
                joined += item;
            }
            return joined;
        }
    case REG_DWORD:
    case REG_DWORD_BIG_ENDIAN:
    case REG_QWORD:
        {
            wchar_t buffer[32] = {};
            // reg.exe displays REG_DWORD_BIG_ENDIAN as little endian, like REG_DWORD
            swprintf_s(buffer, L"0x%llx", value_format::ReadUnsigned({data, size}, type == REG_QWORD ? 8 : 4));
            return buffer;
        }
    default:
        return win32ui::HexBytes({data, size}, L'\0', true);
    }
}

KeyRef ChildRef(const KeyRef& parent, const std::wstring& name)
{
    KeyRef child = parent;
    child.subkey = registry_path::JoinSubkey(parent.subkey, name);
    child.path = parent.path + L"\\" + name;
    child.display = parent.display + L"\\" + name;
    return child;
}

std::wstring ValueName(const RegistryValue& value)
{
    return value.name.empty() ? std::wstring(L"(Default)") : value.name;
}

using registry_backend::KeyContents;

LONG ReadKey(const KeyRef& key, REGSAM view, bool include_data, KeyContents* contents)
{
    return registry_backend::ReadKeyContents(key.root, key.subkey, view, include_data, contents);
}

void SelectValue(const Options& options, std::vector<RegistryValue>* values)
{
    if (options.has_value)
    {
        const std::wstring wanted = options.default_value ? std::wstring() : options.value_name;
        std::erase_if(*values, [&](const RegistryValue& value) { return !util::EqualsInsensitive(value.name, wanted); });
    }
}

// * and ? wildcards match reg query /f, case-insensitive callers fold both sides first
bool WildcardMatch(std::wstring_view text, std::wstring_view pattern)
{
    size_t t = 0;
    size_t p = 0;
    size_t star = std::wstring_view::npos;
    size_t resume = 0;
    while (t < text.size())
    {
        if (p < pattern.size() && (pattern[p] == L'?' || pattern[p] == text[t]))
        {
            ++t;
            ++p;
        }
        else if (p < pattern.size() && pattern[p] == L'*')
        {
            star = p++;
            resume = t;
        }
        else if (star != std::wstring_view::npos)
        {
            p = star + 1;
            t = ++resume;
        }
        else
        {
            return false;
        }
    }
    while (p < pattern.size() && pattern[p] == L'*')
    {
        ++p;
    }
    return p == pattern.size();
}

struct Query
{
    const Options& options;
    // "*find*", folded unless case sensitive
    std::wstring pattern = L"*" + (options.case_sensitive ? options.find : util::ToLower(options.find)) + L"*";
    std::vector<DWORD> types;
    bool search_keys = false;
    bool search_names = false;
    bool search_data = false;
    size_t matches = 0;
    bool found_value = false;
    bool search = false;
    bool started = false;
    bool skipped = false;

    bool Matches(const std::wstring& text) const
    {
        if (options.exact)
        {
            return options.case_sensitive ? text == options.find : util::EqualsInsensitive(text, options.find);
        }
        // folds beyond ASCII like the /e path
        return WildcardMatch(options.case_sensitive ? text : util::ToLower(text), pattern);
    }
    bool Matches(const RegistryValue& value) const
    {
        if (search_names && Matches(value.name))
        {
            return true;
        }
        if (!search_data)
        {
            return false;
        }
        // numbers are searched in decimal, as reg.exe does
        if ((value.type == REG_DWORD || value.type == REG_QWORD) && value.data.size() >= (value.type == REG_DWORD ? 4u : 8u))
        {
            unsigned long long number = 0;
            memcpy(&number, value.data.data(), value.type == REG_DWORD ? 4 : 8);
            return Matches(std::to_wstring(number));
        }
        return Matches(FormatData(value.type, value.data.data(), static_cast<DWORD>(value.data.size()), options.separator));
    }
};

void PrintHeader(const KeyRef& key)
{
    Print(L"");
    Print(key.display);
}

LONG QueryKey(const KeyRef& key, Query& query, bool name_matched)
{
    const Options& options = query.options;
    KeyContents contents;
    const LONG status = ReadKey(key, options.view, true, &contents);
    if (status != ERROR_SUCCESS)
    {
        return status;
    }
    const bool root = !std::exchange(query.started, true);
    if (root)
    {
        Print(L"");
    }
    SelectValue(options, &contents.values);
    std::erase_if(contents.values, [&](const RegistryValue& value) {
        return (!query.types.empty() && std::find(query.types.begin(), query.types.end(), value.type) == query.types.end()) ||
               (options.has_find && !query.Matches(value));
    });
    const bool unset_default = options.default_value && !query.search && contents.values.empty();
    if (unset_default)
    {
        contents.values.push_back({L"", REG_SZ});
    }
    const bool listing = !query.search && !options.has_value;
    // reg.exe lists the queried key itself only when it has values
    if ((listing && !root) || name_matched || !contents.values.empty())
    {
        Print(key.display);
        for (const RegistryValue& value : contents.values)
        {
            const std::wstring type = TypeName(value.type) + (options.verbose ? L" (" + std::to_wstring(static_cast<LONG>(value.type)) + L")" : std::wstring());
            Print(L"    " + ValueName(value) + L"    " + type + L"    " + (unset_default ? L"(value not set)" : FormatData(value.type, value.data.data(), static_cast<DWORD>(value.data.size()), options.separator)));
        }
        Print(L"");
    }
    query.found_value = query.found_value || !contents.values.empty();
    query.matches += contents.values.size() + (name_matched ? 1 : 0);
    if (listing && !options.recurse)
    {
        for (const auto& child : contents.subkeys)
        {
            Print(ChildRef(key, child).display);
        }
        return ERROR_SUCCESS;
    }
    for (const auto& child : contents.subkeys)
    {
        const bool child_matched = options.has_find && query.search_keys && query.Matches(child);
        const KeyRef child_key = ChildRef(key, child);
        if (!options.recurse)
        {
            if (child_matched)
            {
                Print(child_key.display);
                ++query.matches;
            }
            continue;
        }
        // reg.exe silently skips unreadable subkeys, report skips and fail the exit code here
        const LONG child_status = QueryKey(child_key, query, child_matched);
        if (child_status != ERROR_SUCCESS)
        {
            if (child_matched)
            {
                Print(child_key.display);
                Print(L"");
                ++query.matches;
            }
            PrintError(L"Skipped " + regfile::Describe({regfile::Operation::Kind::kKey, child_key.display}, win32ui::FormatWin32Error(static_cast<DWORD>(child_status))));
            query.skipped = true;
        }
    }
    return ERROR_SUCCESS;
}

int CmdQuery(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"v", L"ve", L"s", L"se", L"f", L"k", L"d", L"c", L"e", L"t", L"z", L"reg:32", L"reg:64"}, reg_exe::Verb::kQuery))
    {
        return kFailed;
    }
    if (positional.empty())
    {
        PrintError(L"reg query requires a key name.");
        return kFailed;
    }
    if (positional.size() > 1 || (options.value_names && !options.has_find) || ((options.keys_only || options.data_only || options.case_sensitive || options.exact) && !options.has_find))
    {
        PrintError(L"Invalid syntax.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key))
    {
        return kFailed;
    }
    Query query{options};
    for (size_t start = 0; !options.type_text.empty() && start <= options.type_text.size();)
    {
        const size_t end = std::min(options.type_text.find(L',', start), options.type_text.size());
        DWORD type = REG_NONE;
        if (!reg_exe::ParseType(std::wstring_view(options.type_text).substr(start, end - start), &type))
        {
            PrintError(L"Invalid type: " + options.type_text);
            return kFailed;
        }
        query.types.push_back(type);
        start = end + 1;
    }
    const bool scoped = options.keys_only || options.value_names || options.data_only;
    query.search_keys = !scoped || options.keys_only;
    query.search_names = !scoped || options.value_names;
    query.search_data = !scoped || options.data_only;
    query.search = options.has_find || !query.types.empty() || (options.has_value && options.recurse);
    const LONG status = QueryKey(key, query, false);
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    if (query.search)
    {
        Print(L"End of search: " + std::to_wstring(query.matches) + L" match(es) found.");
        return query.matches && !query.skipped ? kOk : kFailed;
    }
    // key may exist even if /v or /ve value is missing
    if (options.has_value && !query.found_value)
    {
        Print(L"");
        return Fail(ERROR_FILE_NOT_FOUND);
    }
    return query.skipped ? kFailed : kOk;
}

int CmdAdd(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"v", L"ve", L"t", L"s", L"d", L"f", L"reg:32", L"reg:64"}, reg_exe::Verb::kAdd))
    {
        return kFailed;
    }
    if (positional.empty())
    {
        PrintError(L"reg add requires a key name.");
        return kFailed;
    }
    if (positional.size() > 1)
    {
        PrintError(L"Invalid syntax.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key))
    {
        return kFailed;
    }

    // reg.exe writes /d or /t without /v to the default value
    if (!options.has_value && (options.has_data || !options.type_text.empty()))
    {
        options.has_value = options.default_value = true;
    }
    DWORD value_type = REG_SZ;
    std::vector<BYTE> value_data;
    if (options.has_value)
    {
        if (!options.type_text.empty() && !reg_exe::ParseType(options.type_text, &value_type))
        {
            PrintError(L"Invalid type: " + options.type_text);
            return kFailed;
        }
        std::wstring error;
        if (!reg_exe::BuildData(value_type, options.has_data ? options.data : std::wstring(), options.separator, &value_data, &error))
        {
            PrintError(error);
            return kFailed;
        }
    }

    win32ui::UniqueHKey handle;
    LONG status = RegCreateKeyExW(key.root, key.subkey.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE | KEY_QUERY_VALUE | options.view, nullptr, handle.put(), nullptr);
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    if (options.has_value)
    {
        const std::wstring name = options.default_value ? std::wstring() : options.value_name;
        if (!options.force &&
            RegQueryValueExW(handle.get(), name.c_str(), nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS)
        {
            PrintError(L"Value " + (name.empty() ? std::wstring(L"(Default)") : name) + L" already exists. Use /f to overwrite.");
            return kFailed;
        }
        status = RegSetValueExW(handle.get(), name.c_str(), 0, value_type, value_data.data(), static_cast<DWORD>(value_data.size()));
        if (status != ERROR_SUCCESS)
        {
            return Fail(status);
        }
    }
    PrintSuccess();
    return kOk;
}

int CmdDelete(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"v", L"ve", L"va", L"f", L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.empty())
    {
        PrintError(L"reg delete requires a key name.");
        return kFailed;
    }
    if (positional.size() > 1)
    {
        PrintError(L"Invalid syntax.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key))
    {
        return kFailed;
    }

    if (!options.force)
    {
        PrintError(L"This operation deletes registry data. Rerun with /f to confirm.");
        return kFailed;
    }

    if (options.has_value || options.all_values)
    {
        KeyContents contents;
        LONG status = options.all_values ? ReadKey(key, options.view, false, &contents) : ERROR_SUCCESS;
        win32ui::UniqueHKey handle;
        if (status == ERROR_SUCCESS)
        {
            status = RegOpenKeyExW(key.root, key.subkey.c_str(), 0, KEY_SET_VALUE | options.view, handle.put());
        }
        if (!options.all_values)
        {
            contents.values.push_back({options.default_value ? std::wstring() : options.value_name});
        }
        const registry_backend::RegistryKeyHandle target(std::move(handle));
        for (size_t index = 0; status == ERROR_SUCCESS && index < contents.values.size(); ++index)
        {
            status = target.DeleteValue(contents.values[index].name);
        }
        if (status != ERROR_SUCCESS)
        {
            return Fail(status);
        }
        PrintSuccess();
        return kOk;
    }

    if (key.subkey.empty())
    {
        PrintError(L"Refusing to delete a registry root.");
        return kFailed;
    }
    win32ui::UniqueHKey target;
    LONG status = util::OpenRegistryPath(
        key.root,
        key.subkey,
        DELETE | KEY_ENUMERATE_SUB_KEYS | KEY_QUERY_VALUE | options.view,
        true,
        &target
    );
    if (status == ERROR_SUCCESS)
    {
        status = util::DeleteRegistryTree(target.get());
    }
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    PrintSuccess();
    return kOk;
}

LONG CopyTree(const KeyRef& from, HKEY to_parent, const std::wstring& to_name, REGSAM view, bool recurse)
{
    KeyContents contents;
    LONG status = ReadKey(from, view, true, &contents);
    win32ui::UniqueHKey created;
    if (status == ERROR_SUCCESS)
    {
        status = util::CreateRegistryKey(to_parent, to_name, KEY_WRITE | view, REG_OPTION_NON_VOLATILE, &created, nullptr);
    }
    const registry_backend::RegistryKeyHandle target(std::move(created));
    for (size_t index = 0; status == ERROR_SUCCESS && index < contents.values.size(); ++index)
    {
        const RegistryValue& value = contents.values[index];
        status = target.SetValue(value.name, value.type, value.data.data(), static_cast<DWORD>(value.data.size()));
    }
    for (size_t index = 0; recurse && status == ERROR_SUCCESS && index < contents.subkeys.size(); ++index)
    {
        status = CopyTree(ChildRef(from, contents.subkeys[index]), target.get(), contents.subkeys[index], view, true);
    }
    return status;
}

int CmdCopy(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"s", L"f", L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.size() != 2)
    {
        PrintError(positional.size() < 2 ? L"reg copy requires a source and a destination key." : L"Invalid syntax.");
        return kFailed;
    }
    KeyRef from;
    KeyRef to;
    if (!ParseKey(positional[0], &from) || !ParseKey(positional[1], &to))
    {
        return kFailed;
    }
    // recursive copies into their own subtree would copy newly written keys forever
    if (from.hive == to.hive && util::EqualsInsensitive(from.machine, to.machine) && (util::EqualsInsensitive(from.subkey, to.subkey) || (options.recurse && (from.subkey.empty() || registry_path::HasComponentPrefix(to.subkey, from.subkey)))))
    {
        PrintError(L"The registry entry cannot be copied onto itself or into its own subkey.");
        return kFailed;
    }
    const LONG status = CopyTree(from, to.root, to.subkey, options.view, options.recurse);
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    PrintSuccess();
    return kOk;
}

// write readable subkeys even after skips, exit code reports incomplete output
int ExportKeyToFile(const KeyRef& key, const std::wstring& path, REGSAM view)
{
    regfile::Writer writer;
    std::vector<std::wstring> skipped;
    const LONG status = regfile::AppendRegistryTree(&writer, key.root, key.subkey, key.display, view, true, &skipped);
    for (const std::wstring& entry : skipped)
    {
        PrintError(L"Skipped " + entry);
    }
    if (status != ERROR_SUCCESS)
    {
        PrintError(status == ERROR_FILE_NOT_FOUND ? L"The key doesn't exist: " + key.display : win32ui::FormatWin32Error(static_cast<DWORD>(status)));
        return kFailed;
    }
    if (win32ui::WriteTextFileAtomic(path.c_str(), std::move(writer).Finish(), win32ui::TextEncoding::utf16) != ERROR_SUCCESS)
    {
        PrintError(L"Failed to write " + path);
        return kFailed;
    }
    return skipped.empty() ? kOk : kFailed;
}

int CmdExport(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"y", L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.size() != 2)
    {
        PrintError(positional.size() < 2 ? L"reg export requires a key name and a file name." : L"Invalid syntax.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key))
    {
        return kFailed;
    }
    if (!options.force && !util::IsMissing(positional[1].c_str()))
    {
        PrintError(positional[1] + L" already exists. Use /y to overwrite.");
        return kFailed;
    }
    const int result = ExportKeyToFile(key, positional[1], options.view);
    if (result == kOk)
    {
        PrintSuccess();
    }
    return result;
}

int CmdImport(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.size() != 1)
    {
        PrintError(positional.empty() ? L"reg import requires a file name." : L"Invalid syntax.");
        return kFailed;
    }
    std::wstring error;
    if (!ImportRegFileFromPath(positional[0], &error, options.view))
    {
        PrintError(error.empty() ? L"Import failed." : error);
        return kFailed;
    }
    PrintSuccess();
    return kOk;
}

int CmdSave(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"y", L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.size() < 2)
    {
        PrintError(L"reg save requires a key name and a file name.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key, true))
    {
        return kFailed;
    }
    if (!options.force && !util::IsMissing(positional[1].c_str()))
    {
        PrintError(positional[1] + L" already exists. Use /y to overwrite.");
        return kFailed;
    }
    const LONG status = SaveKeyToHive(key.root, key.subkey, options.view, positional[1]);
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    PrintSuccess();
    return kOk;
}

int CmdRestore(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.size() < 2)
    {
        PrintError(L"reg restore requires a key name and a file name.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key, true))
    {
        return kFailed;
    }
    const LONG status = RestoreKeyFromHive(key.root, key.subkey, options.view, positional[1]);
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    PrintSuccess();
    return kOk;
}

int CmdLoad(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.size() != 2)
    {
        PrintError(positional.size() < 2 ? L"reg load requires a key name and a file name." : L"Invalid syntax.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key, true))
    {
        return kFailed;
    }
    const util::PrivilegeScope privileges({SE_RESTORE_NAME, SE_BACKUP_NAME});
    const LONG status = RegLoadKeyW(key.root, key.subkey.c_str(), positional[1].c_str());
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    PrintSuccess();
    return kOk;
}

int CmdUnload(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {}))
    {
        return kFailed;
    }
    if (positional.size() != 1)
    {
        PrintError(positional.empty() ? L"reg unload requires a key name." : L"Invalid syntax.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key, true))
    {
        return kFailed;
    }
    const util::PrivilegeScope privileges({SE_RESTORE_NAME, SE_BACKUP_NAME});
    const LONG status = RegUnLoadKeyW(key.root, key.subkey.c_str());
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    PrintSuccess();
    return kOk;
}

struct Compare
{
    const Options& options;
    bool differs = false;

    // = equal, < left only or different, > right only or different, /oa /od /os /on select printed lines
    void Line(wchar_t mark, const KeyRef& key, const RegistryValue* value)
    {
        const bool equal = mark == L'=';
        differs = differs || !equal;
        const wchar_t output = options.compare_output;
        if (output == L'n' || (output == L'd' && equal) || (output == L's' && !equal))
        {
            return;
        }
        std::wstring line(1, mark);
        if (value)
        {
            line += L" Value: " + key.display + L"  " + ValueName(*value) + L" " + TypeName(value->type) + L" " + FormatData(value->type, value->data.data(), static_cast<DWORD>(value->data.size()));
        }
        else
        {
            line += L" Key: " + key.display;
        }
        Print(line);
    }
};

template <typename Item, typename Name>
class NameIndex
{
  public:
    NameIndex(const std::vector<Item>& items, Name name)
        : items_(items), name_(name), order_(items.size())
    {
        std::iota(order_.begin(), order_.end(), size_t{0});
        std::sort(order_.begin(), order_.end(), [&](size_t a, size_t b) { return util::CompareInsensitive(name_(items_[a]), name_(items_[b])) < 0; });
    }

    const Item* Find(const std::wstring& name) const
    {
        const auto found = std::lower_bound(order_.begin(), order_.end(), name, [&](size_t index, const std::wstring& key) {
            return util::CompareInsensitive(name_(items_[index]), key) < 0;
        });
        return found != order_.end() && util::CompareInsensitive(name_(items_[*found]), name) == 0 ? &items_[*found] : nullptr;
    }

  private:
    const std::vector<Item>& items_;
    Name name_;
    std::vector<size_t> order_;
};

int CompareKeys(const KeyRef& left, const KeyRef& right, Compare& compare, bool top)
{
    const Options& options = compare.options;
    KeyContents left_contents;
    KeyContents right_contents;
    const LONG left_status = ReadKey(left, options.view, true, &left_contents);
    const LONG right_status = ReadKey(right, options.view, true, &right_contents);
    for (const LONG status : {left_status, right_status})
    {
        if (status != ERROR_SUCCESS)
        {
            return Fail(status);
        }
    }
    SelectValue(options, &left_contents.values);
    SelectValue(options, &right_contents.values);
    if (top && options.has_value && (left_contents.values.empty() || right_contents.values.empty()))
    {
        return Fail(ERROR_FILE_NOT_FOUND);
    }
    const auto value_name = [](const RegistryValue& value) -> const std::wstring& { return value.name; };
    const NameIndex left_values(left_contents.values, value_name);
    const NameIndex right_values(right_contents.values, value_name);
    for (const RegistryValue& value : left_contents.values)
    {
        const RegistryValue* other = right_values.Find(value.name);
        if (other && other->type == value.type && other->data == value.data)
        {
            compare.Line(L'=', left, &value);
            continue;
        }
        compare.Line(L'<', left, &value);
        if (other)
        {
            compare.Line(L'>', right, other);
        }
    }
    for (const RegistryValue& value : right_contents.values)
    {
        if (!left_values.Find(value.name))
        {
            compare.Line(L'>', right, &value);
        }
    }
    if (!options.recurse)
    {
        return kOk;
    }
    const auto key_name = [](const std::wstring& name) -> const std::wstring& { return name; };
    const NameIndex left_keys(left_contents.subkeys, key_name);
    const NameIndex right_keys(right_contents.subkeys, key_name);
    // compare shared keys first, then one sided keys without descending into them
    for (const std::wstring& child : left_contents.subkeys)
    {
        if (right_keys.Find(child))
        {
            compare.Line(L'=', ChildRef(left, child), nullptr);
            if (CompareKeys(ChildRef(left, child), ChildRef(right, child), compare, false) == kFailed)
            {
                return kFailed;
            }
        }
    }
    for (const std::wstring& child : left_contents.subkeys)
    {
        if (!right_keys.Find(child))
        {
            compare.Line(L'<', ChildRef(left, child), nullptr);
        }
    }
    for (const std::wstring& child : right_contents.subkeys)
    {
        if (!left_keys.Find(child))
        {
            compare.Line(L'>', ChildRef(right, child), nullptr);
        }
    }
    return kOk;
}

int CmdCompare(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"v", L"ve", L"oa", L"od", L"os", L"on", L"s", L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    if (positional.size() != 2)
    {
        PrintError(positional.size() < 2 ? L"reg compare requires two key names." : L"Invalid syntax.");
        return kFailed;
    }
    KeyRef left;
    KeyRef right;
    // bare \\machine as second key uses the first key path on that machine
    std::wstring second = positional[1];
    if (second.starts_with(L"\\\\") && second.find(L'\\', 2) == std::wstring::npos && ParseKey(positional[0], &left))
    {
        second += L"\\" + left.path;
    }
    if (!ParseKey(positional[0], &left) || !ParseKey(second, &right))
    {
        return kFailed;
    }
    if (util::EqualsInsensitive(left.display, right.display))
    {
        PrintError(L"The registry entry is being compared with itself.");
        return kFailed;
    }
    Compare compare{options};
    if (CompareKeys(left, right, compare, true) == kFailed)
    {
        return kFailed;
    }
    Print(L"");
    Print(compare.differs ? L"Result Compared:  Different" : L"Result Compared:  Identical");
    PrintSuccess();
    // reg.exe exit code 2 means different keys
    return compare.differs ? 2 : kOk;
}

struct ControlFlag
{
    const wchar_t* name;
    ULONG bit;
};

constexpr ControlFlag kControlFlags[] = {
    {L"DONT_VIRTUALIZE", util::kKeyDontVirtualize},
    {L"DONT_SILENT_FAIL", util::kKeyDontSilentFail},
    {L"RECURSE_FLAG", util::kKeyRecurseFlag},
};

LONG ApplyFlags(const KeyRef& key, const Options& options, std::optional<ULONG> set)
{
    win32ui::UniqueHKey handle;
    LONG status = util::OpenRegistryPath(key.root, key.subkey, KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS | (set ? KEY_SET_VALUE : 0) | options.view, false, &handle);
    if (status != ERROR_SUCCESS)
    {
        return status;
    }
    if (set)
    {
        status = util::SetKeyControlFlags(handle.get(), *set);
    }
    else
    {
        const ULONG flags = util::QueryNativeKeyInfo(handle.get()).control_flags.value_or(0);
        PrintHeader(key);
        for (const ControlFlag& flag : kControlFlags)
        {
            Print(std::wstring(L"\tREG_KEY_") + flag.name + ((flags & flag.bit) ? L": SET" : L": CLEAR"));
        }
    }
    if (status != ERROR_SUCCESS || !options.recurse)
    {
        return status;
    }
    KeyContents contents;
    status = ReadKey(key, options.view, false, &contents);
    for (size_t index = 0; status == ERROR_SUCCESS && index < contents.subkeys.size(); ++index)
    {
        status = ApplyFlags(ChildRef(key, contents.subkeys[index]), options, set);
    }
    return status;
}

int CmdFlags(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"s", L"reg:32", L"reg:64"}))
    {
        return kFailed;
    }
    const bool set_verb = positional.size() > 1 && util::EqualsInsensitive(positional[1], L"SET");
    if (positional.empty() || (!set_verb && positional.size() > 1 && (positional.size() > 2 || !util::EqualsInsensitive(positional[1], L"QUERY"))))
    {
        PrintError(L"Invalid syntax.");
        return kFailed;
    }
    KeyRef key;
    if (!ParseKey(positional[0], &key, true))
    {
        return kFailed;
    }
    const std::wstring_view first = std::wstring_view(key.subkey).substr(0, key.subkey.find(L'\\'));
    if (key.connection || key.root != HKEY_LOCAL_MACHINE || !util::EqualsInsensitive(first, L"Software"))
    {
        PrintError(L"This operation can only be performed on subkeys of HKLM\\Software.");
        return kFailed;
    }
    std::optional<ULONG> set;
    if (set_verb)
    {
        set = 0;
        for (size_t index = 2; index < positional.size(); ++index)
        {
            const auto flag = std::find_if(std::begin(kControlFlags), std::end(kControlFlags), [&](const ControlFlag& entry) { return util::EqualsInsensitive(positional[index], entry.name); });
            if (flag == std::end(kControlFlags))
            {
                PrintError(L"Invalid syntax.");
                return kFailed;
            }
            *set |= flag->bit;
        }
    }
    const LONG status = ApplyFlags(key, options, set);
    if (status != ERROR_SUCCESS)
    {
        return Fail(status);
    }
    if (!set)
    {
        Print(L"");
    }
    PrintSuccess();
    return kOk;
}

int CmdConvert(const std::vector<std::wstring>& args)
{
    Options options;
    std::vector<std::wstring> positional;
    if (!ParseOptions(args, 1, &options, &positional, {L"y"}))
    {
        return kFailed;
    }
    regfile::Format format = regfile::Format::kReg;
    if (positional.size() != 2 || !regfile::FormatFromPath(positional[1], &format))
    {
        PrintError(L"Usage: convert <input> <output> [/y], both .reg, .bat, .cmd or .ps1.");
        return kFailed;
    }
    if (!options.force && !util::IsMissing(positional[1].c_str()))
    {
        PrintError(positional[1] + L" already exists. Use /y to overwrite.");
        return kFailed;
    }
    std::vector<regfile::Operation> operations;
    std::wstring error;
    if (!regfile::ReadOperations(positional[0], &operations, &error))
    {
        PrintError(error);
        return kFailed;
    }
    std::vector<std::wstring> skipped;
    const std::wstring text = regfile::RenderOperations(format, operations, true, &skipped);
    for (const std::wstring& entry : skipped)
    {
        PrintError(L"Skipped " + entry);
    }
    if (!regfile::SaveRendered(positional[1], format, text))
    {
        PrintError(L"Failed to write " + positional[1]);
        return kFailed;
    }
    if (!skipped.empty())
    {
        return kFailed;
    }
    PrintSuccess();
    return kOk;
}

void PrintUsage()
{
    Print(L"RegKit command line\n"
          L"\n"
          L"regedit compatible:\n"
          L"  regkit file.reg                 import a .reg file (asks first)\n"
          L"  regkit /s file.reg              import without prompting\n"
          L"  regkit /e file.reg <key>        export a key\n"
          L"  regkit /a file.reg <key>        same as /e, kept for compatibility\n"
          L"  regkit /c /m /l:file /r:file    accepted and ignored (legacy)\n"
          L"\n"
          L"reg.exe compatible (the leading \"reg\" is optional):\n"
          L"  add <key> [/v name | /ve] [/t type] [/s sep] [/d data] [/f]\n"
          L"  delete <key> [/v name | /ve | /va] [/f]\n"
          L"  query <key> [/v [name] | /ve] [/s] [/f data [/k] [/d] [/c] [/e]]\n"
          L"        [/t type[,type]] [/z] [/se separator]\n"
          L"  copy <src> <dst> [/s] [/f]\n"
          L"  export <key> <file.reg> [/y]\n"
          L"  import <file.reg>\n"
          L"  save <key> <file.hiv> [/y]      restore <key> <file.hiv>\n"
          L"  load <key> <file.hiv>           unload <key>\n"
          L"  compare <key1> <key2> [/v name | /ve] [/oa | /od | /os | /on] [/s]\n"
          L"  flags <HKLM\\Software\\key> [QUERY | SET [DONT_VIRTUALIZE]\n"
          L"        [DONT_SILENT_FAIL] [RECURSE_FLAG]] [/s]\n"
          L"  /reg:32 | /reg:64               pick the registry view\n"
          L"\n"
          L"RegKit additions:\n"
          L"  regkit <key>                    open the window at that key\n"
          L"  regkit --goto <key>             same, explicit form\n"
          L"  regkit --edit-reg file.reg      open a .reg file in a tab\n"
          L"  regkit convert <in> <out> [/y]  convert between .reg, .bat, .cmd and .ps1\n"
          L"  regkit --install-edit-context-menu\n"
          L"                                    add the Edit with RegKit context menu\n"
          L"  regkit --uninstall-edit-context-menu\n"
          L"                                    remove this executable's context menu\n"
          L"  regkit --install-regedit-replacement [--override]\n"
          L"                                    replace RegEdit with this executable\n"
          L"                                    fails if another program owns RegEdits\n"
          L"                                    Debugger entry, --override replaces it anyway\n"
          L"  regkit --uninstall-regedit-replacement\n"
          L"                                    remove this executable's replacement\n"
          L"  regkit --setting key=value      change a settings.ini entry (repeatable)\n"
          L"  regkit --restart-system         relaunch as SYSTEM\n"
          L"  regkit --restart-ti             relaunch as TrustedInstaller\n"
          L"  regkit --help                   show this text\n"
          L"\n"
          L"Key names accept HKLM, HKCU, HKCR, HKU, HKCC and their full forms,\n"
          L"and \\\\machine\\HKLM or \\\\machine\\HKU for a remote registry.");
}

int RunVerb(const std::wstring& verb, const std::vector<std::wstring>& args)
{
    struct Verb
    {
        const wchar_t* name;
        int (*run)(const std::vector<std::wstring>&);
    };
    static constexpr Verb kVerbs[] = {
        {L"query", CmdQuery},
        {L"add", CmdAdd},
        {L"delete", CmdDelete},
        {L"copy", CmdCopy},
        {L"export", CmdExport},
        {L"import", CmdImport},
        {L"save", CmdSave},
        {L"restore", CmdRestore},
        {L"load", CmdLoad},
        {L"unload", CmdUnload},
        {L"compare", CmdCompare},
        {L"flags", CmdFlags},
        {L"convert", CmdConvert},
    };
    for (const Verb& entry : kVerbs)
    {
        if (util::EqualsInsensitive(verb, entry.name))
        {
            return entry.run(args);
        }
    }
    return -1;
}
}

bool Execute(const std::vector<std::wstring>& args, int* exit_code)
{
    if (args.empty() || !exit_code)
    {
        return false;
    }

    for (const std::wstring& arg : args)
    {
        if (IsSwitch(arg, L"?") || IsSwitch(arg, L"h") || IsSwitch(arg, L"-help") || IsSwitch(arg, L"help"))
        {
            PrintUsage();
            *exit_code = kOk;
            return true;
        }
    }

    std::vector<std::wstring> verb_args = args;
    if (util::EqualsInsensitive(verb_args[0], L"reg"))
    {
        verb_args.erase(verb_args.begin());
    }
    if (!verb_args.empty())
    {
        const int result = RunVerb(verb_args[0], verb_args);
        if (result >= 0)
        {
            *exit_code = result;
            return true;
        }
    }

    // accept legacy RegEdit import/export syntax
    for (size_t i = 0; i < args.size(); ++i)
    {
        if (IsSwitch(args[i], L"s") && i + 1 < args.size())
        {
            std::wstring error;
            if (!ImportRegFileFromPath(args[i + 1], &error))
            {
                PrintError(error.empty() ? L"Import failed." : error);
                *exit_code = kFailed;
            }
            else
            {
                *exit_code = kOk;
            }
            return true;
        }
        if ((IsSwitch(args[i], L"e") || IsSwitch(args[i], L"a")) && i + 1 < args.size())
        {
            KeyRef key;
            *exit_code = kFailed;
            if (i + 2 >= args.size())
            {
                PrintError(L"Exporting the whole registry isn't supported, name a key.");
            }
            else if (ParseKey(args[i + 2], &key))
            {
                *exit_code = ExportKeyToFile(key, args[i + 1], win32::kDefaultRegistryView);
            }
            return true;
        }
    }
    return false;
}

bool PrintErrorToTerminal(const std::wstring& text)
{
    if (!win32ui::AttachTerminal())
    {
        return false;
    }
    PrintError(text);
    return true;
}

}
