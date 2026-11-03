// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "registry/registry_backends.h"

#include "registry/key_algorithms.h"
#include "win32/shell_paths.h"
#include "win32/translation.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/error_text.h>
#include <win32ui/system_api.h>

#include <algorithm>
#include <iterator>
#include <mutex>
#include <shared_mutex>
#include <vector>

#include <winternl.h>

namespace regkit::registry_backend::offline
{
namespace
{

using ORHKEY = void*;
using OROpenHiveFn = DWORD(WINAPI*)(PCWSTR, ORHKEY*);
using ORCloseHiveFn = DWORD(WINAPI*)(ORHKEY);
using ORSaveHiveFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, DWORD, DWORD);
using OROpenKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, ORHKEY*);
using ORCloseKeyFn = DWORD(WINAPI*)(ORHKEY);
using ORCreateKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, PWSTR, DWORD, PSECURITY_DESCRIPTOR, ORHKEY*, DWORD*);
using ORDeleteKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);
using ORQueryInfoKeyFn = DWORD(WINAPI*)(ORHKEY, PWSTR, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, DWORD*, FILETIME*);
using OREnumKeyFn = DWORD(WINAPI*)(ORHKEY, DWORD, PWSTR, DWORD*, PWSTR, DWORD*, FILETIME*);
using ORGetValueFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, PCWSTR, DWORD*, void*, DWORD*);
using ORSetValueFn = DWORD(WINAPI*)(ORHKEY, PCWSTR, DWORD, const BYTE*, DWORD);
using ORDeleteValueFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);
using OREnumValueFn = DWORD(WINAPI*)(ORHKEY, DWORD, PWSTR, DWORD*, DWORD*, BYTE*, DWORD*);
using ORRenameKeyFn = DWORD(WINAPI*)(ORHKEY, PCWSTR);
using ORGetKeySecurityFn = DWORD(WINAPI*)(ORHKEY, SECURITY_INFORMATION, PSECURITY_DESCRIPTOR, DWORD*);
using ORSetKeySecurityFn = DWORD(WINAPI*)(ORHKEY, SECURITY_INFORMATION, PSECURITY_DESCRIPTOR);

class OffregApi
{
  public:
    OffregApi()
    {
        const std::wstring directory = util::ModuleDirectory();
        const std::wstring path = directory.empty() ? std::wstring() : util::JoinPath(directory, L"offreg.dll");
        const DWORD attributes = path.empty() ? INVALID_FILE_ATTRIBUTES : GetFileAttributesW(path.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
        {
            load_error_ = path.empty()                            ? ERROR_MOD_NOT_FOUND
                          : attributes == INVALID_FILE_ATTRIBUTES ? GetLastError()
                                                                  : ERROR_ACCESS_DENIED;
            return;
        }
        module_ =
            LoadLibraryExW(path.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32);
        if (!module_ && GetLastError() == ERROR_INVALID_PARAMETER)
        {
            module_ = LoadLibraryExW(path.c_str(), nullptr, LOAD_WITH_ALTERED_SEARCH_PATH);
        }
        if (!module_)
        {
            load_error_ = GetLastError();
            return;
        }
        open_hive = win32ui::ImportProc<OROpenHiveFn>(module_, "OROpenHive");
        close_hive = win32ui::ImportProc<ORCloseHiveFn>(module_, "ORCloseHive");
        save_hive = win32ui::ImportProc<ORSaveHiveFn>(module_, "ORSaveHive");
        open_key = win32ui::ImportProc<OROpenKeyFn>(module_, "OROpenKey");
        close_key = win32ui::ImportProc<ORCloseKeyFn>(module_, "ORCloseKey");
        create_key = win32ui::ImportProc<ORCreateKeyFn>(module_, "ORCreateKey");
        delete_key = win32ui::ImportProc<ORDeleteKeyFn>(module_, "ORDeleteKey");
        query_info = win32ui::ImportProc<ORQueryInfoKeyFn>(module_, "ORQueryInfoKey");
        enum_key = win32ui::ImportProc<OREnumKeyFn>(module_, "OREnumKey");
        get_value = win32ui::ImportProc<ORGetValueFn>(module_, "ORGetValue");
        set_value = win32ui::ImportProc<ORSetValueFn>(module_, "ORSetValue");
        delete_value = win32ui::ImportProc<ORDeleteValueFn>(module_, "ORDeleteValue");
        enum_value = win32ui::ImportProc<OREnumValueFn>(module_, "OREnumValue");
        rename_key = win32ui::ImportProc<ORRenameKeyFn>(module_, "ORRenameKey");
        get_key_security = win32ui::ImportProc<ORGetKeySecurityFn>(module_, "ORGetKeySecurity");
        set_key_security = win32ui::ImportProc<ORSetKeySecurityFn>(module_, "ORSetKeySecurity");
        if (!valid())
        {
            load_error_ = ERROR_PROC_NOT_FOUND;
            FreeLibrary(module_);
            module_ = nullptr;
        }
    }

    ~OffregApi()
    {
        if (module_)
        {
            FreeLibrary(module_);
        }
    }

    OffregApi(const OffregApi&) = delete;
    OffregApi& operator=(const OffregApi&) = delete;

    bool valid() const noexcept
    {
        return module_ && open_hive && close_hive && save_hive && open_key && close_key && create_key && delete_key &&
               query_info && enum_key && get_value && set_value && delete_value && enum_value && rename_key;
    }

    OROpenHiveFn open_hive = nullptr;
    ORCloseHiveFn close_hive = nullptr;
    ORSaveHiveFn save_hive = nullptr;
    OROpenKeyFn open_key = nullptr;
    ORCloseKeyFn close_key = nullptr;
    ORCreateKeyFn create_key = nullptr;
    ORDeleteKeyFn delete_key = nullptr;
    ORQueryInfoKeyFn query_info = nullptr;
    OREnumKeyFn enum_key = nullptr;
    ORGetValueFn get_value = nullptr;
    ORSetValueFn set_value = nullptr;
    ORDeleteValueFn delete_value = nullptr;
    OREnumValueFn enum_value = nullptr;
    ORRenameKeyFn rename_key = nullptr;
    ORGetKeySecurityFn get_key_security = nullptr;
    ORSetKeySecurityFn set_key_security = nullptr;

    DWORD load_error() const noexcept
    {
        return load_error_;
    }

  private:
    HMODULE module_ = nullptr;
    DWORD load_error_ = ERROR_SUCCESS;
};

OffregApi& OffregInstance()
{
    static OffregApi api;
    return api;
}

OffregApi* Api()
{
    OffregApi& api = OffregInstance();
    return api.valid() ? &api : nullptr;
}

std::wstring OffregLoadFailure()
{
    std::wstring message = util::Tr(L"offreg.dll couldn't be loaded.");
    const std::wstring detail = win32ui::FormatWin32Error(OffregInstance().load_error());
    if (!detail.empty())
    {
        message += L"\n";
        message += detail;
    }
    return message;
}

struct CloseOfflineKey
{
    void operator()(ORHKEY key) const noexcept
    {
        OffregInstance().close_key(key);
    }
};

std::wstring SelectedControlSet(OffregApi* api, ORHKEY root)
{
    DWORD current = 0;
    DWORD size = sizeof(current);
    DWORD type = 0;
    if (api->get_value(root, L"Select", L"Current", &type, &current, &size) != ERROR_SUCCESS || type != REG_DWORD || current > 999)
    {
        return {};
    }
    wchar_t control_set[16] = {};
    swprintf_s(control_set, L"ControlSet%03lu", current);
    return control_set;
}

// CurrentControlSet is a link made at boot, offline hives use Select\Current unless a real key exists
std::wstring ResolveControlSet(OffregApi* api, ORHKEY root, const std::wstring& subkey)
{
    constexpr std::wstring_view kLink = L"CurrentControlSet";
    ORHKEY real = nullptr;
    if (!registry_path::HasComponentPrefix(subkey, kLink) || api->open_key(root, kLink.data(), &real) == ERROR_SUCCESS)
    {
        if (real)
        {
            api->close_key(real);
        }
        return subkey;
    }
    const std::wstring current = SelectedControlSet(api, root);
    return current.empty() ? subkey : current + subkey.substr(kLink.size());
}

class OfflineKey
{
  public:
    OfflineKey(const RegistryNode& node)
        : api_(Api())
    {
        ORHKEY root = reinterpret_cast<ORHKEY>(node.root);
        // offreg truncates embedded NUL names and would address a different key
        if (!api_ || !root || HasNull(node.subkey))
        {
            return;
        }
        if (node.subkey.empty())
        {
            // caller owns hive roots, this scope owns opened subkeys
            key_ = root;
        }
        else if (api_->open_key(root, ResolveControlSet(api_, root, node.subkey).c_str(), owner_.put()) == ERROR_SUCCESS)
        {
            key_ = owner_.get();
        }
    }
    OfflineKey(OffregApi* api, ORHKEY parent, const std::wstring& name)
        : api_(api)
    {
        if (!HasNull(name) && api_->open_key(parent, name.c_str(), owner_.put()) == ERROR_SUCCESS)
        {
            key_ = owner_.get();
        }
    }

    explicit operator bool() const noexcept
    {
        return key_ != nullptr;
    }
    ORHKEY get() const noexcept
    {
        return key_;
    }
    OffregApi& api() const noexcept
    {
        return *api_;
    }

    LONG QueryInfo(DWORD* subkeys, DWORD* max_subkey_length, DWORD* values, DWORD* max_value_name_length, DWORD* max_value_data_length, FILETIME* last_write, wchar_t* class_name = nullptr, DWORD* class_length = nullptr, DWORD* max_class_length = nullptr) const
    {
        return static_cast<LONG>(api_->query_info(key_, class_name, class_length, subkeys, max_subkey_length, max_class_length, values, max_value_name_length, max_value_data_length, nullptr, last_write));
    }
    LONG EnumKey(DWORD index, wchar_t* name, DWORD* length) const
    {
        return static_cast<LONG>(api_->enum_key(key_, index, name, length, nullptr, nullptr, nullptr));
    }
    LONG EnumValue(DWORD index, wchar_t* name, DWORD* name_length, DWORD* type, BYTE* data, DWORD* data_length) const
    {
        const DWORD capacity = *name_length;
        DWORD result = api_->enum_value(key_, index, name, name_length, type, data, data_length);
        // size only queries return ERROR_MORE_DATA from offreg, success from RegEnumValueW
        if (result == ERROR_MORE_DATA && !data && data_length)
        {
            *name_length = capacity;
            result = api_->enum_value(key_, index, name, name_length, type, nullptr, nullptr);
        }
        return static_cast<LONG>(result);
    }
    LONG GetValue(const std::wstring& name, DWORD* type, BYTE* data, DWORD* size) const
    {
        return HasNull(name) ? ERROR_INVALID_PARAMETER : static_cast<LONG>(api_->get_value(key_, nullptr, ValueNameArg(name), type, data, size));
    }
    LONG SetValue(const std::wstring& name, DWORD type, const BYTE* data, DWORD size) const
    {
        return HasNull(name) ? ERROR_INVALID_PARAMETER : static_cast<LONG>(api_->set_value(key_, ValueNameArg(name), type, data, size));
    }
    LONG DeleteValue(const std::wstring& name) const
    {
        return HasNull(name) ? ERROR_INVALID_PARAMETER : static_cast<LONG>(api_->delete_value(key_, ValueNameArg(name)));
    }
    LONG GetSecurity(SECURITY_INFORMATION information, PSECURITY_DESCRIPTOR descriptor, DWORD* size) const
    {
        return api_->get_key_security ? static_cast<LONG>(api_->get_key_security(key_, information, descriptor, size))
                                      : ERROR_CALL_NOT_IMPLEMENTED;
    }
    LONG SetSecurity(SECURITY_INFORMATION information, PSECURITY_DESCRIPTOR descriptor) const
    {
        return api_->set_key_security ? static_cast<LONG>(api_->set_key_security(key_, information, descriptor))
                                      : ERROR_CALL_NOT_IMPLEMENTED;
    }

  private:
    OffregApi* api_ = nullptr;
    ORHKEY key_ = nullptr;
    win32ui::UniqueNullable<ORHKEY, CloseOfflineKey> owner_;
};

std::shared_mutex g_roots_mutex;
std::vector<HKEY> g_roots;

bool DeleteSubtree(const OfflineKey& parent, const std::wstring& name)
{
    // close child handle before deleting its now empty key
    {
        const OfflineKey child(&parent.api(), parent.get(), name);
        if (!child)
        {
            return false;
        }
        for (const std::wstring& child_name : SubKeyNames(child, false))
        {
            if (!DeleteSubtree(child, child_name))
            {
                return false;
            }
        }
    }
    return parent.api().delete_key(parent.get(), name.c_str()) == ERROR_SUCCESS;
}

template <typename Action>
bool WithApi(std::wstring* error, Action&& action)
{
    if (error)
    {
        error->clear();
    }
    OffregApi* api = Api();
    if (!api)
    {
        if (error)
        {
            *error = OffregLoadFailure();
        }
        return false;
    }
    const DWORD result = action(*api);
    if (result != ERROR_SUCCESS && error)
    {
        *error = win32ui::FormatWin32Error(result);
    }
    return result == ERROR_SUCCESS;
}

}

bool OpenHive(const std::wstring& path, HKEY* root, std::wstring* error)
{
    *root = nullptr;
    return WithApi(error, [&](OffregApi& api) {
        ORHKEY hive = nullptr;
        DWORD result = api.open_hive(path.c_str(), &hive);
        if (result == ERROR_SUCCESS && !hive)
        {
            result = ERROR_INVALID_HANDLE;
        }
        *root = reinterpret_cast<HKEY>(hive);
        return result;
    });
}

bool SaveHive(HKEY root, const std::wstring& path, std::wstring* error)
{
    return root && WithApi(error, [&](OffregApi& api) {
               const RTL_OSVERSIONINFOW& version = win32ui::OsVersion();
               return api.save_hive(reinterpret_cast<ORHKEY>(root), path.c_str(), version.dwMajorVersion ? version.dwMajorVersion : 10, version.dwMinorVersion);
           });
}

bool CloseHive(HKEY root, std::wstring* error)
{
    return !root || WithApi(error, [&](OffregApi& api) { return api.close_hive(reinterpret_cast<ORHKEY>(root)); });
}

void AddRoot(HKEY root)
{
    std::unique_lock lock(g_roots_mutex);
    if (root && std::find(g_roots.begin(), g_roots.end(), root) == g_roots.end())
    {
        g_roots.push_back(root);
    }
}

void RemoveRoot(HKEY root)
{
    std::unique_lock lock(g_roots_mutex);
    std::erase(g_roots, root);
}

bool Owns(HKEY root)
{
    std::shared_lock lock(g_roots_mutex);
    return root && std::find(g_roots.begin(), g_roots.end(), root) != g_roots.end();
}

bool HasSubKeys(const RegistryNode& node)
{
    const OfflineKey key(node);
    return key && registry_backend::HasSubKeys(key);
}

bool QueryKeyInfo(const RegistryNode& node, KeyInfo* info)
{
    const OfflineKey key(node);
    return key && registry_backend::QueryKeyInfo(key, info);
}

bool QuerySymbolicLinkTarget(const RegistryNode& node, std::wstring* target)
{
    target->clear();
    const OfflineKey key(node);
    return key && ReadLinkTarget(key, target) && !target->empty();
}

std::vector<std::wstring> EnumSubKeyNames(const RegistryNode& node, bool sorted)
{
    const OfflineKey key(node);
    return key ? SubKeyNames(key, sorted) : std::vector<std::wstring>();
}

bool EnumKeyStreaming(const RegistryNode& node, bool include_values, bool include_data, bool include_subkeys, RegistryStore::KeyEnumResult* out_info, const RegistryStore::ValueStreamCallback& value_callback, const RegistryStore::SubkeyStreamCallback& subkey_callback, DWORD max_data_size, EnumerationScratch* scratch, bool)
{
    const OfflineKey key(node);
    return key && EnumerateKey(key, include_values, include_data, include_subkeys, out_info, value_callback, subkey_callback, max_data_size, scratch);
}

bool QueryValue(const RegistryNode& node, const std::wstring& value_name, RegistryValue* out)
{
    const OfflineKey key(node);
    return key && registry_backend::QueryValue(key, value_name, out);
}

bool QueryKeyDetails(const RegistryNode& node, KeyDetails* details)
{
    const OfflineKey key(node);
    return key && registry_backend::QueryKeyDetails(key, details) == ERROR_SUCCESS;
}

std::wstring SelectedControlSet(const RegistryNode& node)
{
    OffregApi* api = Api();
    return api && node.root ? SelectedControlSet(api, reinterpret_cast<ORHKEY>(node.root)) : std::wstring();
}

bool CreateKey(const RegistryNode& node, const std::wstring& name, const std::wstring& class_name)
{
    const OfflineKey parent(node);
    if (!parent || HasNull(name))
    {
        return false;
    }
    win32ui::UniqueNullable<ORHKEY, CloseOfflineKey> created;
    DWORD disposition = 0;
    return parent.api().create_key(parent.get(), name.c_str(), class_name.empty() ? nullptr : const_cast<PWSTR>(class_name.c_str()), 0, nullptr, created.put(), &disposition) ==
               ERROR_SUCCESS &&
           disposition == REG_CREATED_NEW_KEY;
}

bool ReadKeySecurity(const RegistryNode& node, SECURITY_INFORMATION* parts, std::vector<BYTE>* descriptor, LONG* error)
{
    descriptor->clear();
    const OfflineKey key(node);
    return key && ReadSecurity(key, parts, descriptor, error);
}

bool WriteKeySecurity(const RegistryNode& node, SECURITY_INFORMATION parts, const std::vector<BYTE>& descriptor)
{
    const OfflineKey key(node);
    return key && (descriptor.empty() || WriteSecurity(key, parts, descriptor));
}

bool DeleteKey(const RegistryNode& node)
{
    RegistryNode parent_node;
    std::wstring name;
    if (!SplitNode(node, &parent_node, &name))
    {
        return false;
    }
    const OfflineKey parent(parent_node);
    return parent && DeleteSubtree(parent, name);
}

bool RenameKey(const RegistryNode& node, const std::wstring& new_name)
{
    const OfflineKey key(node);
    return key && !HasNull(new_name) && key.api().rename_key(key.get(), new_name.c_str()) == ERROR_SUCCESS;
}

bool DeleteValue(const RegistryNode& node, const std::wstring& value_name)
{
    const OfflineKey key(node);
    return key && key.DeleteValue(value_name) == ERROR_SUCCESS;
}

bool SetValue(const RegistryNode& node, const std::wstring& value_name, DWORD type, const std::vector<BYTE>& data)
{
    const OfflineKey key(node);
    return key && key.SetValue(value_name, type, data.empty() ? nullptr : data.data(), static_cast<DWORD>(data.size())) == ERROR_SUCCESS;
}

bool RenameValue(const RegistryNode& node, const std::wstring& old_name, const std::wstring& new_name, bool* both_names_left)
{
    const OfflineKey key(node);
    return key && registry_backend::RenameValue(key, old_name, new_name, both_names_left);
}

}
