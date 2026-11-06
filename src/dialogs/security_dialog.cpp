// Copyright (C) 2026 nohuto
// SPDX-License-Identifier: AGPL-3.0-or-later

#include "dialogs/security_dialog.h"

#include "registry/key_access.h"
#include "registry/registry_path.h"
#include "ui/feedback.h"
#include "win32/process_rights.h"
#include "win32/registry_native.h"
#include "win32/registry_view.h"
#include "win32/translation.h"

#include <win32ui/handle_owner.h>
#include <win32ui/system/error_text.h>

#include <cstring>
#include <string>
#include <vector>

#include <accctrl.h>
#include <aclapi.h>
#include <aclui.h>

namespace regkit
{

namespace
{

// aclui may pass an absolute descriptor
bool SelfRelativeCopy(PSECURITY_DESCRIPTOR descriptor, std::vector<BYTE>* copy)
{
    SECURITY_DESCRIPTOR_CONTROL control = 0;
    DWORD revision = 0;
    if (!GetSecurityDescriptorControl(descriptor, &control, &revision))
    {
        return false;
    }
    if (control & SE_SELF_RELATIVE)
    {
        const auto* bytes = static_cast<const BYTE*>(descriptor);
        copy->assign(bytes, bytes + GetSecurityDescriptorLength(descriptor));
        return true;
    }
    DWORD size = 0;
    MakeSelfRelativeSD(descriptor, nullptr, &size);
    copy->resize(size);
    return size != 0 && MakeSelfRelativeSD(descriptor, copy->data(), &size);
}

// null key is an offline hive, read and written through RegistryStore
class RegistrySecurityInformation : public ISecurityInformation, public IEffectivePermission
{
  public:
    RegistrySecurityInformation(const RegistryNode& node, HKEY key, std::wstring object_name, REGSAM granted, bool audits)
        : node_(node), key_(key), object_name_(std::move(object_name)), granted_(granted), audits_(audits)
    {
    }

    bool changed() const noexcept
    {
        return changed_;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv)
        {
            return E_POINTER;
        }
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == IID_ISecurityInformation)
        {
            *ppv = static_cast<ISecurityInformation*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == IID_IEffectivePermission)
        {
            *ppv = static_cast<IEffectivePermission*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return static_cast<ULONG>(InterlockedIncrement(&references_));
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const LONG remaining = InterlockedDecrement(&references_);
        return static_cast<ULONG>(remaining < 0 ? 0 : remaining);
    }

    HRESULT STDMETHODCALLTYPE GetObjectInformation(PSI_OBJECT_INFO info) override
    {
        if (!info)
        {
            return E_POINTER;
        }
        DWORD flags = SI_ADVANCED | SI_EDIT_OWNER | SI_EDIT_PERMS | SI_CONTAINER | SI_EDIT_EFFECTIVE | (audits_ ? SI_EDIT_AUDITS : 0);
        if (!(granted_ & WRITE_DAC))
        {
            flags |= SI_READONLY;
        }
        if (!(granted_ & WRITE_OWNER))
        {
            flags |= SI_OWNER_READONLY;
        }
        info->dwFlags = flags;
        info->hInstance = nullptr;
        info->pszServerName = nullptr;
        info->pszObjectName = const_cast<wchar_t*>(object_name_.c_str());
        info->pszPageTitle = nullptr;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetSecurity(SECURITY_INFORMATION security_info, PSECURITY_DESCRIPTOR* out_sd, BOOL) override
    {
        if (!out_sd)
        {
            return E_POINTER;
        }
        *out_sd = nullptr;
        if (!key_)
        {
            SECURITY_INFORMATION parts = security_info;
            std::vector<BYTE> stored;
            LONG error = ERROR_ACCESS_DENIED;
            if (!RegistryStore::ReadKeySecurity(node_, &parts, &stored, &error) || parts != security_info)
            {
                return HRESULT_FROM_WIN32(error);
            }
            *out_sd = LocalAlloc(LMEM_FIXED, stored.size());
            if (!*out_sd)
            {
                return E_OUTOFMEMORY;
            }
            std::memcpy(*out_sd, stored.data(), stored.size());
            return S_OK;
        }
        DWORD size = 0;
        LSTATUS status = RegGetKeySecurity(key_, security_info, nullptr, &size);
        if (status != ERROR_INSUFFICIENT_BUFFER || size == 0)
        {
            return HRESULT_FROM_WIN32(status == ERROR_SUCCESS ? ERROR_INVALID_DATA : status);
        }
        PSECURITY_DESCRIPTOR descriptor = LocalAlloc(LPTR, size);
        if (!descriptor)
        {
            return E_OUTOFMEMORY;
        }
        status = RegGetKeySecurity(key_, security_info, descriptor, &size);
        if (status != ERROR_SUCCESS)
        {
            LocalFree(descriptor);
            return HRESULT_FROM_WIN32(status);
        }
        *out_sd = descriptor;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetSecurity(SECURITY_INFORMATION security_info, PSECURITY_DESCRIPTOR sd) override
    {
        if (!sd)
        {
            return E_POINTER;
        }
        if (!key_)
        {
            std::vector<BYTE> descriptor;
            if (!SelfRelativeCopy(sd, &descriptor))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            if (!RegistryStore::WriteKeySecurity(node_, security_info, descriptor))
            {
                return E_FAIL;
            }
            changed_ = true;
            return S_OK;
        }
        const LSTATUS status = RegSetKeySecurity(key_, security_info, sd);
        changed_ = changed_ || status == ERROR_SUCCESS;
        return status == ERROR_SUCCESS ? S_OK : HRESULT_FROM_WIN32(status);
    }

    HRESULT STDMETHODCALLTYPE GetAccessRights(const GUID*, DWORD, PSI_ACCESS* access, ULONG* count, ULONG* default_access) override
    {
        static SI_ACCESS rights[] = {
            {&GUID_NULL, KEY_CREATE_SUB_KEY, const_cast<wchar_t*>(util::Tr(L"Create")), SI_ACCESS_SPECIFIC},
            {&GUID_NULL, KEY_ENUMERATE_SUB_KEYS, const_cast<wchar_t*>(util::Tr(L"Enumerate")), SI_ACCESS_SPECIFIC},
            {&GUID_NULL, KEY_SET_VALUE, const_cast<wchar_t*>(util::Tr(L"Set Value")), SI_ACCESS_SPECIFIC},
            {&GUID_NULL, KEY_QUERY_VALUE, const_cast<wchar_t*>(util::Tr(L"Query Value")), SI_ACCESS_SPECIFIC},
            {&GUID_NULL, KEY_WRITE, const_cast<wchar_t*>(util::Tr(L"Write")), SI_ACCESS_GENERAL},
            {&GUID_NULL, KEY_READ, const_cast<wchar_t*>(util::Tr(L"Read")), SI_ACCESS_GENERAL},
        };
        if (access)
        {
            *access = rights;
        }
        if (count)
        {
            *count = static_cast<ULONG>(_countof(rights));
        }
        if (default_access)
        {
            *default_access = 0;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE MapGeneric(const GUID*, UCHAR*, ACCESS_MASK* mask) override
    {
        if (!mask)
        {
            return E_POINTER;
        }
        GENERIC_MAPPING mapping = {};
        mapping.GenericRead = KEY_READ;
        mapping.GenericWrite = KEY_WRITE;
        mapping.GenericExecute = KEY_EXECUTE;
        mapping.GenericAll = KEY_ALL_ACCESS;
        MapGenericMask(mask, &mapping);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetInheritTypes(PSI_INHERIT_TYPE* types, ULONG* count) override
    {
        static SI_INHERIT_TYPE inherit_types[] = {
            {&GUID_NULL, 0, const_cast<wchar_t*>(util::Tr(L"This key only"))},
            {&GUID_NULL, CONTAINER_INHERIT_ACE, const_cast<wchar_t*>(util::Tr(L"This key and subkeys"))},
        };
        if (types)
        {
            *types = inherit_types;
        }
        if (count)
        {
            *count = static_cast<ULONG>(_countof(inherit_types));
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE PropertySheetPageCallback(HWND, UINT, SI_PAGE_TYPE) override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetEffectivePermission(const GUID*, PSID user, LPCWSTR, PSECURITY_DESCRIPTOR descriptor, POBJECT_TYPE_LIST* object_types, ULONG* object_type_count, PACCESS_MASK* granted_list, ULONG* granted_count) override
    {
        static OBJECT_TYPE_LIST default_types[] = {{ACCESS_OBJECT_GUID, 0, const_cast<GUID*>(&GUID_NULL)}};
        // on failure, page reports that Windows cannot calculate effective permissions
        const std::optional<ACCESS_MASK> access = key_access::ForSid(descriptor, user);
        if (!access)
        {
            return E_FAIL;
        }
        auto* granted = static_cast<ACCESS_MASK*>(LocalAlloc(LPTR, sizeof(ACCESS_MASK)));
        if (!granted)
        {
            return E_OUTOFMEMORY;
        }
        *granted = *access;
        *object_types = default_types;
        *object_type_count = 1;
        *granted_list = granted;
        *granted_count = 1;
        return S_OK;
    }

  private:
    RegistryNode node_;
    HKEY key_ = nullptr;
    std::wstring object_name_;
    REGSAM granted_ = 0;
    bool audits_ = false;
    bool changed_ = false;
    LONG references_ = 1;
};

}

bool ShowRegistryPermissions(HWND owner, const RegistryNode& node)
{
    std::wstring path = registry_path::Build(node);
    if (path.empty())
    {
        return false;
    }

    const util::PrivilegeScope privilege({SE_TAKE_OWNERSHIP_NAME});
    const util::PrivilegeScope restore({SE_RESTORE_NAME});
    // auditing needs SACL access, available only with the security privilege
    const util::PrivilegeScope audit({SE_SECURITY_NAME});
    const REGSAM audits = audit.held() ? ACCESS_SYSTEM_SECURITY : 0;
    const bool offline = RegistryStore::IsOfflineRoot(node.root);
    win32ui::UniqueHKey key;
    REGSAM granted = READ_CONTROL | WRITE_DAC | WRITE_OWNER;
    // without read access aclui still lets the owner rights or a granted WRITE_DAC replace the settings
    const auto open = [&] {
        for (const REGSAM rights : {REGSAM(READ_CONTROL | WRITE_DAC | WRITE_OWNER), REGSAM(READ_CONTROL | WRITE_DAC), REGSAM(READ_CONTROL | WRITE_OWNER), REGSAM(READ_CONTROL), REGSAM(WRITE_DAC | WRITE_OWNER), REGSAM(WRITE_DAC), REGSAM(WRITE_OWNER)})
        {
            for (const REGSAM access : {rights | audits, rights})
            {
                const LONG status = util::OpenRegistryPath(node.root, node.subkey, access | ViewOf(node), true, &key);
                if (status != ERROR_ACCESS_DENIED && status != ERROR_PRIVILEGE_NOT_HELD)
                {
                    granted = access;
                    return status;
                }
            }
        }
        return static_cast<LONG>(ERROR_ACCESS_DENIED);
    };
    const LONG result = offline ? ERROR_SUCCESS : open();

    DWORD error = static_cast<DWORD>(result);
    bool changed = false;
    if (result == ERROR_SUCCESS)
    {
        RegistrySecurityInformation info(node, key.get(), path, granted, offline || (audits && (granted & ACCESS_SYSTEM_SECURITY)));
        error = EditSecurity(owner, &info) ? ERROR_SUCCESS : GetLastError();
        changed = info.changed();
    }
    if (error != ERROR_SUCCESS)
    {
        ui::ShowError(owner, util::TrDetail(L"The key's permissions couldn't be read.", path + L"\n" + win32ui::FormatWin32Error(error)));
    }
    return changed;
}

}
