// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// RegKit is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
//
// You should have received a copy of the GNU Affero General Public License
// along with RegKit.  If not, see <https://www.gnu.org/licenses/>.

#include "app/registry_security.h"

#include <string>

#include <accctrl.h>
#include <aclapi.h>
#include <aclui.h>

#include "registry/registry_provider.h"

namespace regkit {

namespace {

bool SetPrivilege(const wchar_t* name, bool enable) {
  HANDLE token = nullptr;
  if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &token)) {
    return false;
  }
  LUID luid = {};
  if (!LookupPrivilegeValueW(nullptr, name, &luid)) {
    CloseHandle(token);
    return false;
  }
  TOKEN_PRIVILEGES tp = {};
  tp.PrivilegeCount = 1;
  tp.Privileges[0].Luid = luid;
  tp.Privileges[0].Attributes = enable ? SE_PRIVILEGE_ENABLED : 0;
  AdjustTokenPrivileges(token, FALSE, &tp, sizeof(tp), nullptr, nullptr);
  DWORD last_error = GetLastError();
  CloseHandle(token);
  return last_error == ERROR_SUCCESS;
}

class RegistrySecurityInformation : public ISecurityInformation {
public:
  RegistrySecurityInformation(HKEY key, std::wstring object_name, bool read_only) : key_(key), object_name_(std::move(object_name)), read_only_(read_only) {}

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
    if (!ppv) {
      return E_POINTER;
    }
    *ppv = nullptr;
    if (riid == IID_IUnknown || riid == IID_ISecurityInformation) {
      *ppv = static_cast<ISecurityInformation*>(this);
      return S_OK;
    }
    return E_NOINTERFACE;
  }

  ULONG STDMETHODCALLTYPE AddRef() override { return 2; }

  ULONG STDMETHODCALLTYPE Release() override { return 1; }

  HRESULT STDMETHODCALLTYPE GetObjectInformation(PSI_OBJECT_INFO info) override {
    if (!info) {
      return E_POINTER;
    }
    DWORD flags = SI_ADVANCED | SI_EDIT_OWNER | SI_EDIT_PERMS | SI_CONTAINER;
    if (read_only_) {
      flags |= SI_READONLY | SI_OWNER_READONLY;
    }
    info->dwFlags = flags;
    info->hInstance = nullptr;
    info->pszServerName = nullptr;
    info->pszObjectName = const_cast<wchar_t*>(object_name_.c_str());
    info->pszPageTitle = nullptr;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE GetSecurity(SECURITY_INFORMATION security_info, PSECURITY_DESCRIPTOR* out_sd, BOOL) override {
    if (!out_sd) {
      return E_POINTER;
    }
    *out_sd = nullptr;
    DWORD result = GetSecurityInfo(key_, SE_REGISTRY_KEY, security_info, nullptr, nullptr, nullptr, nullptr, out_sd);
    return HRESULT_FROM_WIN32(result);
  }

  HRESULT STDMETHODCALLTYPE SetSecurity(SECURITY_INFORMATION security_info, PSECURITY_DESCRIPTOR sd) override {
    if (!sd) {
      return E_POINTER;
    }
    return SetKernelObjectSecurity(key_, security_info, sd) ? S_OK : HRESULT_FROM_WIN32(GetLastError());
  }

  HRESULT STDMETHODCALLTYPE GetAccessRights(const GUID*, DWORD, PSI_ACCESS* access, ULONG* count, ULONG* default_access) override {
    static SI_ACCESS rights[] = {
        {&GUID_NULL, KEY_CREATE_SUB_KEY, const_cast<wchar_t*>(L"Create"), SI_ACCESS_SPECIFIC}, {&GUID_NULL, KEY_ENUMERATE_SUB_KEYS, const_cast<wchar_t*>(L"Enumerate"), SI_ACCESS_SPECIFIC}, {&GUID_NULL, KEY_SET_VALUE, const_cast<wchar_t*>(L"Set Value"), SI_ACCESS_SPECIFIC}, {&GUID_NULL, KEY_QUERY_VALUE, const_cast<wchar_t*>(L"Query Value"), SI_ACCESS_SPECIFIC}, {&GUID_NULL, KEY_WRITE, const_cast<wchar_t*>(L"Write"), SI_ACCESS_GENERAL}, {&GUID_NULL, KEY_READ, const_cast<wchar_t*>(L"Read"), SI_ACCESS_GENERAL},
    };
    if (access) {
      *access = rights;
    }
    if (count) {
      *count = static_cast<ULONG>(_countof(rights));
    }
    if (default_access) {
      *default_access = 0;
    }
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE MapGeneric(const GUID*, UCHAR*, ACCESS_MASK*) override { return S_OK; }

  HRESULT STDMETHODCALLTYPE GetInheritTypes(PSI_INHERIT_TYPE* types, ULONG* count) override {
    static SI_INHERIT_TYPE inherit_types[] = {
        {&GUID_NULL, 0, const_cast<wchar_t*>(L"This key only")},
        {&GUID_NULL, CONTAINER_INHERIT_ACE, const_cast<wchar_t*>(L"This key and subkeys")},
    };
    if (types) {
      *types = inherit_types;
    }
    if (count) {
      *count = static_cast<ULONG>(_countof(inherit_types));
    }
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE PropertySheetPageCallback(HWND, UINT, SI_PAGE_TYPE) override { return S_OK; }

private:
  HKEY key_ = nullptr;
  std::wstring object_name_;
  bool read_only_ = false;
};

} // namespace

bool ShowRegistryPermissions(HWND owner, const RegistryNode& node) {
  std::wstring path = RegistryProvider::BuildPath(node);
  if (path.empty()) {
    return false;
  }

  bool privilege_enabled = SetPrivilege(SE_TAKE_OWNERSHIP_NAME, true);
  const wchar_t* subkey = node.subkey.empty() ? nullptr : node.subkey.c_str();
  bool read_only = false;
  HKEY key = nullptr;
  LONG result = RegOpenKeyExW(node.root, subkey, 0, READ_CONTROL | WRITE_DAC | WRITE_OWNER, &key);
  if (result == ERROR_ACCESS_DENIED) {
    read_only = true;
    result = RegOpenKeyExW(node.root, subkey, 0, READ_CONTROL, &key);
  }
  if (result == ERROR_ACCESS_DENIED) {
    result = RegOpenKeyExW(node.root, subkey, 0, MAXIMUM_ALLOWED, &key);
  }

  bool ok = false;
  if (result == ERROR_SUCCESS && key) {
    RegistrySecurityInformation info(key, path, read_only);
    ok = SUCCEEDED(EditSecurity(owner, &info));
    RegCloseKey(key);
  }

  if (privilege_enabled) {
    SetPrivilege(SE_TAKE_OWNERSHIP_NAME, false);
  }
  return ok;
}

} // namespace regkit
