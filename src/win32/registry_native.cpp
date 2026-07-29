// Copyright (C) 2026 Noverse (Nohuto)
// This file is part of RegKit https://github.com/nohuto/regkit
//
// RegKit is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

#include "win32/registry_native.h"

#include <winternl.h>

namespace util {
namespace {

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif

#ifndef OBJ_OPENLINK
#define OBJ_OPENLINK 0x00000008L
#endif

using NtOpenKey = NTSTATUS(NTAPI*)(PHANDLE, ACCESS_MASK, POBJECT_ATTRIBUTES);

NtOpenKey ResolveNtOpenKey() {
  HMODULE module = GetModuleHandleW(L"ntdll.dll");
  return module
             ? reinterpret_cast<NtOpenKey>(GetProcAddress(module, "NtOpenKey"))
             : nullptr;
}

} // namespace

UniqueHKey OpenNativeRegistryKey(const std::wstring& path, REGSAM access,
                                 bool open_link) {
  static const NtOpenKey open_key = ResolveNtOpenKey();
  if (!open_key || path.empty()) {
    return {};
  }
  UNICODE_STRING name = {};
  name.Buffer = const_cast<PWSTR>(path.c_str());
  name.Length = static_cast<USHORT>(path.size() * sizeof(wchar_t));
  name.MaximumLength = name.Length;
  OBJECT_ATTRIBUTES attributes = {};
  InitializeObjectAttributes(
      &attributes, &name,
      OBJ_CASE_INSENSITIVE | (open_link ? OBJ_OPENLINK : 0), nullptr,
      nullptr);
  HANDLE handle = nullptr;
  if (!NT_SUCCESS(open_key(&handle, access, &attributes)) || !handle) {
    return {};
  }
  return UniqueHKey(reinterpret_cast<HKEY>(handle));
}

UniqueHKey OpenNativeRegistryRoot() {
  return OpenNativeRegistryKey(L"\\REGISTRY", KEY_READ);
}

} // namespace util
