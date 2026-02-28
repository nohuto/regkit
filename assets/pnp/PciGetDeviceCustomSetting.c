__int64 __fastcall PciGetDeviceCustomSetting(__int64 a1, const WCHAR *a2)
{
  struct _DEVICE_OBJECT *v2; // rcx
  __int64 v3; // rbx
  HANDLE Handle; // [rsp+40h] [rbp+8h] BYREF
  __int64 v7; // [rsp+50h] [rbp+18h] BYREF
  __int64 v8; // [rsp+58h] [rbp+20h] BYREF

  v2 = *(struct _DEVICE_OBJECT **)(a1 + 128);
  Handle = 0LL;
  v3 = 0LL;
  v8 = 0LL;
  if ( IoOpenDeviceRegistryKey(v2, 1u, 0x20019u, &Handle) >= 0 )
  {
    PciGetRegistryValue(a2, L"e5b3b5ac-9725-4f78-963f-03dfb1d828c7", Handle, 4, &v8, (ULONG *)&v7);
    v3 = v8;
  }
  if ( Handle )
    ZwClose(Handle);
  return v3;
}