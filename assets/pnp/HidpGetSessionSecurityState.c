NTSTATUS __fastcall HidpGetSessionSecurityState(__int64 a1, __int64 a2)
{
  __int64 v2; // rsi
  __int64 v4; // rax
  _QWORD *v5; // rbx
  __int64 v6; // rcx
  __int64 v7; // rax
  bool v8; // al
  struct _DEVICE_OBJECT *v9; // rcx
  NTSTATUS result; // eax
  int v11; // edx
  unsigned int *Pool2; // rbx
  int v13; // edx
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-28h] BYREF
  ULONG Length; // [rsp+70h] [rbp+8h] BYREF
  void *DeviceRegKey; // [rsp+78h] [rbp+10h] BYREF

  v2 = a2;
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(a2) = 5;
    WPP_RECORDER_SF_(
      WPP_GLOBAL_Control->DeviceExtension,
      a2,
      3,
      16,
      (__int64)&WPP_0c3075560e1b37693c50d24ef513d07f_Traceguids);
  }
  v4 = *(unsigned int *)(v2 + 4);
  v5 = *(_QWORD **)(a1 + 64);
  DeviceRegKey = 0LL;
  v6 = 5 * v4;
  v7 = v5[24];
  v8 = *(_WORD *)(v7 + 8 * v6) == 13 && (unsigned __int16)(*(_WORD *)(v7 + 8 * v6 + 2) - 1) <= 3u;
  v9 = *(struct _DEVICE_OBJECT **)(a1 + 48);
  *(_BYTE *)(a1 + 257) = v8;
  result = IoOpenDeviceRegistryKey(v9, 1u, 0x1F0000u, &DeviceRegKey);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return result;
    LOBYTE(v11) = 2;
    result = WPP_RECORDER_SF_qd(
               v5[88],
               v11,
               3,
               17,
               (__int64)&WPP_0c3075560e1b37693c50d24ef513d07f_Traceguids,
               v5[4],
               result);
  }
  else
  {
    DestinationString = 0LL;
    RtlInitUnicodeString(&DestinationString, L"SessionSecurityEnabled");
    Length = DestinationString.MaximumLength + 28;
    Pool2 = (unsigned int *)ExAllocatePool2(256LL, Length, 1130654024LL);
    if ( Pool2 )
    {
      if ( ZwQueryValueKey(DeviceRegKey, &DestinationString, KeyValueFullInformation, Pool2, Length, &Length) >= 0
        && Pool2[3] == 4 )
      {
        *(_BYTE *)(a1 + 257) = *(unsigned int *)((char *)Pool2 + Pool2[2]) != 0;
      }
      ExFreePoolWithTag(Pool2, 0);
    }
    result = ZwClose(DeviceRegKey);
    DeviceRegKey = 0LL;
  }
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    if ( LOWORD(WPP_GLOBAL_Control->DeviceType) )
    {
      LOBYTE(v13) = 5;
      return WPP_RECORDER_SF_(
               WPP_GLOBAL_Control->DeviceExtension,
               v13,
               3,
               18,
               (__int64)&WPP_0c3075560e1b37693c50d24ef513d07f_Traceguids);
    }
  }
  return result;
}