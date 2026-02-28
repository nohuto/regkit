int __fastcall HidpGetRetainWWIrpEnabledFromRegistry(__int64 a1, struct _DEVICE_OBJECT *a2)
{
  int result; // eax
  int v4; // edx
  unsigned int *Pool2; // rbx
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-18h] BYREF
  ULONG Length; // [rsp+60h] [rbp+8h] BYREF
  void *DeviceRegKey; // [rsp+70h] [rbp+18h] BYREF

  DeviceRegKey = 0LL;
  *(_BYTE *)(a1 + 621) = 0;
  result = IoOpenDeviceRegistryKey(a2, 1u, 0x1F0000u, &DeviceRegKey);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v4) = 2;
      return WPP_RECORDER_SF_qd(
               *(_QWORD *)(a1 + 704),
               v4,
               3,
               19,
               (__int64)&WPP_0c3075560e1b37693c50d24ef513d07f_Traceguids,
               *(_QWORD *)(a1 + 32),
               result);
    }
  }
  else
  {
    DestinationString = 0LL;
    RtlInitUnicodeString(&DestinationString, L"RetainWWIrpWhenDeviceAbsent");
    Length = DestinationString.MaximumLength + 28;
    Pool2 = (unsigned int *)ExAllocatePool2(256LL, Length, 1130654024LL);
    if ( Pool2 )
    {
      if ( ZwQueryValueKey(DeviceRegKey, &DestinationString, KeyValueFullInformation, Pool2, Length, &Length) >= 0
        && Pool2[3] == 4 )
      {
        *(_BYTE *)(a1 + 621) = *(unsigned int *)((char *)Pool2 + Pool2[2]) != 0;
      }
      ExFreePoolWithTag(Pool2, 0);
    }
    return ZwClose(DeviceRegKey);
  }
  return result;
}