void __fastcall HidpGetPdoReenumerateSelfInterfaceEnabled(__int64 a1, struct _DEVICE_OBJECT *a2)
{
  NTSTATUS v3; // eax
  int v4; // edx
  unsigned int *Pool2; // rax
  int v6; // edx
  unsigned int *v7; // rdi
  NTSTATUS v8; // eax
  int v9; // edx
  int v10; // r9d
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-10h] BYREF
  ULONG Length; // [rsp+60h] [rbp+10h] BYREF
  void *DeviceRegKey; // [rsp+70h] [rbp+20h] BYREF

  DeviceRegKey = 0LL;
  Length = 0;
  *(_BYTE *)(a1 + 1970) = 0;
  DestinationString = 0LL;
  v3 = IoOpenDeviceRegistryKey(a2, 1u, 0x1F0000u, &DeviceRegKey);
  if ( v3 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v4) = 2;
      WPP_RECORDER_SF_qd(
        *(_QWORD *)(a1 + 672),
        v4,
        3,
        29,
        (__int64)&WPP_0c3075560e1b37693c50d24ef513d07f_Traceguids,
        *(_QWORD *)a1,
        v3);
    }
    goto LABEL_16;
  }
  RtlInitUnicodeString(&DestinationString, L"CollectionReenumerateSelfInterfaceEnabled");
  Length = DestinationString.MaximumLength + 28;
  Pool2 = (unsigned int *)ExAllocatePool2(256LL, Length, 1130654024LL);
  v7 = Pool2;
  if ( Pool2 )
  {
    v8 = ZwQueryValueKey(DeviceRegKey, &DestinationString, KeyValueFullInformation, Pool2, Length, &Length);
    if ( v8 >= 0 )
    {
      if ( v7[3] != 4 )
        goto LABEL_15;
      LOBYTE(v8) = *(unsigned int *)((char *)v7 + v7[2]) != 0;
      *(_BYTE *)(a1 + 1970) = v8;
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_15;
      v10 = 32;
      LOBYTE(v9) = 4;
    }
    else
    {
      if ( v8 == -1073741772 || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_15;
      v10 = 31;
      LOBYTE(v9) = 3;
    }
    WPP_RECORDER_SF_qd(
      *(_QWORD *)(a1 + 672),
      v9,
      3,
      v10,
      (__int64)&WPP_0c3075560e1b37693c50d24ef513d07f_Traceguids,
      *(_QWORD *)a1,
      v8);
LABEL_15:
    ExFreePoolWithTag(v7, 0);
    goto LABEL_16;
  }
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v6) = 2;
    WPP_RECORDER_SF_qdL(
      *(_QWORD *)(a1 + 672),
      v6,
      3,
      30,
      (__int64)&WPP_0c3075560e1b37693c50d24ef513d07f_Traceguids,
      *(_QWORD *)a1,
      Length,
      154);
  }
LABEL_16:
  if ( DeviceRegKey )
    ZwClose(DeviceRegKey);
}