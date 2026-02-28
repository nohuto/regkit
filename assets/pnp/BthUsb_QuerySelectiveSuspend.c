unsigned __int8 __fastcall BthUsb_QuerySelectiveSuspend(struct _BTDEVICE *a1)
{
  void *MiniportContext; // r14
  char v2; // di
  unsigned __int8 v4; // si
  __int64 v5; // rcx
  __int64 v6; // rdx
  char v7; // cl
  int v8; // edx
  __int16 v10; // [rsp+50h] [rbp-B0h] BYREF
  int v11; // [rsp+54h] [rbp-ACh]
  void *DeviceRegKey[20]; // [rsp+60h] [rbp-A0h] BYREF

  MiniportContext = a1->MiniportContext;
  v2 = 0;
  DeviceRegKey[0] = 0LL;
  v4 = 0;
  v10 = 0;
  v11 = 1;
  memset(&DeviceRegKey[2], 0, 0x88uLL);
  if ( IoOpenDeviceRegistryKey(a1->PhysicalDeviceObject, 1u, 0x20000u, DeviceRegKey) >= 0 )
  {
    v4 = (int)BthQueryKeyValue(DeviceRegKey[0]) >= 0;// ForceSelectiveSuspend
    ZwClose(DeviceRegKey[0]);
    if ( v4 )
    {
      *(_BYTE *)(*((_QWORD *)MiniportContext + 13) + 7LL) |= 0x60u;
      v5 = *((_QWORD *)MiniportContext + 13);
      *((_BYTE *)MiniportContext + 1170) = 1;
      *((_BYTE *)MiniportContext + 1168) = (*(_BYTE *)(v5 + 7) & 0x20) != 0;
      *((_BYTE *)MiniportContext + 1169) = (*(_BYTE *)(v5 + 7) & 0x40) != 0;
LABEL_12:
      *((_BYTE *)MiniportContext + 1172) = 1;
      goto LABEL_13;
    }
  }
  LODWORD(DeviceRegKey[2]) = 1245320;
  DeviceRegKey[7] = &v10;
  HIDWORD(DeviceRegKey[6]) = 2;
  DeviceRegKey[8] = 0LL;
  WORD2(DeviceRegKey[18]) = 0;
  DeviceRegKey[9] = 0LL;
  if ( (int)USBCallSyncEx(*((PVOID *)MiniportContext + 5)) >= 0 )
  {
    v6 = *((_QWORD *)MiniportContext + 13);
    v7 = v10 & 1;
    *((_BYTE *)MiniportContext + 1170) = v10 & 1;
    *((_BYTE *)MiniportContext + 1168) = (*(_BYTE *)(v6 + 7) & 0x20) != 0;
    *((_BYTE *)MiniportContext + 1169) = (*(_BYTE *)(v6 + 7) & 0x40) != 0;
    if ( v7 )
    {
      if ( (*(_BYTE *)(v6 + 7) & 0x60) == 0x60 )
      {
        *((_BYTE *)MiniportContext + 1171) = 0;
        v4 = 1;
        if ( IoOpenDeviceRegistryKey(a1->PhysicalDeviceObject, 1u, 0x20000u, DeviceRegKey) < 0 )
          goto LABEL_12;
        if ( (int)BthQueryKeyValue(DeviceRegKey[0]) >= 0 && !v11 )// SelectiveSuspendSupported
        {
          *((_BYTE *)MiniportContext + 1171) = 1;
          v4 = 0;
        }
        ZwClose(DeviceRegKey[0]);
        if ( v4 )
          goto LABEL_12;
      }
    }
  }
LABEL_13:
  if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
    v2 = 1;
  v8 = 29;
  LOBYTE(v8) = v2;
  WPP_RECORDER_AND_TRACE_SF_d(
    WPP_GLOBAL_Control->AttachedDevice,
    v8,
    v4,
    WPP_GLOBAL_Control->DeviceExtension,
    4,
    1,
    29,
    (__int64)&WPP_1f9a50d1de1b36addda1af3e8dc29e49_Traceguids,
    v4);
  return v4;
}