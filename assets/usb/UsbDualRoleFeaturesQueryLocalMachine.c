void __fastcall UsbDualRoleFeaturesQueryLocalMachine(unsigned int *a1)
{
  int PersistedKeyPath; // eax
  int v3; // edx
  __int64 v4; // rcx
  PVOID v5; // r14
  int v6; // esi
  int v7; // r9d
  __int64 v8; // rcx
  __int64 v9; // rdx
  int v10; // eax
  unsigned int v11; // edx
  __int64 v12; // rcx
  PVOID v13; // rsi
  int v14; // r9d
  int v15; // eax
  int v16; // r8d
  int v17; // r8d
  int v18; // r14d
  int v19; // r9d
  int v20; // r9d
  int v21; // [rsp+20h] [rbp-20h]
  HANDLE Handle; // [rsp+80h] [rbp+40h] BYREF
  PVOID P; // [rsp+88h] [rbp+48h] BYREF
  PVOID v24; // [rsp+90h] [rbp+50h] BYREF

  Handle = 0LL;
  P = 0LL;
  PersistedKeyPath = GetPersistedKeyPath(&P);
  v5 = P;
  v6 = PersistedKeyPath;
  if ( PersistedKeyPath < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_14;
    v7 = 13;
    goto LABEL_4;
  }
  PersistedKeyPath = MyRegOpenKeyForRead(v4, P, &Handle);
  v6 = PersistedKeyPath;
  if ( PersistedKeyPath >= 0 )
  {
    PersistedKeyPath = MyRegQueryUlong(Handle, L"DualRoleFeaturesTestOverride", a1);
    v6 = PersistedKeyPath;
    if ( PersistedKeyPath >= 0 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v3) = 4;
        WPP_RECORDER_SF_d(
          WPP_GLOBAL_Control->DeviceExtension,
          v3,
          1,
          16,
          (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids,
          *a1);
      }
    }
    else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v7 = 15;
      LOBYTE(v3) = 4;
      goto LABEL_5;
    }
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v7 = 14;
LABEL_4:
    LOBYTE(v3) = 2;
LABEL_5:
    WPP_RECORDER_SF_d(
      WPP_GLOBAL_Control->DeviceExtension,
      v3,
      1,
      v7,
      (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids,
      PersistedKeyPath);
  }
LABEL_14:
  if ( v5 )
    ExFreePoolWithTag(v5, 0);
  if ( Handle )
    ZwClose(Handle);
  if ( v6 < 0 )
  {
    ReadManifestAssignedValue(a1);
    *a1 &= 0xFFFFFFF1;
    if ( CheckUSBFnIncludeDefaultCfg(v8) )
      CheckUSBFnConfiguration(a1, L"Default");
    if ( (int)ReadUSBFnFeaturesFromCurrentConfiguration(a1, 0LL) < 0 )
    {
      LOBYTE(v9) = 1;
      ReadUSBFnFeaturesFromCurrentConfiguration(a1, v9);
    }
  }
  P = 0LL;
  v24 = 0LL;
  LODWORD(Handle) = 0;
  v10 = GetPersistedKeyPath(&v24);
  v13 = v24;
  if ( v10 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_36;
    v14 = 33;
    LOBYTE(v11) = 2;
    goto LABEL_26;
  }
  v15 = MyRegOpenKeyForRead(v12, v24, &P);
  if ( v15 >= 0 )
  {
    v10 = MyRegQueryUlong(P, L"UcmIsPresent", &Handle);
    if ( v10 >= 0 )
    {
      v18 = (int)Handle;
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v11) = 4;
        WPP_RECORDER_SF_Sd(WPP_GLOBAL_Control->DeviceExtension, v11, v17, 36, v21, (__int64)v13, (char)Handle);
      }
      v11 = v18 != 0 ? 0x80000000 : 0;
      *a1 = v11 | *a1 & 0x7FFFFFFF;
    }
    else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v14 = 35;
      LOBYTE(v11) = 3;
LABEL_26:
      WPP_RECORDER_SF_d(
        WPP_GLOBAL_Control->DeviceExtension,
        v11,
        1,
        v14,
        (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids,
        v10);
    }
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v11) = 4;
    WPP_RECORDER_SF_Sd(WPP_GLOBAL_Control->DeviceExtension, v11, v16, 34, v21, (__int64)v13, v15);
  }
LABEL_36:
  if ( v13 )
    ExFreePoolWithTag(v13, 0);
  if ( P )
    ZwClose(P);
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v11) = 4;
    WPP_RECORDER_SF_d(
      WPP_GLOBAL_Control->DeviceExtension,
      v11,
      1,
      37,
      (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids,
      *a1);
  }
  if ( (*(_BYTE *)a1 & 1) != 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_48;
    v19 = 38;
  }
  else
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_48;
    v19 = 39;
  }
  LOBYTE(v11) = 4;
  WPP_RECORDER_SF_(
    WPP_GLOBAL_Control->DeviceExtension,
    v11,
    1,
    v19,
    (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids);
LABEL_48:
  if ( (*(_BYTE *)a1 & 2) != 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return;
    v20 = 40;
  }
  else
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return;
    v20 = 41;
  }
  LOBYTE(v11) = 4;
  WPP_RECORDER_SF_(
    WPP_GLOBAL_Control->DeviceExtension,
    v11,
    1,
    v20,
    (__int64)&WPP_5169c4c8089132207a438b4341aed5b6_Traceguids);
}