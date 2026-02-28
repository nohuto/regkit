__int64 __fastcall WinUSB_GetRegParams(__int64 a1, int *a2)
{
  int *v2; // r14
  __int64 v4; // rax
  int v5; // eax
  int v6; // edx
  unsigned int v7; // edi
  int v8; // r9d
  unsigned __int64 v9; // rax
  int v10; // edx
  int v11; // edx
  int v12; // edx
  int v13; // edx
  char v14; // al
  char v15; // al
  char v16; // al
  int v17; // eax
  int v18; // edx
  int v19; // edx
  __int64 v21; // [rsp+40h] [rbp-39h] BYREF
  __int64 v22; // [rsp+48h] [rbp-31h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+50h] [rbp-29h] BYREF
  __int128 v24; // [rsp+60h] [rbp-19h] BYREF
  __int64 v25; // [rsp+70h] [rbp-9h]
  __int64 v26; // [rsp+78h] [rbp-1h]
  __int128 v27; // [rsp+80h] [rbp+7h]
  __int64 v28; // [rsp+90h] [rbp+17h]
  unsigned int v29; // [rsp+F0h] [rbp+77h] BYREF
  __int64 v30; // [rsp+F8h] [rbp+7Fh] BYREF

  DestinationString = 0LL;
  v29 = 0;
  v2 = a2;
  v30 = 0LL;
  v22 = 0LL;
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(a2) = 5;
    WPP_RECORDER_SF_(
      WPP_GLOBAL_Control->DeviceExtension,
      (_DWORD)a2,
      1,
      67,
      (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
  }
  v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1632))(WdfDriverGlobals, a1);
  v5 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 384))(
         WdfDriverGlobals,
         v4,
         1LL,
         2031616LL,
         0LL,
         &v30);
  v7 = v5;
  if ( v5 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_69;
    v8 = 68;
    goto LABEL_68;
  }
  RtlInitUnicodeString(&DestinationString, L"DeviceInterfaceGUIDs");
  if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, _QWORD, _QWORD))(WdfFunctions_01015 + 1896))(
         WdfDriverGlobals,
         v30,
         &DestinationString,
         0LL,
         *(_QWORD *)(a1 + 152)) < 0 )
  {
    v21 = 0LL;
    v28 = 0LL;
    v9 = *(_QWORD *)(a1 + 152);
    v25 = 0LL;
    v26 = 0x100000001LL;
    v27 = v9;
    v24 = 0LL;
    LODWORD(v24) = 56;
    v5 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int128 *, __int64 *))(WdfFunctions_01015 + 2464))(
           WdfDriverGlobals,
           0LL,
           &v24,
           &v21);
    v7 = v5;
    if ( v5 < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_69;
      v8 = 69;
      goto LABEL_68;
    }
    RtlInitUnicodeString(&DestinationString, L"DeviceInterfaceGUID");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64))(WdfFunctions_01015
                                                                                                + 1912))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           v21) >= 0 )
    {
      v5 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64))(WdfFunctions_01015 + 120))(
             WdfDriverGlobals,
             *(_QWORD *)(a1 + 152),
             v21);
      v7 = v5;
      if ( v5 < 0 )
      {
        if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          goto LABEL_69;
        v8 = 70;
        goto LABEL_68;
      }
    }
  }
  RtlInitUnicodeString(&DestinationString, L"ResetPortEnabled");
  if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64))(WdfFunctions_01015 + 1920))(
         WdfDriverGlobals,
         v30,
         &DestinationString,
         a1 + 304) < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v10) = 4;
      WPP_RECORDER_SF_(
        WPP_GLOBAL_Control->DeviceExtension,
        v10,
        11,
        71,
        (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
    }
    *(_DWORD *)(a1 + 304) = 0;
  }
  RtlInitUnicodeString(&DestinationString, L"CyclePortEnabled");
  if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64))(WdfFunctions_01015 + 1920))(
         WdfDriverGlobals,
         v30,
         &DestinationString,
         a1 + 308) < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v11) = 4;
      WPP_RECORDER_SF_(
        WPP_GLOBAL_Control->DeviceExtension,
        v11,
        11,
        72,
        (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
    }
    *(_DWORD *)(a1 + 308) = 0;
  }
  RtlInitUnicodeString(&DestinationString, L"WinusbIsochUsed");
  if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64))(WdfFunctions_01015 + 1920))(
         WdfDriverGlobals,
         v30,
         &DestinationString,
         a1 + 404) < 0 )
    *(_DWORD *)(a1 + 404) = 0;
  if ( *(_BYTE *)(a1 + 160) )
  {
    RtlInitUnicodeString(&DestinationString, L"SystemWakeEnabled");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           &v29) >= 0 )
    {
      *(_BYTE *)(a1 + 236) = v29 != 0;
    }
    else
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v12) = 4;
        WPP_RECORDER_SF_(
          WPP_GLOBAL_Control->DeviceExtension,
          v12,
          11,
          73,
          (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
      }
      *(_BYTE *)(a1 + 236) = 0;
    }
    RtlInitUnicodeString(&DestinationString, L"DeviceIdleEnabled");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           &v29) >= 0 )
    {
      *(_BYTE *)(a1 + 232) = v29 != 0;
    }
    else
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v13) = 4;
        WPP_RECORDER_SF_(
          WPP_GLOBAL_Control->DeviceExtension,
          v13,
          11,
          74,
          (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
      }
      *(_BYTE *)(a1 + 232) = 0;
    }
    RtlInitUnicodeString(&DestinationString, L"DeviceIdleIgnoreWakeEnable");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           &v29) < 0
      || (v14 = 1, !v29) )
    {
      v14 = 0;
    }
    *(_BYTE *)(a1 + 234) = v14;
    RtlInitUnicodeString(&DestinationString, L"UserSetDeviceIdleEnabled");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           &v29) < 0
      || (v15 = 1, !v29) )
    {
      v15 = 0;
    }
    *(_BYTE *)(a1 + 233) = v15;
    RtlInitUnicodeString(&DestinationString, L"DefaultIdleState");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           &v29) < 0
      || (v16 = 1, !v29) )
    {
      v16 = 0;
    }
    *(_BYTE *)(a1 + 235) = v16;
    RtlInitUnicodeString(&DestinationString, L"DefaultIdleTimeout");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64))(WdfFunctions_01015
                                                                                                + 1920))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           a1 + 228) < 0 )
      *(_DWORD *)(a1 + 228) = 0;
    RtlInitUnicodeString(&DestinationString, L"DevicePowerUpOnS0Entry");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           &v29) >= 0 )
      v17 = v29 != 0;
    else
      v17 = 2;
    *v2 = v17;
    RtlInitUnicodeString(&DestinationString, L"e5b3b5ac-9725-4f78-963f-03dfb1d828c7");
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
           WdfDriverGlobals,
           v30,
           &DestinationString,
           131097LL,
           0LL,
           &v22) < 0
      || (RtlInitUnicodeString(&DestinationString, L"D3ColdSupported"),
          (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
            WdfDriverGlobals,
            v22,
            &DestinationString,
            &v29) < 0) )
    {
      *(_BYTE *)(a1 + 276) = 0;
    }
    else
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v18) = 4;
        WPP_RECORDER_SF_(
          WPP_GLOBAL_Control->DeviceExtension,
          v18,
          11,
          75,
          (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
      }
      *(_BYTE *)(a1 + 276) = v29 != 0;
    }
  }
  RtlInitUnicodeString(&DestinationString, L"WinRtInterfaceRestrictionLevel");
  v5 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, unsigned int *))(WdfFunctions_01015 + 1920))(
         WdfDriverGlobals,
         v30,
         &DestinationString,
         &v29);
  v7 = v5;
  if ( v5 < 0 )
  {
    *(_DWORD *)(a1 + 440) = 255;
    if ( v5 == -1073741772 )
    {
      v7 = 0;
      goto LABEL_69;
    }
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v8 = 78;
LABEL_68:
      LOBYTE(v6) = 2;
      WPP_RECORDER_SF_d(
        WPP_GLOBAL_Control->DeviceExtension,
        v6,
        11,
        v8,
        (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
        v5);
    }
  }
  else
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v6) = 4;
      WPP_RECORDER_SF_d(
        WPP_GLOBAL_Control->DeviceExtension,
        v6,
        11,
        76,
        (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
        v29);
    }
    if ( v29 <= 1 )
    {
      *(_DWORD *)(a1 + 440) = v29;
    }
    else
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v6) = 4;
        WPP_RECORDER_SF_d(
          WPP_GLOBAL_Control->DeviceExtension,
          v6,
          11,
          77,
          (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
          0);
      }
      *(_DWORD *)(a1 + 440) = 0;
    }
  }
LABEL_69:
  if ( v30 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  v19 = v22;
  if ( v22 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(v19) = 5;
    WPP_RECORDER_SF_d(
      WPP_GLOBAL_Control->DeviceExtension,
      v19,
      1,
      79,
      (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
      v7);
  }
  return v7;
}