__int64 __fastcall HUBREG_QueryValuesInDeviceHardwareKey(__int64 a1)
{
  __int64 v1; // rdx
  __int64 v3; // rax
  int v4; // eax
  int v5; // edx
  unsigned int v6; // ebx
  int v7; // r9d
  void *Pool2; // rax
  int v9; // edx
  void *v10; // rdi
  __int64 v12; // [rsp+40h] [rbp-20h] BYREF
  void *Src[2]; // [rsp+48h] [rbp-18h] BYREF
  int v14; // [rsp+A0h] [rbp+40h] BYREF
  __int64 v15; // [rsp+A8h] [rbp+48h] BYREF
  __int64 v16; // [rsp+B0h] [rbp+50h] BYREF
  __int64 v17; // [rsp+B8h] [rbp+58h] BYREF

  v1 = *(_QWORD *)(a1 + 16);
  v14 = 0;
  *(_OWORD *)Src = 0LL;
  v16 = 0LL;
  v15 = 0LL;
  v12 = 0LL;
  v17 = 0LL;
  v3 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1632))(WdfDriverGlobals, v1);
  v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 384))(
         WdfDriverGlobals,
         v3,
         1LL,
         131097LL,
         0LL,
         &v15);
  v6 = v4;
  if ( v4 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_50;
    v7 = 85;
    goto LABEL_4;
  }
  v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
         WdfDriverGlobals,
         v15,
         L"02",                                 // DeviceSelectiveSuspended
         4LL,
         &v14,
         0LL,
         0LL);
  v6 = v4;
  if ( v4 < 0 )
  {
    if ( v4 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        v7 = 86;
        goto LABEL_4;
      }
      goto LABEL_50;
    }
  }
  else if ( v14 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 1632), 0x400u);
  }
  v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, _QWORD, __int64 *))(WdfFunctions_01015 + 2464))(
         WdfDriverGlobals,
         0LL,
         0LL,
         &v17);
  v6 = v4;
  if ( v4 >= 0 )
  {
    v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64))(WdfFunctions_01015 + 1912))(
           WdfDriverGlobals,
           v15,
           &g_FriendlyName,                     // FriendlyName
           v17);
    v6 = v4;
    if ( v4 < 0 )
    {
      if ( v4 != -1073741772 )
      {
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          v7 = 89;
          goto LABEL_4;
        }
        goto LABEL_50;
      }
    }
    else
    {
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void **))(WdfFunctions_01015 + 2472))(
        WdfDriverGlobals,
        v17,
        Src);
      Pool2 = (void *)ExAllocatePool2(64LL, LOWORD(Src[0]), 1681082453LL);
      v10 = Pool2;
      if ( !Pool2 )
      {
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          LOBYTE(v9) = 2;
          WPP_RECORDER_SF_(
            *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
            v9,
            5,
            88,
            (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids);
        }
        goto LABEL_50;
      }
      memmove(Pool2, Src[1], LOWORD(Src[0]));
      *(_DWORD *)(a1 + 2164) = LOWORD(Src[0]);
      *(_QWORD *)(a1 + 2168) = v10;
    }
    v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
           WdfDriverGlobals,
           v15,
           L"HJ",                               // e5b3b5ac-9725-4f78-963f-03dfb1d828c7
           131097LL,
           0LL,
           &v12);
    v6 = v4;
    if ( v4 >= 0 )
    {
      v14 = 0;
      v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v12,
             &g_D3ColdSupported,                // D3ColdSupported
             4LL,
             &v14,
             0LL,
             0LL);
      v6 = v4;
      if ( v4 < 0 )
      {
        if ( v4 != -1073741772 )
        {
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          {
            v7 = 92;
            goto LABEL_4;
          }
          goto LABEL_50;
        }
      }
      else if ( v14 )
      {
        _InterlockedOr((volatile signed __int32 *)(a1 + 1636), 0x1000u);
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          LOBYTE(v5) = 4;
          WPP_RECORDER_SF_(
            *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
            v5,
            5,
            91,
            (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids);
        }
      }
    }
    else if ( v4 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        v7 = 90;
        goto LABEL_4;
      }
      goto LABEL_50;
    }
    v14 = 0;
    v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v15,
           L" \"",                              // AllowIdleIrpInD3
           4LL,
           &v14,
           0LL,
           0LL);
    v6 = v4;
    if ( v4 < 0 )
    {
      if ( v4 != -1073741772 )
      {
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          v7 = 93;
          goto LABEL_4;
        }
        goto LABEL_50;
      }
    }
    else if ( v14 )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 1632), 0x4000u);
    }
    *(_DWORD *)(*(_QWORD *)(a1 + 8) + 1440LL) = 1000;
    v14 = 0;
    v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v15,
           L",.",                               // D3ColdReconnectTimeout
           4LL,
           &v14,
           0LL,
           0LL);
    v6 = v4;
    if ( v4 < 0 )
    {
      if ( v4 != -1073741772 )
      {
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          v7 = 94;
          goto LABEL_4;
        }
        goto LABEL_50;
      }
    }
    else
    {
      *(_DWORD *)(*(_QWORD *)(a1 + 8) + 1440LL) = v14;
    }
    v4 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 104))(
           WdfDriverGlobals,
           0LL,
           &v16);
    v6 = v4;
    if ( v4 >= 0 )
    {
      if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, _QWORD, __int64))(WdfFunctions_01015
                                                                                                 + 1896))(
             WdfDriverGlobals,
             v15,
             L"$&",                             // EndpointPriorities
             0LL,
             v16) >= 0 )
        HUBREG_ValidateAndPopulateEndpointPriorities(a1, v16);
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1664))(WdfDriverGlobals, v16);
      v6 = 0;
    }
    else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v7 = 95;
      goto LABEL_4;
    }
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v7 = 87;
LABEL_4:
    LOBYTE(v5) = 2;
    WPP_RECORDER_SF_d(
      *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
      v5,
      5,
      v7,
      (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
      v4);
  }
LABEL_50:
  if ( v12 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( v15 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( v17 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1664))(WdfDriverGlobals);
  return v6;
}