__int64 __fastcall HUBREG_QueryGlobalHubValues(__int64 a1)
{
  __int64 result; // rax
  int v3; // edx
  char v4; // al
  int v5; // r9d
  int v6; // [rsp+80h] [rbp+40h] BYREF
  __int64 v7; // [rsp+88h] [rbp+48h] BYREF

  v6 = 0;
  v7 = 0LL;
  _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x80u);
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
             WdfDriverGlobals,
             0LL,
             &g_HubGlobalKeyName, // \Registry\Machine\System\CurrentControlSet\Services\USBHUB\hubg
             131097LL,
             0LL,
             &v7);
  if ( (int)result < 0 )
    goto LABEL_58;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"24", // DisableSelectiveSuspendUI
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_58;
      v5 = 42;
      goto LABEL_57;
    }
  }
  else if ( v6 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 4), 2u);
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"$&", // "MsOsDescriptorMode
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_58;
      v5 = 44;
      goto LABEL_57;
    }
  }
  else if ( v6 && (unsigned int)(v6 - 1) >= 2 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v3) = 2;
      WPP_RECORDER_SF_d(*(_QWORD *)(a1 + 64), v3, 2, 43, (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids, v6);
    }
  }
  else
  {
    *(_DWORD *)(a1 + 8) = v6;
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"(*", // EnableDiagnosticMode
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_58;
      v5 = 45;
      goto LABEL_57;
    }
  }
  else if ( v6 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 4), 8u);
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"&(", // DisableOnSoftRemove
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_58;
      v5 = 46;
      goto LABEL_57;
    }
  }
  else if ( !v6 )
  {
    _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFFFF7F);
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"\"$", // DisableUxdSupport
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_58;
      v5 = 47;
      goto LABEL_57;
    }
  }
  else if ( v6 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x10u);
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"02", // EnableExtendedValidation
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_58;
      v5 = 48;
      goto LABEL_57;
    }
  }
  else
  {
    v4 = v6;
    if ( v6 )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x20u);
      v4 = v6;
    }
    if ( (v4 & 8) != 0 )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x2000u);
      v4 = v6;
    }
    if ( (v4 & 4) != 0 )
      _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x4000u);
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             &g_WakeOnConnectUI,
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_58;
      v5 = 49;
      goto LABEL_57;
    }
  }
  else if ( v6 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x40u);
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"NP", // PreventDebounceTimeForSuperSpeedDevices
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result == -1073741772 || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_58;
    v5 = 50;
LABEL_57:
    LOBYTE(v3) = 2;
    result = WPP_RECORDER_SF_d(
               *(_QWORD *)(a1 + 64),
               v3,
               2,
               v5,
               (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
               result);
    goto LABEL_58;
  }
  if ( v6 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x10000u);
LABEL_58:
  if ( v7 )
    return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return result;
}