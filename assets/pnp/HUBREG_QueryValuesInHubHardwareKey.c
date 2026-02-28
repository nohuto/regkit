__int64 __fastcall HUBREG_QueryValuesInHubHardwareKey(__int64 a1)
{
  __int64 v1; // rdx
  __int64 result; // rax
  int v4; // edx
  int v5; // r9d
  unsigned int v6; // [rsp+60h] [rbp+20h] BYREF
  __int64 v7; // [rsp+68h] [rbp+28h] BYREF

  v1 = *(_QWORD *)(a1 + 16);
  v6 = 0;
  v7 = 0LL;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 384))(
             WdfDriverGlobals,
             v1,
             1LL,
             131097LL,
             0LL,
             &v7);
  if ( (int)result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_26;
    v5 = 51;
    goto LABEL_25;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, unsigned int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"&(",                             // WakeSystemOnConnect
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_26;
      v5 = 52;
      goto LABEL_25;
    }
  }
  else if ( v6 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x100u);
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, unsigned int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             &g_HubInstHardResetCount,          // HardResetCount
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_26;
      v5 = 53;
      goto LABEL_25;
    }
  }
  else
  {
    *(_DWORD *)(a1 + 56) = v6;
  }
  v6 = 0;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, unsigned int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v7,
             L"&(",                             // OvercurrentDetected
             4LL,
             &v6,
             0LL,
             0LL);
  if ( (int)result >= 0 )
  {
    *(_DWORD *)(a1 + 40) ^= (*(_DWORD *)(a1 + 40) ^ (v6 << 29)) & 0x20000000;
LABEL_11:
    v6 = 0;
    result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, unsigned int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
               WdfDriverGlobals,
               v7,
               L"&(",                           // HubFWUpdateProtocol
               4LL,
               &v6,
               0LL,
               0LL);
    if ( (int)result >= 0 )
    {
      result = v6;
      *(_DWORD *)(a1 + 160) = v6;
      goto LABEL_26;
    }
    if ( (_DWORD)result == -1073741772 || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_26;
    v5 = 55;
    goto LABEL_25;
  }
  if ( (_DWORD)result == -1073741772 )
    goto LABEL_11;
  if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    goto LABEL_26;
  v5 = 54;
LABEL_25:
  LOBYTE(v4) = 2;
  result = WPP_RECORDER_SF_d(
             *(_QWORD *)(a1 + 2520),
             v4,
             3,
             v5,
             (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
             result);
LABEL_26:
  if ( v7 )
    return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return result;
}