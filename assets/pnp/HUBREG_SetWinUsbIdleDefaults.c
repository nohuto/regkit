__int64 __fastcall HUBREG_SetWinUsbIdleDefaults(__int64 a1)
{
  __int64 v2; // rdi
  __int64 result; // rax
  int v4; // edx
  int v5; // r9d
  int v6; // r8d
  int v7; // ebx
  int v8; // ebx
  char v9; // [rsp+28h] [rbp-18h]
  int v10; // [rsp+68h] [rbp+28h] BYREF
  __int64 v11; // [rsp+70h] [rbp+30h] BYREF
  __int64 v12; // [rsp+78h] [rbp+38h] BYREF

  v10 = 0;
  v12 = 0LL;
  v11 = 0LL;
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *))(WdfFunctions_01015 + 1616))(
         WdfDriverGlobals,
         a1,
         off_1C006A170);
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 384))(
             WdfDriverGlobals,
             a1,
             1LL,
             131103LL,
             0LL,
             &v11);
  if ( (int)result >= 0 )
  {
    result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, _QWORD, __int64 *))(WdfFunctions_01015 + 2464))(
               WdfDriverGlobals,
               0LL,
               0LL,
               &v12);
    if ( (int)result >= 0 )
    {
      v7 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64))(WdfFunctions_01015 + 1912))(
             WdfDriverGlobals,
             v11,
             L"&(",                             // DeviceInterfaceGUID
             v12);
      result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1664))(
                 WdfDriverGlobals,
                 v12);
      if ( v7 == -1073741772 )
      {
        result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, const wchar_t *, _QWORD, __int64 *))(WdfFunctions_01015 + 2464))(
                   WdfDriverGlobals,
                   L"LN",                       // {52783fc2-0179-4eca-bb46-128bba61975e}
                   0LL,
                   &v12);
        if ( (int)result < 0 )
        {
          if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
            goto LABEL_33;
          v5 = 149;
          goto LABEL_4;
        }
        v8 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64))(WdfFunctions_01015
                                                                                               + 1960))(
               WdfDriverGlobals,
               v11,
               L"&(",                           // DeviceInterfaceGUID
               v12);
        result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1664))(
                   WdfDriverGlobals,
                   v12);
        if ( v8 < 0 )
        {
          if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
            goto LABEL_33;
          v5 = 150;
          v9 = v8;
          v6 = 3;
          goto LABEL_32;
        }
      }
      else if ( v7 < 0 )
      {
        if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          goto LABEL_33;
        v5 = 151;
        v9 = v7;
        v6 = 5;
        goto LABEL_32;
      }
      result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
                 WdfDriverGlobals,
                 v11,
                 L"\"$",                        // DeviceIdleEnabled
                 4LL,
                 &v10,
                 0LL,
                 0LL);
      if ( (int)result >= 0 )
        goto LABEL_33;
      result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
                 WdfDriverGlobals,
                 v11,
                 L" \"",                        // DefaultIdleState
                 4LL,
                 &v10,
                 0LL,
                 0LL);
      if ( (int)result >= 0 )
        goto LABEL_33;
      result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
                 WdfDriverGlobals,
                 v11,
                 L"46",                         // DeviceIdleIgnoreWakeEnable
                 4LL,
                 &v10,
                 0LL,
                 0LL);
      if ( (int)result >= 0 )
        goto LABEL_33;
      v10 = 1;
      result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, int *))(WdfFunctions_01015 + 1928))(
                 WdfDriverGlobals,
                 v11,
                 L"\"$",                        // DeviceIdleEnabled
                 4LL,
                 4,
                 &v10);
      if ( (int)result >= 0 )
      {
        v10 = 1;
        result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, int *))(WdfFunctions_01015 + 1928))(
                   WdfDriverGlobals,
                   v11,
                   L" \"",                      // DefaultIdleState
                   4LL,
                   4,
                   &v10);
        if ( (int)result >= 0 )
        {
          v10 = 1;
          result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, int *))(WdfFunctions_01015 + 1928))(
                     WdfDriverGlobals,
                     v11,
                     L"46",                     // DeviceIdleIgnoreWakeEnable
                     4LL,
                     4,
                     &v10);
          if ( (int)result >= 0 || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
            goto LABEL_33;
          v5 = 154;
        }
        else
        {
          if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
            goto LABEL_33;
          v5 = 153;
        }
      }
      else
      {
        if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          goto LABEL_33;
        v5 = 152;
      }
      v6 = 5;
      goto LABEL_31;
    }
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_33;
    v5 = 148;
  }
  else
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_33;
    v5 = 147;
  }
LABEL_4:
  v6 = 3;
LABEL_31:
  v9 = result;
LABEL_32:
  LOBYTE(v4) = 2;
  result = WPP_RECORDER_SF_d(
             *(_QWORD *)(*(_QWORD *)v2 + 2520LL),
             v4,
             v6,
             v5,
             (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
             v9);
LABEL_33:
  if ( v11 )
    return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return result;
}