signed int __fastcall HUBREG_UpdateSqmFlags(__int64 a1)
{
  __int64 v2; // rax
  signed int result; // eax
  int v4; // r9d
  char v5; // di
  __int64 v6; // rax
  void *v7; // rdx
  int v8; // r9d
  void *v9; // rdx
  int v10; // [rsp+70h] [rbp+20h] BYREF
  __int64 v11; // [rsp+78h] [rbp+28h] BYREF
  __int64 v12; // [rsp+80h] [rbp+30h] BYREF

  v10 = 0;
  v12 = 0LL;
  v11 = 0LL;
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD))(WdfFunctions_01015 + 1632))(
         WdfDriverGlobals,
         *(_QWORD *)(a1 + 16));
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 384))(
             WdfDriverGlobals,
             v2,
             1LL,
             131103LL,
             0LL,
             &v12);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 119;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, _DWORD, _QWORD, _QWORD, __int64 *))(WdfFunctions_01015 + 1840))(
             WdfDriverGlobals,
             v12,
             L"\b\n",                           // Ceip
             131103LL,
             0,
             0LL,
             0LL,
             &v11);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 120;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, int *))(WdfFunctions_01015 + 1920))(
             WdfDriverGlobals,
             v11,
             L"\"$",                            // DeviceInformation
             &v10);
  v5 = result;
  if ( result >= 0 )
  {
    v8 = v10;
  }
  else
  {
    if ( result != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        v6 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, WDFDRIVER__ *, void *))(WdfFunctions_01015 + 1616))(
               WdfDriverGlobals,
               WdfDriverGlobals->Driver,
               off_1C006A1E8);
        v7 = &WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids;
        LOBYTE(v7) = 2;
        result = WPP_RECORDER_SF_d(
                   *(_QWORD *)(v6 + 64),
                   (_DWORD)v7,
                   2,
                   121,
                   (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
                   v5);
      }
      goto LABEL_42;
    }
    v8 = 0;
  }
  v10 = *(_DWORD *)(a1 + 1640) | 8 | v8;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *))(WdfFunctions_01015 + 1968))(
             WdfDriverGlobals,
             v11,
             L"\"$");                           // DeviceInformation
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 122;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, _QWORD))(WdfFunctions_01015 + 1968))(
             WdfDriverGlobals,
             v11,
             L"(*",                             // PortInterconnectType
             *(unsigned int *)(*(_QWORD *)(a1 + 8) + 216LL));
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 123;
    goto LABEL_41;
  }
  result = RtlNumberOfSetBits((PRTL_BITMAP)(a1 + 2592));
  if ( !result )
    goto LABEL_42;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, __int64))(WdfFunctions_01015 + 1928))(
             WdfDriverGlobals,
             v11,
             L"24",                             // DescriptorValidationInfo0
             4LL,
             4,
             a1 + 2608);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 124;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, __int64))(WdfFunctions_01015 + 1928))(
             WdfDriverGlobals,
             v11,
             L"24",                             // DescriptorValidationInfo1
             4LL,
             4,
             a1 + 2612);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 125;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, __int64))(WdfFunctions_01015 + 1928))(
             WdfDriverGlobals,
             v11,
             L"24",
             4LL,                               // DescriptorValidationInfo2
             4,
             a1 + 2616);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 126;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, __int64))(WdfFunctions_01015 + 1928))(
             WdfDriverGlobals,
             v11,
             L"24",                             // DescriptorValidationInfo3
             4LL,
             4,
             a1 + 2620);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 127;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, __int64))(WdfFunctions_01015 + 1928))(
             WdfDriverGlobals,
             v11,
             L"24",                             // DescriptorValidationInfo4
             4LL,
             4,
             a1 + 2624);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 128;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, __int64))(WdfFunctions_01015 + 1928))(
             WdfDriverGlobals,
             v11,
             L"24",                             // DescriptorValidationInfo5
             4LL,
             4,
             a1 + 2628);
  if ( result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_42;
    v4 = 129;
    goto LABEL_41;
  }
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int, __int64))(WdfFunctions_01015 + 1928))(
             WdfDriverGlobals,
             v11,
             L"24",                             // DescriptorValidationInfo6
             4LL,
             4,
             a1 + 2632);
  if ( result < 0 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v4 = 130;
LABEL_41:
    v9 = &WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids;
    LOBYTE(v9) = 2;
    result = WPP_RECORDER_SF_d(
               *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
               (_DWORD)v9,
               5,
               v4,
               (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
               result);
  }
LABEL_42:
  if ( v11 )
    result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( v12 )
    return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return result;
}