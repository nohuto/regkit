__int64 __fastcall HUBREG_QueryExtPropDescSemaphoreInDeviceHardwareKey(__int64 a1)
{
  __int64 v2; // rax
  __int64 result; // rax
  int v4; // edx
  int v5; // r9d
  int v6; // eax
  int v7; // edx
  int v8; // eax
  int v9; // edx
  int v10; // [rsp+70h] [rbp+30h] BYREF
  int v11; // [rsp+78h] [rbp+38h] BYREF
  int v12; // [rsp+80h] [rbp+40h] BYREF
  __int64 v13; // [rsp+88h] [rbp+48h] BYREF

  v12 = 0;
  v10 = 0;
  v11 = 0;
  v13 = 0LL;
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD))(WdfFunctions_01015 + 1632))(
         WdfDriverGlobals,
         *(_QWORD *)(a1 + 16));
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 384))(
             WdfDriverGlobals,
             v2,
             1LL,
             131097LL,
             0LL,
             &v13);
  if ( (int)result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_23;
    v5 = 81;
    goto LABEL_22;
  }
  v6 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
         WdfDriverGlobals,
         v13,
         &g_RevisionId,
         4LL,
         &v10,
         0LL,
         0LL);
  if ( (int)(v6 + 0x80000000) >= 0
    && v6 != -1073741772
    && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v7) = 2;
    WPP_RECORDER_SF_d(
      *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
      v7,
      5,
      82,
      (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
      v6);
  }
  v8 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
         WdfDriverGlobals,
         v13,
         &g_VendorRevision,
         4LL,
         &v11,
         0LL,
         0LL);
  if ( ((v8 + 0x80000000) & 0x80000000) == 0
    && v8 != -1073741772
    && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v9) = 2;
    WPP_RECORDER_SF_d(
      *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
      v9,
      5,
      83,
      (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
      v8);
  }
  _InterlockedAnd((volatile signed __int32 *)(a1 + 1632), 0xFFFFFDFF);
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
             WdfDriverGlobals,
             v13,
             L"(*", // ExtPropDescSemaphore
             4LL,
             &v12,
             0LL,
             0LL);
  if ( (int)result < 0 )
  {
    if ( (_DWORD)result == -1073741772 || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_23;
    v5 = 84;
LABEL_22:
    LOBYTE(v4) = 2;
    result = WPP_RECORDER_SF_d(
               *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
               v4,
               5,
               v5,
               (__int64)&WPP_7a0afab5c79d3741c23ff4ee70090e0b_Traceguids,
               result);
    goto LABEL_23;
  }
  result = *(unsigned __int16 *)(a1 + 2000);
  if ( v10 == (_DWORD)result )
  {
    if ( (result = *(_DWORD *)(a1 + 2464) & 0x400, (*(_DWORD *)(a1 + 2464) & 0x400) == 0) && !v11
      || (_DWORD)result && (result = *(_QWORD *)(a1 + 2528), v11 == *(unsigned __int16 *)(result + 4)) )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 1632), 0x200u);
    }
  }
LABEL_23:
  if ( v13 )
    return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return result;
}