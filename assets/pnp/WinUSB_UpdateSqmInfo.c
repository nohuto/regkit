__int64 __fastcall WinUSB_UpdateSqmInfo(__int64 a1)
{
  __int64 v2; // rax
  __int64 result; // rax
  int v4; // edx
  char v5; // bl
  int v6; // r9d
  int v7; // edx
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-28h] BYREF
  __int64 v9; // [rsp+78h] [rbp+10h] BYREF

  DestinationString = 0LL;
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
    WPP_RECORDER_SF_(
      WPP_GLOBAL_Control->DeviceExtension,
      5,
      1,
      83,
      (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
  v9 = 0LL;
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1632))(WdfDriverGlobals, a1);
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 384))(
             WdfDriverGlobals,
             v2,
             1LL,
             2031616LL,
             0LL,
             &v9);
  v5 = result;
  if ( (int)result < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_11;
    v6 = 84;
    goto LABEL_10;
  }
  RtlInitUnicodeString(&DestinationString, L"WinusbIsochUsed");
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, _QWORD))(WdfFunctions_01015 + 1968))(
             WdfDriverGlobals,
             v9,
             &DestinationString,
             *(unsigned int *)(a1 + 404));
  v5 = result;
  if ( (int)result < 0 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v6 = 85;
LABEL_10:
    LOBYTE(v4) = 2;
    result = WPP_RECORDER_SF_d(
               WPP_GLOBAL_Control->DeviceExtension,
               v4,
               11,
               v6,
               (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
               result);
  }
LABEL_11:
  v7 = v9;
  if ( v9 )
    result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    if ( LOWORD(WPP_GLOBAL_Control->DeviceType) )
    {
      LOBYTE(v7) = 5;
      return WPP_RECORDER_SF_d(
               WPP_GLOBAL_Control->DeviceExtension,
               v7,
               1,
               86,
               (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
               v5);
    }
  }
  return result;
}