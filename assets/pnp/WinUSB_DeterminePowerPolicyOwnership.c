__int64 __fastcall WinUSB_DeterminePowerPolicyOwnership(__int64 a1, _BYTE *a2)
{
  _BYTE *v2; // rdi
  __int64 v4; // rax
  PWDF_DRIVER_GLOBALS v5; // rcx
  int v6; // eax
  int v7; // edx
  int v8; // ebx
  int v9; // edx
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-38h] BYREF
  int v12; // [rsp+88h] [rbp+10h] BYREF
  __int64 v13; // [rsp+90h] [rbp+18h] BYREF

  DestinationString = 0LL;
  v12 = 0;
  v2 = a2;
  v13 = 0LL;
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(a2) = 5;
    WPP_RECORDER_SF_(
      WPP_GLOBAL_Control->DeviceExtension,
      (_DWORD)a2,
      1,
      80,
      (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids);
  }
  v4 = WdfFunctions_01015;
  v5 = WdfDriverGlobals;
  *v2 = 1;
  v6 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(v4 + 1000))(
         v5,
         a1,
         1LL,
         0x80000000LL,
         0LL,
         &v13);
  v8 = v6;
  if ( v6 >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"WinUsbPowerPolicyOwnershipDisabled");
    v8 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, int *))(WdfFunctions_01015
                                                                                                  + 1920))(
           WdfDriverGlobals,
           v13,
           &DestinationString,
           &v12);
    if ( v8 >= 0 )
    {
      if ( v12 )
        *v2 = 0;
    }
    else
    {
      v8 = 0;
    }
    if ( !*v2 )
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, _QWORD))(WdfFunctions_01015 + 456))(
        WdfDriverGlobals,
        a1,
        0LL);
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v7) = 2;
    WPP_RECORDER_SF_d(
      WPP_GLOBAL_Control->DeviceExtension,
      v7,
      11,
      81,
      (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
      v6);
  }
  v9 = v13;
  if ( v13 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(v9) = 5;
    WPP_RECORDER_SF_d(
      WPP_GLOBAL_Control->DeviceExtension,
      v9,
      1,
      82,
      (__int64)&WPP_26f1e60b7ed939401ac6450f7fef2921_Traceguids,
      v8);
  }
  return (unsigned int)v8;
}