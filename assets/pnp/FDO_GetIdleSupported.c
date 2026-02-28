char __fastcall FDO_GetIdleSupported(__int64 a1) // vhf.sys
{
  char v1; // di
  char v2; // bl
  int v3; // eax
  int v4; // edx
  int v5; // r8d
  int v6; // r9d
  int v7; // edx
  WDFDRIVER Driver; // rdx
  int v10; // [rsp+48h] [rbp-9h] BYREF
  __int64 SystemInformation; // [rsp+50h] [rbp-1h] BYREF
  __int64 v12; // [rsp+58h] [rbp+7h] BYREF
  _QWORD v13[2]; // [rsp+60h] [rbp+Fh] BYREF
  __int128 v14; // [rsp+70h] [rbp+1Fh] BYREF
  __int64 v15; // [rsp+80h] [rbp+2Fh]
  int v16; // [rsp+88h] [rbp+37h]

  v1 = a1;
  v10 = 0;
  SystemInformation = 0LL;
  v12 = 0LL;
  v16 = *(_DWORD *)L"d";
  v14 = *(_OWORD *)L"IdleSupported";
  v2 = 0;
  v13[1] = &v14;
  v15 = *(_QWORD *)L"orted";
  v13[0] = 1835034LL;
  v3 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1000))(
         WdfDriverGlobals,
         a1,
         1LL,
         131097LL,
         0LL,
         &v12);
  if ( v3 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_14;
    v6 = 44;
    LOBYTE(v4) = 2;
    goto LABEL_4;
  }
  v3 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, _QWORD *, int *))(WdfFunctions_01015 + 1920))(
         WdfDriverGlobals,
         v12,
         v13,
         &v10);
  if ( v3 >= 0 )
  {
    if ( v10 )
    {
      LODWORD(SystemInformation) = 8;
      v2 = 1;
      if ( ZwQuerySystemInformation(MaxSystemInfoClass|SystemProcessInformation, &SystemInformation, 8u, 0LL) < 0
        || (SystemInformation & 0x200000000LL) == 0 )
      {
        v2 = 0;
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          LOBYTE(v7) = 3;
          WPP_RECORDER_SF_qq(
            WPP_GLOBAL_Control->DeviceExtension,
            v7,
            v5,
            46,
            (__int64)&WPP_7239fe49e07c3f43551e44488ee0f0b8_Traceguids,
            (char)WdfDriverGlobals->Driver,
            v1);
        }
      }
    }
  }
  else if ( v3 != -1073741772 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v6 = 45;
    LOBYTE(v4) = 3;
LABEL_4:
    WPP_RECORDER_SF_qd(
      WPP_GLOBAL_Control->DeviceExtension,
      v4,
      v5,
      v6,
      (__int64)&WPP_7239fe49e07c3f43551e44488ee0f0b8_Traceguids,
      (char)WdfDriverGlobals->Driver,
      v3);
  }
LABEL_14:
  if ( v12 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    Driver = WdfDriverGlobals->Driver;
    LOBYTE(Driver) = 4;
    WPP_RECORDER_SF_qqd(
      WPP_GLOBAL_Control->DeviceExtension,
      (_DWORD)Driver,
      v5,
      47,
      (__int64)&WPP_7239fe49e07c3f43551e44488ee0f0b8_Traceguids,
      (char)WdfDriverGlobals->Driver,
      v1,
      v2);
  }
  return v2;
}
