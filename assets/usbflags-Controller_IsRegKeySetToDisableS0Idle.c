bool __fastcall Controller_IsRegKeySetToDisableS0Idle(__int64 a1)
{
  int v2; // eax
  int v3; // edx
  int v5; // [rsp+58h] [rbp+10h] BYREF
  __int64 v6; // [rsp+60h] [rbp+18h] BYREF

  v6 = 0LL;
  v5 = 0;
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, const wchar_t *, __int64, _QWORD, __int64 *))(WdfFunctions_01033 + 1832))(
         WdfDriverGlobals,
         0LL,
         L"vx", // \Registry\Machine\System\CurrentControlSet\Control\usbflags
         131097LL,
         0LL,
         &v6);
  if ( v2 >= 0 )
  {
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01033 + 1880))(
           WdfDriverGlobals,
           v6,
           &g_DisableHCS0Idle,
           4LL,
           &v5,
           0LL,
           0LL) < 0 )
      v5 = 0;
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v3) = 3;
    WPP_RECORDER_SF_d(*(_QWORD *)(a1 + 72), v3, 4, 22, (__int64)&WPP_28df95fb689a379d24090c035b397fa9_Traceguids, v2);
  }
  if ( v6 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01033 + 1848))(WdfDriverGlobals);
  return v5 != 0;
}