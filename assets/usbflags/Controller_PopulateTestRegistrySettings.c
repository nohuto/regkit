__int64 __fastcall Controller_PopulateTestRegistrySettings(__int64 a1)
{
  __int64 result; // rax
  int v3; // edx
  unsigned int v4; // [rsp+50h] [rbp+8h] BYREF
  __int64 v5; // [rsp+58h] [rbp+10h] BYREF

  result = g_WdfDriverUsbXhciContext;
  *(_DWORD *)(a1 + 1284) = 0;
  v5 = 0LL;
  v4 = 0;
  if ( *(_BYTE *)(result + 28) )
  {
    result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, const wchar_t *, __int64, _QWORD, __int64 *))(WdfFunctions_01033 + 1832))(
               WdfDriverGlobals,
               0LL,
               L"vx", // \Registry\Machine\System\CurrentControlSet\Control\usbflags
               131097LL,
               0LL,
               &v5);
    if ( (int)result >= 0 )
    {
      result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, unsigned int *, _QWORD, _QWORD))(WdfFunctions_01033 + 1880))(
                 WdfDriverGlobals,
                 v5,
                 L"(*", // TestRunEsmInWorkItem
                 4LL,
                 &v4,
                 0LL,
                 0LL);
      if ( (int)result >= 0 )
      {
        result = v4;
        if ( v4 )
        {
          if ( v4 == 1 )
            *(_DWORD *)(a1 + 1284) |= 1u;
        }
        else
        {
          *(_DWORD *)(a1 + 1284) &= ~1u;
        }
      }
    }
    else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v3) = 3;
      result = WPP_RECORDER_SF_d(
                 *(_QWORD *)(a1 + 72),
                 v3,
                 4,
                 176,
                 (__int64)&WPP_28df95fb689a379d24090c035b397fa9_Traceguids,
                 result);
    }
    if ( v5 )
      return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01033 + 1848))(WdfDriverGlobals);
  }
  return result;
}