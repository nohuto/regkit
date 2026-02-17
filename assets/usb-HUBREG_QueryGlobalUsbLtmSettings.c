__int64 __fastcall HUBREG_QueryGlobalUsbLtmSettings(__int64 a1)
{
  __int64 result; // rax
  int v3; // [rsp+50h] [rbp+8h] BYREF
  __int64 v4; // [rsp+58h] [rbp+10h] BYREF

  v3 = 0;
  v4 = 0LL;
  _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFBFFFF);
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, const wchar_t *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
             WdfDriverGlobals,
             0LL,
             L"z|", // \Registry\Machine\System\CurrentControlSet\Control\usb\UsbLtm',
             131097LL,
             0LL,
             &v4);
  if ( (int)result >= 0 )
  {
    result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
               WdfDriverGlobals,
               v4,
               &g_UsbLtmEnableName,
               4LL,
               &v3,
               0LL,
               0LL);
    if ( (int)result >= 0 )
    {
      if ( v3 )
        _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x40000u);
      else
        _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFBFFFF);
    }
  }
  if ( v4 )
    return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return result;
}