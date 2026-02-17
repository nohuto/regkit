__int64 __fastcall HUBREG_QueryGlobalUsb20HardwareLpmSettings(__int64 a1)
{
  __int64 result; // rax
  unsigned int v3; // [rsp+50h] [rbp+8h] BYREF
  __int64 v4; // [rsp+58h] [rbp+10h] BYREF

  v3 = 0;
  v4 = 0LL;
  _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x8000u);
  *(_BYTE *)(a1 + 72) = 2;
  result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
             WdfDriverGlobals,
             0LL,
             &g_Usb20HardwareLpmKeyName, // \Registry\Machine\System\CurrentControlSet\Control\usb\Usb20HardwareLpm
             131097LL,
             0LL,
             &v4);
  if ( (int)result >= 0 )
  {
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, unsigned int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v4,
           L"02", // Usb20HardwareLpmOverride
           4LL,
           &v3,
           0LL,
           0LL) >= 0 )
    {
      if ( v3 )
        _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x8000u);
      else
        _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFF7FFF);
    }
    v3 = 0;
    result = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, unsigned int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
               WdfDriverGlobals,
               v4,
               L".0", // Usb20HardwareLpmTimeout
               4LL,
               &v3,
               0LL,
               0LL);
    if ( (int)result >= 0 )
    {
      result = v3;
      if ( v3 <= 0xFF )
        *(_BYTE *)(a1 + 72) = v3;
    }
  }
  if ( v4 )
    return (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return result;
}