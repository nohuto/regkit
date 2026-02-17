__int64 __fastcall HUBREG_QueryGlobalUxdSettings(__int64 a1)
{
  int v2; // edi
  int v4; // [rsp+60h] [rbp+20h] BYREF
  __int64 v5; // [rsp+68h] [rbp+28h] BYREF

  v4 = 0;
  v5 = 0LL;
  _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFFFEFF);
  _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFFFDFF);
  _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFFFBFF);
  _InterlockedAnd((volatile signed __int32 *)(a1 + 4), 0xFFFFF7FF);
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015
                                                                                                  + 1832))(
         WdfDriverGlobals,
         0LL,
         &g_UxdGlobalSettingsKey, // \registry\machine\system\currentcontrolset\services\usbhub\uxd_control\policy
         131097LL,
         0LL,
         &v5);
  if ( v2 >= 0 )
  {
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v5,
           L"24", // UxdGlobalDeleteOnShutdown
           4LL,
           &v4,
           0LL,
           0LL) >= 0
      && v4 )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x100u);
    }
    v4 = 0;
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v5,
           L".0", // UxdGlobalDeleteOnReload
           4LL,
           &v4,
           0LL,
           0LL) >= 0
      && v4 )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x200u);
    }
    v4 = 0;
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v5,
           L"68", // UxdGlobalDeleteOnDisconnect
           4LL,
           &v4,
           0LL,
           0LL) >= 0
      && v4 )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x400u);
    }
    v4 = 0;
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v5,
           &g_UxdGlobalEnable, // UxdGlobalEnable
           4LL,
           &v4,
           0LL,
           0LL) >= 0
      && v4 )
    {
      _InterlockedOr((volatile signed __int32 *)(a1 + 4), 0x800u);
    }
    v2 = 0;
  }
  if ( v5 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return (unsigned int)v2;
}