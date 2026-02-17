__int64 __fastcall HUBREG_DeleteUxdDeviceKey(__int64 a1)
{
  NTSTATUS v2; // ebx
  int v3; // edx
  __int64 v5; // [rsp+20h] [rbp-88h]
  __int64 v6; // [rsp+40h] [rbp-68h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+48h] [rbp-60h] BYREF
  __int64 v8; // [rsp+58h] [rbp-50h] BYREF

  *(_QWORD *)&DestinationString.Length = 3407872LL;
  DestinationString.Buffer = (wchar_t *)&v8;
  v6 = 0LL;
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015
                                                                                                  + 1832))(
         WdfDriverGlobals,
         0LL,
         &g_UxdDeviceSettingsKey, // \registry\machine\system\currentcontrolset\services\usbhub\uxd_control\devices
         131097LL,
         0LL,
         &v6);
  if ( v2 >= 0 )
  {
    LODWORD(v5) = *(unsigned __int16 *)(a1 + 2008);
    v2 = RtlUnicodeStringPrintf(
           &DestinationString,
           L"%04X%04X%04X",
           *(unsigned __int16 *)(a1 + 2004),
           *(unsigned __int16 *)(a1 + 2006),
           v5);
    if ( v2 >= 0 )
    {
      v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *))(WdfFunctions_01015 + 1872))(
             WdfDriverGlobals,
             v6,
             &DestinationString);
    }
    else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v3) = 2;
      WPP_RECORDER_SF_d(
        *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
        v3,
        5,
        105,
        (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
        v2);
    }
  }
  if ( v6 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return (unsigned int)v2;
}