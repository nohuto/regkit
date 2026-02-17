__int64 __fastcall HUBREG_GetUxdPnpValue(__int64 a1, const GUID *a2, __int64 a3)
{
  NTSTATUS v6; // eax
  int v7; // edx
  int v8; // r9d
  __int64 v10; // [rsp+40h] [rbp-38h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+48h] [rbp-30h] BYREF

  v10 = 0LL;
  DestinationString = 0LL;
  RtlInitUnicodeString(&DestinationString, 0LL);
  v6 = RtlStringFromGUID(a2, &DestinationString);
  if ( v6 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_11;
    v8 = 106;
    goto LABEL_10;
  }
  v6 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015
                                                                                                  + 1832))(
         WdfDriverGlobals,
         0LL,
         &g_UxdGuidSettingsKey, // \registry\machine\system\currentcontrolset\services\usbhub\uxd_control\pnp
         131097LL,
         0LL,
         &v10);
  if ( v6 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_11;
    v8 = 107;
    goto LABEL_10;
  }
  v6 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64))(WdfFunctions_01015
                                                                                                  + 1912))(
         WdfDriverGlobals,
         v10,
         &DestinationString,
         a3);
  if ( v6 < 0 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    v8 = 108;
LABEL_10:
    LOBYTE(v7) = 2;
    WPP_RECORDER_SF_d(
      *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
      v7,
      5,
      v8,
      (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
      v6);
  }
LABEL_11:
  RtlFreeUnicodeString(&DestinationString);
  if ( v10 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return 0LL;
}