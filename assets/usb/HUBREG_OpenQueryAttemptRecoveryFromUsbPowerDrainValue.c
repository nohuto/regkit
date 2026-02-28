__int64 __fastcall HUBREG_OpenQueryAttemptRecoveryFromUsbPowerDrainValue(_DWORD *a1)
{
  __int64 v2; // rsi
  wchar_t *Pool2; // r14
  char v4; // r13
  char v5; // r12
  int v6; // edx
  int v7; // ebx
  int v8; // eax
  int v9; // edx
  int v10; // edx
  NTSTATUS PersistedStateLocation; // eax
  int v12; // r9d
  char v14; // [rsp+28h] [rbp-28h]
  struct _UNICODE_STRING DestinationString; // [rsp+40h] [rbp-10h] BYREF
  unsigned int v16; // [rsp+98h] [rbp+48h] BYREF
  __int64 v17; // [rsp+A0h] [rbp+50h] BYREF
  __int64 v18; // [rsp+A8h] [rbp+58h] BYREF

  DestinationString = 0LL;
  v16 = 0;
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, WDFDRIVER__ *, void *))(WdfFunctions_01015 + 1616))(
         WdfDriverGlobals,
         WdfDriverGlobals->Driver,
         off_14006D308);
  v17 = 0LL;
  v18 = 0LL;
  Pool2 = 0LL;
  v4 = 0;
  v5 = 0;
  v7 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015
                                                                                                  + 1832))(
         WdfDriverGlobals,
         0LL,
         &g_UsbAutomaticSurpriseRemovalKeyName, // \Registry\Machine\System\CurrentControlSet\Control\usb\AutomaticSurpriseRemoval
         131097LL,
         0LL,
         &v17);
  if ( v7 >= 0 )
  {
    v8 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, _DWORD *))(WdfFunctions_01015 + 1920))(
           WdfDriverGlobals,
           v17,
           L"@B", // AttemptRecoveryFromUsbPowerDrain
           a1);
    v7 = v8;
    if ( v8 >= 0 )
    {
      v4 = 1;
    }
    else
    {
      *a1 = 0;
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v9) = 2;
        WPP_RECORDER_SF_d(
          *(_QWORD *)(v2 + 64),
          v9,
          2,
          138,
          (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
          v8);
      }
    }
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1848))(WdfDriverGlobals, v17);
    v17 = 0LL;
  }
  else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v6) = 2;
    WPP_RECORDER_SF_d(*(_QWORD *)(v2 + 64), v6, 2, 137, (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids, v7);
  }
  if ( (unsigned __int8)RtlIsStateSeparationEnabled() == 1 )
  {
    PersistedStateLocation = RtlGetPersistedStateLocation(
                               L"USB",
                               0LL,
                               L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\usb",
                               0LL,
                               0LL,
                               0,
                               &v16);
    v7 = PersistedStateLocation;
    if ( PersistedStateLocation == -2147483643 )
    {
      Pool2 = (wchar_t *)ExAllocatePool2(64LL, v16, 1681082453LL);
      if ( Pool2 )
      {
        PersistedStateLocation = RtlGetPersistedStateLocation(
                                   L"USB",
                                   0LL,
                                   L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\usb",
                                   0LL,
                                   Pool2,
                                   v16,
                                   0LL);
        v7 = PersistedStateLocation;
        if ( PersistedStateLocation >= 0 )
        {
          PersistedStateLocation = RtlUnicodeStringInit(&DestinationString, Pool2);
          v7 = PersistedStateLocation;
          if ( PersistedStateLocation < 0 )
          {
            if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
              goto LABEL_34;
            v12 = 142;
            goto LABEL_19;
          }
          PersistedStateLocation = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, struct _UNICODE_STRING *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
                                     WdfDriverGlobals,
                                     0LL,
                                     &DestinationString,
                                     131097LL,
                                     0LL,
                                     &v18);
          v7 = PersistedStateLocation;
          if ( PersistedStateLocation < 0 )
          {
            if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
              goto LABEL_34;
            v12 = 143;
            goto LABEL_19;
          }
          PersistedStateLocation = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
                                     WdfDriverGlobals,
                                     v18,
                                     L"02", // AutomaticSurpriseRemova
                                     131097LL,
                                     0LL,
                                     &v17);
          v7 = PersistedStateLocation;
          if ( PersistedStateLocation < 0 )
          {
            if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
              goto LABEL_34;
            v12 = 144;
            goto LABEL_19;
          }
          PersistedStateLocation = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, _DWORD *))(WdfFunctions_01015 + 1920))(
                                     WdfDriverGlobals,
                                     v17,
                                     L"@B", // AttemptRecoveryFromUsbPowerDrain
                                     a1);
          v7 = PersistedStateLocation;
          if ( PersistedStateLocation >= 0 )
          {
            v5 = 1;
            goto LABEL_34;
          }
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          {
            v12 = 145;
            goto LABEL_19;
          }
        }
        else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          v12 = 141;
          goto LABEL_19;
        }
      }
      else
      {
        v7 = -1073741670;
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          v12 = 140;
          v14 = -102;
LABEL_20:
          LOBYTE(v10) = 2;
          WPP_RECORDER_SF_d(
            *(_QWORD *)(v2 + 64),
            v10,
            2,
            v12,
            (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
            v14);
        }
      }
    }
    else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v12 = 139;
LABEL_19:
      v14 = PersistedStateLocation;
      goto LABEL_20;
    }
  }
LABEL_34:
  if ( v4 && !v5 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v10) = 3;
      WPP_RECORDER_SF_d(
        *(_QWORD *)(v2 + 64),
        v10,
        2,
        146,
        (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
        v7);
    }
    v7 = 0;
  }
  if ( Pool2 )
    ExFreePoolWithTag(Pool2, 0x64334855u);
  if ( v18 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( v17 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return (unsigned int)v7;
}