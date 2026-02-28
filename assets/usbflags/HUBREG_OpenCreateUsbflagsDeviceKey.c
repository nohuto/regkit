__int64 __fastcall HUBREG_OpenCreateUsbflagsDeviceKey(
        __int64 a1,
        __int64 a2,
        __int64 a3,
        unsigned int a4,
        _QWORD *a5,
        _QWORD *a6,
        char a7,
        __int64 a8)
{
  void *v8; // rsi
  NTSTATUS PersistedStateLocation; // ebx
  __int64 Pool2; // rax
  __int64 v11; // rcx
  const wchar_t *v12; // rax
  __int16 v13; // cx
  unsigned int v14; // r14d
  int v15; // edx
  int v16; // r9d
  unsigned int v18; // [rsp+50h] [rbp-99h] BYREF
  __int64 v19; // [rsp+58h] [rbp-91h] BYREF
  unsigned int v20; // [rsp+60h] [rbp-89h]
  struct _UNICODE_STRING DestinationString; // [rsp+68h] [rbp-81h] BYREF
  struct _UNICODE_STRING v22; // [rsp+78h] [rbp-71h] BYREF
  __int64 v23; // [rsp+88h] [rbp-61h]
  __int64 v24; // [rsp+90h] [rbp-59h]
  __int64 v25; // [rsp+98h] [rbp-51h]
  _QWORD *v26; // [rsp+A0h] [rbp-49h]
  char v27; // [rsp+A8h] [rbp-41h] BYREF

  v24 = a2;
  v25 = a1;
  v20 = a4;
  v23 = a3;
  v26 = a5;
  *(_QWORD *)&v22.Length = 3407872LL;
  v22.Buffer = (wchar_t *)&v27;
  v18 = 0;
  v19 = 0LL;
  DestinationString = 0LL;
  if ( a5 )
    *a5 = 0LL;
  *a6 = 0LL;
  v8 = 0LL;
  if ( a7 != 1 )
  {
    v11 = 0x7FFFLL;
    v12 = L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\usbflags";
    while ( *v12 )
    {
      ++v12;
      if ( !--v11 )
        goto LABEL_17;
    }
    v13 = 2 * v11;
    DestinationString.Buffer = L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\usbflags";
    DestinationString.Length = -2 - v13;
    DestinationString.MaximumLength = -v13;
    goto LABEL_17;
  }
  PersistedStateLocation = RtlGetPersistedStateLocation(
                             L"UsbFlags",
                             0LL,
                             L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\usbflags",
                             0LL,
                             0LL,
                             0,
                             &v18);
  if ( PersistedStateLocation == -2147483643 )
  {
    Pool2 = ExAllocatePool2(64LL, v18, 1681082453LL);
    v8 = (void *)Pool2;
    if ( Pool2 )
    {
      PersistedStateLocation = RtlGetPersistedStateLocation(
                                 L"UsbFlags",
                                 0LL,
                                 L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\usbflags",
                                 0LL,
                                 Pool2,
                                 v18,
                                 0LL);
      if ( PersistedStateLocation < 0 )
      {
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          WPP_RECORDER_SF_d(
            a8,
            2,
            5,
            10,
            (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
            PersistedStateLocation);
LABEL_34:
        ExFreePoolWithTag(v8, 0x64334855u);
        goto LABEL_35;
      }
      RtlUnicodeStringInit(&DestinationString, (NTSTRSAFE_PCWSTR)v8);
    }
LABEL_17:
    v14 = v20;
    PersistedStateLocation = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, struct _UNICODE_STRING *, _QWORD, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
                               WdfDriverGlobals,
                               0LL,
                               &DestinationString,
                               v20,
                               0LL,
                               &v19);
    if ( PersistedStateLocation == -1073741772 )
    {
      if ( a7 != 1 )
      {
LABEL_21:
        if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          goto LABEL_33;
        v16 = 12;
        goto LABEL_32;
      }
      PersistedStateLocation = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, struct _UNICODE_STRING *, _QWORD, _DWORD, _QWORD, _QWORD, __int64 *))(WdfFunctions_01015 + 1840))(
                                 WdfDriverGlobals,
                                 0LL,
                                 &DestinationString,
                                 v14,
                                 0,
                                 0LL,
                                 0LL,
                                 &v19);
    }
    if ( PersistedStateLocation < 0 )
      goto LABEL_21;
    PersistedStateLocation = RtlUnicodeStringPrintf(&v22, L"%S%S%S", v25, v24, v23);
    if ( PersistedStateLocation < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_33;
      v16 = 13;
      goto LABEL_32;
    }
    PersistedStateLocation = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64, _QWORD, _QWORD *))(WdfFunctions_01015 + 1832))(
                               WdfDriverGlobals,
                               v19,
                               &v22,
                               131097LL,
                               0LL,
                               a6);
    if ( PersistedStateLocation == -1073741772 )
    {
      if ( a7 != 1 )
      {
LABEL_30:
        if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          goto LABEL_33;
        v16 = 14;
LABEL_32:
        LOBYTE(v15) = 2;
        WPP_RECORDER_SF_d(
          a8,
          v15,
          5,
          v16,
          (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
          PersistedStateLocation);
LABEL_33:
        if ( !v8 )
          goto LABEL_35;
        goto LABEL_34;
      }
      PersistedStateLocation = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64, _DWORD, _QWORD, _QWORD, _QWORD *))(WdfFunctions_01015 + 1840))(
                                 WdfDriverGlobals,
                                 v19,
                                 &v22,
                                 983103LL,
                                 0,
                                 0LL,
                                 0LL,
                                 a6);
    }
    if ( PersistedStateLocation >= 0 )
      goto LABEL_33;
    goto LABEL_30;
  }
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    WPP_RECORDER_SF_d(a8, 2, 5, 11, (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids, PersistedStateLocation);
LABEL_35:
  if ( PersistedStateLocation >= 0 )
  {
    if ( !v26 )
    {
LABEL_42:
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
      return (unsigned int)PersistedStateLocation;
    }
    *v26 = v19;
  }
  else
  {
    if ( *a6 )
    {
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
      *a6 = 0LL;
    }
    if ( v19 )
      goto LABEL_42;
  }
  return (unsigned int)PersistedStateLocation;
}