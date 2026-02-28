__int64 __fastcall HUBREG_QueryUsbHardwareVerifierValue(
        __int64 a1,
        __int64 a2,
        __int64 a3,
        __int64 a4,
        __int64 a5,
        __int64 a6,
        _DWORD *a7)
{
  unsigned __int16 v8; // dx
  void *v9; // rsi
  void *v10; // rax
  int v11; // eax
  int v12; // edx
  int v13; // ebx
  __int64 v14; // r13
  NTSTATUS v15; // eax
  int v16; // r9d
  int v17; // esi
  int v18; // eax
  int v19; // eax
  int v20; // eax
  int v21; // eax
  __int64 v22; // r13
  int v23; // eax
  int v24; // eax
  int v25; // eax
  int v26; // eax
  __int64 v28; // [rsp+20h] [rbp-A1h]
  __int64 v29; // [rsp+40h] [rbp-81h] BYREF
  __int64 v30; // [rsp+48h] [rbp-79h] BYREF
  __int64 v31; // [rsp+50h] [rbp-71h]
  __int64 v32; // [rsp+58h] [rbp-69h] BYREF
  __int64 v33; // [rsp+60h] [rbp-61h]
  struct _UNICODE_STRING DestinationString; // [rsp+68h] [rbp-59h] BYREF
  __int64 v35; // [rsp+78h] [rbp-49h]
  __int64 v36; // [rsp+80h] [rbp-41h]
  char v37; // [rsp+88h] [rbp-39h] BYREF

  v33 = a5;
  v35 = a3;
  *a7 = 0;
  v36 = a2;
  v8 = *(_WORD *)(a1 + 2);
  v31 = a6;
  *(_QWORD *)&DestinationString.Length = 3407872LL;
  DestinationString.Buffer = (wchar_t *)&v37;
  v32 = 0LL;
  v30 = 0LL;
  v29 = 0LL;
  if ( v8 )
  {
    if ( v8 > 0x200u )
    {
      v10 = &g_HwVerifierUsb2XName;
      if ( v8 >= 0x300u )
        v10 = &g_HwVerifierUsb30Name;
      v9 = v10;
    }
    else
    {
      v9 = &g_HwVerifierUsbUpto20Name;
    }
  }
  else
  {
    v9 = &g_HwVerifierUsb30Name;
  }
  v11 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
          WdfDriverGlobals,
          0LL,
          &g_HwVerifierKeyName,
          131097LL,
          0LL,
          &v32);
  v12 = 0;
  v13 = v11;
  if ( v11 < 0 )
  {
    v32 = 0LL;
LABEL_43:
    v17 = v31;
    goto LABEL_44;
  }
  v28 = a4;
  v14 = v36;
  v15 = RtlUnicodeStringPrintf(&DestinationString, L"%S%S%S", v36, v35, v28);
  v12 = 0;
  v13 = v15;
  if ( v15 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_43;
    v16 = 15;
LABEL_13:
    v17 = v31;
    LOBYTE(v12) = 2;
    WPP_RECORDER_SF_d(v31, v12, 5, v16, (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids, v15);
    v12 = 0;
LABEL_44:
    *a7 = 0;
    if ( v13 != -1073741772 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v12) = 2;
      WPP_RECORDER_SF_d(v17, v12, 5, 17, (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids, v13);
    }
    goto LABEL_49;
  }
  v18 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
          WdfDriverGlobals,
          v32,
          &DestinationString,
          131097LL,
          0LL,
          &v30);
  v12 = 0;
  v13 = v18;
  if ( v18 >= 0 )
  {
    v19 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
            WdfDriverGlobals,
            v30,
            v9,
            131097LL,
            0LL,
            &v29);
    v12 = 0;
    v13 = v19;
    if ( v19 >= 0 )
    {
      v20 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _DWORD *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
              WdfDriverGlobals,
              v29,
              v33,
              4LL,
              a7,
              0LL,
              0LL);
      v12 = 0;
      v13 = v20;
      if ( v20 >= 0 )
        goto LABEL_47;
    }
    else
    {
      v29 = 0LL;
    }
  }
  else
  {
    v30 = 0LL;
  }
  if ( v13 == -1073741772 )
  {
    if ( v29 )
    {
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
      v29 = 0LL;
    }
    if ( v30 )
    {
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
      v30 = 0LL;
    }
    v15 = RtlUnicodeStringPrintf(&DestinationString, L"%S%S", v14, v35);
    v12 = 0;
    v13 = v15;
    if ( v15 < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_43;
      v16 = 16;
      goto LABEL_13;
    }
    v21 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
            WdfDriverGlobals,
            v32,
            &DestinationString,
            131097LL,
            0LL,
            &v30);
    v12 = 0;
    v13 = v21;
    if ( v21 >= 0 )
    {
      v23 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
              WdfDriverGlobals,
              v30,
              v9,
              131097LL,
              0LL,
              &v29);
      v22 = v33;
      v12 = 0;
      *a7 = 0;
      v13 = v23;
      if ( v23 >= 0 )
      {
        v24 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _DWORD *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
                WdfDriverGlobals,
                v29,
                v22,
                4LL,
                a7,
                0LL,
                0LL);
        v12 = 0;
        v13 = v24;
        if ( v24 >= 0 )
          goto LABEL_47;
      }
      else
      {
        v29 = 0LL;
      }
    }
    else
    {
      v22 = v33;
      v30 = 0LL;
      *a7 = 0;
    }
    if ( v13 != -1073741772 )
      goto LABEL_43;
    if ( v29 )
    {
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
      v29 = 0LL;
    }
    if ( v30 )
    {
      (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
      v30 = 0LL;
    }
    v25 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
            WdfDriverGlobals,
            v32,
            &g_HwVerifierGlobalName,
            131097LL,
            0LL,
            &v30);
    v12 = 0;
    v13 = v25;
    if ( v25 < 0 )
    {
      v30 = 0LL;
      goto LABEL_43;
    }
    v26 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
            WdfDriverGlobals,
            v30,
            v9,
            131097LL,
            0LL,
            &v29);
    v12 = 0;
    v13 = v26;
    if ( v26 < 0 )
    {
      v29 = 0LL;
      goto LABEL_43;
    }
    *a7 = 0;
    v13 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, __int64, __int64, _DWORD *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
            WdfDriverGlobals,
            v29,
            v22,
            4LL,
            a7,
            0LL,
            0LL);
    v12 = 0;
  }
  if ( v13 < 0 )
    goto LABEL_43;
LABEL_47:
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    LOBYTE(v12) = 2;
    WPP_RECORDER_SF_d(v31, v12, 5, 18, (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids, *a7);
  }
LABEL_49:
  if ( v29 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( v30 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( v32 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return (unsigned int)v13;
}