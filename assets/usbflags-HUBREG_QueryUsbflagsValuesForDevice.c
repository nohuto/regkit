__int64 __fastcall HUBREG_QueryUsbflagsValuesForDevice(volatile signed __int32 *a1, __int64 a2, __int64 a3, __int64 a4)
{
  __int64 v4; // rbx
  __int64 v9; // rax
  int v10; // eax
  __int64 v11; // r13
  NTSTATUS v12; // esi
  __int64 v13; // r10
  __int64 v14; // rbx
  int v15; // edx
  int v16; // r9d
  bool v17; // zf
  bool v18; // zf
  bool v19; // zf
  char v20; // al
  __int64 v22; // [rsp+38h] [rbp-C8h]
  int v23; // [rsp+80h] [rbp-80h] BYREF
  char v24; // [rsp+84h] [rbp-7Ch]
  __int64 v25; // [rsp+88h] [rbp-78h] BYREF
  char pszDest[8]; // [rsp+90h] [rbp-70h] BYREF
  __int64 v27; // [rsp+98h] [rbp-68h] BYREF
  __int64 v28; // [rsp+A0h] [rbp-60h] BYREF
  __int64 v29; // [rsp+A8h] [rbp-58h] BYREF
  __int64 v30; // [rsp+B0h] [rbp-50h] BYREF
  __int64 v31; // [rsp+B8h] [rbp-48h] BYREF
  __int64 v32; // [rsp+C0h] [rbp-40h] BYREF
  __int64 v33; // [rsp+C8h] [rbp-38h] BYREF
  __int64 v34; // [rsp+D0h] [rbp-30h] BYREF
  __int64 v35; // [rsp+D8h] [rbp-28h] BYREF
  __int64 v36; // [rsp+E0h] [rbp-20h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+E8h] [rbp-18h] BYREF
  char v38; // [rsp+100h] [rbp+0h] BYREF

  v4 = *(_QWORD *)a1;
  v23 = 0;
  v34 = 0LL;
  v31 = 0LL;
  v32 = 0LL;
  v28 = 0LL;
  v29 = 0LL;
  v30 = 0LL;
  v33 = 0LL;
  v35 = 0LL;
  v24 = *(_BYTE *)(v4 + 200);
  DestinationString.Buffer = (wchar_t *)&v38;
  v9 = *((_QWORD *)a1 + 1);
  *(_QWORD *)&DestinationString.Length = 6291456LL;
  v25 = 0LL;
  v27 = 0LL;
  v22 = *(_QWORD *)(v9 + 1432);
  v36 = 0LL;
  HUBREG_OpenCreateUsbflagsDeviceKey(a2, a3, a4, 0x20019u, &v36, &v25, 0, v22);
  v10 = HUBREG_OpenCreateUsbflagsDeviceKey(
          a2,
          a3,
          a4,
          0x20019u,
          0LL,
          &v27,
          1,
          *(_QWORD *)(*((_QWORD *)a1 + 1) + 1432LL));
  v11 = v36;
  v12 = v10;
  if ( v10 < 0 )
    goto LABEL_157;
  RtlStringCchPrintfA(pszDest, 3uLL, "%02X", *((unsigned __int8 *)a1 + 2000));
  if ( *(_DWORD *)(v4 + 168) == 3 && (v13 = *(_QWORD *)(v4 + 176)) != 0 )
    HUBMISC_QueryKseDeviceFlags(
      (unsigned int)pszDest,
      a2,
      a3,
      a4,
      v13,
      *(_QWORD *)(v4 + 184),
      *(_QWORD *)(v4 + 192),
      (__int64)&v34,
      (__int64)&v31,
      (__int64)&v32,
      (__int64)&v28,
      (__int64)&v29,
      (__int64)&v30,
      (__int64)&v33,
      0,
      *(_QWORD *)(*((_QWORD *)a1 + 1) + 1432LL));
  else
    HUBMISC_QueryKseDeviceFlags(
      (unsigned int)pszDest,
      a2,
      a3,
      a4,
      0LL,
      0LL,
      0LL,
      (__int64)&v34,
      (__int64)&v31,
      (__int64)&v32,
      (__int64)&v28,
      (__int64)&v29,
      (__int64)&v30,
      (__int64)&v33,
      0,
      *(_QWORD *)(*((_QWORD *)a1 + 1) + 1432LL));
  v14 = v34 | v31 | v32 | v28 | v29 | v30 | v33;
  if ( v11 )
  {
    v12 = RtlUnicodeStringPrintf(
            &DestinationString,
            L"IgnoreHWSerNum%04X%04X",
            *((unsigned __int16 *)a1 + 1002),
            *((unsigned __int16 *)a1 + 1003));
    if ( v12 < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 25;
LABEL_156:
      LOBYTE(v15) = 2;
      WPP_RECORDER_SF_d(
        *(_QWORD *)(*((_QWORD *)a1 + 1) + 1432LL),
        v15,
        5,
        v16,
        (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
        v12);
      goto LABEL_157;
    }
    v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, struct _UNICODE_STRING *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
            WdfDriverGlobals,
            v11,
            &DestinationString,
            4LL,
            &v23,
            0LL,
            0LL);
    if ( v12 < 0 )
    {
      if ( v12 != -1073741772 )
      {
        if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          goto LABEL_157;
        v16 = 26;
        goto LABEL_156;
      }
    }
    else if ( v23 )
    {
      _InterlockedOr(a1 + 413, 1u);
    }
  }
  v23 = 0;
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v27,
          L"\b\n", // osvc
          2LL,
          &v23,
          0LL,
          0LL);
  if ( v12 >= 0 )
  {
    if ( v23 )
    {
      *((_BYTE *)a1 + 2060) = BYTE1(v23);
      goto LABEL_28;
    }
    goto LABEL_21;
  }
  if ( v12 != -1073741772 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_157;
    v16 = 27;
    goto LABEL_156;
  }
  if ( (v14 & 1) != 0 )
  {
LABEL_21:
    _InterlockedOr(a1 + 410, 0x80u);
    goto LABEL_28;
  }
  if ( (v28 & 2) != 0 || (v29 & 2) != 0 || (v30 & 2) != 0 || (v31 & 2) != 0 || (v32 & 2) != 0 )
    _InterlockedOr(a1 + 413, 2u);
LABEL_28:
  v23 = 0;
  if ( !v25 )
    goto LABEL_35;
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v25,
          &g_IgnoreHwSerialNumber,
          4LL,
          &v23,
          0LL,
          0LL);
  if ( v12 < 0 )
  {
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 28;
      goto LABEL_156;
    }
LABEL_35:
    if ( (v14 & 0x40) == 0 )
      goto LABEL_37;
    goto LABEL_36;
  }
  if ( v23 )
LABEL_36:
    _InterlockedOr(a1 + 413, 1u);
LABEL_37:
  v23 = 0;
  if ( !v25 )
    goto LABEL_44;
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v25,
          L"68", // UseWin8DescriptorValidation
          4LL,
          &v23,
          0LL,
          0LL);
  if ( v12 < 0 )
  {
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 29;
      goto LABEL_156;
    }
LABEL_44:
    if ( (v14 & 0x80000000) == 0 )
      goto LABEL_46;
    goto LABEL_45;
  }
  if ( v23 )
LABEL_45:
    _InterlockedOr(a1 + 413, 0x200000u);
LABEL_46:
  v23 = 0;
  if ( !v25 )
    goto LABEL_53;
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v25,
          &g_ResetOnResume,
          4LL,
          &v23,
          0LL,
          0LL);
  if ( v12 < 0 )
  {
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 30;
      goto LABEL_156;
    }
LABEL_53:
    if ( (v14 & 4) == 0 )
      goto LABEL_55;
    goto LABEL_54;
  }
  if ( v23 )
LABEL_54:
    _InterlockedOr(a1 + 413, 4u);
LABEL_55:
  v23 = 0;
  _InterlockedOr(a1 + 413, 8u);
  if ( !v25 )
    goto LABEL_62;
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v25,
          L"&(", // DisableOnSoftRemove
          4LL,
          &v23,
          0LL,
          0LL);
  if ( v12 < 0 )
  {
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 31;
      goto LABEL_156;
    }
LABEL_62:
    if ( (v14 & 8) == 0 )
      goto LABEL_64;
    goto LABEL_63;
  }
  if ( v23 )
    goto LABEL_64;
LABEL_63:
  _InterlockedAnd(a1 + 413, 0xFFFFFFF7);
LABEL_64:
  v23 = 0;
  if ( v25 )
  {
    v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
            WdfDriverGlobals,
            v25,
            L"02", // RequestConfigDescOnReset
            4LL,
            &v23,
            0LL,
            0LL);
    if ( v12 >= 0 )
    {
      v17 = v23 == 0;
      goto LABEL_71;
    }
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 32;
      goto LABEL_156;
    }
  }
  v17 = (v14 & 0x10) == 0;
LABEL_71:
  if ( !v17 )
    _InterlockedOr(a1 + 413, 0x10u);
  v23 = 0;
  if ( !v25 )
    goto LABEL_80;
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v25,
          L":<", // DisableRecoveryFromPowerDrain
          4LL,
          &v23,
          0LL,
          0LL);
  if ( v12 < 0 )
  {
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 33;
      goto LABEL_156;
    }
LABEL_80:
    if ( (v14 & 0x1000000000LL) == 0 )
      goto LABEL_82;
    goto LABEL_81;
  }
  if ( !v23 )
    goto LABEL_82;
LABEL_81:
  _InterlockedOr(a1 + 413, 0x800000u);
LABEL_82:
  v23 = 0;
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v27,
          L"(*", // SkipContainerIdQuery
          4LL,
          &v23,
          0LL,
          0LL);
  if ( v12 >= 0 )
  {
    if ( !v23 )
      goto LABEL_88;
LABEL_87:
    _InterlockedOr(a1 + 413, 0x20u);
    goto LABEL_88;
  }
  if ( v12 != -1073741772 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      goto LABEL_157;
    v16 = 34;
    goto LABEL_156;
  }
  if ( (v14 & 0x20) != 0 )
    goto LABEL_87;
LABEL_88:
  v23 = 0;
  if ( v25 )
  {
    v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
            WdfDriverGlobals,
            v25,
            &g_DisableLpm,
            4LL,
            &v23,
            0LL,
            0LL);
    if ( v12 >= 0 )
    {
      v18 = v23 == 0;
      goto LABEL_95;
    }
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 35;
      goto LABEL_156;
    }
  }
  v18 = (v14 & 0x1000) == 0;
LABEL_95:
  if ( !v18 )
    _InterlockedOr(a1 + 413, 0x80u);
  if ( (v14 & 0x400) != 0 )
    _InterlockedOr(a1 + 413, 0x40u);
  if ( (v14 & 0x4000) != 0 )
    _InterlockedOr(a1 + 413, 0x100u);
  if ( (v14 & 0x10000) != 0 && *(_BYTE *)(*(_QWORD *)a1 + 240LL) )
    _InterlockedOr(a1 + 413, 0x80u);
  if ( (v14 & 0x80000) != 0 )
    _InterlockedOr(a1 + 413, 0x400u);
  if ( (v14 & 0x200000) != 0 )
    _InterlockedOr(a1 + 413, 0x800u);
  if ( (v14 & 0x800000) != 0 )
    _InterlockedOr(a1 + 413, 0x1000u);
  if ( (v14 & 0x1000000) != 0 )
    _InterlockedOr(a1 + 413, 0x2000u);
  v23 = 0;
  if ( !v25 )
  {
LABEL_118:
    v19 = (v14 & 0x8000000) == 0;
    goto LABEL_119;
  }
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v25,
          L",.", // SkipBOSDescriptorQuery
          4LL,
          &v23,
          0LL,
          0LL);
  if ( v12 < 0 )
  {
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 36;
      goto LABEL_156;
    }
    goto LABEL_118;
  }
  v19 = v23 == 0;
LABEL_119:
  if ( !v19 )
    _InterlockedOr(a1 + 413, 0x8000u);
  if ( (v14 & 0x2000) != 0 )
    _InterlockedOr(a1 + 413, 0x20000u);
  if ( (v14 & 0x20000) != 0 )
    _InterlockedOr(a1 + 413, 0x40000u);
  if ( (v14 & 0x40000000) != 0 )
    _InterlockedOr(a1 + 413, 0x100000u);
  if ( ((v14 & 0x400000) != 0 || (v14 & 0x4000000000LL) != 0 && v24) && (a1[410] & 2) == 0 )
    _InterlockedOr(a1 + 413, 0x80000u);
  if ( (v14 & 0x100000000LL) != 0 )
    _InterlockedOr(a1 + 413, 0x400000u);
  if ( (v14 & 0x2000000000LL) != 0 )
    _InterlockedOr(a1 + 413, 0x1000000u);
  if ( (v14 & 0x80000000000LL) != 0 )
    _InterlockedOr(a1 + 413, 0x4000000u);
  if ( (v14 & 0x800000000000LL) != 0 )
    _InterlockedOr(a1 + 413, 0x8000000u);
  v12 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, __int64 *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v27,
          L".0", // MsOs20DescriptorSetInfo
          8LL,
          &v35,
          0LL,
          0LL);
  if ( v12 < 0 )
  {
    if ( v12 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_157;
      v16 = 37;
      goto LABEL_156;
    }
  }
  else
  {
    _InterlockedOr(a1 + 619, 4u);
    v20 = BYTE6(v35);
    *((_DWORD *)a1 + 618) |= 4u;
    *((_BYTE *)a1 + 2060) = v20;
    *((_QWORD *)a1 + 311) = v35;
  }
  if ( *((_WORD *)a1 + 1002) == 8457 && *((_WORD *)a1 + 1003) == 2064 && (unsigned __int8)*((_WORD *)a1 + 1004) < 0x89u )
    _InterlockedOr(a1 + 413, 0x10000u);
  if ( v25 )
    HUBREG_QueryUsbflagsAlternateSettingFilter(a1);
  v12 = 0;
LABEL_157:
  if ( v25 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  if ( v11 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64))(WdfFunctions_01015 + 1848))(WdfDriverGlobals, v11);
  if ( v27 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return (unsigned int)v12;
}