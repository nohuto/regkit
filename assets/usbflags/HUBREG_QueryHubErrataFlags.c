__int64 __fastcall HUBREG_QueryHubErrataFlags(__int64 a1, __int64 a2, __int64 a3, __int64 a4)
{
  char v4; // r13
  const CHAR *v5; // r12
  const CHAR *v6; // r15
  const CHAR *v7; // r14
  bool v9; // zf
  __int64 v10; // rdx
  __int64 v11; // rbx
  int v12; // edx
  int v13; // esi
  int v14; // r9d
  int v15; // esi
  int IsEnabledDeviceUsageNoInline; // eax
  int v17; // edx
  __int64 v18; // rax
  __int64 v19; // rcx
  int v20; // ecx
  __int64 v21; // rcx
  int v22; // ecx
  __int64 v24; // [rsp+38h] [rbp-C8h]
  int v25; // [rsp+80h] [rbp-80h] BYREF
  __int64 v26; // [rsp+88h] [rbp-78h] BYREF
  __int64 v27; // [rsp+90h] [rbp-70h] BYREF
  __int64 v28; // [rsp+98h] [rbp-68h] BYREF
  __int64 v29; // [rsp+A0h] [rbp-60h] BYREF
  __int64 v30; // [rsp+A8h] [rbp-58h] BYREF
  __int64 v31; // [rsp+B0h] [rbp-50h] BYREF
  __int64 v32; // [rsp+B8h] [rbp-48h] BYREF
  __int64 v33; // [rsp+C0h] [rbp-40h] BYREF
  __int64 v34; // [rsp+C8h] [rbp-38h] BYREF
  _BYTE v35[32]; // [rsp+D0h] [rbp-30h] BYREF
  __int64 *v36; // [rsp+F0h] [rbp-10h]
  __int64 v37; // [rsp+F8h] [rbp-8h]
  const CHAR *v38; // [rsp+100h] [rbp+0h]
  int v39; // [rsp+108h] [rbp+8h]
  int v40; // [rsp+10Ch] [rbp+Ch]
  const CHAR *v41; // [rsp+110h] [rbp+10h]
  int v42; // [rsp+118h] [rbp+18h]
  int v43; // [rsp+11Ch] [rbp+1Ch]
  const CHAR *v44; // [rsp+120h] [rbp+20h]
  int v45; // [rsp+128h] [rbp+28h]
  int v46; // [rsp+12Ch] [rbp+2Ch]
  __int64 *v47; // [rsp+130h] [rbp+30h]
  __int64 v48; // [rsp+138h] [rbp+38h]

  v4 = *(_BYTE *)(a1 + 200);
  v5 = (const CHAR *)a4;
  v24 = *(_QWORD *)(a1 + 2536);
  v6 = (const CHAR *)a3;
  v7 = (const CHAR *)a2;
  v26 = 0LL;
  HUBREG_OpenCreateUsbflagsDeviceKey(a2, a3, a4, 0x20019u, 0LL, &v26, 0, v24);
  v9 = *(_DWORD *)(a1 + 168) == 3;
  v32 = 0LL;
  v31 = 0LL;
  v30 = 0LL;
  v29 = 0LL;
  v28 = 0LL;
  v27 = 0LL;
  if ( v9 && (v10 = *(_QWORD *)(a1 + 176)) != 0 )
    HUBMISC_QueryKseDeviceFlags(
      0,
      (_DWORD)v7,
      (_DWORD)v6,
      (_DWORD)v5,
      v10,
      *(_QWORD *)(a1 + 184),
      *(_QWORD *)(a1 + 192),
      (__int64)&v32,
      (__int64)&v31,
      (__int64)&v30,
      (__int64)&v29,
      (__int64)&v28,
      (__int64)&v27,
      0LL,
      *(_BYTE *)(a1 + 240) == 0,
      *(_QWORD *)(a1 + 2536));
  else
    HUBMISC_QueryKseDeviceFlags(
      0,
      (_DWORD)v7,
      (_DWORD)v6,
      (_DWORD)v5,
      0LL,
      0LL,
      0LL,
      (__int64)&v32,
      (__int64)&v31,
      (__int64)&v30,
      (__int64)&v29,
      (__int64)&v28,
      (__int64)&v27,
      0LL,
      *(_BYTE *)(a1 + 240) == 0,
      *(_QWORD *)(a1 + 2536));
  v11 = v32 | v31 | v30 | v29 | v28 | v27;
  v25 = 0;
  if ( !v26 )
    goto LABEL_13;
  v13 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
          WdfDriverGlobals,
          v26,
          &g_ResetTTOnCancel, // ResetTTOnCancel
          4LL,
          &v25,
          0LL,
          0LL);
  if ( v13 < 0 )
  {
    if ( v13 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_109;
      v14 = 38;
      goto LABEL_12;
    }
LABEL_13:
    if ( (v11 & 0x100) == 0 )
      goto LABEL_15;
    goto LABEL_14;
  }
  if ( v25 )
LABEL_14:
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x800u);
LABEL_15:
  v25 = 0;
  if ( v26 )
  {
    v13 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
            WdfDriverGlobals,
            v26,
            L".0", // NoClearTTBufferOnCancel
            4LL,
            &v25,
            0LL,
            0LL);
    if ( v13 >= 0 )
    {
      if ( !v25 )
        goto LABEL_24;
      goto LABEL_23;
    }
    if ( v13 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_109;
      v14 = 39;
      goto LABEL_12;
    }
  }
  if ( (v11 & 0x200) != 0 )
  {
LABEL_23:
    _InterlockedAnd((volatile signed __int32 *)(a1 + 40), 0xFFFFF7FF);
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x1000u);
  }
LABEL_24:
  if ( (v11 & 0x800) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x2000u);
  v25 = 0;
  if ( v26 )
  {
    v13 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, void *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
            WdfDriverGlobals,
            v26,
            &g_DisableLpm, // DisableLPM
            4LL,
            &v25,
            0LL,
            0LL);
    if ( v13 >= 0 )
    {
      if ( !v25 )
        goto LABEL_35;
      goto LABEL_34;
    }
    if ( v13 != -1073741772 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_109;
      v14 = 40;
LABEL_12:
      LOBYTE(v12) = 2;
      WPP_RECORDER_SF_d(
        *(_QWORD *)(a1 + 2536),
        v12,
        3,
        v14,
        (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids,
        v13);
      goto LABEL_109;
    }
  }
  if ( (v11 & 0x1000) != 0 )
LABEL_34:
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x8000u);
LABEL_35:
  if ( (v11 & 0x2000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x10000u);
  if ( (v11 & 0x8000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x80000u);
  if ( (v11 & 0x40000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x100000u);
  if ( (v11 & 0x100000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x200000u);
  if ( (v11 & 0x400000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x800000u);
  if ( (v11 & 0x2000000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x1000000u);
  v15 = 1;
  if ( (v11 & 0x4000000) != 0 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 2512), 1u);
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x8000u);
  }
  if ( (v11 & 0x40000000000LL) != 0 )
  {
    _InterlockedOr((volatile signed __int32 *)(a1 + 2512), 4u);
    _InterlockedOr((volatile signed __int32 *)(a1 + 2512), 1u);
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x8000u);
  }
  if ( !*(_BYTE *)(a1 + 240) && (v11 & 0x10000) != 0 )
    *(_DWORD *)(a1 + 2512) |= 1u;
  if ( (v11 & 0x10000000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x2000000u);
  if ( (v11 & 0x8000000000LL) != 0 && v4 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 1u);
  if ( (v11 & 0x20000000) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x10000000u);
  if ( (v11 & 0x100000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 8u);
  if ( (v11 & 0x20000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 4u);
  _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x40000000u);
  if ( (v11 & 8) != 0 )
    _InterlockedAnd((volatile signed __int32 *)(a1 + 40), 0xBFFFFFFF);
  if ( (v11 & 0x800000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x80000000);
  if ( (v11 & 0x10000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 2u);
  if ( (v11 & 0x200000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 0x10u);
  if ( (v11 & 0x400000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 0x40u);
  if ( (v11 & 0x1000000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 0x80u);
  if ( (unsigned int)Feature_RH1S__private_IsEnabledDeviceUsageNoInline() && (v11 & 0x2000000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 0x400u);
  IsEnabledDeviceUsageNoInline = Feature_RH5S__private_IsEnabledDeviceUsageNoInline();
  v17 = 0;
  if ( IsEnabledDeviceUsageNoInline && (v11 & 0x4000000000000LL) != 0 )
    _InterlockedOr((volatile signed __int32 *)(a1 + 44), 0x800u);
  if ( *(_WORD *)(a1 + 2480) == 8457 && *(_WORD *)(a1 + 2482) == 2064 && (unsigned __int8)*(_WORD *)(a1 + 2484) < 0x89u )
    _InterlockedOr((volatile signed __int32 *)(a1 + 40), 0x800000u);
  if ( (*(_DWORD *)(a1 + 40) & 0x800000) != 0 )
  {
    if ( (unsigned int)dword_14006D318 > 5
      && (qword_14006D328 & 0x200000000004LL) != 0
      && (qword_14006D330 & 0x200000000004LL) == qword_14006D330 )
    {
      v33 = 1LL;
      v36 = &v33;
      v18 = -1LL;
      v37 = 8LL;
      if ( v7 )
      {
        v19 = -1LL;
        do
          ++v19;
        while ( v7[v19] );
        v20 = v19 + 1;
      }
      else
      {
        v7 = File;
        v20 = 1;
      }
      v38 = v7;
      v39 = v20;
      v40 = 0;
      if ( v6 )
      {
        v21 = -1LL;
        do
          ++v21;
        while ( v6[v21] );
        v22 = v21 + 1;
      }
      else
      {
        v6 = File;
        v22 = 1;
      }
      v41 = v6;
      v42 = v22;
      v43 = 0;
      if ( v5 )
      {
        do
          ++v18;
        while ( v5[v18] );
        v15 = v18 + 1;
      }
      else
      {
        v5 = File;
      }
      v46 = 0;
      v47 = &v34;
      v44 = v5;
      v45 = v15;
      v34 = 16779264LL;
      v48 = 8LL;
      tlgWriteAgg(v22, (unsigned int)&unk_140068F3E, (unsigned int)File, 7, (__int64)v35);
    }
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      LOBYTE(v17) = 4;
      WPP_RECORDER_SF_(*(_QWORD *)(a1 + 2536), v17, 3, 41, (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids);
    }
  }
  v13 = 0;
LABEL_109:
  if ( v26 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return (unsigned int)v13;
}