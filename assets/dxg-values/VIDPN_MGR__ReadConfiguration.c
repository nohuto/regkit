__int64 __fastcall VIDPN_MGR::_ReadConfiguration(VIDPN_MGR *this)
{
  int RegistryValues; // eax
  int v3; // ebx
  unsigned int v4; // ecx
  struct DXGADAPTER *ContainingAdapter; // rax
  struct DXGADAPTER *v6; // rax
  struct DXGADAPTER *v7; // rax
  struct DXGADAPTER *v8; // rax
  bool v9; // al
  _DWORD *v10; // rbx
  __int64 v12; // [rsp+28h] [rbp-E0h]
  __int64 v13; // [rsp+28h] [rbp-E0h]
  __int64 v14; // [rsp+28h] [rbp-E0h]
  __int64 v15; // [rsp+28h] [rbp-E0h]
  unsigned int v16; // [rsp+38h] [rbp-D0h] BYREF
  int v17; // [rsp+3Ch] [rbp-CCh] BYREF
  int v18; // [rsp+40h] [rbp-C8h] BYREF
  int v19; // [rsp+44h] [rbp-C4h] BYREF
  int v20; // [rsp+48h] [rbp-C0h] BYREF
  int v21; // [rsp+4Ch] [rbp-BCh] BYREF
  __int64 v22; // [rsp+50h] [rbp-B8h] BYREF
  __int64 v23; // [rsp+58h] [rbp-B0h] BYREF
  __int64 v24; // [rsp+60h] [rbp-A8h]
  const wchar_t *v25; // [rsp+68h] [rbp-A0h]
  unsigned int *v26; // [rsp+70h] [rbp-98h]
  __int64 v27; // [rsp+78h] [rbp-90h]
  unsigned int *v28; // [rsp+80h] [rbp-88h]
  int v29; // [rsp+88h] [rbp-80h]
  __int64 v30; // [rsp+90h] [rbp-78h]
  int v31; // [rsp+98h] [rbp-70h]
  const wchar_t *v32; // [rsp+A0h] [rbp-68h]
  char *v33; // [rsp+A8h] [rbp-60h]
  int v34; // [rsp+B0h] [rbp-58h]
  char *v35; // [rsp+B8h] [rbp-50h]
  int v36; // [rsp+C0h] [rbp-48h]
  __int64 v37; // [rsp+C8h] [rbp-40h]
  int v38; // [rsp+D0h] [rbp-38h]
  const wchar_t *v39; // [rsp+D8h] [rbp-30h]
  __int64 *v40; // [rsp+E0h] [rbp-28h]
  int v41; // [rsp+E8h] [rbp-20h]
  __int64 *v42; // [rsp+F0h] [rbp-18h]
  int v43; // [rsp+F8h] [rbp-10h]
  __int64 v44; // [rsp+100h] [rbp-8h]
  int v45; // [rsp+108h] [rbp+0h]
  __int64 v46; // [rsp+110h] [rbp+8h]
  __int128 v47; // [rsp+118h] [rbp+10h]
  __int128 v48; // [rsp+128h] [rbp+20h]
  _QWORD v49[22]; // [rsp+138h] [rbp+30h] BYREF

  if ( !VIDPN_MGR::_BadMonitorSourceModeDiagnosibility )
  {
    v17 = 2;
    memset(v49, 0, 0xA8uLL);
    LODWORD(v49[1]) = 288;
    LODWORD(v49[4]) = 0x4000000;
    v49[2] = L"BadMonitorModeDiag";
    LODWORD(v49[11]) = 0x4000000;
    v49[3] = &v17;
    v49[5] = 0LL;
    v49[9] = L"AssertOnDdiViolation";
    LODWORD(v49[6]) = 0;
    v49[10] = &g_DmmAssertOnDdiViolation;
    v49[7] = 0LL;
    LODWORD(v49[8]) = 288;
    v49[12] = 0LL;
    LODWORD(v49[13]) = 0;
    HIDWORD(v12) = 0;
    RegistryValues = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\DMM", v49);
    v3 = RegistryValues;
    if ( RegistryValues >= 0 )
    {
      v4 = v17;
    }
    else
    {
      WdLogSingleEntry1(7LL, RegistryValues);
      WdLogGlobalForLineNumber = 689;
      if ( v3 != -1073741772 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 692;
      }
      v4 = 2;
      v17 = 2;
    }
    if ( v4 == 1 || v4 == 2 )
    {
      VIDPN_MGR::_BadMonitorSourceModeDiagnosibility = v4;
    }
    else
    {
      WdLogSingleEntry1(2LL, v4);
      WdLogGlobalForLineNumber = 717;
    }
  }
  v18 = 0;
  ContainingAdapter = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v12) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)ContainingAdapter + 27), L"AllowUnspecifiedVSync", &v18, 4LL, v12) >= 0 )
  {
    VIDPN_MGR::_bAllowUnspecifiedVSync = v18 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    WdLogGlobalForLineNumber = 738;
  }
  v19 = 0;
  v6 = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v13) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)v6 + 27), L"AllowUnspecifiedHSync", &v19, 4LL, v13) >= 0 )
  {
    VIDPN_MGR::_bAllowUnspecifiedHSync = v19 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    WdLogGlobalForLineNumber = 761;
  }
  v20 = 0;
  v7 = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v14) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)v7 + 27), L"AllowUnspecifiedPixelRate", &v20, 4LL, v14) >= 0 )
  {
    VIDPN_MGR::_bAllowUnspecifiedPixelRate = v20 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    WdLogGlobalForLineNumber = 784;
  }
  v21 = 0;
  v8 = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v15) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)v8 + 27), L"ForceDualViewBehavior", &v21, 4LL, v15) >= 0 )
  {
    v9 = v21 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    v9 = 0;
    WdLogGlobalForLineNumber = 808;
  }
  *((_BYTE *)this + 520) = v9;
  v10 = (_DWORD *)((char *)this + 544);
  v16 = 1000;
  LODWORD(v27) = 67108868;
  v34 = 67108868;
  v25 = L"RapidHPDTime";
  v41 = 67108868;
  v26 = &v16;
  *((_DWORD *)this + 136) = 5;
  v28 = &v16;
  LODWORD(v22) = 0;
  v32 = L"RapidHPDThresholdCount";
  v23 = 0LL;
  v39 = L"EnableExperimentalRefreshRates";
  v40 = &v22;
  v42 = &v22;
  LODWORD(v24) = 288;
  v29 = 4;
  v30 = 0LL;
  v31 = 288;
  v33 = (char *)this + 544;
  v35 = (char *)this + 544;
  v36 = 4;
  v37 = 0LL;
  v38 = 288;
  v43 = 4;
  v44 = 0LL;
  v45 = 0;
  v46 = 0LL;
  v47 = 0LL;
  v48 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v23);
  if ( v16 > 0xEA60 )
    v16 = 60000;
  *((_DWORD *)this + 135) = 10000 * v16 / KeQueryTimeIncrement();
  if ( *v10 == 1 )
  {
    *v10 = 0;
  }
  else if ( *v10 > 0x20u )
  {
    *v10 = 32;
  }
  return 0LL;
}
