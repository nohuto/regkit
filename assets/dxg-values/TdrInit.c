void TdrInit(void)
{
  int v0; // eax
  int v1; // eax
  int v2; // eax
  int v3; // eax
  int v4; // eax
  int v5; // eax
  int v6; // eax
  __int64 v7; // rcx
  unsigned int v8; // [rsp+38h] [rbp-D0h] BYREF
  unsigned int v9; // [rsp+3Ch] [rbp-CCh] BYREF
  unsigned int v10; // [rsp+40h] [rbp-C8h] BYREF
  unsigned int v11; // [rsp+44h] [rbp-C4h] BYREF
  unsigned int v12; // [rsp+48h] [rbp-C0h] BYREF
  unsigned int v13; // [rsp+4Ch] [rbp-BCh] BYREF
  unsigned int v14; // [rsp+50h] [rbp-B8h] BYREF
  unsigned int v15; // [rsp+54h] [rbp-B4h] BYREF
  int v16; // [rsp+58h] [rbp-B0h] BYREF
  int v17; // [rsp+5Ch] [rbp-ACh] BYREF
  int v18; // [rsp+60h] [rbp-A8h] BYREF
  int v19; // [rsp+64h] [rbp-A4h] BYREF
  int v20; // [rsp+68h] [rbp-A0h] BYREF
  int v21; // [rsp+6Ch] [rbp-9Ch] BYREF
  __int64 v22; // [rsp+70h] [rbp-98h] BYREF
  __int64 v23; // [rsp+78h] [rbp-90h] BYREF
  int v24; // [rsp+80h] [rbp-88h]
  const wchar_t *v25; // [rsp+88h] [rbp-80h]
  unsigned int *v26; // [rsp+90h] [rbp-78h]
  int v27; // [rsp+98h] [rbp-70h]
  int *v28; // [rsp+A0h] [rbp-68h]
  int v29; // [rsp+A8h] [rbp-60h]
  __int64 v30; // [rsp+B0h] [rbp-58h]
  int v31; // [rsp+B8h] [rbp-50h]
  const wchar_t *v32; // [rsp+C0h] [rbp-48h]
  unsigned int *v33; // [rsp+C8h] [rbp-40h]
  int v34; // [rsp+D0h] [rbp-38h]
  int *v35; // [rsp+D8h] [rbp-30h]
  int v36; // [rsp+E0h] [rbp-28h]
  __int64 v37; // [rsp+E8h] [rbp-20h]
  int v38; // [rsp+F0h] [rbp-18h]
  const wchar_t *v39; // [rsp+F8h] [rbp-10h]
  unsigned int *v40; // [rsp+100h] [rbp-8h]
  int v41; // [rsp+108h] [rbp+0h]
  int *v42; // [rsp+110h] [rbp+8h]
  int v43; // [rsp+118h] [rbp+10h]
  __int64 v44; // [rsp+120h] [rbp+18h]
  int v45; // [rsp+128h] [rbp+20h]
  const wchar_t *v46; // [rsp+130h] [rbp+28h]
  unsigned int *v47; // [rsp+138h] [rbp+30h]
  int v48; // [rsp+140h] [rbp+38h]
  int *v49; // [rsp+148h] [rbp+40h]
  int v50; // [rsp+150h] [rbp+48h]
  __int64 v51; // [rsp+158h] [rbp+50h]
  int v52; // [rsp+160h] [rbp+58h]
  const wchar_t *v53; // [rsp+168h] [rbp+60h]
  unsigned int *v54; // [rsp+170h] [rbp+68h]
  int v55; // [rsp+178h] [rbp+70h]
  int *v56; // [rsp+180h] [rbp+78h]
  int v57; // [rsp+188h] [rbp+80h]
  __int64 v58; // [rsp+190h] [rbp+88h]
  int v59; // [rsp+198h] [rbp+90h]
  const wchar_t *v60; // [rsp+1A0h] [rbp+98h]
  unsigned int *v61; // [rsp+1A8h] [rbp+A0h]
  int v62; // [rsp+1B0h] [rbp+A8h]
  int *v63; // [rsp+1B8h] [rbp+B0h]
  int v64; // [rsp+1C0h] [rbp+B8h]
  __int64 v65; // [rsp+1C8h] [rbp+C0h]
  int v66; // [rsp+1D0h] [rbp+C8h]
  const wchar_t *v67; // [rsp+1D8h] [rbp+D0h]
  unsigned int *v68; // [rsp+1E0h] [rbp+D8h]
  int v69; // [rsp+1E8h] [rbp+E0h]
  __int64 *v70; // [rsp+1F0h] [rbp+E8h]
  int v71; // [rsp+1F8h] [rbp+F0h]
  __int64 v72; // [rsp+200h] [rbp+F8h]
  int v73; // [rsp+208h] [rbp+100h]
  const wchar_t *v74; // [rsp+210h] [rbp+108h]
  unsigned int *v75; // [rsp+218h] [rbp+110h]
  int v76; // [rsp+220h] [rbp+118h]
  char *v77; // [rsp+228h] [rbp+120h]
  int v78; // [rsp+230h] [rbp+128h]
  __int64 v79; // [rsp+238h] [rbp+130h]
  int v80; // [rsp+240h] [rbp+138h]
  __int64 v81; // [rsp+248h] [rbp+140h]
  __int128 v82; // [rsp+250h] [rbp+148h]
  __int128 v83; // [rsp+260h] [rbp+158h]

  v22 = 0x20000003CLL;
  v13 = 0;
  v8 = 0;
  v9 = 0;
  v16 = 3;
  v25 = L"TdrLevel";
  v26 = &v13;
  v28 = &v16;
  v10 = 0;
  v17 = 2;
  v32 = L"TdrDelay";
  v18 = 2;
  v33 = &v8;
  v35 = &v17;
  v39 = L"TdrDodPresentDelay";
  v40 = &v9;
  v42 = &v18;
  v46 = L"TdrDodVSyncDelay";
  v47 = &v10;
  v49 = &v19;
  v53 = L"TdrDdiDelay";
  v54 = &v11;
  v56 = &v20;
  v60 = L"TdrLimitCount";
  v61 = &v14;
  v19 = 2;
  v20 = 5;
  v11 = 0;
  v12 = 0;
  v21 = 5;
  v14 = 0;
  v15 = 0;
  v23 = 0LL;
  v24 = 288;
  v27 = 67108868;
  v29 = 4;
  v30 = 0LL;
  v31 = 288;
  v34 = 67108868;
  v36 = 4;
  v37 = 0LL;
  v38 = 288;
  v41 = 67108868;
  v43 = 4;
  v44 = 0LL;
  v45 = 288;
  v48 = 67108868;
  v50 = 4;
  v51 = 0LL;
  v52 = 288;
  v55 = 67108868;
  v57 = 4;
  v58 = 0LL;
  v59 = 288;
  v62 = 67108868;
  v63 = &v21;
  v64 = 4;
  v67 = L"TdrLimitTime";
  v66 = 288;
  v68 = &v15;
  v70 = &v22;
  v74 = L"TdrDebugMode";
  v75 = &v12;
  v69 = 67108868;
  v71 = 4;
  v73 = 288;
  v76 = 67108868;
  v78 = 4;
  v77 = (char *)&v22 + 4;
  v65 = 0LL;
  v72 = 0LL;
  v79 = 0LL;
  v80 = 0;
  v81 = 0LL;
  v82 = 0LL;
  v83 = 0LL;
  v0 = RtlQueryRegistryValuesEx(0LL, L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\GraphicsDrivers", &v23);
  if ( v0 < 0 )
  {
    v13 = 3;
    v8 = 2;
    v9 = 2;
    v10 = 2;
    v11 = 5;
    v12 = 2;
    WdLogSingleEntry1(3LL, v0);
    WdLogGlobalForLineNumber = 2211;
  }
  if ( v13 < 2 || v13 == 3 )
  {
    g_TdrConfig = v13;
  }
  else
  {
    g_TdrConfig = 3;
    WdLogSingleEntry2(3LL, v13, 3LL);
    WdLogGlobalForLineNumber = 2238;
  }
  v1 = v8;
  if ( v8 )
  {
    if ( v8 > 0x384 )
      v1 = 900;
    dword_1C01537DC = v1;
  }
  else
  {
    dword_1C01537DC = 1;
  }
  if ( dword_1C01537DC != v8 )
  {
    WdLogSingleEntry2(3LL, v8, (unsigned int)dword_1C01537DC);
    WdLogGlobalForLineNumber = 2262;
  }
  v2 = v9;
  if ( v9 )
  {
    if ( v9 > 0x384 )
      v2 = 900;
    dword_1C01537E0 = v2;
  }
  else
  {
    dword_1C01537E0 = 1;
  }
  if ( dword_1C01537E0 != v9 )
  {
    WdLogSingleEntry2(3LL, v9, (unsigned int)dword_1C01537E0);
    WdLogGlobalForLineNumber = 2287;
  }
  v3 = v10;
  if ( v10 )
  {
    if ( v10 > 0x384 )
      v3 = 900;
    dword_1C01537E4 = v3;
  }
  else
  {
    dword_1C01537E4 = 1;
  }
  if ( dword_1C01537E4 != v10 )
  {
    WdLogSingleEntry2(3LL, v10, (unsigned int)dword_1C01537E4);
    WdLogGlobalForLineNumber = 2312;
  }
  v4 = v11;
  if ( v11 )
  {
    if ( v11 > 0x384 )
      v4 = 900;
    dword_1C01537E8 = v4;
  }
  else
  {
    dword_1C01537E8 = 1;
  }
  if ( dword_1C01537E8 != v11 )
  {
    WdLogSingleEntry2(3LL, v11, (unsigned int)dword_1C01537E8);
    WdLogGlobalForLineNumber = 2337;
  }
  v5 = v14;
  if ( v14 <= 0x20 )
  {
    if ( !v14 )
      v5 = 1;
    dword_1C01537F0 = v5;
  }
  else
  {
    dword_1C01537F0 = 32;
  }
  if ( dword_1C01537F0 != v14 )
  {
    WdLogSingleEntry2(3LL, v14, (unsigned int)dword_1C01537F0);
    WdLogGlobalForLineNumber = 2362;
  }
  v6 = v15;
  v7 = 3600LL;
  if ( v15 <= 0xE10 )
  {
    if ( v15 < 5 )
      v6 = 5;
    dword_1C01537F4 = v6;
  }
  else
  {
    dword_1C01537F4 = 3600;
  }
  if ( dword_1C01537F4 != v15 )
  {
    WdLogSingleEntry2(3LL, v15, (unsigned int)dword_1C01537F4);
    WdLogGlobalForLineNumber = 2387;
  }
  LOBYTE(v7) = 1;
  byte_1C01537EC = (unsigned __int8)WdIsDebuggerPresent(v7) != 0;
  if ( v12 < 2 || v12 - 2 < 2 )
    g_TdrDebugMode = v12;
  else
    g_TdrDebugMode = 2;
  if ( g_TdrDebugMode != v12 )
  {
    WdLogSingleEntry2(3LL, v12, g_TdrDebugMode);
    WdLogGlobalForLineNumber = 2418;
  }
  TdrHistoryInit(&g_TdrHistory);
}
