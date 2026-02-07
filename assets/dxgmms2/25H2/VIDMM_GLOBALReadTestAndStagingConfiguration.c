void VIDMM_GLOBAL::ReadTestAndStagingConfiguration(void)
{
  int v0; // edx
  int v1; // eax
  int v2; // ecx
  int v3; // eax
  unsigned int v4; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v5; // [rsp+34h] [rbp-CCh] BYREF
  unsigned int v6; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v7; // [rsp+3Ch] [rbp-C4h] BYREF
  int v8; // [rsp+40h] [rbp-C0h] BYREF
  int v9; // [rsp+44h] [rbp-BCh] BYREF
  int v10; // [rsp+48h] [rbp-B8h] BYREF
  int v11; // [rsp+4Ch] [rbp-B4h] BYREF
  int v12; // [rsp+50h] [rbp-B0h] BYREF
  int v13; // [rsp+54h] [rbp-ACh] BYREF
  int v14; // [rsp+58h] [rbp-A8h] BYREF
  int v15; // [rsp+5Ch] [rbp-A4h] BYREF
  int v16; // [rsp+60h] [rbp-A0h] BYREF
  int v17; // [rsp+64h] [rbp-9Ch] BYREF
  int v18; // [rsp+68h] [rbp-98h] BYREF
  int v19; // [rsp+6Ch] [rbp-94h] BYREF
  int v20; // [rsp+70h] [rbp-90h] BYREF
  int v21; // [rsp+74h] [rbp-8Ch] BYREF
  int v22; // [rsp+78h] [rbp-88h] BYREF
  int v23; // [rsp+7Ch] [rbp-84h] BYREF
  __int64 v24; // [rsp+80h] [rbp-80h] BYREF
  int v25; // [rsp+88h] [rbp-78h]
  const wchar_t *v26; // [rsp+90h] [rbp-70h]
  unsigned int *v27; // [rsp+98h] [rbp-68h]
  int v28; // [rsp+A0h] [rbp-60h]
  int *v29; // [rsp+A8h] [rbp-58h]
  int v30; // [rsp+B0h] [rbp-50h]
  __int64 v31; // [rsp+B8h] [rbp-48h]
  int v32; // [rsp+C0h] [rbp-40h]
  const wchar_t *v33; // [rsp+C8h] [rbp-38h]
  unsigned int *v34; // [rsp+D0h] [rbp-30h]
  int v35; // [rsp+D8h] [rbp-28h]
  int *v36; // [rsp+E0h] [rbp-20h]
  int v37; // [rsp+E8h] [rbp-18h]
  __int64 v38; // [rsp+F0h] [rbp-10h]
  int v39; // [rsp+F8h] [rbp-8h]
  const wchar_t *v40; // [rsp+100h] [rbp+0h]
  int *v41; // [rsp+108h] [rbp+8h]
  int v42; // [rsp+110h] [rbp+10h]
  int *v43; // [rsp+118h] [rbp+18h]
  int v44; // [rsp+120h] [rbp+20h]
  __int64 v45; // [rsp+128h] [rbp+28h]
  int v46; // [rsp+130h] [rbp+30h]
  const wchar_t *v47; // [rsp+138h] [rbp+38h]
  int *v48; // [rsp+140h] [rbp+40h]
  int v49; // [rsp+148h] [rbp+48h]
  int *v50; // [rsp+150h] [rbp+50h]
  int v51; // [rsp+158h] [rbp+58h]
  __int64 v52; // [rsp+160h] [rbp+60h]
  int v53; // [rsp+168h] [rbp+68h]
  const wchar_t *v54; // [rsp+170h] [rbp+70h]
  int *v55; // [rsp+178h] [rbp+78h]
  int v56; // [rsp+180h] [rbp+80h]
  int *v57; // [rsp+188h] [rbp+88h]
  int v58; // [rsp+190h] [rbp+90h]
  __int64 v59; // [rsp+198h] [rbp+98h]
  int v60; // [rsp+1A0h] [rbp+A0h]
  const wchar_t *v61; // [rsp+1A8h] [rbp+A8h]
  int *v62; // [rsp+1B0h] [rbp+B0h]
  int v63; // [rsp+1B8h] [rbp+B8h]
  int *v64; // [rsp+1C0h] [rbp+C0h]
  int v65; // [rsp+1C8h] [rbp+C8h]
  __int64 v66; // [rsp+1D0h] [rbp+D0h]
  int v67; // [rsp+1D8h] [rbp+D8h]
  const wchar_t *v68; // [rsp+1E0h] [rbp+E0h]
  unsigned int *v69; // [rsp+1E8h] [rbp+E8h]
  int v70; // [rsp+1F0h] [rbp+F0h]
  int *v71; // [rsp+1F8h] [rbp+F8h]
  int v72; // [rsp+200h] [rbp+100h]
  __int64 v73; // [rsp+208h] [rbp+108h]
  int v74; // [rsp+210h] [rbp+110h]
  const wchar_t *v75; // [rsp+218h] [rbp+118h]
  int *v76; // [rsp+220h] [rbp+120h]
  int v77; // [rsp+228h] [rbp+128h]
  int *v78; // [rsp+230h] [rbp+130h]
  int v79; // [rsp+238h] [rbp+138h]
  __int64 v80; // [rsp+240h] [rbp+140h]
  int v81; // [rsp+248h] [rbp+148h]
  const wchar_t *v82; // [rsp+250h] [rbp+150h]
  unsigned int *v83; // [rsp+258h] [rbp+158h]
  int v84; // [rsp+260h] [rbp+160h]
  int *v85; // [rsp+268h] [rbp+168h]
  int v86; // [rsp+270h] [rbp+170h]
  __int64 v87; // [rsp+278h] [rbp+178h]
  int v88; // [rsp+280h] [rbp+180h]
  const wchar_t *v89; // [rsp+288h] [rbp+188h]
  int *v90; // [rsp+290h] [rbp+190h]
  int v91; // [rsp+298h] [rbp+198h]
  int *v92; // [rsp+2A0h] [rbp+1A0h]
  int v93; // [rsp+2A8h] [rbp+1A8h]
  __int128 v94; // [rsp+2B0h] [rbp+1B0h]
  __int128 v95; // [rsp+2C0h] [rbp+1C0h]
  __int128 v96; // [rsp+2D0h] [rbp+1D0h]
  __int64 v97; // [rsp+2E0h] [rbp+1E0h]

  v15 = 1;
  v7 = 1;
  v16 = 0;
  v14 = 25;
  v4 = 25;
  v8 = 0;
  v18 = 0x400000;
  v11 = 0x400000;
  v20 = 32;
  v5 = 32;
  v21 = -1;
  v12 = -1;
  v26 = L"BudgetThreshold";
  v27 = &v4;
  v29 = &v14;
  v33 = L"PagingQueueFenceIncrement";
  v34 = &v7;
  v36 = &v15;
  v40 = L"RestrictToPreferredSegment";
  v41 = &v8;
  v43 = &v16;
  v47 = L"Use64KPages";
  v48 = &v10;
  v50 = &v17;
  v54 = L"ExpandTo64KBAllocationSizeThreshold";
  v55 = &v11;
  v57 = &v18;
  v17 = 0;
  v10 = 0;
  v19 = 0;
  v9 = 0;
  v22 = 0;
  v6 = 0;
  v23 = 0;
  v13 = 0;
  v24 = 0LL;
  v25 = 288;
  v28 = 67108868;
  v30 = 4;
  v31 = 0LL;
  v32 = 288;
  v35 = 67108868;
  v37 = 4;
  v38 = 0LL;
  v39 = 288;
  v42 = 67108868;
  v44 = 4;
  v45 = 0LL;
  v46 = 288;
  v49 = 67108868;
  v51 = 4;
  v52 = 0LL;
  v53 = 288;
  v56 = 67108868;
  v58 = 4;
  v59 = 0LL;
  v60 = 288;
  v61 = L"AlwaysDecommitOnOffer";
  v63 = 67108868;
  v62 = &v9;
  v65 = 4;
  v64 = &v19;
  v68 = L"LazyDecommitChunkSizeMB";
  v69 = &v5;
  v71 = &v20;
  v75 = L"DxgMms2OfferReclaim";
  v76 = &v12;
  v78 = &v21;
  v82 = L"LargifyUpgradeThresholdPercent";
  v83 = &v6;
  v85 = &v22;
  v89 = L"LargifyUpgradeThresholdBytes";
  v90 = &v13;
  v92 = &v23;
  v67 = 288;
  v70 = 67108868;
  v72 = 4;
  v74 = 288;
  v77 = 67108868;
  v79 = 4;
  v81 = 288;
  v84 = 67108868;
  v86 = 4;
  v88 = 288;
  v91 = 67108868;
  v93 = 4;
  v97 = 0LL;
  v66 = 0LL;
  v73 = 0LL;
  v80 = 0LL;
  v87 = 0LL;
  v94 = 0LL;
  v95 = 0LL;
  v96 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v24, 0LL, 0LL);
  v0 = 100;
  v1 = 100;
  if ( v4 < 0x64 )
    v1 = v4;
  dword_140081590 = v1;
  dword_140081594 = v7;
  if ( v7 <= 0x51EB851 )
  {
    if ( !v7 )
      dword_140081594 = 1;
  }
  else
  {
    dword_140081594 = 85899345;
  }
  v2 = v12;
  dword_140081610 = v8;
  dword_14008161C = v9;
  v3 = 512;
  if ( v5 < 0x200 )
    v3 = v5;
  dword_140081620 = v3;
  dword_140081614 = v10;
  dword_140081618 = v11;
  dword_1400816B4 = v13;
  if ( (unsigned int)(v12 - 3) <= 0xFFFFFFFB )
    v2 = 0;
  dword_1400815E8 = v2;
  if ( v6 < 0x64 )
    v0 = v6;
  dword_1400816B0 = v0;
}