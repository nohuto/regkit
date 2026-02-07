void VIDMM_GLOBAL::ReadTestAndStagingConfiguration(void)
{
  int v0; // eax
  int v1; // ecx
  int v2; // eax
  unsigned int v3; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v4; // [rsp+34h] [rbp-CCh] BYREF
  unsigned int v5; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v6; // [rsp+3Ch] [rbp-C4h] BYREF
  int v7; // [rsp+40h] [rbp-C0h] BYREF
  int v8; // [rsp+44h] [rbp-BCh] BYREF
  int v9; // [rsp+48h] [rbp-B8h] BYREF
  int v10; // [rsp+4Ch] [rbp-B4h] BYREF
  int v11; // [rsp+50h] [rbp-B0h] BYREF
  int v12; // [rsp+54h] [rbp-ACh] BYREF
  int v13; // [rsp+58h] [rbp-A8h] BYREF
  int v14; // [rsp+5Ch] [rbp-A4h] BYREF
  int v15; // [rsp+60h] [rbp-A0h] BYREF
  int v16; // [rsp+64h] [rbp-9Ch] BYREF
  int v17; // [rsp+68h] [rbp-98h] BYREF
  int v18; // [rsp+6Ch] [rbp-94h] BYREF
  int v19; // [rsp+70h] [rbp-90h] BYREF
  int v20; // [rsp+74h] [rbp-8Ch] BYREF
  __int64 v21; // [rsp+80h] [rbp-80h] BYREF
  int v22; // [rsp+88h] [rbp-78h]
  const wchar_t *v23; // [rsp+90h] [rbp-70h]
  unsigned int *v24; // [rsp+98h] [rbp-68h]
  int v25; // [rsp+A0h] [rbp-60h]
  int *v26; // [rsp+A8h] [rbp-58h]
  int v27; // [rsp+B0h] [rbp-50h]
  __int64 v28; // [rsp+B8h] [rbp-48h]
  int v29; // [rsp+C0h] [rbp-40h]
  const wchar_t *v30; // [rsp+C8h] [rbp-38h]
  unsigned int *v31; // [rsp+D0h] [rbp-30h]
  int v32; // [rsp+D8h] [rbp-28h]
  int *v33; // [rsp+E0h] [rbp-20h]
  int v34; // [rsp+E8h] [rbp-18h]
  __int64 v35; // [rsp+F0h] [rbp-10h]
  int v36; // [rsp+F8h] [rbp-8h]
  const wchar_t *v37; // [rsp+100h] [rbp+0h]
  int *v38; // [rsp+108h] [rbp+8h]
  int v39; // [rsp+110h] [rbp+10h]
  int *v40; // [rsp+118h] [rbp+18h]
  int v41; // [rsp+120h] [rbp+20h]
  __int64 v42; // [rsp+128h] [rbp+28h]
  int v43; // [rsp+130h] [rbp+30h]
  const wchar_t *v44; // [rsp+138h] [rbp+38h]
  int *v45; // [rsp+140h] [rbp+40h]
  int v46; // [rsp+148h] [rbp+48h]
  int *v47; // [rsp+150h] [rbp+50h]
  int v48; // [rsp+158h] [rbp+58h]
  __int64 v49; // [rsp+160h] [rbp+60h]
  int v50; // [rsp+168h] [rbp+68h]
  const wchar_t *v51; // [rsp+170h] [rbp+70h]
  int *v52; // [rsp+178h] [rbp+78h]
  int v53; // [rsp+180h] [rbp+80h]
  int *v54; // [rsp+188h] [rbp+88h]
  int v55; // [rsp+190h] [rbp+90h]
  __int64 v56; // [rsp+198h] [rbp+98h]
  int v57; // [rsp+1A0h] [rbp+A0h]
  const wchar_t *v58; // [rsp+1A8h] [rbp+A8h]
  int *v59; // [rsp+1B0h] [rbp+B0h]
  int v60; // [rsp+1B8h] [rbp+B8h]
  int *v61; // [rsp+1C0h] [rbp+C0h]
  int v62; // [rsp+1C8h] [rbp+C8h]
  __int64 v63; // [rsp+1D0h] [rbp+D0h]
  int v64; // [rsp+1D8h] [rbp+D8h]
  const wchar_t *v65; // [rsp+1E0h] [rbp+E0h]
  unsigned int *v66; // [rsp+1E8h] [rbp+E8h]
  int v67; // [rsp+1F0h] [rbp+F0h]
  int *v68; // [rsp+1F8h] [rbp+F8h]
  int v69; // [rsp+200h] [rbp+100h]
  __int64 v70; // [rsp+208h] [rbp+108h]
  int v71; // [rsp+210h] [rbp+110h]
  const wchar_t *v72; // [rsp+218h] [rbp+118h]
  unsigned int *v73; // [rsp+220h] [rbp+120h]
  int v74; // [rsp+228h] [rbp+128h]
  int *v75; // [rsp+230h] [rbp+130h]
  int v76; // [rsp+238h] [rbp+138h]
  __int64 v77; // [rsp+240h] [rbp+140h]
  int v78; // [rsp+248h] [rbp+148h]
  const wchar_t *v79; // [rsp+250h] [rbp+150h]
  int *v80; // [rsp+258h] [rbp+158h]
  int v81; // [rsp+260h] [rbp+160h]
  int *v82; // [rsp+268h] [rbp+168h]
  int v83; // [rsp+270h] [rbp+170h]
  __int128 v84; // [rsp+278h] [rbp+178h]
  __int128 v85; // [rsp+288h] [rbp+188h]
  __int128 v86; // [rsp+298h] [rbp+198h]
  __int64 v87; // [rsp+2A8h] [rbp+1A8h]

  v12 = 25;
  v3 = 25;
  v14 = 0;
  v7 = 0;
  v13 = 1;
  v16 = 0x400000;
  v10 = 0x400000;
  v18 = 32;
  v4 = 32;
  v20 = -1;
  v11 = -1;
  v23 = L"BudgetThreshold";
  v24 = &v3;
  v26 = &v12;
  v30 = L"PagingQueueFenceIncrement";
  v31 = &v6;
  v33 = &v13;
  v37 = L"RestrictToPreferredSegment";
  v38 = &v7;
  v40 = &v14;
  v44 = L"Use64KPages";
  v45 = &v9;
  v47 = &v15;
  v51 = L"ExpandTo64KBAllocationSizeThreshold";
  v52 = &v10;
  v54 = &v16;
  v58 = L"AlwaysDecommitOnOffer";
  v6 = 1;
  v15 = 0;
  v9 = 0;
  v17 = 0;
  v8 = 0;
  v19 = 1;
  v5 = 1;
  v21 = 0LL;
  v22 = 288;
  v25 = 67108868;
  v27 = 4;
  v28 = 0LL;
  v29 = 288;
  v32 = 67108868;
  v34 = 4;
  v35 = 0LL;
  v36 = 288;
  v39 = 67108868;
  v41 = 4;
  v42 = 0LL;
  v43 = 288;
  v46 = 67108868;
  v48 = 4;
  v49 = 0LL;
  v50 = 288;
  v53 = 67108868;
  v55 = 4;
  v56 = 0LL;
  v57 = 288;
  v60 = 67108868;
  v62 = 4;
  v59 = &v8;
  v64 = 288;
  v61 = &v17;
  v65 = L"LazyDecommitChunkSizeMB";
  v66 = &v4;
  v68 = &v18;
  v72 = L"DecommitRepurposeMode";
  v73 = &v5;
  v75 = &v19;
  v79 = L"DxgMms2OfferReclaim";
  v80 = &v11;
  v82 = &v20;
  v67 = 67108868;
  v69 = 4;
  v71 = 288;
  v74 = 67108868;
  v76 = 4;
  v78 = 288;
  v81 = 67108868;
  v83 = 4;
  v87 = 0LL;
  v63 = 0LL;
  v70 = 0LL;
  v77 = 0LL;
  v84 = 0LL;
  v85 = 0LL;
  v86 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v21, 0LL, 0LL);
  v0 = 100;
  if ( v3 < 0x64 )
    v0 = v3;
  dword_1C00764E8 = v0;
  dword_1C00764EC = v6;
  if ( v6 > 0x51EB851 )
  {
    dword_1C00764EC = 85899345;
  }
  else if ( !v6 )
  {
    dword_1C00764EC = 1;
  }
  v1 = v11;
  dword_1C0076560 = v7;
  dword_1C007656C = v8;
  v2 = 512;
  if ( v4 < 0x200 )
    v2 = v4;
  dword_1C0076570 = v2;
  dword_1C0076574 = v5 < 3 ? v5 : 0;
  dword_1C0076564 = v9;
  dword_1C0076568 = v10;
  if ( (unsigned int)(v11 - 3) <= 0xFFFFFFFB )
    v1 = 0;
  dword_1C007653C = v1;
}