void VIDMM_GLOBAL::ReadHeapConfiguration(void)
{
  int v0; // edx
  int v1; // eax
  int v2; // [rsp+30h] [rbp-D0h] BYREF
  int v3; // [rsp+34h] [rbp-CCh] BYREF
  int v4; // [rsp+38h] [rbp-C8h] BYREF
  int v5; // [rsp+3Ch] [rbp-C4h] BYREF
  int v6; // [rsp+40h] [rbp-C0h] BYREF
  int v7; // [rsp+44h] [rbp-BCh] BYREF
  int v8; // [rsp+48h] [rbp-B8h] BYREF
  int v9; // [rsp+4Ch] [rbp-B4h] BYREF
  int v10; // [rsp+50h] [rbp-B0h] BYREF
  int v11; // [rsp+54h] [rbp-ACh] BYREF
  int v12; // [rsp+58h] [rbp-A8h] BYREF
  int v13; // [rsp+5Ch] [rbp-A4h] BYREF
  int v14; // [rsp+60h] [rbp-A0h] BYREF
  int v15; // [rsp+64h] [rbp-9Ch] BYREF
  int v16; // [rsp+68h] [rbp-98h] BYREF
  int v17; // [rsp+6Ch] [rbp-94h] BYREF
  int v18; // [rsp+70h] [rbp-90h] BYREF
  int v19; // [rsp+74h] [rbp-8Ch] BYREF
  int v20; // [rsp+78h] [rbp-88h] BYREF
  int v21; // [rsp+7Ch] [rbp-84h] BYREF
  int v22; // [rsp+80h] [rbp-80h] BYREF
  int v23; // [rsp+84h] [rbp-7Ch] BYREF
  int v24; // [rsp+88h] [rbp-78h] BYREF
  int v25; // [rsp+8Ch] [rbp-74h] BYREF
  int v26; // [rsp+90h] [rbp-70h] BYREF
  int v27; // [rsp+94h] [rbp-6Ch] BYREF
  int v28; // [rsp+98h] [rbp-68h] BYREF
  int v29; // [rsp+9Ch] [rbp-64h] BYREF
  __int64 v30; // [rsp+A0h] [rbp-60h] BYREF
  int v31; // [rsp+A8h] [rbp-58h]
  const wchar_t *v32; // [rsp+B0h] [rbp-50h]
  int *v33; // [rsp+B8h] [rbp-48h]
  int v34; // [rsp+C0h] [rbp-40h]
  int *v35; // [rsp+C8h] [rbp-38h]
  int v36; // [rsp+D0h] [rbp-30h]
  __int64 v37; // [rsp+D8h] [rbp-28h]
  int v38; // [rsp+E0h] [rbp-20h]
  const wchar_t *v39; // [rsp+E8h] [rbp-18h]
  int *v40; // [rsp+F0h] [rbp-10h]
  int v41; // [rsp+F8h] [rbp-8h]
  int *v42; // [rsp+100h] [rbp+0h]
  int v43; // [rsp+108h] [rbp+8h]
  __int64 v44; // [rsp+110h] [rbp+10h]
  int v45; // [rsp+118h] [rbp+18h]
  const wchar_t *v46; // [rsp+120h] [rbp+20h]
  int *v47; // [rsp+128h] [rbp+28h]
  int v48; // [rsp+130h] [rbp+30h]
  int *v49; // [rsp+138h] [rbp+38h]
  int v50; // [rsp+140h] [rbp+40h]
  __int64 v51; // [rsp+148h] [rbp+48h]
  int v52; // [rsp+150h] [rbp+50h]
  const wchar_t *v53; // [rsp+158h] [rbp+58h]
  int *v54; // [rsp+160h] [rbp+60h]
  int v55; // [rsp+168h] [rbp+68h]
  int *v56; // [rsp+170h] [rbp+70h]
  int v57; // [rsp+178h] [rbp+78h]
  __int64 v58; // [rsp+180h] [rbp+80h]
  int v59; // [rsp+188h] [rbp+88h]
  const wchar_t *v60; // [rsp+190h] [rbp+90h]
  int *v61; // [rsp+198h] [rbp+98h]
  int v62; // [rsp+1A0h] [rbp+A0h]
  int *v63; // [rsp+1A8h] [rbp+A8h]
  int v64; // [rsp+1B0h] [rbp+B0h]
  __int64 v65; // [rsp+1B8h] [rbp+B8h]
  int v66; // [rsp+1C0h] [rbp+C0h]
  const wchar_t *v67; // [rsp+1C8h] [rbp+C8h]
  int *v68; // [rsp+1D0h] [rbp+D0h]
  int v69; // [rsp+1D8h] [rbp+D8h]
  int *v70; // [rsp+1E0h] [rbp+E0h]
  int v71; // [rsp+1E8h] [rbp+E8h]
  __int64 v72; // [rsp+1F0h] [rbp+F0h]
  int v73; // [rsp+1F8h] [rbp+F8h]
  const wchar_t *v74; // [rsp+200h] [rbp+100h]
  int *v75; // [rsp+208h] [rbp+108h]
  int v76; // [rsp+210h] [rbp+110h]
  int *v77; // [rsp+218h] [rbp+118h]
  int v78; // [rsp+220h] [rbp+120h]
  __int64 v79; // [rsp+228h] [rbp+128h]
  int v80; // [rsp+230h] [rbp+130h]
  const wchar_t *v81; // [rsp+238h] [rbp+138h]
  int *v82; // [rsp+240h] [rbp+140h]
  int v83; // [rsp+248h] [rbp+148h]
  int *v84; // [rsp+250h] [rbp+150h]
  int v85; // [rsp+258h] [rbp+158h]
  __int64 v86; // [rsp+260h] [rbp+160h]
  int v87; // [rsp+268h] [rbp+168h]
  const wchar_t *v88; // [rsp+270h] [rbp+170h]
  int *v89; // [rsp+278h] [rbp+178h]
  int v90; // [rsp+280h] [rbp+180h]
  int *v91; // [rsp+288h] [rbp+188h]
  int v92; // [rsp+290h] [rbp+190h]
  __int64 v93; // [rsp+298h] [rbp+198h]
  int v94; // [rsp+2A0h] [rbp+1A0h]
  const wchar_t *v95; // [rsp+2A8h] [rbp+1A8h]
  int *v96; // [rsp+2B0h] [rbp+1B0h]
  int v97; // [rsp+2B8h] [rbp+1B8h]
  int *v98; // [rsp+2C0h] [rbp+1C0h]
  int v99; // [rsp+2C8h] [rbp+1C8h]
  __int64 v100; // [rsp+2D0h] [rbp+1D0h]
  int v101; // [rsp+2D8h] [rbp+1D8h]
  const wchar_t *v102; // [rsp+2E0h] [rbp+1E0h]
  int *v103; // [rsp+2E8h] [rbp+1E8h]
  int v104; // [rsp+2F0h] [rbp+1F0h]
  int *v105; // [rsp+2F8h] [rbp+1F8h]
  int v106; // [rsp+300h] [rbp+200h]
  __int64 v107; // [rsp+308h] [rbp+208h]
  int v108; // [rsp+310h] [rbp+210h]
  const wchar_t *v109; // [rsp+318h] [rbp+218h]
  int *v110; // [rsp+320h] [rbp+220h]
  int v111; // [rsp+328h] [rbp+228h]
  int *v112; // [rsp+330h] [rbp+230h]
  int v113; // [rsp+338h] [rbp+238h]
  __int64 v114; // [rsp+340h] [rbp+240h]
  int v115; // [rsp+348h] [rbp+248h]
  const wchar_t *v116; // [rsp+350h] [rbp+250h]
  int *v117; // [rsp+358h] [rbp+258h]
  int v118; // [rsp+360h] [rbp+260h]
  int *v119; // [rsp+368h] [rbp+268h]
  int v120; // [rsp+370h] [rbp+270h]
  __int64 v121; // [rsp+378h] [rbp+278h]
  int v122; // [rsp+380h] [rbp+280h]
  const wchar_t *v123; // [rsp+388h] [rbp+288h]
  int *v124; // [rsp+390h] [rbp+290h]
  int v125; // [rsp+398h] [rbp+298h]
  int *v126; // [rsp+3A0h] [rbp+2A0h]
  int v127; // [rsp+3A8h] [rbp+2A8h]
  __int128 v128; // [rsp+3B0h] [rbp+2B0h]
  __int128 v129; // [rsp+3C0h] [rbp+2C0h]
  __int128 v130; // [rsp+3D0h] [rbp+2D0h]
  __int64 v131; // [rsp+3E0h] [rbp+2E0h]

  v0 = 256;
  v16 = 1;
  v2 = 1;
  v22 = 1;
  v17 = 15;
  v3 = 15;
  v18 = 15;
  v4 = 15;
  v20 = 32;
  v6 = 32;
  v21 = 1024;
  v7 = 1024;
  v24 = 8;
  v10 = 8;
  v1 = 256;
  if ( (unsigned __int64)qword_1C0076288 <= 0x53333333 )
    v1 = 64;
  v8 = 1;
  v26 = v1;
  if ( (unsigned __int64)qword_1C0076288 <= 0x53333333 )
    v0 = 64;
  v12 = v1;
  v27 = v0;
  v13 = v0;
  v32 = L"DebouncedPageManagement";
  v19 = 4;
  v33 = &v2;
  v35 = &v16;
  v39 = L"DebouncedUnlockAge";
  v40 = &v3;
  v42 = &v17;
  v46 = L"DebouncedDecommitAge";
  v47 = &v4;
  v49 = &v18;
  v53 = L"RecycleHeapPackingThreshold";
  v54 = &v5;
  v5 = 4;
  v23 = 4;
  v9 = 4;
  v25 = 64;
  v11 = 64;
  v28 = 0;
  v14 = 0;
  v29 = 64;
  v15 = 64;
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
  v56 = &v19;
  v60 = L"RecycleHeapPackingBlockSize";
  v61 = &v6;
  v63 = &v20;
  v67 = L"RecycleHeapPTDBlockSize";
  v68 = &v7;
  v70 = &v21;
  v74 = L"ZeroedRecyclePages";
  v75 = &v8;
  v77 = &v22;
  v81 = L"LeanRecycleHeapPackingThreshold";
  v82 = &v9;
  v84 = &v23;
  v88 = L"LeanRecycleHeapPackingBlockSize";
  v89 = &v10;
  v91 = &v24;
  v95 = L"LeanRecycleHeapPTDBlockSize";
  v96 = &v11;
  v98 = &v25;
  v102 = L"MaximumDecommitDebounce";
  v103 = &v12;
  v105 = &v26;
  v109 = L"MaximumUnlockDebounce";
  v110 = &v13;
  v58 = 0LL;
  v59 = 288;
  v62 = 67108868;
  v64 = 4;
  v65 = 0LL;
  v66 = 288;
  v69 = 67108868;
  v71 = 4;
  v72 = 0LL;
  v73 = 288;
  v76 = 67108868;
  v78 = 4;
  v79 = 0LL;
  v80 = 288;
  v83 = 67108868;
  v85 = 4;
  v86 = 0LL;
  v87 = 288;
  v90 = 67108868;
  v92 = 4;
  v93 = 0LL;
  v94 = 288;
  v97 = 67108868;
  v99 = 4;
  v100 = 0LL;
  v101 = 288;
  v104 = 67108868;
  v106 = 4;
  v107 = 0LL;
  v108 = 288;
  v111 = 67108868;
  v112 = &v27;
  v114 = 0LL;
  v116 = L"RecycleHistory";
  v115 = 288;
  v117 = &v14;
  v119 = &v28;
  v123 = L"RecycleHistorySize";
  v124 = &v15;
  v126 = &v29;
  v118 = 67108868;
  v121 = 0LL;
  v122 = 288;
  v125 = 67108868;
  v113 = 4;
  v120 = 4;
  v127 = 4;
  v128 = 0LL;
  v131 = 0LL;
  v129 = 0LL;
  v130 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v30, 0LL, 0LL);
  VIDMM_GLOBAL::_Config &= ~0x200u;
  dword_1C0076478 = v2;
  dword_1C007647C = v3;
  dword_1C0076480 = v4;
  dword_1C0076484 = v5;
  dword_1C0076488 = v6;
  dword_1C007648C = v7;
  dword_1C0076490 = v8;
  dword_1C0076494 = v9;
  dword_1C0076498 = v10;
  dword_1C007649C = v11;
  dword_1C00764A0 = v12;
  dword_1C00764A4 = v13;
  dword_1C00764A8 = v14;
  LODWORD(dword_1C00764AC) = v15;
}