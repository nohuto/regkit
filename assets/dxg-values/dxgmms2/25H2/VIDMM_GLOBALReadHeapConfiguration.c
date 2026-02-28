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
  unsigned __int64 v15; // [rsp+64h] [rbp-9Ch] BYREF
  int v16; // [rsp+6Ch] [rbp-94h] BYREF
  int v17; // [rsp+70h] [rbp-90h] BYREF
  int v18; // [rsp+74h] [rbp-8Ch] BYREF
  int v19; // [rsp+78h] [rbp-88h] BYREF
  int v20; // [rsp+7Ch] [rbp-84h] BYREF
  int v21; // [rsp+80h] [rbp-80h] BYREF
  int v22; // [rsp+84h] [rbp-7Ch] BYREF
  int v23; // [rsp+88h] [rbp-78h] BYREF
  int v24; // [rsp+8Ch] [rbp-74h] BYREF
  int v25; // [rsp+90h] [rbp-70h] BYREF
  int v26; // [rsp+94h] [rbp-6Ch] BYREF
  int v27; // [rsp+98h] [rbp-68h] BYREF
  int v28; // [rsp+9Ch] [rbp-64h] BYREF
  int v29; // [rsp+A0h] [rbp-60h] BYREF
  int v30; // [rsp+A4h] [rbp-5Ch] BYREF
  __int64 v31; // [rsp+B0h] [rbp-50h] BYREF
  int v32; // [rsp+B8h] [rbp-48h]
  const wchar_t *v33; // [rsp+C0h] [rbp-40h]
  int *v34; // [rsp+C8h] [rbp-38h]
  int v35; // [rsp+D0h] [rbp-30h]
  int *v36; // [rsp+D8h] [rbp-28h]
  int v37; // [rsp+E0h] [rbp-20h]
  __int64 v38; // [rsp+E8h] [rbp-18h]
  int v39; // [rsp+F0h] [rbp-10h]
  const wchar_t *v40; // [rsp+F8h] [rbp-8h]
  int *v41; // [rsp+100h] [rbp+0h]
  int v42; // [rsp+108h] [rbp+8h]
  int *v43; // [rsp+110h] [rbp+10h]
  int v44; // [rsp+118h] [rbp+18h]
  __int64 v45; // [rsp+120h] [rbp+20h]
  int v46; // [rsp+128h] [rbp+28h]
  const wchar_t *v47; // [rsp+130h] [rbp+30h]
  int *v48; // [rsp+138h] [rbp+38h]
  int v49; // [rsp+140h] [rbp+40h]
  int *v50; // [rsp+148h] [rbp+48h]
  int v51; // [rsp+150h] [rbp+50h]
  __int64 v52; // [rsp+158h] [rbp+58h]
  int v53; // [rsp+160h] [rbp+60h]
  const wchar_t *v54; // [rsp+168h] [rbp+68h]
  int *v55; // [rsp+170h] [rbp+70h]
  int v56; // [rsp+178h] [rbp+78h]
  int *v57; // [rsp+180h] [rbp+80h]
  int v58; // [rsp+188h] [rbp+88h]
  __int64 v59; // [rsp+190h] [rbp+90h]
  int v60; // [rsp+198h] [rbp+98h]
  const wchar_t *v61; // [rsp+1A0h] [rbp+A0h]
  int *v62; // [rsp+1A8h] [rbp+A8h]
  int v63; // [rsp+1B0h] [rbp+B0h]
  int *v64; // [rsp+1B8h] [rbp+B8h]
  int v65; // [rsp+1C0h] [rbp+C0h]
  __int64 v66; // [rsp+1C8h] [rbp+C8h]
  int v67; // [rsp+1D0h] [rbp+D0h]
  const wchar_t *v68; // [rsp+1D8h] [rbp+D8h]
  int *v69; // [rsp+1E0h] [rbp+E0h]
  int v70; // [rsp+1E8h] [rbp+E8h]
  int *v71; // [rsp+1F0h] [rbp+F0h]
  int v72; // [rsp+1F8h] [rbp+F8h]
  __int64 v73; // [rsp+200h] [rbp+100h]
  int v74; // [rsp+208h] [rbp+108h]
  const wchar_t *v75; // [rsp+210h] [rbp+110h]
  int *v76; // [rsp+218h] [rbp+118h]
  int v77; // [rsp+220h] [rbp+120h]
  int *v78; // [rsp+228h] [rbp+128h]
  int v79; // [rsp+230h] [rbp+130h]
  __int64 v80; // [rsp+238h] [rbp+138h]
  int v81; // [rsp+240h] [rbp+140h]
  const wchar_t *v82; // [rsp+248h] [rbp+148h]
  int *v83; // [rsp+250h] [rbp+150h]
  int v84; // [rsp+258h] [rbp+158h]
  int *v85; // [rsp+260h] [rbp+160h]
  int v86; // [rsp+268h] [rbp+168h]
  __int64 v87; // [rsp+270h] [rbp+170h]
  int v88; // [rsp+278h] [rbp+178h]
  const wchar_t *v89; // [rsp+280h] [rbp+180h]
  int *v90; // [rsp+288h] [rbp+188h]
  int v91; // [rsp+290h] [rbp+190h]
  int *v92; // [rsp+298h] [rbp+198h]
  int v93; // [rsp+2A0h] [rbp+1A0h]
  __int64 v94; // [rsp+2A8h] [rbp+1A8h]
  int v95; // [rsp+2B0h] [rbp+1B0h]
  const wchar_t *v96; // [rsp+2B8h] [rbp+1B8h]
  int *v97; // [rsp+2C0h] [rbp+1C0h]
  int v98; // [rsp+2C8h] [rbp+1C8h]
  int *v99; // [rsp+2D0h] [rbp+1D0h]
  int v100; // [rsp+2D8h] [rbp+1D8h]
  __int64 v101; // [rsp+2E0h] [rbp+1E0h]
  int v102; // [rsp+2E8h] [rbp+1E8h]
  const wchar_t *v103; // [rsp+2F0h] [rbp+1F0h]
  int *v104; // [rsp+2F8h] [rbp+1F8h]
  int v105; // [rsp+300h] [rbp+200h]
  int *v106; // [rsp+308h] [rbp+208h]
  int v107; // [rsp+310h] [rbp+210h]
  __int64 v108; // [rsp+318h] [rbp+218h]
  int v109; // [rsp+320h] [rbp+220h]
  const wchar_t *v110; // [rsp+328h] [rbp+228h]
  int *v111; // [rsp+330h] [rbp+230h]
  int v112; // [rsp+338h] [rbp+238h]
  int *v113; // [rsp+340h] [rbp+240h]
  int v114; // [rsp+348h] [rbp+248h]
  __int64 v115; // [rsp+350h] [rbp+250h]
  int v116; // [rsp+358h] [rbp+258h]
  const wchar_t *v117; // [rsp+360h] [rbp+260h]
  int *v118; // [rsp+368h] [rbp+268h]
  int v119; // [rsp+370h] [rbp+270h]
  int *v120; // [rsp+378h] [rbp+278h]
  int v121; // [rsp+380h] [rbp+280h]
  __int64 v122; // [rsp+388h] [rbp+288h]
  int v123; // [rsp+390h] [rbp+290h]
  const wchar_t *v124; // [rsp+398h] [rbp+298h]
  unsigned __int64 *v125; // [rsp+3A0h] [rbp+2A0h]
  int v126; // [rsp+3A8h] [rbp+2A8h]
  int *v127; // [rsp+3B0h] [rbp+2B0h]
  int v128; // [rsp+3B8h] [rbp+2B8h]
  __int64 v129; // [rsp+3C0h] [rbp+2C0h]
  int v130; // [rsp+3C8h] [rbp+2C8h]
  const wchar_t *v131; // [rsp+3D0h] [rbp+2D0h]
  char *v132; // [rsp+3D8h] [rbp+2D8h]
  int v133; // [rsp+3E0h] [rbp+2E0h]
  int *v134; // [rsp+3E8h] [rbp+2E8h]
  int v135; // [rsp+3F0h] [rbp+2F0h]
  __int128 v136; // [rsp+3F8h] [rbp+2F8h]
  __int128 v137; // [rsp+408h] [rbp+308h]
  __int128 v138; // [rsp+418h] [rbp+318h]
  __int64 v139; // [rsp+428h] [rbp+328h]

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
  if ( (unsigned __int64)qword_140081328 <= 0x53333333 )
    v1 = 64;
  v8 = 1;
  v26 = v1;
  if ( (unsigned __int64)qword_140081328 <= 0x53333333 )
    v0 = 64;
  v12 = v1;
  v27 = v0;
  v30 = 0x200000;
  v33 = L"DebouncedPageManagement";
  v34 = &v2;
  v36 = &v16;
  v40 = L"DebouncedUnlockAge";
  v41 = &v3;
  v43 = &v17;
  v47 = L"DebouncedDecommitAge";
  v48 = &v4;
  v50 = &v18;
  v13 = v0;
  v54 = L"RecycleHeapPackingThreshold";
  v19 = 4;
  v5 = 4;
  v23 = 4;
  v9 = 4;
  v25 = 64;
  v11 = 64;
  v28 = 0;
  v14 = 0;
  v29 = 64;
  v15 = 0x20000000000040LL;
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
  v55 = &v5;
  v57 = &v19;
  v61 = L"RecycleHeapPackingBlockSize";
  v62 = &v6;
  v64 = &v20;
  v68 = L"RecycleHeapPTDBlockSize";
  v69 = &v7;
  v71 = &v21;
  v75 = L"ZeroedRecyclePages";
  v76 = &v8;
  v78 = &v22;
  v82 = L"LeanRecycleHeapPackingThreshold";
  v83 = &v9;
  v85 = &v23;
  v89 = L"LeanRecycleHeapPackingBlockSize";
  v90 = &v10;
  v92 = &v24;
  v96 = L"LeanRecycleHeapPTDBlockSize";
  v97 = &v11;
  v99 = &v25;
  v103 = L"MaximumDecommitDebounce";
  v104 = &v12;
  v106 = &v26;
  v110 = L"MaximumUnlockDebounce";
  v58 = 4;
  v59 = 0LL;
  v60 = 288;
  v63 = 67108868;
  v65 = 4;
  v66 = 0LL;
  v67 = 288;
  v70 = 67108868;
  v72 = 4;
  v73 = 0LL;
  v74 = 288;
  v77 = 67108868;
  v79 = 4;
  v80 = 0LL;
  v81 = 288;
  v84 = 67108868;
  v86 = 4;
  v87 = 0LL;
  v88 = 288;
  v91 = 67108868;
  v93 = 4;
  v94 = 0LL;
  v95 = 288;
  v98 = 67108868;
  v100 = 4;
  v101 = 0LL;
  v102 = 288;
  v105 = 67108868;
  v107 = 4;
  v108 = 0LL;
  v109 = 288;
  v111 = &v13;
  v112 = 67108868;
  v113 = &v27;
  v115 = 0LL;
  v117 = L"RecycleHistory";
  v118 = &v14;
  v120 = &v28;
  v124 = L"RecycleHistorySize";
  v125 = &v15;
  v127 = &v29;
  v131 = L"ZeroPageLockThreshold";
  v132 = (char *)&v15 + 4;
  v134 = &v30;
  v116 = 288;
  v119 = 67108868;
  v122 = 0LL;
  v123 = 288;
  v126 = 67108868;
  v129 = 0LL;
  v130 = 288;
  v133 = 67108868;
  v114 = 4;
  v121 = 4;
  v128 = 4;
  v135 = 4;
  v136 = 0LL;
  v139 = 0LL;
  v137 = 0LL;
  v138 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v31, 0LL, 0LL);
  VIDMM_GLOBAL::_Config &= ~0x200u;
  dword_140081518 = v2;
  dword_14008151C = v3;
  dword_140081520 = v4;
  dword_140081524 = v5;
  dword_140081528 = v6;
  dword_14008152C = v7;
  dword_140081530 = v8;
  dword_140081534 = v9;
  dword_140081538 = v10;
  dword_14008153C = v11;
  dword_140081540 = v12;
  dword_140081544 = v13;
  dword_140081548 = v14;
  qword_14008154C = v15;
}