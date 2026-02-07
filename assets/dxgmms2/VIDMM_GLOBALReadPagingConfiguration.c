void VIDMM_GLOBAL::ReadPagingConfiguration(void)
{
  unsigned int v0; // edx
  int v1; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v2; // [rsp+34h] [rbp-CCh] BYREF
  unsigned int v3; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v4; // [rsp+3Ch] [rbp-C4h] BYREF
  unsigned int v5; // [rsp+40h] [rbp-C0h] BYREF
  unsigned int v6; // [rsp+44h] [rbp-BCh] BYREF
  int v7; // [rsp+48h] [rbp-B8h] BYREF
  int v8; // [rsp+4Ch] [rbp-B4h] BYREF
  int v9; // [rsp+50h] [rbp-B0h] BYREF
  unsigned int v10; // [rsp+54h] [rbp-ACh] BYREF
  int v11; // [rsp+58h] [rbp-A8h] BYREF
  int v12; // [rsp+5Ch] [rbp-A4h] BYREF
  int v13; // [rsp+60h] [rbp-A0h] BYREF
  int v14; // [rsp+64h] [rbp-9Ch] BYREF
  int v15; // [rsp+68h] [rbp-98h] BYREF
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
  __int64 v29; // [rsp+A0h] [rbp-60h] BYREF
  __int64 v30; // [rsp+A8h] [rbp-58h] BYREF
  __int64 v31; // [rsp+B0h] [rbp-50h]
  __int64 v32; // [rsp+C0h] [rbp-40h] BYREF
  int v33; // [rsp+C8h] [rbp-38h]
  const wchar_t *v34; // [rsp+D0h] [rbp-30h]
  int *v35; // [rsp+D8h] [rbp-28h]
  int v36; // [rsp+E0h] [rbp-20h]
  int *v37; // [rsp+E8h] [rbp-18h]
  int v38; // [rsp+F0h] [rbp-10h]
  __int64 v39; // [rsp+F8h] [rbp-8h]
  int v40; // [rsp+100h] [rbp+0h]
  const wchar_t *v41; // [rsp+108h] [rbp+8h]
  unsigned int *v42; // [rsp+110h] [rbp+10h]
  int v43; // [rsp+118h] [rbp+18h]
  int *v44; // [rsp+120h] [rbp+20h]
  int v45; // [rsp+128h] [rbp+28h]
  __int64 v46; // [rsp+130h] [rbp+30h]
  int v47; // [rsp+138h] [rbp+38h]
  const wchar_t *v48; // [rsp+140h] [rbp+40h]
  unsigned int *v49; // [rsp+148h] [rbp+48h]
  int v50; // [rsp+150h] [rbp+50h]
  int *v51; // [rsp+158h] [rbp+58h]
  int v52; // [rsp+160h] [rbp+60h]
  __int64 v53; // [rsp+168h] [rbp+68h]
  int v54; // [rsp+170h] [rbp+70h]
  const wchar_t *v55; // [rsp+178h] [rbp+78h]
  unsigned int *v56; // [rsp+180h] [rbp+80h]
  int v57; // [rsp+188h] [rbp+88h]
  int *v58; // [rsp+190h] [rbp+90h]
  int v59; // [rsp+198h] [rbp+98h]
  __int64 v60; // [rsp+1A0h] [rbp+A0h]
  int v61; // [rsp+1A8h] [rbp+A8h]
  const wchar_t *v62; // [rsp+1B0h] [rbp+B0h]
  unsigned int *v63; // [rsp+1B8h] [rbp+B8h]
  int v64; // [rsp+1C0h] [rbp+C0h]
  int *v65; // [rsp+1C8h] [rbp+C8h]
  int v66; // [rsp+1D0h] [rbp+D0h]
  __int64 v67; // [rsp+1D8h] [rbp+D8h]
  int v68; // [rsp+1E0h] [rbp+E0h]
  const wchar_t *v69; // [rsp+1E8h] [rbp+E8h]
  unsigned int *v70; // [rsp+1F0h] [rbp+F0h]
  int v71; // [rsp+1F8h] [rbp+F8h]
  int *v72; // [rsp+200h] [rbp+100h]
  int v73; // [rsp+208h] [rbp+108h]
  __int64 v74; // [rsp+210h] [rbp+110h]
  int v75; // [rsp+218h] [rbp+118h]
  const wchar_t *v76; // [rsp+220h] [rbp+120h]
  int *v77; // [rsp+228h] [rbp+128h]
  int v78; // [rsp+230h] [rbp+130h]
  int *v79; // [rsp+238h] [rbp+138h]
  int v80; // [rsp+240h] [rbp+140h]
  __int64 v81; // [rsp+248h] [rbp+148h]
  int v82; // [rsp+250h] [rbp+150h]
  const wchar_t *v83; // [rsp+258h] [rbp+158h]
  int *v84; // [rsp+260h] [rbp+160h]
  int v85; // [rsp+268h] [rbp+168h]
  int *v86; // [rsp+270h] [rbp+170h]
  int v87; // [rsp+278h] [rbp+178h]
  __int64 v88; // [rsp+280h] [rbp+180h]
  int v89; // [rsp+288h] [rbp+188h]
  const wchar_t *v90; // [rsp+290h] [rbp+190h]
  __int64 *v91; // [rsp+298h] [rbp+198h]
  int v92; // [rsp+2A0h] [rbp+1A0h]
  __int64 *v93; // [rsp+2A8h] [rbp+1A8h]
  int v94; // [rsp+2B0h] [rbp+1B0h]
  __int64 v95; // [rsp+2B8h] [rbp+1B8h]
  int v96; // [rsp+2C0h] [rbp+1C0h]
  const wchar_t *v97; // [rsp+2C8h] [rbp+1C8h]
  int *v98; // [rsp+2D0h] [rbp+1D0h]
  int v99; // [rsp+2D8h] [rbp+1D8h]
  int *v100; // [rsp+2E0h] [rbp+1E0h]
  int v101; // [rsp+2E8h] [rbp+1E8h]
  __int64 v102; // [rsp+2F0h] [rbp+1F0h]
  int v103; // [rsp+2F8h] [rbp+1F8h]
  const wchar_t *v104; // [rsp+300h] [rbp+200h]
  unsigned int *v105; // [rsp+308h] [rbp+208h]
  int v106; // [rsp+310h] [rbp+210h]
  int *v107; // [rsp+318h] [rbp+218h]
  int v108; // [rsp+320h] [rbp+220h]
  __int64 v109; // [rsp+328h] [rbp+228h]
  int v110; // [rsp+330h] [rbp+230h]
  const wchar_t *v111; // [rsp+338h] [rbp+238h]
  int *v112; // [rsp+340h] [rbp+240h]
  int v113; // [rsp+348h] [rbp+248h]
  int *v114; // [rsp+350h] [rbp+250h]
  int v115; // [rsp+358h] [rbp+258h]
  __int64 v116; // [rsp+360h] [rbp+260h]
  int v117; // [rsp+368h] [rbp+268h]
  const wchar_t *v118; // [rsp+370h] [rbp+270h]
  int *v119; // [rsp+378h] [rbp+278h]
  int v120; // [rsp+380h] [rbp+280h]
  int *v121; // [rsp+388h] [rbp+288h]
  int v122; // [rsp+390h] [rbp+290h]
  __int64 v123; // [rsp+398h] [rbp+298h]
  int v124; // [rsp+3A0h] [rbp+2A0h]
  const wchar_t *v125; // [rsp+3A8h] [rbp+2A8h]
  int *v126; // [rsp+3B0h] [rbp+2B0h]
  int v127; // [rsp+3B8h] [rbp+2B8h]
  int *v128; // [rsp+3C0h] [rbp+2C0h]
  int v129; // [rsp+3C8h] [rbp+2C8h]
  __int64 v130; // [rsp+3D0h] [rbp+2D0h]
  int v131; // [rsp+3D8h] [rbp+2D8h]
  const wchar_t *v132; // [rsp+3E0h] [rbp+2E0h]
  int *v133; // [rsp+3E8h] [rbp+2E8h]
  int v134; // [rsp+3F0h] [rbp+2F0h]
  int *v135; // [rsp+3F8h] [rbp+2F8h]
  int v136; // [rsp+400h] [rbp+300h]
  __int128 v137; // [rsp+408h] [rbp+308h]
  __int128 v138; // [rsp+418h] [rbp+318h]
  __int128 v139; // [rsp+428h] [rbp+328h]
  __int64 v140; // [rsp+438h] [rbp+338h]

  v30 = 16LL;
  v15 = 1;
  v1 = 1;
  v16 = 500;
  v3 = 500;
  v17 = 500;
  v4 = 500;
  v18 = 1000;
  v5 = 1000;
  v19 = 1000;
  v6 = 1000;
  v21 = 48;
  v7 = 48;
  v22 = 5000;
  v8 = 5000;
  v29 = 0x2000000LL;
  v31 = 0x2000000LL;
  v34 = L"DemotionWithinDeviceEnabled";
  v35 = &v1;
  v37 = &v15;
  v41 = L"DeviceSuspendPeriodMin";
  v42 = &v3;
  v44 = &v16;
  v48 = L"DeviceSuspendPeriodMax";
  v49 = &v4;
  v51 = &v17;
  v55 = L"DeviceResumePeriodMin";
  v20 = 50;
  v2 = 50;
  v23 = 50;
  v9 = 50;
  v24 = 1;
  v10 = 1;
  v25 = 1;
  v11 = 1;
  v56 = &v5;
  v26 = 0;
  v12 = 0;
  v27 = 0;
  v13 = 0;
  v28 = 0;
  v14 = 0;
  v32 = 0LL;
  v33 = 288;
  v36 = 67108868;
  v38 = 4;
  v39 = 0LL;
  v40 = 288;
  v43 = 67108868;
  v45 = 4;
  v46 = 0LL;
  v47 = 288;
  v50 = 67108868;
  v52 = 4;
  v53 = 0LL;
  v54 = 288;
  v57 = 67108868;
  v58 = &v18;
  v62 = L"DeviceResumePeriodMax";
  v63 = &v6;
  v65 = &v19;
  v69 = L"PagingQueueProcessingPeriodTime";
  v70 = &v2;
  v72 = &v20;
  v76 = L"InitialPromotionInterval";
  v77 = &v7;
  v79 = &v21;
  v83 = L"MaximumPromotionInterval";
  v84 = &v8;
  v86 = &v22;
  v90 = L"PromotionTargetSizePerInterval";
  v91 = &v30;
  v93 = &v29;
  v97 = L"PromotionNumberCapPerInterval";
  v98 = &v9;
  v100 = &v23;
  v104 = L"TransferFlushThreshold";
  v105 = &v10;
  v107 = &v24;
  v111 = L"EnableAsyncResidency";
  v112 = &v11;
  v114 = &v25;
  v59 = 4;
  v60 = 0LL;
  v61 = 288;
  v64 = 67108868;
  v66 = 4;
  v67 = 0LL;
  v68 = 288;
  v71 = 67108868;
  v73 = 4;
  v74 = 0LL;
  v75 = 288;
  v78 = 67108868;
  v80 = 4;
  v81 = 0LL;
  v82 = 288;
  v85 = 67108868;
  v87 = 4;
  v88 = 0LL;
  v89 = 288;
  v92 = 184549387;
  v94 = 8;
  v95 = 0LL;
  v96 = 288;
  v99 = 67108868;
  v101 = 4;
  v102 = 0LL;
  v103 = 288;
  v106 = 67108868;
  v108 = 4;
  v109 = 0LL;
  v110 = 288;
  v113 = 67108868;
  v115 = 4;
  v117 = 288;
  v118 = L"ForceUncommitGpuVAOnEvict";
  v119 = &v12;
  v121 = &v26;
  v125 = L"ForceSynchronousEvict";
  v126 = &v13;
  v128 = &v27;
  v132 = L"BreakOnPagingFailure";
  v133 = &v14;
  v135 = &v28;
  v120 = 67108868;
  v122 = 4;
  v124 = 288;
  v127 = 67108868;
  v129 = 4;
  v131 = 288;
  v134 = 67108868;
  v136 = 4;
  v140 = 0LL;
  v116 = 0LL;
  v123 = 0LL;
  v130 = 0LL;
  v137 = 0LL;
  v138 = 0LL;
  v139 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v32, 0LL, 0LL);
  v0 = v2;
  VIDMM_GLOBAL::_Config = (v1 != 0 ? 0x40 : 0) | VIDMM_GLOBAL::_Config & 0xFFFFFFBF;
  if ( v2 >= 0x12D )
  {
    v0 = 300;
  }
  else if ( v2 < 0x10 )
  {
    v0 = 16;
  }
  qword_1C00764F0 = 10000LL * v3;
  qword_1C00764F8 = 10000LL * v4;
  qword_1C0076500 = 10000LL * v5;
  qword_1C0076508 = 10000LL * v6;
  qword_1C0076528 = v31;
  dword_1C0076530 = v9;
  qword_1C0076510 = 10000LL * v0;
  qword_1C0076518 = (unsigned int)(10000 * v7);
  qword_1C0076550 = (unsigned __int64)v10 << 20;
  dword_1C007646C = v11;
  dword_1C0076534 = v12;
  dword_1C0076538 = v13;
  dword_1C00765C0 = v14;
  qword_1C0076520 = (unsigned int)(10000 * v8);
}