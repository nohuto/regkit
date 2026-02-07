void VIDMM_GLOBAL::ReadPagingConfiguration(void)
{
  unsigned int v0; // edx
  unsigned int v1; // [rsp+30h] [rbp-D0h] BYREF
  int v2; // [rsp+34h] [rbp-CCh] BYREF
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
  int v29; // [rsp+A0h] [rbp-60h] BYREF
  int v30; // [rsp+A4h] [rbp-5Ch] BYREF
  int v31; // [rsp+A8h] [rbp-58h] BYREF
  int v32; // [rsp+ACh] [rbp-54h] BYREF
  __int64 v33; // [rsp+B0h] [rbp-50h] BYREF
  __int64 v34; // [rsp+B8h] [rbp-48h] BYREF
  __int64 v35; // [rsp+C0h] [rbp-40h]
  __int64 v36; // [rsp+D0h] [rbp-30h] BYREF
  int v37; // [rsp+D8h] [rbp-28h]
  const wchar_t *v38; // [rsp+E0h] [rbp-20h]
  int *v39; // [rsp+E8h] [rbp-18h]
  int v40; // [rsp+F0h] [rbp-10h]
  int *v41; // [rsp+F8h] [rbp-8h]
  int v42; // [rsp+100h] [rbp+0h]
  __int64 v43; // [rsp+108h] [rbp+8h]
  int v44; // [rsp+110h] [rbp+10h]
  const wchar_t *v45; // [rsp+118h] [rbp+18h]
  unsigned int *v46; // [rsp+120h] [rbp+20h]
  int v47; // [rsp+128h] [rbp+28h]
  int *v48; // [rsp+130h] [rbp+30h]
  int v49; // [rsp+138h] [rbp+38h]
  __int64 v50; // [rsp+140h] [rbp+40h]
  int v51; // [rsp+148h] [rbp+48h]
  const wchar_t *v52; // [rsp+150h] [rbp+50h]
  unsigned int *v53; // [rsp+158h] [rbp+58h]
  int v54; // [rsp+160h] [rbp+60h]
  int *v55; // [rsp+168h] [rbp+68h]
  int v56; // [rsp+170h] [rbp+70h]
  __int64 v57; // [rsp+178h] [rbp+78h]
  int v58; // [rsp+180h] [rbp+80h]
  const wchar_t *v59; // [rsp+188h] [rbp+88h]
  unsigned int *v60; // [rsp+190h] [rbp+90h]
  int v61; // [rsp+198h] [rbp+98h]
  int *v62; // [rsp+1A0h] [rbp+A0h]
  int v63; // [rsp+1A8h] [rbp+A8h]
  __int64 v64; // [rsp+1B0h] [rbp+B0h]
  int v65; // [rsp+1B8h] [rbp+B8h]
  const wchar_t *v66; // [rsp+1C0h] [rbp+C0h]
  unsigned int *v67; // [rsp+1C8h] [rbp+C8h]
  int v68; // [rsp+1D0h] [rbp+D0h]
  int *v69; // [rsp+1D8h] [rbp+D8h]
  int v70; // [rsp+1E0h] [rbp+E0h]
  __int64 v71; // [rsp+1E8h] [rbp+E8h]
  int v72; // [rsp+1F0h] [rbp+F0h]
  const wchar_t *v73; // [rsp+1F8h] [rbp+F8h]
  unsigned int *v74; // [rsp+200h] [rbp+100h]
  int v75; // [rsp+208h] [rbp+108h]
  int *v76; // [rsp+210h] [rbp+110h]
  int v77; // [rsp+218h] [rbp+118h]
  __int64 v78; // [rsp+220h] [rbp+120h]
  int v79; // [rsp+228h] [rbp+128h]
  const wchar_t *v80; // [rsp+230h] [rbp+130h]
  int *v81; // [rsp+238h] [rbp+138h]
  int v82; // [rsp+240h] [rbp+140h]
  int *v83; // [rsp+248h] [rbp+148h]
  int v84; // [rsp+250h] [rbp+150h]
  __int64 v85; // [rsp+258h] [rbp+158h]
  int v86; // [rsp+260h] [rbp+160h]
  const wchar_t *v87; // [rsp+268h] [rbp+168h]
  int *v88; // [rsp+270h] [rbp+170h]
  int v89; // [rsp+278h] [rbp+178h]
  int *v90; // [rsp+280h] [rbp+180h]
  int v91; // [rsp+288h] [rbp+188h]
  __int64 v92; // [rsp+290h] [rbp+190h]
  int v93; // [rsp+298h] [rbp+198h]
  const wchar_t *v94; // [rsp+2A0h] [rbp+1A0h]
  int *v95; // [rsp+2A8h] [rbp+1A8h]
  int v96; // [rsp+2B0h] [rbp+1B0h]
  int *v97; // [rsp+2B8h] [rbp+1B8h]
  int v98; // [rsp+2C0h] [rbp+1C0h]
  __int64 v99; // [rsp+2C8h] [rbp+1C8h]
  int v100; // [rsp+2D0h] [rbp+1D0h]
  const wchar_t *v101; // [rsp+2D8h] [rbp+1D8h]
  __int64 *v102; // [rsp+2E0h] [rbp+1E0h]
  int v103; // [rsp+2E8h] [rbp+1E8h]
  __int64 *v104; // [rsp+2F0h] [rbp+1F0h]
  int v105; // [rsp+2F8h] [rbp+1F8h]
  __int64 v106; // [rsp+300h] [rbp+200h]
  int v107; // [rsp+308h] [rbp+208h]
  const wchar_t *v108; // [rsp+310h] [rbp+210h]
  int *v109; // [rsp+318h] [rbp+218h]
  int v110; // [rsp+320h] [rbp+220h]
  int *v111; // [rsp+328h] [rbp+228h]
  int v112; // [rsp+330h] [rbp+230h]
  __int64 v113; // [rsp+338h] [rbp+238h]
  int v114; // [rsp+340h] [rbp+240h]
  const wchar_t *v115; // [rsp+348h] [rbp+248h]
  unsigned int *v116; // [rsp+350h] [rbp+250h]
  int v117; // [rsp+358h] [rbp+258h]
  int *v118; // [rsp+360h] [rbp+260h]
  int v119; // [rsp+368h] [rbp+268h]
  __int64 v120; // [rsp+370h] [rbp+270h]
  int v121; // [rsp+378h] [rbp+278h]
  const wchar_t *v122; // [rsp+380h] [rbp+280h]
  int *v123; // [rsp+388h] [rbp+288h]
  int v124; // [rsp+390h] [rbp+290h]
  int *v125; // [rsp+398h] [rbp+298h]
  int v126; // [rsp+3A0h] [rbp+2A0h]
  __int64 v127; // [rsp+3A8h] [rbp+2A8h]
  int v128; // [rsp+3B0h] [rbp+2B0h]
  const wchar_t *v129; // [rsp+3B8h] [rbp+2B8h]
  int *v130; // [rsp+3C0h] [rbp+2C0h]
  int v131; // [rsp+3C8h] [rbp+2C8h]
  int *v132; // [rsp+3D0h] [rbp+2D0h]
  int v133; // [rsp+3D8h] [rbp+2D8h]
  __int64 v134; // [rsp+3E0h] [rbp+2E0h]
  int v135; // [rsp+3E8h] [rbp+2E8h]
  const wchar_t *v136; // [rsp+3F0h] [rbp+2F0h]
  int *v137; // [rsp+3F8h] [rbp+2F8h]
  int v138; // [rsp+400h] [rbp+300h]
  int *v139; // [rsp+408h] [rbp+308h]
  int v140; // [rsp+410h] [rbp+310h]
  __int64 v141; // [rsp+418h] [rbp+318h]
  int v142; // [rsp+420h] [rbp+320h]
  const wchar_t *v143; // [rsp+428h] [rbp+328h]
  int *v144; // [rsp+430h] [rbp+330h]
  int v145; // [rsp+438h] [rbp+338h]
  int *v146; // [rsp+440h] [rbp+340h]
  int v147; // [rsp+448h] [rbp+348h]
  __int64 v148; // [rsp+450h] [rbp+350h]
  int v149; // [rsp+458h] [rbp+358h]
  const wchar_t *v150; // [rsp+460h] [rbp+360h]
  int *v151; // [rsp+468h] [rbp+368h]
  int v152; // [rsp+470h] [rbp+370h]
  int *v153; // [rsp+478h] [rbp+378h]
  int v154; // [rsp+480h] [rbp+380h]
  __int128 v155; // [rsp+488h] [rbp+388h]
  __int128 v156; // [rsp+498h] [rbp+398h]
  __int128 v157; // [rsp+4A8h] [rbp+3A8h]
  __int64 v158; // [rsp+4B8h] [rbp+3B8h]

  v34 = 16LL;
  v29 = 0;
  v17 = 1;
  v18 = 500;
  v3 = 500;
  v19 = 500;
  v4 = 500;
  v20 = 1000;
  v5 = 1000;
  v21 = 1000;
  v6 = 1000;
  v24 = 48;
  v7 = 48;
  v25 = 5000;
  v8 = 5000;
  v33 = 0x2000000LL;
  v35 = 0x2000000LL;
  v38 = L"DemotionWithinDeviceEnabled";
  v39 = &v2;
  v41 = &v17;
  v45 = L"DeviceSuspendPeriodMin";
  v46 = &v3;
  v48 = &v18;
  v52 = L"DeviceSuspendPeriodMax";
  v53 = &v4;
  v55 = &v19;
  v2 = 1;
  v22 = 50;
  v1 = 50;
  v23 = 1;
  v16 = 1;
  v26 = 50;
  v9 = 50;
  v27 = 1;
  v10 = 1;
  v28 = 1;
  v11 = 1;
  v59 = L"DeviceResumePeriodMin";
  v12 = 0;
  v30 = 0;
  v13 = 0;
  v31 = 0;
  v14 = 0;
  v32 = 0;
  v15 = 0;
  v36 = 0LL;
  v37 = 288;
  v40 = 67108868;
  v42 = 4;
  v43 = 0LL;
  v44 = 288;
  v47 = 67108868;
  v49 = 4;
  v50 = 0LL;
  v51 = 288;
  v54 = 67108868;
  v56 = 4;
  v57 = 0LL;
  v58 = 288;
  v60 = &v5;
  v62 = &v20;
  v66 = L"DeviceResumePeriodMax";
  v67 = &v6;
  v69 = &v21;
  v73 = L"PagingQueueProcessingPeriodTime";
  v74 = &v1;
  v76 = &v22;
  v80 = L"EnablePromotion";
  v81 = &v16;
  v83 = &v23;
  v87 = L"InitialPromotionInterval";
  v88 = &v7;
  v90 = &v24;
  v94 = L"MaximumPromotionInterval";
  v95 = &v8;
  v97 = &v25;
  v101 = L"PromotionTargetSizePerInterval";
  v102 = &v34;
  v104 = &v33;
  v108 = L"PromotionNumberCapPerInterval";
  v109 = &v9;
  v111 = &v26;
  v115 = L"TransferFlushThreshold";
  v116 = &v10;
  v61 = 67108868;
  v63 = 4;
  v64 = 0LL;
  v65 = 288;
  v68 = 67108868;
  v70 = 4;
  v71 = 0LL;
  v72 = 288;
  v75 = 67108868;
  v77 = 4;
  v78 = 0LL;
  v79 = 288;
  v82 = 67108868;
  v84 = 4;
  v85 = 0LL;
  v86 = 288;
  v89 = 67108868;
  v91 = 4;
  v92 = 0LL;
  v93 = 288;
  v96 = 67108868;
  v98 = 4;
  v99 = 0LL;
  v100 = 288;
  v103 = 184549387;
  v105 = 8;
  v106 = 0LL;
  v107 = 288;
  v110 = 67108868;
  v112 = 4;
  v113 = 0LL;
  v114 = 288;
  v117 = 67108868;
  v119 = 4;
  v118 = &v27;
  v121 = 288;
  v122 = L"EnableAsyncResidency";
  v123 = &v11;
  v125 = &v28;
  v129 = L"ForceUncommitGpuVAOnEvict";
  v130 = &v12;
  v132 = &v29;
  v136 = L"ForceSynchronousEvict";
  v137 = &v13;
  v139 = &v30;
  v143 = L"BreakOnPagingFailure";
  v144 = &v14;
  v146 = &v31;
  v150 = L"TemporaryResourcePolicy";
  v151 = &v15;
  v153 = &v32;
  v124 = 67108868;
  v126 = 4;
  v128 = 288;
  v131 = 67108868;
  v133 = 4;
  v135 = 288;
  v138 = 67108868;
  v140 = 4;
  v142 = 288;
  v145 = 67108868;
  v147 = 4;
  v149 = 288;
  v152 = 67108868;
  v154 = 4;
  v158 = 0LL;
  v120 = 0LL;
  v127 = 0LL;
  v134 = 0LL;
  v141 = 0LL;
  v148 = 0LL;
  v155 = 0LL;
  v156 = 0LL;
  v157 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v36, 0LL, 0LL);
  v0 = v1;
  VIDMM_GLOBAL::_Config = (v2 != 0 ? 0x40 : 0) | VIDMM_GLOBAL::_Config & 0xFFFFFFBF;
  if ( v1 < 0x12D )
  {
    if ( v1 >= 0x10 )
      goto LABEL_6;
    v0 = 16;
  }
  else
  {
    v0 = 300;
  }
  v1 = v0;
LABEL_6:
  qword_140081598 = 10000LL * v3;
  qword_1400815A0 = 10000LL * v4;
  qword_1400815A8 = 10000LL * v5;
  qword_1400815B0 = 10000LL * v6;
  qword_1400815D0 = v35;
  dword_1400815D8 = v9;
  qword_1400815B8 = 10000LL * v0;
  qword_1400815C0 = (unsigned int)(10000 * v7);
  qword_140081600 = (unsigned __int64)v10 << 20;
  dword_14008150C = v11;
  dword_1400815E0 = v12;
  dword_1400815E4 = v13;
  dword_140081668 = v14;
  dword_14008166C = v15;
  qword_1400815C8 = (unsigned int)(10000 * v8);
  if ( (unsigned int)Feature_Servicing_GraphicsKernel_PromotionRegistryKey__private_IsEnabledDeviceUsageNoInline() )
    byte_1400815DC = v16 != 0;
}