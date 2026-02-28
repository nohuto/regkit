void VIDMM_GLOBAL::ReadBudgetConfiguration(void)
{
  int v0; // ebx
  unsigned int v1; // r8d
  unsigned int v2; // eax
  unsigned int v3; // eax
  unsigned int v4; // edx
  unsigned int v5; // ecx
  int v6; // eax
  int v7; // eax
  unsigned __int64 v8; // rax
  unsigned __int64 v9; // rax
  int v10; // eax
  int v11; // edx
  unsigned int v12; // eax
  int v13; // eax
  int v14; // edx
  unsigned int v15; // eax
  int v16; // eax
  int v17; // edx
  unsigned int v18; // eax
  unsigned int v19; // [rsp+38h] [rbp-D0h] BYREF
  unsigned int v20; // [rsp+3Ch] [rbp-CCh] BYREF
  unsigned int v21; // [rsp+40h] [rbp-C8h] BYREF
  unsigned int v22; // [rsp+44h] [rbp-C4h] BYREF
  unsigned int v23; // [rsp+48h] [rbp-C0h] BYREF
  int v24; // [rsp+4Ch] [rbp-BCh] BYREF
  unsigned int v25; // [rsp+50h] [rbp-B8h] BYREF
  unsigned int v26; // [rsp+54h] [rbp-B4h] BYREF
  unsigned int v27; // [rsp+58h] [rbp-B0h] BYREF
  int v28; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v29; // [rsp+60h] [rbp-A8h] BYREF
  unsigned int v30; // [rsp+64h] [rbp-A4h] BYREF
  unsigned int v31; // [rsp+68h] [rbp-A0h] BYREF
  unsigned int v32; // [rsp+6Ch] [rbp-9Ch] BYREF
  unsigned int v33; // [rsp+70h] [rbp-98h] BYREF
  unsigned int v34; // [rsp+74h] [rbp-94h] BYREF
  unsigned int v35; // [rsp+78h] [rbp-90h] BYREF
  unsigned int v36; // [rsp+7Ch] [rbp-8Ch] BYREF
  int v37; // [rsp+80h] [rbp-88h] BYREF
  int v38; // [rsp+84h] [rbp-84h] BYREF
  int v39; // [rsp+88h] [rbp-80h] BYREF
  int v40; // [rsp+8Ch] [rbp-7Ch] BYREF
  int v41; // [rsp+90h] [rbp-78h] BYREF
  int v42; // [rsp+94h] [rbp-74h] BYREF
  int v43; // [rsp+98h] [rbp-70h] BYREF
  int v44; // [rsp+9Ch] [rbp-6Ch] BYREF
  int v45; // [rsp+A0h] [rbp-68h] BYREF
  int v46; // [rsp+A4h] [rbp-64h] BYREF
  int v47; // [rsp+A8h] [rbp-60h] BYREF
  int v48; // [rsp+ACh] [rbp-5Ch] BYREF
  int v49; // [rsp+B0h] [rbp-58h] BYREF
  int v50; // [rsp+B4h] [rbp-54h] BYREF
  int v51; // [rsp+B8h] [rbp-50h] BYREF
  int v52; // [rsp+BCh] [rbp-4Ch] BYREF
  int v53; // [rsp+C0h] [rbp-48h] BYREF
  int v54; // [rsp+C4h] [rbp-44h] BYREF
  __int64 v55; // [rsp+C8h] [rbp-40h] BYREF
  __int64 v56; // [rsp+D0h] [rbp-38h] BYREF
  __int64 v57; // [rsp+D8h] [rbp-30h]
  __int64 v58; // [rsp+E8h] [rbp-20h] BYREF
  int v59; // [rsp+F0h] [rbp-18h]
  const wchar_t *v60; // [rsp+F8h] [rbp-10h]
  __int64 *v61; // [rsp+100h] [rbp-8h]
  int v62; // [rsp+108h] [rbp+0h]
  __int64 *v63; // [rsp+110h] [rbp+8h]
  int v64; // [rsp+118h] [rbp+10h]
  __int64 v65; // [rsp+120h] [rbp+18h]
  int v66; // [rsp+128h] [rbp+20h]
  const wchar_t *v67; // [rsp+130h] [rbp+28h]
  int *v68; // [rsp+138h] [rbp+30h]
  int v69; // [rsp+140h] [rbp+38h]
  int *v70; // [rsp+148h] [rbp+40h]
  int v71; // [rsp+150h] [rbp+48h]
  __int64 v72; // [rsp+158h] [rbp+50h]
  int v73; // [rsp+160h] [rbp+58h]
  const wchar_t *v74; // [rsp+168h] [rbp+60h]
  unsigned int *v75; // [rsp+170h] [rbp+68h]
  int v76; // [rsp+178h] [rbp+70h]
  int *v77; // [rsp+180h] [rbp+78h]
  int v78; // [rsp+188h] [rbp+80h]
  __int64 v79; // [rsp+190h] [rbp+88h]
  int v80; // [rsp+198h] [rbp+90h]
  const wchar_t *v81; // [rsp+1A0h] [rbp+98h]
  unsigned int *v82; // [rsp+1A8h] [rbp+A0h]
  int v83; // [rsp+1B0h] [rbp+A8h]
  int *v84; // [rsp+1B8h] [rbp+B0h]
  int v85; // [rsp+1C0h] [rbp+B8h]
  __int64 v86; // [rsp+1C8h] [rbp+C0h]
  int v87; // [rsp+1D0h] [rbp+C8h]
  const wchar_t *v88; // [rsp+1D8h] [rbp+D0h]
  unsigned int *v89; // [rsp+1E0h] [rbp+D8h]
  int v90; // [rsp+1E8h] [rbp+E0h]
  int *v91; // [rsp+1F0h] [rbp+E8h]
  int v92; // [rsp+1F8h] [rbp+F0h]
  __int64 v93; // [rsp+200h] [rbp+F8h]
  int v94; // [rsp+208h] [rbp+100h]
  const wchar_t *v95; // [rsp+210h] [rbp+108h]
  unsigned int *v96; // [rsp+218h] [rbp+110h]
  int v97; // [rsp+220h] [rbp+118h]
  int *v98; // [rsp+228h] [rbp+120h]
  int v99; // [rsp+230h] [rbp+128h]
  __int64 v100; // [rsp+238h] [rbp+130h]
  int v101; // [rsp+240h] [rbp+138h]
  const wchar_t *v102; // [rsp+248h] [rbp+140h]
  unsigned int *v103; // [rsp+250h] [rbp+148h]
  int v104; // [rsp+258h] [rbp+150h]
  int *v105; // [rsp+260h] [rbp+158h]
  int v106; // [rsp+268h] [rbp+160h]
  __int64 v107; // [rsp+270h] [rbp+168h]
  int v108; // [rsp+278h] [rbp+170h]
  const wchar_t *v109; // [rsp+280h] [rbp+178h]
  int *v110; // [rsp+288h] [rbp+180h]
  int v111; // [rsp+290h] [rbp+188h]
  int *v112; // [rsp+298h] [rbp+190h]
  int v113; // [rsp+2A0h] [rbp+198h]
  __int64 v114; // [rsp+2A8h] [rbp+1A0h]
  int v115; // [rsp+2B0h] [rbp+1A8h]
  const wchar_t *v116; // [rsp+2B8h] [rbp+1B0h]
  unsigned int *v117; // [rsp+2C0h] [rbp+1B8h]
  int v118; // [rsp+2C8h] [rbp+1C0h]
  int *v119; // [rsp+2D0h] [rbp+1C8h]
  int v120; // [rsp+2D8h] [rbp+1D0h]
  __int64 v121; // [rsp+2E0h] [rbp+1D8h]
  int v122; // [rsp+2E8h] [rbp+1E0h]
  const wchar_t *v123; // [rsp+2F0h] [rbp+1E8h]
  unsigned int *v124; // [rsp+2F8h] [rbp+1F0h]
  int v125; // [rsp+300h] [rbp+1F8h]
  int *v126; // [rsp+308h] [rbp+200h]
  int v127; // [rsp+310h] [rbp+208h]
  __int64 v128; // [rsp+318h] [rbp+210h]
  int v129; // [rsp+320h] [rbp+218h]
  const wchar_t *v130; // [rsp+328h] [rbp+220h]
  unsigned int *v131; // [rsp+330h] [rbp+228h]
  int v132; // [rsp+338h] [rbp+230h]
  int *v133; // [rsp+340h] [rbp+238h]
  int v134; // [rsp+348h] [rbp+240h]
  __int64 v135; // [rsp+350h] [rbp+248h]
  int v136; // [rsp+358h] [rbp+250h]
  const wchar_t *v137; // [rsp+360h] [rbp+258h]
  unsigned int *v138; // [rsp+368h] [rbp+260h]
  int v139; // [rsp+370h] [rbp+268h]
  int *v140; // [rsp+378h] [rbp+270h]
  int v141; // [rsp+380h] [rbp+278h]
  __int64 v142; // [rsp+388h] [rbp+280h]
  int v143; // [rsp+390h] [rbp+288h]
  const wchar_t *v144; // [rsp+398h] [rbp+290h]
  unsigned int *v145; // [rsp+3A0h] [rbp+298h]
  int v146; // [rsp+3A8h] [rbp+2A0h]
  int *v147; // [rsp+3B0h] [rbp+2A8h]
  int v148; // [rsp+3B8h] [rbp+2B0h]
  __int64 v149; // [rsp+3C0h] [rbp+2B8h]
  int v150; // [rsp+3C8h] [rbp+2C0h]
  const wchar_t *v151; // [rsp+3D0h] [rbp+2C8h]
  unsigned int *v152; // [rsp+3D8h] [rbp+2D0h]
  int v153; // [rsp+3E0h] [rbp+2D8h]
  int *v154; // [rsp+3E8h] [rbp+2E0h]
  int v155; // [rsp+3F0h] [rbp+2E8h]
  __int64 v156; // [rsp+3F8h] [rbp+2F0h]
  int v157; // [rsp+400h] [rbp+2F8h]
  const wchar_t *v158; // [rsp+408h] [rbp+300h]
  unsigned int *v159; // [rsp+410h] [rbp+308h]
  int v160; // [rsp+418h] [rbp+310h]
  int *v161; // [rsp+420h] [rbp+318h]
  int v162; // [rsp+428h] [rbp+320h]
  __int64 v163; // [rsp+430h] [rbp+328h]
  int v164; // [rsp+438h] [rbp+330h]
  const wchar_t *v165; // [rsp+440h] [rbp+338h]
  unsigned int *v166; // [rsp+448h] [rbp+340h]
  int v167; // [rsp+450h] [rbp+348h]
  int *v168; // [rsp+458h] [rbp+350h]
  int v169; // [rsp+460h] [rbp+358h]
  __int64 v170; // [rsp+468h] [rbp+360h]
  int v171; // [rsp+470h] [rbp+368h]
  const wchar_t *v172; // [rsp+478h] [rbp+370h]
  unsigned int *v173; // [rsp+480h] [rbp+378h]
  int v174; // [rsp+488h] [rbp+380h]
  int *v175; // [rsp+490h] [rbp+388h]
  int v176; // [rsp+498h] [rbp+390h]
  __int64 v177; // [rsp+4A0h] [rbp+398h]
  int v178; // [rsp+4A8h] [rbp+3A0h]
  const wchar_t *v179; // [rsp+4B0h] [rbp+3A8h]
  unsigned int *v180; // [rsp+4B8h] [rbp+3B0h]
  int v181; // [rsp+4C0h] [rbp+3B8h]
  int *v182; // [rsp+4C8h] [rbp+3C0h]
  int v183; // [rsp+4D0h] [rbp+3C8h]
  __int64 v184; // [rsp+4D8h] [rbp+3D0h]
  int v185; // [rsp+4E0h] [rbp+3D8h]
  const wchar_t *v186; // [rsp+4E8h] [rbp+3E0h]
  unsigned int *v187; // [rsp+4F0h] [rbp+3E8h]
  int v188; // [rsp+4F8h] [rbp+3F0h]
  int *v189; // [rsp+500h] [rbp+3F8h]
  int v190; // [rsp+508h] [rbp+400h]
  __int128 v191; // [rsp+510h] [rbp+408h]
  __int128 v192; // [rsp+520h] [rbp+418h]
  __int128 v193; // [rsp+530h] [rbp+428h]
  __int64 v194; // [rsp+540h] [rbp+438h]

  v56 = 16LL;
  v55 = 0LL;
  v57 = 0LL;
  v58 = 0LL;
  v59 = 288;
  v37 = 1;
  v24 = 1;
  v38 = 40;
  v19 = 40;
  v0 = 50;
  v39 = 10;
  v41 = 90000;
  v27 = 90000;
  v40 = 90000;
  v25 = 90000;
  v42 = 10000;
  v26 = 10000;
  v43 = 2000;
  v28 = 2000;
  v47 = 512;
  v29 = 512;
  v48 = 256;
  v30 = 256;
  v49 = 30;
  v32 = 30;
  v51 = 30;
  v34 = 30;
  v53 = 30;
  v36 = 30;
  v60 = L"GlobalCommitmentBudget";
  v61 = &v56;
  v63 = &v55;
  v67 = L"EnableTrimWnfCallback";
  v68 = &v24;
  v70 = &v37;
  v74 = L"StartPeriodicTrimThreshold";
  v20 = 10;
  v44 = 10;
  v21 = 10;
  v75 = &v19;
  v45 = 5;
  v22 = 5;
  v46 = 5;
  v23 = 5;
  v50 = 50;
  v31 = 50;
  v52 = 50;
  v33 = 50;
  v54 = 50;
  v35 = 50;
  v62 = 184549387;
  v64 = 8;
  v65 = 0LL;
  v66 = 288;
  v69 = 67108868;
  v71 = 4;
  v72 = 0LL;
  v73 = 288;
  v76 = 67108868;
  v78 = 4;
  v77 = &v38;
  v81 = L"CriticalPeriodicTrimThreshold";
  v82 = &v20;
  v84 = &v39;
  v88 = L"IdleTrimInterval";
  v89 = &v25;
  v91 = &v40;
  v95 = L"ForegroundTrimInterval";
  v96 = &v27;
  v98 = &v41;
  v102 = L"MaximumTrimInterval";
  v103 = &v26;
  v105 = &v42;
  v109 = L"MinimumTrimInterval";
  v110 = &v28;
  v112 = &v43;
  v116 = L"VideoMemoryFragmentationBuffer";
  v117 = &v21;
  v119 = &v44;
  v123 = L"SystemMemoryFragmentationBuffer";
  v124 = &v22;
  v126 = &v45;
  v130 = L"ProcessBudgetCapBuffer";
  v131 = &v23;
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
  v113 = 4;
  v114 = 0LL;
  v115 = 288;
  v118 = 67108868;
  v120 = 4;
  v121 = 0LL;
  v122 = 288;
  v125 = 67108868;
  v127 = 4;
  v128 = 0LL;
  v129 = 288;
  v132 = 67108868;
  v133 = &v46;
  v137 = L"MaxVideoMemoryFragmentationBuffer";
  v138 = &v29;
  v140 = &v47;
  v144 = L"MaxProcessBudgetCapBuffer";
  v145 = &v30;
  v147 = &v48;
  v151 = L"L_LocalMemoryBudgetDWMTarget";
  v152 = &v32;
  v154 = &v49;
  v158 = L"L_LocalMemoryBudgetFocusTarget";
  v159 = &v31;
  v161 = &v50;
  v165 = L"LNL_LocalMemoryBudgetDWMTarget";
  v166 = &v34;
  v168 = &v51;
  v172 = L"LNL_LocalMemoryBudgetFocusTarget";
  v173 = &v33;
  v175 = &v52;
  v179 = L"LNL_NonLocalMemoryBudgetDWMTarget";
  v180 = &v36;
  v182 = &v53;
  v186 = L"LNL_NonLocalMemoryBudgetFocusTarget";
  v187 = &v35;
  v189 = &v54;
  v134 = 4;
  v135 = 0LL;
  v136 = 288;
  v139 = 67108868;
  v141 = 4;
  v142 = 0LL;
  v143 = 288;
  v146 = 67108868;
  v148 = 4;
  v149 = 0LL;
  v150 = 288;
  v153 = 67108868;
  v155 = 4;
  v156 = 0LL;
  v157 = 288;
  v160 = 67108868;
  v162 = 4;
  v163 = 0LL;
  v164 = 288;
  v167 = 67108868;
  v169 = 4;
  v170 = 0LL;
  v171 = 288;
  v174 = 67108868;
  v176 = 4;
  v177 = 0LL;
  v178 = 288;
  v181 = 67108868;
  v183 = 4;
  v184 = 0LL;
  v185 = 288;
  v188 = 67108868;
  v190 = 4;
  v194 = 0LL;
  v191 = 0LL;
  v192 = 0LL;
  v193 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v58, 0LL, 0LL);
  v1 = v27;
  qword_1C0076470 = v57;
  dword_1C00764B0 = v24;
  v2 = 100;
  if ( v19 < 0x64 )
    v2 = v19;
  dword_1C00764B4 = v2;
  if ( v20 < v2 )
    v2 = v20;
  dword_1C00764B8 = v2;
  v3 = v25;
  if ( v26 > v25 )
    v3 = v26;
  if ( v3 > v27 )
    v1 = v3;
  if ( v3 >= 0x124F80 )
    v3 = 1200000;
  if ( v1 >= 0x124F80 )
    v1 = 1200000;
  v4 = 10000 * v28;
  v5 = 10000 * v26;
  dword_1C00764C8 = 10000 * v28;
  dword_1C00764BC = 10000 * v3;
  dword_1C00764C4 = 10000 * v26;
  dword_1C00764C0 = 10000 * v1;
  if ( 10000 * v26 <= 0xEA60 )
  {
    if ( v5 < 0x10 )
    {
      v5 = 16;
      dword_1C00764C4 = 16;
    }
  }
  else
  {
    v5 = 60000;
    dword_1C00764C4 = 60000;
  }
  if ( v4 < 0x10 )
  {
    dword_1C00764C8 = 16;
  }
  else if ( v4 >= v5 )
  {
    dword_1C00764C8 = v5;
  }
  v6 = 50;
  if ( v21 < 0x32 )
    v6 = v21;
  dword_1C00764CC = v6;
  v7 = 50;
  if ( v22 < 0x32 )
    v7 = v22;
  dword_1C00764D0 = v7;
  if ( v23 < 0x32 )
    v0 = v23;
  v8 = (unsigned __int64)v29 << 20;
  dword_1C00764D4 = v0;
  if ( v8 < 0x2000000 )
    v8 = 0x2000000LL;
  qword_1C00764D8 = v8;
  v9 = (unsigned __int64)v30 << 20;
  if ( v9 < 0x2000000 )
    v9 = 0x2000000LL;
  qword_1C00764E0 = v9;
  if ( v31 <= 4 )
  {
    dword_1C007658C = 5;
  }
  else
  {
    v10 = v31;
    if ( v31 > 0x5A )
      v10 = 90;
    dword_1C007658C = v10;
  }
  v11 = v32;
  v12 = v32;
  if ( 95 - v31 < v32 )
    v12 = 95 - v31;
  if ( v12 < 5 )
  {
    dword_1C0076588 = 5;
  }
  else
  {
    if ( 95 - v31 < v32 )
      v11 = 95 - v31;
    dword_1C0076588 = v11;
  }
  if ( v33 <= 4 )
  {
    dword_1C0076594 = 5;
  }
  else
  {
    v13 = v33;
    if ( v33 > 0x5A )
      v13 = 90;
    dword_1C0076594 = v13;
  }
  v14 = v34;
  v15 = v34;
  if ( 95 - v33 < v34 )
    v15 = 95 - v33;
  if ( v15 < 5 )
  {
    dword_1C0076590 = 5;
  }
  else
  {
    if ( 95 - v33 < v34 )
      v14 = 95 - v33;
    dword_1C0076590 = v14;
  }
  if ( v35 <= 4 )
  {
    dword_1C007659C = 5;
  }
  else
  {
    v16 = v35;
    if ( v35 > 0x5A )
      v16 = 90;
    dword_1C007659C = v16;
  }
  v17 = v36;
  v18 = v36;
  if ( 95 - v35 < v36 )
    v18 = 95 - v35;
  if ( v18 < 5 )
  {
    dword_1C0076598 = 5;
  }
  else
  {
    if ( 95 - v35 < v36 )
      v17 = 95 - v35;
    dword_1C0076598 = v17;
  }
}