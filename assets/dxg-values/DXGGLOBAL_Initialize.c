__int64 __fastcall DXGGLOBAL::Initialize(DXGGLOBAL *this)
{
  __int64 v1; // rdi
  __int128 v2; // xmm1
  __int128 v3; // xmm0
  __int128 v4; // xmm1
  __int128 v5; // xmm0
  NTSTATUS v6; // eax
  __int64 v7; // r14
  const wchar_t *v8; // r9
  NTSTATUS v10; // eax
  struct _ERESOURCE *v11; // rax
  NTSTATUS v12; // eax
  unsigned int v13; // ebx
  NTSTATUS v14; // eax
  NTSTATUS v15; // eax
  unsigned __int8 v16; // r9
  int v17; // ecx
  int v18; // r8d
  int v19; // eax
  int v20; // eax
  int v21; // eax
  __int128 v22; // xmm1
  __int128 v23; // xmm0
  __int128 v24; // xmm1
  __int128 v25; // xmm0
  __int128 v26; // xmm1
  __int128 v27; // xmm0
  int v28; // eax
  __int64 v29; // rcx
  int DxgkSharedObjectTypes; // eax
  unsigned int v31; // ecx
  unsigned int v32; // edx
  unsigned __int64 v33; // rbx
  __int64 v34; // rax
  __int64 v35; // rax
  __int64 v36; // rax
  __int64 v37; // rax
  DXGSESSIONMGR *v38; // rax
  DXGSESSIONMGR *v39; // rax
  int v40; // ecx
  __int64 v41; // rbx
  __int64 v42; // rax
  ULONG *v43; // rax
  __int64 v44; // rax
  _BYTE *v45; // rbx
  NTSTATUS v46; // eax
  __int64 v47; // rbx
  NTSTATUS v48; // eax
  NTSTATUS v49; // eax
  __int64 v50; // rdi
  ULONG Flags[2]; // [rsp+28h] [rbp-E0h]
  ULONG Flagsa[2]; // [rsp+28h] [rbp-E0h]
  ULONG Flagsb[2]; // [rsp+28h] [rbp-E0h]
  int OutputBuffer; // [rsp+58h] [rbp-B0h] BYREF
  unsigned int v55; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v56; // [rsp+60h] [rbp-A8h] BYREF
  unsigned int v57; // [rsp+64h] [rbp-A4h] BYREF
  unsigned int v58; // [rsp+68h] [rbp-A0h] BYREF
  unsigned int v59; // [rsp+6Ch] [rbp-9Ch] BYREF
  unsigned int v60; // [rsp+70h] [rbp-98h] BYREF
  unsigned int v61; // [rsp+74h] [rbp-94h] BYREF
  unsigned int v62; // [rsp+78h] [rbp-90h] BYREF
  int v63; // [rsp+7Ch] [rbp-8Ch] BYREF
  int v64; // [rsp+80h] [rbp-88h] BYREF
  int v65; // [rsp+84h] [rbp-84h] BYREF
  int v66; // [rsp+88h] [rbp-80h] BYREF
  int v67; // [rsp+8Ch] [rbp-7Ch] BYREF
  int v68; // [rsp+90h] [rbp-78h] BYREF
  int v69; // [rsp+94h] [rbp-74h] BYREF
  int v70; // [rsp+98h] [rbp-70h] BYREF
  int v71; // [rsp+9Ch] [rbp-6Ch] BYREF
  int v72; // [rsp+A0h] [rbp-68h] BYREF
  int v73; // [rsp+A4h] [rbp-64h] BYREF
  unsigned int v74; // [rsp+A8h] [rbp-60h] BYREF
  unsigned int v75; // [rsp+ACh] [rbp-5Ch] BYREF
  int v76; // [rsp+B0h] [rbp-58h] BYREF
  int v77; // [rsp+B4h] [rbp-54h] BYREF
  int v78; // [rsp+B8h] [rbp-50h] BYREF
  int v79; // [rsp+BCh] [rbp-4Ch] BYREF
  int v80; // [rsp+C0h] [rbp-48h] BYREF
  int v81; // [rsp+C4h] [rbp-44h] BYREF
  int v82; // [rsp+C8h] [rbp-40h] BYREF
  int v83; // [rsp+CCh] [rbp-3Ch] BYREF
  int v84; // [rsp+D0h] [rbp-38h] BYREF
  int v85; // [rsp+D4h] [rbp-34h] BYREF
  int v86; // [rsp+D8h] [rbp-30h] BYREF
  int v87; // [rsp+DCh] [rbp-2Ch] BYREF
  int v88; // [rsp+E0h] [rbp-28h] BYREF
  int v89; // [rsp+E4h] [rbp-24h] BYREF
  int v90; // [rsp+E8h] [rbp-20h] BYREF
  struct _UNICODE_STRING v91; // [rsp+F0h] [rbp-18h] BYREF
  struct _UNICODE_STRING v92; // [rsp+100h] [rbp-8h] BYREF
  _QWORD v93[13]; // [rsp+110h] [rbp+8h] BYREF
  __int64 v94; // [rsp+178h] [rbp+70h] BYREF
  int v95; // [rsp+180h] [rbp+78h]
  const wchar_t *v96; // [rsp+188h] [rbp+80h]
  unsigned int *v97; // [rsp+190h] [rbp+88h]
  int v98; // [rsp+198h] [rbp+90h]
  _QWORD *v99; // [rsp+1A0h] [rbp+98h]
  int v100; // [rsp+1A8h] [rbp+A0h]
  __int64 v101; // [rsp+1B0h] [rbp+A8h]
  int v102; // [rsp+1B8h] [rbp+B0h]
  const wchar_t *v103; // [rsp+1C0h] [rbp+B8h]
  int *v104; // [rsp+1C8h] [rbp+C0h]
  int v105; // [rsp+1D0h] [rbp+C8h]
  int *v106; // [rsp+1D8h] [rbp+D0h]
  int v107; // [rsp+1E0h] [rbp+D8h]
  __int64 v108; // [rsp+1E8h] [rbp+E0h]
  int v109; // [rsp+1F0h] [rbp+E8h]
  const wchar_t *v110; // [rsp+1F8h] [rbp+F0h]
  unsigned int *v111; // [rsp+200h] [rbp+F8h]
  int v112; // [rsp+208h] [rbp+100h]
  int *v113; // [rsp+210h] [rbp+108h]
  int v114; // [rsp+218h] [rbp+110h]
  __int64 v115; // [rsp+220h] [rbp+118h]
  int v116; // [rsp+228h] [rbp+120h]
  const wchar_t *v117; // [rsp+230h] [rbp+128h]
  unsigned int *v118; // [rsp+238h] [rbp+130h]
  int v119; // [rsp+240h] [rbp+138h]
  int *v120; // [rsp+248h] [rbp+140h]
  int v121; // [rsp+250h] [rbp+148h]
  __int64 v122; // [rsp+258h] [rbp+150h]
  int v123; // [rsp+260h] [rbp+158h]
  const wchar_t *v124; // [rsp+268h] [rbp+160h]
  int *v125; // [rsp+270h] [rbp+168h]
  int v126; // [rsp+278h] [rbp+170h]
  int *v127; // [rsp+280h] [rbp+178h]
  int v128; // [rsp+288h] [rbp+180h]
  __int64 v129; // [rsp+290h] [rbp+188h]
  int v130; // [rsp+298h] [rbp+190h]
  const wchar_t *v131; // [rsp+2A0h] [rbp+198h]
  int *v132; // [rsp+2A8h] [rbp+1A0h]
  int v133; // [rsp+2B0h] [rbp+1A8h]
  int *v134; // [rsp+2B8h] [rbp+1B0h]
  int v135; // [rsp+2C0h] [rbp+1B8h]
  __int64 v136; // [rsp+2C8h] [rbp+1C0h]
  int v137; // [rsp+2D0h] [rbp+1C8h]
  const wchar_t *v138; // [rsp+2D8h] [rbp+1D0h]
  int *v139; // [rsp+2E0h] [rbp+1D8h]
  int v140; // [rsp+2E8h] [rbp+1E0h]
  int *v141; // [rsp+2F0h] [rbp+1E8h]
  int v142; // [rsp+2F8h] [rbp+1F0h]
  __int64 v143; // [rsp+300h] [rbp+1F8h]
  int v144; // [rsp+308h] [rbp+200h]
  const wchar_t *v145; // [rsp+310h] [rbp+208h]
  int *v146; // [rsp+318h] [rbp+210h]
  int v147; // [rsp+320h] [rbp+218h]
  int *v148; // [rsp+328h] [rbp+220h]
  int v149; // [rsp+330h] [rbp+228h]
  __int64 v150; // [rsp+338h] [rbp+230h]
  int v151; // [rsp+340h] [rbp+238h]
  const wchar_t *v152; // [rsp+348h] [rbp+240h]
  int *v153; // [rsp+350h] [rbp+248h]
  int v154; // [rsp+358h] [rbp+250h]
  int *v155; // [rsp+360h] [rbp+258h]
  int v156; // [rsp+368h] [rbp+260h]
  __int64 v157; // [rsp+370h] [rbp+268h]
  int v158; // [rsp+378h] [rbp+270h]
  const wchar_t *v159; // [rsp+380h] [rbp+278h]
  int *v160; // [rsp+388h] [rbp+280h]
  int v161; // [rsp+390h] [rbp+288h]
  int *v162; // [rsp+398h] [rbp+290h]
  int v163; // [rsp+3A0h] [rbp+298h]
  __int64 v164; // [rsp+3A8h] [rbp+2A0h]
  int v165; // [rsp+3B0h] [rbp+2A8h]
  const wchar_t *v166; // [rsp+3B8h] [rbp+2B0h]
  unsigned int *v167; // [rsp+3C0h] [rbp+2B8h]
  int v168; // [rsp+3C8h] [rbp+2C0h]
  int *v169; // [rsp+3D0h] [rbp+2C8h]
  int v170; // [rsp+3D8h] [rbp+2D0h]
  __int64 v171; // [rsp+3E0h] [rbp+2D8h]
  int v172; // [rsp+3E8h] [rbp+2E0h]
  const wchar_t *v173; // [rsp+3F0h] [rbp+2E8h]
  unsigned int *v174; // [rsp+3F8h] [rbp+2F0h]
  int v175; // [rsp+400h] [rbp+2F8h]
  unsigned int *v176; // [rsp+408h] [rbp+300h]
  int v177; // [rsp+410h] [rbp+308h]
  __int64 v178; // [rsp+418h] [rbp+310h]
  int v179; // [rsp+420h] [rbp+318h]
  const wchar_t *v180; // [rsp+428h] [rbp+320h]
  unsigned int *v181; // [rsp+430h] [rbp+328h]
  int v182; // [rsp+438h] [rbp+330h]
  int *v183; // [rsp+440h] [rbp+338h]
  int v184; // [rsp+448h] [rbp+340h]
  __int64 v185; // [rsp+450h] [rbp+348h]
  int v186; // [rsp+458h] [rbp+350h]
  const wchar_t *v187; // [rsp+460h] [rbp+358h]
  unsigned int *v188; // [rsp+468h] [rbp+360h]
  int v189; // [rsp+470h] [rbp+368h]
  int *v190; // [rsp+478h] [rbp+370h]
  int v191; // [rsp+480h] [rbp+378h]
  __int64 v192; // [rsp+488h] [rbp+380h]
  int v193; // [rsp+490h] [rbp+388h]
  const wchar_t *v194; // [rsp+498h] [rbp+390h]
  unsigned int *v195; // [rsp+4A0h] [rbp+398h]
  int v196; // [rsp+4A8h] [rbp+3A0h]
  int *v197; // [rsp+4B0h] [rbp+3A8h]
  int v198; // [rsp+4B8h] [rbp+3B0h]
  __int64 v199; // [rsp+4C0h] [rbp+3B8h]
  int v200; // [rsp+4C8h] [rbp+3C0h]
  const wchar_t *v201; // [rsp+4D0h] [rbp+3C8h]
  int *v202; // [rsp+4D8h] [rbp+3D0h]
  int v203; // [rsp+4E0h] [rbp+3D8h]
  int *v204; // [rsp+4E8h] [rbp+3E0h]
  int v205; // [rsp+4F0h] [rbp+3E8h]
  __int64 v206; // [rsp+4F8h] [rbp+3F0h]
  int v207; // [rsp+500h] [rbp+3F8h]
  const wchar_t *v208; // [rsp+508h] [rbp+400h]
  int *v209; // [rsp+510h] [rbp+408h]
  int v210; // [rsp+518h] [rbp+410h]
  int *v211; // [rsp+520h] [rbp+418h]
  int v212; // [rsp+528h] [rbp+420h]
  __int64 v213; // [rsp+530h] [rbp+428h]
  int v214; // [rsp+538h] [rbp+430h]
  const wchar_t *v215; // [rsp+540h] [rbp+438h]
  unsigned int *v216; // [rsp+548h] [rbp+440h]
  int v217; // [rsp+550h] [rbp+448h]
  __int64 v218; // [rsp+558h] [rbp+450h]
  int v219; // [rsp+560h] [rbp+458h]
  __int64 v220; // [rsp+568h] [rbp+460h]
  int v221; // [rsp+570h] [rbp+468h]
  const wchar_t *v222; // [rsp+578h] [rbp+470h]
  unsigned int *v223; // [rsp+580h] [rbp+478h]
  int v224; // [rsp+588h] [rbp+480h]
  __int64 v225; // [rsp+590h] [rbp+488h]
  int v226; // [rsp+598h] [rbp+490h]
  __int64 v227; // [rsp+5A0h] [rbp+498h]
  int v228; // [rsp+5A8h] [rbp+4A0h]
  const wchar_t *v229; // [rsp+5B0h] [rbp+4A8h]
  unsigned int *v230; // [rsp+5B8h] [rbp+4B0h]
  int v231; // [rsp+5C0h] [rbp+4B8h]
  __int64 v232; // [rsp+5C8h] [rbp+4C0h]
  int v233; // [rsp+5D0h] [rbp+4C8h]
  __int64 v234; // [rsp+5D8h] [rbp+4D0h]
  int v235; // [rsp+5E0h] [rbp+4D8h]
  const wchar_t *v236; // [rsp+5E8h] [rbp+4E0h]
  unsigned int *v237; // [rsp+5F0h] [rbp+4E8h]
  int v238; // [rsp+5F8h] [rbp+4F0h]
  __int64 v239; // [rsp+600h] [rbp+4F8h]
  int v240; // [rsp+608h] [rbp+500h]
  __int64 v241; // [rsp+610h] [rbp+508h]
  int v242; // [rsp+618h] [rbp+510h]
  const wchar_t *v243; // [rsp+620h] [rbp+518h]
  unsigned int *v244; // [rsp+628h] [rbp+520h]
  int v245; // [rsp+630h] [rbp+528h]
  __int64 v246; // [rsp+638h] [rbp+530h]
  int v247; // [rsp+640h] [rbp+538h]
  __int64 v248; // [rsp+648h] [rbp+540h]
  int v249; // [rsp+650h] [rbp+548h]
  const wchar_t *v250; // [rsp+658h] [rbp+550h]
  unsigned int *v251; // [rsp+660h] [rbp+558h]
  int v252; // [rsp+668h] [rbp+560h]
  __int64 v253; // [rsp+670h] [rbp+568h]
  int v254; // [rsp+678h] [rbp+570h]
  __int64 v255; // [rsp+680h] [rbp+578h]
  int v256; // [rsp+688h] [rbp+580h]
  const wchar_t *v257; // [rsp+690h] [rbp+588h]
  int *v258; // [rsp+698h] [rbp+590h]
  int v259; // [rsp+6A0h] [rbp+598h]
  __int64 v260; // [rsp+6A8h] [rbp+5A0h]
  int v261; // [rsp+6B0h] [rbp+5A8h]
  __int64 v262; // [rsp+6B8h] [rbp+5B0h]
  int v263; // [rsp+6C0h] [rbp+5B8h]
  const wchar_t *v264; // [rsp+6C8h] [rbp+5C0h]
  int *v265; // [rsp+6D0h] [rbp+5C8h]
  int v266; // [rsp+6D8h] [rbp+5D0h]
  __int64 v267; // [rsp+6E0h] [rbp+5D8h]
  int v268; // [rsp+6E8h] [rbp+5E0h]
  __int64 v269; // [rsp+6F0h] [rbp+5E8h]
  int v270; // [rsp+6F8h] [rbp+5F0h]
  __int64 v271; // [rsp+700h] [rbp+5F8h]
  __int128 v272; // [rsp+708h] [rbp+600h]
  __int128 v273; // [rsp+718h] [rbp+610h]
  _OWORD v274[2]; // [rsp+728h] [rbp+620h] BYREF
  wchar_t v275; // [rsp+748h] [rbp+640h]
  _OWORD v276[9]; // [rsp+758h] [rbp+650h] BYREF
  int v277; // [rsp+7E8h] [rbp+6E0h]
  wchar_t v278; // [rsp+7ECh] [rbp+6E4h]

  v1 = *(_QWORD *)&DXGGLOBAL::m_pGlobal;
  memset(&v93[1], 0, 0x58uLL);
  v2 = *(_OWORD *)&v93[3];
  *(_OWORD *)(*(_QWORD *)&DXGGLOBAL::m_pGlobal + 64LL) = *(_OWORD *)&v93[1];
  v3 = *(_OWORD *)&v93[5];
  *(_OWORD *)(v1 + 80) = v2;
  v4 = *(_OWORD *)&v93[7];
  *(_OWORD *)(v1 + 96) = v3;
  v5 = *(_OWORD *)&v93[9];
  *(_OWORD *)(v1 + 112) = v4;
  *(_QWORD *)&v4 = v93[11];
  *(_OWORD *)(v1 + 128) = v5;
  *(_QWORD *)(v1 + 144) = v4;
  g_WindowsSubsystem = ZwAllocateVirtualMemory;
  qword_1C01538D0 = ZwAllocateVirtualMemoryEx;
  qword_1C01538D8 = (__int64)ZwFreeVirtualMemory;
  qword_1C01538E0 = MmMapViewOfSection;
  qword_1C01538E8 = MmUnmapViewOfSection;
  qword_1C01538F0 = (__int64)MmMapLockedPagesSpecifyCache;
  qword_1C01538F8 = (__int64)MmUnmapLockedPages;
  g_WslSubsystem = ZwAllocateVirtualMemory;
  qword_1C0153898 = ZwAllocateVirtualMemoryEx;
  qword_1C01538A0 = (__int64)ZwFreeVirtualMemory;
  qword_1C01538A8 = MmMapViewOfSection;
  qword_1C01538B0 = MmUnmapViewOfSection;
  qword_1C01538B8 = (__int64)MmMapLockedPagesSpecifyCache;
  qword_1C01538C0 = (__int64)MmUnmapLockedPages;
  v6 = ExInitializeLookasideListEx(
         (PLOOKASIDE_LIST_EX)(v1 + 305392),
         0LL,
         0LL,
         (POOL_TYPE)512,
         0,
         0x10uLL,
         0x4B677844u,
         0);
  v7 = v6;
  if ( v6 < 0 )
  {
    WdLogSingleEntry2(2LL, v1, v6);
    v8 = L"DXGGlobal 0x%I64x: Unable to initialize the lookaside list for lock order tracker, returning 0x%I64x";
    WdLogGlobalForLineNumber = 1800;
LABEL_3:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v8, v1, v7, 0LL, 0LL, 0LL);
    return (unsigned int)v7;
  }
  *(_BYTE *)(v1 + 305376) = 1;
  v10 = ExInitializeLookasideListEx(
          (PLOOKASIDE_LIST_EX)(v1 + 160),
          0LL,
          0LL,
          (POOL_TYPE)512,
          0,
          0xA0uLL,
          0x576B7844u,
          0);
  v7 = v10;
  if ( v10 < 0 )
  {
    WdLogSingleEntry2(2LL, v1, v10);
    v8 = L"DXGGlobal 0x%I64x: Unable to initialize m_VmBusPacketWorkItemList, returning 0x%I64x";
    WdLogGlobalForLineNumber = 1812;
    goto LABEL_3;
  }
  *(_BYTE *)(v1 + 1347) = 1;
  if ( !HMGRTABLE::ExpandTable((HMGRTABLE *)(v1 + 336)) )
  {
    WdLogSingleEntry1(6LL, -1073741801LL);
    WdLogGlobalForLineNumber = 1824;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed the initial shared resource handle table expansion, returning 0x%I64x",
      -1073741801LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v11 = (struct _ERESOURCE *)operator new(104LL, 1265072196LL);
  *(_QWORD *)(v1 + 600) = v11;
  if ( !v11 )
  {
    WdLogSingleEntry2(3LL, v1, -1073741801LL);
    WdLogGlobalForLineNumber = 1837;
    return 3221225495LL;
  }
  v12 = ExInitializeResourceLite(v11);
  v13 = v12;
  if ( v12 < 0 )
  {
    WdLogSingleEntry2(3LL, v1, v12);
    WdLogGlobalForLineNumber = 1847;
    return v13;
  }
  v14 = ExInitializeLookasideListEx((PLOOKASIDE_LIST_EX)(v1 + 1136), 0LL, 0LL, PagedPool, 0, 0x5F8uLL, 0x4B677844u, 0);
  v13 = v14;
  if ( v14 < 0 )
  {
    WdLogSingleEntry3(3LL, v1, v14, 0LL, *(_QWORD *)Flags);
    WdLogGlobalForLineNumber = 1856;
    return v13;
  }
  *(_BYTE *)(v1 + 1345) = 1;
  v15 = ExInitializeLookasideListEx((PLOOKASIDE_LIST_EX)(v1 + 1232), 0LL, 0LL, PagedPool, 0, 0x5E0uLL, 0x4B677844u, 0);
  v13 = v15;
  if ( v15 < 0 )
  {
    WdLogSingleEntry3(3LL, v1, v15, 0LL, *(_QWORD *)Flagsa);
    WdLogGlobalForLineNumber = 1866;
    return v13;
  }
  v16 = g_bSkuSupportMultipleUsers;
  *(_BYTE *)(v1 + 1346) = 1;
  v79 = 32;
  v93[0] = 0x4000000LL;
  v62 = 0;
  v77 = 0;
  v63 = 0;
  v78 = 1;
  v61 = 0;
  v60 = 0;
  v65 = 0;
  v80 = 0;
  v81 = 0;
  v66 = 0;
  v67 = 0;
  v82 = 0;
  v83 = 0;
  v68 = 0;
  v84 = 0;
  v69 = 0;
  v85 = 0;
  v64 = 0;
  v72 = 0;
  if ( v16 )
    v17 = g_IsInternalReleaseOrDbg != 0 ? 0x100000 : 0x80000;
  else
    v17 = g_IsInternalReleaseOrDbg != 0 ? 0x80000 : 0x10000;
  v86 = v17;
  if ( v16 )
    v18 = g_IsInternalReleaseOrDbg != 0 ? 8 : 4;
  else
    v18 = 2;
  v75 = v18;
  if ( v16 )
    v19 = g_IsInternalReleaseOrDbg != 0 ? 0x80000 : 0x10000;
  else
    v19 = g_IsInternalReleaseOrDbg != 0 ? 0x80000 : 0x4000;
  v87 = v19;
  v56 = v19;
  v88 = 300;
  v58 = 300;
  v55 = v17;
  v76 = 1;
  v57 = v18;
  v59 = 1;
  v89 = 5000;
  v70 = 0;
  v90 = 15000;
  v71 = 0;
  v73 = *(_DWORD *)(v1 + 305880);
  v96 = L"TerminationListSizeLimit";
  v97 = &v62;
  v99 = v93;
  v103 = L"ValidateWDDMCaps";
  v104 = &v63;
  v106 = &v77;
  v110 = L"WDDM2LockManagement";
  v111 = &v61;
  v113 = &v78;
  v117 = L"MaximumAdapterCount";
  v118 = &v60;
  v120 = &v79;
  v124 = L"InvestigationDebugParameter";
  v125 = &v65;
  v127 = &v80;
  v131 = L"EnableIgnoreWin32ProcessStatus";
  v132 = &v66;
  v134 = &v81;
  v138 = L"EnableHMDTestMode";
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
  v112 = 67108868;
  v114 = 4;
  v115 = 0LL;
  v116 = 288;
  v119 = 67108868;
  v121 = 4;
  v122 = 0LL;
  v123 = 288;
  v126 = 67108868;
  v128 = 4;
  v129 = 0LL;
  v130 = 288;
  v133 = 67108868;
  v135 = 4;
  v136 = 0LL;
  v137 = 288;
  v140 = 67108868;
  v139 = &v67;
  v141 = &v82;
  v145 = L"PreserveFirmwareMode";
  v146 = &v68;
  v148 = &v83;
  v152 = L"PreventFullscreenWireFormatChange";
  v153 = &v69;
  v155 = &v84;
  v159 = L"EnableFuzzing";
  v160 = &v64;
  v162 = &v85;
  v166 = L"InternalDiagnosticsBufferSize";
  v167 = &v55;
  v169 = &v86;
  v173 = L"InternalDiagnosticsBufferMultiplier";
  v174 = &v57;
  v176 = &v75;
  v180 = L"ExternalDiagnosticsBufferSize";
  v181 = &v56;
  v183 = &v87;
  v187 = L"ExternalDiagnosticsBufferMultiplier";
  v188 = &v59;
  v190 = &v76;
  v194 = L"DiagnosticsBufferExpansionTime";
  v142 = 4;
  v143 = 0LL;
  v144 = 288;
  v147 = 67108868;
  v149 = 4;
  v150 = 0LL;
  v151 = 288;
  v154 = 67108868;
  v156 = 4;
  v157 = 0LL;
  v158 = 288;
  v161 = 67108868;
  v163 = 4;
  v164 = 0LL;
  v165 = 288;
  v168 = 67108868;
  v170 = 4;
  v171 = 0LL;
  v172 = 288;
  v175 = 67108868;
  v177 = 4;
  v178 = 0LL;
  v179 = 288;
  v182 = 67108868;
  v184 = 4;
  v185 = 0LL;
  v186 = 288;
  v189 = 67108868;
  v191 = 4;
  v192 = 0LL;
  v193 = 288;
  v195 = &v58;
  v197 = &v88;
  v201 = L"RapidHpdTimeoutInMilliseconds";
  v202 = &v70;
  v204 = &v89;
  v208 = L"RapidHpdMaxChainInMilliseconds";
  v209 = &v71;
  v211 = &v90;
  v215 = L"ForceUsb4MonitorSupport";
  v216 = &g_bDbgForceUsb4MonitorSupport;
  v222 = L"Usb4MonitorTargetId";
  v223 = &g_DbgUsb4MonitorTargetId;
  v229 = L"Usb4MonitorDpcdUSB4_Driver_ID";
  v230 = &g_DbgUsb4MonitorDpcdUSB4_Driver_ID;
  v236 = L"Usb4MonitorDpcdDP_IN_Adapter_Number";
  v237 = &g_DbgUsb4MonitorDpcdDP_IN_Adapter_Number;
  v243 = L"Usb4MonitorPowerOnDelayInSeconds";
  v244 = &g_DbgUsb4MonitorPowerOnDelayInSeconds;
  v250 = L"TreatUsb4MonitorAsNormal";
  v251 = &g_bDbgTreatUsb4MonitorAsNormal;
  v196 = 67108868;
  v198 = 4;
  v199 = 0LL;
  v200 = 288;
  v203 = 67108868;
  v205 = 4;
  v206 = 0LL;
  v207 = 288;
  v210 = 67108868;
  v212 = 4;
  v213 = 0LL;
  v214 = 288;
  v217 = 67108868;
  v218 = 0LL;
  v219 = 0;
  v220 = 0LL;
  v221 = 288;
  v224 = 67108868;
  v225 = 0LL;
  v226 = 0;
  v227 = 0LL;
  v228 = 288;
  v231 = 67108868;
  v232 = 0LL;
  v233 = 0;
  v234 = 0LL;
  v235 = 288;
  v238 = 67108868;
  v239 = 0LL;
  v240 = 0;
  v241 = 0LL;
  v242 = 288;
  v245 = 67108868;
  v246 = 0LL;
  v247 = 0;
  v248 = 0LL;
  v249 = 288;
  v252 = 67108868;
  v253 = 0LL;
  v254 = 0;
  v255 = 0LL;
  v256 = 288;
  v259 = 67108868;
  v263 = 288;
  v257 = L"AllowAdvancedEtwLogging";
  v266 = 67108868;
  v258 = &v72;
  v260 = 0LL;
  v261 = 0;
  v264 = L"NodeUsageTelemetryTimerInterval";
  v265 = &v73;
  v262 = 0LL;
  v267 = 0LL;
  v268 = 0;
  v269 = 0LL;
  v270 = 0;
  v271 = 0LL;
  v272 = 0LL;
  Flagsb[1] = 0;
  v273 = 0LL;
  if ( (int)RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v94) < 0 )
  {
    *(_QWORD *)(v1 + 880) = 0x4000000LL;
    *(_DWORD *)(v1 + 1364) = 32;
    *(_BYTE *)(v1 + 888) = 0;
    *(_DWORD *)(v1 + 1360) = 1;
    *(_DWORD *)(v1 + 1648) = 0;
    *(_DWORD *)(v1 + 1664) = 0;
  }
  else
  {
    *(_QWORD *)(v1 + 880) = v62;
    *(_BYTE *)(v1 + 888) = v63 != 0;
    *(_BYTE *)(v1 + 304834) = v64 != 0;
    v20 = 1;
    if ( v61 < 2 )
      v20 = v61;
    *(_DWORD *)(v1 + 1360) = v20;
    v21 = v60;
    if ( v60 >= 4 )
    {
      if ( v60 > 0x400 )
      {
        v21 = 1024;
        v60 = 1024;
      }
    }
    else
    {
      v21 = 4;
      v60 = 4;
    }
    *(_DWORD *)(v1 + 1364) = v21;
    *(_DWORD *)(v1 + 1648) = v65;
    *(_DWORD *)(v1 + 1664) = v66;
    *(_BYTE *)(v1 + 304833) = v67 == 1;
    *(_BYTE *)(v1 + 304888) = v68 != 0;
    *(_BYTE *)(v1 + 304889) = v69 != 0;
    if ( v70 )
      *(_DWORD *)(v1 + 305600) = v70;
    if ( v71 )
      *(_DWORD *)(v1 + 305604) = v71;
    if ( !g_IsInternalRelease && !g_OSTestSigningEnabled )
    {
      g_bDbgForceUsb4MonitorSupport = 0;
      g_bDbgTreatUsb4MonitorAsNormal = 0;
      g_DbgUsb4MonitorPowerOnDelayInSeconds = 0;
    }
    *(_BYTE *)(v1 + 305672) = v72 != 0;
    *(_DWORD *)(v1 + 305880) = v73;
    DXGGLOBAL::SetNodeUsageTelemetryTimer((DXGGLOBAL *)v1);
  }
  *(_DWORD *)(v1 + 872) = 0;
  v22 = *(_OWORD *)L"Y\\MACHINE\\System\\ControlSet001\\Control\\Terminal Server\\WinStations";
  v74 = 0;
  v276[0] = *(_OWORD *)L"\\REGISTRY\\MACHINE\\System\\ControlSet001\\Control\\Terminal Server\\WinStations";
  *(_QWORD *)&v92.Length = 9830548LL;
  v23 = *(_OWORD *)L"E\\System\\ControlSet001\\Control\\Terminal Server\\WinStations";
  *(_QWORD *)&v91.Length = 2228256LL;
  v276[1] = v22;
  v24 = *(_OWORD *)L"\\ControlSet001\\Control\\Terminal Server\\WinStations";
  v276[2] = v23;
  v25 = *(_OWORD *)L"Set001\\Control\\Terminal Server\\WinStations";
  v276[3] = v24;
  v26 = *(_OWORD *)L"ontrol\\Terminal Server\\WinStations";
  v276[4] = v25;
  v27 = *(_OWORD *)L"erminal Server\\WinStations";
  v276[5] = v26;
  v276[6] = v27;
  v276[7] = *(_OWORD *)L"Server\\WinStations";
  v28 = *(_DWORD *)L"ns";
  v276[8] = *(_OWORD *)L"inStations";
  v277 = v28;
  v278 = aRegistryMachin_13[74];
  v92.Buffer = (wchar_t *)v276;
  v275 = aDwmframeinterv[16];
  v91.Buffer = (wchar_t *)v274;
  v274[0] = *(_OWORD *)L"DWMFRAMEINTERVAL";
  v274[1] = *(_OWORD *)L"INTERVAL";
  if ( ReadRegistryDwordKeyValue(&v92, &v91, &v74) >= 0 && v74 )
    *(_DWORD *)(v1 + 305152) = v74;
  DxgkSharedObjectTypes = CreateDxgkSharedObjectTypes(v29);
  v13 = DxgkSharedObjectTypes;
  if ( DxgkSharedObjectTypes < 0 )
  {
    WdLogSingleEntry1(3LL, DxgkSharedObjectTypes);
    WdLogGlobalForLineNumber = 2143;
    return v13;
  }
  v31 = v57;
  if ( !v57 || ((v57 - 1) & v57) != 0 )
  {
    v31 = v75;
    v57 = v75;
  }
  v32 = v59;
  if ( !v59 || ((v59 - 1) & v59) != 0 )
  {
    v32 = v76;
    v59 = v76;
  }
  if ( !g_OSTestSigningEnabled )
  {
    if ( v55 < 0x1000 || v55 * v31 > 0x1000000 )
    {
      v55 = 0x1000000;
      v57 = 1;
    }
    if ( v56 < 0x1000 || v56 * v32 > 0x1000000 )
    {
      v56 = 0x1000000;
      v59 = 1;
    }
  }
  if ( v58 > 0xE10 )
    v58 = 3600;
  v33 = (-(__int64)(g_IsInternalReleaseOrDbg != 0) & 0xFFFFFFFFFFFFFF40uLL) + 256;
  v34 = operator new(112LL, 1265072196LL);
  if ( v34 )
    v35 = DXGDIAGNOSTICS::DXGDIAGNOSTICS(v34, v55, v57, v33, v58);
  else
    v35 = 0LL;
  *(_QWORD *)(v1 + 928) = v35;
  v36 = operator new(112LL, 1265072196LL);
  if ( v36 )
  {
    Flagsb[0] = v58;
    v37 = DXGDIAGNOSTICS::DXGDIAGNOSTICS(v36, v56, v59, v33, *(_QWORD *)Flagsb);
  }
  else
  {
    v37 = 0LL;
  }
  *(_QWORD *)(v1 + 936) = v37;
  if ( !*(_QWORD *)(v1 + 928) )
  {
    WdLogSingleEntry1(6LL, v55);
    WdLogGlobalForLineNumber = 2197;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for internal diagnostics buffers (SmallInternalDiagnosticsSize = 0x%I64x).",
      v55,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  if ( !v37 )
  {
    WdLogSingleEntry1(6LL, v56);
    WdLogGlobalForLineNumber = 2203;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for external diagnostics buffers (SmallExternalDiagnosticsSize = 0x%I64x).",
      v56,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v38 = (DXGSESSIONMGR *)operator new(448LL, 1265072196LL);
  if ( v38 )
    v39 = DXGSESSIONMGR::DXGSESSIONMGR(v38);
  else
    v39 = 0LL;
  *(_QWORD *)(v1 + 944) = v39;
  if ( !v39 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2210;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for dxgkrnl session manager.",
      2210LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v40 = *(_DWORD *)(v1 + 1364);
  v41 = (unsigned int)(v40 + 31) >> 5;
  v42 = 4LL * ((unsigned int)v41 + ((unsigned int)(1055 - v40) >> 5));
  if ( !is_mul_ok((unsigned int)v41 + ((unsigned int)(1055 - v40) >> 5), 4uLL) )
    v42 = -1LL;
  v43 = (ULONG *)operator new[](v42, 1265072196LL, 256LL);
  *(_QWORD *)(v1 + 864) = v43;
  if ( !v43 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2219;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for dxgkrnl adapter ordinal bits.",
      2219LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  RtlInitializeBitMap((PRTL_BITMAP)(v1 + 832), v43, *(_DWORD *)(v1 + 1364));
  RtlInitializeBitMap((PRTL_BITMAP)(v1 + 848), (PULONG)(*(_QWORD *)(v1 + 864) + 4 * v41), 1024 - *(_DWORD *)(v1 + 1364));
  if ( DXGPROCESS::CreateDxgProcess((struct DXGPROCESS **)(v1 + 1368), 0LL, 0LL, 0, 0LL) < 0 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2233;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for system process.",
      2233LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  if ( PsInitialSystemProcess != *(PEPROCESS *)(*(_QWORD *)(v1 + 1368) + 56LL) )
  {
    WdLogSingleEntry0(1LL);
    WdLogGlobalForLineNumber = 2236;
    DxgkLogInternalTriageEvent(
      0,
      262146,
      -1,
      (unsigned int)L"PsInitialSystemProcess == m_pSystemDxgProcess->GetEProcess()",
      2236LL,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  v44 = operator new(640LL, 1265072196LL);
  v45 = (_BYTE *)v44;
  if ( v44 )
  {
    *(_BYTE *)v44 = 1;
    *(_QWORD *)(v44 + 16) = 0LL;
    *(_QWORD *)(v44 + 24) = 0LL;
    *(_QWORD *)(v44 + 32) = 0LL;
    *(_DWORD *)(v44 + 40) = 0;
    *(_DWORD *)(v44 + 44) = 69;
    *(_DWORD *)(v44 + 48) = 1;
    *(_DWORD *)(v44 + 632) = 0;
    memset((void *)(v44 + 56), 0, 0x240uLL);
    *v45 = 0;
  }
  else
  {
    v45 = 0LL;
  }
  *(_QWORD *)(v1 + 1464) = v45;
  if ( !v45 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2241;
    DxgkLogInternalTriageEvent(0, 262145, -1, (unsigned int)L"Failed to Qdc cache.", 2241LL, 0LL, 0LL, 0LL, 0LL);
    return 3221225495LL;
  }
  KeInitializeSpinLock(&qword_1C01537D0);
  DXGVALIDATION::InitializeBootSettings((DXGVALIDATION *)(v1 + 1652));
  DXGGLOBAL::CsExitInitiatedWnfSubscription((DXGGLOBAL *)v1);
  KeInitializeTimer((PKTIMER)(v1 + 1904));
  KeInitializeDpc((PRKDPC)(v1 + 1968), CsExitInitiatedReleaseComponentReferences, (PVOID)v1);
  LOBYTE(OutputBuffer) = 0;
  v46 = ZwPowerInformation((POWER_INFORMATION_LEVEL)66, 0LL, 0, &OutputBuffer, 1u);
  if ( v46 >= 0 )
  {
    if ( (_BYTE)OutputBuffer )
      DXGGLOBAL::SubscribeWNFForCSAccounting((DXGGLOBAL *)v1);
  }
  else
  {
    v47 = v46;
    WdLogSingleEntry1(2LL, v46);
    WdLogGlobalForLineNumber = 2278;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Failed to get the platformInformation. Status : 0x%I64x",
      v47,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  *(_QWORD *)(v1 + 2056) = v1;
  *(_QWORD *)(v1 + 2048) = CsExitInitiatedReleaseComponentReferencesPassiveLevel;
  *(_QWORD *)(v1 + 2032) = 0LL;
  DXGGLOBAL::InitializeResourceManagerSid((DXGGLOBAL *)v1);
  *(_DWORD *)(v1 + 304820) &= ~1u;
  *(_DWORD *)(v1 + 304808) = 10;
  *(_DWORD *)(v1 + 304812) = 50;
  *(_DWORD *)(v1 + 304816) = 30;
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 1752));
  DisplayDiagnostics::Initialize((PVOID)(v1 + 304960));
  v48 = PoRegisterPowerSettingCallback(
          0LL,
          &GUID_ADVANCED_COLOR_QUALITY_BIAS,
          DXGGLOBAL::AdvancedColorPowerSettingsCallback,
          (PVOID)v1,
          0LL);
  v7 = v48;
  if ( v48 < 0 )
  {
    WdLogSingleEntry1(2LL, v48);
    WdLogGlobalForLineNumber = 2315;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"PoRegisterPowerSettingCallback for GUID_HDR_DISPLAY_QUALITY_BIAS failed with status:0x%I64x",
      v7,
      0LL,
      0LL,
      0LL,
      0LL);
    return (unsigned int)v7;
  }
  v49 = PoRegisterPowerSettingCallback(0LL, &GUID_ACDC_POWER_SOURCE, DXGGLOBAL::AcDcPowerSourceCallback, (PVOID)v1, 0LL);
  v50 = v49;
  if ( v49 < 0 )
  {
    WdLogSingleEntry1(2LL, v49);
    WdLogGlobalForLineNumber = 2325;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"PoRegisterPowerSettingCallback for GUID_ACDC_POWER_SOURCE failed with status:0x%I64x",
      v50,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  return (unsigned int)v50;
}
