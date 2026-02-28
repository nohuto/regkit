bool __fastcall VidSchiReadGlobalConfiguration(__int64 a1)
{
  unsigned int v2; // ecx
  __int64 v3; // rdx
  int v4; // edi
  bool v5; // sf
  bool v6; // of
  int NodeConfiguration; // eax
  unsigned int v8; // ecx
  __int64 v9; // r11
  unsigned int v10; // r9d
  int *v11; // rdx
  __int64 v12; // r8
  __int64 v13; // rdx
  unsigned int v14; // r9d
  _DWORD *v15; // rax
  int *v16; // rax
  int v17; // r10d
  int *v18; // rax
  int v19; // ecx
  int v20; // ecx
  bool v21; // zf
  __int64 v22; // rcx
  __int64 v23; // rcx
  __int64 v24; // rcx
  int v25; // ecx
  int v26; // eax
  bool IsEnabled; // al
  int v28; // ecx
  _QWORD *v29; // rdx
  unsigned int v30; // eax
  __int64 v31; // rcx
  int v32; // eax
  bool v33; // cc
  int v34; // eax
  __int64 v35; // rax
  int v36; // eax
  int v37; // edx
  unsigned int v38; // edx
  int v39; // ecx
  unsigned int v40; // edx
  unsigned int v41; // ecx
  __int64 v42; // rdx
  int v43; // eax
  bool result; // al
  _DWORD *v45; // rax
  int v46; // ecx
  int v47; // r8d
  int v48; // eax
  unsigned int v49; // [rsp+58h] [rbp-B0h] BYREF
  unsigned int v50; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v51; // [rsp+60h] [rbp-A8h] BYREF
  unsigned int v52; // [rsp+64h] [rbp-A4h] BYREF
  unsigned int v53; // [rsp+68h] [rbp-A0h] BYREF
  unsigned int v54; // [rsp+6Ch] [rbp-9Ch] BYREF
  int v55; // [rsp+70h] [rbp-98h] BYREF
  int v56; // [rsp+74h] [rbp-94h] BYREF
  int v57; // [rsp+78h] [rbp-90h] BYREF
  int v58; // [rsp+7Ch] [rbp-8Ch] BYREF
  int v59; // [rsp+80h] [rbp-88h] BYREF
  int v60; // [rsp+84h] [rbp-84h] BYREF
  int v61; // [rsp+88h] [rbp-80h] BYREF
  int v62; // [rsp+8Ch] [rbp-7Ch] BYREF
  int v63; // [rsp+90h] [rbp-78h] BYREF
  int v64; // [rsp+94h] [rbp-74h] BYREF
  int v65; // [rsp+98h] [rbp-70h] BYREF
  int v66; // [rsp+9Ch] [rbp-6Ch] BYREF
  int v67; // [rsp+A0h] [rbp-68h] BYREF
  int v68; // [rsp+A4h] [rbp-64h] BYREF
  int v69; // [rsp+A8h] [rbp-60h] BYREF
  int v70; // [rsp+ACh] [rbp-5Ch] BYREF
  int v71; // [rsp+B0h] [rbp-58h] BYREF
  int v72; // [rsp+B4h] [rbp-54h] BYREF
  int v73; // [rsp+B8h] [rbp-50h] BYREF
  int v74; // [rsp+BCh] [rbp-4Ch] BYREF
  int v75; // [rsp+C0h] [rbp-48h] BYREF
  int v76; // [rsp+C4h] [rbp-44h] BYREF
  int v77; // [rsp+C8h] [rbp-40h] BYREF
  int v78; // [rsp+CCh] [rbp-3Ch] BYREF
  int v79; // [rsp+D0h] [rbp-38h] BYREF
  int v80; // [rsp+D4h] [rbp-34h] BYREF
  int v81; // [rsp+D8h] [rbp-30h] BYREF
  int v82; // [rsp+DCh] [rbp-2Ch] BYREF
  BOOL v83; // [rsp+E0h] [rbp-28h] BYREF
  int v84; // [rsp+E4h] [rbp-24h] BYREF
  int v85; // [rsp+E8h] [rbp-20h] BYREF
  int v86; // [rsp+ECh] [rbp-1Ch] BYREF
  int v87; // [rsp+F0h] [rbp-18h] BYREF
  int v88; // [rsp+F4h] [rbp-14h] BYREF
  int v89; // [rsp+F8h] [rbp-10h] BYREF
  int v90; // [rsp+FCh] [rbp-Ch] BYREF
  int v91; // [rsp+100h] [rbp-8h] BYREF
  int v92; // [rsp+104h] [rbp-4h] BYREF
  int v93; // [rsp+108h] [rbp+0h] BYREF
  int v94; // [rsp+10Ch] [rbp+4h] BYREF
  int v95; // [rsp+110h] [rbp+8h] BYREF
  int v96; // [rsp+114h] [rbp+Ch] BYREF
  int v97; // [rsp+118h] [rbp+10h] BYREF
  int v98; // [rsp+11Ch] [rbp+14h] BYREF
  int v99; // [rsp+120h] [rbp+18h] BYREF
  int v100; // [rsp+124h] [rbp+1Ch] BYREF
  int v101; // [rsp+128h] [rbp+20h] BYREF
  int v102; // [rsp+12Ch] [rbp+24h] BYREF
  int v103; // [rsp+130h] [rbp+28h] BYREF
  int v104; // [rsp+134h] [rbp+2Ch] BYREF
  int v105; // [rsp+138h] [rbp+30h] BYREF
  int v106; // [rsp+13Ch] [rbp+34h] BYREF
  int v107; // [rsp+140h] [rbp+38h] BYREF
  unsigned int v108; // [rsp+144h] [rbp+3Ch] BYREF
  int v109; // [rsp+148h] [rbp+40h] BYREF
  int v110; // [rsp+14Ch] [rbp+44h] BYREF
  int v111; // [rsp+150h] [rbp+48h] BYREF
  BOOL v112; // [rsp+154h] [rbp+4Ch] BYREF
  _QWORD v113[232]; // [rsp+158h] [rbp+50h] BYREF

  v85 = 25000;
  v84 = 0;
  v86 = 50000;
  v2 = *(_DWORD *)(a1 + 228);
  v90 = 0;
  v87 = 1;
  v3 = *(_QWORD *)(a1 + 16);
  v4 = 16;
  v88 = 2;
  v106 = 16;
  v89 = 3;
  v91 = 0;
  v92 = 1;
  v93 = 1;
  v94 = 0;
  v97 = 0;
  v95 = 20;
  v96 = 2;
  v56 = 7;
  v99 = 0;
  v100 = 900;
  v101 = 1000;
  v98 = 1;
  v102 = 8;
  v103 = 0;
  v68 = 10;
  v104 = 1;
  v105 = 0;
  v109 = 0;
  v110 = 0;
  v111 = 0;
  v107 = 100;
  v75 = 64;
  v108 = v2;
  v6 = __OFSUB__(*(_DWORD *)(v3 + 2820), 2000);
  v5 = *(_DWORD *)(v3 + 2820) - 2000 < 0;
  v58 = 0;
  v51 = 25000;
  v112 = v5 == v6;
  v83 = v112;
  v52 = 50000;
  v60 = 1;
  v57 = 2;
  v53 = 3;
  v81 = 0;
  v59 = 0;
  v61 = 1;
  v77 = 1;
  v78 = 0;
  v79 = 0;
  v54 = 20;
  v80 = 2;
  v55 = 7;
  v62 = 0;
  v70 = 900;
  v71 = 1000;
  v76 = 1;
  v73 = 8;
  v69 = 0;
  v67 = 10;
  v63 = 1;
  v64 = 0;
  v65 = 0;
  v66 = 0;
  v72 = 16;
  v74 = 100;
  v49 = 64;
  v82 = 0;
  v50 = v2;
  if ( *(int *)(v3 + 2820) >= 1300 && *(_BYTE *)(v3 + 2757) )
  {
    v56 = 1;
    v55 = 1;
  }
  memset(v113, 0, 0x738uLL);
  v113[7] = 0LL;
  LODWORD(v113[1]) = 288;
  LODWORD(v113[4]) = 67108868;
  LODWORD(v113[6]) = 4;
  v113[2] = L"AutoSyncToCPUPriority";
  v113[3] = &v58;
  v113[5] = &v84;
  v113[9] = L"QuantumUnit";
  v113[10] = &v51;
  v113[12] = &v85;
  v113[16] = L"PreemptionQuantumUnit";
  v113[17] = &v52;
  v113[19] = &v86;
  v113[23] = L"EnablePreemption";
  v113[24] = &v60;
  v113[26] = &v87;
  v113[30] = L"HwQueuedRenderPacketGroupLimit";
  v113[31] = &v57;
  v113[33] = &v88;
  v113[37] = L"QueuedPresentLimit";
  v113[38] = &v53;
  v113[40] = &v89;
  v113[44] = L"InitDriverFenceId";
  v113[45] = &v81;
  v113[47] = &v90;
  v113[51] = L"CarryOverUsedQuantum";
  LODWORD(v113[8]) = 288;
  LODWORD(v113[11]) = 67108868;
  LODWORD(v113[13]) = 4;
  v113[14] = 0LL;
  LODWORD(v113[15]) = 288;
  LODWORD(v113[18]) = 67108868;
  LODWORD(v113[20]) = 4;
  v113[21] = 0LL;
  LODWORD(v113[22]) = 288;
  LODWORD(v113[25]) = 67108868;
  LODWORD(v113[27]) = 4;
  v113[28] = 0LL;
  LODWORD(v113[29]) = 288;
  LODWORD(v113[32]) = 67108868;
  LODWORD(v113[34]) = 4;
  v113[35] = 0LL;
  LODWORD(v113[36]) = 288;
  LODWORD(v113[39]) = 67108868;
  LODWORD(v113[41]) = 4;
  v113[42] = 0LL;
  LODWORD(v113[43]) = 288;
  LODWORD(v113[46]) = 67108868;
  LODWORD(v113[48]) = 4;
  v113[49] = 0LL;
  LODWORD(v113[50]) = 288;
  v113[52] = &v59;
  v113[54] = &v91;
  v113[58] = L"EnableFlipImmediateSwFlipQueue";
  v113[59] = &v61;
  v113[61] = &v92;
  v113[65] = L"AdjustWorkerThreadPriority";
  v113[66] = &v77;
  v113[68] = &v93;
  v113[72] = L"CountFlipTowardHwLimit";
  v113[73] = &v78;
  v113[75] = &v94;
  v113[79] = L"NumberOfDmaPacketPool";
  v113[80] = &v54;
  v113[82] = &v95;
  v113[86] = L"ProfileLevel";
  v113[87] = &v80;
  v113[89] = &v96;
  v113[93] = L"VSyncIdleTimeout";
  v113[94] = &v55;
  v113[96] = &v56;
  v113[100] = L"CountPresentTowardHwLimit";
  v113[101] = &v79;
  v113[103] = &v97;
  v113[107] = L"EnableContextDelay";
  v113[108] = &v76;
  LODWORD(v113[53]) = 67108868;
  LODWORD(v113[55]) = 4;
  v113[56] = 0LL;
  LODWORD(v113[57]) = 288;
  LODWORD(v113[60]) = 67108868;
  LODWORD(v113[62]) = 4;
  v113[63] = 0LL;
  LODWORD(v113[64]) = 288;
  LODWORD(v113[67]) = 67108868;
  LODWORD(v113[69]) = 4;
  v113[70] = 0LL;
  LODWORD(v113[71]) = 288;
  LODWORD(v113[74]) = 67108868;
  LODWORD(v113[76]) = 4;
  v113[77] = 0LL;
  LODWORD(v113[78]) = 288;
  LODWORD(v113[81]) = 67108868;
  LODWORD(v113[83]) = 4;
  v113[84] = 0LL;
  LODWORD(v113[85]) = 288;
  LODWORD(v113[88]) = 67108868;
  LODWORD(v113[90]) = 4;
  v113[91] = 0LL;
  LODWORD(v113[92]) = 288;
  LODWORD(v113[95]) = 67108868;
  LODWORD(v113[97]) = 4;
  v113[98] = 0LL;
  LODWORD(v113[99]) = 288;
  LODWORD(v113[102]) = 67108868;
  LODWORD(v113[104]) = 4;
  v113[105] = 0LL;
  LODWORD(v113[106]) = 288;
  LODWORD(v113[109]) = 67108868;
  v113[110] = &v98;
  v113[114] = L"LogDriverVSyncCallback";
  v113[115] = &v62;
  v113[117] = &v99;
  v113[121] = L"MaximumAllowedPreemptionDelay";
  v113[122] = &v70;
  v113[124] = &v100;
  v113[128] = L"ContextSchedulingPenaltyDelay";
  v113[129] = &v71;
  v113[131] = &v101;
  v113[135] = L"BackgroundProcessMaximumAllowedPreemptionDelay";
  v113[136] = &v73;
  v113[138] = &v102;
  v113[142] = L"ForceEnableFlipFenceModel";
  v113[143] = &v69;
  v113[145] = &v103;
  v113[149] = L"YieldPercentage";
  v113[150] = &v67;
  v113[152] = &v68;
  v113[156] = L"ForegroundPriorityBoost";
  v113[157] = &v63;
  v113[159] = &v104;
  v113[163] = L"ForceFlipTrueImmediateMode";
  v113[164] = &v64;
  LODWORD(v113[111]) = 4;
  v113[112] = 0LL;
  LODWORD(v113[113]) = 288;
  LODWORD(v113[116]) = 67108868;
  LODWORD(v113[118]) = 4;
  v113[119] = 0LL;
  LODWORD(v113[120]) = 288;
  LODWORD(v113[123]) = 67108868;
  LODWORD(v113[125]) = 4;
  v113[126] = 0LL;
  LODWORD(v113[127]) = 288;
  LODWORD(v113[130]) = 67108868;
  LODWORD(v113[132]) = 4;
  v113[133] = 0LL;
  LODWORD(v113[134]) = 288;
  LODWORD(v113[137]) = 67108868;
  LODWORD(v113[139]) = 4;
  v113[140] = 0LL;
  LODWORD(v113[141]) = 288;
  LODWORD(v113[144]) = 67108868;
  LODWORD(v113[146]) = 4;
  v113[147] = 0LL;
  LODWORD(v113[148]) = 288;
  LODWORD(v113[151]) = 67108868;
  LODWORD(v113[153]) = 4;
  v113[154] = 0LL;
  LODWORD(v113[155]) = 288;
  LODWORD(v113[158]) = 67108868;
  LODWORD(v113[160]) = 4;
  v113[161] = 0LL;
  LODWORD(v113[162]) = 288;
  LODWORD(v113[165]) = 67108868;
  LODWORD(v113[167]) = 4;
  v113[166] = &v105;
  v113[170] = L"MaxYieldInterval";
  v113[171] = &v72;
  v113[173] = &v106;
  v113[177] = L"MaxFocusGpuQuantumWithoutPresent";
  v113[178] = &v74;
  v113[180] = &v107;
  v113[184] = L"HistoryLogSize";
  v113[185] = &v49;
  v113[187] = &v75;
  v113[191] = L"HwQueuePacketCap";
  v113[192] = &v50;
  v113[194] = &v108;
  v113[198] = L"FlipDoNotFlipMode";
  v113[199] = &v65;
  v113[201] = &v109;
  v113[205] = L"DdiSuspendMode";
  v113[206] = &v66;
  v113[208] = &v110;
  v113[212] = L"PfnCpuOverride";
  v113[213] = &v82;
  v113[215] = &v111;
  v113[219] = L"PerSourceCustomDuration";
  v113[220] = &v83;
  v113[168] = 0LL;
  LODWORD(v113[169]) = 288;
  LODWORD(v113[172]) = 67108868;
  LODWORD(v113[174]) = 4;
  v113[175] = 0LL;
  LODWORD(v113[176]) = 288;
  LODWORD(v113[179]) = 67108868;
  LODWORD(v113[181]) = 4;
  v113[182] = 0LL;
  LODWORD(v113[183]) = 288;
  LODWORD(v113[186]) = 67108868;
  LODWORD(v113[188]) = 4;
  v113[189] = 0LL;
  LODWORD(v113[190]) = 288;
  LODWORD(v113[193]) = 67108868;
  LODWORD(v113[195]) = 4;
  v113[196] = 0LL;
  LODWORD(v113[197]) = 288;
  LODWORD(v113[200]) = 67108868;
  LODWORD(v113[202]) = 4;
  v113[203] = 0LL;
  LODWORD(v113[204]) = 288;
  LODWORD(v113[207]) = 67108868;
  LODWORD(v113[209]) = 4;
  v113[210] = 0LL;
  LODWORD(v113[211]) = 288;
  LODWORD(v113[214]) = 67108868;
  LODWORD(v113[216]) = 4;
  v113[217] = 0LL;
  LODWORD(v113[218]) = 288;
  LODWORD(v113[221]) = 67108868;
  v113[222] = &v112;
  LODWORD(v113[223]) = 4;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\Scheduler", v113, 0LL, 0LL);
  NodeConfiguration = VidSchiReadNodeConfiguration(a1, *(_QWORD *)(a1 + 2568));
  v8 = 0;
  if ( *(_DWORD *)(a1 + 80) )
  {
    v9 = NodeConfiguration;
    do
    {
      v10 = *(_DWORD *)(a1 + 2608);
      if ( v9 < 0 )
        goto LABEL_7;
      v45 = *(_DWORD **)(a1 + 2568);
      v12 = v8;
      if ( v8 < v10 )
        v45 += v8;
      if ( !*v45 )
      {
LABEL_7:
        v11 = *(int **)(a1 + 2568);
        v12 = v8;
        if ( v8 < v10 )
          v11 += v8;
        *v11 = v57;
      }
      v13 = *(_QWORD *)(a1 + 2568);
      v14 = *(_DWORD *)(a1 + 2608);
      v15 = (_DWORD *)(v13 + 4 * v12);
      if ( v8 >= v14 )
        v15 = *(_DWORD **)(a1 + 2568);
      if ( *v15 <= 1u )
      {
        v17 = 1;
      }
      else
      {
        v16 = (int *)(v13 + 4 * v12);
        if ( v8 >= v14 )
          v16 = *(int **)(a1 + 2568);
        v17 = *v16;
      }
      v18 = (int *)(v13 + 4 * v12);
      if ( v8 >= v14 )
        v18 = *(int **)(a1 + 2568);
      ++v8;
      *v18 = v17;
    }
    while ( v8 < *(_DWORD *)(a1 + 80) );
  }
  v19 = v64;
  *(_DWORD *)(a1 + 2536) = (v63 != 0 ? 0x400 : 0) | (v62 != 0 ? 0x100 : 0) | (v61 != 0 ? 0x10 : 0) | (v60 != 0) | (v59 != 0 ? 4 : 0) | (v58 != 0 ? 2 : 0) | *(_DWORD *)(a1 + 2536) & 0xFFFFFAE8;
  if ( !v19 || (unsigned int)(v19 - 1) <= 1 )
    *(_DWORD *)(a1 + 2548) = v19;
  if ( !v65 || (unsigned int)(v65 - 1) <= 1 )
    *(_DWORD *)(a1 + 2552) = v65;
  if ( !v66 || (unsigned int)(v66 - 1) <= 1 )
    *(_DWORD *)(a1 + 2556) = v66;
  v20 = v67;
  if ( (unsigned int)(v67 - 1) > 0x53 )
    v20 = v68;
  v21 = v69 == 0;
  *(_DWORD *)(a1 + 208) = v20;
  *(_DWORD *)(a1 + 212) = v20 + 15;
  v22 = (unsigned int)(10000 * v70);
  *(_BYTE *)(a1 + 57) = !v21;
  *(_QWORD *)(a1 + 2792) = 1000LL;
  *(_QWORD *)(a1 + 2800) = 2500LL;
  *(_QWORD *)(a1 + 2808) = 5000LL;
  *(_QWORD *)(a1 + 2768) = v22;
  v23 = (unsigned int)(10000 * v71);
  *(_QWORD *)(a1 + 2816) = 10000LL;
  *(_QWORD *)(a1 + 2824) = 25000LL;
  *(_QWORD *)(a1 + 2832) = 50000LL;
  *(_QWORD *)(a1 + 2840) = 100000LL;
  *(_QWORD *)(a1 + 2776) = v23;
  v24 = (unsigned int)(10000 * v72);
  *(_QWORD *)(a1 + 2848) = 250000LL;
  *(_QWORD *)(a1 + 2856) = 500000LL;
  *(_QWORD *)(a1 + 2864) = v24;
  *(_QWORD *)(a1 + 2784) = (unsigned int)(10000 * v73);
  *(_QWORD *)(a1 + 2872) = (unsigned int)(10000 * v74);
  v25 = v49;
  if ( v49 < 0x10 )
  {
    v25 = 16;
LABEL_64:
    v49 = v25;
    goto LABEL_30;
  }
  if ( v49 > 0x10000 )
  {
    v25 = 0x10000;
    v49 = 0x10000;
    goto LABEL_30;
  }
  if ( ((v49 - 1) & v49) != 0 )
  {
    WdLogSingleEntry1(1LL, v49);
    DxgkLogInternalTriageEvent(
      v46,
      0x40000,
      v47,
      (unsigned int)L"History log size value 0x%x is not a power of 2",
      v49,
      0LL,
      0LL,
      0LL);
    v25 = v75;
    goto LABEL_64;
  }
LABEL_30:
  *(_DWORD *)(a1 + 224) = v25;
  v26 = 14;
  if ( v50 <= 0xE )
  {
    v26 = v50;
    if ( !v50 )
      v26 = 1;
  }
  v50 = v26;
  *(_DWORD *)(a1 + 228) = v26;
  if ( !v76 || (IsEnabled = TdrIsEnabled(), v28 = 512, !IsEnabled) )
    v28 = 0;
  v29 = (_QWORD *)(a1 + 2680);
  v30 = v28 | *(_DWORD *)(a1 + 2536) & 0xFFFFFDFF;
  v31 = 0LL;
  *(_DWORD *)(a1 + 2536) = v30;
  do
  {
    v32 = 1;
    if ( v51 > 1 )
      v32 = v51;
    v33 = v52 <= 1;
    *(v29 - 6) = (unsigned int)(*(_DWORD *)((char *)&gulQuantumMultiplierTableByPriorityClass + v31) * v32);
    v34 = 1;
    if ( !v33 )
      v34 = v52;
    v35 = (unsigned int)(*(_DWORD *)((char *)&gulPreemptionQuantumMultiplierTableByPriorityClass + v31) * v34);
    v31 += 4LL;
    *v29++ = v35;
  }
  while ( v31 < 24 );
  v36 = 1;
  v37 = *(_DWORD *)(a1 + 2536);
  if ( v53 > 1 )
    v36 = v53;
  *(_DWORD *)(a1 + 2560) = v36;
  v38 = (v78 != 0 ? 0x40 : 0) | (v77 != 0 ? 0x20 : 0) | v37 & 0xFFFFFF9F;
  v39 = -(v79 != 0);
  *(_DWORD *)(a1 + 6472) = v80;
  v40 = v39 & 0x80 | v38 & 0xFFFFFF7F;
  v41 = v55;
  v33 = v54 <= 0x10;
  *(_DWORD *)(a1 + 2536) = v40;
  if ( !v33 )
    v4 = v54;
  v42 = *(_QWORD *)(a1 + 16);
  *(_DWORD *)(a1 + 2620) = v4;
  *(_DWORD *)(a1 + 2404) = v41;
  if ( *(int *)(v42 + 2820) < 1300 )
  {
    if ( v41 >= 4 )
    {
      v48 = v41;
      if ( v41 > 0xFFFFFFFD )
        v48 = -3;
      *(_DWORD *)(a1 + 2404) = v48;
    }
    else
    {
      *(_DWORD *)(a1 + 2404) = 4;
    }
  }
  v43 = v81;
  *(_DWORD *)(a1 + 2760) = v81;
  *(_DWORD *)(a1 + 2752) = v43;
  *(_DWORD *)(a1 + 2744) = v43;
  *(_DWORD *)(a1 + 2736) = v43;
  *(_DWORD *)(a1 + 2728) = v43;
  switch ( v82 )
  {
    case 0:
      if ( (**(_DWORD **)(v42 + 2824) & 0x1000) == 0 )
        break;
LABEL_49:
      *(_DWORD *)(a1 + 232) = 1;
      break;
    case 1:
      goto LABEL_49;
    case 2:
      *(_DWORD *)(a1 + 232) = 2;
      break;
    case 3:
      *(_DWORD *)(a1 + 232) = 0;
      break;
  }
  result = v83;
  *(_BYTE *)(a1 + 6633) = v83;
  return result;
}