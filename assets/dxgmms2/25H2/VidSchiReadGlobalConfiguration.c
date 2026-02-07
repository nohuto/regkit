__int64 __fastcall VidSchiReadGlobalConfiguration(__int64 a1)
{
  __int64 v1; // r8
  int v3; // edi
  int v4; // edx
  int v5; // r9d
  bool v6; // sf
  bool v7; // of
  int NodeConfiguration; // r11d
  unsigned int i; // ecx
  _DWORD *v10; // rdx
  int *v11; // rdx
  unsigned int v12; // r8d
  int *v13; // rdx
  int *v14; // r9
  int *v15; // r9
  int v16; // r10d
  int v17; // ecx
  int v18; // ecx
  bool v19; // zf
  __int64 v20; // rcx
  __int64 v21; // rcx
  __int64 v22; // rcx
  int v23; // ecx
  int v24; // ecx
  int v25; // r8d
  unsigned int v26; // eax
  __int64 *v27; // rcx
  __int64 v28; // rdx
  __int64 v29; // rax
  bool v30; // cc
  __int64 v31; // rax
  _QWORD *v32; // rdx
  __int64 j; // rcx
  int v34; // eax
  int v35; // eax
  __int64 v36; // rax
  int v37; // eax
  int v38; // edx
  unsigned int v39; // edx
  int v40; // ecx
  unsigned int v41; // eax
  unsigned int v42; // edx
  __int64 v43; // rcx
  __int64 result; // rax
  unsigned int v45; // [rsp+50h] [rbp-B0h] BYREF
  unsigned int v46; // [rsp+54h] [rbp-ACh] BYREF
  unsigned int v47; // [rsp+58h] [rbp-A8h] BYREF
  unsigned int v48; // [rsp+5Ch] [rbp-A4h] BYREF
  unsigned int v49; // [rsp+60h] [rbp-A0h] BYREF
  unsigned int v50; // [rsp+64h] [rbp-9Ch] BYREF
  unsigned int v51; // [rsp+68h] [rbp-98h] BYREF
  unsigned int v52; // [rsp+6Ch] [rbp-94h] BYREF
  int v53; // [rsp+70h] [rbp-90h] BYREF
  int v54; // [rsp+74h] [rbp-8Ch] BYREF
  int v55; // [rsp+78h] [rbp-88h] BYREF
  int v56; // [rsp+7Ch] [rbp-84h] BYREF
  int v57; // [rsp+80h] [rbp-80h] BYREF
  int v58; // [rsp+84h] [rbp-7Ch] BYREF
  int v59; // [rsp+88h] [rbp-78h] BYREF
  int v60; // [rsp+8Ch] [rbp-74h] BYREF
  int v61; // [rsp+90h] [rbp-70h] BYREF
  int v62; // [rsp+94h] [rbp-6Ch] BYREF
  int v63; // [rsp+98h] [rbp-68h] BYREF
  int v64; // [rsp+9Ch] [rbp-64h] BYREF
  int v65; // [rsp+A0h] [rbp-60h] BYREF
  int v66; // [rsp+A4h] [rbp-5Ch] BYREF
  int v67; // [rsp+A8h] [rbp-58h] BYREF
  int v68; // [rsp+ACh] [rbp-54h] BYREF
  unsigned int v69; // [rsp+B0h] [rbp-50h] BYREF
  int v70; // [rsp+B4h] [rbp-4Ch] BYREF
  int v71; // [rsp+B8h] [rbp-48h] BYREF
  int v72; // [rsp+BCh] [rbp-44h] BYREF
  int v73; // [rsp+C0h] [rbp-40h] BYREF
  int v74; // [rsp+C4h] [rbp-3Ch] BYREF
  int v75; // [rsp+C8h] [rbp-38h] BYREF
  int v76; // [rsp+CCh] [rbp-34h] BYREF
  int v77; // [rsp+D0h] [rbp-30h] BYREF
  int v78; // [rsp+D4h] [rbp-2Ch] BYREF
  BOOL v79; // [rsp+D8h] [rbp-28h] BYREF
  int v80; // [rsp+DCh] [rbp-24h] BYREF
  int v81; // [rsp+E0h] [rbp-20h] BYREF
  int v82; // [rsp+E4h] [rbp-1Ch] BYREF
  unsigned int v83; // [rsp+E8h] [rbp-18h] BYREF
  int v84; // [rsp+ECh] [rbp-14h] BYREF
  int v85; // [rsp+F0h] [rbp-10h] BYREF
  int v86; // [rsp+F4h] [rbp-Ch] BYREF
  int v87; // [rsp+F8h] [rbp-8h] BYREF
  int v88; // [rsp+FCh] [rbp-4h] BYREF
  int v89; // [rsp+100h] [rbp+0h] BYREF
  int v90; // [rsp+104h] [rbp+4h] BYREF
  int v91; // [rsp+108h] [rbp+8h] BYREF
  int v92; // [rsp+10Ch] [rbp+Ch] BYREF
  int v93; // [rsp+110h] [rbp+10h] BYREF
  int v94; // [rsp+114h] [rbp+14h] BYREF
  int v95; // [rsp+118h] [rbp+18h] BYREF
  int v96; // [rsp+11Ch] [rbp+1Ch] BYREF
  int v97; // [rsp+120h] [rbp+20h] BYREF
  int v98; // [rsp+124h] [rbp+24h] BYREF
  int v99; // [rsp+128h] [rbp+28h] BYREF
  int v100; // [rsp+12Ch] [rbp+2Ch] BYREF
  int v101; // [rsp+130h] [rbp+30h] BYREF
  int v102; // [rsp+134h] [rbp+34h] BYREF
  int v103; // [rsp+138h] [rbp+38h] BYREF
  int v104; // [rsp+13Ch] [rbp+3Ch] BYREF
  int v105; // [rsp+140h] [rbp+40h] BYREF
  int v106; // [rsp+144h] [rbp+44h] BYREF
  unsigned int v107; // [rsp+148h] [rbp+48h] BYREF
  int v108; // [rsp+14Ch] [rbp+4Ch] BYREF
  int v109; // [rsp+150h] [rbp+50h] BYREF
  BOOL v110; // [rsp+154h] [rbp+54h] BYREF
  int v111; // [rsp+158h] [rbp+58h] BYREF
  int v112; // [rsp+15Ch] [rbp+5Ch] BYREF
  int v113; // [rsp+160h] [rbp+60h] BYREF
  int v114; // [rsp+164h] [rbp+64h] BYREF
  _QWORD v115[252]; // [rsp+170h] [rbp+70h] BYREF

  v1 = *(_QWORD *)(a1 + 16);
  v84 = 0;
  v85 = 25000;
  v86 = 50000;
  v91 = 3;
  v89 = 1;
  v90 = 2;
  v3 = 16;
  v93 = 1;
  v4 = 100;
  v96 = 2;
  v92 = 0;
  v94 = 0;
  v98 = 0;
  v95 = 20;
  v54 = 7;
  v5 = *(_DWORD *)(*(_QWORD *)(v1 + 3008) + 4LL) & 1;
  v99 = 900;
  v97 = v5;
  v100 = 1000;
  v101 = 8;
  v102 = 0;
  v64 = 10;
  v103 = 1;
  v104 = 0;
  v108 = 0;
  v109 = 0;
  v105 = 16;
  v106 = 8000;
  v55 = 100;
  if ( *(_BYTE *)(a1 + 7062) )
  {
    v4 = 10;
    v55 = 10;
  }
  v107 = *(_DWORD *)(a1 + 244);
  v72 = 64;
  v7 = __OFSUB__(*(_DWORD *)(v1 + 3004), 2000);
  v6 = *(_DWORD *)(v1 + 3004) - 2000 < 0;
  v111 = 2;
  v112 = 1;
  v110 = v6 == v7;
  v79 = v110;
  v113 = 24;
  v114 = 17;
  v87 = 30000;
  v88 = 60000;
  v57 = 0;
  v49 = 25000;
  v50 = 50000;
  v59 = 1;
  v56 = 2;
  v51 = 3;
  v58 = 0;
  v74 = 1;
  v75 = 0;
  v76 = 0;
  v52 = 20;
  v77 = 2;
  v53 = 7;
  v73 = v5;
  v66 = 900;
  v67 = 1000;
  v70 = 8;
  v65 = 0;
  v63 = 10;
  v60 = 1;
  v61 = 0;
  v62 = 0;
  v68 = 16;
  v69 = 8000;
  v71 = v4;
  v45 = 64;
  v78 = 0;
  v80 = 2;
  v81 = 1;
  v82 = 24;
  v83 = 17;
  v47 = 30000;
  v48 = 60000;
  v46 = v107;
  if ( *(int *)(v1 + 3004) >= 1300 && *(_BYTE *)(v1 + 2941) )
  {
    v54 = 1;
    v53 = 1;
  }
  memset(v115, 0, sizeof(v115));
  v115[7] = 0LL;
  LODWORD(v115[1]) = 288;
  LODWORD(v115[4]) = 67108868;
  LODWORD(v115[6]) = 4;
  v115[2] = L"AutoSyncToCPUPriority";
  v115[3] = &v57;
  v115[5] = &v84;
  v115[9] = L"QuantumUnit";
  v115[10] = &v49;
  v115[12] = &v85;
  v115[16] = L"PreemptionQuantumUnit";
  v115[17] = &v50;
  v115[19] = &v86;
  v115[23] = L"NpuContextSwitchQuantum";
  v115[24] = &v47;
  v115[26] = &v87;
  v115[30] = L"NpuPreemptionQuantum";
  v115[31] = &v48;
  v115[33] = &v88;
  v115[37] = L"EnablePreemption";
  v115[38] = &v59;
  v115[40] = &v89;
  v115[44] = L"HwQueuedRenderPacketGroupLimit";
  v115[45] = &v56;
  v115[47] = &v90;
  v115[51] = L"QueuedPresentLimit";
  LODWORD(v115[8]) = 288;
  LODWORD(v115[11]) = 67108868;
  LODWORD(v115[13]) = 4;
  v115[14] = 0LL;
  LODWORD(v115[15]) = 288;
  LODWORD(v115[18]) = 67108868;
  LODWORD(v115[20]) = 4;
  v115[21] = 0LL;
  LODWORD(v115[22]) = 288;
  LODWORD(v115[25]) = 67108868;
  LODWORD(v115[27]) = 4;
  v115[28] = 0LL;
  LODWORD(v115[29]) = 288;
  LODWORD(v115[32]) = 67108868;
  LODWORD(v115[34]) = 4;
  v115[35] = 0LL;
  LODWORD(v115[36]) = 288;
  LODWORD(v115[39]) = 67108868;
  LODWORD(v115[41]) = 4;
  v115[42] = 0LL;
  LODWORD(v115[43]) = 288;
  LODWORD(v115[46]) = 67108868;
  LODWORD(v115[48]) = 4;
  v115[49] = 0LL;
  LODWORD(v115[50]) = 288;
  v115[52] = &v51;
  v115[54] = &v91;
  v115[58] = L"CarryOverUsedQuantum";
  v115[59] = &v58;
  v115[61] = &v92;
  v115[65] = L"AdjustWorkerThreadPriority";
  v115[66] = &v74;
  v115[68] = &v93;
  v115[72] = L"CountFlipTowardHwLimit";
  v115[73] = &v75;
  v115[75] = &v94;
  v115[79] = L"NumberOfDmaPacketPool";
  v115[80] = &v52;
  v115[82] = &v95;
  v115[86] = L"ProfileLevel";
  v115[87] = &v77;
  v115[89] = &v96;
  v115[93] = L"VSyncIdleTimeout";
  v115[94] = &v53;
  v115[96] = &v54;
  v115[100] = L"EnableDirectSubmission";
  v115[101] = &v73;
  v115[103] = &v97;
  v115[107] = L"CountPresentTowardHwLimit";
  v115[108] = &v76;
  LODWORD(v115[53]) = 67108868;
  LODWORD(v115[55]) = 4;
  v115[56] = 0LL;
  LODWORD(v115[57]) = 288;
  LODWORD(v115[60]) = 67108868;
  LODWORD(v115[62]) = 4;
  v115[63] = 0LL;
  LODWORD(v115[64]) = 288;
  LODWORD(v115[67]) = 67108868;
  LODWORD(v115[69]) = 4;
  v115[70] = 0LL;
  LODWORD(v115[71]) = 288;
  LODWORD(v115[74]) = 67108868;
  LODWORD(v115[76]) = 4;
  v115[77] = 0LL;
  LODWORD(v115[78]) = 288;
  LODWORD(v115[81]) = 67108868;
  LODWORD(v115[83]) = 4;
  v115[84] = 0LL;
  LODWORD(v115[85]) = 288;
  LODWORD(v115[88]) = 67108868;
  LODWORD(v115[90]) = 4;
  v115[91] = 0LL;
  LODWORD(v115[92]) = 288;
  LODWORD(v115[95]) = 67108868;
  LODWORD(v115[97]) = 4;
  v115[98] = 0LL;
  LODWORD(v115[99]) = 288;
  LODWORD(v115[102]) = 67108868;
  LODWORD(v115[104]) = 4;
  v115[105] = 0LL;
  LODWORD(v115[106]) = 288;
  LODWORD(v115[109]) = 67108868;
  v115[110] = &v98;
  v115[114] = L"MaximumAllowedPreemptionDelay";
  v115[115] = &v66;
  v115[117] = &v99;
  v115[121] = L"ContextSchedulingPenaltyDelay";
  v115[122] = &v67;
  v115[124] = &v100;
  v115[128] = L"BackgroundProcessMaximumAllowedPreemptionDelay";
  v115[129] = &v70;
  v115[131] = &v101;
  v115[135] = L"ForceEnableFlipFenceModel";
  v115[136] = &v65;
  v115[138] = &v102;
  v115[142] = L"YieldPercentage";
  v115[143] = &v63;
  v115[145] = &v64;
  v115[149] = L"ForegroundPriorityBoost";
  v115[150] = &v60;
  v115[152] = &v103;
  v115[156] = L"ForceFlipTrueImmediateMode";
  v115[157] = &v61;
  v115[159] = &v104;
  v115[163] = L"MaxYieldInterval";
  v115[164] = &v68;
  LODWORD(v115[111]) = 4;
  v115[112] = 0LL;
  LODWORD(v115[113]) = 288;
  LODWORD(v115[116]) = 67108868;
  LODWORD(v115[118]) = 4;
  v115[119] = 0LL;
  LODWORD(v115[120]) = 288;
  LODWORD(v115[123]) = 67108868;
  LODWORD(v115[125]) = 4;
  v115[126] = 0LL;
  LODWORD(v115[127]) = 288;
  LODWORD(v115[130]) = 67108868;
  LODWORD(v115[132]) = 4;
  v115[133] = 0LL;
  LODWORD(v115[134]) = 288;
  LODWORD(v115[137]) = 67108868;
  LODWORD(v115[139]) = 4;
  v115[140] = 0LL;
  LODWORD(v115[141]) = 288;
  LODWORD(v115[144]) = 67108868;
  LODWORD(v115[146]) = 4;
  v115[147] = 0LL;
  LODWORD(v115[148]) = 288;
  LODWORD(v115[151]) = 67108868;
  LODWORD(v115[153]) = 4;
  v115[154] = 0LL;
  LODWORD(v115[155]) = 288;
  LODWORD(v115[158]) = 67108868;
  LODWORD(v115[160]) = 4;
  v115[161] = 0LL;
  LODWORD(v115[162]) = 288;
  LODWORD(v115[165]) = 67108868;
  LODWORD(v115[167]) = 4;
  v115[166] = &v105;
  v115[170] = L"MinYieldInterval";
  v115[171] = &v69;
  v115[173] = &v106;
  v115[177] = L"MaxFocusGpuQuantumWithoutPresent";
  v115[178] = &v71;
  v115[180] = &v55;
  v115[184] = L"HistoryLogSize";
  v115[185] = &v45;
  v115[187] = &v72;
  v115[191] = L"HwQueuePacketCap";
  v115[192] = &v46;
  v115[194] = &v107;
  v115[198] = L"FlipDoNotFlipMode";
  v115[199] = &v62;
  v115[201] = &v108;
  v115[205] = L"PfnCpuOverride";
  v115[206] = &v78;
  v115[208] = &v109;
  v115[212] = L"PerSourceCustomDuration";
  v115[213] = &v79;
  v115[215] = &v110;
  v115[219] = L"HwSchThreadOffloadMode";
  v115[220] = &v80;
  v115[168] = 0LL;
  LODWORD(v115[169]) = 288;
  LODWORD(v115[172]) = 67108868;
  LODWORD(v115[174]) = 4;
  v115[175] = 0LL;
  LODWORD(v115[176]) = 288;
  LODWORD(v115[179]) = 67108868;
  LODWORD(v115[181]) = 4;
  v115[182] = 0LL;
  LODWORD(v115[183]) = 288;
  LODWORD(v115[186]) = 67108868;
  LODWORD(v115[188]) = 4;
  v115[189] = 0LL;
  LODWORD(v115[190]) = 288;
  LODWORD(v115[193]) = 67108868;
  LODWORD(v115[195]) = 4;
  v115[196] = 0LL;
  LODWORD(v115[197]) = 288;
  LODWORD(v115[200]) = 67108868;
  LODWORD(v115[202]) = 4;
  v115[203] = 0LL;
  LODWORD(v115[204]) = 288;
  LODWORD(v115[207]) = 67108868;
  LODWORD(v115[209]) = 4;
  v115[210] = 0LL;
  LODWORD(v115[211]) = 288;
  LODWORD(v115[214]) = 67108868;
  LODWORD(v115[216]) = 4;
  v115[217] = 0LL;
  LODWORD(v115[218]) = 288;
  LODWORD(v115[221]) = 67108868;
  v115[222] = &v111;
  LODWORD(v115[225]) = 288;
  v115[226] = L"DebugLargeSmoothenedDuration";
  LODWORD(v115[228]) = 67108868;
  v115[227] = &v81;
  v115[229] = &v112;
  v115[233] = L"AudioDgAutoBoostPriority";
  v115[234] = &v82;
  v115[236] = &v113;
  v115[240] = L"FrameServerAutoBoostPriority";
  v115[241] = &v83;
  LODWORD(v115[232]) = 288;
  LODWORD(v115[235]) = 67108868;
  LODWORD(v115[239]) = 288;
  LODWORD(v115[242]) = 67108868;
  v115[243] = &v114;
  LODWORD(v115[223]) = 4;
  v115[224] = 0LL;
  LODWORD(v115[230]) = 4;
  v115[231] = 0LL;
  LODWORD(v115[237]) = 4;
  v115[238] = 0LL;
  LODWORD(v115[244]) = 4;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\Scheduler", v115, 0LL, 0LL);
  NodeConfiguration = VidSchiReadNodeConfiguration(a1, *(_QWORD *)(a1 + 2816));
  for ( i = 0; i < *(_DWORD *)(a1 + 80); *v13 = v16 )
  {
    if ( NodeConfiguration < 0 )
      goto LABEL_11;
    v10 = *(_DWORD **)(a1 + 2816);
    if ( i < *(_DWORD *)(a1 + 2856) )
      v10 += i;
    if ( !*v10 )
    {
LABEL_11:
      v11 = *(int **)(a1 + 2816);
      if ( i < *(_DWORD *)(a1 + 2856) )
        v11 += i;
      *v11 = v56;
    }
    v12 = *(_DWORD *)(a1 + 2856);
    v13 = *(int **)(a1 + 2816);
    if ( i >= v12 )
      v14 = *(int **)(a1 + 2816);
    else
      v14 = &v13[i];
    if ( (unsigned int)*v14 <= 1 )
    {
      v16 = 1;
    }
    else
    {
      if ( i >= v12 )
        v15 = *(int **)(a1 + 2816);
      else
        v15 = &v13[i];
      v16 = *v15;
    }
    if ( i < v12 )
      v13 += i;
    ++i;
  }
  v17 = v61;
  *(_DWORD *)(a1 + 2792) = (v60 != 0 ? 0x200 : 0) | (v59 != 0) | (v58 != 0 ? 8 : 0) | (v57 != 0 ? 4 : 0) | *(_DWORD *)(a1 + 2792) & 0xFFFFFDF2;
  if ( !v17 || (unsigned int)(v17 - 1) <= 1 )
    *(_DWORD *)(a1 + 2804) = v17;
  if ( !v62 || (unsigned int)(v62 - 1) <= 1 )
    *(_DWORD *)(a1 + 2808) = v62;
  v18 = v63;
  if ( (unsigned int)(v63 - 1) > 0x53 )
    v18 = v64;
  v19 = v65 == 0;
  *(_DWORD *)(a1 + 224) = v18;
  *(_DWORD *)(a1 + 228) = v18 + 15;
  v20 = (unsigned int)(10000 * v66);
  *(_BYTE *)(a1 + 57) = !v19;
  *(_QWORD *)(a1 + 3080) = v69;
  *(_QWORD *)(a1 + 3000) = 1000LL;
  *(_QWORD *)(a1 + 2976) = v20;
  v21 = (unsigned int)(10000 * v67);
  *(_QWORD *)(a1 + 3008) = 2500LL;
  *(_QWORD *)(a1 + 3016) = 5000LL;
  *(_QWORD *)(a1 + 3024) = 10000LL;
  *(_QWORD *)(a1 + 3032) = 25000LL;
  *(_QWORD *)(a1 + 2984) = v21;
  v22 = (unsigned int)(10000 * v68);
  *(_QWORD *)(a1 + 3040) = 50000LL;
  *(_QWORD *)(a1 + 3048) = 100000LL;
  *(_QWORD *)(a1 + 3056) = 250000LL;
  *(_QWORD *)(a1 + 3064) = 500000LL;
  *(_QWORD *)(a1 + 3072) = v22;
  *(_QWORD *)(a1 + 2992) = (unsigned int)(10000 * v70);
  *(_QWORD *)(a1 + 3088) = (unsigned int)(10000 * v71);
  v23 = v45;
  if ( v45 < 0x10 )
  {
    v23 = 16;
LABEL_40:
    v45 = v23;
    goto LABEL_41;
  }
  if ( v45 > 0x10000 )
  {
    v23 = 0x10000;
    v45 = 0x10000;
    goto LABEL_41;
  }
  if ( ((v45 - 1) & v45) != 0 )
  {
    WdLogSingleEntry1(1LL, v45);
    WdLogGlobalForLineNumber = 15446;
    DxgkLogInternalTriageEvent(
      v24,
      0x40000,
      v25,
      (unsigned int)L"History log size value 0x%x is not a power of 2",
      v45,
      0LL,
      0LL,
      0LL);
    v23 = v72;
    goto LABEL_40;
  }
LABEL_41:
  v26 = v46;
  *(_DWORD *)(a1 + 240) = v23;
  if ( v26 <= 0xE )
  {
    if ( v26 )
      goto LABEL_46;
    v26 = 1;
  }
  else
  {
    v26 = 14;
  }
  v46 = v26;
LABEL_46:
  *(_DWORD *)(a1 + 244) = v26;
  *(_DWORD *)(a1 + 2792) = (TdrIsEnabled() << 8) | *(_DWORD *)(a1 + 2792) & 0xFFFFFEFF;
  if ( *(_BYTE *)(a1 + 7062) )
  {
    v27 = (__int64 *)(a1 + 2928);
    v28 = 6LL;
    do
    {
      v29 = 1LL;
      if ( v47 > 1 )
        v29 = v47;
      v30 = v48 <= 1;
      *(v27 - 6) = v29;
      v31 = 1LL;
      if ( !v30 )
        v31 = v48;
      *v27++ = v31;
      --v28;
    }
    while ( v28 );
  }
  else
  {
    v32 = (_QWORD *)(a1 + 2928);
    for ( j = 0LL; j < 24; j += 4LL )
    {
      v34 = 1;
      if ( v49 > 1 )
        v34 = v49;
      v30 = v50 <= 1;
      *(v32 - 6) = (unsigned int)(*(_DWORD *)((char *)&gulQuantumMultiplierTableByPriorityClass + j) * v34);
      v35 = 1;
      if ( !v30 )
        v35 = v50;
      v36 = (unsigned int)(*(_DWORD *)((char *)&gulPreemptionQuantumMultiplierTableByPriorityClass + j) * v35);
      *v32++ = v36;
    }
  }
  v37 = 1;
  v38 = *(_DWORD *)(a1 + 2792);
  if ( v51 > 1 )
    v37 = v51;
  *(_DWORD *)(a1 + 2812) = v37;
  v39 = (v75 != 0 ? 0x40 : 0) | (v74 != 0 ? 0x20 : 0) | (v73 != 0 ? 2 : 0) | v38 & 0xFFFFFF9D;
  v40 = -(v76 != 0);
  *(_DWORD *)(a1 + 6704) = v77;
  v41 = v53;
  *(_DWORD *)(a1 + 2660) = v53;
  v42 = v40 & 0x80 | v39 & 0xFFFFFF7F;
  v43 = *(_QWORD *)(a1 + 16);
  v30 = v52 <= 0x10;
  *(_DWORD *)(a1 + 2792) = v42;
  if ( !v30 )
    v3 = v52;
  *(_DWORD *)(a1 + 2868) = v3;
  if ( *(int *)(v43 + 3004) < 1300 )
  {
    if ( v41 >= 4 )
    {
      if ( v41 > 0xFFFFFFFD )
        *(_DWORD *)(a1 + 2660) = -3;
    }
    else
    {
      *(_DWORD *)(a1 + 2660) = 4;
    }
  }
  switch ( v78 )
  {
    case 0:
      if ( (**(_DWORD **)(v43 + 3008) & 0x1000) == 0 )
        break;
LABEL_76:
      *(_DWORD *)(a1 + 248) = 1;
      break;
    case 1:
      goto LABEL_76;
    case 2:
      *(_DWORD *)(a1 + 248) = 2;
      break;
    case 3:
      *(_DWORD *)(a1 + 248) = 0;
      break;
  }
  v19 = v81 == 0;
  *(_BYTE *)(a1 + 7057) = v79;
  *(_DWORD *)(a1 + 304) = v80;
  *(_BYTE *)(a1 + 7065) = !v19;
  *(_DWORD *)(a1 + 208) = v82;
  result = v83;
  *(_DWORD *)(a1 + 212) = v83;
  return result;
}