void VIDMM_GLOBAL::ReadConfiguration(void)
{
  int v0; // ebx
  unsigned __int64 v1; // rsi
  __int64 PhysicalMemoryRanges; // rax
  _QWORD *v3; // rcx
  __int64 v4; // rax
  int v5; // edx
  int v6; // eax
  int v7; // eax
  int v8; // eax
  unsigned int *v9; // r14
  ULONG v10; // esi
  __int64 v11; // r15
  unsigned int v12; // eax
  int v13; // eax
  int v14; // eax
  int v15; // eax
  int v16; // ecx
  int v17; // r8d
  unsigned int v18; // eax
  unsigned int v19; // [rsp+58h] [rbp-B0h] BYREF
  unsigned int v20; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v21; // [rsp+60h] [rbp-A8h] BYREF
  unsigned int v22; // [rsp+64h] [rbp-A4h] BYREF
  int v23; // [rsp+68h] [rbp-A0h] BYREF
  int v24; // [rsp+6Ch] [rbp-9Ch] BYREF
  int v25; // [rsp+70h] [rbp-98h] BYREF
  int v26; // [rsp+74h] [rbp-94h] BYREF
  int v27; // [rsp+78h] [rbp-90h] BYREF
  int v28; // [rsp+7Ch] [rbp-8Ch] BYREF
  int v29; // [rsp+80h] [rbp-88h] BYREF
  int v30; // [rsp+84h] [rbp-84h] BYREF
  int v31; // [rsp+88h] [rbp-80h] BYREF
  int v32; // [rsp+8Ch] [rbp-7Ch] BYREF
  int v33; // [rsp+90h] [rbp-78h] BYREF
  int v34; // [rsp+94h] [rbp-74h] BYREF
  int v35; // [rsp+98h] [rbp-70h] BYREF
  int v36; // [rsp+9Ch] [rbp-6Ch] BYREF
  int v37; // [rsp+A0h] [rbp-68h] BYREF
  int v38; // [rsp+A4h] [rbp-64h] BYREF
  int v39; // [rsp+A8h] [rbp-60h] BYREF
  int v40; // [rsp+ACh] [rbp-5Ch] BYREF
  int v41; // [rsp+B0h] [rbp-58h] BYREF
  int v42; // [rsp+B4h] [rbp-54h] BYREF
  int v43; // [rsp+B8h] [rbp-50h] BYREF
  int v44; // [rsp+BCh] [rbp-4Ch] BYREF
  int v45; // [rsp+C0h] [rbp-48h] BYREF
  int v46; // [rsp+C4h] [rbp-44h] BYREF
  int v47; // [rsp+C8h] [rbp-40h] BYREF
  unsigned int v48; // [rsp+CCh] [rbp-3Ch] BYREF
  int v49; // [rsp+D0h] [rbp-38h] BYREF
  int v50; // [rsp+D4h] [rbp-34h] BYREF
  int v51; // [rsp+D8h] [rbp-30h] BYREF
  int v52; // [rsp+DCh] [rbp-2Ch] BYREF
  unsigned int v53; // [rsp+E0h] [rbp-28h] BYREF
  int v54; // [rsp+E4h] [rbp-24h] BYREF
  int v55; // [rsp+E8h] [rbp-20h] BYREF
  int v56; // [rsp+ECh] [rbp-1Ch] BYREF
  int v57; // [rsp+F0h] [rbp-18h] BYREF
  int v58; // [rsp+F4h] [rbp-14h] BYREF
  int v59; // [rsp+F8h] [rbp-10h] BYREF
  int v60; // [rsp+FCh] [rbp-Ch] BYREF
  int v61; // [rsp+100h] [rbp-8h] BYREF
  int v62; // [rsp+104h] [rbp-4h] BYREF
  int v63; // [rsp+108h] [rbp+0h] BYREF
  int v64; // [rsp+10Ch] [rbp+4h] BYREF
  int v65; // [rsp+110h] [rbp+8h] BYREF
  int v66; // [rsp+114h] [rbp+Ch] BYREF
  int v67; // [rsp+118h] [rbp+10h] BYREF
  int v68; // [rsp+11Ch] [rbp+14h] BYREF
  int v69; // [rsp+120h] [rbp+18h] BYREF
  int v70; // [rsp+124h] [rbp+1Ch] BYREF
  int v71; // [rsp+128h] [rbp+20h] BYREF
  int v72; // [rsp+12Ch] [rbp+24h] BYREF
  int v73; // [rsp+130h] [rbp+28h] BYREF
  int v74; // [rsp+134h] [rbp+2Ch] BYREF
  int v75; // [rsp+138h] [rbp+30h] BYREF
  int v76; // [rsp+13Ch] [rbp+34h] BYREF
  int v77; // [rsp+140h] [rbp+38h] BYREF
  int v78; // [rsp+144h] [rbp+3Ch] BYREF
  int v79; // [rsp+148h] [rbp+40h] BYREF
  int v80; // [rsp+14Ch] [rbp+44h] BYREF
  int v81; // [rsp+150h] [rbp+48h] BYREF
  int v82; // [rsp+154h] [rbp+4Ch] BYREF
  int v83; // [rsp+158h] [rbp+50h] BYREF
  int v84; // [rsp+15Ch] [rbp+54h] BYREF
  struct _UNICODE_STRING Destination; // [rsp+160h] [rbp+58h] BYREF
  struct _UNICODE_STRING String; // [rsp+170h] [rbp+68h] BYREF
  __int64 v87; // [rsp+180h] [rbp+78h] BYREF
  __int64 v88; // [rsp+188h] [rbp+80h] BYREF
  __int64 v89; // [rsp+190h] [rbp+88h] BYREF
  __int128 v90; // [rsp+198h] [rbp+90h]
  __int128 v91; // [rsp+1A8h] [rbp+A0h]
  __int128 v92; // [rsp+1B8h] [rbp+B0h]
  __int64 v93; // [rsp+1C8h] [rbp+C0h]
  struct _UNICODE_STRING DestinationString; // [rsp+1D0h] [rbp+C8h] BYREF
  __int64 v95; // [rsp+1E0h] [rbp+D8h] BYREF
  SIZE_T v96; // [rsp+1E8h] [rbp+E0h]
  __int64 v97; // [rsp+1F0h] [rbp+E8h] BYREF
  PHYSICAL_ADDRESS v98; // [rsp+1F8h] [rbp+F0h]
  __int64 v99; // [rsp+200h] [rbp+F8h] BYREF
  PHYSICAL_ADDRESS v100; // [rsp+208h] [rbp+100h]
  _OWORD v101[126]; // [rsp+218h] [rbp+110h] BYREF
  char v102; // [rsp+9F8h] [rbp+8F0h] BYREF
  _BYTE v103[64]; // [rsp+A08h] [rbp+900h] BYREF

  v0 = 0;
  v1 = 0LL;
  PhysicalMemoryRanges = MmGetPhysicalMemoryRangesEx(0LL);
  v3 = (_QWORD *)PhysicalMemoryRanges;
  if ( PhysicalMemoryRanges )
  {
    v4 = *(_QWORD *)(PhysicalMemoryRanges + 8);
    v5 = 0;
    while ( v4 )
    {
      v1 += v4;
      v4 = v3[2 * (unsigned int)++v5 + 1];
    }
    ExFreePoolWithTag(v3, 0);
  }
  else
  {
    _InterlockedAdd(&dword_1C0076890, 1u);
    WdLogSingleEntry1(6LL, 47LL);
    DxgkLogInternalTriageEvent(
      v16,
      262145,
      v17,
      (unsigned int)L"Couldn't allocate buffer to query system memory size",
      47LL,
      0LL,
      0LL,
      0LL);
    v1 = 0x20000000LL;
  }
  qword_1C0076290 = v1;
  v23 = 25;
  v19 = 25;
  qword_1C0076288 = v1;
  v24 = 40;
  v20 = 40;
  v54 = 0;
  v21 = 0;
  v55 = 10;
  v56 = 15;
  v36 = 15;
  v35 = 10;
  v57 = 5;
  v37 = 5;
  v58 = 300;
  v38 = 300;
  v6 = 256;
  if ( v1 > 0x20000000 )
    v6 = 1024;
  v64 = 10;
  v60 = v6;
  v26 = v6;
  v7 = 0x800000;
  if ( v1 > 0x20000000 )
    v7 = 0x2000000;
  v30 = 10;
  v62 = v7;
  v28 = v7;
  v8 = 0x400000;
  if ( v1 > 0x20000000 )
    v8 = 0x1000000;
  v59 = 0;
  v63 = v8;
  v29 = v8;
  v25 = 0;
  v61 = 4;
  v27 = 4;
  v66 = 1;
  v32 = 1;
  v67 = 1;
  v34 = 1;
  v65 = g_IsInternalRelease != 0 ? 0x40 : 0;
  v31 = v65;
  v69 = 0x100000;
  v39 = 0x100000;
  v68 = 1;
  v33 = 1;
  v74 = 8;
  v44 = 8;
  v70 = 0x800000;
  v40 = 0x800000;
  v71 = 60;
  v41 = 60;
  v72 = 60;
  v42 = 60;
  v73 = 1;
  v43 = 1;
  v75 = 2;
  v45 = 2;
  v76 = 0;
  v46 = 0;
  v77 = 0;
  v47 = 0;
  v79 = 1;
  v49 = 1;
  v78 = 200;
  v48 = 200;
  v80 = 4096;
  v50 = 4096;
  v82 = 20;
  v52 = 20;
  v89 = 0xFFFFFFFFLL;
  v100.QuadPart = 0xFFFFFFFFLL;
  v83 = 900;
  v53 = 900;
  *(_QWORD *)&v101[1] = L"PinnedMemoryLimit";
  *((_QWORD *)&v101[1] + 1) = &v19;
  *((_QWORD *)&v101[2] + 1) = &v23;
  *((_QWORD *)&v101[4] + 1) = L"PinnedApertureMemoryLimit";
  *(_QWORD *)&v101[5] = &v20;
  *(_QWORD *)&v101[6] = &v24;
  *(_QWORD *)&v101[8] = L"PagesHistory";
  *((_QWORD *)&v101[8] + 1) = &v21;
  *((_QWORD *)&v101[9] + 1) = &v54;
  *((_QWORD *)&v101[11] + 1) = L"MemTransferThreshold";
  *(_QWORD *)&v101[12] = &v35;
  *(_QWORD *)&v101[13] = &v55;
  *(_QWORD *)&v101[15] = L"ExcessiveMemTransferFlipThreshold";
  *((_QWORD *)&v101[15] + 1) = &v36;
  *((_QWORD *)&v101[16] + 1) = &v56;
  v81 = 6;
  v51 = 6;
  v87 = 0LL;
  v95 = 16LL;
  v96 = 0LL;
  v88 = 0LL;
  v97 = 16LL;
  v98.QuadPart = 0LL;
  v99 = 16LL;
  *(_QWORD *)&v101[0] = 0LL;
  DWORD2(v101[0]) = 288;
  LODWORD(v101[2]) = 67108868;
  LODWORD(v101[3]) = 4;
  *((_QWORD *)&v101[3] + 1) = 0LL;
  LODWORD(v101[4]) = 288;
  DWORD2(v101[5]) = 67108868;
  DWORD2(v101[6]) = 4;
  *(_QWORD *)&v101[7] = 0LL;
  DWORD2(v101[7]) = 288;
  LODWORD(v101[9]) = 67108868;
  LODWORD(v101[10]) = 4;
  *((_QWORD *)&v101[10] + 1) = 0LL;
  LODWORD(v101[11]) = 288;
  DWORD2(v101[12]) = 67108868;
  DWORD2(v101[13]) = 4;
  *(_QWORD *)&v101[14] = 0LL;
  DWORD2(v101[14]) = 288;
  LODWORD(v101[16]) = 67108868;
  LODWORD(v101[17]) = 4;
  *((_QWORD *)&v101[17] + 1) = 0LL;
  LODWORD(v101[18]) = 288;
  *((_QWORD *)&v101[18] + 1) = L"ExcessiveMemTransferPenalty";
  *(_QWORD *)&v101[19] = &v37;
  *(_QWORD *)&v101[20] = &v57;
  *(_QWORD *)&v101[22] = L"EventThrottleThreshold";
  *((_QWORD *)&v101[22] + 1) = &v38;
  *((_QWORD *)&v101[23] + 1) = &v58;
  *((_QWORD *)&v101[25] + 1) = L"DisablePrefetching";
  *(_QWORD *)&v101[26] = &v25;
  *(_QWORD *)&v101[27] = &v59;
  *(_QWORD *)&v101[29] = L"NbDmaBufferLimitPerDevice";
  *((_QWORD *)&v101[29] + 1) = &v26;
  *((_QWORD *)&v101[30] + 1) = &v60;
  *((_QWORD *)&v101[32] + 1) = L"NbCddDmaBufferLimitPerDevice";
  *(_QWORD *)&v101[33] = &v27;
  *(_QWORD *)&v101[34] = &v61;
  *(_QWORD *)&v101[36] = L"DmaBufferBytesLimitAllDevices";
  *((_QWORD *)&v101[36] + 1) = &v28;
  *((_QWORD *)&v101[37] + 1) = &v62;
  *((_QWORD *)&v101[39] + 1) = L"DmaBufferListBytesLimitAllDevices";
  *(_QWORD *)&v101[40] = &v29;
  *(_QWORD *)&v101[41] = &v63;
  *(_QWORD *)&v101[43] = L"NbDmaBufferLimitCompareWatermark";
  *((_QWORD *)&v101[43] + 1) = &v30;
  *((_QWORD *)&v101[44] + 1) = &v64;
  *((_QWORD *)&v101[46] + 1) = L"NbPagingHistoryRecords";
  DWORD2(v101[19]) = 67108868;
  DWORD2(v101[20]) = 4;
  *(_QWORD *)&v101[21] = 0LL;
  DWORD2(v101[21]) = 288;
  LODWORD(v101[23]) = 67108868;
  LODWORD(v101[24]) = 4;
  *((_QWORD *)&v101[24] + 1) = 0LL;
  LODWORD(v101[25]) = 288;
  DWORD2(v101[26]) = 67108868;
  DWORD2(v101[27]) = 4;
  *(_QWORD *)&v101[28] = 0LL;
  DWORD2(v101[28]) = 288;
  LODWORD(v101[30]) = 67108868;
  LODWORD(v101[31]) = 4;
  *((_QWORD *)&v101[31] + 1) = 0LL;
  LODWORD(v101[32]) = 288;
  DWORD2(v101[33]) = 67108868;
  DWORD2(v101[34]) = 4;
  *(_QWORD *)&v101[35] = 0LL;
  DWORD2(v101[35]) = 288;
  LODWORD(v101[37]) = 67108868;
  LODWORD(v101[38]) = 4;
  *((_QWORD *)&v101[38] + 1) = 0LL;
  LODWORD(v101[39]) = 288;
  DWORD2(v101[40]) = 67108868;
  DWORD2(v101[41]) = 4;
  *(_QWORD *)&v101[42] = 0LL;
  DWORD2(v101[42]) = 288;
  LODWORD(v101[44]) = 67108868;
  LODWORD(v101[45]) = 4;
  *((_QWORD *)&v101[45] + 1) = 0LL;
  LODWORD(v101[46]) = 288;
  DWORD2(v101[47]) = 67108868;
  *(_QWORD *)&v101[47] = &v31;
  *(_QWORD *)&v101[48] = &v65;
  *(_QWORD *)&v101[50] = L"PinDWMAllocationBackingStore";
  *((_QWORD *)&v101[50] + 1) = &v32;
  *((_QWORD *)&v101[51] + 1) = &v66;
  *((_QWORD *)&v101[53] + 1) = L"RemovePagesFromWorkingSetOnPagingForDwm";
  *(_QWORD *)&v101[54] = &v34;
  *(_QWORD *)&v101[55] = &v67;
  *(_QWORD *)&v101[57] = L"UseUnreset";
  *((_QWORD *)&v101[57] + 1) = &v33;
  *((_QWORD *)&v101[58] + 1) = &v68;
  *((_QWORD *)&v101[60] + 1) = L"PrivateHeapPackingThreshold";
  *(_QWORD *)&v101[61] = &v39;
  *(_QWORD *)&v101[62] = &v69;
  *(_QWORD *)&v101[64] = L"PrivateHeapPackingBlockSize";
  *((_QWORD *)&v101[64] + 1) = &v40;
  *((_QWORD *)&v101[65] + 1) = &v70;
  *((_QWORD *)&v101[67] + 1) = L"EvictTemporaryPeriod";
  *(_QWORD *)&v101[68] = &v41;
  *(_QWORD *)&v101[69] = &v71;
  *(_QWORD *)&v101[71] = L"EvictUnusedPeriod";
  *((_QWORD *)&v101[71] + 1) = &v42;
  *((_QWORD *)&v101[72] + 1) = &v72;
  *((_QWORD *)&v101[74] + 1) = L"ProcessPendingOfferPeriod";
  DWORD2(v101[48]) = 4;
  *(_QWORD *)&v101[49] = 0LL;
  DWORD2(v101[49]) = 288;
  LODWORD(v101[51]) = 67108868;
  LODWORD(v101[52]) = 4;
  *((_QWORD *)&v101[52] + 1) = 0LL;
  LODWORD(v101[53]) = 288;
  DWORD2(v101[54]) = 67108868;
  DWORD2(v101[55]) = 4;
  *(_QWORD *)&v101[56] = 0LL;
  DWORD2(v101[56]) = 288;
  LODWORD(v101[58]) = 67108868;
  LODWORD(v101[59]) = 4;
  *((_QWORD *)&v101[59] + 1) = 0LL;
  LODWORD(v101[60]) = 288;
  DWORD2(v101[61]) = 67108868;
  DWORD2(v101[62]) = 4;
  *(_QWORD *)&v101[63] = 0LL;
  DWORD2(v101[63]) = 288;
  LODWORD(v101[65]) = 67108868;
  LODWORD(v101[66]) = 4;
  *((_QWORD *)&v101[66] + 1) = 0LL;
  LODWORD(v101[67]) = 288;
  DWORD2(v101[68]) = 67108868;
  DWORD2(v101[69]) = 4;
  *(_QWORD *)&v101[70] = 0LL;
  DWORD2(v101[70]) = 288;
  LODWORD(v101[72]) = 67108868;
  LODWORD(v101[73]) = 4;
  *((_QWORD *)&v101[73] + 1) = 0LL;
  LODWORD(v101[74]) = 288;
  *(_QWORD *)&v101[75] = &v43;
  *(_QWORD *)&v101[76] = &v73;
  *(_QWORD *)&v101[78] = L"ProcessSysmemOfferPeriod";
  *((_QWORD *)&v101[78] + 1) = &v44;
  *((_QWORD *)&v101[79] + 1) = &v74;
  *((_QWORD *)&v101[81] + 1) = L"SegmentBalancingPolicy";
  *(_QWORD *)&v101[82] = &v45;
  *(_QWORD *)&v101[83] = &v75;
  *(_QWORD *)&v101[85] = L"BugcheckOnApertureCorruption";
  *((_QWORD *)&v101[85] + 1) = &v46;
  *((_QWORD *)&v101[86] + 1) = &v76;
  *((_QWORD *)&v101[88] + 1) = L"QuickApertureCorruptionCheck";
  *(_QWORD *)&v101[89] = &v47;
  *(_QWORD *)&v101[90] = &v77;
  *(_QWORD *)&v101[92] = L"DirectFlipMemoryRequirement";
  *((_QWORD *)&v101[92] + 1) = &v48;
  *((_QWORD *)&v101[93] + 1) = &v78;
  *((_QWORD *)&v101[95] + 1) = L"CommitProcessHeapOnDemand";
  *(_QWORD *)&v101[96] = &v49;
  *(_QWORD *)&v101[97] = &v79;
  *(_QWORD *)&v101[99] = L"SegmentCleanupSizeThreshold";
  *((_QWORD *)&v101[99] + 1) = &v50;
  *((_QWORD *)&v101[100] + 1) = &v80;
  *((_QWORD *)&v101[102] + 1) = L"SegmentCleanupCountThreshold";
  *(_QWORD *)&v101[103] = &v51;
  DWORD2(v101[75]) = 67108868;
  DWORD2(v101[76]) = 4;
  *(_QWORD *)&v101[77] = 0LL;
  DWORD2(v101[77]) = 288;
  LODWORD(v101[79]) = 67108868;
  LODWORD(v101[80]) = 4;
  *((_QWORD *)&v101[80] + 1) = 0LL;
  LODWORD(v101[81]) = 288;
  DWORD2(v101[82]) = 67108868;
  DWORD2(v101[83]) = 4;
  *(_QWORD *)&v101[84] = 0LL;
  DWORD2(v101[84]) = 288;
  LODWORD(v101[86]) = 67108868;
  LODWORD(v101[87]) = 4;
  *((_QWORD *)&v101[87] + 1) = 0LL;
  LODWORD(v101[88]) = 288;
  DWORD2(v101[89]) = 67108868;
  DWORD2(v101[90]) = 4;
  *(_QWORD *)&v101[91] = 0LL;
  DWORD2(v101[91]) = 288;
  LODWORD(v101[93]) = 67108868;
  LODWORD(v101[94]) = 4;
  *((_QWORD *)&v101[94] + 1) = 0LL;
  LODWORD(v101[95]) = 288;
  DWORD2(v101[96]) = 67108868;
  DWORD2(v101[97]) = 4;
  *(_QWORD *)&v101[98] = 0LL;
  DWORD2(v101[98]) = 288;
  LODWORD(v101[100]) = 67108868;
  LODWORD(v101[101]) = 4;
  *((_QWORD *)&v101[101] + 1) = 0LL;
  LODWORD(v101[102]) = 288;
  DWORD2(v101[103]) = 67108868;
  DWORD2(v101[105]) = 288;
  LODWORD(v101[107]) = 67108868;
  *(_QWORD *)&v101[104] = &v81;
  *(_QWORD *)&v101[106] = L"SegmentCleanupTime";
  *((_QWORD *)&v101[106] + 1) = &v52;
  *((_QWORD *)&v101[107] + 1) = &v82;
  *((_QWORD *)&v101[109] + 1) = L"PhysicalHeapSize";
  *(_QWORD *)&v101[110] = &v95;
  *(_QWORD *)&v101[111] = &v87;
  *(_QWORD *)&v101[113] = L"PhysicalHeapLowestAddress";
  *((_QWORD *)&v101[113] + 1) = &v97;
  *((_QWORD *)&v101[114] + 1) = &v88;
  *((_QWORD *)&v101[116] + 1) = L"PhysicalHeapHighestAddress";
  *(_QWORD *)&v101[117] = &v99;
  *(_QWORD *)&v101[118] = &v89;
  *(_QWORD *)&v101[120] = L"SelfRefreshVramForceEvictionTimer";
  *((_QWORD *)&v101[120] + 1) = &v53;
  *((_QWORD *)&v101[121] + 1) = &v83;
  LODWORD(v101[109]) = 288;
  DWORD2(v101[110]) = 184549387;
  DWORD2(v101[111]) = 8;
  DWORD2(v101[112]) = 288;
  LODWORD(v101[114]) = 184549387;
  LODWORD(v101[115]) = 8;
  LODWORD(v101[116]) = 288;
  DWORD2(v101[117]) = 184549387;
  DWORD2(v101[118]) = 8;
  DWORD2(v101[119]) = 288;
  LODWORD(v101[121]) = 67108868;
  *((_QWORD *)&v101[125] + 1) = 0LL;
  DWORD2(v101[104]) = 4;
  *(_QWORD *)&v101[105] = 0LL;
  LODWORD(v101[108]) = 4;
  *((_QWORD *)&v101[108] + 1) = 0LL;
  *(_QWORD *)&v101[112] = 0LL;
  *((_QWORD *)&v101[115] + 1) = 0LL;
  *(_QWORD *)&v101[119] = 0LL;
  LODWORD(v101[122]) = 4;
  *(_OWORD *)((char *)&v101[122] + 8) = 0LL;
  *(_OWORD *)((char *)&v101[123] + 8) = 0LL;
  *(_OWORD *)((char *)&v101[124] + 8) = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", v101, 0LL, 0LL);
  v9 = (unsigned int *)&unk_1C00762B8;
  memset(&unk_1C00762B8, 0, 0x80uLL);
  memset(v101, 0, sizeof(v101));
  v10 = 0;
  v11 = 0LL;
  do
  {
    memset(v103, 0, sizeof(v103));
    *(_QWORD *)&Destination.Length = 0x400000LL;
    Destination.Buffer = (PWSTR)v103;
    DestinationString = 0LL;
    String = 0LL;
    RtlInitUnicodeString(&DestinationString, L"MaxSegmentSize");
    if ( RtlAppendUnicodeStringToString(&Destination, &DestinationString) >= 0 )
    {
      *(_DWORD *)&String.Length = 0x100000;
      String.Buffer = (PWSTR)&v102;
      if ( RtlIntegerToUnicodeString(v10, 0, &String) >= 0 && RtlAppendUnicodeStringToString(&Destination, &String) >= 0 )
      {
        *(_QWORD *)&v91 = Destination.Buffer;
        *(_QWORD *)&v90 = 0LL;
        *((_QWORD *)&v91 + 1) = &v22;
        *((_QWORD *)&v90 + 1) = 288LL;
        *((_QWORD *)&v92 + 1) = &v84;
        v101[1] = v91;
        v101[0] = v90;
        *(_QWORD *)&v92 = 67108868LL;
        v93 = 4LL;
        v101[2] = v92;
        *(_QWORD *)&v101[3] = 4LL;
        v84 = 0;
        v22 = 0;
        RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", v101, 0LL, 0LL);
        v12 = v22;
        *v9 = v22;
        if ( v12 )
        {
          v18 = (v12 + 4095) & 0xFFFFF000;
          if ( v18 < 0x800000 )
            v18 = 0x800000;
          *v9 = v18;
          WdLogSingleEntry2(4LL, v11);
        }
      }
    }
    ++v10;
    ++v11;
    ++v9;
  }
  while ( v10 < 0x20 );
  WdLogSingleEntry1(4LL, v19);
  v13 = v23;
  if ( v19 < 0x5A )
    v13 = v19;
  dword_1C00762A8 = v13;
  v14 = v24;
  if ( v20 < 0x5A )
    v14 = v20;
  dword_1C00762AC = v14;
  v15 = 0x7FFFFFF;
  dword_1C00762B0 = 0;
  if ( v21 < 0x7FFFFFF )
    v15 = v21;
  dword_1C00762B4 = v15;
  dword_1C00763B8 = v26;
  dword_1C00763BC = v27;
  dword_1C00763C0 = v28;
  dword_1C00763C4 = v29;
  dword_1C00763C8 = v30;
  dword_1C00763CC = v31;
  qword_1C00763D0 = (unsigned int)(v35 << 20);
  dword_1C00763D8 = v36;
  dword_1C00763DC = v37;
  dword_1C00763E8 = v39;
  dword_1C00763EC = v40;
  dword_1C0076410 = v45;
  qword_1C00763E0 = (unsigned int)(10000000 * v38);
  qword_1C00763F0 = (unsigned int)(10000000 * v41);
  qword_1C00763F8 = (unsigned int)(10000000 * v42);
  qword_1C0076400 = (unsigned int)(10000000 * v43);
  qword_1C0076408 = (unsigned int)(10000000 * v44);
  VIDMM_GLOBAL::_Config = (v46 != 0 ? 0x10 : 0) | (v32 != 0 ? 2 : 0) | (VIDMM_GLOBAL::_Config ^ (VIDMM_GLOBAL::_Config ^ v25) & 1) & 0xFFFFFFE1 | (4 * (v34 & 1 | (unsigned __int8)(2 * (v33 & 1)))) & 0xEF;
  dword_1C0076274 = v47 != 0;
  qword_1C0076280 = (unsigned __int64)v48 << 20;
  LOBYTE(v0) = v49 != 0;
  qword_1C0076440 = (unsigned int)(v50 << 10);
  dword_1C0076448 = v51;
  NumberOfBytes = v96;
  LowestAcceptableAddress = v98;
  HighestAcceptableAddress = v100;
  qword_1C0076450 = (unsigned int)(10000 * v52);
  dword_1C0076278 = v0;
  qword_1C0076600 = 10000000LL * v53;
  VIDMM_GLOBAL::ReadCommitLimitInformation();
  VIDMM_GLOBAL::ReadWorkingSetConfiguration();
  VIDMM_GLOBAL::ReadUnusedAllocationConfiguration();
  VIDMM_GLOBAL::ReadPreparationPeriodConfiguration();
  VIDMM_GLOBAL::ReadHeapConfiguration();
  VIDMM_GLOBAL::ReadPowerConfiguration();
  VIDMM_GLOBAL::ReadGpuVaPagingHistoryConfiguration();
  VIDMM_GLOBAL::ReadGpuVaConfiguration();
  VIDMM_GLOBAL::ReadPagingConfiguration();
  VIDMM_GLOBAL::ReadTestAndStagingConfiguration();
  VIDMM_GLOBAL::ReadVPRConfiguration();
  VIDMM_GLOBAL::ReadBudgetConfiguration();
}