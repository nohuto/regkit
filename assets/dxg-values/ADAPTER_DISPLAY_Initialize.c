__int64 __fastcall ADAPTER_DISPLAY::Initialize(ADAPTER_DISPLAY *this)
{
  int *v1; // rdi
  __int64 v3; // rcx
  __int64 v4; // rdx
  unsigned int v5; // eax
  __int64 v6; // rbx
  __int64 v7; // rax
  unsigned __int64 v8; // kr00_8
  bool v9; // cf
  __int64 v10; // rax
  _QWORD *v11; // rax
  __int64 v12; // rcx
  _QWORD *v13; // rdi
  DISPLAY_SOURCE *i; // r14
  unsigned int j; // ebx
  MONITOR_MGR **v16; // r14
  MONITOR_MGR *v17; // rax
  MONITOR_MGR *v18; // rax
  MONITOR_MGR *v19; // rdi
  unsigned int v20; // ebx
  __int64 result; // rax
  unsigned int *v22; // r15
  int RegistryValues; // eax
  int v24; // r14d
  int v25; // eax
  unsigned int v26; // eax
  int v27; // ecx
  __int64 v28; // rcx
  int v29; // edx
  __int64 v30; // rax
  __int64 v31; // rcx
  __int64 v32; // rdx
  __int64 v33; // rax
  bool v34; // al
  __int64 v35; // rcx
  __int64 v36; // rax
  __int64 v37; // rcx
  bool v38; // sf
  bool v39; // of
  __int64 v40; // rcx
  int v41; // r12d
  int v42; // ebx
  struct DXGGLOBAL *v43; // rax
  int v44; // eax
  __int64 v45; // rcx
  __int64 v46; // rax
  int v47; // ecx
  struct _LUID v48; // rcx
  __int64 v49; // rax
  DXGGLOBAL *Global; // rax
  __int64 v51; // rax
  __int64 v52; // rcx
  bool v53; // zf
  __int64 v54; // rcx
  _DWORD *v55; // rcx
  __int64 v56; // rax
  struct DXGGLOBAL *v57; // rax
  struct DXGDODPRESENT *DodPresent; // rax
  __int64 v59; // rcx
  int (__fastcall *v60)(_QWORD, __int128 *); // rax
  __int64 v61; // rcx
  _DWORD *v62; // rdx
  int v63; // eax
  __int64 v64; // rcx
  unsigned int k; // r10d
  __int64 v66; // rax
  struct _KEVENT *v67; // rax
  struct OUTPUTDUPL_MGR **v68; // [rsp+28h] [rbp-E0h]
  __int64 v69; // [rsp+28h] [rbp-E0h]
  __int64 v70; // [rsp+28h] [rbp-E0h]
  __int64 v71; // [rsp+28h] [rbp-E0h]
  __int64 v72; // [rsp+28h] [rbp-E0h]
  __int64 v73; // [rsp+28h] [rbp-E0h]
  __int64 v74; // [rsp+28h] [rbp-E0h]
  __int64 v75; // [rsp+30h] [rbp-D8h]
  __int64 v76; // [rsp+30h] [rbp-D8h]
  __int64 v77; // [rsp+30h] [rbp-D8h]
  __int64 v78; // [rsp+30h] [rbp-D8h]
  __int64 v79; // [rsp+38h] [rbp-D0h]
  __int64 v80; // [rsp+38h] [rbp-D0h]
  int v81; // [rsp+58h] [rbp-B0h] BYREF
  int v82; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v83; // [rsp+60h] [rbp-A8h] BYREF
  int v84; // [rsp+64h] [rbp-A4h] BYREF
  void *EventHandle; // [rsp+68h] [rbp-A0h] BYREF
  struct _LUID v86; // [rsp+70h] [rbp-98h] BYREF
  struct _LUID v87; // [rsp+78h] [rbp-90h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v88; // [rsp+80h] [rbp-88h] BYREF
  __int128 v89; // [rsp+B0h] [rbp-58h] BYREF
  __int64 v90; // [rsp+C0h] [rbp-48h]
  _QWORD v91[50]; // [rsp+C8h] [rbp-40h] BYREF

  v1 = (int *)((char *)this + 24);
  *((_DWORD *)this + 6) = 0;
  v3 = *((_QWORD *)this + 2);
  v4 = v3;
  if ( *(_DWORD *)(v3 + 2280) >= 0x5010u && !*(_BYTE *)(v3 + 209) && (*(_DWORD *)(v3 + 2976) & 8) == 0 )
  {
    *(_QWORD *)&v88.Type = 16LL;
    *(_QWORD *)&v88.InputDataSize = 0LL;
    *(_QWORD *)&v88.Flags.0 = 0LL;
    HIDWORD(v88.hKmdProcessHandle) = 0;
    v88.pInputData = 0LL;
    v88.pOutputData = v1;
    v88.OutputDataSize = 4;
    v44 = DXGADAPTER::DdiQueryAdapterInfo((DXGADAPTER *)v3, &v88);
    if ( v44 < 0 )
    {
      *(_QWORD *)(WdLogNewEntry5_WdTrace(v45) + 24) = v44;
      *v1 = 0;
      v46 = *((_QWORD *)this + 2);
      WdLogGlobalForLineNumber = 4707;
      if ( *(int *)(v46 + 2736) >= 8704 )
        *v1 |= 2u;
    }
    v4 = *((_QWORD *)this + 2);
    v47 = *v1;
    if ( *(int *)(v4 + 2736) >= 9472 )
    {
      if ( (v47 & 0xC) == 0xC )
      {
        WdLogSingleEntry1(2LL, this);
        WdLogGlobalForLineNumber = 4736;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Adapter 0x%I64x: Both HdrFP16ScanoutSupport and HdrARGB10ScanoutSupport can't be set to 1 at the same time",
          (__int64)this,
          0LL,
          0LL,
          0LL,
          0LL);
        return 3221225485LL;
      }
    }
    else
    {
      v47 &= 0xFFFFFFF3;
      *v1 = v47;
    }
    if ( *(int *)(v4 + 2736) < 9984 )
    {
      v47 &= ~0x10u;
      *v1 = v47;
    }
    if ( *(int *)(v4 + 2736) < 10496 || *(_QWORD *)(v4 + 832) || !*(_DWORD *)(v4 + 1856) || (v47 & 2) == 0 )
    {
      v47 &= ~0x20u;
      *v1 = v47;
    }
    if ( *(int *)(v4 + 2736) < 12288 )
    {
      v47 &= ~0x40u;
      *v1 = v47;
    }
    if ( g_bDbgForceUsb4MonitorSupport )
      *v1 = v47 | 0x40;
  }
  v5 = *(_DWORD *)(v4 + 1856);
  *((_DWORD *)this + 24) = v5;
  v6 = v5;
  v8 = v5;
  v7 = 4000LL * v5;
  if ( !is_mul_ok(v8, 0xFA0uLL) )
    v7 = -1LL;
  v9 = __CFADD__(v7, 8LL);
  v10 = v7 + 8;
  if ( v9 )
    v10 = -1LL;
  v11 = (_QWORD *)operator new[](v10, 1265072196LL, 64LL);
  if ( v11 )
  {
    *v11 = v6;
    v13 = v11 + 1;
    for ( i = (DISPLAY_SOURCE *)(v11 + 1); v6; --v6 )
    {
      DISPLAY_SOURCE::DISPLAY_SOURCE(i);
      i = (DISPLAY_SOURCE *)((char *)i + 4000);
    }
  }
  else
  {
    v13 = 0LL;
  }
  *((_QWORD *)this + 16) = v13;
  if ( !v13 )
  {
    WdLogSingleEntry3(6LL, *((unsigned int *)this + 24), *((_QWORD *)this + 2), -1073741801LL);
    v76 = *((_QWORD *)this + 2);
    v71 = *((unsigned int *)this + 24);
    WdLogGlobalForLineNumber = 4792;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate 0x%I64x of display sources for adapter 0x%I64x, returning 0x%I64x",
      v71,
      v76,
      -1073741801LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  for ( j = 0; j < *((_DWORD *)this + 24); ++j )
  {
    result = DISPLAY_SOURCE::Initialize((DISPLAY_SOURCE *)(*((_QWORD *)this + 16) + 4000LL * j), this, j);
    if ( (int)result < 0 )
      return result;
  }
  v16 = (MONITOR_MGR **)((char *)this + 112);
  *(_QWORD *)(WdLogNewEntry5_WdTrace(v12) + 24) = this;
  WdLogGlobalForLineNumber = 253;
  if ( this == (ADAPTER_DISPLAY *)-112LL )
  {
    WdLogSingleEntry2(2LL, -112LL, 0LL);
    WdLogGlobalForLineNumber = 265;
    return (unsigned int)-1073741811;
  }
  *v16 = 0LL;
  v17 = (MONITOR_MGR *)operator new(688LL, 1298626628LL);
  if ( !v17 || (v18 = MONITOR_MGR::MONITOR_MGR(v17, this), (v19 = v18) == 0LL) )
  {
    WdLogSingleEntry1(2LL, *((_QWORD *)this + 2));
    WdLogGlobalForLineNumber = 285;
    return (unsigned int)-1073741811;
  }
  v20 = MONITOR_MGR::_InitializeMonitorManager(v18);
  if ( (v20 & 0x80000000) != 0 )
  {
    MONITOR_MGR::`vector deleting destructor(v19, 1u);
    return v20;
  }
  *v16 = v19;
  result = VIDPN_MGR_CLASSFACTORY::CreateVidPnMgr(this, (struct VIDPN_MGR **)this + 13);
  if ( (int)result > -1071774937 && (unsigned int)(result + 1071774934) > 0x3FE1FCD5 )
  {
    if ( (*(_DWORD *)(*((_QWORD *)this + 2) + 444LL) & 0x100) != 0 )
    {
      v48 = (struct _LUID)*((_QWORD *)DXGGLOBAL::GetGlobal() + 123);
      v49 = *((_QWORD *)this + 2);
      v87 = v48;
      v86 = *(struct _LUID *)(v49 + 412);
      result = CreateOutputDuplManager(*((_DWORD *)this + 24), 0LL, &v87, &v86, (struct OUTPUTDUPL_MGR **)this + 15);
      if ( (int)result < 0 )
        return result;
      Global = DXGGLOBAL::GetGlobal();
      DXGGLOBAL::AddIndirectOutputDuplMgr(
        Global,
        (struct OUTPUTDUPL_MGR_INDIRECT *)((*((_QWORD *)this + 15) - 24LL) & -(__int64)(*((_QWORD *)this + 15) != 0LL)));
    }
    else
    {
      result = CreateOutputDuplManager(*((_DWORD *)this + 24), this, 0LL, 0LL, (struct OUTPUTDUPL_MGR **)this + 15);
      if ( (int)result < 0 )
        return result;
    }
    v81 = 1;
    *((_QWORD *)this + 76) = (char *)this + 600;
    *((_QWORD *)this + 75) = (char *)this + 600;
    v22 = (unsigned int *)((char *)this + 528);
    *((_DWORD *)this + 130) = 0;
    *((_DWORD *)this + 132) = 1000;
    *((_DWORD *)this + 131) = 200;
    *((_DWORD *)this + 133) = 20000000;
    *((_DWORD *)this + 134) = 0;
    memset(v91, 0, 0x188uLL);
    v91[5] = 0LL;
    LODWORD(v91[4]) = 0x4000000;
    LODWORD(v91[1]) = 288;
    v91[2] = L"ModeListCaching";
    LODWORD(v91[8]) = 288;
    v91[3] = &v81;
    LODWORD(v91[11]) = 0x4000000;
    v91[9] = L"SetTimingsFlags";
    v91[16] = L"ShortLinkTrainingTimeout";
    v91[23] = L"LongLinkTrainingTimeout";
    v91[30] = L"HPDFilterLimit";
    LODWORD(v91[15]) = 288;
    LODWORD(v91[18]) = 0x4000000;
    LODWORD(v91[22]) = 288;
    LODWORD(v91[25]) = 0x4000000;
    LODWORD(v91[29]) = 288;
    LODWORD(v91[32]) = 0x4000000;
    LODWORD(v91[36]) = 288;
    LODWORD(v91[39]) = 0x4000000;
    v91[37] = L"EnableVirtualRefreshRateOnExternalMonitor";
    LODWORD(v91[6]) = 0;
    v91[7] = 0LL;
    v91[10] = (char *)this + 520;
    v91[12] = 0LL;
    LODWORD(v91[13]) = 0;
    v91[14] = 0LL;
    v91[17] = (char *)this + 524;
    v91[19] = 0LL;
    LODWORD(v91[20]) = 0;
    v91[21] = 0LL;
    v91[24] = (char *)this + 528;
    v91[26] = 0LL;
    LODWORD(v91[27]) = 0;
    v91[28] = 0LL;
    v91[31] = (char *)this + 532;
    v91[33] = 0LL;
    LODWORD(v91[34]) = 0;
    v91[35] = 0LL;
    v91[38] = (char *)this + 536;
    v91[40] = 0LL;
    LODWORD(v91[41]) = 0;
    HIDWORD(v68) = 0;
    RegistryValues = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\DMM", v91);
    v24 = RegistryValues;
    if ( RegistryValues < 0 )
    {
      WdLogSingleEntry1(4LL, RegistryValues);
      WdLogGlobalForLineNumber = 4941;
      if ( v24 != -1073741772 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 4944;
        DxgkLogInternalTriageEvent(
          0,
          262146,
          -1,
          (unsigned int)L"Status == STATUS_OBJECT_NAME_NOT_FOUND",
          4944LL,
          0LL,
          0LL,
          0LL,
          0LL);
      }
      *((_DWORD *)this + 131) = 200;
      v25 = 1;
      v81 = 1;
      v24 = 0;
      *((_DWORD *)this + 130) = 0;
      *v22 = 1000;
    }
    else
    {
      v25 = v81;
    }
    *((_BYTE *)this + 292) = v25 == 1;
    v26 = *v22;
    if ( !*v22 || *((_DWORD *)this + 131) >= v26 || v26 >= 0x7530 )
    {
      WdLogSingleEntry3(2LL, *((unsigned int *)this + 131), *((unsigned int *)this + 131), *((_QWORD *)this + 2));
      v79 = *((_QWORD *)this + 2);
      v69 = *((unsigned int *)this + 131);
      WdLogGlobalForLineNumber = 4969;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Invalid link training timeout registry value (0x%I64x, 0x%I64x) on adapter 0x%I64x, fallback to th"
                       "e default value.",
        v69,
        v69,
        v79,
        0LL,
        0LL);
      *((_DWORD *)this + 131) = 200;
      *((_DWORD *)this + 132) = 1000;
    }
    v27 = *((_DWORD *)this + 133);
    if ( (unsigned int)(v27 - 1000000) > 0x5E69EC0 )
    {
      if ( v27 )
      {
        WdLogSingleEntry3(2LL, *((unsigned int *)this + 133), 20000000LL, *((_QWORD *)this + 2));
        v80 = *((_QWORD *)this + 2);
        v72 = *((unsigned int *)this + 133);
        WdLogGlobalForLineNumber = 4984;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Invalid hot-plug filter limit of %#x on adapter 0x%I64x.  Using default of %#x.",
          v72,
          20000000LL,
          v80,
          0LL,
          0LL);
      }
      *((_DWORD *)this + 133) = 20000000;
    }
    if ( (*((_DWORD *)this + 130) & 1) != 0 )
    {
      v51 = *((_QWORD *)this + 2);
      if ( !*(_QWORD *)(v51 + 656) )
      {
        v20 = -1073741735;
        WdLogSingleEntry3(2LL, *(int *)(v51 + 416), *(unsigned int *)(v51 + 412), -1073741735LL);
        v52 = *((_QWORD *)this + 2);
        v77 = *(unsigned int *)(v52 + 412);
        v73 = *(int *)(v52 + 416);
        WdLogGlobalForLineNumber = 5001;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Miniport driver wants t fallback to use DdiCommitVidPn but it does not supply pfnCommitVidPn on "
                         "adapter (0x%I64x%08I64x), returning 0x%I64x.",
          v73,
          v77,
          -1073741735LL,
          0LL,
          0LL);
        return v20;
      }
    }
    v28 = *((_QWORD *)this + 2);
    v29 = *(_DWORD *)(v28 + 420);
    if ( (*(_DWORD *)(v28 + 444) & 0x400) != 0 )
    {
      if ( v29 == 1297040209
        && *(_DWORD *)(*(_QWORD *)(*(_QWORD *)(*(_QWORD *)(v28 + 216) + 64LL) + 40LL) + 28LL) < 0x700Au )
      {
        *((_BYTE *)this + 289) = 1;
        v34 = 1;
      }
      else
      {
        v82 = (*((_DWORD *)this + 6) >> 1) & 1;
        memset(v91, 0, 0x188uLL);
        LODWORD(v91[1]) = 288;
        v91[2] = L"ForceEnableDWMClone";
        LODWORD(v91[4]) = 67108868;
        v91[3] = &v82;
        LODWORD(v91[6]) = 4;
        v91[5] = &v82;
        HIDWORD(v68) = 0;
        RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", v91);
        v53 = v82 == 0;
        *((_BYTE *)this + 289) = v82 != 0;
        v34 = !v53;
      }
    }
    else
    {
      if ( v29 == 1297040209 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 5058;
        DxgkLogInternalTriageEvent(
          0,
          262146,
          -1,
          (unsigned int)L"GetAdapter()->GetAdapterVendorId() != VENDOR_ID_QUALCOMM",
          5058LL,
          0LL,
          0LL,
          0LL,
          0LL);
      }
      v30 = *((_QWORD *)this + 2);
      v31 = *(unsigned int *)(v30 + 412);
      v32 = *(int *)(v30 + 416);
      if ( (*((_DWORD *)this + 6) & 2) != 0 )
      {
        v20 = -1073741735;
        WdLogSingleEntry3(2LL, v32, (unsigned int)v31, -1073741735LL);
        v33 = *((_QWORD *)this + 2);
        v75 = *(unsigned int *)(v33 + 412);
        v70 = *(int *)(v33 + 416);
        WdLogGlobalForLineNumber = 5070;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Force to stop DWM clone supported adapter (0x%I64x%08I64x) due to target ID does not support DWM"
                         " clone, returning 0x%I64x.",
          v70,
          v75,
          -1073741735LL,
          0LL,
          0LL);
        return v20;
      }
      WdLogSingleEntry2(4LL, v32, v31);
      v34 = 0;
      WdLogGlobalForLineNumber = 5078;
      *((_BYTE *)this + 289) = 0;
    }
    *((_BYTE *)this + 290) = v34;
    v35 = *((_QWORD *)this + 2);
    if ( *(int *)(v35 + 3004) < 2000 )
    {
      v54 = *(_QWORD *)(v35 + 216);
      v84 = 0;
      LODWORD(v68) = 2;
      if ( (int)DpiReadPnpRegistryValue(v54, L"EnableVirtualTopologySupport", &v84, 4LL, v68) >= 0 )
      {
        if ( v84 )
        {
          v55 = (_DWORD *)*((_QWORD *)this + 2);
          if ( (v55[111] & 0x800) == 0 )
          {
            v20 = -1073741735;
            WdLogSingleEntry3(2LL, (int)v55[104], (unsigned int)v55[103], -1073741735LL);
            v56 = *((_QWORD *)this + 2);
            v78 = *(unsigned int *)(v56 + 412);
            v74 = *(int *)(v56 + 416);
            WdLogGlobalForLineNumber = 5104;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"Force to stop adapter (0x%I64x%08I64x) due to target ID does not support reduced hash size a"
                             "nd registry requested to use virtual topologies, returning 0x%I64x.",
              v74,
              v78,
              -1073741735LL,
              0LL,
              0LL);
            return v20;
          }
          *((_BYTE *)this + 290) = 1;
          v57 = DXGGLOBAL::GetGlobal();
          DXGADAPTERSOURCEHASH::ForceReducedHashSize((struct DXGGLOBAL *)((char *)v57 + 1384));
        }
      }
    }
    v36 = *((_QWORD *)this + 2);
    if ( !*(_QWORD *)(v36 + 3128) )
    {
      DodPresent = DxgkpCreateDodPresent(this, *(_QWORD *)(v36 + 696) != 0LL);
      v59 = *((_QWORD *)this + 2);
      *((_QWORD *)this + 57) = DodPresent;
      if ( !DodPresent )
        v24 = -1073741801;
      v90 = 0LL;
      v89 = 0LL;
      v60 = *(int (__fastcall **)(_QWORD, __int128 *))(v59 + 2368);
      if ( v60 && v60(*(_QWORD *)(v59 + 2296), &v89) >= 0 )
      {
        v61 = 0LL;
        v62 = (_DWORD *)((char *)this + 432);
        do
        {
          v63 = *((unsigned __int8 *)&v89 + v61++);
          *v62++ = v63;
        }
        while ( v61 < 4 );
        *((_DWORD *)this + 113) = BYTE4(v90);
        *((_DWORD *)this + 112) = BYTE5(v90);
      }
      else
      {
        *((_DWORD *)this + 108) = 1;
      }
      v64 = *(_QWORD *)(*((_QWORD *)this + 2) + 216LL);
      if ( *(_DWORD *)(*(_QWORD *)(*(_QWORD *)(v64 + 64) + 40LL) + 28LL) >= 0x3007u )
        DpiSetSchedulerCallbackState(v64, 3LL);
    }
    if ( *((_QWORD *)this + 57) )
    {
      for ( k = 0;
            k < *((_DWORD *)this + 24);
            *(_QWORD *)(2936 * v66 + *(_QWORD *)(*((_QWORD *)this + 57) + 8LL) + 400) = *(_QWORD *)(4000 * v66
                                                                                                  + *((_QWORD *)this + 16)
                                                                                                  + 928) )
      {
        v66 = k++;
      }
    }
    v37 = *((_QWORD *)this + 2);
    LODWORD(v68) = 2;
    v39 = __OFSUB__(*(_DWORD *)(v37 + 2736), 8704);
    v38 = *(_DWORD *)(v37 + 2736) - 8704 < 0;
    v40 = *(_QWORD *)(v37 + 216);
    v41 = v38 ^ v39;
    v83 = v41;
    if ( (int)DpiReadPnpRegistryValue(v40, L"NeedToSuspendVidSchBeforeSetGammaRamp", &v83, 4LL, v68) >= 0 )
    {
      v42 = v83;
      if ( v83 != v41 )
      {
        WdLogSingleEntry2(3LL, v83, *((_QWORD *)this + 2));
        WdLogGlobalForLineNumber = 5203;
      }
    }
    else
    {
      v42 = v41;
      v83 = v41;
    }
    *((_BYTE *)this + 291) = v42 != 0;
    v43 = DXGGLOBAL::GetGlobal();
    if ( (int)DXGADAPTERSOURCEHASH::AddNewAdapterEntry(
                (struct DXGGLOBAL *)((char *)v43 + 1384),
                (const struct _LUID *)(*((_QWORD *)this + 2) + 412LL),
                *((unsigned __int8 *)this + 290)) < 0 )
    {
      WdLogSingleEntry0(1LL);
      WdLogGlobalForLineNumber = 5216;
      DxgkLogInternalTriageEvent(0, 262146, -1, (unsigned int)L"NT_SUCCESS(TmpStatus)", 5216LL, 0LL, 0LL, 0LL, 0LL);
    }
    if ( v24 >= 0 )
    {
      EventHandle = 0LL;
      v67 = IoCreateNotificationEvent(0LL, &EventHandle);
      *((_QWORD *)this + 83) = v67;
      if ( v67 )
      {
        KeClearEvent(v67);
        ObfReferenceObject(*((PVOID *)this + 83));
        ZwClose(EventHandle);
      }
      else
      {
        WdLogSingleEntry0(6LL);
        WdLogGlobalForLineNumber = 5227;
        DxgkLogInternalTriageEvent(
          0,
          262145,
          -1,
          (unsigned int)L"Failed to create adapter VidPnSourceUsedBySession event object.",
          5227LL,
          0LL,
          0LL,
          0LL,
          0LL);
        return (unsigned int)-1073741801;
      }
    }
    return (unsigned int)v24;
  }
  return result;
}
