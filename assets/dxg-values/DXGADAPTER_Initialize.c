__int64 __fastcall DXGADAPTER::Initialize(DXGADAPTER *this, PDEVICE_OBJECT DeviceObject, struct _DXGK_ADAPTER_CAPS *a3)
{
  struct _ERESOURCE *v6; // rax
  __int64 result; // rax
  NTSTATUS v8; // eax
  NTSTATUS LocallyUniqueId; // ebx
  PDEVICE_OBJECT DeviceAttachmentBaseRef; // rax
  __int64 v11; // rax
  const wchar_t *v12; // r9
  int v13; // edx
  struct _ERESOURCE *v14; // rax
  NTSTATUS v15; // eax
  int v16; // eax
  __int64 v17; // r15
  int AdapterInfo; // eax
  struct _LUID *v19; // rdx
  int (__fastcall *v20)(_QWORD, __int128 *); // rax
  unsigned __int8 v21; // bl
  DXGSESSIONDATA *SessionDataForSpecifiedSession; // rax
  unsigned int v23; // eax
  __int64 v24; // rax
  const wchar_t *v25; // r9
  unsigned int v26; // eax
  const struct _GUID *v27; // rdx
  unsigned __int16 v28; // r8
  unsigned __int16 v29; // r9
  int v30; // eax
  const wchar_t *v31; // r9
  int v32; // eax
  unsigned __int8 v33; // dl
  __int64 v34; // rcx
  __int64 v35; // r14
  __int64 v36; // rax
  unsigned int v37; // r13d
  unsigned __int8 v38; // r8
  __int64 v39; // rax
  const wchar_t *v40; // r9
  int v41; // eax
  int v42; // ecx
  __int64 v43; // rax
  int v44; // ecx
  __int64 v45; // r15
  int v46; // eax
  int v47; // ecx
  __int64 v48; // rcx
  int v49; // eax
  int v50; // ecx
  char v51; // al
  int v52; // eax
  unsigned int v53; // ebx
  __int64 v54; // rax
  __int64 v55; // rax
  char v56; // r12
  unsigned int v57; // eax
  unsigned int v58; // r8d
  __int64 v59; // r9
  UINT PhysicalAdapterCapsSizeFromDdiVersion; // r15d
  int v61; // eax
  __int64 v62; // rdx
  __int64 v63; // rcx
  __int64 v64; // rax
  DXGGLOBAL *Global; // rax
  __int64 v66; // rcx
  unsigned int v67; // r8d
  __int64 v68; // rdx
  __int64 v69; // rcx
  __int64 v70; // rdi
  __int64 v71; // rbx
  __int64 v72; // rdi
  __int64 v73; // rbx
  __int64 v74; // rax
  __int64 v75; // rax
  int v76; // eax
  __int64 RenderCore; // rdi
  unsigned int v78; // edx
  unsigned __int64 v79; // r8
  __int64 v80; // r13
  __int64 v81; // r12
  int v82; // ecx
  int v83; // edi
  int v84; // ebx
  unsigned __int8 IsGpuVaIoMmuGlobalSupported; // al
  const wchar_t *v86; // r9
  int v87; // eax
  char v88; // al
  int v89; // eax
  __int64 v90; // rax
  int v91; // eax
  int v92; // ecx
  int v93; // eax
  struct _DXGK_ADAPTER_CAPS *v94; // r12
  char v95; // cl
  char v96; // dl
  char v97; // al
  char v98; // r8
  char v99; // cl
  char v100; // dl
  char v101; // cl
  char v102; // al
  char v103; // al
  char v104; // cl
  unsigned int v105; // eax
  int v106; // ecx
  __int64 v107; // rax
  struct DXGGLOBAL *v108; // rax
  struct DXGGLOBAL *v109; // rax
  struct DXGGLOBAL *v110; // rax
  char v111; // r9
  char v112; // r8
  unsigned int v113; // ecx
  unsigned int v114; // edx
  __int64 v115; // rax
  __int64 v116; // rax
  unsigned int v117; // ebx
  DXGGLOBAL *v118; // rax
  int v119; // eax
  int v120; // ecx
  __int64 v121; // rax
  __int64 v122; // rcx
  int v123; // eax
  char v124; // cl
  int v125; // eax
  __int64 v126; // rax
  char *v127; // rbx
  int DisplayCore; // eax
  bool v129; // zf
  char v130; // cl
  char v131; // dl
  int v132; // eax
  char v133; // al
  __int64 v134; // rdx
  DXGADAPTER *v135; // rcx
  int v136; // eax
  __int64 v137; // rax
  bool v138; // cf
  __int64 v139; // rdx
  __int64 v140; // r8
  int v141; // eax
  DXGGLOBAL *v142; // rax
  int v143; // eax
  __int64 v144; // rax
  int v145; // eax
  DXGADAPTER *v146; // rcx
  __int64 v147; // r14
  __int64 v148; // rbx
  struct DXGGLOBAL *v149; // rax
  int v150; // eax
  struct DXGGLOBAL *v151; // rax
  __int64 v152; // rdx
  DXGGLOBAL *v153; // rax
  __int64 v154; // [rsp+20h] [rbp-E0h]
  void *v155; // [rsp+20h] [rbp-E0h]
  void *v156; // [rsp+20h] [rbp-E0h]
  void *v157; // [rsp+20h] [rbp-E0h]
  void *v158; // [rsp+28h] [rbp-D8h]
  __int64 v159; // [rsp+28h] [rbp-D8h]
  __int64 v160; // [rsp+28h] [rbp-D8h]
  __int64 v161; // [rsp+30h] [rbp-D0h]
  unsigned int v162; // [rsp+50h] [rbp-B0h] BYREF
  bool IsAdapterSessionized; // [rsp+54h] [rbp-ACh]
  unsigned int v164; // [rsp+58h] [rbp-A8h] BYREF
  int v165; // [rsp+5Ch] [rbp-A4h] BYREF
  int v166; // [rsp+60h] [rbp-A0h] BYREF
  int v167; // [rsp+64h] [rbp-9Ch] BYREF
  int v168; // [rsp+68h] [rbp-98h] BYREF
  unsigned int v169; // [rsp+6Ch] [rbp-94h]
  __int64 v170; // [rsp+70h] [rbp-90h] BYREF
  unsigned int v171; // [rsp+78h] [rbp-88h]
  __int64 v172; // [rsp+80h] [rbp-80h] BYREF
  _DXGKARG_QUERYADAPTERINFO v173; // [rsp+88h] [rbp-78h] BYREF
  struct _DXGK_ADAPTER_CAPS *v174[2]; // [rsp+B8h] [rbp-48h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v175; // [rsp+C8h] [rbp-38h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v176; // [rsp+F8h] [rbp-8h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v177; // [rsp+128h] [rbp+28h] BYREF
  __int128 v178; // [rsp+158h] [rbp+58h] BYREF
  unsigned int v179[2]; // [rsp+168h] [rbp+68h] BYREF

  v174[0] = a3;
  if ( KeGetCurrentIrql() )
  {
    WdLogSingleEntry0(1LL);
    WdLogGlobalForLineNumber = 6949;
    DxgkLogInternalTriageEvent(
      0,
      262146,
      -1,
      (unsigned int)L"KeGetCurrentIrql() == PASSIVE_LEVEL",
      6949LL,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  if ( *((_DWORD *)this + 50) )
    return 3221225485LL;
  v6 = (struct _ERESOURCE *)operator new(104LL, 1265072196LL);
  *((_QWORD *)this + 21) = v6;
  if ( !v6 )
  {
    WdLogSingleEntry2(3LL, this, -1073741801LL);
    WdLogGlobalForLineNumber = 6967;
    return 3221225495LL;
  }
  v8 = ExInitializeResourceLite(v6);
  LocallyUniqueId = v8;
  if ( v8 < 0 )
  {
    WdLogSingleEntry2(3LL, this, v8);
    WdLogGlobalForLineNumber = 6978;
    return (unsigned int)LocallyUniqueId;
  }
  *((_QWORD *)this + 27) = DeviceObject;
  *((_QWORD *)this + 28) = DpiGetSysMmAdapterFromDevice(DeviceObject);
  DeviceAttachmentBaseRef = IoGetDeviceAttachmentBaseRef(DeviceObject);
  *((_QWORD *)this + 29) = DeviceAttachmentBaseRef;
  ObfDereferenceObject(DeviceAttachmentBaseRef);
  LocallyUniqueId = ZwAllocateLocallyUniqueId((PLUID)((char *)this + 4756));
  if ( LocallyUniqueId < 0 )
  {
    WdLogSingleEntry0(6LL);
    v11 = 6999LL;
    v12 = L"ZwAllocateLocallyUniqueId failed";
LABEL_12:
    v13 = 262145;
LABEL_13:
    WdLogGlobalForLineNumber = v11;
    DxgkLogInternalTriageEvent(0, v13, -1, (_DWORD)v12, v11, 0LL, 0LL, 0LL, 0LL);
    return (unsigned int)LocallyUniqueId;
  }
  v14 = (struct _ERESOURCE *)operator new(104LL, 1265072196LL);
  *((_QWORD *)this + 35) = v14;
  if ( !v14 )
  {
    WdLogSingleEntry2(3LL, this, -1073741801LL);
    WdLogGlobalForLineNumber = 7012;
    return 3221225495LL;
  }
  v15 = ExInitializeResourceLite(v14);
  LocallyUniqueId = v15;
  if ( v15 < 0 )
  {
    WdLogSingleEntry2(3LL, this, v15);
    WdLogGlobalForLineNumber = 7023;
    return (unsigned int)LocallyUniqueId;
  }
  _InterlockedIncrement64((volatile signed __int64 *)this + 3);
  v170 = 0LL;
  *((_QWORD *)this + 5) = -1LL;
  if ( *((_BYTE *)DeviceObject->DeviceExtension + 481) )
  {
    v16 = DXGADAPTER::InitializeParavirtualizedAdapter(this, (struct DRIVER_WORKAROUNDS *)&v170);
    v17 = v16;
    if ( v16 < 0 )
    {
      WdLogSingleEntry1(2LL, v16);
      WdLogGlobalForLineNumber = 7046;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"InitializeParavirtualizedAdapter failed: 0x%I64x",
        v17,
        0LL,
        0LL,
        0LL,
        0LL);
      return (unsigned int)v17;
    }
  }
  else
  {
    *((_BYTE *)this + 1785) = 0;
    AdapterInfo = DpiGetAdapterInfo((int)DeviceObject, (char *)this + 1744, (char *)this + 288);
    LocallyUniqueId = AdapterInfo;
    if ( AdapterInfo < 0 )
    {
      WdLogSingleEntry2(3LL, this, AdapterInfo);
      WdLogGlobalForLineNumber = 7063;
      return (unsigned int)LocallyUniqueId;
    }
  }
  DpiFdoSetFeatureDatabaseDxgAdapter(*((_QWORD *)this + 27), this);
  *(_QWORD *)v179 = 0LL;
  v20 = (int (__fastcall *)(_QWORD, __int128 *))*((_QWORD *)this + 296);
  v178 = 0LL;
  if ( v20 && v20(*((_QWORD *)this + 287), &v178) >= 0 )
  {
    *(_QWORD *)((char *)this + 4828) = *((_QWORD *)&v178 + 1);
    *((_DWORD *)this + 1209) = v179[0];
  }
  IsAdapterSessionized = DXGADAPTER::IsAdapterSessionized(this, v19, v179, 0LL);
  v21 = IsAdapterSessionized;
  if ( IsAdapterSessionized )
  {
    SessionDataForSpecifiedSession = DXGSESSIONMGR::GetSessionDataForSpecifiedSession(
                                       *(DXGSESSIONMGR **)(*((_QWORD *)this + 2) + 944LL),
                                       v179[0]);
    if ( !SessionDataForSpecifiedSession
      || (v23 = DXGSESSIONDATA::AcquireSessionAdapterOrdinal(SessionDataForSpecifiedSession),
          *((_DWORD *)this + 61) = v23,
          v23 == -1) )
    {
      WdLogSingleEntry2(2LL, v179[0], -1073741801LL);
      v24 = v179[0];
      v25 = L"Exceeded the maximum number of sessionized adapter in session 0x%I64x, returning 0x%I64x.";
      v159 = -1073741801LL;
      WdLogGlobalForLineNumber = 7096;
LABEL_31:
      DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v25, v24, v159, 0LL, 0LL, 0LL);
      return 3221225495LL;
    }
  }
  v26 = DXGGLOBAL::AcquireAdapterOrdinal(*((DXGGLOBAL **)this + 2), v21);
  *((_DWORD *)this + 60) = v26;
  if ( v26 == -1 )
    return 3221225495LL;
  if ( (*((_DWORD *)this + 111) & 0x200) != 0 )
    *((_BYTE *)DXGGLOBAL::GetGlobal() + 304832) = 1;
  v30 = *((_DWORD *)this + 111);
  if ( (v30 & 8) != 0 && (v30 & 0x10) != 0 )
  {
    WdLogSingleEntry0(1LL);
    WdLogGlobalForLineNumber = 7120;
    DxgkLogInternalTriageEvent(
      0,
      262146,
      -1,
      (unsigned int)L"!(IsSoftGPU() && IsWarpAdapter())",
      7120LL,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  if ( !*((_QWORD *)this + 57) )
  {
    WdLogSingleEntry0(2LL);
    v31 = L"Miniport did not provide required DDIs";
    v160 = 0LL;
    v154 = 7127LL;
    WdLogGlobalForLineNumber = 7127;
LABEL_40:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v31, v154, v160, 0LL, 0LL, 0LL);
    return 3221225561LL;
  }
  if ( !*((_QWORD *)this + 74) )
    *((_QWORD *)this + 74) = DXGADAPTER::DefaultDdiEscape;
  if ( !*((_QWORD *)this + 135) )
    *((_QWORD *)this + 135) = W32kStub_GreSfmOpenTokenEvent;
  v32 = DXGADAPTER::CallDriverQueryInterface(this, v27, v28, v29, (char *)this + 2096, v158);
  v35 = v32;
  if ( v32 >= 0 )
  {
    if ( *((_WORD *)this + 1049) >= 4u )
      goto LABEL_49;
  }
  else
  {
    v36 = WdLogNewEntry5_WdTrace(v34);
    *(_QWORD *)(v36 + 24) = this;
    *(_QWORD *)(v36 + 32) = v35;
    WdLogGlobalForLineNumber = 7158;
  }
  memset((char *)this + 2096, 0, 0xB8uLL);
LABEL_49:
  v37 = *(_DWORD *)(*(_QWORD *)(*(_QWORD *)(*((_QWORD *)this + 27) + 64LL) + 40LL) + 28LL);
  v171 = v37;
  *((_DWORD *)this + 570) = v37;
  if ( v37 < 0x7000 )
  {
    if ( v37 < 0x6002 )
      goto LABEL_58;
  }
  else
  {
    if ( !*((_DWORD *)this + 464) )
      goto LABEL_58;
    if ( *((_DWORD *)this + 465) )
    {
      v38 = 0;
LABEL_57:
      DXGADAPTER::SetModeBehavior(this, v33, v38);
      goto LABEL_58;
    }
  }
  if ( *((_DWORD *)this + 464) && *((_DWORD *)this + 465) )
  {
    v38 = 1;
    goto LABEL_57;
  }
LABEL_58:
  if ( v37 - 20480 <= 5 )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7202LL;
    v40 = L"Cannot load an M1 threshold driver on later builds.";
LABEL_60:
    WdLogGlobalForLineNumber = v39;
LABEL_61:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v40, v39, 0LL, 0LL, 0LL, 0LL);
    return 3221225485LL;
  }
  *(_QWORD *)&v173.InputDataSize = 0LL;
  v173.pOutputData = (char *)this + 2400;
  *(_QWORD *)&v173.Type = 1LL;
  *(_QWORD *)&v173.Flags.0 = 0LL;
  HIDWORD(v173.hKmdProcessHandle) = 0;
  v173.pInputData = 0LL;
  v173.OutputDataSize = GetDriverCapsSizeFromDdiVersion(v37);
  if ( !v173.OutputDataSize )
    return 3221225485LL;
  v41 = DXGADAPTER::DdiQueryAdapterInfo(this, &v173);
  v17 = v41;
  if ( v41 < 0 )
  {
    WdLogSingleEntry1(2LL, v41);
    WdLogGlobalForLineNumber = 7225;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Miniport failed DdiQueryAdapterInfo(DXGKQAITYPE_DRIVERCAPS) with status 0x%I64x",
      v17,
      0LL,
      0LL,
      0LL,
      0LL);
    return (unsigned int)v17;
  }
  v42 = *((_DWORD *)this + 684);
  if ( v42 <= 9472 )
  {
    if ( v42 < 4864 )
    {
      v45 = 0LL;
      if ( *((_QWORD *)this + 104) )
      {
        v44 = 1300;
      }
      else if ( v42 == 4608 )
      {
        v44 = 1200;
      }
      else if ( !*((_QWORD *)this + 100) || (v44 = 1105, (*((_DWORD *)this + 613) & 4) == 0) )
      {
        v44 = 1000;
      }
      *((_DWORD *)this + 751) = v44;
      goto LABEL_79;
    }
  }
  else if ( *((_DWORD *)DeviceObject->DeviceExtension + 687) <= 0xA00Bu )
  {
    WdLogSingleEntry1(2LL, *((int *)this + 684));
    v43 = *((int *)this + 684);
    WdLogGlobalForLineNumber = 7231;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Miniport returned incorrect WDDMVersion: 0x%I64x",
      v43,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225485LL;
  }
  v44 = DxgkConvertWddmVersionToD3DKMTDriverVersion();
  *((_DWORD *)this + 751) = v44;
  v45 = 0LL;
LABEL_79:
  v46 = *((_DWORD *)this + 744);
  if ( v44 >= 2600 )
  {
    v47 = *((_DWORD *)this + 111);
    if ( (v46 & 8) != 0 )
    {
      *((_DWORD *)this + 111) = v47 | 0x80000;
    }
    else if ( (v47 & 0x80000) != 0 && v37 >= 0x11002 )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7285LL;
      v40 = L"MiscCaps.ComputeOnly is not set, but the device belongs to the ComputeAccelerator class";
      goto LABEL_60;
    }
  }
  else
  {
    v46 &= ~8u;
    *((_DWORD *)this + 744) = v46;
  }
  if ( *((_BYTE *)this + 1784) && (v46 & 0xC) == 0 )
  {
    WdLogSingleEntry0(2LL);
    WdLogGlobalForLineNumber = 7292;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"UMD name is missing and device is not compute only",
      7292LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225524LL;
  }
  v48 = *((_QWORD *)this + 27);
  v165 = 0;
  LODWORD(v155) = 2;
  v49 = DpiReadPnpRegistryValue(v48, L"ACGSupported", &v165, 4LL, v155);
  v50 = v165;
  if ( v49 < 0 )
    v50 = 0;
  v165 = v50;
  if ( v50 || (v51 = 0, *((int *)this + 751) >= 2200) )
    v51 = 1;
  *((_BYTE *)this + 212) = v51;
  if ( *((_BYTE *)this + 209) )
  {
    *((_BYTE *)a3 + 1) &= ~1u;
    *(_BYTE *)a3 &= 0x7Bu;
    *((_DWORD *)this + 744) &= 0xFFFFFFEB;
    *((_DWORD *)this + 617) &= 0xFFFFD2FF;
    *((_BYTE *)this + 2940) = 0;
    *((_BYTE *)this + 2968) = 1;
    *((_BYTE *)this + 2942) = 1;
    if ( *((_BYTE *)this + 210) )
      *((_DWORD *)this + 613) &= ~0x100000u;
  }
  else if ( v37 >= 0x5023 )
  {
    if ( g_bCreateParavirtualizedGpu )
    {
      v52 = *((_DWORD *)this + 111);
      if ( (v52 & 4) == 0 && (v52 & 0x10) == 0 && !*(_BYTE *)(*((_QWORD *)DeviceObject->DeviceExtension + 5) + 133LL) )
        *((_DWORD *)this + 617) |= 0x400u;
    }
  }
  v169 = *((_DWORD *)this + 74);
  v53 = v169;
  v54 = 344LL * v169;
  if ( !is_mul_ok(v169, 0x158uLL) )
    v54 = -1LL;
  v55 = operator new[](v54, 1265072196LL, 64LL);
  *((_QWORD *)this + 374) = v55;
  if ( !v55 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 7351;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate DXGK_PHYSICALADAPTERINFO",
      7351LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v56 = 0;
  if ( *((int *)this + 684) < 0x2000 || v37 < 0x5005 )
    goto LABEL_147;
  *((_DWORD *)this + 750) = 0;
  v57 = 0;
  v162 = 0;
  if ( v53 )
  {
    PhysicalAdapterCapsSizeFromDdiVersion = GetPhysicalAdapterCapsSizeFromDdiVersion(v37);
    while ( 1 )
    {
      v175.pInputData = &v162;
      *(_QWORD *)&v175.Type = 15LL;
      *(_QWORD *)&v175.InputDataSize = 4LL;
      v175.pOutputData = (void *)(v59 + 344LL * v58);
      *(_QWORD *)&v175.Flags.0 = 0LL;
      HIDWORD(v175.hKmdProcessHandle) = 0;
      v175.OutputDataSize = PhysicalAdapterCapsSizeFromDdiVersion;
      v61 = DXGADAPTER::DdiQueryAdapterInfo(this, &v175);
      if ( v61 < 0 )
        break;
      if ( v37 >= 0xC003 )
      {
        v62 = *((_QWORD *)this + 374);
        v63 = 344LL * v162;
        if ( (*(_DWORD *)(v63 + v62 + 16) & 0x20) != 0 )
        {
          v64 = *(unsigned int *)(v63 + v62 + 24);
          if ( (unsigned int)v64 >= *(unsigned __int16 *)(v63 + v62) )
          {
            WdLogSingleEntry3(2LL, this, v64, *(unsigned __int16 *)(v63 + v62));
            v75 = *((_QWORD *)this + 374);
            WdLogGlobalForLineNumber = 7396;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"Adapter 0x%I64x: VirtualCopyEngineSupported but node index is invalid (VirtualCopyIndex:%u, "
                             "NumExecutionNodes:%u)",
              (__int64)this,
              *(unsigned int *)(344LL * v162 + v75 + 24),
              *(unsigned __int16 *)(344LL * v162 + v75),
              0LL,
              0LL);
            return 3221225485LL;
          }
          if ( (*((_DWORD *)this + 617) & 0x2000) == 0 )
          {
            WdLogSingleEntry1(2LL, this);
            WdLogGlobalForLineNumber = 7403;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"Adapter 0x%I64x: IoMmuSecureModeRequired must be set for a device exposing a virtual copy engine",
              (__int64)this,
              0LL,
              0LL,
              0LL,
              0LL);
            return 3221225485LL;
          }
        }
      }
      Global = DXGGLOBAL::GetGlobal();
      if ( DXGGLOBAL::GpuVaIoMmuEnabled(Global) )
      {
        v66 = *((_QWORD *)this + 27);
        v166 = 0;
        v167 = 0;
        LODWORD(v156) = 2;
        if ( (int)DpiReadPnpRegistryValue(v66, L"DxgkGpuVaIommuRequired", &v166, 4LL, v156) >= 0 )
          *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) = (v166 != 0 ? 0x40 : 0) | *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) & 0xFFFFFFBF;
        LODWORD(v157) = 2;
        if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)this + 27), L"DxgkGpuVaIommuGlobalSupported", &v167, 4LL, v157) >= 0 )
          *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) = (v167 != 0 ? 0x80 : 0) | *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) & 0xFFFFFF7F;
      }
      v67 = v162;
      v68 = *((_QWORD *)this + 374);
      v69 = 344LL * v162;
      if ( (*(_DWORD *)(v69 + v68 + 16) & 2) != 0 )
      {
        *(_BYTE *)(v69 + v68 + 49) = 1;
        v67 = v162;
      }
      v70 = *((_QWORD *)this + 374);
      v71 = 344LL * v67;
      if ( (*(_DWORD *)(v71 + v70 + 16) & 0x40) != 0 )
      {
        if ( !DXGADAPTER::IsGpuVaIoMmuSupported(this) )
        {
          WdLogSingleEntry1(2LL, this);
          WdLogGlobalForLineNumber = 7434;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Adapter 0x%I64x: GpuVaIommuRequired is set for a physical adapter, but not in IOMMU_CAPS",
            (__int64)this,
            0LL,
            0LL,
            0LL,
            0LL);
          return 3221225485LL;
        }
        *(_BYTE *)(v71 + v70 + 49) = 1;
        *(_BYTE *)(344LL * v162 + *((_QWORD *)this + 374) + 48) = 1;
        v67 = v162;
      }
      v72 = *((_QWORD *)this + 374);
      v73 = 344LL * v67;
      if ( (*(_DWORD *)(v73 + v72 + 16) & 0x80u) != 0 )
      {
        if ( !DXGADAPTER::IsGpuVaIoMmuGlobalSupported(this) )
        {
          WdLogSingleEntry1(2LL, this);
          WdLogGlobalForLineNumber = 7445;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Adapter 0x%I64x: GpuVaIommuGlobalRequired is set for a physical adapter, but not in IOMMU_CAPS",
            (__int64)this,
            0LL,
            0LL,
            0LL,
            0LL);
          return 3221225485LL;
        }
        *(_BYTE *)(v73 + v72 + 49) = 1;
        *(_BYTE *)(344LL * v162 + *((_QWORD *)this + 374) + 48) = 1;
        v67 = v162;
      }
      v59 = *((_QWORD *)this + 374);
      v53 = v169;
      v74 = v67;
      v58 = v67 + 1;
      v57 = *(unsigned __int16 *)(344 * v74 + v59) + *((_DWORD *)this + 750);
      v162 = v58;
      *((_DWORD *)this + 750) = v57;
      if ( v58 >= v53 )
        goto LABEL_130;
    }
    WdLogSingleEntry1(4LL, v61);
    WdLogGlobalForLineNumber = 7378;
    v56 = 1;
  }
  else
  {
LABEL_130:
    if ( *((int *)this + 751) <= 2400 && v57 > 0x40 )
    {
      WdLogSingleEntry3(2LL, this, 64LL, v57);
      v161 = *((unsigned int *)this + 750);
      WdLogGlobalForLineNumber = 7463;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Adapter 0x%I64x: Exceeded maximum number of %I64d nodes on pre-WDDM 2.5 adapter. Total node count: %I64d",
        (__int64)this,
        64LL,
        v161,
        0LL,
        0LL);
      return 3221225485LL;
    }
    if ( (*((_DWORD *)this + 616) & 1) == 0 )
    {
      WdLogSingleEntry1(2LL, this);
      WdLogGlobalForLineNumber = 7468;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Adapter 0x%I64x: SchedulingCaps.MultiEngineAware is not set by WDDMv2 driver",
        (__int64)this,
        0LL,
        0LL,
        0LL,
        0LL);
      return 3221225485LL;
    }
  }
  v45 = 0LL;
  if ( (*((_DWORD *)this + 617) & 0x800) != 0 )
  {
    v164 = 0;
    if ( v53 )
    {
      while ( 1 )
      {
        v172 = 0LL;
        v173.pInputData = &v164;
        v173.Type = DXGKQAITYPE_FRAMEBUFFERSAVESIZE;
        v173.pOutputData = &v172;
        v173.InputDataSize = 4;
        v173.OutputDataSize = 8;
        v76 = DXGADAPTER::DdiQueryAdapterInfo(this, &v173);
        RenderCore = v76;
        if ( v76 < 0 )
          break;
        if ( (v172 & 0xFFF) != 0 )
        {
          WdLogSingleEntry1(2LL, v172);
          v39 = v172;
          v40 = L"Frame buffer reserve size must be a multiple of PAGE_SIZE. Size=%I64u";
          WdLogGlobalForLineNumber = 7493;
          goto LABEL_61;
        }
        *(_QWORD *)(344LL * v164 + *((_QWORD *)this + 374) + 56) = v172;
        v78 = v164;
        v79 = *(_QWORD *)(344LL * v164 + *((_QWORD *)this + 374) + 56);
        if ( v79 )
        {
          result = DXGADAPTER::CreateFrameBufferSaveAreaSection(this, v164, v79);
          if ( (int)result < 0 )
            return result;
          v78 = v164;
        }
        v164 = v78 + 1;
        if ( v78 + 1 >= v53 )
          goto LABEL_146;
      }
      WdLogSingleEntry1(2LL, v76);
      v86 = L"Failed to query frame buffer save area size. Status 0x%I64x";
      WdLogGlobalForLineNumber = 7487;
      goto LABEL_160;
    }
  }
LABEL_146:
  if ( v56 )
  {
LABEL_147:
    if ( v53 )
    {
      v80 = v53;
      do
      {
        v81 = *((_QWORD *)this + 374);
        *(_WORD *)(v45 + v81) = *((_WORD *)this + 1238);
        v82 = *(_DWORD *)(v45 + v81 + 16) ^ ((unsigned __int8)*(_DWORD *)(v45 + v81 + 16) ^ (unsigned __int8)(*((_DWORD *)this + 617) >> 7)) & 1;
        *(_DWORD *)(v45 + v81 + 16) = v82;
        v83 = v82 ^ (v82 ^ (*((_DWORD *)this + 617) >> 5)) & 2;
        *(_DWORD *)(v45 + v81 + 16) = v83;
        v84 = v83 ^ ((unsigned __int8)v83 ^ (unsigned __int8)(DXGADAPTER::IsGpuVaIoMmuSupported(this) << 6)) & 0x40;
        *(_DWORD *)(v45 + v81 + 16) = v84;
        IsGpuVaIoMmuGlobalSupported = DXGADAPTER::IsGpuVaIoMmuGlobalSupported(this);
        *(_DWORD *)(v45 + v81 + 16) = v84 ^ ((unsigned __int8)v84 ^ (unsigned __int8)(IsGpuVaIoMmuGlobalSupported << 7)) & 0x80;
        *(_WORD *)(v45 + v81 + 2) = *((_WORD *)this + 1236);
        *(_QWORD *)(v45 + v81 + 8) = *((_QWORD *)this + 27);
        if ( (((unsigned __int8)v84 ^ ((unsigned __int8)v84 ^ (unsigned __int8)(IsGpuVaIoMmuGlobalSupported << 7)) & 0x80) & 0xC2) != 0 )
          *(_WORD *)(v45 + v81 + 48) = 257;
        v45 += 344LL;
        --v80;
      }
      while ( v80 );
      v37 = v171;
    }
  }
  if ( *((int *)this + 751) >= 2400 )
  {
    if ( *((_DWORD *)this + 744) >= 0x200u )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7547LL;
      v40 = L"Driver should not set MiscCaps.Reserved bits";
      goto LABEL_60;
    }
    *((_BYTE *)this + 3057) = *((_BYTE *)this + 2976) & 1;
  }
  v87 = *((_DWORD *)this + 744);
  if ( (v87 & 0x10) != 0 && !*((_QWORD *)this + 175) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7557LL;
    v40 = L"Driver sets IndependentVidPnVSyncControl cap but does not support DxgkDdiControlInterrupt3, returning failure";
    goto LABEL_60;
  }
  if ( *((_BYTE *)this + 3220) )
    *((_DWORD *)this + 744) = v87 & 0xFFFFFFEF;
  if ( v37 >= 0x3001 )
  {
    v89 = *((_DWORD *)this + 684);
    if ( v89 != 4096
      && v89 != 4608
      && v89 != 4864
      && v89 != 0x2000
      && v89 != 8448
      && v89 != 8704
      && v89 != 8960
      && v89 != 9216
      && v89 != 9472
      && v89 != 9728
      && v89 != 9984
      && v89 != 10240
      && v89 != 10496
      && v89 != 12288
      && v89 != 12544
      && v89 != 12800 )
    {
      WdLogSingleEntry1(2LL, *((int *)this + 684));
      v90 = *((int *)this + 684);
      v31 = L"Miniport returned unknown WDDM version 0x%I64x";
      v160 = 0LL;
      WdLogGlobalForLineNumber = 7615;
LABEL_203:
      v154 = v90;
      goto LABEL_40;
    }
  }
  else
  {
    *((_DWORD *)this + 684) = 4096;
  }
  if ( !*((_BYTE *)DXGGLOBAL::GetGlobal() + 888) || (v88 = 1, (*((_DWORD *)this + 111) & 8) != 0) )
    v88 = 0;
  *((_BYTE *)this + 3016) = v88;
  if ( v88 )
  {
    if ( *((int *)this + 684) < 4608
      && (*((_DWORD *)this + 732)
       || *((_DWORD *)this + 733)
       || *((_BYTE *)this + 2936)
       || *((_BYTE *)this + 2937)
       || *((_BYTE *)this + 2938)
       || (*((_DWORD *)this + 613) & 0x10000000) != 0
       || (*((_DWORD *)this + 616) & 0x14) != 0
       || *((_BYTE *)this + 2939)
       || *((_BYTE *)this + 2941)
       || *((_BYTE *)this + 2942)) )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7641LL;
      v40 = L"Driver reports WDDM version less than 1.2 but implements some WDDM 1.2 features.";
      goto LABEL_60;
    }
    v91 = *((_DWORD *)this + 684);
    if ( v91 >= 4864 )
    {
      if ( v91 >= 0x2000 )
        goto LABEL_213;
    }
    else if ( (*((_DWORD *)this + 615) & 0x10) != 0
           || (*((_DWORD *)this + 617) & 0x10) != 0
           || *((_BYTE *)this + 2943)
           || *((_DWORD *)this + 736) )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7656LL;
      v40 = L"Driver reports WDDM version less than 1.3 but implements some WDDM 1.3 features.";
      goto LABEL_60;
    }
    if ( *((_BYTE *)this + 2948) )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7684LL;
      v40 = L"Pre-WDDM 2.0 driver should not set the HybridIntegrated cap.";
      goto LABEL_60;
    }
  }
LABEL_213:
  v92 = *((_DWORD *)this + 617);
  if ( (v92 & 0x10000) != 0 )
  {
    if ( (*((_DWORD *)this + 617) & 0x8010) != 0x8010 )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7698LL;
      v40 = L"Driver reports CrossAdapterResourceScanout but does not report lower tier support.";
      goto LABEL_60;
    }
  }
  else if ( (v92 & 0x8000) != 0 && (v92 & 0x10) == 0 )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7706LL;
    v40 = L"Driver reports CrossAdapterResourceTexture but does not report lower tier support.";
    goto LABEL_60;
  }
  if ( v37 >= 0x4000 )
  {
    if ( v37 >= 0x5011 )
      goto LABEL_226;
  }
  else
  {
    v92 &= ~0x10u;
    *((_BYTE *)this + 2943) = 0;
    *((_DWORD *)this + 617) = v92;
  }
  v93 = *((_DWORD *)this + 111);
  if ( (v93 & 1) != 0 && (v92 & 0x10) != 0 && (v93 & 0x1000) != 0 )
    *((_BYTE *)this + 2948) = 1;
LABEL_226:
  v94 = v174[0];
  v95 = *(_BYTE *)v174[0] ^ (*(_BYTE *)v174[0] ^ (4 * *((_BYTE *)this + 2936))) & 4;
  *(_BYTE *)v174[0] = v95;
  v96 = v95 & 0xF7 | (*((_BYTE *)this + 2942) != 0 ? 8 : 0);
  *(_BYTE *)v94 = v96;
  v97 = v96 ^ (v96 ^ (32 * (*((_DWORD *)this + 617) >> 4))) & 0x20;
  *(_BYTE *)v94 = v97;
  v98 = v97 ^ (v97 ^ (*((_BYTE *)this + 2943) << 6)) & 0x40;
  *(_BYTE *)v94 = v98;
  *((_DWORD *)v94 + 1) = *((_DWORD *)this + 609);
  v99 = *((_BYTE *)v94 + 1) ^ (*((_BYTE *)this + 2948) ^ *((_BYTE *)v94 + 1)) & 1;
  *((_BYTE *)v94 + 1) = v99;
  *((_DWORD *)v94 + 2) = *((_DWORD *)this + 684);
  v100 = v99 ^ (v99 ^ (32 * (*((_DWORD *)this + 744) >> 3))) & 0x20;
  v101 = v98 & 0xEF;
  *((_BYTE *)v94 + 1) = v100;
  *(_BYTE *)v94 = v98 & 0xEF;
  if ( v37 >= 0x5021 )
  {
    v101 = v98 ^ (v98 ^ (16 * *((_BYTE *)this + 2968))) & 0x10;
    *(_BYTE *)v94 = v101;
  }
  if ( *((_BYTE *)this + 209) )
    goto LABEL_259;
  if ( (v101 & 0x40) != 0 )
  {
    if ( v37 < 0x5005 && (*((_DWORD *)this + 464) || *((_DWORD *)this + 465)) )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports device 0x%I64x is hybrid discrete device but it has VidPn source and target.";
      WdLogGlobalForLineNumber = 7769;
      goto LABEL_61;
    }
    v102 = v100 ^ (v100 ^ (2 * *((_BYTE *)this + 2971))) & 2;
    *((_BYTE *)v94 + 1) = v102;
    v103 = v102 & 1;
    goto LABEL_236;
  }
  v103 = v100 & 1;
  if ( (v100 & 1) != 0 )
  {
LABEL_236:
    if ( (v101 & 0x20) == 0 )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports device 0x%I64x as hybrid device but does not support cross adapter resource.";
      WdLogGlobalForLineNumber = 7783;
      goto LABEL_61;
    }
  }
  v104 = v101 & 0x40;
  if ( v103 )
  {
    if ( v104 )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports both HybridIntegrated and HybridDiscrete caps 0x%I64x";
      WdLogGlobalForLineNumber = 7790;
      goto LABEL_61;
    }
    if ( !*((_DWORD *)this + 465) )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports the HybridIntegrated cap, but the adapter has no outputs. 0x%I64x";
      WdLogGlobalForLineNumber = 7798;
      goto LABEL_61;
    }
  }
  if ( *((_BYTE *)this + 2938) && (!*((_QWORD *)this + 101) || !*((_QWORD *)this + 102) || !*((_QWORD *)this + 103)) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7812LL;
    v40 = L"Driver reports SupportPerEngineTDR cap but does not fill in all of the required DDIs.";
    goto LABEL_60;
  }
  if ( (*((_DWORD *)this + 613) & 4) != 0 && !*((_QWORD *)this + 100) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7819LL;
    v40 = L"Driver reports SupportKernelModeCommandBuffer cap but does not fill in the pfnRenderKm DDI.";
    goto LABEL_60;
  }
  if ( *((_BYTE *)this + 2941) && (!*((_QWORD *)this + 105) || !*((_QWORD *)this + 106)) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7827LL;
    v40 = L"Driver reports SupportRuntimePowerManagement cap but does not fill in the pfnSetPowerComponentFState or pfnPow"
           "erRuntimeControlRequest DDI.";
    goto LABEL_60;
  }
  if ( v37 < 0x300C && *((_QWORD *)this + 105) && *((_QWORD *)this + 106) )
    *((_BYTE *)this + 2941) = 1;
LABEL_259:
  *((_WORD *)this + 1509) = 0;
  *((_BYTE *)this + 3020) = 0;
  if ( !*((_BYTE *)this + 2940) )
    goto LABEL_297;
  if ( v37 < 0x300B )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7849LL;
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but it is not compiled with expected header files.";
    goto LABEL_60;
  }
  if ( v37 < 0x4000 )
  {
    *((_BYTE *)this + 3018) = 1;
    goto LABEL_279;
  }
  if ( v37 == 0x4000 )
  {
    *((_BYTE *)this + 3019) = 1;
    goto LABEL_279;
  }
  v105 = *((_DWORD *)this + 736);
  if ( !v105 )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7862LL;
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but doesn't report any overlay planes or panel fitter.";
    goto LABEL_60;
  }
  if ( v105 <= 8 )
  {
    if ( v37 > 0x5000 )
      *((_BYTE *)this + 3020) = 1;
    goto LABEL_279;
  }
  v106 = *((_DWORD *)this + 684);
  if ( v106 < 8704 )
  {
    if ( v106 < 0x2000 || v105 != 10 )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7885LL;
      goto LABEL_272;
    }
    *((_DWORD *)this + 736) = 8;
  }
  else if ( v105 > 0xA )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7872LL;
LABEL_272:
    v40 = L"Driver reports more than the supported number of overlay planes.";
    goto LABEL_60;
  }
LABEL_279:
  v107 = *((_QWORD *)this + 109);
  if ( !v107 && !*((_QWORD *)this + 125) && !*((_QWORD *)this + 129) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7901LL;
LABEL_283:
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but does not fill in all of the required DDIs.";
    goto LABEL_60;
  }
  if ( v37 > 0x4002 && !*((_QWORD *)this + 113) && !*((_QWORD *)this + 124) && !*((_QWORD *)this + 128) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7913LL;
    goto LABEL_283;
  }
  if ( !*((_BYTE *)this + 2939) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7923LL;
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but DirectFlip is not supported.";
    goto LABEL_60;
  }
  if ( v107 )
  {
    v108 = DXGGLOBAL::GetGlobal();
    DXGGLOBAL::RecordFeatureUsage(v108, 1LL, 1LL);
  }
  if ( *((_QWORD *)this + 125) )
  {
    v109 = DXGGLOBAL::GetGlobal();
    DXGGLOBAL::RecordFeatureUsage(v109, 2LL, 1LL);
  }
  if ( *((_QWORD *)this + 129) )
  {
    v110 = DXGGLOBAL::GetGlobal();
    DXGGLOBAL::RecordFeatureUsage(v110, 3LL, 1LL);
  }
LABEL_297:
  v111 = *((_BYTE *)this + 209);
  *((_BYTE *)this + 3055) = 0;
  if ( v111 )
    goto LABEL_310;
  v112 = 0;
  if ( v37 >= 0x700A && *((int *)this + 684) >= 8704 && (!*((_QWORD *)this + 83) || *((_QWORD *)this + 146)) )
  {
    *((_BYTE *)this + 3055) = 1;
    v112 = 1;
  }
  if ( *((int *)this + 684) < 8960 )
  {
LABEL_310:
    *((_DWORD *)this + 612) &= 0xFFFFFFE3;
  }
  else
  {
    v113 = (*((_DWORD *)this + 612) >> 3) & 1;
    v114 = (*((_DWORD *)this + 612) >> 2) & 1;
    if ( v114 < v113 || v113 < ((*((_DWORD *)this + 612) >> 4) & 1u) )
    {
      WdLogSingleEntry2(2LL, *((_QWORD *)this + 27), -1073741811LL);
      v116 = *((_QWORD *)this + 27);
      WdLogGlobalForLineNumber = 7973;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Driver reports support higher level of colorSpaceTransform but not lower levels on device 0x%I64x,"
                       " returning 0x%I64x.",
        v116,
        -1073741811LL,
        0LL,
        0LL,
        0LL);
      return 3221225485LL;
    }
    if ( !v112 && v114 )
    {
      WdLogSingleEntry2(2LL, *((_QWORD *)this + 27), -1073741811LL);
      v115 = *((_QWORD *)this + 27);
      WdLogGlobalForLineNumber = 7981;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"ColorSpaceTransform is supported on the device 0x%I64x which does not have pfnSetTargetGamma, returning 0x%I64x.",
        v115,
        -1073741811LL,
        0LL,
        0LL,
        0LL);
      return 3221225485LL;
    }
  }
  if ( !*(_BYTE *)(*(_QWORD *)(*(_QWORD *)(*((_QWORD *)this + 27) + 64LL) + 40LL) + 133LL) && !v111 )
  {
    v117 = *((_DWORD *)this + 684) >= 0x2000;
    v118 = DXGGLOBAL::GetGlobal();
    v119 = DXGGLOBAL::DeferredInitialize(v118, v117);
    RenderCore = v119;
    if ( v119 < 0 )
    {
      WdLogSingleEntry1(2LL, v119);
      v86 = L"DXGGLOBAL::DeferredInitialize failed (Status = 0x%I64x).";
      WdLogGlobalForLineNumber = 8008;
LABEL_160:
      DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v86, RenderCore, 0LL, 0LL, 0LL, 0LL);
      return (unsigned int)RenderCore;
    }
  }
  DXGADAPTER::Config = 0;
  DXGADAPTER::ReadConfig(this, v94);
  DXGADAPTER::InitializeDriverWorkarounds(this);
  if ( *((_BYTE *)this + 209) )
  {
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFFDFFFF | v170 & 0x20000;
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFFE7FFF | v170 & 0x18000;
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFEFFFFF | v170 & 0x100000;
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFF3FFFF | v170 & 0xC0000;
    *((_BYTE *)this + 3021) = 0;
  }
  else if ( (*((_DWORD *)this + 111) & 0x10) != 0 && *((_BYTE *)this + 3071) )
  {
    *((_DWORD *)this + 617) |= 0x400u;
  }
  v120 = *((_DWORD *)this + 684);
  if ( v120 < 9216 )
    goto LABEL_323;
  v121 = *((_QWORD *)this + 167);
  if ( *((_QWORD *)this + 166) )
  {
    if ( v121 )
      goto LABEL_324;
LABEL_335:
    WdLogSingleEntry0(2LL);
    v39 = 8062LL;
    v40 = L"Driver cannot support only one of DdiQueryDiagnosticTypesSupport and DdiControlDiagnosticReporting.";
    goto LABEL_60;
  }
  if ( v121 )
    goto LABEL_335;
LABEL_323:
  *((_QWORD *)this + 166) = W32kStub_UserRemoveWindowedSwapChain;
  *((_QWORD *)this + 167) = DXGADAPTER::DefaultDdiControlDiagnosticReporting;
LABEL_324:
  if ( v120 >= 12800 && v37 >= 0x11001 )
  {
    memset(&v176, 0, 24);
    v176.Type = DXGKQAITYPE_POWERCOMPONENTINFO|0x20;
    *(_OWORD *)&v176.OutputDataSize = 0LL;
    v176.pOutputData = (char *)this + 5088;
    v176.OutputDataSize = 4;
    if ( (int)DXGADAPTER::DdiQueryAdapterInfo(this, &v176) < 0 )
    {
      *(_QWORD *)(WdLogNewEntry5_WdTrace(v122) + 24) = this;
      WdLogGlobalForLineNumber = 8077;
    }
  }
  v168 = 0;
  memset(&v177, 0, 24);
  v177.Type = DXGKQAITYPE_PHYSICALADAPTERCAPS|0x20;
  v177.pOutputData = &v168;
  *(_OWORD *)&v177.OutputDataSize = 0LL;
  v177.OutputDataSize = 4;
  v123 = DXGADAPTER::DdiQueryAdapterInfo(this, &v177);
  v124 = *((_BYTE *)this + 3072) & 0xFD;
  if ( v123 >= 0 )
    v124 |= 2 * (v168 & 1);
  *((_BYTE *)this + 3072) = v124;
  result = DXGADAPTER::CheckMcdmDdiOverall(this);
  if ( (int)result >= 0 )
  {
    DXGADAPTER::InitializeDriverDiagnosticReporting(this);
    DXGADAPTER::QueryFeatureEnablement(this);
    if ( (*((_DWORD *)this + 616) & 0x800) != 0 )
    {
      if ( (*((_DWORD *)this + 1257) & 0x40) == 0 )
      {
        WdLogSingleEntry0(2LL);
        v39 = 8117LL;
        v40 = L"Driver reports NativeGpuFence cap when NativeFence feature is disabled, returning failure";
        goto LABEL_60;
      }
      v173.Type = DXGKQAITYPE_QUERYSEGMENT3|0x20;
      v173.pOutputData = (char *)this + 5032;
      v173.OutputDataSize = 56;
      v125 = DXGADAPTER::DdiQueryAdapterInfo(this, &v173);
      RenderCore = v125;
      if ( v125 < 0 )
      {
        WdLogSingleEntry1(2LL, v125);
        v86 = L"Failed to get DXGK_NATIVE_FENCE_CAPS. Status 0x%I64x";
        WdLogGlobalForLineNumber = 8128;
        goto LABEL_160;
      }
    }
    RenderCore = (int)ADAPTER_RENDER::CreateRenderCore(this, (struct ADAPTER_RENDER **)this + 391);
    v126 = *((_QWORD *)this + 391);
    if ( (int)RenderCore < 0 )
    {
      if ( v126 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 8140;
        DxgkLogInternalTriageEvent(0, 262146, -1, (unsigned int)L"m_pRenderCore == NULL", 8140LL, 0LL, 0LL, 0LL, 0LL);
      }
      WdLogSingleEntry2(2LL, this, RenderCore);
      WdLogGlobalForLineNumber = 8143;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Failed to create ADAPTER_RENDER on adapter 0x%I64x (Status = 0x%I64x).",
        (__int64)this,
        RenderCore,
        0LL,
        0LL,
        0LL);
      return (unsigned int)RenderCore;
    }
    if ( v126 )
    {
      if ( IsAdapterSessionized )
      {
        WdLogSingleEntry0(2LL);
        v31 = L"Render capable adapter should NOT be sessionized!";
        v90 = 8159LL;
        WdLogGlobalForLineNumber = 8159;
        v160 = 0LL;
        goto LABEL_203;
      }
      if ( (*((_DWORD *)this + 744) & 0xC) == 0 )
        *((_BYTE *)this + 3072) |= 1u;
    }
    v127 = (char *)this + 3120;
    DisplayCore = ADAPTER_DISPLAY::CreateDisplayCore(this, (struct ADAPTER_DISPLAY **)this + 390);
    RenderCore = DisplayCore;
    if ( DisplayCore < 0 )
    {
      if ( *(_QWORD *)v127 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 8174;
        DxgkLogInternalTriageEvent(0, 262146, -1, (unsigned int)L"m_pDisplayCore == NULL", 8174LL, 0LL, 0LL, 0LL, 0LL);
      }
      WdLogSingleEntry2(2LL, this, RenderCore);
      WdLogGlobalForLineNumber = 8177;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Failed to create ADAPTER_DISPLAY on adapter 0x%I64x (Status = 0x%I64x).",
        (__int64)this,
        RenderCore,
        0LL,
        0LL,
        0LL);
      return (unsigned int)RenderCore;
    }
    if ( *((_QWORD *)this + 391) )
    {
      v129 = *(_QWORD *)v127 == 0LL;
    }
    else
    {
      v129 = *(_QWORD *)v127 == 0LL;
      if ( !*(_QWORD *)v127 )
      {
        WdLogSingleEntry2(2LL, this, -1073741735LL);
        v31 = L"Current adapter 0x%I64x does not have display or render capabilities (Status = 0x%I64x).";
        v160 = -1073741735LL;
        v154 = (__int64)this;
        WdLogGlobalForLineNumber = 8190;
        goto LABEL_40;
      }
    }
    v130 = *(_BYTE *)v94 & 0xFE | !v129;
    *(_BYTE *)v94 = v130;
    v131 = v130 & 0xFD | (*((_QWORD *)this + 391) != 0LL ? 2 : 0);
    *(_BYTE *)v94 = v131;
    if ( *(_QWORD *)v127 )
      v132 = *(_DWORD *)(*(_QWORD *)v127 + 24LL);
    else
      LOBYTE(v132) = 0;
    v133 = v131 & 0x7F | ((_BYTE)v132 << 7);
    *(_BYTE *)v94 = v133;
    if ( (v133 & 1) != 0 )
      *((_BYTE *)v94 + 1) = *((_BYTE *)v94 + 1) & 0xFB | (DXGADAPTER::DriverSupportSetTimingsFromVidPn(this) != 0 ? 4 : 0);
    else
      *((_BYTE *)v94 + 1) &= ~4u;
    if ( !*((_QWORD *)this + 391) )
      *((_DWORD *)this + 613) |= 1u;
    if ( DXGADAPTER::IsDxgmms2(this) )
    {
      v136 = *((_DWORD *)this + 111);
      if ( (v136 & 4) == 0
        && (v136 & 8) == 0
        && v134
        && v37 >= 0x5008
        && (!*((_QWORD *)this + 114) || !*((_QWORD *)this + 126)) )
      {
        WdLogSingleEntry0(2LL);
        v39 = 8231LL;
        v40 = L"Driver is compiled against DXGKDDI_INTERFACE_VERSION_WDDM2_0_M2_2_1 or greater, but does not fill in the p"
               "fnCalibrateGpuClock or pfnSetStablePowerState DDI.";
        goto LABEL_60;
      }
    }
    if ( *((_BYTE *)this + 3016) && DXGADAPTER::IsFullWDDMAdapter(v135) && *((int *)this + 684) >= 4608 )
    {
      if ( !*((_BYTE *)this + 2939) )
      {
        WdLogSingleEntry0(2LL);
        v39 = 8246LL;
        v40 = L"Driver reports WDDM version 1.2 but does not implement all mandatory WDDM 1.2 full adapter features.";
        goto LABEL_60;
      }
    }
    else if ( !*((_BYTE *)this + 2939) )
    {
      goto LABEL_381;
    }
    if ( *((_BYTE *)this + 209) )
      goto LABEL_382;
    v137 = *((_QWORD *)this + 391);
    if ( !v137
      || !(*(unsigned __int8 (__fastcall **)(_QWORD))(*(_QWORD *)(*(_QWORD *)(v137 + 760) + 8LL) + 656LL))(*(_QWORD *)(v137 + 768)) )
    {
      *(_WORD *)((char *)this + 2939) = 0;
    }
LABEL_381:
    if ( !*((_BYTE *)this + 209) )
    {
LABEL_383:
      v138 = DXGADAPTER::IsBddFallbackDriver(this) != 0;
      v141 = *((_DWORD *)this + 111);
      *((_DWORD *)this + 50) = v138 ? 3 : 1;
      if ( (v141 & 0x10) != 0 && !*((_QWORD *)this + 390) )
      {
        DXGGLOBALSHAREMUTEX::DXGGLOBALSHAREMUTEX((DXGGLOBALSHAREMUTEX *)v174);
        DXGAUTOMUTEX::Acquire((DXGAUTOMUTEX *)v174);
        if ( *((_QWORD *)DXGGLOBAL::GetGlobal() + 119) )
        {
          WdLogSingleEntry2(2LL, this, -1073741735LL);
          WdLogGlobalForLineNumber = 8296;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Current adapter 0x%I64x does not have display or render capabilities (Status = 0x%I64x).",
            (__int64)this,
            -1073741735LL,
            0LL,
            0LL,
            0LL);
        }
        else
        {
          _InterlockedIncrement64((volatile signed __int64 *)this + 3);
          *((_QWORD *)this + 4) = -1LL;
          v142 = DXGGLOBAL::GetGlobal();
          DXGGLOBAL::SetWarpAdapter(v142, this);
        }
        DXGPROCESSCOPYPROTECTIONMUTEX::~DXGPROCESSCOPYPROTECTIONMUTEX((DXGPROCESSCOPYPROTECTIONMUTEX *)v174);
      }
      if ( *((_BYTE *)this + 209)
        || (v143 = DXGADAPTER::InitializePowerManagement(this, v139, v140), RenderCore = v143, v143 >= 0) )
      {
        if ( *((_BYTE *)this + 3016) )
        {
          if ( *((int *)this + 684) >= 4864 )
          {
            if ( DXGADAPTER::IsFullWDDMAdapter(this) )
            {
              v145 = *((_DWORD *)this + 111);
              if ( (v145 & 4) == 0 && (v145 & 0x20) == 0 && (*((_DWORD *)this + 615) & 0x10) == 0 )
              {
                WdLogSingleEntry0(2LL);
                v39 = 8327LL;
                v40 = L"WDDM 1.3 driver must support independent flip.";
                goto LABEL_60;
              }
            }
          }
        }
      }
      else
      {
        WdLogSingleEntry2(2LL, this, v143);
        WdLogGlobalForLineNumber = 8314;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Failed to initialize power management for the adapter 0x%I64x (Status = 0x%I64x).",
          (__int64)this,
          RenderCore,
          0LL,
          0LL,
          0LL);
      }
      if ( (*((_DWORD *)this + 111) & 0x10) != 0 )
        *((_BYTE *)this + 3058) = 1;
      if ( v37 >= 0xA008 )
        *((_BYTE *)this + 3058) = 1;
      v144 = operator new(40LL, 1265072196LL);
      if ( v144 )
      {
        *(_OWORD *)v144 = 0LL;
        *(_OWORD *)(v144 + 16) = 0LL;
        *(_QWORD *)(v144 + 32) = 0LL;
      }
      else
      {
        v144 = 0LL;
      }
      *((_QWORD *)this + 621) = v144;
      if ( !v144 )
      {
        WdLogSingleEntry0(2LL);
        v25 = L"Failed to allocate MockDriverState object";
        v24 = 8365LL;
        WdLogGlobalForLineNumber = 8365;
        v159 = 0LL;
        goto LABEL_31;
      }
      LocallyUniqueId = MOCKDRIVERSTATE::Initialize((MOCKDRIVERSTATE *)v144, this);
      if ( LocallyUniqueId < 0 )
      {
        WdLogSingleEntry0(2LL);
        v11 = 8372LL;
        v12 = L"Failed to initialize MockDriverState object";
        v13 = 0x40000;
        goto LABEL_13;
      }
      *((_BYTE *)this + 4976) = 0;
      LocallyUniqueId = DXGADAPTER::InitializeVSyncPhaseState(this);
      if ( LocallyUniqueId < 0 )
      {
        WdLogSingleEntry0(6LL);
        v11 = 8385LL;
        v12 = L"Failed to allocate VSync Phase Timer state";
        goto LABEL_12;
      }
      if ( (int)DXGADAPTER::InitializeCABCStateV2(v146) < 0 )
      {
        WdLogSingleEntry0(2LL);
        WdLogGlobalForLineNumber = 8400;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Failed to initialize CABC State",
          8400LL,
          0LL,
          0LL,
          0LL,
          0LL);
      }
      v147 = *((_QWORD *)this + 391);
      if ( v147 && !*((_BYTE *)this + 209) )
      {
        v148 = *(_QWORD *)(v147 + 736);
        v149 = DXGGLOBAL::GetGlobal();
        (*(void (__fastcall **)(_QWORD, __int64))(*(_QWORD *)(v148 + 8) + 920LL))(
          *(_QWORD *)(v147 + 744),
          (__int64)v149 + 1328);
      }
      if ( (*((_DWORD *)this + 111) & 1) != 0 )
        *((_QWORD *)DXGGLOBAL::GetGlobal() + 123) = *(_QWORD *)((char *)this + 412);
      if ( (int)RenderCore < 0 )
        return (unsigned int)RenderCore;
      if ( v169 <= 1 )
        goto LABEL_426;
      v150 = *((_DWORD *)this + 105);
      if ( v150 == 4318 )
      {
        v151 = DXGGLOBAL::GetGlobal();
        v152 = 7LL;
      }
      else
      {
        if ( v150 != 4098 )
        {
LABEL_426:
          v153 = DXGGLOBAL::GetGlobal();
          DXGGLOBAL::RecordFeatureUsageWddmVersion(v153, this);
          return (unsigned int)RenderCore;
        }
        v151 = DXGGLOBAL::GetGlobal();
        v152 = 8LL;
      }
      DXGGLOBAL::RecordFeatureUsage(v151, v152, 1LL);
      goto LABEL_426;
    }
LABEL_382:
    *((_QWORD *)this + 114) = 0LL;
    goto LABEL_383;
  }
  return result;
}
