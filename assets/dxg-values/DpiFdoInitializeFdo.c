__int64 __fastcall DpiFdoInitializeFdo(_QWORD *StartContext)
{
  __int64 v1; // rdi
  char v3; // r12
  char v4; // si
  char v5; // r15
  int v6; // eax
  size_t v7; // rbx
  void *Pool2; // rax
  __int64 v9; // rbx
  int DevicePropertyString; // eax
  int v11; // eax
  __int64 v12; // rcx
  struct _DEVICE_OBJECT *v13; // rcx
  int MiniportInterface; // eax
  struct _DEVICE_OBJECT *v15; // rcx
  NTSTATUS v16; // eax
  __int64 v17; // rax
  NTSTATUS v18; // eax
  __int64 v19; // rax
  NTSTATUS v20; // eax
  __int64 v21; // rax
  _WORD *v22; // rsi
  int v23; // eax
  bool v24; // bl
  __int64 v25; // rdx
  __int64 v26; // rax
  int v27; // eax
  size_t v28; // r8
  void *v29; // rcx
  void *v30; // rcx
  void *v31; // rcx
  void *v32; // rcx
  void *v33; // rcx
  void *v34; // rcx
  void *v35; // rcx
  void *v36; // rcx
  void (__fastcall *v37)(_QWORD); // rax
  void (__fastcall *v38)(_QWORD); // rax
  struct SYSMM_ADAPTER *v39; // rcx
  int Size; // [rsp+28h] [rbp-E0h]
  ULONG Sizea[2]; // [rsp+28h] [rbp-E0h]
  int Sizec; // [rsp+28h] [rbp-E0h]
  int Sizeb; // [rsp+28h] [rbp-E0h]
  char v45; // [rsp+48h] [rbp-C0h] BYREF
  char v46; // [rsp+49h] [rbp-BFh] BYREF
  char Data; // [rsp+4Ah] [rbp-BEh] BYREF
  ULONG RequiredSize; // [rsp+4Ch] [rbp-BCh] BYREF
  ULONG Type; // [rsp+50h] [rbp-B8h] BYREF
  unsigned int v50; // [rsp+54h] [rbp-B4h] BYREF
  _QWORD SymbolicLinkName[3]; // [rsp+58h] [rbp-B0h] BYREF
  __int64 v52; // [rsp+70h] [rbp-98h] BYREF
  void *ThreadHandle; // [rsp+78h] [rbp-90h] BYREF
  PVOID Object; // [rsp+80h] [rbp-88h] BYREF
  __int64 v55; // [rsp+88h] [rbp-80h] BYREF
  int v56; // [rsp+90h] [rbp-78h]
  const wchar_t *v57; // [rsp+98h] [rbp-70h]
  unsigned int *v58; // [rsp+A0h] [rbp-68h]
  int v59; // [rsp+A8h] [rbp-60h]
  unsigned int *v60; // [rsp+B0h] [rbp-58h]
  int v61; // [rsp+B8h] [rbp-50h]
  __int64 v62; // [rsp+C0h] [rbp-48h]
  int v63; // [rsp+C8h] [rbp-40h]
  const wchar_t *v64; // [rsp+D0h] [rbp-38h]
  unsigned int *v65; // [rsp+D8h] [rbp-30h]
  int v66; // [rsp+E0h] [rbp-28h]
  unsigned int *v67; // [rsp+E8h] [rbp-20h]
  int v68; // [rsp+F0h] [rbp-18h]
  __int64 v69; // [rsp+F8h] [rbp-10h]
  int v70; // [rsp+100h] [rbp-8h]
  const wchar_t *v71; // [rsp+108h] [rbp+0h]
  int *v72; // [rsp+110h] [rbp+8h]
  int v73; // [rsp+118h] [rbp+10h]
  int *v74; // [rsp+120h] [rbp+18h]
  int v75; // [rsp+128h] [rbp+20h]
  __int64 v76; // [rsp+130h] [rbp+28h]
  int v77; // [rsp+138h] [rbp+30h]
  const wchar_t *v78; // [rsp+140h] [rbp+38h]
  int *v79; // [rsp+148h] [rbp+40h]
  int v80; // [rsp+150h] [rbp+48h]
  int *v81; // [rsp+158h] [rbp+50h]
  int v82; // [rsp+160h] [rbp+58h]
  __int64 v83; // [rsp+168h] [rbp+60h]
  int v84; // [rsp+170h] [rbp+68h]
  const wchar_t *v85; // [rsp+178h] [rbp+70h]
  __int64 *v86; // [rsp+180h] [rbp+78h]
  int v87; // [rsp+188h] [rbp+80h]
  __int64 v88; // [rsp+190h] [rbp+88h]
  int v89; // [rsp+198h] [rbp+90h]
  __int64 v90; // [rsp+1A0h] [rbp+98h]
  int v91; // [rsp+1A8h] [rbp+A0h]
  __int64 v92; // [rsp+1B0h] [rbp+A8h]
  __int128 v93; // [rsp+1B8h] [rbp+B0h]
  __int128 v94; // [rsp+1C8h] [rbp+C0h]

  v1 = StartContext[8];
  RequiredSize = 0;
  Type = 0;
  ThreadHandle = 0LL;
  v3 = 0;
  *(_OWORD *)&SymbolicLinkName[1] = 0LL;
  *(_QWORD *)(v1 + 112) = &DpiFdoDispatchInternalIoctl;
  *(_QWORD *)(v1 + 144) = DpiFdoDispatchSystemControl;
  v4 = 0;
  v5 = 0;
  *(_QWORD *)(v1 + 352) = &DpiFdoHandleQueryInterface;
  *(_QWORD *)(v1 + 344) = &DpiFdoHandleQueryDeviceRelations;
  LODWORD(v52) = 0;
  v55 = 0LL;
  v57 = L"GpuVirtualizationFlags";
  v56 = 288;
  v50 = g_VgpuReplaceWarp != 0 ? 8 : 0;
  v58 = &v50;
  v59 = 67108868;
  v60 = &v50;
  v61 = 4;
  v64 = L"DisableVaBackedVm";
  v65 = &g_VgpuDisableVaBackedVm;
  v67 = &g_VgpuDisableVaBackedVm;
  v71 = L"VirtualGpuOnly";
  v72 = &g_VirtualGpuOnly;
  v74 = &g_VirtualGpuOnly;
  v78 = L"LimitNumberOfVfs";
  v79 = &g_LimitNumberOfVfs;
  v81 = &g_LimitNumberOfVfs;
  v85 = L"DisableVersionMismatchCheck";
  v86 = &v52;
  v62 = 0LL;
  v63 = 288;
  v66 = 67108868;
  v68 = 4;
  v69 = 0LL;
  v70 = 288;
  v73 = 67108868;
  v75 = 4;
  v76 = 0LL;
  v77 = 288;
  v80 = 67108868;
  v82 = 4;
  v83 = 0LL;
  v84 = 288;
  v87 = 67108868;
  v88 = 0LL;
  v89 = 0;
  v90 = 0LL;
  v91 = 0;
  v92 = 0LL;
  v93 = 0LL;
  v94 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v55);
  g_bCreateParavirtualizedGpu = v50 & 1;
  g_VgpuReplaceWarp = (v50 >> 3) & 1;
  v6 = *(_DWORD *)(v1 + 504);
  g_ForceSecureVirtualMachine = (v50 >> 2) & 1;
  if ( v6 )
  {
    v7 = (unsigned int)(8 * v6);
    Pool2 = (void *)ExAllocatePool2(64LL, v7, 1953656900LL);
    *(_QWORD *)(v1 + 2832) = Pool2;
    if ( !Pool2 )
    {
      LODWORD(v9) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 9400;
      goto LABEL_140;
    }
    memset(Pool2, 0, v7);
    **(_QWORD **)(v1 + 2832) = StartContext;
    *(_DWORD *)(v1 + 2840) = 1;
  }
  *(_DWORD *)(v1 + 3604) = -1;
  DevicePropertyString = DpiGetDevicePropertyString(
                           *(PDEVICE_OBJECT *)(v1 + 152),
                           DevicePropertyDeviceDescription,
                           (__int64)&RequiredSize);
  LODWORD(v9) = DevicePropertyString;
  if ( DevicePropertyString < 0 )
  {
    WdLogSingleEntry1(2LL, DevicePropertyString);
    WdLogGlobalForLineNumber = 9427;
LABEL_41:
    v4 = 0;
    goto LABEL_140;
  }
  DpiGetDevicePropertyDataString(
    *(PDEVICE_OBJECT *)(v1 + 152),
    (DEVPROPKEY *)&DEVPKEY_Device_DriverVersion,
    v1 + 4952,
    (__int64)&RequiredSize);
  IoGetDevicePropertyData(
    *(PDEVICE_OBJECT *)(v1 + 152),
    &DEVPKEY_Device_DriverDate,
    0,
    0,
    8u,
    (PVOID)(v1 + 4960),
    &RequiredSize,
    &Type);
  IoGetDevicePropertyData(
    *(PDEVICE_OBJECT *)(v1 + 152),
    &DEVPKEY_Device_DriverRank,
    0,
    0,
    4u,
    (PVOID)(v1 + 4968),
    &RequiredSize,
    &Type);
  if ( !(_DWORD)v52 )
  {
    v11 = DpiFdoValidateKmdAndPnpVersionMatch(v1);
    LODWORD(v9) = v11;
    if ( v11 < 0 )
    {
      WdLogSingleEntry1(2LL, v11);
      WdLogGlobalForLineNumber = 9474;
      goto LABEL_41;
    }
  }
  v12 = *(_QWORD *)(v1 + 152);
  v46 = 0;
  if ( (int)DpiGetDevicePropertyDataBoolean(v12, &DEVPKEY_Device_InstallInProgress, &v46) >= 0 && v46 )
  {
    v13 = *(struct _DEVICE_OBJECT **)(v1 + 152);
    v45 = 0;
    IoSetDevicePropertyData(v13, &DEVPKEY_Device_InstallInProgress, 0, 0, 0x11u, 1u, &v45);
  }
  if ( *(_BYTE *)(v1 + 1153) )
  {
    if ( *(_BYTE *)(v1 + 480) )
    {
      MiniportInterface = DpiQueryMiniportInterface(
                            (_DWORD)StartContext,
                            (unsigned int)&GUID_DEVINTERFACE_MSBDD_FALLBACK,
                            56,
                            1,
                            Size,
                            v1 + 944);
      LODWORD(v9) = MiniportInterface;
      if ( MiniportInterface < 0 || !*(_QWORD *)(v1 + 976) || !*(_QWORD *)(v1 + 984) || !*(_QWORD *)(v1 + 992) )
      {
        WdLogSingleEntry3(0LL, 275LL, 21LL, MiniportInterface, *(_QWORD *)Sizea);
        WdLogGlobalForLineNumber = 9527;
        goto LABEL_137;
      }
    }
  }
  v3 = 1;
  if ( *(_BYTE *)(v1 + 1158) )
  {
    v15 = *(struct _DEVICE_OBJECT **)(v1 + 152);
    Data = 0;
    if ( IoGetDevicePropertyData(v15, &DEVPKEY_Gpu_IddVirtualMonitorDevice, 0, 0, 1u, &Data, &RequiredSize, &Type) >= 0
      && Type == 17
      && RequiredSize == 1
      && Data == -1 )
    {
      *(_BYTE *)(v1 + 1159) = 1;
    }
  }
  v16 = IoRegisterDeviceInterface(
          *(PDEVICE_OBJECT *)(v1 + 152),
          &GUID_COMPUTE_DEVICE_ARRIVAL,
          0LL,
          (PUNICODE_STRING)&SymbolicLinkName[1]);
  LODWORD(v9) = v16;
  if ( v16 < 0 )
  {
    WdLogSingleEntry1(2LL, v16);
    WdLogGlobalForLineNumber = 9566;
    goto LABEL_41;
  }
  v4 = 1;
  v17 = ExAllocatePool2(64LL, WORD1(SymbolicLinkName[1]), 1953656900LL);
  *(_QWORD *)(v1 + 2856) = v17;
  if ( !v17 )
  {
    LODWORD(v9) = -1073741801;
    WdLogSingleEntry1(6LL, -1073741801LL);
    WdLogGlobalForLineNumber = 9587;
    goto LABEL_140;
  }
  *(_DWORD *)(v1 + 2848) = SymbolicLinkName[1];
  RtlCopyUnicodeString((PUNICODE_STRING)(v1 + 2848), (PCUNICODE_STRING)&SymbolicLinkName[1]);
  RtlFreeUnicodeString((PUNICODE_STRING)&SymbolicLinkName[1]);
  if ( !*(_BYTE *)(v1 + 2722) )
  {
    v18 = IoRegisterDeviceInterface(
            *(PDEVICE_OBJECT *)(v1 + 152),
            &GUID_DISPLAY_DEVICE_ARRIVAL,
            0LL,
            (PUNICODE_STRING)&SymbolicLinkName[1]);
    LODWORD(v9) = v18;
    if ( v18 < 0 )
    {
      WdLogSingleEntry1(2LL, v18);
      WdLogGlobalForLineNumber = 9617;
      goto LABEL_41;
    }
    v19 = ExAllocatePool2(64LL, WORD1(SymbolicLinkName[1]), 1953656900LL);
    *(_QWORD *)(v1 + 2872) = v19;
    if ( !v19 )
    {
      LODWORD(v9) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 9638;
      goto LABEL_140;
    }
    *(_DWORD *)(v1 + 2864) = SymbolicLinkName[1];
    RtlCopyUnicodeString((PUNICODE_STRING)(v1 + 2864), (PCUNICODE_STRING)&SymbolicLinkName[1]);
    RtlFreeUnicodeString((PUNICODE_STRING)&SymbolicLinkName[1]);
  }
  *(_BYTE *)(v1 + 482) = 0;
  *(_BYTE *)(v1 + 484) = 0;
  *(_QWORD *)(v1 + 488) = 0LL;
  if ( !*(_BYTE *)(v1 + 480) )
  {
    KeInitializeEvent((PRKEVENT)(v1 + 4056), SynchronizationEvent, 0);
    *(_QWORD *)(v1 + 4096) = v1 + 4088;
    *(_QWORD *)(v1 + 4088) = v1 + 4088;
    KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 4208));
    KeInitializeEvent((PRKEVENT)(v1 + 4224), NotificationEvent, 1u);
    KeInitializeEvent((PRKEVENT)(v1 + 4248), NotificationEvent, 1u);
    *(_BYTE *)(v1 + 484) = 1;
    *(_QWORD *)(v1 + 4272) = 0LL;
    *(_DWORD *)(v1 + 4216) = 0;
    memset((void *)(v1 + 4112), 0, 0x60uLL);
    *(_DWORD *)(v1 + 4128) = 1953656900;
    *(_DWORD *)(v1 + 4132) = 11;
    *(_DWORD *)(v1 + 4152) = 64;
    KeInitializeTimer((PKTIMER)(v1 + 4288));
    KeInitializeDpc((PRKDPC)(v1 + 4352), DpiSuspendAdapterDpc, (PVOID)v1);
    v20 = PsCreateSystemThread(&ThreadHandle, 0x1FFFFFu, 0LL, 0LL, 0LL, DpiPowerArbiterThread, StartContext);
    LODWORD(v9) = v20;
    if ( v20 < 0 )
    {
      WdLogSingleEntry1(2LL, v20);
      WdLogGlobalForLineNumber = 9716;
      goto LABEL_41;
    }
    Object = 0LL;
    v9 = ObReferenceObjectByHandle(ThreadHandle, 0x1FFFFFu, 0LL, 0, &Object, 0LL);
    *(_QWORD *)(v1 + 4048) = Object;
    ZwClose(ThreadHandle);
    if ( (int)v9 < 0 )
    {
      WdLogSingleEntry1(2LL, v9);
      WdLogGlobalForLineNumber = 9738;
      goto LABEL_41;
    }
  }
  KeInitializeEvent((PRKEVENT)(v1 + 3816), NotificationEvent, 1u);
  *(_QWORD *)(v1 + 3592) = v1 + 3584;
  *(_QWORD *)(v1 + 3584) = v1 + 3584;
  ExInitializeResourceLite((PERESOURCE)(v1 + 3424));
  *(_QWORD *)(v1 + 3624) = v1 + 3616;
  *(_QWORD *)(v1 + 3616) = v1 + 3616;
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 3608));
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 3640));
  KeInitializeEvent((PRKEVENT)(v1 + 3648), NotificationEvent, 1u);
  *(_QWORD *)(v1 + 5456) = v1 + 5448;
  *(_QWORD *)(v1 + 5448) = v1 + 5448;
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 5464));
  IoCsqInitialize(
    (PIO_CSQ)(v1 + 5384),
    DpiPendingIrpCancelQueueInsert,
    DpiPendingIrpCancelQueueRemove,
    DpiPendingIrpCancelQueuePick,
    DpiPendingIrpCancelQueueAcquireLock,
    DpiPendingIrpCancelQueueReleaseLock,
    DpiPendingIrpCancelQueueComplete);
  *(_QWORD *)(v1 + 5536) = 0LL;
  *(_QWORD *)(v1 + 5544) = 0LL;
  KeInitializeEvent((PRKEVENT)(v1 + 5552), NotificationEvent, 0);
  *(_DWORD *)(v1 + 5528) = 1;
  *(_DWORD *)(v1 + 5496) = 0;
  KeInitializeMutex((PRKMUTEX)(v1 + 3528), 0);
  KeInitializeMutex((PRKMUTEX)(v1 + 3704), 0);
  *(_QWORD *)(v1 + 3776) = v1 + 3768;
  *(_QWORD *)(v1 + 3768) = v1 + 3768;
  *(_QWORD *)(v1 + 3800) = v1 + 3792;
  *(_QWORD *)(v1 + 3792) = v1 + 3792;
  *(_QWORD *)(v1 + 3696) = v1 + 3688;
  *(_QWORD *)(v1 + 3688) = v1 + 3688;
  ExInitializeResourceLite((PERESOURCE)(v1 + 3912));
  LODWORD(v9) = DpiFdoInitializeAdapterUniqueString(StartContext);
  v4 = 0;
  if ( (int)v9 < 0 )
  {
LABEL_139:
    ExDeleteResourceLite((PERESOURCE)(v1 + 3912));
    ExDeleteResourceLite((PERESOURCE)(v1 + 3424));
    goto LABEL_140;
  }
  v5 = 1;
  DpiQueryBusInterface(*(PDEVICE_OBJECT *)(v1 + 152), v1 + 2992);
  DpiQueryBusInterface(*(PDEVICE_OBJECT *)(v1 + 152), v1 + 3040);
  DpiQueryMiniportInterface((_DWORD)StartContext, (unsigned int)&GUID_DEVINTERFACE_I2C, 48, 1, Sizec, v1 + 3088);
  v21 = *(_QWORD *)(v1 + 40);
  *(_DWORD *)(v1 + 6008) = 1;
  if ( !*(_BYTE *)(v21 + 133) && !*(_BYTE *)(v1 + 1158) )
  {
    v22 = (_WORD *)(v1 + 5880);
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_WDDM_INTERFACE_DISPLAYMUX_2,
                128,
                2,
                Sizeb,
                v1 + 5880) >= 0 )
    {
      if ( *v22 != 128
        || *(_WORD *)(v1 + 5882) != 2
        || !*(_QWORD *)(v1 + 5912)
        || !*(_QWORD *)(v1 + 5920)
        || !*(_QWORD *)(v1 + 5928)
        || !*(_QWORD *)(v1 + 5936)
        || !*(_QWORD *)(v1 + 5944)
        || !*(_QWORD *)(v1 + 5952)
        || !*(_QWORD *)(v1 + 5960)
        || !*(_QWORD *)(v1 + 5968)
        || !*(_QWORD *)(v1 + 5976)
        || !*(_QWORD *)(v1 + 5984)
        || !*(_QWORD *)(v1 + 5992)
        || !*(_QWORD *)(v1 + 6000) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        WdLogGlobalForLineNumber = 9887;
        goto LABEL_85;
      }
      LODWORD(SymbolicLinkName[0]) = 0;
      if ( (int)DpiDxgkDdiDisplayMuxGetDriverSupportLevel(v1, SymbolicLinkName) < 0 )
      {
        *(_DWORD *)(v1 + 6008) = 1;
      }
      else
      {
        v23 = SymbolicLinkName[0];
        *(_DWORD *)(v1 + 6008) = SymbolicLinkName[0];
        if ( v23 != 1 )
        {
          v24 = DISPLAY_MUX_MGR::DisplayMuxPresent(qword_1C0154100);
          if ( DISPLAY_MUX_MGR::ShouldHideMuxFromDriver(qword_1C0154100) )
          {
            WdLogSingleEntry0(4LL);
            WdLogGlobalForLineNumber = 9915;
            v24 = 0;
          }
          LOBYTE(v25) = v24;
          DpiDxgkDdiDisplayMuxReportPresence(v1, v25);
          *(_BYTE *)(v1 + 6377) = v24;
        }
      }
    }
  }
  v26 = *(_QWORD *)(v1 + 40);
  *(_DWORD *)(v1 + 3136) = 0;
  if ( !*(_BYTE *)(v26 + 133) || *(_BYTE *)(v1 + 1158) )
  {
    v22 = (_WORD *)(v1 + 3144);
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM_3,
                128,
                4,
                Sizeb,
                v1 + 3144) >= 0 )
    {
      if ( *v22 != 128
        || (v27 = 4, *(_WORD *)(v1 + 3146) != 4)
        || !*(_QWORD *)(v1 + 3176)
        || !*(_QWORD *)(v1 + 3184)
        || !*(_QWORD *)(v1 + 3192)
        || !*(_QWORD *)(v1 + 3200)
        || !*(_QWORD *)(v1 + 3208)
        || !*(_QWORD *)(v1 + 3216)
        || !*(_QWORD *)(v1 + 3224)
        || !*(_QWORD *)(v1 + 3232)
        || !*(_QWORD *)(v1 + 3240)
        || !*(_QWORD *)(v1 + 3248)
        || !*(_QWORD *)(v1 + 3256)
        || !*(_QWORD *)(v1 + 3264) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        WdLogGlobalForLineNumber = 10051;
LABEL_85:
        v28 = 128LL;
LABEL_86:
        v29 = v22;
LABEL_87:
        memset(v29, 0, v28);
        v4 = 0;
        goto LABEL_139;
      }
      goto LABEL_115;
    }
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM_2,
                112,
                3,
                Sizeb,
                v1 + 3144) >= 0 )
    {
      if ( *v22 != 112
        || (v27 = 3, *(_WORD *)(v1 + 3146) != 3)
        || !*(_QWORD *)(v1 + 3176)
        || !*(_QWORD *)(v1 + 3184)
        || !*(_QWORD *)(v1 + 3192)
        || !*(_QWORD *)(v1 + 3200)
        || !*(_QWORD *)(v1 + 3208)
        || !*(_QWORD *)(v1 + 3216)
        || !*(_QWORD *)(v1 + 3224)
        || !*(_QWORD *)(v1 + 3232)
        || !*(_QWORD *)(v1 + 3240)
        || !*(_QWORD *)(v1 + 3248) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        v28 = 112LL;
        WdLogGlobalForLineNumber = 10102;
        goto LABEL_86;
      }
      goto LABEL_115;
    }
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM_2_JTP,
                120,
                2,
                Sizeb,
                v1 + 3144) >= 0 )
    {
      v27 = 2;
      if ( *v22 != 120
        || *(_WORD *)(v1 + 3146) != 2
        || !*(_QWORD *)(v1 + 3176)
        || !*(_QWORD *)(v1 + 3184)
        || !*(_QWORD *)(v1 + 3192)
        || !*(_QWORD *)(v1 + 3200)
        || !*(_QWORD *)(v1 + 3208)
        || !*(_QWORD *)(v1 + 3216)
        || !*(_QWORD *)(v1 + 3224)
        || !*(_QWORD *)(v1 + 3232)
        || !*(_QWORD *)(v1 + 3240)
        || !*(_QWORD *)(v1 + 3256) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        v28 = 120LL;
        WdLogGlobalForLineNumber = 10155;
        goto LABEL_86;
      }
LABEL_115:
      *(_DWORD *)(v1 + 3136) = v27;
      goto LABEL_119;
    }
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM,
                104,
                1,
                Sizeb,
                v1 + 3144) >= 0 )
      *(_DWORD *)(v1 + 3136) = 1;
  }
LABEL_119:
  *(_DWORD *)(v1 + 3344) = -1;
  if ( byte_1C0153A96
    && *(_DWORD *)(*(_QWORD *)(StartContext[8] + 40LL) + 28LL) >= 0x4000u
    && (!*(_BYTE *)(*(_QWORD *)(v1 + 40) + 133LL) || *(_BYTE *)(v1 + 1158)) )
  {
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_MIRACAST_DISPLAY,
                64,
                1,
                Sizeb,
                v1 + 3272) < 0 )
    {
      memset((void *)(v1 + 3272), 0, 0x40uLL);
    }
    else if ( *(_WORD *)(v1 + 3272) < 0x40u
           || *(_WORD *)(v1 + 3274) != 1
           || !*(_QWORD *)(v1 + 3304)
           || !*(_QWORD *)(v1 + 3312)
           || !*(_QWORD *)(v1 + 3320)
           || !*(_QWORD *)(v1 + 3328) )
    {
      LODWORD(v9) = -1073741811;
      WdLogSingleEntry1(2LL, -1073741811LL);
      v29 = (void *)(v1 + 3272);
      WdLogGlobalForLineNumber = 10232;
      v28 = 64LL;
      goto LABEL_87;
    }
  }
  if ( *(_BYTE *)(v1 + 1159) )
    *(_QWORD *)(v1 + 120) = DpiFdoDispatchIoctl;
  if ( *(_BYTE *)(v1 + 1158) )
  {
    *(_QWORD *)(v1 + 104) = &DpiFdoDispatchCreate;
    *(_QWORD *)(v1 + 96) = &DpiFdoDispatchCleanupAndClose;
  }
  v9 = StartContext[8];
  memset((void *)(v9 + 4504), 0, 0x48uLL);
  memset((void *)(v9 + 4576), 0, 0x48uLL);
  memset((void *)(v9 + 4648), 0, 0x58uLL);
  *(_OWORD *)(v9 + 4736) = 0LL;
  *(_OWORD *)(v9 + 4752) = 0LL;
  *(_OWORD *)(v9 + 4768) = 0LL;
  *(_QWORD *)(v9 + 4784) = 0LL;
  memset((void *)(v9 + 4792), 0, 0x58uLL);
  LODWORD(v9) = DpiInitializeBlockList(StartContext);
LABEL_137:
  v5 = v3;
  if ( (int)v9 >= 0 )
    return (unsigned int)v9;
  v4 = 0;
  if ( v3 == 1 )
    goto LABEL_139;
LABEL_140:
  if ( *(_QWORD *)(v1 + 4048) )
    DpiRequestIoPowerState(StartContext, 7LL);
  if ( v4 == 1 )
    RtlFreeUnicodeString((PUNICODE_STRING)&SymbolicLinkName[1]);
  if ( v5 )
  {
    RtlFreeUnicodeString((PUNICODE_STRING)(v1 + 4880));
    RtlFreeUnicodeString((PUNICODE_STRING)(v1 + 4896));
  }
  v30 = *(void **)(v1 + 3416);
  *(_DWORD *)(v1 + 3400) = 0;
  if ( v30 )
  {
    ExFreePoolWithTag(v30, 0);
    *(_QWORD *)(v1 + 3416) = 0LL;
  }
  v31 = *(void **)(v1 + 3408);
  if ( v31 )
  {
    ExFreePoolWithTag(v31, 0);
    *(_QWORD *)(v1 + 3408) = 0LL;
  }
  v32 = *(void **)(v1 + 4944);
  if ( v32 )
  {
    ExFreePoolWithTag(v32, 0);
    *(_QWORD *)(v1 + 4944) = 0LL;
  }
  v33 = *(void **)(v1 + 4952);
  if ( v33 )
  {
    ExFreePoolWithTag(v33, 0);
    *(_QWORD *)(v1 + 4952) = 0LL;
  }
  v34 = *(void **)(v1 + 2832);
  if ( v34 )
  {
    ExFreePoolWithTag(v34, 0);
    *(_QWORD *)(v1 + 2832) = 0LL;
  }
  v35 = *(void **)(v1 + 2856);
  if ( v35 )
  {
    ExFreePoolWithTag(v35, 0);
    *(_QWORD *)(v1 + 2856) = 0LL;
  }
  v36 = *(void **)(v1 + 2872);
  if ( v36 )
  {
    ExFreePoolWithTag(v36, 0);
    *(_QWORD *)(v1 + 2872) = 0LL;
  }
  v37 = *(void (__fastcall **)(_QWORD))(v1 + 3016);
  if ( v37 )
  {
    v37(*(_QWORD *)(v1 + 3000));
    *(_OWORD *)(v1 + 2992) = 0LL;
    *(_OWORD *)(v1 + 3008) = 0LL;
    *(_OWORD *)(v1 + 3024) = 0LL;
  }
  v38 = *(void (__fastcall **)(_QWORD))(v1 + 3064);
  if ( v38 )
  {
    v38(*(_QWORD *)(v1 + 3048));
    *(_OWORD *)(v1 + 3040) = 0LL;
    *(_OWORD *)(v1 + 3056) = 0LL;
    *(_OWORD *)(v1 + 3072) = 0LL;
  }
  v39 = *(struct SYSMM_ADAPTER **)(v1 + 5808);
  if ( v39 )
    SysMmDestroyAdapter(v39);
  return (unsigned int)v9;
}
