NTSTATUS __stdcall DriverEntry(_DRIVER_OBJECT *DriverObject, PUNICODE_STRING RegistryPath)
{
  NTSTATUS result; // eax
  int v5; // eax
  __int64 v6; // rdi
  const wchar_t *v7; // r9
  NTSTATUS v8; // eax
  int ProcessNotifyRoutineEx2; // eax
  __int64 v10; // rbx
  unsigned __int8 v11; // al
  __int64 v12; // rcx
  __int64 v13; // rcx
  __int64 v14; // rcx
  int v15; // eax
  __int64 v16; // rdx
  __int64 v17; // rbx
  __int64 v18; // rcx
  __int64 v19; // r8
  NTSTATUS v20; // eax
  int v21; // eax
  int v22; // eax
  int v23; // eax
  NTSTATUS v24; // eax
  __int64 v25; // rcx
  __int64 v26; // r8
  __int64 v27; // rcx
  __int64 v28; // r8
  BOOLEAN Size; // [rsp+28h] [rbp-D8h]
  unsigned int v30; // [rsp+50h] [rbp-B0h] BYREF
  __int64 v31; // [rsp+58h] [rbp-A8h]
  char v32; // [rsp+60h] [rbp-A0h]
  _QWORD v33[2]; // [rsp+68h] [rbp-98h] BYREF
  UNICODE_STRING DefaultSDDLString; // [rsp+78h] [rbp-88h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+88h] [rbp-78h] BYREF
  __int64 v36; // [rsp+A0h] [rbp-60h] BYREF
  int v37; // [rsp+A8h] [rbp-58h]
  const wchar_t *v38; // [rsp+B0h] [rbp-50h]
  unsigned __int8 *v39; // [rsp+B8h] [rbp-48h]
  int v40; // [rsp+C0h] [rbp-40h]
  unsigned __int8 *v41; // [rsp+C8h] [rbp-38h]
  int v42; // [rsp+D0h] [rbp-30h]
  __int64 v43; // [rsp+D8h] [rbp-28h]
  int v44; // [rsp+E0h] [rbp-20h]
  __int64 v45; // [rsp+E8h] [rbp-18h]
  __int128 v46; // [rsp+F0h] [rbp-10h]
  __int128 v47; // [rsp+100h] [rbp+0h]
  __int64 SystemInformation; // [rsp+130h] [rbp+30h] BYREF

  g_pDriverObject = (PDEVICE_OBJECT)DriverObject;
  g_RegistryPath.Buffer = (wchar_t *)operator new[](RegistryPath->MaximumLength, 1265072196LL, 256LL);
  if ( !g_RegistryPath.Buffer )
  {
    WdLogSingleEntry0(2LL);
    WdLogGlobalForLineNumber = 298;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Failed to allocate registry path buffer.",
      298LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return -1073741801;
  }
  g_RegistryPath.MaximumLength = RegistryPath->MaximumLength;
  RtlCopyUnicodeString(&g_RegistryPath, RegistryPath);
  v5 = PsTlsAlloc(DxgkThreadPsTslCallback, 0LL, &g_DxgkThreadTlsId);
  v6 = v5;
  if ( v5 < 0 )
  {
    WdLogSingleEntry1(2LL, v5);
    v7 = L"Failed to allocate a PsTls slot for DxgkThread, returning 0x%I64x.";
    WdLogGlobalForLineNumber = 311;
LABEL_5:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v7, v6, 0LL, 0LL, 0LL, 0LL);
    return v6;
  }
  v8 = ExInitializeLookasideListEx(&g_DxgkThreadLookasideList, 0LL, 0LL, (POOL_TYPE)512, 0, 0x40uLL, 0x54677844u, 0);
  v6 = v8;
  if ( v8 < 0 )
  {
    PsTlsFree(g_DxgkThreadTlsId);
    WdLogSingleEntry1(2LL, v6);
    v7 = L"Failed to initialize the lookaside list for DXGTHREAD, returning 0x%I64x";
    WdLogGlobalForLineNumber = 326;
    goto LABEL_5;
  }
  ProcessNotifyRoutineEx2 = PsSetCreateProcessNotifyRoutineEx2(0LL, DxgkProcessNotify, 0LL);
  if ( ProcessNotifyRoutineEx2 < 0 )
  {
    v10 = ProcessNotifyRoutineEx2;
    WdLogSingleEntry1(2LL, ProcessNotifyRoutineEx2);
    WdLogGlobalForLineNumber = 337;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"PsSetCreateProcessNotifyRoutineEx failed 0x%I64x",
      v10,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  SystemInformation = 8LL;
  if ( ZwQuerySystemInformation(MaxSystemInfoClass|SystemProcessInformation, &SystemInformation, 8u, 0LL) < 0
    || (v11 = 1, (SystemInformation & 0x200000000LL) == 0) )
  {
    v11 = 0;
  }
  g_OSTestSigningEnabled = v11;
  v36 = 0LL;
  v37 = 288;
  v40 = 67108868;
  v38 = L"IsInternalRelease";
  v42 = 4;
  v39 = &g_IsInternalRelease;
  v41 = &g_IsInternalRelease;
  v43 = 0LL;
  v44 = 0;
  v45 = 0LL;
  v46 = 0LL;
  v47 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v36);
  g_IsInternalRelease = g_IsInternalRelease != 0;
  g_IsInternalReleaseOrDbg = g_IsInternalRelease;
  g_bSkuSupportMultipleUsers = (RtlGetSuiteMask(v12) & 0x110) == 16;
  wil_InitializeFeatureStaging(v13);
  InitializeTelemetryAssertsKMByDriverObject(DriverObject);
  WdInitialize(v14);
  result = DXGGLOBAL::CreateGlobal();
  if ( result >= 0 )
  {
    result = DpiInitializeGlobalState();
    if ( result >= 0 )
    {
      result = CCD_BTL::CreateGlobal();
      if ( result >= 0 )
      {
        DxgkInitializeTelemetry();
        Size = 0;
        v15 = ExSubscribeWnfStateChange(&gScreenStudyEventSubscription, &WNF_SRUM_SCREENONSTUDY_SESSION, 1LL);
        if ( v15 < 0 )
        {
          v17 = v15;
          WdLogSingleEntry1(2LL, v15);
          WdLogGlobalForLineNumber = 447;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"ExSubscribeWnfStateChange failed, returing 0x%I64x",
            v17,
            0LL,
            0LL,
            0LL,
            0LL);
          gScreenStudyEventSubscription = 0LL;
        }
        bTracingEnabled = 0;
        McGenEventRegister_EtwRegister(&DxgkControlGuid, v16, &DxgkControlGuid_Context, &DxgkControlGuid_Context);
        v30 = -1;
        v31 = 0LL;
        if ( (qword_1C0151480 & 2) != 0 )
        {
          v32 = 1;
          v30 = 0;
          if ( (Microsoft_Windows_DxgKrnlEnableBits & 0x10000) != 0 )
            McTemplateK0q_EtwWriteTransfer(v18, &EventProfilerEnter, v19, 0LL);
        }
        else
        {
          v32 = 0;
        }
        DXGETWPROFILER_BASE_PushProfilerEntry(&v30, 0LL);
        v33[0] = &DxgkControlGuid;
        v33[1] = &Dxgk_WDI_NotifyUser;
        WdDiagInit(v33);
        DestinationString = 0LL;
        RtlInitUnicodeString(&DestinationString, L"\\Device\\DxgKrnl");
        DriverObject->MajorFunction[0] = (PDRIVER_DISPATCH)&DxgkCreateClose;
        DriverObject->MajorFunction[2] = (PDRIVER_DISPATCH)&DxgkCreateClose;
        DriverObject->MajorFunction[14] = (PDRIVER_DISPATCH)&DxgkDeviceIoctl;
        DriverObject->MajorFunction[15] = (PDRIVER_DISPATCH)&DxgkInternalDeviceIoctl;
        DriverObject->MajorFunction[16] = (PDRIVER_DISPATCH)&DxgkShutdown;
        DriverObject->DriverUnload = (PDRIVER_UNLOAD)DxgkUnload;
        DefaultSDDLString = 0LL;
        RtlInitUnicodeString(&DefaultSDDLString, L"D:P(A;;GRGW;;;S-1-5-83-0)");
        v20 = WdmlibIoCreateDeviceSecure(
                DriverObject,
                0,
                &DestinationString,
                0x22u,
                0x100u,
                Size,
                &DefaultSDDLString,
                &GUID_SD_DXGKRNL_DRIVER_OBJECT,
                &g_pDeviceObject);
        LODWORD(v6) = v20;
        if ( v20 < 0 )
        {
          WdLogSingleEntry1(3LL, v20);
          WdLogGlobalForLineNumber = 500;
LABEL_33:
          DxgkCleanupPower();
          MonitorCleanupGlobal();
          if ( g_pDeviceObject )
          {
            IoDeleteDevice(g_pDeviceObject);
            g_pDeviceObject = 0LL;
          }
          if ( g_RegistryPath.Buffer )
          {
            DXGQUOTAALLOCATOR<256,1835156294>::operator delete(g_RegistryPath.Buffer);
            g_RegistryPath = 0LL;
          }
          DXGGLOBAL::DestroyGlobal();
          PsTlsFree(g_DxgkThreadTlsId);
          ExDeleteLookasideListEx(&g_DxgkThreadLookasideList);
          DXGETWPROFILER_BASE::PopProfilerEntry((DXGETWPROFILER_BASE *)&v30);
          if ( v32 )
          {
            if ( (Microsoft_Windows_DxgKrnlEnableBits & 0x10000) != 0 )
              McTemplateK0q_EtwWriteTransfer(v25, &EventProfilerExit, v26, v30);
          }
          return v6;
        }
        v21 = DxgkInitialPower();
        LODWORD(v6) = v21;
        if ( v21 < 0 )
        {
          WdLogSingleEntry1(3LL, v21);
          WdLogGlobalForLineNumber = 513;
          goto LABEL_33;
        }
        v22 = MonitorInitializeGlobal();
        LODWORD(v6) = v22;
        if ( v22 < 0 )
        {
          WdLogSingleEntry1(3LL, v22);
          WdLogGlobalForLineNumber = 526;
          goto LABEL_33;
        }
        SysMmInitializeGlobal();
        DxgkInitTest();
        DxgDbgInit();
        TdrInit();
        v23 = SMgrRegisterSessionChangeCallout(DxgkNotifySessionStateChange);
        v6 = v23;
        if ( v23 < 0 )
        {
          WdLogSingleEntry1(2LL, v23);
          WdLogGlobalForLineNumber = 559;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Could not register session change callout with session manager, returning 0x%I64x.",
            v6,
            0LL,
            0LL,
            0LL,
            0LL);
          goto LABEL_33;
        }
        v24 = IoRegisterShutdownNotification(g_pDeviceObject);
        v6 = v24;
        if ( v24 < 0 )
        {
          WdLogSingleEntry1(2LL, v24);
          WdLogGlobalForLineNumber = 569;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Could not register for shutdown notification, returning 0x%I64x.",
            v6,
            0LL,
            0LL,
            0LL,
            0LL);
          SMgrUnregisterSessionChangeCallout(DxgkNotifySessionStateChange);
          goto LABEL_33;
        }
        DXGETWPROFILER_BASE::PopProfilerEntry((DXGETWPROFILER_BASE *)&v30);
        if ( v32 && (Microsoft_Windows_DxgKrnlEnableBits & 0x10000) != 0 )
          McTemplateK0q_EtwWriteTransfer(v27, &EventProfilerExit, v28, v30);
        return 0;
      }
    }
  }
  return result;
}
