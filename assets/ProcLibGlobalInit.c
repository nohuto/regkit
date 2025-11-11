__int64 __fastcall ProcLibGlobalInit(PDEVICE_OBJECT DeviceObject)
{
  NTSTATUS v2; // eax
  int v3; // edx
  unsigned int v4; // edi
  int v5; // r9d
  char v6; // al
  int v7; // r9d
  int v8; // r8d
  struct _IO_WORKITEM *WorkItem; // rax
  __int64 v10; // rcx
  _QWORD *v11; // rax
  __int64 v12; // rdx
  __int64 v13; // rcx
  bool v14; // si
  char v15; // r15
  __int64 v21; // rdx
  __int16 v22; // r11
  bool v23; // zf
  unsigned int v24; // ebx
  __int64 v25; // rax
  unsigned __int64 v26; // rdx
  __int64 v27; // rax
  const char *v28; // rdx
  const char *v29; // rax
  __int64 v30; // rdx
  char v32; // [rsp+30h] [rbp-51h]
  char v33; // [rsp+38h] [rbp-49h]
  char v34; // [rsp+48h] [rbp-39h] BYREF
  _BYTE v35[3]; // [rsp+49h] [rbp-38h] BYREF
  int v36; // [rsp+4Ch] [rbp-35h] BYREF
  int v37; // [rsp+50h] [rbp-31h] BYREF
  int v38; // [rsp+54h] [rbp-2Dh] BYREF
  int v39; // [rsp+58h] [rbp-29h] BYREF
  int v40; // [rsp+5Ch] [rbp-25h] BYREF
  __int128 v41; // [rsp+60h] [rbp-21h] BYREF
  __int128 v42; // [rsp+70h] [rbp-11h]
  __int64 v43; // [rsp+80h] [rbp-1h]
  _OWORD InputBuffer[2]; // [rsp+88h] [rbp+7h] BYREF

  v40 = 0;
  v38 = 0;
  v39 = 0;
  v35[0] = 0;
  v34 = 0;
  LODWORD(v43) = 0;
  memset(InputBuffer, 0, sizeof(InputBuffer));
  v41 = 0LL;
  v42 = 0LL;
  v2 = ZwPowerInformation(ProcessorStateHandler, 0LL, (ULONG)0, &dword_1C00137A8, (ULONG)280);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v5 = 10;
LABEL_122:
      v8 = 3;
      goto LABEL_123;
    }
    return v4;
  }
  v6 = dword_1C00137A8;
  if ( dword_1C00137A8 != 82 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return (unsigned int)-1073741735;
    v7 = 11;
    v33 = 82;
LABEL_7:
    LOBYTE(v3) = 2;
    WPP_RECORDER_SF_DD(
      WPP_GLOBAL_Control->DeviceExtension,
      v3,
      3,
      v7,
      (__int64)&WPP_4e1b20cf9f023c365f1b3d32753808d1_Traceguids,
      v6,
      v33);
    return (unsigned int)-1073741735;
  }
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 2496))(
         WdfDriverGlobals,
         0LL,
         &qword_1C0013498);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return v4;
    v5 = 12;
LABEL_12:
    v8 = 4;
LABEL_123:
    v32 = v2;
    goto LABEL_124;
  }
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 2496))(
         WdfDriverGlobals,
         0LL,
         &qword_1C00134A8);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return v4;
    v5 = 13;
    goto LABEL_12;
  }
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 2496))(
         WdfDriverGlobals,
         0LL,
         &qword_1C00134A0);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return v4;
    v5 = 14;
    goto LABEL_12;
  }
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 2496))(
         WdfDriverGlobals,
         0LL,
         &qword_1C0013AD8);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return v4;
    v5 = 15;
    goto LABEL_12;
  }
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 2496))(
         WdfDriverGlobals,
         0LL,
         &qword_1C00134B0);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return v4;
    v5 = 16;
    goto LABEL_12;
  }
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 2520))(
         WdfDriverGlobals,
         0LL,
         &qword_1C0013AF8);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return v4;
    v5 = 17;
    goto LABEL_12;
  }
  v2 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, __int64 *))(WdfFunctions_01015 + 2520))(
         WdfDriverGlobals,
         0LL,
         &qword_1C0013AD0);
  v4 = v2;
  if ( v2 < 0 )
  {
    if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      return v4;
    v5 = 18;
    goto LABEL_12;
  }
  KeInitializeEvent(&Event, NotificationEvent, 1u);
  WorkItem = IoAllocateWorkItem(DeviceObject);
  word_1C0013B0C = 0;
  v10 = 2LL;
  qword_1C0013B00 = WorkItem;
  qword_1C00134E0 = (__int64)&qword_1C00134D8;
  qword_1C00134D8 = (__int64)&qword_1C00134D8;
  qword_1C00134F0 = (__int64)&qword_1C00134E8;
  qword_1C00134E8 = (__int64)&qword_1C00134E8;
  qword_1C0013500 = (__int64)&qword_1C00134F8;
  qword_1C00134F8 = (__int64)&qword_1C00134F8;
  qword_1C0013510 = (__int64)&qword_1C0013508;
  qword_1C0013508 = (__int64)&qword_1C0013508;
  qword_1C0013520 = (__int64)&qword_1C0013518;
  qword_1C0013518 = (__int64)&qword_1C0013518;
  qword_1C0013AE8 = (__int64)&qword_1C0013AE0;
  qword_1C0013AE0 = (__int64)&qword_1C0013AE0;
  qword_1C0013DD8 = (__int64)&qword_1C0013DD0;
  qword_1C0013DD0 = (__int64)&qword_1C0013DD0;
  v11 = &unk_1C00134B8;
  qword_1C0013DC8 = 0LL;
  do
  {
    v11[1] = v11;
    *v11 = v11;
    v11 += 2;
    --v10;
  }
  while ( v10 );
  GetRegistryDwordValue(
    L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Processor",
    L"AllowPepPerfStates",
    &v40);
  GetRegistryDwordValue(
    L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Processor",
    L"Overrides",
    &dword_1C0013490);
  GetRegistryQwordValue(v13, v12, &qword_1C0013488);
  GetRegistryDwordValue(L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Processor", L"DisableAsserts", &v38);
  if ( v38 )
    byte_1C0013B28 = 1;
  GetRegistryDwordValue(
    L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Session Manager\\Throttle",
    L"PerfEnablePackageIdle",
    &v39);
  qword_1C00139C8 = (__int64)RegisterKernelIdleStates;
  word_1C0013DA1 = 0;
  qword_1C0013A10 = (__int64)RegisterHiddenIdleStates;
  byte_1C0013A18 = v39 == 0;
  dword_1C0013A1C = 0;
  qword_1C00139D0 = (__int64)RegisterKernelPerfStates;
  qword_1C00139E0 = (__int64)RegisterKernelPerfFeedback;
  qword_1C00139E8 = (__int64)RegisterKernelLegacyPcc;
  qword_1C00139D8 = (__int64)RegisterKernelCap;
  qword_1C00139F0 = (__int64)RegisterKernelCpc;
  qword_1C00139F8 = (__int64)RegisterKernelPepPerf;
  qword_1C0013A00 = (__int64)GetNtProcessorNumber;
  qword_1C0013A08 = (__int64)RegisterKernelPackage;
  v37 = 0;
  HviGetHypervisorFeatures(InputBuffer);
  v14 = 0;
  byte_1C0013A20 = 0;
  v15 = 0;
  if ( (unsigned __int8)HviIsAnyHypervisorPresent() )
  {
    v14 = (BYTE8(InputBuffer[0]) & 0x20) != 0;
    LOBYTE(word_1C0013DA1) = 1;
    if ( (unsigned __int8)HviIsHypervisorMicrosoftCompatible() )
    {
      _RAX = 1073741828LL;
      __asm { cpuid }
    }
    InputBuffer[0] = 0LL;
    HviGetHypervisorFeatures(InputBuffer);
    if ( (*(_QWORD *)&InputBuffer[0] & 0x100000000000LL) != 0 && (v22 & 0x1000) == 0 )
    {
      GetHvPpmCapabilities(&v34, v21, v35);
      if ( v34 )
      {
        v4 = InitializeHvProcessorInfo();
        if ( (v4 & 0x80000000) != 0 )
          return v4;
        dword_1C0013A1C = 1;
        qword_1C00139D0 = (__int64)RegisterHvPerfStatesCounters;
        qword_1C0013A10 = (__int64)RegisterHvIdleStates;
        qword_1C00139E0 = (__int64)RegisterHvPerfFeedbackCounters;
        qword_1C00139E8 = (__int64)RegisterHvLegacyPccCounters;
        qword_1C00139F0 = (__int64)RegisterHvCpcCounters;
        if ( v35[0] )
          word_1C0013DA1 = 256;
        else
          qword_1C00139C8 = (__int64)RegisterHvIdleStates;
        byte_1C0013DE0 = 1;
        qword_1C0013A00 = (__int64)GetLpIndex;
        qword_1C0013A08 = (__int64)RegisterHvPackage;
      }
      dword_1C0013D98 = GetHiddenProcessorPresence();
      goto LABEL_60;
    }
    v15 = 1;
    qword_1C00139C8 = (__int64)RegisterGuestIdleStates;
    if ( (v22 & 0x1000) != 0 )
    {
      GetHvPpmCapabilities(&v34, v21, v35);
      if ( v34 )
        qword_1C00139E0 = (__int64)RegisterHvPerfFeedbackCounters;
    }
    GetRegistryDwordValue(
      L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Processor",
      L"AllowGuestPerfStates",
      &v37);
    if ( !v37 )
    {
      qword_1C00139D0 = (__int64)RegisterNoop;
      qword_1C00139E8 = (__int64)RegisterNoop;
      qword_1C00139D8 = (__int64)RegisterNoop;
      qword_1C00139F0 = (__int64)RegisterNoop;
      qword_1C00139F8 = (__int64)RegisterNoop;
      if ( v14 )
      {
        byte_1C0013A20 = 1;
      }
      else
      {
        qword_1C00139E0 = (__int64)RegisterNoop;
        qword_1C0013A08 = (__int64)RegisterNoop;
      }
    }
  }
  if ( (int)HalPrivateDispatchTable[143]((__int64)&v41) >= 0 )
  {
    v6 = v41;
    if ( (_DWORD)v41 != 1 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        return (unsigned int)-1073741735;
      v7 = 19;
      v33 = 1;
      goto LABEL_7;
    }
    qword_1C0013860 = *((_QWORD *)&v41 + 1);
    xmmword_1C0013868 = v42;
    qword_1C0013878 = v43;
  }
  dword_1C0013D9C = dword_1C0013494 + HalPrivateDispatchTable[145](0xFFFFFFFFLL);
  dword_1C0013D98 = 2;
LABEL_60:
  DevExts = ExAllocatePool2(64LL, 0x4000LL, 1919119952LL);
  if ( DevExts )
  {
    v2 = EtwRegister(&PPM_ETW_PROVIDER, ProcLibTraceControlCallback, 0LL, &ProcLibEtwHandle);
    v4 = v2;
    if ( v2 < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        v5 = 21;
        goto LABEL_122;
      }
    }
    else
    {
      TraceLoggingRegisterEx_EtwRegister_EtwSetInformation(&dword_1C0013068);
      ProcLibEtwRegistered = 1;
      *((_QWORD *)&InputBuffer[0] + 1) = 0LL;
      *(_QWORD *)&InputBuffer[0] = ProcessSystemSleepStateNotify;
      v2 = ZwPowerInformation(SystemPowerStateNotifyHandler, InputBuffer, (ULONG)16, 0LL, (ULONG)0);
      v4 = v2;
      if ( v2 >= 0 )
      {
        v2 = CollectAcpiBiosInfo();
        v4 = v2;
        if ( v2 >= 0 )
        {
          v36 = 1;
          EmClientQueryRuleState(&GUID_EM_RULE_DISABLE_PSTATES, &v36);
          v23 = v36 == 2;
          v36 = 1;
          v24 = 0;
          if ( v23 )
            v24 = 1879048192;
          EmClientQueryRuleState(&GUID_EM_RULE_DISABLE_ACPI1_CSTATE_C2, &v36);
          if ( v36 == 2 )
            v24 |= 2u;
          v36 = 1;
          EmClientQueryRuleState(&GUID_EM_RULE_DISABLE_TSTATES, &v36);
          if ( v36 == 2 )
            v24 |= 0x3300000u;
          v36 = 1;
          EmClientQueryRuleState(&GUID_EM_RULE_DISABLE_PCC, &v36);
          if ( v36 == 2 )
            v24 |= 0x80000000;
          v25 = 0x180891100277LL;
          qword_1C0013488 = v24 | (unsigned __int64)qword_1C0013488;
          dword_1C0013A34 = v24;
          dword_1C0013630 = 1;
          dword_1C0013634 = 376;
          if ( v40 )
            v25 = 0x181891500277LL;
          Globals = v25 | 0x2010408800400LL;
          if ( (unsigned __int8)PoEnergyEstimationEnabled() )
          {
            Globals |= 0x2000000000uLL;
            PopulateEnergyEstimationParameters();
          }
          TraceLoggingRegisterEx_EtwRegister_EtwSetInformation(&dword_1C00130A0);
          v4 = 0;
          v27 = Globals | 0x40000000;
          Globals |= 0x40000000uLL;
          if ( v15 )
          {
            qword_1C0013658 = 0LL;
            qword_1C0013670 = 0LL;
            qword_1C0013678 = 0LL;
            qword_1C0013680 = 0LL;
            qword_1C00136A0 = 0LL;
            qword_1C0013688 = 0LL;
            qword_1C0013690 = 0LL;
            qword_1C00136B0 = 0LL;
            qword_1C00136B8 = 0LL;
            qword_1C00136C0 = 0LL;
            byte_1C0013718 = 0;
            qword_1C0013720 = 0LL;
            qword_1C0013728 = 0LL;
            if ( !v14 )
              qword_1C00136A8 = 0LL;
            v27 &= 0xFFFE5FFFFFFFFFFFuLL;
            Globals = v27;
            if ( !v37 )
            {
              v27 &= ~0x800000000uLL;
              Globals = v27;
            }
          }
          if ( dword_1C0013A1C )
          {
            v26 = 0xFFFFFFFDFFFFFFFFuLL;
            v27 &= ~0x200000000uLL;
            Globals = v27;
          }
          if ( dword_1C0013A1C != 1 )
          {
            v26 = 0xFFFFBFFFFFFFFFFFuLL;
            Globals = v27 & 0xFFFFBFFFFFFFFFFFuLL;
          }
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          {
            v28 = "Enabled";
            v29 = "Disabled";
            if ( dword_1C0013A1C )
              v29 = "Enabled";
            LOBYTE(v28) = 4;
            WPP_RECORDER_SF_s(
              WPP_GLOBAL_Control->DeviceExtension,
              (_DWORD)v28,
              2,
              25,
              (__int64)&WPP_4e1b20cf9f023c365f1b3d32753808d1_Traceguids,
              (__int64)v29);
            if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
            {
              if ( LOWORD(WPP_GLOBAL_Control->DeviceType) )
              {
                LOBYTE(v26) = 5;
                WPP_RECORDER_SF_(
                  WPP_GLOBAL_Control->DeviceExtension,
                  v26,
                  2,
                  26,
                  (__int64)&WPP_4e1b20cf9f023c365f1b3d32753808d1_Traceguids);
              }
              if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED
                && LOWORD(WPP_GLOBAL_Control->DeviceType) )
              {
                LOBYTE(v26) = 5;
                WPP_RECORDER_SF_(
                  WPP_GLOBAL_Control->DeviceExtension,
                  v26,
                  2,
                  27,
                  (__int64)&WPP_4e1b20cf9f023c365f1b3d32753808d1_Traceguids);
              }
            }
          }
          LOBYTE(v26) = 5;
          DisplayPPMFlags(Globals, v26);
          if ( (v24 & (unsigned int)Globals & 0x7F077LL) != 0 )
            ProcLibTraceIdleStatesErrata(0LL);
          if ( (v24 & (unsigned int)Globals & 0x70000000) != 0 )
            ProcLibTracePerfStatesErrata(0LL);
          if ( (v24 & (unsigned int)Globals & 0x3300000) != 0 )
            ProcLibTraceThrottleStatesErrata(0LL);
          if ( (v24 & (unsigned int)Globals & 0x80000000) != 0 )
            ProcLibTracePccErrata(0LL);
          if ( qword_1C0013488 )
          {
            if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED
              && LOWORD(WPP_GLOBAL_Control->DeviceType) )
            {
              LOBYTE(v30) = 5;
              WPP_RECORDER_SF_(
                WPP_GLOBAL_Control->DeviceExtension,
                v30,
                2,
                28,
                (__int64)&WPP_4e1b20cf9f023c365f1b3d32753808d1_Traceguids);
            }
            LOBYTE(v30) = 5;
            DisplayPPMFlags(~qword_1C0013488, v30);
            Globals &= ~qword_1C0013488;
          }
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          {
            LOBYTE(v30) = 4;
            WPP_RECORDER_SF_(
              WPP_GLOBAL_Control->DeviceExtension,
              v30,
              2,
              29,
              (__int64)&WPP_4e1b20cf9f023c365f1b3d32753808d1_Traceguids);
          }
          LOBYTE(v30) = 4;
          DisplayPPMFlags(Globals, v30);
          if ( qword_1C0013670 && (dword_1C0013490 & 0x70000000) != 0 )
            qword_1C0013670 = 0LL;
          if ( _bittest64(&Globals, 0x23u) )
            HwDebugInitializeRegistryDebugRegisters(0LL);
        }
        else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          v5 = 23;
          goto LABEL_122;
        }
      }
      else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        v5 = 22;
        goto LABEL_122;
      }
    }
  }
  else
  {
    v4 = -1073741670;
    if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    {
      v5 = 20;
      v32 = -102;
      v8 = 3;
LABEL_124:
      LOBYTE(v3) = 2;
      WPP_RECORDER_SF_d(
        WPP_GLOBAL_Control->DeviceExtension,
        v3,
        v8,
        v5,
        (__int64)&WPP_4e1b20cf9f023c365f1b3d32753808d1_Traceguids,
        v32);
    }
  }
  return v4;
}