__int64 CiConfigInitialize()
{
  NTSTATUS v0; // ebx
  unsigned int DWORD; // edx
  __int64 v2; // r9
  unsigned int v3; // eax
  __int64 v4; // r9
  bool v5; // cl
  unsigned int v6; // eax
  __int64 v7; // r9
  __int64 DpcData_high; // r9
  __int64 ActiveThreadCount; // r9
  ULONG v10; // eax
  __int64 v11; // r9
  unsigned int v12; // eax
  __int64 v13; // r9
  unsigned int v14; // eax
  __int64 v15; // r9
  struct _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+20h] [rbp-30h] BYREF
  void *KeyHandle; // [rsp+60h] [rbp+10h] BYREF
  HANDLE Handle; // [rsp+68h] [rbp+18h] BYREF

  *(_QWORD *)&ObjectAttributes.Length = 48LL;
  KeyHandle = 0LL;
  Handle = 0LL;
  ObjectAttributes.RootDirectory = 0LL;
  *(_QWORD *)&ObjectAttributes.Attributes = 64LL;
  ObjectAttributes.ObjectName = (PUNICODE_STRING)0x1C0011110LL; // \Registry\Machine\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile
  *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
  v0 = ZwOpenKey(&KeyHandle, 0x80000100, &ObjectAttributes);
  if ( v0 < 0 )
  {
    if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 2u )
      WPP_SF_d(
        WPP_GLOBAL_Control->AttachedDevice,
        28LL,
        &WPP_350503daac883abe7be9cf63f89038d9_Traceguids,
        (unsigned int)v0);
  }
  else
  {
    DWORD = CiConfigReadDWORD(KeyHandle, 0x1C0011090LL, 100LL); // SystemResponsiveness, missing = 100
    if ( DWORD - 10 > 0x5A ) // data - 10 > 90 = 20
      v2 = 20LL;
    else
      v2 = 10 * (DWORD / 0xA); // 10 step
    CiSystemResponsiveness = v2;
    if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
      WPP_SF_d(WPP_GLOBAL_Control->AttachedDevice, 18LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids, v2);
    if ( CiSystemResponsiveness == 100 )
    {
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 2u )
        WPP_SF_(WPP_GLOBAL_Control->AttachedDevice, 19LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids);
      v0 = -1073741696; // STATUS_SERVER_DISABLED
    }
    else
    {
      v3 = CiConfigReadDWORD(KeyHandle, 0x1C00110A0LL, 10LL); // NetworkThrottlingIndex, missing = 10
      LODWORD(WPP_MAIN_CB.Dpc.DpcData) = v3;
      v4 = v3;
      if ( v3 )
      {
        if ( v3 - 71 <= 0xFFFFFFB7 ) // >=71 = 70
        {
          v4 = 70LL;
          LODWORD(WPP_MAIN_CB.Dpc.DpcData) = 70; // range 1-70
        }
      }
      else
      {
        v4 = 1LL; // zero -> force 1
        LODWORD(WPP_MAIN_CB.Dpc.DpcData) = 1;
      }
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(WPP_GLOBAL_Control->AttachedDevice, 20LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids, v4);
      v5 = (unsigned __int8)CiConfigReadDWORD(KeyHandle, 0x1C0011080LL, 0LL) != 0; // NoLazyMode, missing = 0, nonzero = true
      CiSchedulerDisallowLazyMode = v5;
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(WPP_GLOBAL_Control->AttachedDevice, 21LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids, v5);
      v6 = CiConfigReadDWORD(KeyHandle, 0x1C00110B0LL, 2LL); // IdleDetectionCycles, missing = 2
      v7 = v6;
      CiSchedulerIdleDetectionCycles = v6;
      if ( v6 - 1 > 0x1E ) // data - 1 > 30 = 2
      {
        v7 = 2LL;
        CiSchedulerIdleDetectionCycles = 2;
      }
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(WPP_GLOBAL_Control->AttachedDevice, 22LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids, v7);
      CiSchedulerIdleCycleBitMask = (1 << CiSchedulerIdleDetectionCycles) - 1;
      HIDWORD(WPP_MAIN_CB.Dpc.DpcData) = CiConfigReadDWORD(KeyHandle, 0x1C00110C0LL, 1000000LL); // LazyModeTimeout, missing = 1000000
      DpcData_high = HIDWORD(WPP_MAIN_CB.Dpc.DpcData);
      if ( !HIDWORD(WPP_MAIN_CB.Dpc.DpcData) )
      {
        DpcData_high = 1000000LL; // 0 = 1000000
        HIDWORD(WPP_MAIN_CB.Dpc.DpcData) = 1000000;
      }
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(
          WPP_GLOBAL_Control->AttachedDevice,
          23LL,
          &WPP_350503daac883abe7be9cf63f89038d9_Traceguids,
          DpcData_high);
      WPP_MAIN_CB.ActiveThreadCount = CiConfigReadDWORD(KeyHandle, 0x1C00110D0LL, 10000LL); // SchedulerTimerResolution, missing = 10000
      ActiveThreadCount = WPP_MAIN_CB.ActiveThreadCount;
      if ( WPP_MAIN_CB.ActiveThreadCount > 0x2710 ) // >10000 = 10000
      {
        ActiveThreadCount = 10000LL;
        WPP_MAIN_CB.ActiveThreadCount = 10000;
      }
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(
          WPP_GLOBAL_Control->AttachedDevice,
          24LL,
          &WPP_350503daac883abe7be9cf63f89038d9_Traceguids,
          ActiveThreadCount);
      v10 = CiConfigReadDWORD(KeyHandle, 0x1C00110E0LL, 100000LL); // SchedulerPeriod, missing = 100000
      v11 = v10;
      *(&WPP_MAIN_CB.ActiveThreadCount + 1) = v10;
      if ( v10 - 50000 > 0xE7EF0 ) // data - 50000 > 950000 = 100000
      {
        v11 = 100000LL;
        *(&WPP_MAIN_CB.ActiveThreadCount + 1) = 100000;
      }
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(WPP_GLOBAL_Control->AttachedDevice, 25LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids, v11);
      v12 = CiConfigReadDWORD(KeyHandle, 0x1C00110F0LL, 32LL); // MaxThreadsPerProcess, missing = 32
      v13 = v12;
      CiMaxThreadsPerProcess = v12;
      if ( v12 - 8 > 0x78 ) // data - 8 > 120 = 32
      {
        v13 = 32LL;
        CiMaxThreadsPerProcess = 32;
      }
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(WPP_GLOBAL_Control->AttachedDevice, 26LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids, v13);
      v14 = CiConfigReadDWORD(KeyHandle, 0x1C0011100LL, 256LL); // MaxThreadsTotal, missing = 256
      v15 = v14;
      CiMaxThreadsTotal = v14;
      if ( v14 - 64 > 0xFFBF ) // data - 64 > 65471 = 256
      {
        v15 = 256LL;
        CiMaxThreadsTotal = 256;
      }
      if ( (HIDWORD(WPP_GLOBAL_Control->Timer) & 1) != 0 && BYTE1(WPP_GLOBAL_Control->Timer) >= 4u )
        WPP_SF_d(WPP_GLOBAL_Control->AttachedDevice, 27LL, &WPP_350503daac883abe7be9cf63f89038d9_Traceguids, v15);
      ObjectAttributes.RootDirectory = KeyHandle;
      ObjectAttributes.Length = 48;
      ObjectAttributes.ObjectName = (PUNICODE_STRING)0x1C0011120LL; // Tasks
      ObjectAttributes.Attributes = 64;
      *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
      v0 = ZwOpenKey(&Handle, 0x80000100, &ObjectAttributes);
      if ( v0 >= 0 )
      {
        v0 = CiConfigInitializeFromRegistry(Handle);
        ZwClose(Handle);
      }
    }
    ZwClose(KeyHandle);
  }
  return (unsigned int)v0;
}
