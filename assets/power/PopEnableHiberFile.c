__int64 __fastcall PopEnableHiberFile(char a1)
{
  int v1; // ebx
  PVOID v3; // r15
  __int64 v4; // rsi
  PVOID v5; // r12
  unsigned int v6; // r14d
  int v7; // ecx
  bool v8; // zf
  char v9; // r11
  char v10; // di
  HANDLE v11; // rbx
  __int64 DumpHibernateResources; // rax
  unsigned __int64 v13; // rdi
  int HiberFile; // eax
  void *Pool2; // rax
  char v17[8]; // [rsp+38h] [rbp-39h] BYREF
  __int64 v18; // [rsp+40h] [rbp-31h] BYREF
  ULONG ResultLength; // [rsp+48h] [rbp-29h] BYREF
  HANDLE KeyHandle; // [rsp+50h] [rbp-21h] BYREF
  UNICODE_STRING DestinationString; // [rsp+58h] [rbp-19h] BYREF
  UNICODE_STRING ValueName; // [rsp+68h] [rbp-9h] BYREF
  UNICODE_STRING v23; // [rsp+78h] [rbp+7h] BYREF
  __int128 KeyValueInformation; // [rsp+88h] [rbp+17h] BYREF
  int v25; // [rsp+98h] [rbp+27h]

  v1 = 0;
  v18 = 0LL;
  v17[0] = 0;
  ResultLength = 0;
  v25 = 0;
  DestinationString = 0LL;
  KeyHandle = 0LL;
  v3 = 0LL;
  v23 = 0LL;
  v4 = 0LL;
  v5 = 0LL;
  KeyValueInformation = 0LL;
  ValueName = 0LL;
  PopRemoveReasonRecordByReasonCode(6LL);
  PopRemoveReasonRecordByReasonCode(8LL);
  PopRemoveReasonRecordByReasonCode(22LL);
  PopRemoveReasonRecordByReasonCode(23LL);
  PopRemoveReasonRecordByReasonCode(24LL);
  PopRemoveReasonRecordByReasonCode(25LL);
  v6 = (unsigned __int64)MmGetHighestPhysicalPage(0LL) >= 0x100000000LL ? 8 : 0;
  if ( !(unsigned __int8)PopCheckDisabledReason(2LL) && !(unsigned __int8)PopCheckDisabledReason(1LL) )
    PopCheckDisabledReason(15LL);
  if ( (unsigned __int8)PopCheckDisabledReason(16LL) )
  {
    v7 = -1073741637;
    v1 = -1073741637;
    LODWORD(v18) = -1073741637;
    goto LABEL_13;
  }
  v8 = (unsigned __int8)PopCheckDisabledReason(13LL) == 0;
  v10 = v9;
  if ( !v8 )
    v10 = 1;
  if ( a1 )
  {
    if ( FileObject )
      goto LABEL_56;
    dword_140F0A8A4 = 1;
    dword_140F0A884 = 1601;
    byte_140F0A8A1 = 0;
    dword_140F0A938 = 0;
    if ( (int)PopOpenPowerKey((__int64)&KeyHandle) >= 0 )
    {
      RtlInitUnicodeString(&DestinationString, L"MaxHuffRatio");
      v11 = KeyHandle;
      if ( ZwQueryValueKey(
             KeyHandle,
             &DestinationString,
             KeyValuePartialInformation,
             &KeyValueInformation,
             0x14u,
             &ResultLength) >= 0
        && *(_QWORD *)((char *)&KeyValueInformation + 4) == 0x400000004LL )
      {
        dword_140F0A8A4 = HIDWORD(KeyValueInformation);
        if ( (unsigned int)(HIDWORD(KeyValueInformation) - 1) > 0x62 )
          dword_140F0A8A4 = 1;
      }
      RtlInitUnicodeString(&ValueName, L"HybridBootAnimationTime");
      v25 = 0;
      KeyValueInformation = 0LL;
      if ( ZwQueryValueKey(v11, &ValueName, KeyValuePartialInformation, &KeyValueInformation, 0x14u, &ResultLength) >= 0
        && *(_QWORD *)((char *)&KeyValueInformation + 4) == 0x400000004LL )
      {
        dword_140F0A884 = HIDWORD(KeyValueInformation);
      }
      RtlInitUnicodeString(&v23, L"MultiPhaseResumeDisabled");
      v25 = 0;
      KeyValueInformation = 0LL;
      if ( ZwQueryValueKey(v11, &v23, KeyValuePartialInformation, &KeyValueInformation, 0x14u, &ResultLength) >= 0
        && *(_QWORD *)((char *)&KeyValueInformation + 4) == 0x400000004LL )
      {
        byte_140F0A8A1 = HIDWORD(KeyValueInformation) == 1;
        dword_140F0A938 |= 0x20u;
      }
      ZwClose(v11);
    }
    PopHiberEnabled = 1;
    if ( v10 )
    {
      v7 = -1073741637;
LABEL_41:
      v1 = v7;
      LODWORD(v18) = v7;
      goto LABEL_45;
    }
    DumpHibernateResources = MmAllocateDumpHibernateResources(77824LL);
    v4 = DumpHibernateResources;
    if ( !DumpHibernateResources )
    {
      v6 = 23;
LABEL_40:
      v7 = -1073741670;
      goto LABEL_41;
    }
    v13 = DumpHibernateResources + 0x200000;
    if ( (DumpHibernateResources & 0x1FFFFF) != 0 )
      v13 = (DumpHibernateResources + 0x1FFFFF) & 0xFFFFFFFFFFE00000uLL;
    if ( v13 - DumpHibernateResources >= 0xA000 )
      v13 = DumpHibernateResources;
    PopCalculateHiberFileSize(&v18, v17);
    HiberFile = PopCreateHiberFile(v18);
    LODWORD(v18) = HiberFile;
    v1 = HiberFile;
    if ( HiberFile >= 0 )
    {
      *(_QWORD *)&xmmword_140F0A888 = v4;
      v4 = 0LL;
      *((_QWORD *)&xmmword_140F0A888 + 1) = v13;
      Pool2 = (void *)ExAllocatePool2(0x40uLL);
      v3 = Pool2;
      if ( !Pool2 )
      {
        v6 = 24;
        goto LABEL_40;
      }
      MemoryMap = Pool2;
      v3 = 0LL;
      HiberFile = PopPreallocateHibernateMemory();
      LODWORD(v18) = HiberFile;
      v1 = HiberFile;
      if ( HiberFile >= 0 )
      {
        LODWORD(v18) = 1;
        EmClientQueryRuleState(EM_RULE_DISABLE_MULTI_PHASE_RESUME, &v18);
        if ( (_DWORD)v18 == 2 )
        {
          dword_140F0A938 |= 0x10u;
          byte_140F0A8A1 = 1;
        }
        byte_140F0B196 = v17[0];
        byte_140F0B188 = 1;
        if ( !(_DWORD)InitSafeBootMode )
          byte_140F0B192 = 1;
        if ( (BYTE8(PopBsdPowerTransitionAtBoot) & 1) == 0 )
          PopClearHiberFileSignature();
        if ( (unsigned int)Feature_Servicing_HiberResetPolicyFix__private_IsEnabledDeviceUsageNoInline() )
          PopResetCurrentPolicies();
        v1 = 0;
        goto LABEL_56;
      }
      v6 = 25;
    }
    else
    {
      v6 = 6;
    }
    v7 = HiberFile;
LABEL_45:
    if ( (PopSimulateHiberBugcheck & 0x800) != 0 )
      KeBugCheckEx(0xA0u, 9uLL, v7, 0xFFFFFFFFFFFFFFFFuLL, v6);
    goto LABEL_57;
  }
  PopHiberEnabled = 0;
  if ( !FileObject )
  {
LABEL_56:
    LODWORD(v18) = 0;
    goto LABEL_57;
  }
  if ( (unsigned int)MmZeroPageFileAtShutdown() )
    PopZeroHiberFile(PopHiberInfo);
  ObfDereferenceObjectWithTag(FileObject, 0x62486F50u);
  ZwClose(PopHiberInfo);
  ExFreePoolWithTag(qword_140F0A878, 0x72626968u);
  v4 = xmmword_140F0A888;
  v5 = qword_140F0A8B0;
  v3 = MemoryMap;
  xmmword_140F0A888 = 0LL;
  byte_140F0B188 = 0;
  byte_140F0B196 = 0;
  byte_140F0B192 = 0;
  qword_140F0A898 = 0LL;
  v1 = PopResetCurrentPolicies();
  LODWORD(v18) = v1;
  v7 = v1;
  if ( v1 < 0 )
  {
LABEL_13:
    if ( !a1 )
      goto LABEL_57;
    goto LABEL_45;
  }
LABEL_57:
  if ( v6 )
  {
    PopLogSleepDisabled(v6, 8LL, &v18, 4LL);
    v1 = v18;
  }
  if ( v4 )
    MmReleaseDumpHibernateResources(v4, 77824LL);
  if ( v5 )
  {
    MmReturnChargesToLockPagedPool(v5, Length);
    ExFreePoolWithTag(v5, 0);
    memset_0(&qword_140F0A8B0, 0, 0x88uLL);
  }
  if ( v3 )
  {
    ExFreePoolWithTag(v3, 0x70616D48u);
    MemoryMap = 0LL;
  }
  if ( !a1 )
    memset_0(&PopHiberInfo, 0, 0xE8uLL);
  return (unsigned int)v1;
}
