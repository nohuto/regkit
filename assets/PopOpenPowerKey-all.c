__int64 __fastcall PopOpenPowerKey(__int64 a1)
{
  return PopOpenKey(a1, L"Control\\Session Manager\\Power");
}

__int64 __fastcall PopApplyPolicy(char a1, char a2, _OWORD *a3, unsigned int a4)
{
  __int64 v8; // r8
  __int128 v9; // xmm1
  __int128 v10; // xmm0
  __int128 v11; // xmm1
  __int128 v12; // xmm0
  __int128 v13; // xmm1
  __int128 v14; // xmm0
  __int128 v15; // xmm0
  _OWORD *v16; // rbx
  __int64 v17; // rax
  __int128 v18; // xmm0
  __int128 v19; // xmm1
  __int128 v20; // xmm0
  __int128 v21; // xmm1
  __int128 v22; // xmm0
  __int64 result; // rax
  int v24; // ebx
  _QWORD *v25; // rdi
  char v26; // r14
  __int64 i; // r8
  __int64 v28; // rcx
  _OWORD *v29; // rcx
  __int128 v30; // xmm1
  __int128 v31; // xmm0
  __int128 v32; // xmm1
  __int128 v33; // xmm0
  __int128 v34; // xmm1
  __int128 v35; // xmm0
  __int128 v36; // xmm1
  __int128 v37; // xmm0
  __int128 v38; // xmm1
  __int128 v39; // xmm0
  __int128 v40; // xmm1
  __int128 v41; // xmm0
  __int128 v42; // xmm1
  __int64 v43; // rax
  __int64 v44; // rcx
  HANDLE DestinationString; // [rsp+38h] [rbp-D0h] BYREF
  UNICODE_STRING DestinationString_8; // [rsp+40h] [rbp-C8h] BYREF
  _OWORD Buf1[6]; // [rsp+58h] [rbp-B0h] BYREF
  __int128 v48; // [rsp+B8h] [rbp-50h]
  _OWORD v49[7]; // [rsp+C8h] [rbp-40h]
  __int64 v50; // [rsp+138h] [rbp+30h]
  _OWORD Data[14]; // [rsp+148h] [rbp+40h] BYREF
  __int64 v52; // [rsp+228h] [rbp+120h]

  memset_0(Buf1, 0, 0xE8uLL);
  DestinationString = 0LL;
  DestinationString_8 = 0LL;
  if ( a4 < 0xE8 )
    return 3221225507LL;
  if ( a4 > 0xE8 )
    return 2147483653LL;
  v9 = a3[1];
  Data[0] = *a3;
  v10 = a3[2];
  Data[1] = v9;
  v11 = a3[3];
  Data[2] = v10;
  v12 = a3[4];
  Data[3] = v11;
  v13 = a3[5];
  Data[4] = v12;
  v14 = a3[6];
  Data[5] = v13;
  Data[6] = v14;
  v15 = a3[7];
  v16 = a3 + 8;
  Data[7] = v15;
  v17 = *((_QWORD *)v16 + 12);
  v18 = v16[1];
  Data[8] = *v16;
  v19 = v16[2];
  Data[9] = v18;
  v20 = v16[3];
  Data[10] = v19;
  v21 = v16[4];
  Data[11] = v20;
  v22 = v16[5];
  Data[12] = v21;
  Data[13] = v22;
  v52 = v17;
  result = PopVerifySystemPowerPolicy(Data, Buf1, v8);
  v24 = result;
  if ( (int)result >= 0 )
  {
    v25 = PopPolicy;
    if ( !memcmp(Buf1, PopPolicy, 0xE8uLL) && !a1 )
    {
      return 0LL;
    }
    else
    {
      v26 = 0;
      for ( i = 0LL; (unsigned int)i < 4; i = (unsigned int)(i + 1) )
      {
        v28 = *((_QWORD *)&v49[-1] + 3 * i) - v25[3 * i + 12];
        if ( !v28 )
        {
          v28 = *((_QWORD *)&v48 + 3 * i + 1) - v25[3 * i + 13];
          if ( !v28 )
            v28 = *((_QWORD *)v49 + 3 * i) - v25[3 * i + 14];
        }
        if ( v28 )
        {
          v26 = 1;
          break;
        }
      }
      v29 = PopPolicy;
      v30 = Buf1[1];
      *(_OWORD *)PopPolicy = Buf1[0];
      v31 = Buf1[2];
      v29[1] = v30;
      v32 = Buf1[3];
      v29[2] = v31;
      v33 = Buf1[4];
      v29[3] = v32;
      v34 = Buf1[5];
      v29[4] = v33;
      v35 = v48;
      v29[5] = v34;
      v36 = v49[0];
      v29[6] = v35;
      v29 += 8;
      v37 = v49[1];
      *(v29 - 1) = v36;
      v38 = v49[2];
      *v29 = v37;
      v39 = v49[3];
      v29[1] = v38;
      v40 = v49[4];
      v29[2] = v39;
      v41 = v49[5];
      v29[3] = v40;
      v42 = v49[6];
      v43 = v50;
      v29[4] = v41;
      v29[5] = v42;
      *((_QWORD *)v29 + 12) = v43;
      PopSetNotificationWork(2LL);
      if ( v26 && !a2 )
      {
        LOBYTE(v44) = -125;
        PopResetCBTriggers(v44);
      }
      PopUpdateSystemIdleContext(3LL);
      if ( a1 )
      {
        v24 = PopOpenPowerKey((__int64)&DestinationString);
        if ( v24 >= 0 )
        {
          RtlInitUnicodeString(&DestinationString_8, L"SystemPowerPolicy");
          v24 = ZwSetValueKey(DestinationString, &DestinationString_8, 0, 3u, Data, 0xE8u);
          ZwClose(DestinationString);
        }
      }
      return (unsigned int)v24;
    }
  }
  return result;
}


__int64 PopResetCurrentPolicies()
{
  __int64 result; // rax
  NTSTATUS v1; // ebx
  ULONG v2; // r9d
  ULONG ResultLength; // [rsp+30h] [rbp-D0h] BYREF
  HANDLE KeyHandle; // [rsp+38h] [rbp-C8h] BYREF
  UNICODE_STRING DestinationString; // [rsp+40h] [rbp-C0h] BYREF
  _BYTE KeyValueInformation[12]; // [rsp+50h] [rbp-B0h] BYREF
  _OWORD v7[15]; // [rsp+5Ch] [rbp-A4h] BYREF

  KeyHandle = 0LL;
  ResultLength = 0;
  DestinationString = 0LL;
  memset_0(KeyValueInformation, 0, 0xF8uLL);
  result = PopOpenPowerKey((__int64)&KeyHandle);
  if ( (int)result >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"SystemPowerPolicy");
    v1 = ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           KeyValueInformation,
           0xF8u,
           &ResultLength);
    if ( v1 >= 0 )
    {
      v2 = ResultLength - 12;
    }
    else
    {
      PopDefaultPolicy(v7);
      v2 = 232;
    }
    ResultLength = v2;
    PopApplyPolicy(0, 0, v7, v2);
    ZwClose(KeyHandle);
    return (unsigned int)v1;
  }
  return result;
}



BOOLEAN __fastcall PopDiagTraceHiberStats(int a1)
{
  int v1; // eax
  HANDLE v2; // r14
  __int64 ResumeTimestamp; // rax
  unsigned __int64 v4; // rbx
  _QWORD *v5; // r12
  union _EVENT_DATA_DESCRIPTOR::$535316677C6A15A6ECBA40D88E1D787B *p_Reserved; // r15
  _BYTE *Data; // r13
  __int64 *v8; // rdi
  __int64 v9; // rax
  char *v10; // rax
  unsigned __int64 v11; // rax
  int v12; // ecx
  int v13; // esi
  ULONG DataSize; // ebx
  BOOLEAN result; // al
  REGHANDLE v16; // rbx
  HANDLE DestinationString[3]; // [rsp+38h] [rbp-D0h] BYREF
  __int64 v18; // [rsp+50h] [rbp-B8h]
  __int64 v19; // [rsp+58h] [rbp-B0h]
  __int64 v20; // [rsp+60h] [rbp-A8h]
  __int128 v21; // [rsp+68h] [rbp-A0h] BYREF
  _BYTE v22[480]; // [rsp+78h] [rbp-90h] BYREF
  struct _EVENT_DATA_DESCRIPTOR UserData; // [rsp+258h] [rbp+150h] BYREF

  LODWORD(v18) = a1;
  v21 = 0LL;
  memset(DestinationString, 0, sizeof(DestinationString));
  v1 = PopOpenPowerKey((__int64)DestinationString);
  v2 = DestinationString[0];
  if ( v1 < 0 )
    v2 = 0LL;
  DestinationString[0] = v2;
  ResumeTimestamp = PopSstDiagQueryResumeTimestamp();
  v4 = qword_140F0AAB8;
  qword_140F0AAC0 = ResumeTimestamp;
  dword_140F0ABC4 = dword_140F0A8A4;
  dword_140F0ABC0 = dword_140F0A938;
  v20 = qword_140F0AAB8;
  LODWORD(qword_140F0AB70) = PopQpcTimeInMs(&qword_140F0A998, &qword_140F0A9A0);
  qword_140F0AA20 = (unsigned int)PopQpcTimeInMs(&qword_140F0A968, &qword_140F0AA28);
  PopComputeDerivedHiberStats(&qword_140F0A9D8, v4, &v21);
  v5 = v22;
  v19 = 59LL;
  p_Reserved = (union _EVENT_DATA_DESCRIPTOR::$535316677C6A15A6ECBA40D88E1D787B *)&UserData.Reserved;
  Data = v22;
  v8 = &qword_1400038D0;
  do
  {
    v9 = *(v8 - 1);
    if ( (*(_DWORD *)v8 & 0x40000000) != 0 )
      v10 = &v22[v9 - 16];
    else
      v10 = (char *)&qword_140F0A9D8 + v9;
    if ( (*(_DWORD *)v8 & 2) != 0 )
      v11 = *(_QWORD *)v10;
    else
      v11 = *(unsigned int *)v10;
    *v5 = v11;
    v12 = *(_DWORD *)v8;
    if ( *(int *)v8 < 0 )
    {
      v11 /= v4;
      *v5 = v11;
    }
    if ( (v12 & 0x10000000) != 0 )
      *v5 = PpmConvertTime(v11, PopQpcFrequency, 1000LL);
    v13 = *(_DWORD *)v8 & 0x20;
    DataSize = v13 != 0 ? 8 : 4;
    if ( v2 )
    {
      RtlInitUnicodeString((PUNICODE_STRING)&DestinationString[1], (PCWSTR)*(v8 - 2));
      ZwSetValueKey(v2, (PUNICODE_STRING)&DestinationString[1], 0, v13 != 0 ? 11 : 4, Data, DataSize);
    }
    *(_QWORD *)&p_Reserved[-3].Reserved = Data;
    p_Reserved[-1].Reserved = DataSize;
    Data += 8;
    v4 = v20;
    v8 += 3;
    p_Reserved->Reserved = 0;
    ++v5;
    p_Reserved += 4;
    --v19;
  }
  while ( v19 );
  qword_140F0AB38 /= v4;
  qword_140F0A9E0 /= v4;
  qword_140F0ABC8 = 1000 * qword_140F0AAC0 / PopQpcFrequency
                  - (unsigned int)qword_140F0AA30
                  - (unsigned int)dword_140F0AA38;
  if ( v2 )
  {
    RtlInitUnicodeString((PUNICODE_STRING)&DestinationString[1], L"KernelResumeIoCpuTime");
    ZwSetValueKey(v2, (PUNICODE_STRING)&DestinationString[1], 0, 4u, &qword_140F0AB38, 4u);
    RtlInitUnicodeString((PUNICODE_STRING)&DestinationString[1], L"HiberIoCpuTime");
    ZwSetValueKey(v2, (PUNICODE_STRING)&DestinationString[1], 0, 4u, &qword_140F0A9E0, 4u);
    if ( qword_140F0AB60 )
    {
      dword_140F0A884 += PopQpcTimeInMs(&qword_140F0A988, &qword_140F0AB68);
      RtlInitUnicodeString((PUNICODE_STRING)&DestinationString[1], L"HybridBootAnimationTime");
      ZwSetValueKey(v2, (PUNICODE_STRING)&DestinationString[1], 0, 4u, &dword_140F0A884, 4u);
    }
    qword_140F0ABD0 = (((unsigned __int64)MEMORY[0xFFFFF78000000004] << 32)
                     * (unsigned __int128)(unsigned __int64)(MEMORY[0xFFFFF78000000320] << 8)) >> 64;
    RtlInitUnicodeString((PUNICODE_STRING)&DestinationString[1], L"ResumeCompleteTimestamp");
    ZwSetValueKey(DestinationString[0], (PUNICODE_STRING)&DestinationString[1], 0, 0xBu, &qword_140F0ABD0, 8u);
    ZwClose(DestinationString[0]);
  }
  result = PopPotsLogHibernatePerformance(&qword_140F0A9D8, (unsigned int)v18);
  if ( PopDiagHandleRegistered )
  {
    v16 = PopDiagHandle;
    result = EtwEventEnabled(PopDiagHandle, &POP_ETW_EVENT_HIBER_STATS);
    if ( result )
      return EtwWrite(v16, &POP_ETW_EVENT_HIBER_STATS, 0LL, 0x3Bu, &UserData);
  }
  return result;
}

NTSTATUS __fastcall PopReadHiberbootPolicy(_BYTE *a1)
{
  char v1; // bl
  NTSTATUS result; // eax
  _BYTE v4[4]; // [rsp+30h] [rbp-40h] BYREF
  ULONG ResultLength; // [rsp+34h] [rbp-3Ch] BYREF
  HANDLE KeyHandle; // [rsp+38h] [rbp-38h] BYREF
  UNICODE_STRING DestinationString; // [rsp+40h] [rbp-30h] BYREF
  __int128 KeyValueInformation; // [rsp+50h] [rbp-20h] BYREF
  int v9; // [rsp+60h] [rbp-10h]

  v1 = 0;
  v4[0] = 0;
  KeyHandle = 0LL;
  ResultLength = 0;
  DestinationString = 0LL;
  result = PopReadHiberbootGroupPolicy(v4);
  if ( result >= 0 && v4[0] )
  {
    v1 = v4[0];
  }
  else
  {
    result = PopOpenPowerKey(&KeyHandle);
    if ( result >= 0 )
    {
      RtlInitUnicodeString(&DestinationString, L"HiberbootEnabled");
      v9 = 0;
      KeyValueInformation = 0LL;
      v1 = 0;
      if ( ZwQueryValueKey(
             KeyHandle,
             &DestinationString,
             KeyValuePartialInformation,
             &KeyValueInformation,
             0x14u,
             &ResultLength) >= 0 )
        v1 = BYTE12(KeyValueInformation);
      result = ZwClose(KeyHandle);
    }
  }
  *a1 = v1;
  return result;
}

int PopReadSystemAwayModePolicy()
{
  bool v0; // bl
  int result; // eax
  ULONG ResultLength; // [rsp+30h] [rbp-40h] BYREF
  HANDLE KeyHandle; // [rsp+38h] [rbp-38h] BYREF
  UNICODE_STRING DestinationString; // [rsp+40h] [rbp-30h] BYREF
  __int128 KeyValueInformation; // [rsp+50h] [rbp-20h] BYREF
  int v6; // [rsp+60h] [rbp-10h]

  KeyHandle = 0LL;
  ResultLength = 0;
  v0 = 0;
  v6 = 0;
  DestinationString = 0LL;
  KeyValueInformation = 0LL;
  if ( byte_140F0B112 )
    v0 = dword_140E01A08 != 0;
  result = PopOpenPowerKey((__int64)&KeyHandle);
  if ( result >= 0 )
  {
    if ( byte_140F0B112 )
    {
      RtlInitUnicodeString(&DestinationString, L"AwayModeEnabled");
      if ( ZwQueryValueKey(
             KeyHandle,
             &DestinationString,
             KeyValuePartialInformation,
             &KeyValueInformation,
             0x14u,
             &ResultLength) >= 0
        && *(_QWORD *)((char *)&KeyValueInformation + 4) == 0x400000004LL
        && HIDWORD(KeyValueInformation) )
      {
        v0 = 1;
      }
    }
    result = ZwClose(KeyHandle);
  }
  byte_140F0B110 = v0;
  return result;
}

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