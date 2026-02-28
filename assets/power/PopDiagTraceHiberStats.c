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
