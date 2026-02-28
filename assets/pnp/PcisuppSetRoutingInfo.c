NTSTATUS __fastcall PcisuppSetRoutingInfo(struct _DEVICE_OBJECT *a1, __int64 a2)
{
  NTSTATUS result; // eax
  NTSTATUS v4; // ebx
  __int64 v5; // rcx
  _BYTE *Data; // rdi
  __int64 DataSize; // rax
  struct _UNICODE_STRING DestinationString; // [rsp+30h] [rbp-10h] BYREF
  HANDLE KeyHandle; // [rsp+60h] [rbp+20h] BYREF
  PVOID P; // [rsp+68h] [rbp+28h] BYREF

  KeyHandle = 0LL;
  DestinationString = 0LL;
  result = IrqPolicyGetSubKey(a1, L"Routing Info", 1u, &KeyHandle);
  if ( result >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"Flags");
    v4 = ZwSetValueKey(KeyHandle, &DestinationString, 0, 4u, (PVOID)(a2 + 12), 1u);
    if ( v4 >= 0 )
    {
      if ( *(_QWORD *)a2 )
      {
        v5 = *(_QWORD *)(*(_QWORD *)a2 + 600LL);
        P = 0LL;
        v4 = ACPIAmliBuildObjectPathname(v5, &P, 0LL);
        if ( v4 >= 0 )
        {
          Data = P;
          RtlInitUnicodeString(&DestinationString, L"LinkNode");
          DataSize = -1LL;
          do
            ++DataSize;
          while ( Data[DataSize] );
          v4 = ZwSetValueKey(KeyHandle, &DestinationString, 0, 3u, Data, DataSize);
          if ( Data )
            ExFreePoolWithTag(Data, 0);
        }
      }
      else
      {
        RtlInitUnicodeString(&DestinationString, L"StaticVector");
        v4 = ZwSetValueKey(KeyHandle, &DestinationString, 0, 4u, (PVOID)(a2 + 8), 4u);
      }
    }
    ZwClose(KeyHandle);
    return v4;
  }
  return result;
}