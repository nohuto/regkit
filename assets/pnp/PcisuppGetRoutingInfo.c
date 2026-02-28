NTSTATUS __fastcall PcisuppGetRoutingInfo(struct _DEVICE_OBJECT *a1, _OWORD *a2)
{
  NTSTATUS result; // eax
  int v4; // edi
  ULONG ResultLength; // [rsp+38h] [rbp-81h] BYREF
  HANDLE KeyHandle; // [rsp+40h] [rbp-79h] BYREF
  __int128 v7; // [rsp+48h] [rbp-71h]
  __int64 v8; // [rsp+58h] [rbp-61h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+60h] [rbp-59h] BYREF
  _BYTE KeyValueInformation[4]; // [rsp+70h] [rbp-49h] BYREF
  int v11; // [rsp+74h] [rbp-45h]
  unsigned int v12; // [rsp+78h] [rbp-41h]
  _BYTE v13[116]; // [rsp+7Ch] [rbp-3Dh] BYREF

  KeyHandle = 0LL;
  ResultLength = 0;
  v8 = 0LL;
  DestinationString = 0LL;
  v7 = 0LL;
  result = IrqPolicyGetSubKey(a1, L"Routing Info", 1u, &KeyHandle);
  if ( result >= 0 )
  {
    if ( (int)OSGetRegistryValue(KeyHandle) >= 0 )
    {
      if ( MEMORY[4] && MEMORY[0] == 4 )
        BYTE12(v7) = MEMORY[8];
      ExFreePoolWithTag(0LL, 0);
    }
    RtlInitUnicodeString(&DestinationString, L"LinkNode");
    if ( ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           KeyValueInformation,
           0x78u,
           &ResultLength) >= 0
      && v11 == 3
      && ResultLength < 0x78
      && v12 + 12 < 0x78 )
    {
      if ( v12 > 0x6B )
      {
        v4 = -1073741789;
LABEL_20:
        ZwClose(KeyHandle);
        return v4;
      }
      v13[v12] = 0;
      v4 = LinkNodeFindByName(v13, &v8);
      if ( v4 < 0 )
        goto LABEL_20;
      DWORD2(v7) = 0;
      *(_QWORD *)&v7 = v8;
    }
    else
    {
      *(_QWORD *)&v7 = 0LL;
      v4 = OSGetRegistryValue(KeyHandle);
      if ( v4 < 0 )
        goto LABEL_20;
      if ( MEMORY[4] )
      {
        if ( MEMORY[0] == 4 )
          DWORD2(v7) = MEMORY[8];
      }
    }
    *a2 = v7;
    goto LABEL_20;
  }
  return result;
}