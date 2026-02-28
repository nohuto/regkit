__int64 __fastcall IrqPolicySetDeviceAffinity(__int64 a1, unsigned __int16 *a2)
{
  __int64 result; // rax
  unsigned int v4; // ebx
  struct _UNICODE_STRING DestinationString; // [rsp+30h] [rbp-10h] BYREF
  int Data; // [rsp+60h] [rbp+20h] BYREF
  HANDLE KeyHandle; // [rsp+68h] [rbp+28h] BYREF

  KeyHandle = 0LL;
  DestinationString = 0LL;
  result = IrqPolicyGetSubKey(a1, L"Affinity Policy - Temporal", 1LL, &KeyHandle);
  if ( (int)result >= 0 )
  {
    Data = a2[4];
    RtlInitUnicodeString(&DestinationString, L"TargetGroup");
    ZwSetValueKey(KeyHandle, &DestinationString, 0, 4u, &Data, 4u);
    RtlInitUnicodeString(&DestinationString, L"TargetSet");
    v4 = ZwSetValueKey(KeyHandle, &DestinationString, 0, 0xBu, a2, 8u);
    ZwClose(KeyHandle);
    return v4;
  }
  return result;
}