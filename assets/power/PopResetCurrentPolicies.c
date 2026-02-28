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
