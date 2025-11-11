NTSTATUS PpmInitIllegalThrottleLogging()
{
  NTSTATUS result; // eax
  ULONG ResultLength; // [rsp+30h] [rbp-40h] BYREF
  HANDLE KeyHandle; // [rsp+38h] [rbp-38h] BYREF
  UNICODE_STRING DestinationString; // [rsp+40h] [rbp-30h] BYREF
  __int128 KeyValueInformation; // [rsp+50h] [rbp-20h] BYREF
  int v5; // [rsp+60h] [rbp-10h]

  ResultLength = 0;
  KeyHandle = 0LL;
  DestinationString = 0LL;
  result = PopOpenKey(&KeyHandle, L"Control\\Session Manager\\Power");
  if ( result >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"ProcessorThrottleLogInterval");
    v5 = 0;
    KeyValueInformation = 0LL;
    if ( ZwQueryValueKey(
           KeyHandle,
           &DestinationString,
           KeyValuePartialInformation,
           &KeyValueInformation,
           0x14u,
           &ResultLength) >= 0
      && *(_QWORD *)((char *)&KeyValueInformation + 4) == 0x400000004LL )
    {
      PopProcessorThrottleLogInterval = HIDWORD(KeyValueInformation);
      if ( HIDWORD(KeyValueInformation) > 0x2710 )
        PopProcessorThrottleLogInterval = 10000;
    }
    return ZwClose(KeyHandle);
  }
  return result;
}