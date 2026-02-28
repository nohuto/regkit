__int64 RegQueryGenericCompositeUSBDeviceString()
{
  __int64 (__fastcall *SystemRoutineAddress)(__int64, const wchar_t *, void **); // rax
  struct _UNICODE_STRING DestinationString; // [rsp+30h] [rbp-29h] BYREF
  void *v3; // [rsp+40h] [rbp-19h] BYREF
  int v4; // [rsp+48h] [rbp-11h]
  const wchar_t *v5; // [rsp+50h] [rbp-9h]
  PVOID *v6; // [rsp+58h] [rbp-1h]
  int v7; // [rsp+60h] [rbp+7h]
  __int64 v8; // [rsp+68h] [rbp+Fh]
  int v9; // [rsp+70h] [rbp+17h]
  __int64 v10; // [rsp+78h] [rbp+1Fh]
  int v11; // [rsp+80h] [rbp+27h]
  __int64 v12; // [rsp+88h] [rbp+2Fh]

  v4 = 4;
  v3 = &GetConfigValue;
  v7 = 0;
  v5 = L"GenericCompositeUSBDeviceString\n";
  v8 = 0LL;
  v6 = &GenericCompositeUSBDeviceString;
  v9 = 0;
  v10 = 0LL;
  v11 = 0;
  v12 = 0LL;
  DestinationString = 0LL;
  RtlInitUnicodeString(&DestinationString, L"RtlQueryRegistryValuesEx");
  SystemRoutineAddress = (__int64 (__fastcall *)(__int64, const wchar_t *, void **))MmGetSystemRoutineAddress(&DestinationString);
  if ( !SystemRoutineAddress )
    SystemRoutineAddress = (__int64 (__fastcall *)(__int64, const wchar_t *, void **))RtlQueryRegistryValues;
  return SystemRoutineAddress(2LL, L"usbflags", &v3);
}