int __fastcall UsbhGetD3Policy(PDEVICE_OBJECT DeviceObject)
{
  __int64 v2; // rax
  __int64 v3; // rax
  NTSTATUS v4; // ebx
  NTSTATUS v5; // ebx
  ULONG ResultLength; // [rsp+30h] [rbp-29h] BYREF
  void *DeviceRegKey; // [rsp+38h] [rbp-21h] BYREF
  void *KeyHandle; // [rsp+40h] [rbp-19h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+48h] [rbp-11h] BYREF
  _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+58h] [rbp-1h] BYREF
  __int128 KeyValueInformation; // [rsp+88h] [rbp+2Fh] BYREF
  int v13; // [rsp+98h] [rbp+3Fh]

  *(&ObjectAttributes.Attributes + 1) = 0;
  KeyHandle = 0LL;
  DeviceRegKey = 0LL;
  ResultLength = 0;
  *(&ObjectAttributes.Length + 1) = 0;
  DestinationString = 0LL;
  v13 = 0;
  KeyValueInformation = 0LL;
  v2 = ((__int64 (*)(void))PdoExt)();
  *(_DWORD *)(v2 + 1420) &= ~0x400000u;
  LODWORD(v3) = IoOpenDeviceRegistryKey(DeviceObject, 1u, 0xF003Fu, &DeviceRegKey);
  if ( (int)v3 >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"e5b3b5ac-9725-4f78-963f-03dfb1d828c7");
    ObjectAttributes.RootDirectory = DeviceRegKey;
    ObjectAttributes.Length = 48;
    ObjectAttributes.ObjectName = &DestinationString;
    ObjectAttributes.Attributes = 576;
    *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
    v4 = ZwOpenKey(&KeyHandle, 0xF003Fu, &ObjectAttributes);
    LODWORD(v3) = ZwClose(DeviceRegKey);
    if ( v4 >= 0 )
    {
      RtlInitUnicodeString(&DestinationString, L"D3ColdSupported");
      v5 = ZwQueryValueKey(
             KeyHandle,
             &DestinationString,
             KeyValuePartialInformation,
             &KeyValueInformation,
             0x14u,
             &ResultLength);
      LODWORD(v3) = ZwClose(KeyHandle);
      if ( v5 >= 0 )
      {
        if ( HIDWORD(KeyValueInformation) )
        {
          v3 = PdoExt(DeviceObject);
          *(_DWORD *)(v3 + 1420) |= 0x400000u;
        }
      }
    }
  }
  return v3;
}