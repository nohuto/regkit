void __fastcall HidpToggleRemoteWakeWorker(PDEVICE_OBJECT DeviceObject, _BYTE *Context)
{
  char v2; // r14
  PVOID v3; // rsi
  IRP *v4; // rdi
  char v5; // r15
  unsigned __int8 v6; // bp
  __int64 v7; // rbx
  KIRQL v8; // al
  char v9; // dl
  struct _DEVICE_OBJECT *v10; // rcx
  int v11; // edx
  struct _UNICODE_STRING DestinationString; // [rsp+30h] [rbp-48h] BYREF
  int Data; // [rsp+88h] [rbp+10h] BYREF
  void *DeviceRegKey; // [rsp+90h] [rbp+18h] BYREF

  v2 = Context[24];
  v3 = Context;
  v4 = 0LL;
  v5 = 0;
  v6 = v2 != 0;
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(Context) = 5;
    WPP_RECORDER_SF_(
      WPP_GLOBAL_Control->DeviceExtension,
      (_DWORD)Context,
      7,
      20,
      (__int64)&WPP_f33e33053943332cc47e0fb2468cff69_Traceguids);
  }
  v7 = *((_QWORD *)v3 + 1);
  v8 = KeAcquireSpinLockRaiseToDpc((PKSPIN_LOCK)(v7 + 88));
  v9 = *(_BYTE *)(v7 + 84);
  if ( v6 != v9 )
  {
    v5 = 1;
    if ( v9 )
      v4 = (IRP *)_InterlockedExchange64((volatile __int64 *)(v7 + 96), 0LL);
    *(_BYTE *)(v7 + 84) = v6;
  }
  KeReleaseSpinLock((PKSPIN_LOCK)(v7 + 88), v8);
  if ( v5 )
  {
    v10 = *(struct _DEVICE_OBJECT **)(v7 + 48);
    Data = v6;
    DeviceRegKey = 0LL;
    DestinationString = 0LL;
    if ( IoOpenDeviceRegistryKey(v10, 1u, 0x1F0000u, &DeviceRegKey) >= 0 )
    {
      RtlInitUnicodeString(&DestinationString, L"RemoteWakeEnabled");
      ZwSetValueKey(DeviceRegKey, &DestinationString, 0, 4u, &Data, 4u);
      ZwClose(DeviceRegKey);
    }
    if ( v2 )
      HidpCreateRemoteWakeIrp((PVOID)v7);
  }
  if ( v4 )
    IoCancelIrp(v4);
  IoReleaseRemoveLockEx((PIO_REMOVE_LOCK)(v7 + 16), (PVOID)0x54575752, 0x20u);
  IoFreeWorkItem(*((PIO_WORKITEM *)v3 + 2));
  ExFreePoolWithTag(v3, 0);
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    if ( LOWORD(WPP_GLOBAL_Control->DeviceType) )
    {
      LOBYTE(v11) = 5;
      WPP_RECORDER_SF_(
        WPP_GLOBAL_Control->DeviceExtension,
        v11,
        7,
        21,
        (__int64)&WPP_f33e33053943332cc47e0fb2468cff69_Traceguids);
    }
  }
}