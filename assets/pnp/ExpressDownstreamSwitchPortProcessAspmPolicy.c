__int64 __fastcall ExpressDownstreamSwitchPortProcessAspmPolicy(__int64 a1, _BYTE *a2)
{
  __int64 v2; // rax
  _DWORD *v5; // rcx
  struct _DEVICE_OBJECT *v6; // rsi
  void *DeviceExtension; // rcx
  int v9; // [rsp+40h] [rbp+8h] BYREF
  ULONG v10; // [rsp+50h] [rbp+18h] BYREF

  v2 = *(_QWORD *)(a1 + 2552);
  v9 = 0;
  if ( !*(_BYTE *)(v2 + 28) )
  {
    v5 = *(_DWORD **)(a1 + 8);
    if ( (v5[230] & 0x100LL) != 0 )
      goto LABEL_9;
    v10 = 4;
    v6 = *(struct _DEVICE_OBJECT **)(*(_QWORD *)v5 + 280LL);
    if ( IoGetDeviceProperty(v6, DevicePropertyInstallState, 4u, &v9, &v10) < 0 || v9 )
      goto LABEL_9;
    DeviceExtension = v6->DeviceExtension;
    if ( (*(_DWORD *)(*(_QWORD *)(a1 + 8) + 920LL) & 0x200LL) != 0 )
    {
      if ( (unsigned int)PciIsDeviceFeatureEnabled(DeviceExtension, L"ASPMOptOut") != 2 )
        return 0LL;
      goto LABEL_9;
    }
    if ( (unsigned int)PciIsDeviceFeatureEnabled(DeviceExtension, L"ASPMOptIn") != 2 )
    {
LABEL_9:
      *a2 = 1;
      *(_DWORD *)(a1 + 2492) |= 0x40004u;
    }
  }
  return 0LL;
}
