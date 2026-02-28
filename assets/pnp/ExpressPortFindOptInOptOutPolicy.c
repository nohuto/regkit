char __fastcall ExpressPortFindOptInOptOutPolicy(__int64 a1)
{
  __int64 v1; // rbx
  int v5; // [rsp+50h] [rbp+8h] BYREF
  ULONG v6; // [rsp+58h] [rbp+10h] BYREF

  v1 = *(_QWORD *)(a1 + 8);
  v5 = 0;
  v6 = 4;
  if ( IoGetDeviceProperty(*(PDEVICE_OBJECT *)(v1 + 128), DevicePropertyInstallState, 4u, &v5, &v6) < 0 || v5 )
    return 1;
  if ( (*(_DWORD *)(a1 + 96) & 1) != 0 )
    return (unsigned int)PciIsDeviceFeatureEnabled(v1, L"ASPMOptOut") == 2;
  else
    return (unsigned int)PciIsDeviceFeatureEnabled(v1, L"ASPMOptIn") != 2;
}
