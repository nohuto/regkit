LSTATUS UmpoRestoreEsOverrideState()
{
  LSTATUS result; // eax
  int pvData; // [rsp+50h] [rbp+8h] BYREF
  DWORD pcbData; // [rsp+58h] [rbp+10h] BYREF
  int v3; // [rsp+60h] [rbp+18h] BYREF

  v3 = 1;
  pvData = 2;
  pcbData = 4;
  result = GetPersistedRegistryLocationW(
             L"Power",
             L"System\\CurrentControlSet\\Control\\Power",
             &UmpoEcoModeRedirectedRegistryPath,
             128LL,
             0LL);
  if ( !result )
  {
    result = RegGetValueW(
               HKEY_LOCAL_MACHINE,
               &UmpoEcoModeRedirectedRegistryPath,
               L"EnergySaverState",
               0x18u,
               0LL,
               &pvData,
               &pcbData);
    if ( !result && pvData == 1 )
      return RtlPublishWnfStateData(WNF_PO_ENERGY_SAVER_OVERRIDE, 0LL, &v3, 4LL, 0LL);
  }
  return result;
}
