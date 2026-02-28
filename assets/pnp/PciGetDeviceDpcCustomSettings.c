char __fastcall PciGetDeviceDpcCustomSettings(__int64 a1)
{
  unsigned int *DeviceCustomSetting; // rax
  __int64 v3; // r9
  int v4; // edx
  unsigned __int64 v5; // rax
  __int64 v6; // r8
  int v7; // edx

  DeviceCustomSetting = (unsigned int *)PciGetDeviceCustomSetting(a1, L"DeviceDpcResetActionOverride");
  v3 = *(_QWORD *)(a1 + 2264);
  *(_QWORD *)(a1 + 2264) = v3 & 0xFFFFFFFFFFFFE01FuLL;
  if ( DeviceCustomSetting )
  {
    *(_QWORD *)(a1 + 2264) = (32LL * *DeviceCustomSetting) ^ (v3 ^ (32LL * *DeviceCustomSetting)) & 0xFFFFFFFFFFFFE01FuLL;
    ExFreePoolWithTag(DeviceCustomSetting, 0x42696350u);
    if ( (*(_DWORD *)(a1 + 2264) & 0x1FE0u) >= 0xA0uLL )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && !PciSkipWppTracing )
      {
        v4 = (unsigned __int8)((unsigned __int64)*(unsigned int *)(a1 + 2264) >> 5);
        LOBYTE(v4) = 3;
        WPP_RECORDER_SF_DDDi(
          *(_QWORD *)(*(_QWORD *)(a1 + 144) + 880LL),
          v4,
          2,
          27,
          (__int64)&WPP_caf28d1a72823227dce4fab2e34e6a5e_Traceguids,
          *(_DWORD *)(*(_QWORD *)(a1 + 144) + 292LL),
          *(_DWORD *)(a1 + 36) & 0x1F,
          (unsigned __int8)*(_DWORD *)(a1 + 36) >> 5,
          (unsigned __int64)*(unsigned int *)(a1 + 2264) >> 5);
      }
      *(_QWORD *)(a1 + 2264) &= 0xFFFFFFFFFFFFE01FuLL;
    }
  }
  v5 = PciGetDeviceCustomSetting(a1, L"DeviceDpcCleanUpActionOverride");
  v6 = *(_QWORD *)(a1 + 2264);
  *(_QWORD *)(a1 + 2264) = v6 & 0xFFFFFFFFFFE01FFFuLL;
  if ( v5 )
  {
    *(_QWORD *)(a1 + 2264) = ((unsigned __int64)*(unsigned int *)v5 << 13) ^ (v6 ^ ((unsigned __int64)*(unsigned int *)v5 << 13)) & 0xFFFFFFFFFFE01FFFuLL;
    ExFreePoolWithTag((PVOID)v5, 0x42696350u);
    v5 = *(_DWORD *)(a1 + 2264) & 0x1FE000;
    if ( v5 >= 0x4000 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v5) = PciSkipWppTracing;
        if ( !PciSkipWppTracing )
        {
          LOBYTE(v7) = 3;
          LOBYTE(v5) = WPP_RECORDER_SF_DDDi(
                         *(_QWORD *)(*(_QWORD *)(a1 + 144) + 880LL),
                         v7,
                         2,
                         28,
                         (__int64)&WPP_caf28d1a72823227dce4fab2e34e6a5e_Traceguids,
                         *(_DWORD *)(*(_QWORD *)(a1 + 144) + 292LL),
                         *(_DWORD *)(a1 + 36) & 0x1F,
                         (*(_DWORD *)(a1 + 36) >> 5) & 7,
                         (unsigned __int64)*(unsigned int *)(a1 + 2264) >> 13);
        }
      }
      *(_QWORD *)(a1 + 2264) &= 0xFFFFFFFFFFE01FFFuLL;
    }
  }
  return v5;
}