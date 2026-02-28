char __fastcall PciGetDeviceCustomSettings(__int64 a1)
{
  _DWORD *DeviceCustomSetting; // rax
  int v3; // edx
  char v4; // r8
  int v5; // edx
  _DWORD *v6; // rax
  char result; // al
  char v8; // [rsp+38h] [rbp-20h]

  *(_QWORD *)(a1 + 2264) = 0LL;
  DeviceCustomSetting = (_DWORD *)PciGetDeviceCustomSetting(a1, L"DeviceD0DelayTime");
  if ( DeviceCustomSetting )
  {
    *(_QWORD *)(a1 + 2264) |= 1uLL;
    *(_DWORD *)(a1 + 2272) = *DeviceCustomSetting;
    ExFreePoolWithTag(DeviceCustomSetting, 0x42696350u);
    if ( *(_DWORD *)(a1 + 2272) > 0x64u )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && !PciSkipWppTracing )
      {
        v3 = *(_DWORD *)(a1 + 36);
        v4 = v3 & 0x1F;
        v5 = (unsigned __int8)v3 >> 5;
        v8 = v5;
        LOBYTE(v5) = 3;
        WPP_RECORDER_SF_DDDd(
          *(_QWORD *)(*(_QWORD *)(a1 + 144) + 880LL),
          v5,
          6,
          25,
          (__int64)&WPP_caf28d1a72823227dce4fab2e34e6a5e_Traceguids,
          *(_DWORD *)(*(_QWORD *)(a1 + 144) + 292LL),
          v4,
          v8,
          *(_DWORD *)(a1 + 2272));
      }
      *(_QWORD *)(a1 + 2264) &= ~1uLL;
      *(_DWORD *)(a1 + 2272) = 0;
    }
  }
  v6 = (_DWORD *)PciGetDeviceCustomSetting(a1, L"DevicePowerResetDelayTime");
  if ( v6 )
  {
    *(_QWORD *)(a1 + 2264) |= 2uLL;
    *(_DWORD *)(a1 + 2276) = *v6;
    ExFreePoolWithTag(v6, 0x42696350u);
  }
  result = PciGetDeviceDpcCustomSettings(a1);
  if ( *(_BYTE *)(a1 + 52) == 1 )
    return PciGetBridgeCustomSettings(a1);
  return result;
}