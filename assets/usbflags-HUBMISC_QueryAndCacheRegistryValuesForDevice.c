__int64 __fastcall HUBMISC_QueryAndCacheRegistryValuesForDevice(__int64 a1)
{
  __int64 v1; // rbx
  int UsbflagsValuesForDevice; // edi
  __int64 v4; // rcx
  _BYTE v6[8]; // [rsp+40h] [rbp-28h] BYREF
  _BYTE v7[8]; // [rsp+48h] [rbp-20h] BYREF
  _BYTE v8[8]; // [rsp+50h] [rbp-18h] BYREF

  v1 = a1 + 1996;
  HUBMISC_ConvertUsbDeviceIdsToString(a1 + 1996, v8, v7, v6);
  UsbflagsValuesForDevice = HUBREG_QueryUsbflagsValuesForDevice(a1, v8, v7, v6);
  HUBREG_QueryUsbHardwareVerifierValue(
    v1,
    (__int64)v8,
    (__int64)v7,
    (__int64)v6,
    (__int64)&g_HwVerifierDeviceName,
    *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
    (_DWORD *)(a1 + 2444));
  if ( UsbflagsValuesForDevice < 0 )
  {
    *(_DWORD *)(a1 + 2440) = 1073807366;
    if ( Microsoft_Windows_USB_USBHUB3EnableBits < 0 )
      McTemplateK0pq_EtwWriteTransfer(
        v4,
        &USBHUB3_ETW_EVENT_REGISTRY_FAILURE,
        a1 + 1524,
        *(_QWORD *)(a1 + 24),
        UsbflagsValuesForDevice);
  }
  return ((UsbflagsValuesForDevice >> 31) & 0xFFFFFFF4) + 4077;
}