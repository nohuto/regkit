__int64 Registry_InitializeAllow64KLowOrFullSpeedControlTransfersFlag()
{
  int v0; // ebx
  __int64 v1; // rdx
  int v3; // [rsp+50h] [rbp+8h] BYREF
  __int64 v4; // [rsp+58h] [rbp+10h] BYREF

  v4 = 0LL;
  v3 = 0;
  v0 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, _QWORD, const wchar_t *, __int64, _QWORD, __int64 *))(WdfFunctions_01015 + 1832))(
         WdfDriverGlobals,
         0LL,
         L"vx", // \Registry\Machine\System\CurrentControlSet\Control\usbflags
         131097LL,
         0LL,
         &v4);
  if ( v0 >= 0 )
  {
    v0 = (*(__int64 (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, __int64, int *, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           v4,
           L"LN", // Allow64KLowOrFullSpeedControlTransfers
           4LL,
           &v3,
           0LL,
           0LL);
    v1 = v4;
    if ( v0 >= 0 )
      g_Allow64KLowOrFullSpeedControlTransfersFlag = v3 == 1;
  }
  else
  {
    v1 = 0LL;
    v4 = 0LL;
  }
  if ( v1 )
    (*(void (__fastcall **)(PWDF_DRIVER_GLOBALS))(WdfFunctions_01015 + 1848))(WdfDriverGlobals);
  return (unsigned int)v0;
}

// .data:000000014002DF88 g_Allow64KLowOrFullSpeedControlTransfersFlag db 0
// ..data:000000014002DF88                                         ; DATA XREF: Endpoint_CalculateMaximumTransferSize+AE↑r
// ..data:000000014002DF88                                         ; DriverEntry+481↑r ...