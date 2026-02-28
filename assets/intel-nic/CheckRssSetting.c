__int64 __fastcall CheckRssSetting(struct _MP_PORT *a1)
{
  __int64 v1; // rbx
  char CurrentIrql; // bl
  _DWORD *DfltActivityPtr; // rax
  int v5; // esi
  char v6; // bl
  _DWORD *v7; // rax
  char v9; // bl
  _DWORD *v10; // rax
  char v11; // bl
  _DWORD *v12; // rax
  _QWORD StringsList[2]; // [rsp+40h] [rbp-28h] BYREF
  int v14; // [rsp+70h] [rbp+8h] BYREF

  v1 = *((_QWORD *)a1 + 147) + 21720LL;
  if ( *((_DWORD *)a1 + 1030) != 1 )
    return 0LL;
  if ( (*((unsigned __int8 (__fastcall **)(_QWORD))a1 + 235))(*((_QWORD *)a1 + 148)) )
  {
    StringsList[0] = v1;
    StringsList[1] = v1;
    NdisWriteEventLogEntry_0(g_pDriverObject, -2147024862, 0, 2u, StringsList, 0, 0LL);
    if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
      && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
      && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
    {
      CurrentIrql = KeGetCurrentIrql();
      DfltActivityPtr = (_DWORD *)TraceGetDfltActivityPtr();
      WPP_SF_sDD(
        WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
        80,
        (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
        (unsigned int)"IPOIB",
        CurrentIrql,
        *DfltActivityPtr);
    }
  }
  v14 = 0;
  v5 = ReadRegistryDword(
         L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Services\\Tcpip",
         L"\\Parameters",
         L"EnableRSS",
         1u,
         &v14);
  if ( v5 < 0 )
  {
    _mm_lfence();
    if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
      && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
      && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
    {
      v6 = KeGetCurrentIrql();
      v7 = (_DWORD *)TraceGetDfltActivityPtr();
      WPP_SF_sDDd(
        WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
        81,
        (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
        (unsigned int)"IPOIB",
        v6,
        *v7,
        v5);
    }
    return (unsigned int)v5;
  }
  if ( v14 == 1 )
    return 0LL;
  if ( !*((_BYTE *)a1 + 3529) )
  {
    if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
      && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
      && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
    {
      v11 = KeGetCurrentIrql();
      v12 = (_DWORD *)TraceGetDfltActivityPtr();
      WPP_SF_sDD(
        WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
        83,
        (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
        (unsigned int)"IPOIB",
        v11,
        *v12);
    }
    *((_DWORD *)a1 + 1030) = 0;
    StringsList[0] = (char *)a1 + 272;
    NdisWriteEventLogEntry_0(g_pDriverObject, -2147024881, 0, 1u, StringsList, 0, 0LL);
    return 0LL;
  }
  StringsList[0] = (char *)a1 + 272;
  NdisWriteEventLogEntry_0(g_pDriverObject, -1073283068, 0, 1u, StringsList, 0, 0LL);
  if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
    && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
    && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
  {
    v9 = KeGetCurrentIrql();
    v10 = (_DWORD *)TraceGetDfltActivityPtr();
    WPP_SF_sDD(
      WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
      82,
      (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
      (unsigned int)"IPOIB",
      v9,
      *v10);
  }
  return 3221225473LL;
}
