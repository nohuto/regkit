__int64 __fastcall PciGetDeviceD0DelayTime(__int64 a1, __int64 *a2)
{
  unsigned int v2; // ebx
  char v4; // si
  __int64 *v5; // r14
  __int64 v6; // rcx
  __int64 result; // rax
  void (__fastcall *v8)(_QWORD, _BYTE **); // rax
  int v9; // edx
  unsigned int v10; // eax
  _BYTE *v11; // [rsp+70h] [rbp+8h] BYREF

  v11 = 0LL;
  v2 = *(unsigned __int16 *)(a1 + 1280);
  v4 = 0;
  v5 = a2;
  v6 = *(_QWORD *)(*(_QWORD *)(a1 + 144) + 144LL);
  if ( v6 )
  {
    v8 = *(void (__fastcall **)(_QWORD, _BYTE **))(v6 + 1064);
    if ( v8 )
    {
      v8(*(_QWORD *)(v6 + 1024), &v11);
      if ( (*v11 & 2) != 0 )
      {
        v2 = 0;
        v4 = 1;
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED
          && !*((_BYTE *)&WPP_MAIN_CB.ActiveThreadCount + 4) )
        {
          v9 = *(_DWORD *)(a1 + 32) & 0x1F;
          LOBYTE(v9) = 4;
          WPP_RECORDER_SF_DDD(
            *(_QWORD *)(*(_QWORD *)(a1 + 144) + 864LL),
            v9,
            6,
            54,
            (__int64)&WPP_759aeb2f4d9d3ab90e9f5c845d120a43_Traceguids,
            *(_DWORD *)(*(_QWORD *)(a1 + 144) + 292LL),
            *(_DWORD *)(a1 + 32) & 0x1F,
            (unsigned __int8)*(_DWORD *)(a1 + 32) >> 5);
        }
      }
    }
  }
  if ( (*(_BYTE *)(a1 + 1968) & 1) != 0 )
  {
    v10 = *(_DWORD *)(a1 + 1976);
    if ( v10 <= 0x64 )
    {
      if ( v4 )
      {
        v2 = *(_DWORD *)(a1 + 1976);
      }
      else if ( v10 < v2 )
      {
        v2 = *(_DWORD *)(a1 + 1976);
      }
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED
        && !*((_BYTE *)&WPP_MAIN_CB.ActiveThreadCount + 4) )
      {
        LOBYTE(a2) = 4;
        WPP_RECORDER_SF_DDDd(
          *(_QWORD *)(*(_QWORD *)(a1 + 144) + 864LL),
          (_DWORD)a2,
          6,
          55,
          (__int64)&WPP_759aeb2f4d9d3ab90e9f5c845d120a43_Traceguids,
          *(_DWORD *)(*(_QWORD *)(a1 + 144) + 292LL),
          *(_DWORD *)(a1 + 32) & 0x1F,
          (unsigned __int8)*(_DWORD *)(a1 + 32) >> 5,
          v2);
      }
    }
  }
  result = PciBus_WaitTimeForBusSettle(*(_QWORD *)(a1 + 144), v2);
  *v5 = result;
  return result;
}