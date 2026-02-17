void __fastcall HUBREG_QueryUsbflagsAlternateSettingFilter(__int64 a1, __int64 a2)
{
  int v4; // edx
  int v5; // r9d
  __int64 Pool2; // rax
  int v7; // edx
  unsigned int v8; // [rsp+60h] [rbp+18h] BYREF
  int v9; // [rsp+68h] [rbp+20h] BYREF

  v9 = 0;
  v8 = 0;
  if ( (*(unsigned int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, _QWORD, _QWORD, unsigned int *, int *))(WdfFunctions_01015 + 1880))(
         WdfDriverGlobals,
         a2,
         L",.", // AlternateSettingFilter
         0LL,
         0LL,
         &v8,
         &v9) == -2147483643 )
  {
    if ( !v8 || (v8 & 1) != 0 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        return;
      v5 = 20;
      goto LABEL_19;
    }
    if ( v9 != 3 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        return;
      v5 = 21;
LABEL_19:
      LOBYTE(v4) = 2;
      WPP_RECORDER_SF_(
        *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
        v4,
        5,
        v5,
        (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids);
      return;
    }
    Pool2 = ExAllocatePool2(64LL, v8 + 6LL, 1681082453LL);
    *(_QWORD *)(a1 + 2456) = Pool2;
    if ( !Pool2 )
    {
      if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        return;
      v5 = 22;
      goto LABEL_19;
    }
    if ( (*(int (__fastcall **)(PWDF_DRIVER_GLOBALS, __int64, const wchar_t *, _QWORD, __int64, _QWORD, _QWORD))(WdfFunctions_01015 + 1880))(
           WdfDriverGlobals,
           a2,
           L",.", // AlternateSettingFilter
           v8,
           Pool2 + 4,
           0LL,
           0LL) >= 0 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v7) = 4;
        WPP_RECORDER_SF_(
          *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
          v7,
          5,
          24,
          (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids);
      }
      **(_DWORD **)(a1 + 2456) = v8 >> 1;
    }
    else
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v7) = 2;
        WPP_RECORDER_SF_(
          *(_QWORD *)(*(_QWORD *)(a1 + 8) + 1432LL),
          v7,
          5,
          23,
          (__int64)&WPP_6348287eaa4439ce1c5af6747761b290_Traceguids);
      }
      ExFreePoolWithTag(*(PVOID *)(a1 + 2456), 0x64334855u);
      *(_QWORD *)(a1 + 2456) = 0LL;
    }
  }
}