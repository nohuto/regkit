NTSTATUS __fastcall HidpFdoConfigureIdleSettings(char *Context, __int16 a2)
{
  unsigned int v2; // esi
  unsigned int v4; // eax
  __int64 v5; // rax
  NTSTATUS result; // eax
  int v7; // ebx
  int v8; // edx
  NTSTATUS v9; // eax
  int v10; // edx
  int v11; // r8d
  unsigned int v12; // eax
  int v13; // edx
  int v14; // eax
  int v15; // eax
  NTSTATUS v16; // eax
  int v17; // edx
  int v18; // r8d
  NTSTATUS v19; // eax
  int v20; // edx
  int v21; // r8d
  NTSTATUS v22; // eax
  int v23; // edx
  int v24; // r8d
  int v25; // edx
  char v26; // bl
  int v27; // eax
  int v28; // edx
  unsigned int v29; // r9d
  char v30; // r8
  __int64 v31; // r10
  __int64 v32; // rcx
  __int16 v33; // dx
  __int64 v34; // rax
  __int16 v35; // cx
  int v36; // eax
  int v37; // ecx
  NTSTATUS v38; // eax
  int v39; // r9d
  GUID *v40; // rdx
  int v41; // eax
  NTSTATUS v42; // eax
  int Length; // [rsp+20h] [rbp-39h]
  int Lengtha; // [rsp+20h] [rbp-39h]
  int Lengthb; // [rsp+20h] [rbp-39h]
  int Lengthc; // [rsp+20h] [rbp-39h]
  ULONG ResultLength; // [rsp+40h] [rbp-19h] BYREF
  void *DeviceRegKey; // [rsp+48h] [rbp-11h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+50h] [rbp-9h] BYREF
  int v50; // [rsp+60h] [rbp+7h] BYREF
  __int64 v51; // [rsp+68h] [rbp+Fh] BYREF
  __int128 KeyValueInformation; // [rsp+70h] [rbp+17h] BYREF
  __int64 v53; // [rsp+80h] [rbp+27h] BYREF

  v2 = 0;
  DeviceRegKey = 0LL;
  ResultLength = 0;
  DestinationString = 0LL;
  v51 = 0LL;
  KeyValueInformation = 0LL;
  v53 = WNF_PO_INPUT_SUPPRESS_NOTIFICATION_EX;
  v50 = 0;
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED && LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(a2) = 5;
    WPP_RECORDER_SF_(
      WPP_GLOBAL_Control->DeviceExtension,
      a2,
      9,
      15,
      (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids);
  }
  v4 = *((_DWORD *)Context + 439) & 0xFFFFFEB8;
  *(_QWORD *)(Context + 1772) = 0LL;
  *((_QWORD *)Context + 220) = 0LL;
  *((_DWORD *)Context + 439) = v4 | 0x38;
  v5 = *(_QWORD *)Context;
  *((_DWORD *)Context + 433) = 1000;
  *((_DWORD *)Context + 432) = 5000;
  *((_DWORD *)Context + 435) = 5000;
  *((_DWORD *)Context + 434) = 1000;
  *((_DWORD *)Context + 436) = 1000;
  *((_DWORD *)Context + 437) = 1000;
  result = IoOpenDeviceRegistryKey(**(PDEVICE_OBJECT **)(v5 + 64), 1u, 0x20000u, &DeviceRegKey);
  v7 = result;
  if ( result >= 0 )
  {
    RtlInitUnicodeString(&DestinationString, L"EnhancedPowerManagementEnabled");
    v9 = ZwQueryValueKey(
           DeviceRegKey,
           &DestinationString,
           KeyValuePartialInformation,
           &KeyValueInformation,
           0x10u,
           &ResultLength);
    if ( v9 < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v10) = 3;
        WPP_RECORDER_SF_qSd(
          *((_QWORD *)Context + 84),
          v10,
          v11,
          17,
          Length,
          *(_QWORD *)Context,
          (__int64)L"EnhancedPowerManagementEnabled",
          v9);
      }
    }
    else
    {
      v12 = (BYTE12(KeyValueInformation) != 0 ? 0x10 : 0) | *((_DWORD *)Context + 439) & 0xFFFFFFCF;
      *((_DWORD *)Context + 439) = v12;
      if ( (v12 & 0x30) == 0x10 )
      {
        RtlInitUnicodeString(&DestinationString, L"EnhancedPowerManagementUseMonitor");
        if ( ZwQueryValueKey(
               DeviceRegKey,
               &DestinationString,
               KeyValuePartialInformation,
               &KeyValueInformation,
               0x10u,
               &ResultLength) >= 0 )
        {
          if ( BYTE12(KeyValueInformation) )
            *((_DWORD *)Context + 439) |= 0x40u;
        }
      }
    }
    RtlInitUnicodeString(&DestinationString, L"WakeScreenOnInputSupport");
    if ( (int)HidpFdoRegistryQueryULong(Context, DeviceRegKey, &DestinationString, &v50) >= 0 )
    {
      v14 = v50;
      *((_DWORD *)Context + 441) = v50;
      if ( v14 )
      {
        v15 = *((_DWORD *)Context + 86);
        if ( v15 > 1 )
        {
          RtlInitUnicodeString(&DestinationString, L"WakeScreenOnInputTimeout");
          HidpFdoRegistryQueryULong(Context, DeviceRegKey, &DestinationString, Context + 1748);
        }
        else
        {
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
            WPP_RECORDER_SF_qD(
              *((_QWORD *)Context + 85),
              v13,
              9,
              18,
              (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
              *(_QWORD *)Context,
              v15);
          *((_DWORD *)Context + 441) = 0;
        }
      }
    }
    RtlInitUnicodeString(&DestinationString, L"SelectiveSuspendOn");
    v16 = ZwQueryValueKey(
            DeviceRegKey,
            &DestinationString,
            KeyValuePartialInformation,
            &KeyValueInformation,
            0x10u,
            &ResultLength);
    if ( v16 < 0 )
    {
      if ( v16 == -1073741772 )
      {
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          LOBYTE(v17) = 4;
          WPP_RECORDER_SF_qSd(
            *((_QWORD *)Context + 84),
            v17,
            v18,
            19,
            Lengtha,
            *(_QWORD *)Context,
            (__int64)L"SelectiveSuspendOn",
            52);
        }
      }
      else if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v17) = 3;
        WPP_RECORDER_SF_qSd(
          *((_QWORD *)Context + 84),
          v17,
          v18,
          20,
          Lengtha,
          *(_QWORD *)Context,
          (__int64)L"SelectiveSuspendOn",
          v16);
      }
    }
    else if ( !BYTE12(KeyValueInformation) )
    {
      *((_DWORD *)Context + 439) &= ~8u;
    }
    RtlInitUnicodeString(&DestinationString, L"SelectiveSuspendEnabled");
    v19 = ZwQueryValueKey(
            DeviceRegKey,
            &DestinationString,
            KeyValuePartialInformation,
            &KeyValueInformation,
            0x10u,
            &ResultLength);
    if ( v19 < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v20) = 3;
        WPP_RECORDER_SF_qSd(
          *((_QWORD *)Context + 84),
          v20,
          v21,
          21,
          Lengthb,
          *(_QWORD *)Context,
          (__int64)L"SelectiveSuspendEnabled",
          v19);
      }
    }
    else if ( BYTE12(KeyValueInformation) )
    {
      *((_DWORD *)Context + 439) |= 4u;
    }
    RtlInitUnicodeString(&DestinationString, L"SelectiveSuspendTimeout");
    if ( ZwQueryValueKey(
           DeviceRegKey,
           &DestinationString,
           KeyValuePartialInformation,
           &KeyValueInformation,
           0x10u,
           &ResultLength) >= 0
      && DWORD2(KeyValueInformation) == 4 )
    {
      *((_DWORD *)Context + 432) = HIDWORD(KeyValueInformation);
    }
    RtlInitUnicodeString(&DestinationString, L"SuppressInputInCS");
    v22 = ZwQueryValueKey(
            DeviceRegKey,
            &DestinationString,
            KeyValuePartialInformation,
            &KeyValueInformation,
            0x10u,
            &ResultLength);
    if ( v22 < 0 )
    {
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v23) = 3;
        WPP_RECORDER_SF_qSd(
          *((_QWORD *)Context + 84),
          v23,
          v24,
          23,
          Lengthc,
          *(_QWORD *)Context,
          (__int64)L"SuppressInputInCS",
          v22);
      }
    }
    else if ( BYTE12(KeyValueInformation) )
    {
      *((_DWORD *)Context + 439) |= 0x80u;
      if ( *((_DWORD *)Context + 441) )
      {
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          LOBYTE(v23) = 2;
          WPP_RECORDER_SF_q(
            *((_QWORD *)Context + 85),
            v23,
            9,
            22,
            (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
            *(_QWORD *)Context);
        }
        *((_DWORD *)Context + 441) = 0;
      }
    }
    RtlInitUnicodeString(&DestinationString, L"SystemInputSuppressionEnabled");
    if ( (int)HidpFdoRegistryQueryULong(Context, DeviceRegKey, &DestinationString, &v50) >= 0 && v50 )
      *((_DWORD *)Context + 440) = 1;
    RtlInitUnicodeString(&DestinationString, L"TestIdleTimeoutNoHandlesInitial");
    if ( ZwQueryValueKey(
           DeviceRegKey,
           &DestinationString,
           KeyValuePartialInformation,
           &KeyValueInformation,
           0x10u,
           &ResultLength) >= 0
      && DWORD2(KeyValueInformation) == 4 )
    {
      *((_DWORD *)Context + 435) = HIDWORD(KeyValueInformation);
    }
    RtlInitUnicodeString(&DestinationString, L"TestIdleTimeoutNoHandles");
    if ( ZwQueryValueKey(
           DeviceRegKey,
           &DestinationString,
           KeyValuePartialInformation,
           &KeyValueInformation,
           0x10u,
           &ResultLength) >= 0
      && DWORD2(KeyValueInformation) == 4 )
    {
      *((_DWORD *)Context + 434) = 1000 * HIDWORD(KeyValueInformation);
    }
    RtlInitUnicodeString(&DestinationString, L"TestIdleMonitorDim");
    if ( ZwQueryValueKey(
           DeviceRegKey,
           &DestinationString,
           KeyValuePartialInformation,
           &KeyValueInformation,
           0x10u,
           &ResultLength) >= 0
      && DWORD2(KeyValueInformation) == 4 )
    {
      *((_DWORD *)Context + 439) |= 1u;
      *((_DWORD *)Context + 436) = 1000 * HIDWORD(KeyValueInformation);
    }
    ZwClose(DeviceRegKey);
    HidpGetDeviceFlags(Context, &v51);
    v26 = v51;
    if ( (v51 & 1) != 0 )
    {
      *((_DWORD *)Context + 439) |= 1u;
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v25) = 4;
        WPP_RECORDER_SF_q(
          *((_QWORD *)Context + 84),
          v25,
          9,
          24,
          (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
          *(_QWORD *)Context);
      }
    }
    if ( (v26 & 2) != 0 )
    {
      *((_DWORD *)Context + 439) |= 2u;
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v25) = 4;
        WPP_RECORDER_SF_q(
          *((_QWORD *)Context + 84),
          v25,
          9,
          25,
          (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
          *(_QWORD *)Context);
      }
    }
    if ( (v26 & 4) != 0 )
    {
      *((_DWORD *)Context + 439) &= ~4u;
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v25) = 4;
        WPP_RECORDER_SF_q(
          *((_QWORD *)Context + 84),
          v25,
          9,
          26,
          (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
          *(_QWORD *)Context);
      }
    }
    if ( (v26 & 0x10) != 0 )
    {
      *((_DWORD *)Context + 439) |= 0x100u;
      if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v25) = 4;
        WPP_RECORDER_SF_q(
          *((_QWORD *)Context + 84),
          v25,
          9,
          27,
          (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
          *(_QWORD *)Context);
      }
    }
    if ( (Context[1756] & 0x30) != 0x30 || Context[2048] )
    {
      v27 = HidpRegisterSleepstudyBlockerReasons(**(_QWORD **)(*(_QWORD *)Context + 64LL), Context);
      if ( v27 < 0 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v28) = 2;
        WPP_RECORDER_SF_qd(
          *((_QWORD *)Context + 84),
          v28,
          11,
          28,
          (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
          *(_QWORD *)Context,
          v27);
      }
    }
    if ( *((_DWORD *)Context + 440) != 1 && Context[2048] && (*((_DWORD *)Context + 439) & 0x30) != 0 )
    {
      v29 = *((_DWORD *)Context + 42);
      v30 = 0;
      if ( v29 )
      {
        v31 = *((_QWORD *)Context + 19);
        do
        {
          v32 = 424LL * v2;
          v33 = *(_WORD *)(v32 + v31 + 8);
          v34 = v32 + v31;
          if ( v33 == 1 && ((v35 = *(_WORD *)(v34 + 10), v35 == 2) || v35 == 6) )
          {
            v30 = 1;
          }
          else if ( v33 == 13 && *(_WORD *)(v34 + 10) == 5 )
          {
            *((_DWORD *)Context + 440) = 3;
            goto LABEL_91;
          }
          ++v2;
        }
        while ( v2 < v29 );
        if ( !v30 )
          goto LABEL_90;
        *((_DWORD *)Context + 440) = 1;
      }
      else
      {
LABEL_90:
        *((_DWORD *)Context + 440) = 2;
      }
    }
LABEL_91:
    TraceLoggingIdleConfiguration(Context);
    v36 = HidpFdoRegisterWithPoFx(Context);
    *((_DWORD *)Context + 445) |= 1u;
    v7 = v36;
    if ( v36 < 0 )
      goto LABEL_118;
    HidFdoStartRunTimePolicyEngine(Context);
    v37 = *((_DWORD *)Context + 439);
    if ( (v37 & 1) != 0 )
    {
      v38 = PoRegisterPowerSettingCallback(
              0LL,
              &GUID_CONSOLE_DISPLAY_STATE,
              HidpFdoPowerSettingCallback,
              Context,
              (PVOID *)Context + 77);
      v7 = v38;
      if ( v38 >= 0 || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_107;
      v39 = 29;
    }
    else
    {
      if ( (*((_DWORD *)Context + 439) & 0x30) != 0x10
        && ((v37 & 0x30) == 0 || !Context[2048])
        && !*((_DWORD *)Context + 441) )
      {
        goto LABEL_107;
      }
      if ( (v37 & 0x40) != 0 || (v40 = &GUID_LOW_POWER_EPOCH, *((_DWORD *)Context + 441)) )
        v40 = &GUID_MONITOR_POWER_ON;
      v38 = PoRegisterPowerSettingCallback(0LL, v40, HidpFdoPowerSettingCallback, Context, (PVOID *)Context + 77);
      if ( v38 >= 0 || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        goto LABEL_107;
      v39 = 30;
    }
    LOBYTE(v8) = 2;
    WPP_RECORDER_SF_qd(
      *((_QWORD *)Context + 84),
      v8,
      9,
      v39,
      (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
      *(_QWORD *)Context,
      v38);
LABEL_107:
    if ( *((_DWORD *)Context + 440) == 1 )
    {
      v41 = ExSubscribeWnfStateChange(Context + 632, &v53, 1LL);
      if ( v41 < 0 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v8) = 2;
        WPP_RECORDER_SF_qd(
          *((_QWORD *)Context + 84),
          v8,
          9,
          31,
          (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
          *(_QWORD *)Context,
          v41);
      }
    }
    if ( v7 >= 0 )
    {
      v42 = IoWMIRegistrationControl(*(PDEVICE_OBJECT *)Context, 1u);
      v7 = v42;
      if ( v42 < 0 && WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
      {
        LOBYTE(v8) = 2;
        WPP_RECORDER_SF_qd(
          *((_QWORD *)Context + 84),
          v8,
          9,
          32,
          (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
          *(_QWORD *)Context,
          v42);
      }
    }
    if ( *((_QWORD *)Context + 270) && (*((_DWORD *)Context + 439) & 0x30) == 0 )
      SleepstudyHelper_ComponentActive();
    goto LABEL_118;
  }
  if ( WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
    return result;
  v8 = (int)WPP_GLOBAL_Control;
  if ( LOWORD(WPP_GLOBAL_Control->DeviceType) )
  {
    LOBYTE(v8) = 5;
    WPP_RECORDER_SF_qd(
      *((_QWORD *)Context + 84),
      v8,
      9,
      16,
      (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids,
      *(_QWORD *)Context,
      result);
  }
LABEL_118:
  if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
  {
    if ( LOWORD(WPP_GLOBAL_Control->DeviceType) )
    {
      LOBYTE(v8) = 5;
      WPP_RECORDER_SF_(
        WPP_GLOBAL_Control->DeviceExtension,
        (_WORD)v8,
        9,
        33,
        (__int64)&WPP_bf5a577156b7397c7ad6ea14446ed258_Traceguids);
    }
  }
  return v7;
}