__int64 __fastcall DpiFdoHandleStartDevice(PDEVICE_OBJECT DeviceObject, PIRP Irp)
{
  int *DeviceExtension; // rdi
  struct _IO_STACK_LOCATION *CurrentStackLocation; // r14
  __int64 v6; // rcx
  int v7; // eax
  int v8; // esi
  ULONG_PTR v9; // r8
  bool v10; // r15
  __int64 Status; // rsi
  PIO_SECURITY_CONTEXT SecurityContext; // rdx
  __int64 v14; // r9
  struct _UNICODE_STRING *v15; // rax
  __int64 v16; // rcx
  int v17; // eax
  PUNICODE_STRING v18; // rcx
  size_t v19; // r12
  void *Pool2; // rax
  int v21; // eax
  bool v22; // zf
  _WORD *StartContext; // r14
  int v24; // eax
  NTSTATUS v25; // eax
  void *v26; // rdx
  PIRP v27; // rax
  void *v28; // rdx
  PIRP v29; // rax
  int v31; // eax
  void *v32; // rcx
  PCLIENT_ID ClientId; // [rsp+20h] [rbp-A9h]
  PCLIENT_ID ClientIda; // [rsp+20h] [rbp-A9h]
  int StartRoutine; // [rsp+28h] [rbp-A1h]
  char v36; // [rsp+40h] [rbp-89h] BYREF
  char v37; // [rsp+41h] [rbp-88h]
  unsigned int v38; // [rsp+44h] [rbp-85h] BYREF
  int v39; // [rsp+48h] [rbp-81h] BYREF
  int v40; // [rsp+4Ch] [rbp-7Dh] BYREF
  size_t Size; // [rsp+50h] [rbp-79h] BYREF
  ULONG_PTR v42; // [rsp+58h] [rbp-71h] BYREF
  struct _UNICODE_STRING *v43; // [rsp+60h] [rbp-69h] BYREF
  ULONG_PTR v44; // [rsp+68h] [rbp-61h]
  struct _UNICODE_STRING *FileName; // [rsp+70h] [rbp-59h]
  void *ThreadHandle; // [rsp+78h] [rbp-51h] BYREF
  __int64 v47; // [rsp+80h] [rbp-49h] BYREF
  int v48; // [rsp+88h] [rbp-41h]
  const wchar_t *v49; // [rsp+90h] [rbp-39h]
  int *v50; // [rsp+98h] [rbp-31h]
  int v51; // [rsp+A0h] [rbp-29h]
  int *v52; // [rsp+A8h] [rbp-21h]
  int v53; // [rsp+B0h] [rbp-19h]
  __int64 v54; // [rsp+B8h] [rbp-11h]
  int v55; // [rsp+C0h] [rbp-9h]
  __int64 v56; // [rsp+C8h] [rbp-1h]
  __int128 v57; // [rsp+D0h] [rbp+7h]
  __int128 v58; // [rsp+E0h] [rbp+17h]
  char v60; // [rsp+140h] [rbp+77h]
  char v61; // [rsp+148h] [rbp+7Fh] BYREF

  DeviceExtension = (int *)DeviceObject->DeviceExtension;
  CurrentStackLocation = Irp->Tail.Overlay.CurrentStackLocation;
  v60 = 0;
  v37 = 0;
  v44 = 0LL;
  FileName = 0LL;
  LODWORD(Size) = 0;
  v38 = 0;
  AcquireMiniportListMutex();
  KeEnterCriticalRegion();
  if ( *((_BYTE *)DeviceExtension + 484) )
    DpiCheckForOutstandingD3Requests(DeviceExtension);
  ExAcquireResourceExclusiveLite(*((PERESOURCE *)DeviceExtension + 21), 1u);
  v7 = WindowsQueryLicenseDWORD(v6, &v38);
  if ( v7 < 0 )
  {
    v8 = 1;
    v38 = 1;
    WdLogSingleEntry1(4LL, v7);
    WdLogGlobalForLineNumber = 7066;
  }
  else
  {
    v8 = v38;
    WdLogSingleEntry1(4LL, v38);
    WdLogGlobalForLineNumber = 7053;
  }
  v39 = 1;
  v47 = 0LL;
  v54 = 0LL;
  v49 = L"MultiMonSupport";
  v55 = 0;
  v56 = 0LL;
  v50 = &v39;
  HIDWORD(ClientId) = 0;
  v52 = &v39;
  v48 = 288;
  v51 = 67108868;
  v53 = 4;
  v57 = 0LL;
  v58 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v47);
  v9 = 0LL;
  if ( !v39 )
    v8 = 0;
  v38 = v8;
  if ( !v8
    && DeviceExtension[4] == 1953656900
    && DeviceExtension[5] == 2
    && !(unsigned __int8)DpiFdoIsPostDevice(DeviceObject)
    && *((_BYTE *)DeviceExtension + 480) == (_BYTE)v9 )
  {
    v10 = 1;
    LODWORD(Status) = -1071774664;
    WdLogSingleEntry1((unsigned int)(v9 + 3), *((_QWORD *)DeviceExtension + 3));
    WdLogGlobalForLineNumber = 7113;
LABEL_96:
    if ( DeviceExtension[59] == 1 )
    {
      v31 = DeviceExtension[60];
      --DeviceExtension[69];
      DeviceExtension[59] = v31;
      DeviceExtension[60] = DeviceExtension[(DeviceExtension[69] & 7) + 61];
    }
    v32 = (void *)*((_QWORD *)DeviceExtension + 164);
    if ( v32 )
    {
      ExFreePoolWithTag(v32, 0);
      *((_QWORD *)DeviceExtension + 164) = 0LL;
    }
    if ( !v10 )
    {
      LOBYTE(StartRoutine) = 0;
      DxgCreateLiveDumpWithWdLogs(403LL, 2050LL, (int)Status, DeviceExtension[59], DeviceExtension[60], StartRoutine);
    }
    goto LABEL_88;
  }
  if ( *((_BYTE *)DeviceExtension + 1158) != (_BYTE)v9 )
  {
    ExEnterCriticalRegionAndAcquireFastMutexUnsafe(&dword_1C01540A0);
    if ( !dword_1C01540D8++ )
      KeClearEvent(&Object);
    ExReleaseFastMutexUnsafeAndLeaveCriticalRegion(&dword_1C01540A0);
    v9 = 0LL;
    v37 = 1;
  }
  DeviceExtension[678] = v8;
  if ( *((_BYTE *)DeviceExtension + 1155) == 1 )
  {
    SecurityContext = CurrentStackLocation->Parameters.Create.SecurityContext;
    v42 = v9;
    v43 = (struct _UNICODE_STRING *)v9;
    if ( SecurityContext )
    {
      FileName = CurrentStackLocation->Parameters.QueryDirectory.FileName;
      v44 = (ULONG_PTR)SecurityContext;
      DpiFilterOutVgaResources(DeviceExtension, SecurityContext, &v42, 0LL, 0LL);
      LOBYTE(v14) = 1;
      DpiFilterOutVgaResources(
        DeviceExtension,
        CurrentStackLocation->Parameters.QueryDirectory.FileName,
        &v43,
        v14,
        ClientIda);
      if ( v42 )
      {
        v15 = v43;
        if ( v43 )
        {
          CurrentStackLocation->Parameters.WMI.ProviderId = v42;
          CurrentStackLocation->Parameters.QueryDirectory.FileName = v15;
          v60 = 1;
        }
      }
    }
    else
    {
      WdLogSingleEntry1(2LL, 0LL);
      WdLogGlobalForLineNumber = 7181;
    }
  }
  if ( !(unsigned __int8)DpiFdoIsPostDevice(DeviceObject) && DeviceExtension[4] == 1953656900 && DeviceExtension[5] == 2 )
  {
    v40 = 0;
    LODWORD(ClientId) = 2;
    v17 = DpiReadPnpRegistryValue(v16, L"DisableNonPOSTDevice", &v40, 4LL, ClientId);
    if ( v17 >= 0 )
    {
      if ( v40 )
      {
        LODWORD(Status) = -1073741823;
        WdLogSingleEntry1(2LL, -1073741823LL);
        WdLogGlobalForLineNumber = 7230;
LABEL_36:
        v10 = 0;
        goto LABEL_96;
      }
    }
    else
    {
      WdLogSingleEntry1(4LL, v17);
      WdLogGlobalForLineNumber = 7217;
    }
  }
  IoForwardIrpSynchronously(*((PDEVICE_OBJECT *)DeviceExtension + 20), Irp);
  Status = Irp->IoStatus.Status;
  if ( (int)Status < 0 )
  {
    WdLogSingleEntry5(
      2LL,
      (unsigned int)DeviceExtension[136],
      Status,
      (unsigned int)DeviceExtension[281],
      (unsigned int)DeviceExtension[282],
      *(_QWORD *)(*((_QWORD *)DeviceExtension + 5) + 152LL));
    WdLogGlobalForLineNumber = 7255;
    v10 = (_DWORD)Status == -1073741810
       && *(_BYTE *)(*((_QWORD *)DeviceExtension + 5) + 134LL)
       && RtlCompareMemory(DeviceExtension + 136, &GUID_BUS_TYPE_USB, 0x10uLL) == 16;
    goto LABEL_96;
  }
  v18 = CurrentStackLocation->Parameters.QueryDirectory.FileName;
  if ( v18 )
  {
    DpiDetermineResourceListSize(v18, &Size);
    v19 = (unsigned int)Size;
    Pool2 = (void *)ExAllocatePool2(256LL, (unsigned int)Size, 1953656900LL);
    *((_QWORD *)DeviceExtension + 164) = Pool2;
    if ( !Pool2 )
    {
      LODWORD(Status) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 7301;
      goto LABEL_36;
    }
    memmove(Pool2, CurrentStackLocation->Parameters.QueryDirectory.FileName, v19);
    if ( v60 == 1 )
    {
      ExFreePoolWithTag(CurrentStackLocation->Parameters.Create.SecurityContext, 0);
      ExFreePoolWithTag(CurrentStackLocation->Parameters.QueryDirectory.FileName, 0);
      CurrentStackLocation->Parameters.WMI.ProviderId = v44;
      CurrentStackLocation->Parameters.QueryDirectory.FileName = FileName;
    }
  }
  DeviceExtension[(DeviceExtension[69] & 7) + 61] = DeviceExtension[60];
  v21 = DeviceExtension[59];
  ++DeviceExtension[69];
  DeviceExtension[60] = v21;
  DeviceExtension[59] = 1;
  if ( DeviceExtension[4] != 1953656900 || DeviceExtension[5] != 2 )
  {
LABEL_47:
    if ( !(_BYTE)word_1C0153AE0 )
      goto LABEL_50;
    goto LABEL_48;
  }
  if ( !*((_BYTE *)DeviceExtension + 2717) )
  {
    HIBYTE(word_1C0153AE0) = 1;
    goto LABEL_47;
  }
  LOBYTE(word_1C0153AE0) = 1;
LABEL_48:
  if ( HIBYTE(word_1C0153AE0) )
    KeSetEvent(&stru_1C0153AE8, 0, 0);
LABEL_50:
  v61 = 0;
  v36 = 0;
  if ( *((_BYTE *)DeviceExtension + 2716) || (int)DpiFdoIsMdmDeviceAndOwnsMux(DeviceObject, &v61, &v36) < 0 || !v61 )
  {
    if ( qword_1C0153DE8 )
      goto LABEL_64;
    if ( DeviceExtension[4] != 1953656900 || DeviceExtension[5] != 2 )
    {
      if ( *(_BYTE *)(*((_QWORD *)DeviceExtension + 21) + 108LL) && *((_QWORD *)DeviceExtension + 354) )
        qword_1C0153DE8 = *((_QWORD *)DeviceExtension + 354);
      goto LABEL_64;
    }
    if ( (unsigned __int8)DpiFdoIsPostDevice(DeviceObject) )
    {
LABEL_60:
      qword_1C0153DE8 = (__int64)DeviceObject;
      goto LABEL_64;
    }
    v22 = *(_BYTE *)(*((_QWORD *)DeviceExtension + 21) + 108LL) == 0;
  }
  else
  {
    v22 = v36 == 0;
  }
  if ( !v22 )
    goto LABEL_60;
LABEL_64:
  if ( !*((_BYTE *)DeviceExtension + 480)
    && *((_BYTE *)DeviceExtension + 1153)
    && !(unsigned __int8)DpiFdoIsMsBddAnchoredDevice(DeviceObject) )
  {
    WdLogSingleEntry1(4LL, DeviceObject);
    WdLogGlobalForLineNumber = 7432;
    v10 = 1;
    LODWORD(Status) = -1071774664;
    goto LABEL_96;
  }
  if ( byte_1C0153AE2 && !*((_BYTE *)DeviceExtension + 1158) )
  {
    ThreadHandle = 0LL;
    StartContext = (_WORD *)ExAllocatePool2(256LL, 1552LL, 1953656900LL);
    if ( !StartContext )
    {
      LODWORD(Status) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 7476;
LABEL_72:
      v10 = 0;
      goto LABEL_96;
    }
    if ( *(_BYTE *)(*((_QWORD *)DeviceExtension + 21) + 108LL)
      || (v22 = (unsigned __int8)DpiFdoIsMsBddAnchoredDevice(DeviceObject) == 0, v24 = 0, !v22) )
    {
      v24 = 2;
    }
    *(_DWORD *)StartContext = v24;
    StartContext[2] = 0;
    *((_DWORD *)StartContext + 131) = 0;
    v25 = PsCreateSystemThread(&ThreadHandle, 0x1FFFFFu, 0LL, 0LL, 0LL, DpiFdoStartAdapterThread, StartContext);
    LODWORD(Status) = v25;
    if ( v25 < 0 )
    {
      WdLogSingleEntry1(2LL, v25);
      WdLogGlobalForLineNumber = 7507;
      ExFreePoolWithTag(StartContext, 0x74727044u);
      goto LABEL_72;
    }
    ZwClose(ThreadHandle);
  }
  DeviceExtension[71] = 1;
  DeviceExtension[70] = 1;
  PoSetPowerState(DeviceObject, DevicePowerState, (POWER_STATE)1);
  v10 = 0;
  if ( (int)Status < 0 )
    goto LABEL_96;
  if ( DeviceExtension[4] == 1953656900 && DeviceExtension[5] == 2 )
  {
    v26 = (void *)*((_QWORD *)DeviceExtension + 684);
    if ( v26 )
    {
      v27 = IoCsqRemoveNextIrp((PIO_CSQ)(DeviceExtension + 1346), v26);
      *((_QWORD *)DeviceExtension + 684) = 0LL;
      if ( v27 )
      {
        *((_BYTE *)DeviceExtension + 5500) = 1;
        v27->IoStatus.Status = 0;
        v27->IoStatus.Information = 0LL;
        IofCompleteRequest(v27, 0);
        IoInvalidateDeviceState(*((PDEVICE_OBJECT *)DeviceExtension + 19));
      }
    }
    v28 = (void *)*((_QWORD *)DeviceExtension + 688);
    if ( v28 )
    {
      v29 = IoCsqRemoveNextIrp((PIO_CSQ)(DeviceExtension + 1346), v28);
      *((_QWORD *)DeviceExtension + 688) = 0LL;
      if ( v29 )
      {
        *((_BYTE *)DeviceExtension + 5532) = 1;
        v29->IoStatus.Status = 0;
        v29->IoStatus.Information = 0LL;
        IofCompleteRequest(v29, 0);
      }
    }
  }
LABEL_88:
  if ( v37 )
  {
    ExEnterCriticalRegionAndAcquireFastMutexUnsafe(&dword_1C01540A0);
    if ( !--dword_1C01540D8 )
      KeSetEvent(&Object, 0, 0);
    ExReleaseFastMutexUnsafeAndLeaveCriticalRegion(&dword_1C01540A0);
  }
  if ( *((_BYTE *)DeviceExtension + 484) )
    DpiEnableD3Requests(*((_QWORD *)DeviceExtension + 3));
  ExReleaseResourceLite(*((PERESOURCE *)DeviceExtension + 21));
  KeLeaveCriticalRegion();
  ReleaseMiniportListMutex();
  Irp->IoStatus.Status = Status;
  IofCompleteRequest(Irp, 1);
  return (unsigned int)Status;
}
