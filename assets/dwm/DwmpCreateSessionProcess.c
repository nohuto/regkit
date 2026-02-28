__int64 __fastcall DwmpCreateSessionProcess(PVOID Parameter)
{
  unsigned __int64 v1; // r14
  signed int v2; // ebx
  HANDLE Thread; // rdi
  int v4; // ecx
  int v5; // r8d
  int v6; // r9d
  signed int LastError; // eax
  GUID *v9; // [rsp+50h] [rbp-10h] BYREF
  __int64 v10; // [rsp+58h] [rbp-8h] BYREF
  int pvData; // [rsp+98h] [rbp+38h] BYREF
  DWORD pcbData; // [rsp+A0h] [rbp+40h] BYREF
  GUID *v13; // [rsp+A8h] [rbp+48h] BYREF

  v1 = (unsigned int)Parameter;
  v2 = 0;
  Thread = 0LL;
  if ( GetModuleHandleW(L"wininit.exe")
    && (pvData = 0,
        pcbData = 4,
        RegGetValueW(
          HKEY_LOCAL_MACHINE,
          L"Software\\Microsoft\\Windows\\DWM",
          L"OneCoreNoBootDWM",
          0x20000010u,
          0LL,
          &pvData,
          &pcbData),
        pvData) )
  {
    v2 = 1;
  }
  else if ( gDwmFirstLaunch )
  {
    SetLastError(0);
    Thread = CreateThread(0LL, 0LL, DwmpCreateSessionProcessWorker, (LPVOID)v1, 0, 0LL);
    if ( !Thread )
    {
      LastError = GetLastError();
      v2 = LastError;
      if ( LastError > 0 )
        v2 = (unsigned __int16)LastError | 0x80070000;
      if ( v2 >= 0 )
        v2 = -2003304445;
      DoStackCaptureDirect(v2, 0x5C7u);
    }
  }
  else
  {
    DwmpCreateSessionProcessWorker((PVOID)v1);
  }
  if ( (unsigned int)dword_180016000 > 5
    && (qword_180016010 & 0x400000000000LL) != 0
    && (qword_180016018 & 0x400000000000LL) == qword_180016018 )
  {
    pvData = v1;
    v13 = &gDwmInitTargetAppSessionGuid;
    pcbData = v2;
    v9 = &gDwmInitTelemetryActivityId;
    v10 = 0x1000000LL;
    _tlgWriteTemplate<long (_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),&long _tlgWriteTransfer_EventWriteTransfer(_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),_GUID const *,_GUID const *>::Write<_tlgWrapperByVal<8>,_tlgWrapperByRef<16>,_tlgWrapperByVal<4>,_tlgWrapperByVal<4>,_tlgWrapperByRef<16>>(
      v4,
      (unsigned int)&unk_180012317,
      v5,
      v6,
      (__int64)&v10,
      (__int64)&v9,
      (__int64)&pcbData,
      (__int64)&pvData,
      (__int64)&v13);
  }
  if ( Thread )
    CloseHandle(Thread);
  return (unsigned int)v2;
}
