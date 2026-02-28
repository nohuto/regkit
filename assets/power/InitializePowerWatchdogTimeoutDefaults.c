void __fastcall InitializePowerWatchdogTimeoutDefaults(__int64 a1, __int64 a2, __int64 a3, __int64 a4)
{
  __int64 UserSessionState; // rax
  __int64 v5; // r8
  __int64 v6; // rax
  char *v7; // rdx
  _DWORD *v8; // rcx
  _QWORD v10[15]; // [rsp+28h] [rbp-31h] BYREF

  v10[12] = L"PowerWatchdogDrvSetMonitorTimeoutMsec";
  v10[10] = 30000LL;
  v10[11] = 80LL;
  v10[13] = 10000LL;
  UserSessionState = W32GetUserSessionState(
                       a1,
                       a2,
                       a3,
                       a4,
                       16LL,
                       L"PowerWatchdogRequestQueueTimeoutMsec",
                       30000LL,
                       32LL,
                       L"PowerWatchdogPoCalloutTimeoutMsec",
                       10000LL,
                       48LL,
                       L"PowerWatchdogPowerOnGdiTimeoutMsec",
                       30000LL,
                       64LL,
                       L"PowerWatchdogDwmSyncFlushTimeoutMsec");
  v5 = 5LL;
  v6 = UserSessionState + 360;
  v7 = (char *)&v10[-1] - v6;
  v8 = (_DWORD *)(v6 + 16);
  do
  {
    *(v8 - 4) = *(_DWORD *)((char *)v8 + (_QWORD)v7 - 16);
    *((_QWORD *)v8 - 1) = *(_QWORD *)((char *)v8 + (_QWORD)v10 - v6 - 16);
    *v8 = *(_DWORD *)((char *)v8 + (_QWORD)v7);
    v8 += 6;
    --v5;
  }
  while ( v5 );
}
