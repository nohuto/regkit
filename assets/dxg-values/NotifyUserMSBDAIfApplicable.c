NTSTATUS NotifyUserMSBDAIfApplicable()
{
  ULONGLONG v0; // rax
  NTSTATUS result; // eax
  _DWORD v2[4]; // [rsp+30h] [rbp-D0h] BYREF
  __int64 v3; // [rsp+40h] [rbp-C0h] BYREF
  int v4; // [rsp+48h] [rbp-B8h]
  const wchar_t *v5; // [rsp+50h] [rbp-B0h]
  _DWORD *v6; // [rsp+58h] [rbp-A8h]
  int v7; // [rsp+60h] [rbp-A0h]
  _DWORD *v8; // [rsp+68h] [rbp-98h]
  int v9; // [rsp+70h] [rbp-90h]
  __int64 v10; // [rsp+78h] [rbp-88h]
  int v11; // [rsp+80h] [rbp-80h]
  __int64 v12; // [rsp+88h] [rbp-78h]
  __int128 v13; // [rsp+90h] [rbp-70h]
  __int128 v14; // [rsp+A0h] [rbp-60h]
  _OSVERSIONINFOEXW VersionInfo; // [rsp+B0h] [rbp-50h] BYREF

  memset(&VersionInfo, 0, sizeof(VersionInfo));
  VersionInfo.wProductType = 1;
  v0 = VerSetConditionMask(0LL, 0x80u, 1u);
  result = RtlVerifyVersionInfo(&VersionInfo, 0x80u, v0);
  if ( result >= 0 )
  {
    v2[0] = 0;
    v3 = 0LL;
    v10 = 0LL;
    v11 = 0;
    v12 = 0LL;
    v5 = L"BasicDisplayUserNotified";
    v6 = v2;
    v8 = v2;
    v4 = 288;
    v7 = 67108868;
    v9 = 4;
    v13 = 0LL;
    v14 = 0LL;
    result = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\BasicDisplay", &v3);
    if ( !v2[0] )
      return WdDiagNotifyUser(0LL, 8LL, 0LL, 0LL, 0LL);
  }
  return result;
}
