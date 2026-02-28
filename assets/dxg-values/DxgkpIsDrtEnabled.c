char DxgkpIsDrtEnabled()
{
  struct DXGPROCESS *Current; // rax
  char result; // al
  __int64 v2; // [rsp+30h] [rbp-19h] BYREF
  int v3; // [rsp+38h] [rbp-11h]
  const wchar_t *v4; // [rsp+40h] [rbp-9h]
  int *v5; // [rsp+48h] [rbp-1h]
  int v6; // [rsp+50h] [rbp+7h]
  int *v7; // [rsp+58h] [rbp+Fh]
  int v8; // [rsp+60h] [rbp+17h]
  __int64 v9; // [rsp+68h] [rbp+1Fh]
  int v10; // [rsp+70h] [rbp+27h]
  __int64 v11; // [rsp+78h] [rbp+2Fh]
  __int128 v12; // [rsp+80h] [rbp+37h]
  __int128 v13; // [rsp+90h] [rbp+47h]
  int v14; // [rsp+B0h] [rbp+67h] BYREF

  Current = DXGPROCESS::GetCurrent();
  if ( Current && (*((_DWORD *)Current + 102) & 0x1000) != 0 )
    return 1;
  v14 = 0;
  v2 = 0LL;
  v4 = L"DRTTestEnable";
  v9 = 0LL;
  v10 = 0;
  v5 = &v14;
  v11 = 0LL;
  v7 = &v14;
  v3 = 288;
  v6 = 67108868;
  v8 = 4;
  v12 = 0LL;
  v13 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v2);
  if ( v14 == 1484026436 )
    return 1;
  WdLogSingleEntry0(4LL);
  result = 0;
  WdLogGlobalForLineNumber = 50;
  return result;
}
