bool _TdrIsTestMode(void)
{
  int v0; // eax
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
  int v15; // [rsp+B8h] [rbp+6Fh] BYREF

  v15 = 0;
  v14 = 0;
  v2 = 0LL;
  v9 = 0LL;
  v10 = 0;
  v11 = 0LL;
  v4 = L"TdrTestMode";
  v5 = &v14;
  v7 = &v15;
  v3 = 288;
  v6 = 67108868;
  v8 = 4;
  v12 = 0LL;
  v13 = 0LL;
  if ( (int)RtlQueryRegistryValuesEx(
              0LL,
              L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\GraphicsDrivers",
              &v2) >= 0 )
    v0 = v14;
  else
    v0 = 0;
  return v0 != 0;
}
