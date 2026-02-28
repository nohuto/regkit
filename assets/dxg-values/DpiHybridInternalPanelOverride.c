bool DpiHybridInternalPanelOverride()
{
  __int64 v1; // [rsp+30h] [rbp-19h] BYREF
  int v2; // [rsp+38h] [rbp-11h]
  const wchar_t *v3; // [rsp+40h] [rbp-9h]
  int *v4; // [rsp+48h] [rbp-1h]
  int v5; // [rsp+50h] [rbp+7h]
  int *v6; // [rsp+58h] [rbp+Fh]
  int v7; // [rsp+60h] [rbp+17h]
  __int64 v8; // [rsp+68h] [rbp+1Fh]
  int v9; // [rsp+70h] [rbp+27h]
  __int64 v10; // [rsp+78h] [rbp+2Fh]
  __int128 v11; // [rsp+80h] [rbp+37h]
  __int128 v12; // [rsp+90h] [rbp+47h]
  int v13; // [rsp+B0h] [rbp+67h] BYREF

  if ( !g_OSTestSigningEnabled )
    return 0;
  v13 = 0;
  v1 = 0LL;
  v8 = 0LL;
  v9 = 0;
  v10 = 0LL;
  v3 = L"HybridInternalPanelOverrideEnable";
  v4 = &v13;
  v6 = &v13;
  v2 = 288;
  v5 = 67108868;
  v7 = 4;
  v11 = 0LL;
  v12 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v1);
  return v13 != 0;
}
