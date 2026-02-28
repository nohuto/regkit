void VIDMM_GLOBAL::ReadPowerConfiguration(void)
{
  unsigned int v0; // [rsp+30h] [rbp-79h] BYREF
  unsigned int v1; // [rsp+34h] [rbp-75h] BYREF
  int v2; // [rsp+38h] [rbp-71h] BYREF
  int v3; // [rsp+3Ch] [rbp-6Dh] BYREF
  __int64 v4; // [rsp+40h] [rbp-69h] BYREF
  int v5; // [rsp+48h] [rbp-61h]
  const wchar_t *v6; // [rsp+50h] [rbp-59h]
  unsigned int *v7; // [rsp+58h] [rbp-51h]
  int v8; // [rsp+60h] [rbp-49h]
  int *v9; // [rsp+68h] [rbp-41h]
  int v10; // [rsp+70h] [rbp-39h]
  __int64 v11; // [rsp+78h] [rbp-31h]
  int v12; // [rsp+80h] [rbp-29h]
  const wchar_t *v13; // [rsp+88h] [rbp-21h]
  unsigned int *v14; // [rsp+90h] [rbp-19h]
  int v15; // [rsp+98h] [rbp-11h]
  int *v16; // [rsp+A0h] [rbp-9h]
  int v17; // [rsp+A8h] [rbp-1h]
  __int128 v18; // [rsp+B0h] [rbp+7h]
  __int128 v19; // [rsp+C0h] [rbp+17h]
  __int128 v20; // [rsp+D0h] [rbp+27h]
  __int64 v21; // [rsp+E0h] [rbp+37h]

  v2 = 300;
  v0 = 300;
  v3 = 300;
  v1 = 300;
  v5 = 288;
  v8 = 67108868;
  v6 = L"MemoryComponentActiveThreshold";
  v7 = &v0;
  v9 = &v2;
  v13 = L"SelfRefreshMemoryEvictionThreshold";
  v14 = &v1;
  v16 = &v3;
  v10 = 4;
  v12 = 288;
  v15 = 67108868;
  v17 = 4;
  v4 = 0LL;
  v11 = 0LL;
  v18 = 0LL;
  v21 = 0LL;
  v19 = 0LL;
  v20 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v4, 0LL, 0LL);
  qword_1C0076548 = (unsigned __int64)v0 << 20;
  qword_1C0076558 = (unsigned __int64)v1 << 20;
}