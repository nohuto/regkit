void VIDMM_GLOBAL::ReadGpuVaPagingHistoryConfiguration(void)
{
  int v0; // eax
  bool v1; // bl
  int v2; // eax
  int v3; // edx
  unsigned int v4; // ecx
  unsigned int v5; // eax
  __int64 v6; // [rsp+30h] [rbp-29h] BYREF
  int v7; // [rsp+38h] [rbp-21h]
  const wchar_t *v8; // [rsp+40h] [rbp-19h]
  unsigned int *v9; // [rsp+48h] [rbp-11h]
  int v10; // [rsp+50h] [rbp-9h]
  __int64 v11; // [rsp+58h] [rbp-1h]
  int v12; // [rsp+60h] [rbp+7h]
  __int128 v13; // [rsp+68h] [rbp+Fh]
  __int128 v14; // [rsp+78h] [rbp+1Fh]
  __int128 v15; // [rsp+88h] [rbp+2Fh]
  __int64 v16; // [rsp+98h] [rbp+3Fh]
  int v17; // [rsp+C0h] [rbp+67h] BYREF
  unsigned int v18; // [rsp+C8h] [rbp+6Fh] BYREF

  Feature_GpuVaPagingHistoryFre__private_ReportDeviceUsage();
  v17 = 391174;
  v7 = 292;
  v10 = 0x4000000;
  VIDMM_GLOBAL::GpuVaPagingHistoryFreEnabled = 1;
  v18 = (unsigned __int64)qword_1C0076288 > 0x53333333 ? 0x40 : 0;
  v6 = 0LL;
  v8 = L"GpuVaPagingHistorySize";
  v11 = 0LL;
  v9 = &v18;
  v12 = 0;
  v13 = 0LL;
  v16 = 0LL;
  v14 = 0LL;
  v15 = 0LL;
  v0 = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v6, 0LL, 0LL);
  v6 = 0LL;
  v7 = 292;
  v8 = L"GpuVaPagingHistoryMask";
  v1 = v0 >= 0;
  v10 = 0x4000000;
  v11 = 0LL;
  v9 = (unsigned int *)&v17;
  v12 = 0;
  v16 = 0LL;
  v13 = 0LL;
  v14 = 0LL;
  v15 = 0LL;
  v2 = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v6, 0LL, 0LL);
  v3 = v17;
  if ( v2 < 0 && v1 )
    v3 = -1;
  v4 = v18;
  if ( VIDMM_GLOBAL::GpuVaPagingHistoryFreEnabled && v18 )
  {
    v5 = v18 >= 0x1FFFFFF ? 0x7FFFFFFF : v18 << 6;
    v4 = (((((((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 8) | ((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 16) | ((((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 8) | ((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1))
       - ((((((((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 8) | ((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 16) | ((((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 8) | ((((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 4) | ((v5 | (v5 >> 1)) >> 2) | v5 | (v5 >> 1)) >> 1);
    if ( v4 <= 0x1000 )
      v4 = 4096;
  }
  dword_1C007645C = v4;
  dword_1C0076460 = v3;
  if ( v4 >= 0x3FFFFFF )
    dword_1C0076464 = 0x7FFFFFFF;
  else
    dword_1C0076464 = 32 * v4;
}