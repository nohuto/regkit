__int64 __fastcall VidSchiReadDeviceConfiguration(__int64 a1)
{
  int v1; // ebx
  __int64 result; // rax
  _QWORD v4[16]; // [rsp+30h] [rbp-29h] BYREF
  int v5; // [rsp+C8h] [rbp+6Fh] BYREF
  int v6; // [rsp+D0h] [rbp+77h] BYREF

  v1 = 0;
  v6 = 0;
  v5 = 0;
  memset(v4, 0, 0x70uLL);
  LODWORD(v4[1]) = 288;
  v4[2] = L"FlipOverrideMode";
  LODWORD(v4[4]) = 67108868;
  v4[3] = &v5;
  LODWORD(v4[6]) = 4;
  v4[5] = &v6;
  result = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\Scheduler", v4, 0LL, 0LL);
  if ( v5 == 1 )
  {
    *(_DWORD *)(a1 + 960) = 1;
  }
  else
  {
    if ( v5 == 2 )
      v1 = 2;
    *(_DWORD *)(a1 + 960) = v1;
  }
  return result;
}