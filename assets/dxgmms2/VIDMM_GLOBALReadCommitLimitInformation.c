void VIDMM_GLOBAL::ReadCommitLimitInformation(void)
{
  int v0; // ecx
  __int64 v1; // r8
  int v2; // edx
  int v3; // eax
  unsigned int v4; // [rsp+30h] [rbp-D0h] BYREF
  int v5; // [rsp+34h] [rbp-CCh] BYREF
  int v6; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v7; // [rsp+3Ch] [rbp-C4h] BYREF
  int v8; // [rsp+40h] [rbp-C0h] BYREF
  int v9; // [rsp+44h] [rbp-BCh] BYREF
  _QWORD v10[42]; // [rsp+50h] [rbp-B0h] BYREF

  v4 = 50;
  v8 = 0;
  qword_1C00762A0 = 0LL;
  v5 = 0;
  v6 = 0;
  v9 = 80;
  v7 = 80;
  memset(v10, 0, sizeof(v10));
  v10[7] = 0LL;
  LODWORD(v10[1]) = 288;
  LODWORD(v10[4]) = 67108868;
  LODWORD(v10[6]) = 4;
  v10[2] = L"PinnedBackingStoreLimit";
  LODWORD(v10[8]) = 288;
  LODWORD(v10[15]) = 288;
  v10[3] = &qword_1C00762A0;
  v10[5] = &v8;
  v10[9] = L"MinimumSystemMemoryCommitLimit";
  v10[10] = &v5;
  v10[16] = L"SmallSystemMemorySize";
  v10[17] = &v6;
  v10[23] = L"SystemPartitionCommitLimitPercentage";
  v10[24] = &v4;
  v10[26] = &v4;
  v10[30] = L"SecondaryPartitionCommitLimitPercentage";
  v10[31] = &v7;
  LODWORD(v10[22]) = 288;
  LODWORD(v10[25]) = 67108868;
  LODWORD(v10[27]) = 4;
  LODWORD(v10[29]) = 288;
  LODWORD(v10[32]) = 67108868;
  LODWORD(v10[34]) = 4;
  v10[33] = &v9;
  LODWORD(v10[11]) = 0x4000000;
  v10[12] = 0LL;
  LODWORD(v10[13]) = 0;
  v10[14] = 0LL;
  LODWORD(v10[18]) = 0x4000000;
  v10[19] = 0LL;
  LODWORD(v10[20]) = 0;
  v10[21] = 0LL;
  v10[28] = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", v10, 0LL, 0LL);
  v0 = v4;
  v1 = qword_1C00762A0 << 20;
  v2 = v5 << 20;
  qword_1C00762A0 <<= 20;
  if ( (unsigned int)(v5 << 20) <= 0x4000000 )
    v2 = 0x4000000;
  if ( v4 >= 0x64 )
  {
    v0 = 100;
  }
  else if ( v4 <= 5 )
  {
    v0 = 5;
  }
  v3 = v7;
  if ( v7 >= 0x64 )
  {
    v3 = 100;
  }
  else if ( v7 <= 5 )
  {
    v3 = 5;
  }
  dword_1C00765B0 = v2;
  dword_1C00765B4 = v6 << 20;
  dword_1C00765B8 = v0;
  dword_1C00765BC = v3;
  if ( !v1 )
    qword_1C00762A0 = (unsigned __int64)qword_1C0076288 >> 3;
}