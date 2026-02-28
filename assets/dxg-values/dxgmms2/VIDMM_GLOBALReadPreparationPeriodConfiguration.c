void VIDMM_GLOBAL::ReadPreparationPeriodConfiguration(void)
{
  int v0; // [rsp+30h] [rbp-D0h] BYREF
  int v1; // [rsp+34h] [rbp-CCh] BYREF
  int v2; // [rsp+38h] [rbp-C8h] BYREF
  int v3; // [rsp+3Ch] [rbp-C4h] BYREF
  int v4; // [rsp+40h] [rbp-C0h] BYREF
  int v5; // [rsp+44h] [rbp-BCh] BYREF
  _QWORD v6[56]; // [rsp+50h] [rbp-B0h] BYREF

  v0 = 1;
  v3 = 1;
  qword_1C0076338 = 1LL;
  dword_1C0076348 = 1;
  v1 = 4;
  v2 = 64;
  v4 = 32;
  dword_1C0076344 = 64;
  dword_1C007634C = 32;
  v5 = 0x7FFFFFFF;
  dword_1C0076340 = 4;
  dword_1C0076350 = 0x7FFFFFFF;
  memset(v6, 0, sizeof(v6));
  LODWORD(v6[6]) = 4;
  LODWORD(v6[1]) = 288;
  LODWORD(v6[4]) = 67108868;
  v6[7] = 0LL;
  v6[2] = L"PreparationPeriod";
  v6[3] = &qword_1C0076338;
  v6[5] = &v0;
  v6[9] = L"Period.MinimumPolicyHeldPeriod";
  v6[10] = &dword_1C0076340;
  v6[12] = &v1;
  v6[16] = L"Period.MaximumPolicyHeldPeriod";
  v6[17] = &dword_1C0076344;
  v6[19] = &v2;
  v6[23] = L"Period.AlwaysForceMemReset";
  v6[24] = &dword_1C0076348;
  v6[26] = &v3;
  v6[30] = L"Period.EvictionThresholdForMemReset";
  v6[31] = &dword_1C007634C;
  v6[33] = &v4;
  v6[37] = L"Period.NbOfAllocationsThresholdToMRU";
  v6[38] = &dword_1C0076350;
  LODWORD(v6[8]) = 288;
  LODWORD(v6[11]) = 67108868;
  LODWORD(v6[13]) = 4;
  v6[14] = 0LL;
  LODWORD(v6[15]) = 288;
  LODWORD(v6[18]) = 67108868;
  LODWORD(v6[20]) = 4;
  v6[21] = 0LL;
  LODWORD(v6[22]) = 288;
  LODWORD(v6[25]) = 67108868;
  LODWORD(v6[27]) = 4;
  v6[28] = 0LL;
  LODWORD(v6[29]) = 288;
  LODWORD(v6[32]) = 67108868;
  LODWORD(v6[34]) = 4;
  v6[35] = 0LL;
  LODWORD(v6[36]) = 288;
  LODWORD(v6[39]) = 67108868;
  v6[40] = &v5;
  LODWORD(v6[41]) = 4;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", v6, 0LL, 0LL);
  qword_1C0076338 *= 10000000LL;
  dword_1C007634C <<= 20;
}