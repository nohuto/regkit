void VIDMM_GLOBAL::ReadUnusedAllocationConfiguration(void)
{
  __int128 *v0; // rbx
  int v1; // [rsp+30h] [rbp-D0h] BYREF
  int v2; // [rsp+34h] [rbp-CCh] BYREF
  int v3; // [rsp+38h] [rbp-C8h] BYREF
  int v4; // [rsp+3Ch] [rbp-C4h] BYREF
  int v5; // [rsp+40h] [rbp-C0h] BYREF
  int v6; // [rsp+44h] [rbp-BCh] BYREF
  int v7; // [rsp+48h] [rbp-B8h] BYREF
  int v8; // [rsp+4Ch] [rbp-B4h] BYREF
  int v9; // [rsp+50h] [rbp-B0h] BYREF
  int v10; // [rsp+54h] [rbp-ACh] BYREF
  int v11; // [rsp+58h] [rbp-A8h] BYREF
  int v12; // [rsp+5Ch] [rbp-A4h] BYREF
  int v13; // [rsp+60h] [rbp-A0h] BYREF
  int v14; // [rsp+64h] [rbp-9Ch] BYREF
  int v15; // [rsp+68h] [rbp-98h] BYREF
  _QWORD v16[112]; // [rsp+70h] [rbp-90h] BYREF

  xmmword_1C0076360 = (__int128)_mm_load_si128((const __m128i *)&_xmm);
  xmmword_1C0076380 = (__int128)_mm_load_si128((const __m128i *)&_xmm);
  xmmword_1C0076370 = (__int128)_mm_load_si128((const __m128i *)&_xmm);
  xmmword_1C00763A0 = (__int128)_mm_load_si128((const __m128i *)&_xmm);
  v6 = 1000000;
  v11 = 1000000;
  xmmword_1C0076390 = (__int128)_mm_load_si128((const __m128i *)&_xmm);
  v1 = 1;
  v3 = 15;
  v8 = 1;
  v12 = 15;
  v13 = 15;
  qword_1C0076358 = 1LL;
  xmmword_1C0076430 = (__int128)_mm_load_si128((const __m128i *)&_xmm);
  v2 = 0;
  v4 = 45;
  v5 = 120;
  v7 = 0;
  v9 = 2;
  v10 = 5;
  v14 = 30;
  v15 = 30;
  xmmword_1C0076420 = (__int128)_mm_load_si128((const __m128i *)&_xmm);
  memset(v16, 0, sizeof(v16));
  v16[7] = 0LL;
  LODWORD(v16[1]) = 288;
  LODWORD(v16[4]) = 67108868;
  LODWORD(v16[6]) = 4;
  v16[2] = L"UnusedTrimmingPeriod";
  v16[3] = &qword_1C0076358;
  v16[5] = &v1;
  v16[9] = L"Unused.MinimumThreshold";
  v16[10] = &xmmword_1C0076360;
  v16[12] = &v2;
  v16[16] = L"Unused.LowThreshold";
  v16[17] = (char *)&xmmword_1C0076360 + 8;
  v16[19] = &v3;
  v16[23] = L"Unused.NormalThreshold";
  v16[24] = &xmmword_1C0076370;
  v16[26] = &v4;
  LODWORD(v16[8]) = 288;
  LODWORD(v16[11]) = 67108868;
  LODWORD(v16[13]) = 4;
  v16[14] = 0LL;
  LODWORD(v16[15]) = 288;
  LODWORD(v16[18]) = 67108868;
  LODWORD(v16[20]) = 4;
  v16[21] = 0LL;
  LODWORD(v16[22]) = 288;
  LODWORD(v16[25]) = 67108868;
  LODWORD(v16[27]) = 4;
  v16[28] = 0LL;
  v16[30] = L"Unused.HighThreshold";
  v0 = &xmmword_1C0076420;
  LODWORD(v16[29]) = 288;
  v16[31] = (char *)&xmmword_1C0076370 + 8;
  v16[33] = &v5;
  v16[37] = L"Unused.MaximumThreshold";
  v16[38] = &xmmword_1C0076380;
  v16[40] = &v6;
  v16[44] = L"Unused.SelfTrimMinimumThreshold";
  v16[45] = (char *)&xmmword_1C0076380 + 8;
  v16[47] = &v7;
  v16[51] = L"Unused.SelfTrimLowThreshold";
  v16[52] = &xmmword_1C0076390;
  v16[54] = &v8;
  v16[58] = L"Unused.SelfTrimNormalThreshold";
  v16[59] = (char *)&xmmword_1C0076390 + 8;
  v16[61] = &v9;
  v16[65] = L"Unused.SelfTrimHighThreshold";
  v16[66] = &xmmword_1C00763A0;
  v16[68] = &v10;
  v16[72] = L"Unused.SelfTrimMaximumThreshold";
  v16[73] = (char *)&xmmword_1C00763A0 + 8;
  v16[75] = &v11;
  v16[79] = L"Unused.EvictApertureOfferLowThreshold";
  v16[82] = &v12;
  LODWORD(v16[32]) = 67108868;
  LODWORD(v16[34]) = 4;
  v16[35] = 0LL;
  LODWORD(v16[36]) = 288;
  LODWORD(v16[39]) = 67108868;
  LODWORD(v16[41]) = 4;
  v16[42] = 0LL;
  LODWORD(v16[43]) = 288;
  LODWORD(v16[46]) = 67108868;
  LODWORD(v16[48]) = 4;
  v16[49] = 0LL;
  LODWORD(v16[50]) = 288;
  LODWORD(v16[53]) = 67108868;
  LODWORD(v16[55]) = 4;
  v16[56] = 0LL;
  LODWORD(v16[57]) = 288;
  LODWORD(v16[60]) = 67108868;
  LODWORD(v16[62]) = 4;
  v16[63] = 0LL;
  LODWORD(v16[64]) = 288;
  LODWORD(v16[67]) = 67108868;
  LODWORD(v16[69]) = 4;
  v16[70] = 0LL;
  LODWORD(v16[71]) = 288;
  LODWORD(v16[74]) = 67108868;
  LODWORD(v16[76]) = 4;
  v16[77] = 0LL;
  LODWORD(v16[78]) = 288;
  v16[80] = &xmmword_1C0076420;
  LODWORD(v16[81]) = 67108868;
  LODWORD(v16[83]) = 4;
  v16[84] = 0LL;
  LODWORD(v16[85]) = 288;
  v16[86] = L"Unused.EvictApertureOfferNormalThreshold";
  LODWORD(v16[88]) = 67108868;
  v16[87] = (char *)&xmmword_1C0076420 + 8;
  v16[89] = &v13;
  v16[93] = L"Unused.EvictApertureOfferHighThreshold";
  v16[94] = &xmmword_1C0076430;
  v16[96] = &v14;
  v16[100] = L"Unused.EvictApertureOfferMaximumThreshold";
  v16[101] = (char *)&xmmword_1C0076430 + 8;
  LODWORD(v16[90]) = 4;
  LODWORD(v16[92]) = 288;
  LODWORD(v16[95]) = 67108868;
  LODWORD(v16[97]) = 4;
  LODWORD(v16[99]) = 288;
  LODWORD(v16[102]) = 67108868;
  LODWORD(v16[104]) = 4;
  v16[103] = &v15;
  v16[91] = 0LL;
  v16[98] = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", v16, 0LL, 0LL);
  qword_1C0076358 *= 10000000LL;
  *(_QWORD *)&xmmword_1C0076360 = 10000000 * xmmword_1C0076360;
  *((_QWORD *)&xmmword_1C0076360 + 1) *= 10000000LL;
  *(_QWORD *)&xmmword_1C0076370 = 10000000 * xmmword_1C0076370;
  *((_QWORD *)&xmmword_1C0076370 + 1) *= 10000000LL;
  *(_QWORD *)&xmmword_1C0076380 = 10000000 * xmmword_1C0076380;
  *((_QWORD *)&xmmword_1C0076380 + 1) *= 10000000LL;
  *(_QWORD *)&xmmword_1C0076390 = 10000000 * xmmword_1C0076390;
  *((_QWORD *)&xmmword_1C0076390 + 1) *= 10000000LL;
  *(_QWORD *)&xmmword_1C00763A0 = 10000000 * xmmword_1C00763A0;
  *((_QWORD *)&xmmword_1C00763A0 + 1) *= 10000000LL;
  do
  {
    *(_QWORD *)v0 *= 10000000LL;
    v0 = (__int128 *)((char *)v0 + 8);
  }
  while ( (__int64)v0 < (__int64)&qword_1C0076440 );
}