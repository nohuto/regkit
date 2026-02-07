void VIDMM_GLOBAL::ReadVPRConfiguration(void)
{
  unsigned int v0; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v1; // [rsp+34h] [rbp-CCh] BYREF
  unsigned int v2; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v3; // [rsp+3Ch] [rbp-C4h] BYREF
  int v4; // [rsp+40h] [rbp-C0h] BYREF
  int v5; // [rsp+44h] [rbp-BCh] BYREF
  int v6; // [rsp+48h] [rbp-B8h] BYREF
  int v7; // [rsp+4Ch] [rbp-B4h] BYREF
  __int64 v8; // [rsp+50h] [rbp-B0h] BYREF
  int v9; // [rsp+58h] [rbp-A8h]
  const wchar_t *v10; // [rsp+60h] [rbp-A0h]
  unsigned int *v11; // [rsp+68h] [rbp-98h]
  int v12; // [rsp+70h] [rbp-90h]
  int *v13; // [rsp+78h] [rbp-88h]
  int v14; // [rsp+80h] [rbp-80h]
  __int64 v15; // [rsp+88h] [rbp-78h]
  int v16; // [rsp+90h] [rbp-70h]
  const wchar_t *v17; // [rsp+98h] [rbp-68h]
  unsigned int *v18; // [rsp+A0h] [rbp-60h]
  int v19; // [rsp+A8h] [rbp-58h]
  int *v20; // [rsp+B0h] [rbp-50h]
  int v21; // [rsp+B8h] [rbp-48h]
  __int64 v22; // [rsp+C0h] [rbp-40h]
  int v23; // [rsp+C8h] [rbp-38h]
  const wchar_t *v24; // [rsp+D0h] [rbp-30h]
  unsigned int *v25; // [rsp+D8h] [rbp-28h]
  int v26; // [rsp+E0h] [rbp-20h]
  int *v27; // [rsp+E8h] [rbp-18h]
  int v28; // [rsp+F0h] [rbp-10h]
  __int64 v29; // [rsp+F8h] [rbp-8h]
  int v30; // [rsp+100h] [rbp+0h]
  const wchar_t *v31; // [rsp+108h] [rbp+8h]
  unsigned int *v32; // [rsp+110h] [rbp+10h]
  int v33; // [rsp+118h] [rbp+18h]
  int *v34; // [rsp+120h] [rbp+20h]
  int v35; // [rsp+128h] [rbp+28h]
  __int128 v36; // [rsp+130h] [rbp+30h]
  __int128 v37; // [rsp+140h] [rbp+40h]
  __int128 v38; // [rsp+150h] [rbp+50h]
  __int64 v39; // [rsp+160h] [rbp+60h]

  v8 = 0LL;
  v15 = 0LL;
  v22 = 0LL;
  v29 = 0LL;
  v9 = 288;
  v12 = 67108868;
  v6 = 1;
  v3 = 1;
  v16 = 288;
  v10 = L"VPRGrowRatioNumerator";
  v11 = &v1;
  v13 = &v4;
  v17 = L"VPRGrowRatioDenominator";
  v18 = &v0;
  v20 = &v5;
  v24 = L"VPRCapacityRatioNumerator";
  v25 = &v3;
  v27 = &v6;
  v31 = L"VPRCapacityRatioDenominator";
  v32 = &v2;
  v34 = &v7;
  v19 = 67108868;
  v23 = 288;
  v26 = 67108868;
  v30 = 288;
  v33 = 67108868;
  v39 = 0LL;
  v4 = 4;
  v1 = 4;
  v5 = 5;
  v0 = 5;
  v7 = 5;
  v2 = 5;
  v14 = 4;
  v21 = 4;
  v28 = 4;
  v35 = 4;
  v36 = 0LL;
  v37 = 0LL;
  v38 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v8, 0LL, 0LL);
  if ( v0 > v1 && v1 )
  {
    dword_1C0076578 = v1;
    dword_1C007657C = v0;
  }
  else
  {
    dword_1C0076578 = 4;
    dword_1C007657C = 5;
  }
  if ( v2 > v3 && v3 )
  {
    dword_1C0076580 = v3;
    dword_1C0076584 = v2;
  }
  else
  {
    dword_1C0076580 = 4;
    dword_1C0076584 = 5;
  }
}