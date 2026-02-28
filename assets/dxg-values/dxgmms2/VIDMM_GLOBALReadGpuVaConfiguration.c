void VIDMM_GLOBAL::ReadGpuVaConfiguration(void)
{
  int v0; // [rsp+30h] [rbp-D0h] BYREF
  int v1; // [rsp+34h] [rbp-CCh] BYREF
  int v2; // [rsp+38h] [rbp-C8h] BYREF
  int v3; // [rsp+3Ch] [rbp-C4h] BYREF
  int v4; // [rsp+40h] [rbp-C0h] BYREF
  int v5; // [rsp+44h] [rbp-BCh] BYREF
  int v6; // [rsp+48h] [rbp-B8h] BYREF
  int v7; // [rsp+4Ch] [rbp-B4h] BYREF
  int v8; // [rsp+50h] [rbp-B0h] BYREF
  int v9; // [rsp+54h] [rbp-ACh] BYREF
  int v10; // [rsp+58h] [rbp-A8h] BYREF
  int v11; // [rsp+5Ch] [rbp-A4h] BYREF
  int v12; // [rsp+60h] [rbp-A0h] BYREF
  int v13; // [rsp+64h] [rbp-9Ch] BYREF
  __int64 v14; // [rsp+70h] [rbp-90h] BYREF
  int v15; // [rsp+78h] [rbp-88h]
  const wchar_t *v16; // [rsp+80h] [rbp-80h]
  int *v17; // [rsp+88h] [rbp-78h]
  int v18; // [rsp+90h] [rbp-70h]
  int *v19; // [rsp+98h] [rbp-68h]
  int v20; // [rsp+A0h] [rbp-60h]
  __int64 v21; // [rsp+A8h] [rbp-58h]
  int v22; // [rsp+B0h] [rbp-50h]
  const wchar_t *v23; // [rsp+B8h] [rbp-48h]
  int *v24; // [rsp+C0h] [rbp-40h]
  int v25; // [rsp+C8h] [rbp-38h]
  int *v26; // [rsp+D0h] [rbp-30h]
  int v27; // [rsp+D8h] [rbp-28h]
  __int64 v28; // [rsp+E0h] [rbp-20h]
  int v29; // [rsp+E8h] [rbp-18h]
  const wchar_t *v30; // [rsp+F0h] [rbp-10h]
  int *v31; // [rsp+F8h] [rbp-8h]
  int v32; // [rsp+100h] [rbp+0h]
  int *v33; // [rsp+108h] [rbp+8h]
  int v34; // [rsp+110h] [rbp+10h]
  __int64 v35; // [rsp+118h] [rbp+18h]
  int v36; // [rsp+120h] [rbp+20h]
  const wchar_t *v37; // [rsp+128h] [rbp+28h]
  int *v38; // [rsp+130h] [rbp+30h]
  int v39; // [rsp+138h] [rbp+38h]
  int *v40; // [rsp+140h] [rbp+40h]
  int v41; // [rsp+148h] [rbp+48h]
  __int64 v42; // [rsp+150h] [rbp+50h]
  int v43; // [rsp+158h] [rbp+58h]
  const wchar_t *v44; // [rsp+160h] [rbp+60h]
  int *v45; // [rsp+168h] [rbp+68h]
  int v46; // [rsp+170h] [rbp+70h]
  int *v47; // [rsp+178h] [rbp+78h]
  int v48; // [rsp+180h] [rbp+80h]
  __int64 v49; // [rsp+188h] [rbp+88h]
  int v50; // [rsp+190h] [rbp+90h]
  const wchar_t *v51; // [rsp+198h] [rbp+98h]
  int *v52; // [rsp+1A0h] [rbp+A0h]
  int v53; // [rsp+1A8h] [rbp+A8h]
  int *v54; // [rsp+1B0h] [rbp+B0h]
  int v55; // [rsp+1B8h] [rbp+B8h]
  __int64 v56; // [rsp+1C0h] [rbp+C0h]
  int v57; // [rsp+1C8h] [rbp+C8h]
  const wchar_t *v58; // [rsp+1D0h] [rbp+D0h]
  int *v59; // [rsp+1D8h] [rbp+D8h]
  int v60; // [rsp+1E0h] [rbp+E0h]
  int *v61; // [rsp+1E8h] [rbp+E8h]
  int v62; // [rsp+1F0h] [rbp+F0h]
  __int128 v63; // [rsp+1F8h] [rbp+F8h]
  __int128 v64; // [rsp+208h] [rbp+108h]
  __int128 v65; // [rsp+218h] [rbp+118h]
  __int64 v66; // [rsp+228h] [rbp+128h]

  v7 = 0;
  v0 = 0;
  v8 = 0;
  v1 = 0;
  v9 = 0;
  v10 = 30;
  v3 = 30;
  v11 = 0x10000;
  v4 = 0x10000;
  v16 = L"DisableUncommitGpuVaInPagingProcess";
  v17 = &v0;
  v19 = &v7;
  v23 = L"EnableZeroFlagInPde";
  v24 = &v1;
  v26 = &v8;
  v30 = L"DisableMakeIoMmuAddressValid";
  v31 = &v2;
  v33 = &v9;
  v37 = L"PagingProcessVaSpaceBitCount";
  v38 = &v3;
  v40 = &v10;
  v44 = L"GpuVaFirstValidAddress";
  v45 = &v4;
  v47 = &v11;
  v51 = L"EnableGpuVaGuardPages";
  v52 = &v5;
  v54 = &v12;
  v2 = 0;
  v12 = 0;
  v5 = 0;
  v13 = 0;
  v6 = 0;
  v14 = 0LL;
  v15 = 288;
  v18 = 67108868;
  v20 = 4;
  v21 = 0LL;
  v22 = 288;
  v25 = 67108868;
  v27 = 4;
  v28 = 0LL;
  v29 = 288;
  v32 = 67108868;
  v34 = 4;
  v35 = 0LL;
  v36 = 288;
  v39 = 67108868;
  v41 = 4;
  v42 = 0LL;
  v43 = 288;
  v46 = 67108868;
  v48 = 4;
  v49 = 0LL;
  v50 = 288;
  v53 = 67108868;
  v55 = 4;
  v56 = 0LL;
  v57 = 288;
  v60 = 67108868;
  v58 = L"AllocateGpuVaFromHighAddresses";
  v62 = 4;
  v59 = &v6;
  v61 = &v13;
  v63 = 0LL;
  v66 = 0LL;
  v64 = 0LL;
  v65 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v14, 0LL, 0LL);
  dword_1C0076458 = v3;
  VIDMM_GLOBAL::_Config = (v2 != 0 ? 0x20 : 0) | (v1 != 0 ? 0x100 : 0) | (v0 != 0 ? 0x80 : 0) | VIDMM_GLOBAL::_Config & 0xFFFFFE5F;
  dword_1C0076468 = v4 & 0xFFFFF000;
  dword_1C0076540 = v5;
  dword_1C0076544 = v6;
}