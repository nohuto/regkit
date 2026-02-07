void VIDMM_GLOBAL::ReadGpuVaConfiguration(void)
{
  LONG v0; // eax
  unsigned int v1; // [rsp+30h] [rbp-D0h] BYREF
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
  int v16; // [rsp+6Ch] [rbp-94h] BYREF
  __int64 v17; // [rsp+70h] [rbp-90h] BYREF
  int v18; // [rsp+78h] [rbp-88h]
  const wchar_t *v19; // [rsp+80h] [rbp-80h]
  int *v20; // [rsp+88h] [rbp-78h]
  int v21; // [rsp+90h] [rbp-70h]
  int *v22; // [rsp+98h] [rbp-68h]
  int v23; // [rsp+A0h] [rbp-60h]
  __int64 v24; // [rsp+A8h] [rbp-58h]
  int v25; // [rsp+B0h] [rbp-50h]
  const wchar_t *v26; // [rsp+B8h] [rbp-48h]
  int *v27; // [rsp+C0h] [rbp-40h]
  int v28; // [rsp+C8h] [rbp-38h]
  int *v29; // [rsp+D0h] [rbp-30h]
  int v30; // [rsp+D8h] [rbp-28h]
  __int64 v31; // [rsp+E0h] [rbp-20h]
  int v32; // [rsp+E8h] [rbp-18h]
  const wchar_t *v33; // [rsp+F0h] [rbp-10h]
  int *v34; // [rsp+F8h] [rbp-8h]
  int v35; // [rsp+100h] [rbp+0h]
  int *v36; // [rsp+108h] [rbp+8h]
  int v37; // [rsp+110h] [rbp+10h]
  __int64 v38; // [rsp+118h] [rbp+18h]
  int v39; // [rsp+120h] [rbp+20h]
  const wchar_t *v40; // [rsp+128h] [rbp+28h]
  int *v41; // [rsp+130h] [rbp+30h]
  int v42; // [rsp+138h] [rbp+38h]
  int *v43; // [rsp+140h] [rbp+40h]
  int v44; // [rsp+148h] [rbp+48h]
  __int64 v45; // [rsp+150h] [rbp+50h]
  int v46; // [rsp+158h] [rbp+58h]
  const wchar_t *v47; // [rsp+160h] [rbp+60h]
  int *v48; // [rsp+168h] [rbp+68h]
  int v49; // [rsp+170h] [rbp+70h]
  int *v50; // [rsp+178h] [rbp+78h]
  int v51; // [rsp+180h] [rbp+80h]
  __int64 v52; // [rsp+188h] [rbp+88h]
  int v53; // [rsp+190h] [rbp+90h]
  const wchar_t *v54; // [rsp+198h] [rbp+98h]
  int *v55; // [rsp+1A0h] [rbp+A0h]
  int v56; // [rsp+1A8h] [rbp+A8h]
  int *v57; // [rsp+1B0h] [rbp+B0h]
  int v58; // [rsp+1B8h] [rbp+B8h]
  __int64 v59; // [rsp+1C0h] [rbp+C0h]
  int v60; // [rsp+1C8h] [rbp+C8h]
  const wchar_t *v61; // [rsp+1D0h] [rbp+D0h]
  int *v62; // [rsp+1D8h] [rbp+D8h]
  int v63; // [rsp+1E0h] [rbp+E0h]
  int *v64; // [rsp+1E8h] [rbp+E8h]
  int v65; // [rsp+1F0h] [rbp+F0h]
  __int64 v66; // [rsp+1F8h] [rbp+F8h]
  int v67; // [rsp+200h] [rbp+100h]
  const wchar_t *v68; // [rsp+208h] [rbp+108h]
  unsigned int *v69; // [rsp+210h] [rbp+110h]
  int v70; // [rsp+218h] [rbp+118h]
  int *v71; // [rsp+220h] [rbp+120h]
  int v72; // [rsp+228h] [rbp+128h]
  __int128 v73; // [rsp+230h] [rbp+130h]
  __int128 v74; // [rsp+240h] [rbp+140h]
  __int128 v75; // [rsp+250h] [rbp+150h]
  __int64 v76; // [rsp+260h] [rbp+160h]

  v16 = 128;
  v9 = 0;
  v2 = 0;
  v10 = 0;
  v3 = 0;
  v12 = 30;
  v5 = 30;
  v13 = 0x10000;
  v6 = 0x10000;
  v19 = L"DisableUncommitGpuVaInPagingProcess";
  v20 = &v2;
  v22 = &v9;
  v26 = L"EnableZeroFlagInPde";
  v27 = &v3;
  v29 = &v10;
  v33 = L"DisableMakeIoMmuAddressValid";
  v34 = &v4;
  v36 = &v11;
  v40 = L"PagingProcessVaSpaceBitCount";
  v41 = &v5;
  v43 = &v12;
  v47 = L"GpuVaFirstValidAddress";
  v48 = &v6;
  v50 = &v13;
  v54 = L"EnableGpuVaGuardPages";
  v55 = &v7;
  v57 = &v14;
  v11 = 0;
  v4 = 0;
  v14 = 0;
  v7 = 0;
  v15 = 0;
  v8 = 0;
  v1 = 128;
  v17 = 0LL;
  v18 = 288;
  v21 = 67108868;
  v23 = 4;
  v24 = 0LL;
  v25 = 288;
  v28 = 67108868;
  v30 = 4;
  v31 = 0LL;
  v32 = 288;
  v35 = 67108868;
  v37 = 4;
  v38 = 0LL;
  v39 = 288;
  v42 = 67108868;
  v44 = 4;
  v45 = 0LL;
  v46 = 288;
  v49 = 67108868;
  v51 = 4;
  v52 = 0LL;
  v53 = 288;
  v56 = 67108868;
  v58 = 4;
  v60 = 288;
  v61 = L"AllocateGpuVaFromHighAddresses";
  v62 = &v8;
  v64 = &v15;
  v68 = L"CompanionContextMaxPendingOperations";
  v69 = &v1;
  v71 = &v16;
  v63 = 67108868;
  v65 = 4;
  v67 = 288;
  v70 = 67108868;
  v72 = 4;
  v59 = 0LL;
  v66 = 0LL;
  v73 = 0LL;
  v76 = 0LL;
  v74 = 0LL;
  v75 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\MemoryManager", &v17, 0LL, 0LL);
  dword_1400814F8 = v5;
  dword_140081508 = v6 & 0xFFFFF000;
  dword_1400815EC = v7;
  dword_1400815F0 = v8;
  v0 = 0x7FFFFFFF;
  VIDMM_GLOBAL::_Config = (v4 != 0 ? 0x20 : 0) | (v3 != 0 ? 0x100 : 0) | (v2 != 0 ? 0x80 : 0) | VIDMM_GLOBAL::_Config & 0xFFFFFE5F;
  if ( v1 < 0x7FFFFFFF )
    v0 = v1;
  Count = v0;
}