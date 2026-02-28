__int64 SmmQueryRegistry()
{
  int v0; // ebx
  __int64 v1; // rcx
  __int64 result; // rax
  unsigned int v3; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v4; // [rsp+34h] [rbp-CCh] BYREF
  unsigned int v5; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v6; // [rsp+3Ch] [rbp-C4h] BYREF
  int v7; // [rsp+40h] [rbp-C0h] BYREF
  int v8; // [rsp+44h] [rbp-BCh] BYREF
  int v9; // [rsp+48h] [rbp-B8h] BYREF
  int v10; // [rsp+4Ch] [rbp-B4h] BYREF
  int v11; // [rsp+50h] [rbp-B0h] BYREF
  int v12; // [rsp+54h] [rbp-ACh] BYREF
  int v13; // [rsp+58h] [rbp-A8h] BYREF
  int v14; // [rsp+5Ch] [rbp-A4h] BYREF
  int v15; // [rsp+60h] [rbp-A0h] BYREF
  int v16; // [rsp+64h] [rbp-9Ch] BYREF
  __int64 v17; // [rsp+70h] [rbp-90h] BYREF
  int v18; // [rsp+78h] [rbp-88h]
  const wchar_t *v19; // [rsp+80h] [rbp-80h]
  unsigned int *v20; // [rsp+88h] [rbp-78h]
  int v21; // [rsp+90h] [rbp-70h]
  unsigned int *v22; // [rsp+98h] [rbp-68h]
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
  unsigned int *v34; // [rsp+F8h] [rbp-8h]
  int v35; // [rsp+100h] [rbp+0h]
  unsigned int *v36; // [rsp+108h] [rbp+8h]
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
  __int128 v66; // [rsp+1F8h] [rbp+F8h]
  __int128 v67; // [rsp+208h] [rbp+108h]
  __int128 v68; // [rsp+218h] [rbp+118h]
  __int64 v69; // [rsp+228h] [rbp+128h]

  v0 = 0;
  v19 = L"ForceEnableIommu";
  v6 = 0;
  v3 = 0;
  v20 = &v3;
  v12 = 0;
  v22 = &v6;
  v26 = L"EnablePageTracking";
  v27 = &v8;
  v29 = &v12;
  v33 = L"LogicalAddressMode";
  v34 = &v4;
  v36 = &v5;
  v40 = L"PreferHighLogicalAddresses";
  v41 = &v10;
  v43 = &v13;
  v47 = L"DebugMode";
  v48 = &v11;
  v50 = &v14;
  v54 = L"IdentityMappedPassthrough";
  v55 = &v7;
  v57 = &v15;
  v8 = 0;
  v5 = 0;
  v4 = 0;
  v13 = 0;
  v10 = 0;
  v16 = 0;
  v9 = 0;
  v14 = 0;
  v11 = 0;
  v15 = 0;
  v7 = 0;
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
  v59 = 0LL;
  v60 = 288;
  v61 = L"ForceDmaRemapping";
  v63 = 67108868;
  v65 = 4;
  v62 = &v9;
  v64 = &v16;
  v69 = 0LL;
  v66 = 0LL;
  v67 = 0LL;
  v68 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\Smm", &v17);
  if ( v4 >= 3 )
    v4 = v5;
  if ( v3 >= 3 )
    v3 = v6;
  if ( !(unsigned __int8)HviIsHypervisorMicrosoftCompatible(v1) && SmmGetIommuInterfaceVersion() >= 2 )
    v0 = v7;
  result = v0 != 0 ? 0x400 : 0;
  dword_1C0154380 = result | (v8 != 0 ? 4 : 0) | v3 & 3 | dword_1C0154380 & 0xFFFFFB00 | (unsigned __int8)(8 * (v4 & 3 | (4 * (v11 & 1 | (2 * (v10 & 1 | (2 * (v9 & 1))))))));
  return result;
}
