void __fastcall DXGVALIDATION::InitializeBootSettings(DXGVALIDATION *this)
{
  int v2; // eax
  bool v3; // zf
  bool v4; // al
  bool v5; // al
  int v6; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v7; // [rsp+34h] [rbp-CCh] BYREF
  int v8; // [rsp+38h] [rbp-C8h] BYREF
  int v9; // [rsp+3Ch] [rbp-C4h] BYREF
  int v10; // [rsp+40h] [rbp-C0h] BYREF
  int v11; // [rsp+44h] [rbp-BCh] BYREF
  __int64 v12; // [rsp+50h] [rbp-B0h] BYREF
  int v13; // [rsp+58h] [rbp-A8h]
  const wchar_t *v14; // [rsp+60h] [rbp-A0h]
  unsigned int *v15; // [rsp+68h] [rbp-98h]
  int v16; // [rsp+70h] [rbp-90h]
  int *v17; // [rsp+78h] [rbp-88h]
  int v18; // [rsp+80h] [rbp-80h]
  __int64 v19; // [rsp+88h] [rbp-78h]
  int v20; // [rsp+90h] [rbp-70h]
  const wchar_t *v21; // [rsp+98h] [rbp-68h]
  int *v22; // [rsp+A0h] [rbp-60h]
  int v23; // [rsp+A8h] [rbp-58h]
  int *v24; // [rsp+B0h] [rbp-50h]
  int v25; // [rsp+B8h] [rbp-48h]
  __int64 v26; // [rsp+C0h] [rbp-40h]
  int v27; // [rsp+C8h] [rbp-38h]
  const wchar_t *v28; // [rsp+D0h] [rbp-30h]
  int *v29; // [rsp+D8h] [rbp-28h]
  int v30; // [rsp+E0h] [rbp-20h]
  int *v31; // [rsp+E8h] [rbp-18h]
  int v32; // [rsp+F0h] [rbp-10h]
  __int64 v33; // [rsp+F8h] [rbp-8h]
  int v34; // [rsp+100h] [rbp+0h]
  const wchar_t *v35; // [rsp+108h] [rbp+8h]
  int *v36; // [rsp+110h] [rbp+10h]
  int v37; // [rsp+118h] [rbp+18h]
  int *v38; // [rsp+120h] [rbp+20h]
  int v39; // [rsp+128h] [rbp+28h]
  __int64 v40; // [rsp+130h] [rbp+30h]
  int v41; // [rsp+138h] [rbp+38h]
  const wchar_t *v42; // [rsp+140h] [rbp+40h]
  int *v43; // [rsp+148h] [rbp+48h]
  int v44; // [rsp+150h] [rbp+50h]
  int *v45; // [rsp+158h] [rbp+58h]
  int v46; // [rsp+160h] [rbp+60h]
  __int64 v47; // [rsp+168h] [rbp+68h]
  int v48; // [rsp+170h] [rbp+70h]
  __int64 v49; // [rsp+178h] [rbp+78h]
  __int128 v50; // [rsp+180h] [rbp+80h]
  __int128 v51; // [rsp+190h] [rbp+90h]

  v14 = L"Level";
  v13 = 288;
  v16 = 67108868;
  v15 = &v7;
  v20 = 288;
  v18 = 4;
  v17 = &v6;
  v23 = 67108868;
  v21 = L"FailEscapeDDI";
  v25 = 4;
  v22 = &v8;
  v24 = &v6;
  v28 = L"FailRenderDDI";
  v29 = &v9;
  v31 = &v6;
  v35 = L"FailReserveGPUVA";
  v36 = &v10;
  v38 = &v6;
  v42 = L"ReportVirtualMachine";
  v43 = &v11;
  v27 = 288;
  v30 = 67108868;
  v32 = 4;
  v34 = 288;
  v37 = 67108868;
  v39 = 4;
  v41 = 288;
  v44 = 67108868;
  v46 = 4;
  v45 = &v6;
  v7 = 0;
  v8 = 0;
  v9 = 0;
  v10 = 0;
  v11 = 0;
  v6 = 0;
  v12 = 0LL;
  v19 = 0LL;
  v26 = 0LL;
  v33 = 0LL;
  v40 = 0LL;
  v47 = 0LL;
  v48 = 0;
  v49 = 0LL;
  v50 = 0LL;
  v51 = 0LL;
  if ( (int)RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\Validation", &v12) >= 0 )
  {
    v2 = v7;
    if ( v7 >= 2 )
      v2 = 2;
    *(_DWORD *)this = v2;
    if ( v2 )
    {
      v3 = v9 == 1;
      *((_BYTE *)this + 4) = v8 == 1;
      v4 = v3;
      v3 = v10 == 1;
      *((_BYTE *)this + 5) = v4;
      v5 = v3;
      v3 = v11 == 1;
      *((_BYTE *)this + 6) = v5;
      *((_BYTE *)this + 7) = v3;
    }
  }
}
