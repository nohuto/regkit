void __fastcall VIDMM_GLOBAL::ReadPhysicalAdapterConfiguration(VIDMM_GLOBAL *this, unsigned int a2)
{
  __int64 v2; // rsi
  __int64 v3; // rcx
  unsigned int v4; // edi
  WCHAR *v5; // rbx
  __int64 v6; // rax
  __int64 v7; // rax
  __int64 v8; // rcx
  __int64 v9; // rcx
  unsigned int v10; // eax
  __int64 v11; // rcx
  unsigned __int64 v12; // rdx
  char v13; // al
  char v14; // cl
  unsigned int v15; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v16; // [rsp+34h] [rbp-CCh] BYREF
  unsigned int v17; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v18; // [rsp+3Ch] [rbp-C4h] BYREF
  int v19; // [rsp+40h] [rbp-C0h] BYREF
  int v20; // [rsp+44h] [rbp-BCh] BYREF
  int v21; // [rsp+48h] [rbp-B8h] BYREF
  int v22; // [rsp+4Ch] [rbp-B4h] BYREF
  int v23; // [rsp+50h] [rbp-B0h] BYREF
  int v24; // [rsp+54h] [rbp-ACh] BYREF
  int v25; // [rsp+58h] [rbp-A8h] BYREF
  int v26; // [rsp+5Ch] [rbp-A4h] BYREF
  PCUNICODE_STRING Source; // [rsp+60h] [rbp-A0h] BYREF
  struct _UNICODE_STRING Destination; // [rsp+68h] [rbp-98h] BYREF
  WCHAR *v29; // [rsp+80h] [rbp-80h] BYREF
  _BYTE v30[256]; // [rsp+88h] [rbp-78h] BYREF
  unsigned int v31; // [rsp+188h] [rbp+88h]
  __int64 v32; // [rsp+190h] [rbp+90h] BYREF
  int v33; // [rsp+198h] [rbp+98h]
  const wchar_t *v34; // [rsp+1A0h] [rbp+A0h]
  unsigned int *v35; // [rsp+1A8h] [rbp+A8h]
  int v36; // [rsp+1B0h] [rbp+B0h]
  int *v37; // [rsp+1B8h] [rbp+B8h]
  int v38; // [rsp+1C0h] [rbp+C0h]
  __int64 v39; // [rsp+1C8h] [rbp+C8h]
  int v40; // [rsp+1D0h] [rbp+D0h]
  const wchar_t *v41; // [rsp+1D8h] [rbp+D8h]
  unsigned int *v42; // [rsp+1E0h] [rbp+E0h]
  int v43; // [rsp+1E8h] [rbp+E8h]
  int *v44; // [rsp+1F0h] [rbp+F0h]
  int v45; // [rsp+1F8h] [rbp+F8h]
  __int64 v46; // [rsp+200h] [rbp+100h]
  int v47; // [rsp+208h] [rbp+108h]
  const wchar_t *v48; // [rsp+210h] [rbp+110h]
  unsigned int *v49; // [rsp+218h] [rbp+118h]
  int v50; // [rsp+220h] [rbp+120h]
  int *v51; // [rsp+228h] [rbp+128h]
  int v52; // [rsp+230h] [rbp+130h]
  __int64 v53; // [rsp+238h] [rbp+138h]
  int v54; // [rsp+240h] [rbp+140h]
  const wchar_t *v55; // [rsp+248h] [rbp+148h]
  unsigned int *v56; // [rsp+250h] [rbp+150h]
  int v57; // [rsp+258h] [rbp+158h]
  int *v58; // [rsp+260h] [rbp+160h]
  int v59; // [rsp+268h] [rbp+168h]
  __int64 v60; // [rsp+270h] [rbp+170h]
  int v61; // [rsp+278h] [rbp+178h]
  const wchar_t *v62; // [rsp+280h] [rbp+180h]
  int *v63; // [rsp+288h] [rbp+188h]
  int v64; // [rsp+290h] [rbp+190h]
  int *v65; // [rsp+298h] [rbp+198h]
  int v66; // [rsp+2A0h] [rbp+1A0h]
  __int64 v67; // [rsp+2A8h] [rbp+1A8h]
  int v68; // [rsp+2B0h] [rbp+1B0h]
  const wchar_t *v69; // [rsp+2B8h] [rbp+1B8h]
  int *v70; // [rsp+2C0h] [rbp+1C0h]
  int v71; // [rsp+2C8h] [rbp+1C8h]
  int *v72; // [rsp+2D0h] [rbp+1D0h]
  int v73; // [rsp+2D8h] [rbp+1D8h]
  __int128 v74; // [rsp+2E0h] [rbp+1E0h]
  __int128 v75; // [rsp+2F0h] [rbp+1F0h]
  __int128 v76; // [rsp+300h] [rbp+200h]
  __int64 v77; // [rsp+310h] [rbp+210h]

  v2 = *(_QWORD *)(*((_QWORD *)this + 5029) + 8LL * a2);
  v3 = *(_QWORD *)(*(_QWORD *)(*((_QWORD *)this + 3) + 2992LL) + 344LL * a2 + 8);
  Source = 0LL;
  DpiGetPnpRegistryKeyName(v3, a2, &Source);
  v4 = (Source->Length >> 1) + 16;
  v29 = 0LL;
  v31 = 0;
  if ( v4 <= 0x80 )
  {
    v5 = (WCHAR *)v30;
    v29 = (WCHAR *)v30;
    if ( v4 )
    {
      v7 = 0LL;
      v8 = v4;
      do
      {
        v5[v7++] = 0;
        v5 = v29;
        --v8;
      }
      while ( v8 );
    }
  }
  else
  {
    if ( 0xFFFFFFFFFFFFFFFFuLL / v4 < 2 )
    {
      v5 = 0LL;
      goto LABEL_12;
    }
    v6 = 2LL * v4;
    if ( !is_mul_ok(v4, 2uLL) )
      v6 = -1LL;
    v5 = (WCHAR *)operator new[](v6, 1265072196LL, 256LL);
    v29 = v5;
  }
  v31 = v4;
  if ( v5 )
  {
    *(&Destination.MaximumLength + 2) = 0;
    *(_DWORD *)&Destination.MaximumLength = (unsigned __int16)(2 * v4);
    Destination.Buffer = v5;
    Destination.Length = 0;
    RtlAppendUnicodeStringToString(&Destination, Source);
    RtlAppendUnicodeToString(&Destination, L"\\MemoryManager");
  }
LABEL_12:
  v9 = 0LL;
  v16 = 0;
  v10 = 0;
  v15 = 0;
  v23 = 900;
  v17 = 900;
  v24 = 900;
  v18 = 900;
  v21 = 0;
  v22 = 0;
  v25 = 0;
  v19 = 0;
  v26 = 1;
  v20 = 1;
  if ( v5 )
  {
    v32 = 0LL;
    v38 = 4;
    v33 = 288;
    v36 = 67108868;
    v40 = 288;
    v34 = L"MaxLocalSegmentSize";
    v43 = 67108868;
    v35 = &v15;
    v37 = &v21;
    v41 = L"MaxNonLocalSegmentSize";
    v42 = &v16;
    v44 = &v22;
    v48 = L"SelfRefreshVramForceEvictionTimerDC";
    v49 = &v17;
    v51 = &v23;
    v55 = L"SelfRefreshVramForceEvictionTimerAC";
    v56 = &v18;
    v58 = &v24;
    v62 = L"Supports64KBPages";
    v63 = &v19;
    v65 = &v25;
    v69 = L"EnablePromotion";
    v70 = &v20;
    v72 = &v26;
    v45 = 4;
    v47 = 288;
    v50 = 67108868;
    v52 = 4;
    v54 = 288;
    v57 = 67108868;
    v59 = 4;
    v61 = 288;
    v64 = 67108868;
    v66 = 4;
    v68 = 288;
    v71 = 67108868;
    v73 = 4;
    v77 = 0LL;
    v39 = 0LL;
    v46 = 0LL;
    v53 = 0LL;
    v60 = 0LL;
    v67 = 0LL;
    v74 = 0LL;
    v75 = 0LL;
    v76 = 0LL;
    RtlQueryRegistryValuesEx(0LL, v5, &v32, 0LL, 0LL);
    v10 = v15;
    v9 = v16;
  }
  v11 = v9 << 20;
  v12 = (unsigned __int64)v10 << 20;
  if ( v12 - 1 <= 0xFFFFFFF )
    v12 = 0x10000000LL;
  *(_QWORD *)v2 = v12;
  if ( (unsigned __int64)(v11 - 1) <= 0x1FFFFFFF )
    v11 = 0x20000000LL;
  *(_QWORD *)(v2 + 8) = v11;
  *(_QWORD *)(v2 + 16) = 10000000LL * v17;
  v13 = *(_BYTE *)(v2 + 36);
  *(_QWORD *)(v2 + 24) = 10000000LL * v18;
  v14 = v19 & 1;
  *(_DWORD *)(v2 + 32) = 0;
  *(_BYTE *)(v2 + 36) = v13 & 0xFE | v14;
  if ( (unsigned int)Feature_Servicing_GraphicsKernel_PromotionRegistryKeyPerAdapter__private_IsEnabledDeviceUsageNoInline() )
    *(_BYTE *)(v2 + 36) = *(_BYTE *)(v2 + 36) & 0xFD | (2 * (v20 & 1));
  PagedPoolArray<unsigned short,128>::~PagedPoolArray<unsigned short,128>(&v29);
}