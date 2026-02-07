void __fastcall VIDMM_GLOBAL::ReadPhysicalAdapterConfiguration(VIDMM_GLOBAL *this, unsigned int a2)
{
  unsigned int v2; // ebx
  __int64 v3; // r14
  __int64 v4; // rcx
  unsigned int v5; // esi
  WCHAR *v6; // rdi
  __int64 v7; // rax
  __int64 v8; // rcx
  unsigned int v9; // eax
  unsigned __int64 v10; // rcx
  unsigned __int64 v11; // rdx
  __int64 v12; // rax
  unsigned int v13; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v14; // [rsp+34h] [rbp-CCh] BYREF
  int v15; // [rsp+38h] [rbp-C8h] BYREF
  int v16; // [rsp+3Ch] [rbp-C4h] BYREF
  int v17; // [rsp+40h] [rbp-C0h] BYREF
  int v18; // [rsp+44h] [rbp-BCh] BYREF
  PCUNICODE_STRING Source; // [rsp+48h] [rbp-B8h] BYREF
  struct _UNICODE_STRING Destination; // [rsp+50h] [rbp-B0h] BYREF
  WCHAR *v21; // [rsp+60h] [rbp-A0h] BYREF
  _BYTE v22[256]; // [rsp+68h] [rbp-98h] BYREF
  unsigned int v23; // [rsp+168h] [rbp+68h]
  __int64 v24; // [rsp+170h] [rbp+70h] BYREF
  int v25; // [rsp+178h] [rbp+78h]
  const wchar_t *v26; // [rsp+180h] [rbp+80h]
  unsigned int *v27; // [rsp+188h] [rbp+88h]
  int v28; // [rsp+190h] [rbp+90h]
  int *v29; // [rsp+198h] [rbp+98h]
  int v30; // [rsp+1A0h] [rbp+A0h]
  __int64 v31; // [rsp+1A8h] [rbp+A8h]
  int v32; // [rsp+1B0h] [rbp+B0h]
  const wchar_t *v33; // [rsp+1B8h] [rbp+B8h]
  unsigned int *v34; // [rsp+1C0h] [rbp+C0h]
  int v35; // [rsp+1C8h] [rbp+C8h]
  int *v36; // [rsp+1D0h] [rbp+D0h]
  int v37; // [rsp+1D8h] [rbp+D8h]
  __int64 v38; // [rsp+1E0h] [rbp+E0h]
  int v39; // [rsp+1E8h] [rbp+E8h]
  const wchar_t *v40; // [rsp+1F0h] [rbp+F0h]
  int *v41; // [rsp+1F8h] [rbp+F8h]
  int v42; // [rsp+200h] [rbp+100h]
  int *v43; // [rsp+208h] [rbp+108h]
  int v44; // [rsp+210h] [rbp+110h]
  __int128 v45; // [rsp+218h] [rbp+118h]
  __int128 v46; // [rsp+228h] [rbp+128h]
  __int128 v47; // [rsp+238h] [rbp+138h]
  __int64 v48; // [rsp+248h] [rbp+148h]

  v2 = 0;
  v3 = *((_QWORD *)this + 5028) + 1616LL * a2;
  v4 = *(_QWORD *)(*(_QWORD *)(*((_QWORD *)this + 3) + 2808LL) + 344LL * a2 + 8);
  Source = 0LL;
  DpiGetPnpRegistryKeyName(v4, a2, &Source);
  v5 = (Source->Length >> 1) + 16;
  v21 = 0LL;
  v23 = 0;
  if ( v5 > 0x80 )
  {
    if ( 0xFFFFFFFFFFFFFFFFuLL / v5 < 2 )
    {
      v6 = 0LL;
      goto LABEL_7;
    }
    v12 = 2LL * v5;
    if ( !is_mul_ok(v5, 2uLL) )
      v12 = -1LL;
    v6 = (WCHAR *)operator new[](v12, 1265072196LL, 256LL);
    v21 = v6;
  }
  else
  {
    v6 = (WCHAR *)v22;
    v21 = (WCHAR *)v22;
    if ( v5 )
    {
      v7 = 0LL;
      v8 = v5;
      do
      {
        v6[v7++] = 0;
        v6 = v21;
        --v8;
      }
      while ( v8 );
    }
  }
  v23 = v5;
  if ( v6 )
  {
    *(&Destination.MaximumLength + 2) = 0;
    *(_DWORD *)&Destination.MaximumLength = (unsigned __int16)(2 * v5);
    Destination.Buffer = v6;
    Destination.Length = 0;
    RtlAppendUnicodeStringToString(&Destination, Source);
    RtlAppendUnicodeToString(&Destination, L"\\MemoryManager");
  }
LABEL_7:
  v16 = 0;
  v9 = 0;
  v13 = 0;
  v17 = 0;
  v14 = 0;
  v18 = 0;
  v15 = 0;
  if ( v6 )
  {
    v24 = 0LL;
    v30 = 4;
    v25 = 288;
    v28 = 67108868;
    v32 = 288;
    v26 = L"MaxLocalSegmentSize";
    v35 = 67108868;
    v27 = &v13;
    v29 = &v16;
    v33 = L"MaxNonLocalSegmentSize";
    v34 = &v14;
    v36 = &v17;
    v40 = L"Supports64KBPages";
    v41 = &v15;
    v43 = &v18;
    v37 = 4;
    v39 = 288;
    v42 = 67108868;
    v44 = 4;
    v48 = 0LL;
    v31 = 0LL;
    v38 = 0LL;
    v45 = 0LL;
    v46 = 0LL;
    v47 = 0LL;
    RtlQueryRegistryValuesEx(0LL, v6, &v24, 0LL, 0LL);
    v9 = v14;
    v2 = v13;
  }
  v10 = (unsigned __int64)v9 << 20;
  v11 = (unsigned __int64)v2 << 20;
  if ( v11 - 1 <= 0xFFFFFFF )
    v11 = 0x10000000LL;
  *(_QWORD *)v3 = v11;
  if ( v10 - 1 <= 0x1FFFFFFF )
    v10 = 0x20000000LL;
  *(_QWORD *)(v3 + 8) = v10;
  *(_BYTE *)(v3 + 16) ^= (*(_BYTE *)(v3 + 16) ^ v15) & 1;
  PagedPoolArray<unsigned short,128>::~PagedPoolArray<unsigned short,128>(&v21);
}