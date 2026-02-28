void __fastcall TIMESTAMP::ReadRegistryParameters(TIMESTAMP *this, struct ADAPTER_CONTEXT *a2, void *a3)
{
  REGISTRY *v3; // rcx
  int v4; // [rsp+30h] [rbp-D0h] BYREF
  const wchar_t *v5; // [rsp+38h] [rbp-C8h]
  __int64 v6; // [rsp+40h] [rbp-C0h]
  int v7; // [rsp+48h] [rbp-B8h]
  __int64 v8; // [rsp+4Ch] [rbp-B4h]
  int v9; // [rsp+54h] [rbp-ACh]
  int v10; // [rsp+58h] [rbp-A8h]
  __int16 v11; // [rsp+5Ch] [rbp-A4h]
  int v12; // [rsp+60h] [rbp-A0h]
  const wchar_t *v13; // [rsp+68h] [rbp-98h]
  __int64 v14; // [rsp+70h] [rbp-90h]
  int v15; // [rsp+78h] [rbp-88h]
  __int64 v16; // [rsp+7Ch] [rbp-84h]
  __int64 v17; // [rsp+84h] [rbp-7Ch]
  __int16 v18; // [rsp+8Ch] [rbp-74h]
  int v19; // [rsp+90h] [rbp-70h]
  const wchar_t *v20; // [rsp+98h] [rbp-68h]
  __int64 v21; // [rsp+A0h] [rbp-60h]
  int v22; // [rsp+A8h] [rbp-58h]
  __int64 v23; // [rsp+ACh] [rbp-54h]
  int v24; // [rsp+B4h] [rbp-4Ch]
  int v25; // [rsp+B8h] [rbp-48h]
  __int16 v26; // [rsp+BCh] [rbp-44h]
  int v27; // [rsp+C0h] [rbp-40h]
  const wchar_t *v28; // [rsp+C8h] [rbp-38h]
  __int64 v29; // [rsp+D0h] [rbp-30h]
  int v30; // [rsp+D8h] [rbp-28h]
  __int64 v31; // [rsp+DCh] [rbp-24h]
  __int64 v32; // [rsp+E4h] [rbp-1Ch]
  __int16 v33; // [rsp+ECh] [rbp-14h]
  int v34; // [rsp+F0h] [rbp-10h]
  const wchar_t *v35; // [rsp+F8h] [rbp-8h]
  __int64 v36; // [rsp+100h] [rbp+0h]
  int v37; // [rsp+108h] [rbp+8h]
  __int64 v38; // [rsp+10Ch] [rbp+Ch]
  __int64 v39; // [rsp+114h] [rbp+14h]
  __int16 v40; // [rsp+11Ch] [rbp+1Ch]
  int v41; // [rsp+120h] [rbp+20h]
  const wchar_t *v42; // [rsp+128h] [rbp+28h]
  __int64 v43; // [rsp+130h] [rbp+30h]
  int v44; // [rsp+138h] [rbp+38h]
  __int64 v45; // [rsp+13Ch] [rbp+3Ch]
  __int64 v46; // [rsp+144h] [rbp+44h]
  __int16 v47; // [rsp+14Ch] [rbp+4Ch]

  v4 = 2490404;
  v6 = 0LL;
  v5 = L"AdvertiseTimestamp";
  v7 = 118096;
  v13 = L"AllTransmitHw";
  v8 = 1LL;
  v20 = L"TaggedTransmitHw";
  v9 = 1;
  v28 = L"*PtpHardwareTimestamp";
  v35 = L"*SoftwareTimestamp";
  v42 = L"TimeSync";
  v3 = (REGISTRY *)*((_QWORD *)a2 + 14912);
  v10 = 1;
  v11 = 256;
  v12 = 1835034;
  v14 = 0LL;
  v15 = 118104;
  v16 = 1LL;
  v17 = 1LL;
  v18 = 256;
  v19 = 2228256;
  v21 = 0LL;
  v22 = 118105;
  v23 = 1LL;
  v24 = 1;
  v25 = 1;
  v26 = 256;
  v27 = 2883626;
  v29 = 0LL;
  v30 = 118097;
  v31 = 1LL;
  v32 = 1LL;
  v33 = 256;
  v34 = 2490404;
  v36 = 0LL;
  v37 = 118100;
  v38 = 4LL;
  v39 = 5LL;
  v40 = 256;
  v41 = 1179664;
  v43 = 0LL;
  v44 = 118106;
  v45 = 1LL;
  v46 = 1LL;
  v47 = 256;
  REGISTRY::RegReadRegTable(v3, a2, a3, (struct REGTABLE_ENTRY *)&v4, 6u);
}
