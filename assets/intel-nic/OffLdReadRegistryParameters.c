void __fastcall OffLdReadRegistryParameters(REGISTRY **DstBuf, NDIS_HANDLE ConfigurationHandle)
{
  int v4; // [rsp+50h] [rbp-B0h] BYREF
  const wchar_t *v5; // [rsp+58h] [rbp-A8h]
  __int64 v6; // [rsp+60h] [rbp-A0h]
  int v7; // [rsp+68h] [rbp-98h]
  __int64 v8; // [rsp+6Ch] [rbp-94h]
  int v9; // [rsp+74h] [rbp-8Ch]
  int v10; // [rsp+78h] [rbp-88h]
  __int16 v11; // [rsp+7Ch] [rbp-84h]
  int v12; // [rsp+80h] [rbp-80h]
  const wchar_t *v13; // [rsp+88h] [rbp-78h]
  __int64 v14; // [rsp+90h] [rbp-70h]
  int v15; // [rsp+98h] [rbp-68h]
  __int64 v16; // [rsp+9Ch] [rbp-64h]
  int v17; // [rsp+A4h] [rbp-5Ch]
  int v18; // [rsp+A8h] [rbp-58h]
  __int16 v19; // [rsp+ACh] [rbp-54h]
  int v20; // [rsp+B0h] [rbp-50h]
  const wchar_t *v21; // [rsp+B8h] [rbp-48h]
  __int64 v22; // [rsp+C0h] [rbp-40h]
  int v23; // [rsp+C8h] [rbp-38h]
  __int64 v24; // [rsp+CCh] [rbp-34h]
  int v25; // [rsp+D4h] [rbp-2Ch]
  int v26; // [rsp+D8h] [rbp-28h]
  __int16 v27; // [rsp+DCh] [rbp-24h]
  int v28; // [rsp+E0h] [rbp-20h]
  const wchar_t *v29; // [rsp+E8h] [rbp-18h]
  __int64 v30; // [rsp+F0h] [rbp-10h]
  int v31; // [rsp+F8h] [rbp-8h]
  __int64 v32; // [rsp+FCh] [rbp-4h]
  int v33; // [rsp+104h] [rbp+4h]
  int v34; // [rsp+108h] [rbp+8h]
  __int16 v35; // [rsp+10Ch] [rbp+Ch]
  int v36; // [rsp+110h] [rbp+10h]
  const wchar_t *v37; // [rsp+118h] [rbp+18h]
  __int64 v38; // [rsp+120h] [rbp+20h]
  int v39; // [rsp+128h] [rbp+28h]
  __int64 v40; // [rsp+12Ch] [rbp+2Ch]
  int v41; // [rsp+134h] [rbp+34h]
  int v42; // [rsp+138h] [rbp+38h]
  __int16 v43; // [rsp+13Ch] [rbp+3Ch]
  int v44; // [rsp+140h] [rbp+40h]
  const wchar_t *v45; // [rsp+148h] [rbp+48h]
  __int64 v46; // [rsp+150h] [rbp+50h]
  int v47; // [rsp+158h] [rbp+58h]
  __int64 v48; // [rsp+15Ch] [rbp+5Ch]
  __int64 v49; // [rsp+164h] [rbp+64h]
  __int16 v50; // [rsp+16Ch] [rbp+6Ch]
  int v51; // [rsp+170h] [rbp+70h]
  const wchar_t *v52; // [rsp+178h] [rbp+78h]
  __int64 v53; // [rsp+180h] [rbp+80h]
  int v54; // [rsp+188h] [rbp+88h]
  __int64 v55; // [rsp+18Ch] [rbp+8Ch]
  int v56; // [rsp+194h] [rbp+94h]
  int v57; // [rsp+198h] [rbp+98h]
  __int16 v58; // [rsp+19Ch] [rbp+9Ch]
  int v59; // [rsp+1A0h] [rbp+A0h]
  const wchar_t *v60; // [rsp+1A8h] [rbp+A8h]
  __int64 v61; // [rsp+1B0h] [rbp+B0h]
  int v62; // [rsp+1B8h] [rbp+B8h]
  __int64 v63; // [rsp+1BCh] [rbp+BCh]
  int v64; // [rsp+1C4h] [rbp+C4h]
  int v65; // [rsp+1C8h] [rbp+C8h]
  __int16 v66; // [rsp+1CCh] [rbp+CCh]

  v4 = 3014700;
  v5 = L"*IPChecksumOffloadIPv4";
  v13 = L"*TCPChecksumOffloadIPv4";
  v6 = 0LL;
  v21 = L"*TCPChecksumOffloadIPv6";
  v7 = 119368;
  v29 = L"*UDPChecksumOffloadIPv4";
  v8 = 4LL;
  v37 = L"*UDPChecksumOffloadIPv6";
  v45 = L"*LsoV1IPv4";
  v52 = L"*LsoV2IPv4";
  v60 = L"*LsoV2IPv6";
  v9 = 3;
  v10 = 3;
  v11 = 256;
  v12 = 3145774;
  v14 = 0LL;
  v15 = 119372;
  v16 = 4LL;
  v17 = 3;
  v18 = 3;
  v19 = 256;
  v20 = 3145774;
  v22 = 0LL;
  v23 = 119380;
  v24 = 4LL;
  v25 = 3;
  v26 = 3;
  v27 = 256;
  v28 = 3145774;
  v30 = 0LL;
  v31 = 119376;
  v32 = 4LL;
  v33 = 3;
  v34 = 3;
  v35 = 256;
  v36 = 3145774;
  v38 = 0LL;
  v39 = 119384;
  v40 = 4LL;
  v41 = 3;
  v42 = 3;
  v43 = 256;
  v44 = 1441812;
  v46 = 0LL;
  v47 = 119388;
  v48 = 1LL;
  v49 = 1LL;
  v50 = 256;
  v51 = 1441812;
  v53 = 0LL;
  v54 = 119389;
  v55 = 1LL;
  v56 = 1;
  v57 = 1;
  v58 = 256;
  v59 = 1441812;
  v61 = 0LL;
  v62 = 119390;
  v63 = 1LL;
  v64 = 1;
  v65 = 1;
  v66 = 256;
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(DstBuf + 14928),
    (struct ADAPTER_CONTEXT *)DstBuf,
    ConfigurationHandle,
    (PUCHAR)"*EncapsulatedPacketTaskOffloadVxlan",
    0,
    1u,
    0,
    0,
    0);
  REGKEY<short>::Initialize(
    (enum _REGKEY_STATE *)(DstBuf + 14929),
    (struct ADAPTER_CONTEXT *)DstBuf,
    ConfigurationHandle,
    (PUCHAR)"*VxlanUDPPortNumber",
    1u,
    0xFFFFu,
    0x12B5u,
    0,
    0);
  REGISTRY::RegReadRegTable(
    DstBuf[14912],
    (struct ADAPTER_CONTEXT *)DstBuf,
    ConfigurationHandle,
    (struct REGTABLE_ENTRY *)&v4,
    8u);
}
