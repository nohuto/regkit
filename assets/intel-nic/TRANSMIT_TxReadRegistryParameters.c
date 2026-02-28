void __fastcall TRANSMIT::TxReadRegistryParameters(TRANSMIT *this, struct ADAPTER_CONTEXT *a2, void *a3)
{
  REGISTRY *v3; // rcx
  int v4; // [rsp+30h] [rbp-D0h] BYREF
  const wchar_t *v5; // [rsp+38h] [rbp-C8h]
  __int64 v6; // [rsp+40h] [rbp-C0h]
  int v7; // [rsp+48h] [rbp-B8h]
  int v8; // [rsp+4Ch] [rbp-B4h]
  int v9; // [rsp+50h] [rbp-B0h]
  int v10; // [rsp+54h] [rbp-ACh]
  int v11; // [rsp+58h] [rbp-A8h]
  __int16 v12; // [rsp+5Ch] [rbp-A4h]
  int v13; // [rsp+60h] [rbp-A0h]
  const wchar_t *v14; // [rsp+68h] [rbp-98h]
  __int64 v15; // [rsp+70h] [rbp-90h]
  int v16; // [rsp+78h] [rbp-88h]
  __int64 v17; // [rsp+7Ch] [rbp-84h]
  int v18; // [rsp+84h] [rbp-7Ch]
  int v19; // [rsp+88h] [rbp-78h]
  __int16 v20; // [rsp+8Ch] [rbp-74h]
  int v21; // [rsp+90h] [rbp-70h]
  const wchar_t *v22; // [rsp+98h] [rbp-68h]
  __int64 v23; // [rsp+A0h] [rbp-60h]
  int v24; // [rsp+A8h] [rbp-58h]
  __int64 v25; // [rsp+ACh] [rbp-54h]
  __int64 v26; // [rsp+B4h] [rbp-4Ch]
  __int16 v27; // [rsp+BCh] [rbp-44h]
  int v28; // [rsp+C0h] [rbp-40h]
  const wchar_t *v29; // [rsp+C8h] [rbp-38h]
  __int64 v30; // [rsp+D0h] [rbp-30h]
  int v31; // [rsp+D8h] [rbp-28h]
  __int64 v32; // [rsp+DCh] [rbp-24h]
  __int64 v33; // [rsp+E4h] [rbp-1Ch]
  __int16 v34; // [rsp+ECh] [rbp-14h]
  int v35; // [rsp+F0h] [rbp-10h]
  const wchar_t *v36; // [rsp+F8h] [rbp-8h]
  __int64 v37; // [rsp+100h] [rbp+0h]
  int v38; // [rsp+108h] [rbp+8h]
  __int64 v39; // [rsp+10Ch] [rbp+Ch]
  __int64 v40; // [rsp+114h] [rbp+14h]
  __int16 v41; // [rsp+11Ch] [rbp+1Ch]
  int v42; // [rsp+120h] [rbp+20h]
  const wchar_t *v43; // [rsp+128h] [rbp+28h]
  __int64 v44; // [rsp+130h] [rbp+30h]
  int v45; // [rsp+138h] [rbp+38h]
  __int64 v46; // [rsp+13Ch] [rbp+3Ch]
  __int64 v47; // [rsp+144h] [rbp+44h]
  __int16 v48; // [rsp+14Ch] [rbp+4Ch]
  int v49; // [rsp+150h] [rbp+50h]
  const wchar_t *v50; // [rsp+158h] [rbp+58h]
  __int64 v51; // [rsp+160h] [rbp+60h]
  int v52; // [rsp+168h] [rbp+68h]
  __int64 v53; // [rsp+16Ch] [rbp+6Ch]
  int v54; // [rsp+174h] [rbp+74h]
  int v55; // [rsp+178h] [rbp+78h]
  __int16 v56; // [rsp+17Ch] [rbp+7Ch]
  int v57; // [rsp+180h] [rbp+80h]
  const wchar_t *v58; // [rsp+188h] [rbp+88h]
  __int64 v59; // [rsp+190h] [rbp+90h]
  int v60; // [rsp+198h] [rbp+98h]
  __int64 v61; // [rsp+19Ch] [rbp+9Ch]
  __int64 v62; // [rsp+1A4h] [rbp+A4h]
  __int16 v63; // [rsp+1ACh] [rbp+ACh]
  int v64; // [rsp+1B0h] [rbp+B0h]
  const wchar_t *v65; // [rsp+1B8h] [rbp+B8h]
  __int64 v66; // [rsp+1C0h] [rbp+C0h]
  int v67; // [rsp+1C8h] [rbp+C8h]
  int v68; // [rsp+1CCh] [rbp+CCh]
  int v69; // [rsp+1D0h] [rbp+D0h]
  int v70; // [rsp+1D4h] [rbp+D4h]
  int v71; // [rsp+1D8h] [rbp+D8h]
  __int16 v72; // [rsp+1DCh] [rbp+DCh]
  int v73; // [rsp+1E0h] [rbp+E0h]
  const wchar_t *v74; // [rsp+1E8h] [rbp+E8h]
  __int64 v75; // [rsp+1F0h] [rbp+F0h]
  int v76; // [rsp+1F8h] [rbp+F8h]
  __int64 v77; // [rsp+1FCh] [rbp+FCh]
  __int64 v78; // [rsp+204h] [rbp+104h]
  __int16 v79; // [rsp+20Ch] [rbp+10Ch]
  int v80; // [rsp+210h] [rbp+110h]
  const wchar_t *v81; // [rsp+218h] [rbp+118h]
  __int64 v82; // [rsp+220h] [rbp+120h]
  int v83; // [rsp+228h] [rbp+128h]
  __int64 v84; // [rsp+22Ch] [rbp+12Ch]
  int v85; // [rsp+234h] [rbp+134h]
  int v86; // [rsp+238h] [rbp+138h]
  __int16 v87; // [rsp+23Ch] [rbp+13Ch]
  int v88; // [rsp+240h] [rbp+140h]
  const wchar_t *v89; // [rsp+248h] [rbp+148h]
  __int64 v90; // [rsp+250h] [rbp+150h]
  int v91; // [rsp+258h] [rbp+158h]
  __int64 v92; // [rsp+25Ch] [rbp+15Ch]
  int v93; // [rsp+264h] [rbp+164h]
  int v94; // [rsp+268h] [rbp+168h]
  __int16 v95; // [rsp+26Ch] [rbp+16Ch]
  int v96; // [rsp+270h] [rbp+170h]
  const wchar_t *v97; // [rsp+278h] [rbp+178h]
  __int64 v98; // [rsp+280h] [rbp+180h]
  int v99; // [rsp+288h] [rbp+188h]
  int v100; // [rsp+28Ch] [rbp+18Ch]
  int v101; // [rsp+290h] [rbp+190h]
  int v102; // [rsp+294h] [rbp+194h]
  int v103; // [rsp+298h] [rbp+198h]
  __int16 v104; // [rsp+29Ch] [rbp+19Ch]
  int v105; // [rsp+2A0h] [rbp+1A0h]
  const wchar_t *v106; // [rsp+2A8h] [rbp+1A8h]
  __int64 v107; // [rsp+2B0h] [rbp+1B0h]
  int v108; // [rsp+2B8h] [rbp+1B8h]
  __int64 v109; // [rsp+2BCh] [rbp+1BCh]
  int v110; // [rsp+2C4h] [rbp+1C4h]
  int v111; // [rsp+2C8h] [rbp+1C8h]
  __int16 v112; // [rsp+2CCh] [rbp+1CCh]

  v4 = 2228256;
  v6 = 0LL;
  v5 = L"*TransmitBuffers";
  v7 = 35512;
  v14 = L"EnableTxHeadWB";
  v9 = 64;
  v8 = 4;
  v22 = L"TxWBThresh";
  v29 = L"EnableLocklessTx";
  v36 = L"DropHighlyFragmentedPacket";
  v43 = L"EnableCoalesce";
  v50 = L"CoalesceBufferSize";
  v54 = 2048;
  v55 = 2048;
  v58 = L"EnableUdpTxScaling";
  v65 = L"TxWritebackInterval";
  v10 = 65528;
  v11 = 512;
  v12 = 257;
  v13 = 1966108;
  v15 = 0LL;
  v16 = 35040;
  v17 = 1LL;
  v18 = 1;
  v19 = 1;
  v20 = 256;
  v21 = 1441812;
  v23 = 0LL;
  v24 = 35044;
  v25 = 4LL;
  v26 = 23LL;
  v27 = 257;
  v28 = 2228256;
  v30 = 0LL;
  v31 = 35048;
  v32 = 1LL;
  v33 = 1LL;
  v34 = 256;
  v35 = 3538996;
  v37 = 0LL;
  v38 = 35032;
  v39 = 1LL;
  v40 = 1LL;
  v41 = 256;
  v42 = 1966108;
  v44 = 0LL;
  v45 = 35033;
  v46 = 1LL;
  v47 = 1LL;
  v48 = 256;
  v49 = 2490404;
  v51 = 0LL;
  v52 = 35036;
  v53 = 4LL;
  v56 = 256;
  v57 = 2490404;
  v59 = 0LL;
  v60 = 35049;
  v61 = 1LL;
  v62 = 1LL;
  v63 = 256;
  v64 = 2621478;
  v66 = 0LL;
  v67 = 35544;
  v68 = 4;
  v69 = 1;
  v70 = 1;
  v71 = 1;
  v74 = L"QwaveAPI";
  v72 = 257;
  v81 = L"UserPriorityThresh";
  v89 = L"VerifyTDT_RDTWrite";
  v97 = L"MaxTxPacketsToFlush";
  v106 = L"EnableTss";
  v3 = (REGISTRY *)*((_QWORD *)a2 + 14912);
  v73 = 1179664;
  v75 = 0LL;
  v76 = 35140;
  v77 = 1LL;
  v78 = 1LL;
  v79 = 257;
  v80 = 2490404;
  v82 = 0LL;
  v83 = 35141;
  v84 = 1LL;
  v85 = 7;
  v86 = 2;
  v87 = 257;
  v88 = 2490404;
  v90 = 0LL;
  v91 = 35150;
  v92 = 1LL;
  v93 = 1;
  v94 = 1;
  v95 = 256;
  v96 = 2621478;
  v98 = 0LL;
  v99 = 35136;
  v100 = 4;
  v101 = 10;
  v102 = 2000;
  v103 = 512;
  v104 = 257;
  v105 = 1310738;
  v107 = 0LL;
  v108 = 129008;
  v109 = 1LL;
  v110 = 1;
  v111 = 1;
  v112 = 256;
  REGISTRY::RegReadRegTable(v3, a2, a3, (struct REGTABLE_ENTRY *)&v4, 0xEu);
}
