void __fastcall RSS::RssReadRegistryParameters(RSS *this, struct ADAPTER_CONTEXT *a2, void *a3)
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
  int v17; // [rsp+84h] [rbp-7Ch]
  int v18; // [rsp+88h] [rbp-78h]
  __int16 v19; // [rsp+8Ch] [rbp-74h]
  int v20; // [rsp+90h] [rbp-70h]
  const wchar_t *v21; // [rsp+98h] [rbp-68h]
  __int64 v22; // [rsp+A0h] [rbp-60h]
  int v23; // [rsp+A8h] [rbp-58h]
  __int64 v24; // [rsp+ACh] [rbp-54h]
  int v25; // [rsp+B4h] [rbp-4Ch]
  int v26; // [rsp+B8h] [rbp-48h]
  __int16 v27; // [rsp+BCh] [rbp-44h]
  int v28; // [rsp+C0h] [rbp-40h]
  const wchar_t *v29; // [rsp+C8h] [rbp-38h]
  __int64 v30; // [rsp+D0h] [rbp-30h]
  int v31; // [rsp+D8h] [rbp-28h]
  __int64 v32; // [rsp+DCh] [rbp-24h]
  int v33; // [rsp+E4h] [rbp-1Ch]
  int v34; // [rsp+E8h] [rbp-18h]
  __int16 v35; // [rsp+ECh] [rbp-14h]
  int v36; // [rsp+F0h] [rbp-10h]
  const wchar_t *v37; // [rsp+F8h] [rbp-8h]
  __int64 v38; // [rsp+100h] [rbp+0h]
  int v39; // [rsp+108h] [rbp+8h]
  __int64 v40; // [rsp+10Ch] [rbp+Ch]
  int v41; // [rsp+114h] [rbp+14h]
  int v42; // [rsp+118h] [rbp+18h]
  __int16 v43; // [rsp+11Ch] [rbp+1Ch]
  int v44; // [rsp+120h] [rbp+20h]
  const wchar_t *v45; // [rsp+128h] [rbp+28h]
  __int64 v46; // [rsp+130h] [rbp+30h]
  int v47; // [rsp+138h] [rbp+38h]
  __int64 v48; // [rsp+13Ch] [rbp+3Ch]
  __int64 v49; // [rsp+144h] [rbp+44h]
  __int16 v50; // [rsp+14Ch] [rbp+4Ch]
  int v51; // [rsp+150h] [rbp+50h]
  const wchar_t *v52; // [rsp+158h] [rbp+58h]
  __int64 v53; // [rsp+160h] [rbp+60h]
  int v54; // [rsp+168h] [rbp+68h]
  __int64 v55; // [rsp+16Ch] [rbp+6Ch]
  int v56; // [rsp+174h] [rbp+74h]
  int v57; // [rsp+178h] [rbp+78h]
  __int16 v58; // [rsp+17Ch] [rbp+7Ch]
  int v59; // [rsp+180h] [rbp+80h]
  const wchar_t *v60; // [rsp+188h] [rbp+88h]
  __int64 v61; // [rsp+190h] [rbp+90h]
  int v62; // [rsp+198h] [rbp+98h]
  int v63; // [rsp+19Ch] [rbp+9Ch]
  int v64; // [rsp+1A0h] [rbp+A0h]
  int v65; // [rsp+1A4h] [rbp+A4h]
  int v66; // [rsp+1A8h] [rbp+A8h]
  __int16 v67; // [rsp+1ACh] [rbp+ACh]
  int v68; // [rsp+1B0h] [rbp+B0h]
  const wchar_t *v69; // [rsp+1B8h] [rbp+B8h]
  __int64 v70; // [rsp+1C0h] [rbp+C0h]
  int v71; // [rsp+1C8h] [rbp+C8h]
  __int64 v72; // [rsp+1CCh] [rbp+CCh]
  int v73; // [rsp+1D4h] [rbp+D4h]
  int v74; // [rsp+1D8h] [rbp+D8h]
  __int16 v75; // [rsp+1DCh] [rbp+DCh]
  int v76; // [rsp+1E0h] [rbp+E0h]
  const wchar_t *v77; // [rsp+1E8h] [rbp+E8h]
  __int64 v78; // [rsp+1F0h] [rbp+F0h]
  int v79; // [rsp+1F8h] [rbp+F8h]
  int v80; // [rsp+1FCh] [rbp+FCh]
  int v81; // [rsp+200h] [rbp+100h]
  int v82; // [rsp+204h] [rbp+104h]
  int v83; // [rsp+208h] [rbp+108h]
  __int16 v84; // [rsp+20Ch] [rbp+10Ch]

  v4 = 655368;
  v6 = 0LL;
  v5 = L"*RSS";
  v7 = 3096;
  v13 = L"*RssBaseProcNumber";
  v11 = 256;
  v17 = 0xFFFF;
  v18 = 0xFFFF;
  v25 = 0xFFFF;
  v21 = L"*MaxRssProcessors";
  v29 = L"*NumaNodeId";
  v37 = L"DisablePortScaling";
  v45 = L"ManyCoreScaling";
  v52 = L"*NumRssQueues";
  v26 = 0xFFFF;
  v33 = 0xFFFF;
  v34 = 0xFFFF;
  v60 = L"NumRssQueuesPerVPort";
  v8 = 1LL;
  v9 = 1;
  v10 = 1;
  v12 = 2490404;
  v14 = 0LL;
  v15 = 3104;
  v16 = 4LL;
  v19 = 256;
  v20 = 2359330;
  v22 = 0LL;
  v23 = 3108;
  v24 = 4LL;
  v27 = 256;
  v28 = 1572886;
  v30 = 0LL;
  v31 = 135200;
  v32 = 4LL;
  v35 = 256;
  v36 = 2490404;
  v38 = 0LL;
  v39 = 3097;
  v40 = 1LL;
  v41 = 1;
  v42 = 1;
  v43 = 256;
  v44 = 2097182;
  v46 = 0LL;
  v47 = 3100;
  v48 = 4LL;
  v49 = 1LL;
  v50 = 256;
  v51 = 1835034;
  v53 = 0LL;
  v54 = 3112;
  v55 = 4LL;
  v56 = 16;
  v57 = 1;
  v58 = 256;
  v59 = 2752552;
  v61 = 0LL;
  v62 = 3116;
  v63 = 4;
  v64 = 2;
  v65 = 2;
  v66 = 2;
  v67 = 256;
  v68 = 1835034;
  v69 = L"EnableLHRssWA";
  v82 = 2;
  v77 = L"ReceiveScalingMode";
  v3 = (REGISTRY *)*((_QWORD *)a2 + 14912);
  v83 = 2;
  v70 = 0LL;
  v71 = 3098;
  v72 = 1LL;
  v73 = 1;
  v74 = 1;
  v75 = 256;
  v76 = 2490404;
  v78 = 0LL;
  v79 = 3120;
  v80 = 4;
  v81 = 1;
  v84 = 256;
  REGISTRY::RegReadRegTable(v3, a2, a3, (struct REGTABLE_ENTRY *)&v4, 0xAu);
}
