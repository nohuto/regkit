void __fastcall INTERRUPT::IntReadRegistryParameters(INTERRUPT *this, struct ADAPTER_CONTEXT *a2, void *a3)
{
  REGISTRY *v5; // rcx
  __int64 v6; // rcx
  int v7; // [rsp+30h] [rbp-D0h] BYREF
  const wchar_t *v8; // [rsp+38h] [rbp-C8h]
  __int64 v9; // [rsp+40h] [rbp-C0h]
  int v10; // [rsp+48h] [rbp-B8h]
  __int64 v11; // [rsp+4Ch] [rbp-B4h]
  int v12; // [rsp+54h] [rbp-ACh]
  int v13; // [rsp+58h] [rbp-A8h]
  __int16 v14; // [rsp+5Ch] [rbp-A4h]
  int v15; // [rsp+60h] [rbp-A0h]
  const wchar_t *v16; // [rsp+68h] [rbp-98h]
  __int64 v17; // [rsp+70h] [rbp-90h]
  int v18; // [rsp+78h] [rbp-88h]
  __int64 v19; // [rsp+7Ch] [rbp-84h]
  __int64 v20; // [rsp+84h] [rbp-7Ch]
  __int16 v21; // [rsp+8Ch] [rbp-74h]
  int v22; // [rsp+90h] [rbp-70h]
  const wchar_t *v23; // [rsp+98h] [rbp-68h]
  __int64 v24; // [rsp+A0h] [rbp-60h]
  int v25; // [rsp+A8h] [rbp-58h]
  __int64 v26; // [rsp+ACh] [rbp-54h]
  int v27; // [rsp+B4h] [rbp-4Ch]
  int v28; // [rsp+B8h] [rbp-48h]
  __int16 v29; // [rsp+BCh] [rbp-44h]
  int v30; // [rsp+C0h] [rbp-40h]
  const wchar_t *v31; // [rsp+C8h] [rbp-38h]
  __int64 v32; // [rsp+D0h] [rbp-30h]
  int v33; // [rsp+D8h] [rbp-28h]
  __int64 v34; // [rsp+DCh] [rbp-24h]
  int v35; // [rsp+E4h] [rbp-1Ch]
  int v36; // [rsp+E8h] [rbp-18h]
  __int16 v37; // [rsp+ECh] [rbp-14h]
  int v38; // [rsp+F0h] [rbp-10h]
  const wchar_t *v39; // [rsp+F8h] [rbp-8h]
  __int64 v40; // [rsp+100h] [rbp+0h]
  int v41; // [rsp+108h] [rbp+8h]
  __int64 v42; // [rsp+10Ch] [rbp+Ch]
  int v43; // [rsp+114h] [rbp+14h]
  int v44; // [rsp+118h] [rbp+18h]
  __int16 v45; // [rsp+11Ch] [rbp+1Ch]
  int v46; // [rsp+120h] [rbp+20h]
  const wchar_t *v47; // [rsp+128h] [rbp+28h]
  __int64 v48; // [rsp+130h] [rbp+30h]
  int v49; // [rsp+138h] [rbp+38h]
  __int64 v50; // [rsp+13Ch] [rbp+3Ch]
  int v51; // [rsp+144h] [rbp+44h]
  int v52; // [rsp+148h] [rbp+48h]
  __int16 v53; // [rsp+14Ch] [rbp+4Ch]
  int v54; // [rsp+150h] [rbp+50h]
  const wchar_t *v55; // [rsp+158h] [rbp+58h]
  __int64 v56; // [rsp+160h] [rbp+60h]
  int v57; // [rsp+168h] [rbp+68h]
  __int64 v58; // [rsp+16Ch] [rbp+6Ch]
  __int64 v59; // [rsp+174h] [rbp+74h]
  __int16 v60; // [rsp+17Ch] [rbp+7Ch]
  int v61; // [rsp+180h] [rbp+80h]
  const wchar_t *v62; // [rsp+188h] [rbp+88h]
  __int64 v63; // [rsp+190h] [rbp+90h]
  int v64; // [rsp+198h] [rbp+98h]
  __int64 v65; // [rsp+19Ch] [rbp+9Ch]
  int v66; // [rsp+1A4h] [rbp+A4h]
  int v67; // [rsp+1A8h] [rbp+A8h]
  __int16 v68; // [rsp+1ACh] [rbp+ACh]
  int v69; // [rsp+1B0h] [rbp+B0h]
  const wchar_t *v70; // [rsp+1B8h] [rbp+B8h]
  __int64 v71; // [rsp+1C0h] [rbp+C0h]
  int v72; // [rsp+1C8h] [rbp+C8h]
  __int64 v73; // [rsp+1CCh] [rbp+CCh]
  __int64 v74; // [rsp+1D4h] [rbp+D4h]
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
  int v85; // [rsp+210h] [rbp+110h]
  const wchar_t *v86; // [rsp+218h] [rbp+118h]
  __int64 v87; // [rsp+220h] [rbp+120h]
  int v88; // [rsp+228h] [rbp+128h]
  __int64 v89; // [rsp+22Ch] [rbp+12Ch]
  int v90; // [rsp+234h] [rbp+134h]
  int v91; // [rsp+238h] [rbp+138h]
  __int16 v92; // [rsp+23Ch] [rbp+13Ch]

  v7 = 2752552;
  v9 = 0LL;
  v8 = L"*InterruptModeration";
  v10 = 129009;
  v16 = L"EnableLLI";
  v11 = 1LL;
  v12 = 1;
  v23 = L"ITR";
  v13 = 1;
  v27 = 0xFFFF;
  v28 = 0xFFFF;
  v31 = L"InterruptMode";
  v35 = 3;
  v36 = 3;
  v39 = L"EnableAdvancedDynamicITR";
  v47 = L"AIMLowestLatency";
  v55 = L"EnableIAM";
  v62 = L"EnableEIAM";
  v70 = L"EnableTcpTimer";
  v14 = 256;
  v15 = 1310738;
  v17 = 0LL;
  v18 = 129076;
  v19 = 4LL;
  v20 = 2LL;
  v21 = 256;
  v22 = 524294;
  v24 = 0LL;
  v25 = 129012;
  v26 = 4LL;
  v29 = 257;
  v30 = 1835034;
  v32 = 0LL;
  v33 = 129028;
  v34 = 4LL;
  v37 = 256;
  v38 = 3276848;
  v40 = 0LL;
  v41 = 2680;
  v42 = 1LL;
  v43 = 1;
  v44 = 1;
  v45 = 256;
  v46 = 2228256;
  v48 = 0LL;
  v49 = 2681;
  v50 = 1LL;
  v51 = 1;
  v52 = 1;
  v53 = 256;
  v54 = 1310738;
  v56 = 0LL;
  v57 = 129036;
  v58 = 4LL;
  v59 = 1LL;
  v60 = 256;
  v61 = 1441812;
  v63 = 0LL;
  v64 = 67;
  v65 = 1LL;
  v66 = 1;
  v67 = 1;
  v68 = 256;
  v69 = 1966108;
  v5 = (REGISTRY *)*((_QWORD *)a2 + 14912);
  v77 = L"TcpTimerInterval";
  v82 = 50;
  v86 = L"Ndis61MsixConfig";
  v71 = 0LL;
  v72 = 129016;
  v73 = 1LL;
  v74 = 1LL;
  v75 = 256;
  v76 = 2228256;
  v78 = 0LL;
  v79 = 129020;
  v80 = 4;
  v81 = 1;
  v83 = 2;
  v84 = 257;
  v85 = 2228256;
  v87 = 0LL;
  v88 = 129024;
  v89 = 1LL;
  v90 = 1;
  v91 = 1;
  v92 = 256;
  REGISTRY::RegReadRegTable(v5, a2, a3, (struct REGTABLE_ENTRY *)&v7, 0xBu);
  v6 = *((_QWORD *)a2 + 14910);
  *((_BYTE *)a2 + 129025) = *((_BYTE *)a2 + 129024);
  (*(void (__fastcall **)(__int64, struct ADAPTER_CONTEXT *, void *))(*(_QWORD *)v6 + 64LL))(v6, a2, a3);
}
