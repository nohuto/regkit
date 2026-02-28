void __fastcall RECEIVE::RxReadRegistryParameters(RECEIVE *this, struct ADAPTER_CONTEXT *a2, void *a3)
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
  int v26; // [rsp+B4h] [rbp-4Ch]
  int v27; // [rsp+B8h] [rbp-48h]
  __int16 v28; // [rsp+BCh] [rbp-44h]
  int v29; // [rsp+C0h] [rbp-40h]
  const wchar_t *v30; // [rsp+C8h] [rbp-38h]
  __int64 v31; // [rsp+D0h] [rbp-30h]
  int v32; // [rsp+D8h] [rbp-28h]
  __int64 v33; // [rsp+DCh] [rbp-24h]
  __int64 v34; // [rsp+E4h] [rbp-1Ch]
  __int16 v35; // [rsp+ECh] [rbp-14h]
  int v36; // [rsp+F0h] [rbp-10h]
  const wchar_t *v37; // [rsp+F8h] [rbp-8h]
  __int64 v38; // [rsp+100h] [rbp+0h]
  int v39; // [rsp+108h] [rbp+8h]
  int v40; // [rsp+10Ch] [rbp+Ch]
  int v41; // [rsp+110h] [rbp+10h]
  int v42; // [rsp+114h] [rbp+14h]
  int v43; // [rsp+118h] [rbp+18h]
  __int16 v44; // [rsp+11Ch] [rbp+1Ch]
  int v45; // [rsp+120h] [rbp+20h]
  const wchar_t *v46; // [rsp+128h] [rbp+28h]
  __int64 v47; // [rsp+130h] [rbp+30h]
  int v48; // [rsp+138h] [rbp+38h]
  int v49; // [rsp+13Ch] [rbp+3Ch]
  int v50; // [rsp+140h] [rbp+40h]
  int v51; // [rsp+144h] [rbp+44h]
  int v52; // [rsp+148h] [rbp+48h]
  __int16 v53; // [rsp+14Ch] [rbp+4Ch]
  int v54; // [rsp+150h] [rbp+50h]
  const wchar_t *v55; // [rsp+158h] [rbp+58h]
  __int64 v56; // [rsp+160h] [rbp+60h]
  int v57; // [rsp+168h] [rbp+68h]
  int v58; // [rsp+16Ch] [rbp+6Ch]
  int v59; // [rsp+170h] [rbp+70h]
  int v60; // [rsp+174h] [rbp+74h]
  int v61; // [rsp+178h] [rbp+78h]
  __int16 v62; // [rsp+17Ch] [rbp+7Ch]
  int v63; // [rsp+180h] [rbp+80h]
  const wchar_t *v64; // [rsp+188h] [rbp+88h]
  __int64 v65; // [rsp+190h] [rbp+90h]
  int v66; // [rsp+198h] [rbp+98h]
  __int64 v67; // [rsp+19Ch] [rbp+9Ch]
  int v68; // [rsp+1A4h] [rbp+A4h]
  int v69; // [rsp+1A8h] [rbp+A8h]
  __int16 v70; // [rsp+1ACh] [rbp+ACh]
  int v71; // [rsp+1B0h] [rbp+B0h]
  const wchar_t *v72; // [rsp+1B8h] [rbp+B8h]
  __int64 v73; // [rsp+1C0h] [rbp+C0h]
  int v74; // [rsp+1C8h] [rbp+C8h]
  int v75; // [rsp+1CCh] [rbp+CCh]
  int v76; // [rsp+1D0h] [rbp+D0h]
  int v77; // [rsp+1D4h] [rbp+D4h]
  int v78; // [rsp+1D8h] [rbp+D8h]
  __int16 v79; // [rsp+1DCh] [rbp+DCh]
  int v80; // [rsp+1E0h] [rbp+E0h]
  const wchar_t *v81; // [rsp+1E8h] [rbp+E8h]
  __int64 v82; // [rsp+1F0h] [rbp+F0h]
  int v83; // [rsp+1F8h] [rbp+F8h]
  __int64 v84; // [rsp+1FCh] [rbp+FCh]
  __int64 v85; // [rsp+204h] [rbp+104h]
  __int16 v86; // [rsp+20Ch] [rbp+10Ch]
  int v87; // [rsp+210h] [rbp+110h]
  const wchar_t *v88; // [rsp+218h] [rbp+118h]
  __int64 v89; // [rsp+220h] [rbp+120h]
  int v90; // [rsp+228h] [rbp+128h]
  __int64 v91; // [rsp+22Ch] [rbp+12Ch]
  __int64 v92; // [rsp+234h] [rbp+134h]
  __int16 v93; // [rsp+23Ch] [rbp+13Ch]
  int v94; // [rsp+240h] [rbp+140h]
  const wchar_t *v95; // [rsp+248h] [rbp+148h]
  __int64 v96; // [rsp+250h] [rbp+150h]
  int v97; // [rsp+258h] [rbp+158h]
  __int64 v98; // [rsp+25Ch] [rbp+15Ch]
  __int64 v99; // [rsp+264h] [rbp+164h]
  __int16 v100; // [rsp+26Ch] [rbp+16Ch]
  int v101; // [rsp+270h] [rbp+170h]
  const wchar_t *v102; // [rsp+278h] [rbp+178h]
  __int64 v103; // [rsp+280h] [rbp+180h]
  int v104; // [rsp+288h] [rbp+188h]
  __int64 v105; // [rsp+28Ch] [rbp+18Ch]
  int v106; // [rsp+294h] [rbp+194h]
  int v107; // [rsp+298h] [rbp+198h]
  __int16 v108; // [rsp+29Ch] [rbp+19Ch]
  int v109; // [rsp+2A0h] [rbp+1A0h]
  const wchar_t *v110; // [rsp+2A8h] [rbp+1A8h]
  __int64 v111; // [rsp+2B0h] [rbp+1B0h]
  int v112; // [rsp+2B8h] [rbp+1B8h]
  __int64 v113; // [rsp+2BCh] [rbp+1BCh]
  __int64 v114; // [rsp+2C4h] [rbp+1C4h]
  __int16 v115; // [rsp+2CCh] [rbp+1CCh]
  int v116; // [rsp+2D0h] [rbp+1D0h]
  const wchar_t *v117; // [rsp+2D8h] [rbp+1D8h]
  __int64 v118; // [rsp+2E0h] [rbp+1E0h]
  int v119; // [rsp+2E8h] [rbp+1E8h]
  __int64 v120; // [rsp+2ECh] [rbp+1ECh]
  int v121; // [rsp+2F4h] [rbp+1F4h]
  int v122; // [rsp+2F8h] [rbp+1F8h]
  __int16 v123; // [rsp+2FCh] [rbp+1FCh]
  int v124; // [rsp+300h] [rbp+200h]
  const wchar_t *v125; // [rsp+308h] [rbp+208h]
  __int64 v126; // [rsp+310h] [rbp+210h]
  int v127; // [rsp+318h] [rbp+218h]
  __int64 v128; // [rsp+31Ch] [rbp+21Ch]
  __int64 v129; // [rsp+324h] [rbp+224h]
  __int16 v130; // [rsp+32Ch] [rbp+22Ch]
  int v131; // [rsp+330h] [rbp+230h]
  const wchar_t *v132; // [rsp+338h] [rbp+238h]
  __int64 v133; // [rsp+340h] [rbp+240h]
  int v134; // [rsp+348h] [rbp+248h]
  __int64 v135; // [rsp+34Ch] [rbp+24Ch]
  int v136; // [rsp+354h] [rbp+254h]
  int v137; // [rsp+358h] [rbp+258h]
  __int16 v138; // [rsp+35Ch] [rbp+25Ch]
  int v139; // [rsp+360h] [rbp+260h]
  const wchar_t *v140; // [rsp+368h] [rbp+268h]
  __int64 v141; // [rsp+370h] [rbp+270h]
  int v142; // [rsp+378h] [rbp+278h]
  __int64 v143; // [rsp+37Ch] [rbp+27Ch]
  __int64 v144; // [rsp+384h] [rbp+284h]
  __int16 v145; // [rsp+38Ch] [rbp+28Ch]
  int v146; // [rsp+390h] [rbp+290h]
  const wchar_t *v147; // [rsp+398h] [rbp+298h]
  __int64 v148; // [rsp+3A0h] [rbp+2A0h]
  int v149; // [rsp+3A8h] [rbp+2A8h]
  int v150; // [rsp+3ACh] [rbp+2ACh]
  int v151; // [rsp+3B0h] [rbp+2B0h]
  int v152; // [rsp+3B4h] [rbp+2B4h]
  int v153; // [rsp+3B8h] [rbp+2B8h]
  __int16 v154; // [rsp+3BCh] [rbp+2BCh]
  int v155; // [rsp+3C0h] [rbp+2C0h]
  const wchar_t *v156; // [rsp+3C8h] [rbp+2C8h]
  __int64 v157; // [rsp+3D0h] [rbp+2D0h]
  int v158; // [rsp+3D8h] [rbp+2D8h]
  __int64 v159; // [rsp+3DCh] [rbp+2DCh]
  __int64 v160; // [rsp+3E4h] [rbp+2E4h]
  __int16 v161; // [rsp+3ECh] [rbp+2ECh]
  int v162; // [rsp+3F0h] [rbp+2F0h]
  const wchar_t *v163; // [rsp+3F8h] [rbp+2F8h]
  __int64 v164; // [rsp+400h] [rbp+300h]
  int v165; // [rsp+408h] [rbp+308h]
  __int64 v166; // [rsp+40Ch] [rbp+30Ch]
  int v167; // [rsp+414h] [rbp+314h]
  int v168; // [rsp+418h] [rbp+318h]
  __int16 v169; // [rsp+41Ch] [rbp+31Ch]

  v4 = 2097182;
  v65 = 0LL;
  v5 = L"*ReceiveBuffers";
  v6 = 0LL;
  v14 = L"RxWBThresh";
  v7 = 4520;
  v8 = 4;
  v22 = L"ReceiveBuffersOverride";
  v10 = 2048;
  v30 = L"RxPacketCount";
  v11 = 512;
  v37 = L"MaxPacketCountPerDPC";
  v12 = 257;
  v9 = 64;
  v46 = L"MaxPacketCountPerIndicate";
  v55 = L"RxDescriptorCountPerTailWrite";
  v64 = L"MinHardwareOwnedPacketCount";
  v13 = 1441812;
  v15 = 0LL;
  v16 = 2992;
  v17 = 4LL;
  v18 = 15;
  v19 = 1;
  v20 = 257;
  v21 = 3014700;
  v23 = 0LL;
  v24 = 3022;
  v25 = 1LL;
  v26 = 1;
  v27 = 1;
  v28 = 256;
  v29 = 1835034;
  v31 = 0LL;
  v32 = 4456;
  v33 = 0x4000000004LL;
  v34 = 65528LL;
  v35 = 257;
  v36 = 2752552;
  v38 = 0LL;
  v39 = 4464;
  v40 = 4;
  v41 = 8;
  v42 = 0xFFFF;
  v43 = 256;
  v44 = 257;
  v45 = 3407922;
  v47 = 0LL;
  v48 = 4468;
  v49 = 4;
  v50 = 1;
  v51 = 0xFFFF;
  v52 = 64;
  v53 = 257;
  v54 = 3932218;
  v56 = 0LL;
  v57 = 4648;
  v58 = 4;
  v59 = 4;
  v60 = 65528;
  v61 = 16;
  v62 = 257;
  v63 = 3670070;
  v66 = 4460;
  v67 = 0x800000004LL;
  v73 = 0LL;
  v76 = 0;
  v72 = L"RxBufferPad";
  v68 = 65528;
  v81 = L"MonitorMode";
  v106 = 0xFFFF;
  v82 = 0LL;
  v107 = 0xFFFF;
  v88 = L"MulticastFilterType";
  v89 = 0LL;
  v69 = 32;
  v95 = L"RegForceRxPathSerialization";
  v102 = L"ERT";
  v110 = L"RxPba";
  v117 = L"DynamicLTR";
  v125 = L"EnableRxDescriptorChaining";
  v70 = 257;
  v71 = 1572886;
  v74 = 4380;
  v75 = 4;
  v77 = 63;
  v78 = 10;
  v79 = 257;
  v80 = 1572886;
  v83 = 4048;
  v84 = 4LL;
  v85 = 2LL;
  v86 = 256;
  v87 = 2621478;
  v90 = 119760;
  v91 = 4LL;
  v92 = 3LL;
  v93 = 257;
  v94 = 3670070;
  v96 = 0LL;
  v97 = 3023;
  v98 = 1LL;
  v99 = 1LL;
  v100 = 256;
  v101 = 524294;
  v103 = 0LL;
  v104 = 155620;
  v105 = 4LL;
  v108 = 257;
  v109 = 786442;
  v111 = 0LL;
  v112 = 2984;
  v113 = 4LL;
  v114 = 255LL;
  v115 = 257;
  v116 = 1441812;
  v118 = 0LL;
  v119 = 3020;
  v120 = 1LL;
  v121 = 1;
  v122 = 1;
  v123 = 256;
  v124 = 3538996;
  v126 = 0LL;
  v127 = 3021;
  v128 = 1LL;
  v129 = 1LL;
  v130 = 256;
  v131 = 1441812;
  v133 = 0LL;
  v132 = L"EnableDRBT";
  v134 = 3028;
  v140 = L"*HeaderDataSplit";
  v135 = 4LL;
  v147 = L"HDSplitSize";
  v151 = 128;
  v153 = 128;
  v156 = L"HDSplitMode";
  v163 = L"HDSplitBufferPad";
  v3 = (REGISTRY *)*((_QWORD *)a2 + 14912);
  v136 = 1;
  v137 = 1;
  v138 = 256;
  v139 = 2228256;
  v141 = 0LL;
  v142 = 3072;
  v143 = 1LL;
  v144 = 1LL;
  v145 = 256;
  v146 = 1572886;
  v148 = 0LL;
  v149 = 3076;
  v150 = 4;
  v152 = 960;
  v154 = 256;
  v155 = 1572886;
  v157 = 0LL;
  v158 = 3080;
  v159 = 4LL;
  v160 = 1LL;
  v161 = 256;
  v162 = 2228256;
  v164 = 0LL;
  v165 = 3084;
  v166 = 4LL;
  v167 = 2;
  v168 = 2;
  v169 = 256;
  REGISTRY::RegReadRegTable(v3, a2, a3, (struct REGTABLE_ENTRY *)&v4, 0x15u);
}
