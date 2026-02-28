void __fastcall LINK::LinkReadRegistryParameters(LINK *this, struct ADAPTER_CONTEXT *a2, void *a3)
{
  REGISTRY *v3; // rcx
  int v4; // [rsp+30h] [rbp-D0h] BYREF
  const wchar_t *v5; // [rsp+38h] [rbp-C8h]
  __int64 v6; // [rsp+40h] [rbp-C0h]
  int v7; // [rsp+48h] [rbp-B8h]
  __int64 v8; // [rsp+4Ch] [rbp-B4h]
  __int64 v9; // [rsp+54h] [rbp-ACh]
  __int16 v10; // [rsp+5Ch] [rbp-A4h]
  int v11; // [rsp+60h] [rbp-A0h]
  const wchar_t *v12; // [rsp+68h] [rbp-98h]
  __int64 v13; // [rsp+70h] [rbp-90h]
  int v14; // [rsp+78h] [rbp-88h]
  __int64 v15; // [rsp+7Ch] [rbp-84h]
  int v16; // [rsp+84h] [rbp-7Ch]
  int v17; // [rsp+88h] [rbp-78h]
  __int16 v18; // [rsp+8Ch] [rbp-74h]
  int v19; // [rsp+90h] [rbp-70h]
  const wchar_t *v20; // [rsp+98h] [rbp-68h]
  __int64 v21; // [rsp+A0h] [rbp-60h]
  int v22; // [rsp+A8h] [rbp-58h]
  __int64 v23; // [rsp+ACh] [rbp-54h]
  __int64 v24; // [rsp+B4h] [rbp-4Ch]
  __int16 v25; // [rsp+BCh] [rbp-44h]
  int v26; // [rsp+C0h] [rbp-40h]
  const wchar_t *v27; // [rsp+C8h] [rbp-38h]
  __int64 v28; // [rsp+D0h] [rbp-30h]
  int v29; // [rsp+D8h] [rbp-28h]
  __int64 v30; // [rsp+DCh] [rbp-24h]
  int v31; // [rsp+E4h] [rbp-1Ch]
  int v32; // [rsp+E8h] [rbp-18h]
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
  int v45; // [rsp+13Ch] [rbp+3Ch]
  int v46; // [rsp+140h] [rbp+40h]
  int v47; // [rsp+144h] [rbp+44h]
  int v48; // [rsp+148h] [rbp+48h]
  __int16 v49; // [rsp+14Ch] [rbp+4Ch]
  int v50; // [rsp+150h] [rbp+50h]
  const wchar_t *v51; // [rsp+158h] [rbp+58h]
  __int64 v52; // [rsp+160h] [rbp+60h]
  int v53; // [rsp+168h] [rbp+68h]
  __int64 v54; // [rsp+16Ch] [rbp+6Ch]
  int v55; // [rsp+174h] [rbp+74h]
  int v56; // [rsp+178h] [rbp+78h]
  __int16 v57; // [rsp+17Ch] [rbp+7Ch]
  int v58; // [rsp+180h] [rbp+80h]
  const wchar_t *v59; // [rsp+188h] [rbp+88h]
  __int64 v60; // [rsp+190h] [rbp+90h]
  int v61; // [rsp+198h] [rbp+98h]
  __int64 v62; // [rsp+19Ch] [rbp+9Ch]
  __int64 v63; // [rsp+1A4h] [rbp+A4h]
  __int16 v64; // [rsp+1ACh] [rbp+ACh]
  int v65; // [rsp+1B0h] [rbp+B0h]
  const wchar_t *v66; // [rsp+1B8h] [rbp+B8h]
  __int64 v67; // [rsp+1C0h] [rbp+C0h]
  int v68; // [rsp+1C8h] [rbp+C8h]
  __int64 v69; // [rsp+1CCh] [rbp+CCh]
  __int64 v70; // [rsp+1D4h] [rbp+D4h]
  __int16 v71; // [rsp+1DCh] [rbp+DCh]
  int v72; // [rsp+1E0h] [rbp+E0h]
  const wchar_t *v73; // [rsp+1E8h] [rbp+E8h]
  __int64 v74; // [rsp+1F0h] [rbp+F0h]
  int v75; // [rsp+1F8h] [rbp+F8h]
  int v76; // [rsp+1FCh] [rbp+FCh]
  int v77; // [rsp+200h] [rbp+100h]
  int v78; // [rsp+204h] [rbp+104h]
  int v79; // [rsp+208h] [rbp+108h]
  __int16 v80; // [rsp+20Ch] [rbp+10Ch]
  int v81; // [rsp+210h] [rbp+110h]
  const wchar_t *v82; // [rsp+218h] [rbp+118h]
  __int64 v83; // [rsp+220h] [rbp+120h]
  int v84; // [rsp+228h] [rbp+128h]
  int v85; // [rsp+22Ch] [rbp+12Ch]
  int v86; // [rsp+230h] [rbp+130h]
  int v87; // [rsp+234h] [rbp+134h]
  int v88; // [rsp+238h] [rbp+138h]
  __int16 v89; // [rsp+23Ch] [rbp+13Ch]
  int v90; // [rsp+240h] [rbp+140h]
  const wchar_t *v91; // [rsp+248h] [rbp+148h]
  __int64 v92; // [rsp+250h] [rbp+150h]
  int v93; // [rsp+258h] [rbp+158h]
  int v94; // [rsp+25Ch] [rbp+15Ch]
  int v95; // [rsp+260h] [rbp+160h]
  int v96; // [rsp+264h] [rbp+164h]
  int v97; // [rsp+268h] [rbp+168h]
  __int16 v98; // [rsp+26Ch] [rbp+16Ch]
  int v99; // [rsp+270h] [rbp+170h]
  const wchar_t *v100; // [rsp+278h] [rbp+178h]
  __int64 v101; // [rsp+280h] [rbp+180h]
  int v102; // [rsp+288h] [rbp+188h]
  __int64 v103; // [rsp+28Ch] [rbp+18Ch]
  __int64 v104; // [rsp+294h] [rbp+194h]
  __int16 v105; // [rsp+29Ch] [rbp+19Ch]
  int v106; // [rsp+2A0h] [rbp+1A0h]
  const wchar_t *v107; // [rsp+2A8h] [rbp+1A8h]
  __int64 v108; // [rsp+2B0h] [rbp+1B0h]
  int v109; // [rsp+2B8h] [rbp+1B8h]
  __int64 v110; // [rsp+2BCh] [rbp+1BCh]
  __int64 v111; // [rsp+2C4h] [rbp+1C4h]
  __int16 v112; // [rsp+2CCh] [rbp+1CCh]
  int v113; // [rsp+2D0h] [rbp+1D0h]
  const wchar_t *v114; // [rsp+2D8h] [rbp+1D8h]
  __int64 v115; // [rsp+2E0h] [rbp+1E0h]
  int v116; // [rsp+2E8h] [rbp+1E8h]
  int v117; // [rsp+2ECh] [rbp+1ECh]
  int v118; // [rsp+2F0h] [rbp+1F0h]
  int v119; // [rsp+2F4h] [rbp+1F4h]
  int v120; // [rsp+2F8h] [rbp+1F8h]
  __int16 v121; // [rsp+2FCh] [rbp+1FCh]
  int v122; // [rsp+300h] [rbp+200h]
  const wchar_t *v123; // [rsp+308h] [rbp+208h]
  __int64 v124; // [rsp+310h] [rbp+210h]
  int v125; // [rsp+318h] [rbp+218h]
  int v126; // [rsp+31Ch] [rbp+21Ch]
  int v127; // [rsp+320h] [rbp+220h]
  int v128; // [rsp+324h] [rbp+224h]
  int v129; // [rsp+328h] [rbp+228h]
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
  __int64 v150; // [rsp+3ACh] [rbp+2ACh]
  int v151; // [rsp+3B4h] [rbp+2B4h]
  int v152; // [rsp+3B8h] [rbp+2B8h]
  __int16 v153; // [rsp+3BCh] [rbp+2BCh]
  int v154; // [rsp+3C0h] [rbp+2C0h]
  const wchar_t *v155; // [rsp+3C8h] [rbp+2C8h]
  __int64 v156; // [rsp+3D0h] [rbp+2D0h]
  int v157; // [rsp+3D8h] [rbp+2D8h]
  __int64 v158; // [rsp+3DCh] [rbp+2DCh]
  int v159; // [rsp+3E4h] [rbp+2E4h]
  int v160; // [rsp+3E8h] [rbp+2E8h]
  __int16 v161; // [rsp+3ECh] [rbp+2ECh]
  int v162; // [rsp+3F0h] [rbp+2F0h]
  const wchar_t *v163; // [rsp+3F8h] [rbp+2F8h]
  __int64 v164; // [rsp+400h] [rbp+300h]
  int v165; // [rsp+408h] [rbp+308h]
  __int64 v166; // [rsp+40Ch] [rbp+30Ch]
  int v167; // [rsp+414h] [rbp+314h]
  int v168; // [rsp+418h] [rbp+318h]
  __int16 v169; // [rsp+41Ch] [rbp+31Ch]
  int v170; // [rsp+420h] [rbp+320h]
  const wchar_t *v171; // [rsp+428h] [rbp+328h]
  __int64 v172; // [rsp+430h] [rbp+330h]
  int v173; // [rsp+438h] [rbp+338h]
  __int64 v174; // [rsp+43Ch] [rbp+33Ch]
  __int64 v175; // [rsp+444h] [rbp+344h]
  __int16 v176; // [rsp+44Ch] [rbp+34Ch]

  v4 = 2097182;
  v43 = 0LL;
  v5 = L"DisablePhyReset";
  v6 = 0LL;
  v12 = L"EnableExtraLinkUpRetries";
  v7 = 120614;
  v16 = 10;
  v17 = 10;
  v20 = L"*SpeedDuplex";
  v8 = 1LL;
  v27 = L"WaitAutoNegComplete";
  v9 = 1LL;
  v35 = L"SmartSpeed";
  v42 = L"AutoNegAdvertised";
  v47 = 47;
  v48 = 47;
  v51 = L"AdaptiveIFS";
  v52 = 0LL;
  v59 = L"Mdix";
  v60 = 0LL;
  v10 = 256;
  v11 = 3276848;
  v13 = 0LL;
  v14 = 129440;
  v15 = 4LL;
  v18 = 257;
  v19 = 1703960;
  v21 = 0LL;
  v22 = 129456;
  v23 = 4LL;
  v24 = 6LL;
  v25 = 256;
  v26 = 2621478;
  v28 = 0LL;
  v29 = 129464;
  v30 = 1LL;
  v31 = 2;
  v32 = 2;
  v33 = 256;
  v34 = 1441812;
  v36 = 0LL;
  v37 = 120576;
  v38 = 4LL;
  v39 = 2LL;
  v40 = 257;
  v41 = 2359330;
  v44 = 120600;
  v45 = 2;
  v46 = 1;
  v49 = 256;
  v50 = 1572886;
  v53 = 120303;
  v54 = 1LL;
  v55 = 1;
  v56 = 1;
  v57 = 256;
  v58 = 655368;
  v61 = 120610;
  v62 = 1LL;
  v63 = 3LL;
  v64 = 257;
  v65 = 1572886;
  v74 = 0LL;
  v77 = 0;
  v83 = 0LL;
  v86 = 0;
  v92 = 0LL;
  v95 = 0;
  v97 = 0;
  v115 = 0LL;
  v118 = 0;
  v124 = 0LL;
  v127 = 0;
  v66 = L"MasterSlave";
  v67 = 0LL;
  v73 = L"ReduceSpeedOnPowerDown";
  v76 = 1;
  v78 = 1;
  v79 = 1;
  v82 = L"LinkNegotiationProcess";
  v91 = L"AllowAllSpeedsLPLU";
  v94 = 1;
  v96 = 1;
  v100 = L"ProcessLSCinWorkItem";
  v101 = 0LL;
  v107 = L"DetectForcedLP";
  v108 = 0LL;
  v114 = L"EEELinkAdvertisement";
  v119 = 1;
  v120 = 1;
  v70 = 3LL;
  v128 = 3;
  v129 = 3;
  v123 = L"*FlowControl";
  v68 = 120564;
  v69 = 4LL;
  v71 = 256;
  v72 = 3014700;
  v75 = 129480;
  v80 = 256;
  v81 = 3014700;
  v84 = 129468;
  v85 = 4;
  v87 = 2;
  v88 = 2;
  v89 = 257;
  v90 = 2490404;
  v93 = 129481;
  v98 = 256;
  v99 = 2752552;
  v102 = 129472;
  v103 = 4LL;
  v104 = 1LL;
  v105 = 256;
  v106 = 1966108;
  v109 = 129476;
  v110 = 4LL;
  v111 = 1LL;
  v112 = 256;
  v113 = 2752552;
  v116 = 129500;
  v117 = 4;
  v121 = 257;
  v122 = 1703960;
  v125 = 129460;
  v126 = 4;
  v130 = 256;
  v133 = 0LL;
  v141 = 0LL;
  v132 = L"FlowControlSendXon";
  v131 = 2490404;
  v151 = 0xFFFF;
  v140 = L"FlowControlStrictIEEE";
  v147 = L"FlowControlHighWatermark";
  v155 = L"FlowControlLowWatermark";
  v163 = L"FlowControlPauseTime";
  v171 = L"PhyTimingRecoveryWA";
  v3 = (REGISTRY *)*((_QWORD *)a2 + 14912);
  v159 = 0xFFFF;
  v134 = 120348;
  v135 = 1LL;
  v136 = 1;
  v137 = 1;
  v138 = 257;
  v139 = 2883626;
  v142 = 120349;
  v143 = 1LL;
  v144 = 1LL;
  v145 = 256;
  v146 = 3276848;
  v148 = 0LL;
  v149 = 120336;
  v150 = 4LL;
  v152 = 58982;
  v153 = 257;
  v154 = 3145774;
  v156 = 0LL;
  v157 = 120340;
  v158 = 4LL;
  v160 = 52428;
  v161 = 257;
  v162 = 2752552;
  v164 = 0LL;
  v165 = 120344;
  v166 = 2LL;
  v167 = 0x2000;
  v168 = 1664;
  v169 = 257;
  v170 = 2621478;
  v172 = 0LL;
  v173 = 129482;
  v174 = 1LL;
  v175 = 1LL;
  v176 = 256;
  REGISTRY::RegReadRegTable(v3, a2, a3, (struct REGTABLE_ENTRY *)&v4, 0x16u);
}
