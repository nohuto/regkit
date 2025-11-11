__int64 DpiMiracastGetForcedMode()
{
  __int64 result; // rax
  int v1; // [rsp+30h] [rbp-79h] BYREF
  int v2; // [rsp+34h] [rbp-75h] BYREF
  int v3; // [rsp+38h] [rbp-71h] BYREF
  __int64 v4; // [rsp+40h] [rbp-69h] BYREF
  int v5; // [rsp+48h] [rbp-61h]
  const wchar_t *v6; // [rsp+50h] [rbp-59h]
  int *v7; // [rsp+58h] [rbp-51h]
  int v8; // [rsp+60h] [rbp-49h]
  int *v9; // [rsp+68h] [rbp-41h]
  int v10; // [rsp+70h] [rbp-39h]
  __int64 v11; // [rsp+78h] [rbp-31h]
  int v12; // [rsp+80h] [rbp-29h]
  const wchar_t *v13; // [rsp+88h] [rbp-21h]
  int *v14; // [rsp+90h] [rbp-19h]
  int v15; // [rsp+98h] [rbp-11h]
  int *v16; // [rsp+A0h] [rbp-9h]
  int v17; // [rsp+A8h] [rbp-1h]
  __int64 v18; // [rsp+B0h] [rbp+7h]
  int v19; // [rsp+B8h] [rbp+Fh]
  __int64 v20; // [rsp+C0h] [rbp+17h]
  __int128 v21; // [rsp+C8h] [rbp+1Fh]
  __int128 v22; // [rsp+D8h] [rbp+2Fh]

  v4 = 0LL;
  v11 = 0LL;
  v18 = 0LL;
  v19 = 0;
  v20 = 0LL;
  v6 = L"MiracastUseIhvDriver";
  v7 = &v3;
  v9 = &v1;
  v13 = L"MiracastForceDisable";
  v14 = &v2;
  v5 = 288;
  v8 = 67108868;
  v10 = 4;
  v12 = 288;
  v15 = 67108868;
  v17 = 4;
  v16 = &v1;
  v1 = 2;
  v3 = 2;
  v2 = 2;
  v21 = 0LL;
  v22 = 0LL;
  if ( (int)RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v4) < 0 )
    return 0LL;
  result = 1LL;
  if ( v2 == 1 )
    return 3LL;
  if ( v3 != 1 )
    return v3 == 0 ? 2 : 0;
  return result;
}

void __fastcall DXGADAPTER::ReadConfig(DXGADAPTER *this, struct _DXGK_ADAPTER_CAPS *a2)
{
  bool v4; // zf
  bool v5; // al
  bool v6; // al
  bool v7; // al
  bool v8; // al
  bool v9; // al
  bool v10; // al
  bool v11; // al
  char v12; // al
  bool v13; // al
  char v14; // al
  bool v15; // cf
  int v16; // eax
  bool v17; // al
  bool v18; // al
  bool v19; // al
  bool v20; // al
  char v21; // al
  char v22; // dl
  int v23; // r8d
  char v24; // al
  int v25; // eax
  bool v26; // al
  int v27; // eax
  char v28; // dl
  _DWORD *v29; // rcx
  _DWORD *v30; // r8
  char v31; // al
  int v32; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v33; // [rsp+34h] [rbp-CCh] BYREF
  int v34; // [rsp+38h] [rbp-C8h] BYREF
  int v35; // [rsp+3Ch] [rbp-C4h] BYREF
  int v36; // [rsp+40h] [rbp-C0h] BYREF
  int v37; // [rsp+44h] [rbp-BCh] BYREF
  int v38; // [rsp+48h] [rbp-B8h] BYREF
  int v39; // [rsp+4Ch] [rbp-B4h] BYREF
  int v40; // [rsp+50h] [rbp-B0h] BYREF
  int v41; // [rsp+54h] [rbp-ACh] BYREF
  int v42; // [rsp+58h] [rbp-A8h] BYREF
  int v43; // [rsp+5Ch] [rbp-A4h] BYREF
  int v44; // [rsp+60h] [rbp-A0h] BYREF
  int v45; // [rsp+64h] [rbp-9Ch] BYREF
  int v46; // [rsp+68h] [rbp-98h] BYREF
  int v47; // [rsp+6Ch] [rbp-94h] BYREF
  int v48; // [rsp+70h] [rbp-90h] BYREF
  int v49; // [rsp+74h] [rbp-8Ch] BYREF
  int v50; // [rsp+78h] [rbp-88h] BYREF
  int v51; // [rsp+7Ch] [rbp-84h] BYREF
  int v52; // [rsp+80h] [rbp-80h] BYREF
  int v53; // [rsp+84h] [rbp-7Ch] BYREF
  int v54; // [rsp+88h] [rbp-78h] BYREF
  int v55; // [rsp+8Ch] [rbp-74h] BYREF
  int v56; // [rsp+90h] [rbp-70h] BYREF
  int v57; // [rsp+94h] [rbp-6Ch] BYREF
  int v58; // [rsp+98h] [rbp-68h] BYREF
  int v59; // [rsp+9Ch] [rbp-64h] BYREF
  int v60; // [rsp+A0h] [rbp-60h] BYREF
  int v61; // [rsp+A4h] [rbp-5Ch] BYREF
  int v62; // [rsp+A8h] [rbp-58h] BYREF
  int v63; // [rsp+ACh] [rbp-54h] BYREF
  int v64; // [rsp+B0h] [rbp-50h] BYREF
  int v65; // [rsp+B4h] [rbp-4Ch] BYREF
  int v66; // [rsp+B8h] [rbp-48h] BYREF
  int v67; // [rsp+BCh] [rbp-44h] BYREF
  int v68; // [rsp+C0h] [rbp-40h] BYREF
  int v69; // [rsp+C4h] [rbp-3Ch] BYREF
  int v70; // [rsp+C8h] [rbp-38h] BYREF
  int v71; // [rsp+CCh] [rbp-34h] BYREF
  int v72; // [rsp+D0h] [rbp-30h] BYREF
  int v73; // [rsp+D4h] [rbp-2Ch] BYREF
  int v74; // [rsp+D8h] [rbp-28h] BYREF
  int v75; // [rsp+DCh] [rbp-24h] BYREF
  int v76; // [rsp+E0h] [rbp-20h] BYREF
  int v77; // [rsp+E4h] [rbp-1Ch] BYREF
  int v78; // [rsp+E8h] [rbp-18h] BYREF
  int v79; // [rsp+ECh] [rbp-14h] BYREF
  int v80; // [rsp+F0h] [rbp-10h] BYREF
  int v81; // [rsp+F4h] [rbp-Ch] BYREF
  int v82; // [rsp+F8h] [rbp-8h] BYREF
  int v83; // [rsp+FCh] [rbp-4h] BYREF
  int v84; // [rsp+100h] [rbp+0h] BYREF
  int v85; // [rsp+104h] [rbp+4h] BYREF
  int v86; // [rsp+108h] [rbp+8h] BYREF
  int v87; // [rsp+10Ch] [rbp+Ch] BYREF
  int v88; // [rsp+110h] [rbp+10h] BYREF
  int v89; // [rsp+114h] [rbp+14h] BYREF
  int v90; // [rsp+118h] [rbp+18h] BYREF
  int v91; // [rsp+11Ch] [rbp+1Ch] BYREF
  int v92; // [rsp+120h] [rbp+20h] BYREF
  int v93; // [rsp+124h] [rbp+24h] BYREF
  int v94; // [rsp+128h] [rbp+28h] BYREF
  int v95; // [rsp+12Ch] [rbp+2Ch] BYREF
  int v96; // [rsp+130h] [rbp+30h] BYREF
  int v97; // [rsp+134h] [rbp+34h] BYREF
  int v98; // [rsp+138h] [rbp+38h] BYREF
  int v99; // [rsp+13Ch] [rbp+3Ch] BYREF
  int v100; // [rsp+140h] [rbp+40h] BYREF
  int v101; // [rsp+144h] [rbp+44h] BYREF
  int v102; // [rsp+148h] [rbp+48h] BYREF
  int v103; // [rsp+14Ch] [rbp+4Ch] BYREF
  int v104; // [rsp+150h] [rbp+50h] BYREF
  int v105; // [rsp+154h] [rbp+54h] BYREF
  int v106; // [rsp+158h] [rbp+58h] BYREF
  int v107; // [rsp+15Ch] [rbp+5Ch] BYREF
  int v108; // [rsp+160h] [rbp+60h] BYREF
  int v109; // [rsp+164h] [rbp+64h] BYREF
  int v110; // [rsp+168h] [rbp+68h] BYREF
  int v111; // [rsp+16Ch] [rbp+6Ch] BYREF
  int v112; // [rsp+170h] [rbp+70h] BYREF
  int v113; // [rsp+174h] [rbp+74h] BYREF
  int v114; // [rsp+178h] [rbp+78h] BYREF
  int v115; // [rsp+17Ch] [rbp+7Ch] BYREF
  int v116; // [rsp+180h] [rbp+80h] BYREF
  int v117; // [rsp+184h] [rbp+84h] BYREF
  int v118; // [rsp+188h] [rbp+88h] BYREF
  int v119; // [rsp+18Ch] [rbp+8Ch] BYREF
  int v120; // [rsp+190h] [rbp+90h] BYREF
  int v121; // [rsp+194h] [rbp+94h] BYREF
  __int64 v122; // [rsp+198h] [rbp+98h] BYREF
  __int64 v123; // [rsp+1A0h] [rbp+A0h] BYREF
  __int64 v124; // [rsp+1A8h] [rbp+A8h]
  __int64 v125; // [rsp+1B0h] [rbp+B0h] BYREF
  int v126; // [rsp+1B8h] [rbp+B8h]
  const wchar_t *v127; // [rsp+1C0h] [rbp+C0h]
  int *v128; // [rsp+1C8h] [rbp+C8h]
  int v129; // [rsp+1D0h] [rbp+D0h]
  int *v130; // [rsp+1D8h] [rbp+D8h]
  int v131; // [rsp+1E0h] [rbp+E0h]
  __int64 v132; // [rsp+1E8h] [rbp+E8h]
  int v133; // [rsp+1F0h] [rbp+F0h]
  const wchar_t *v134; // [rsp+1F8h] [rbp+F8h]
  int *v135; // [rsp+200h] [rbp+100h]
  int v136; // [rsp+208h] [rbp+108h]
  int *v137; // [rsp+210h] [rbp+110h]
  int v138; // [rsp+218h] [rbp+118h]
  __int64 v139; // [rsp+220h] [rbp+120h]
  int v140; // [rsp+228h] [rbp+128h]
  const wchar_t *v141; // [rsp+230h] [rbp+130h]
  int *v142; // [rsp+238h] [rbp+138h]
  int v143; // [rsp+240h] [rbp+140h]
  int *v144; // [rsp+248h] [rbp+148h]
  int v145; // [rsp+250h] [rbp+150h]
  __int64 v146; // [rsp+258h] [rbp+158h]
  int v147; // [rsp+260h] [rbp+160h]
  const wchar_t *v148; // [rsp+268h] [rbp+168h]
  __int64 *v149; // [rsp+270h] [rbp+170h]
  int v150; // [rsp+278h] [rbp+178h]
  __int64 *v151; // [rsp+280h] [rbp+180h]
  int v152; // [rsp+288h] [rbp+188h]
  __int64 v153; // [rsp+290h] [rbp+190h]
  int v154; // [rsp+298h] [rbp+198h]
  const wchar_t *v155; // [rsp+2A0h] [rbp+1A0h]
  int *v156; // [rsp+2A8h] [rbp+1A8h]
  int v157; // [rsp+2B0h] [rbp+1B0h]
  int *v158; // [rsp+2B8h] [rbp+1B8h]
  int v159; // [rsp+2C0h] [rbp+1C0h]
  __int64 v160; // [rsp+2C8h] [rbp+1C8h]
  int v161; // [rsp+2D0h] [rbp+1D0h]
  const wchar_t *v162; // [rsp+2D8h] [rbp+1D8h]
  int *v163; // [rsp+2E0h] [rbp+1E0h]
  int v164; // [rsp+2E8h] [rbp+1E8h]
  int *v165; // [rsp+2F0h] [rbp+1F0h]
  int v166; // [rsp+2F8h] [rbp+1F8h]
  __int64 v167; // [rsp+300h] [rbp+200h]
  int v168; // [rsp+308h] [rbp+208h]
  const wchar_t *v169; // [rsp+310h] [rbp+210h]
  int *v170; // [rsp+318h] [rbp+218h]
  int v171; // [rsp+320h] [rbp+220h]
  int *v172; // [rsp+328h] [rbp+228h]
  int v173; // [rsp+330h] [rbp+230h]
  __int64 v174; // [rsp+338h] [rbp+238h]
  int v175; // [rsp+340h] [rbp+240h]
  const wchar_t *v176; // [rsp+348h] [rbp+248h]
  int *v177; // [rsp+350h] [rbp+250h]
  int v178; // [rsp+358h] [rbp+258h]
  int *v179; // [rsp+360h] [rbp+260h]
  int v180; // [rsp+368h] [rbp+268h]
  __int64 v181; // [rsp+370h] [rbp+270h]
  int v182; // [rsp+378h] [rbp+278h]
  const wchar_t *v183; // [rsp+380h] [rbp+280h]
  int *v184; // [rsp+388h] [rbp+288h]
  int v185; // [rsp+390h] [rbp+290h]
  int *v186; // [rsp+398h] [rbp+298h]
  int v187; // [rsp+3A0h] [rbp+2A0h]
  __int64 v188; // [rsp+3A8h] [rbp+2A8h]
  int v189; // [rsp+3B0h] [rbp+2B0h]
  const wchar_t *v190; // [rsp+3B8h] [rbp+2B8h]
  int *v191; // [rsp+3C0h] [rbp+2C0h]
  int v192; // [rsp+3C8h] [rbp+2C8h]
  int *v193; // [rsp+3D0h] [rbp+2D0h]
  int v194; // [rsp+3D8h] [rbp+2D8h]
  __int64 v195; // [rsp+3E0h] [rbp+2E0h]
  int v196; // [rsp+3E8h] [rbp+2E8h]
  const wchar_t *v197; // [rsp+3F0h] [rbp+2F0h]
  int *v198; // [rsp+3F8h] [rbp+2F8h]
  int v199; // [rsp+400h] [rbp+300h]
  int *v200; // [rsp+408h] [rbp+308h]
  int v201; // [rsp+410h] [rbp+310h]
  __int64 v202; // [rsp+418h] [rbp+318h]
  int v203; // [rsp+420h] [rbp+320h]
  const wchar_t *v204; // [rsp+428h] [rbp+328h]
  int *v205; // [rsp+430h] [rbp+330h]
  int v206; // [rsp+438h] [rbp+338h]
  int *v207; // [rsp+440h] [rbp+340h]
  int v208; // [rsp+448h] [rbp+348h]
  __int64 v209; // [rsp+450h] [rbp+350h]
  int v210; // [rsp+458h] [rbp+358h]
  const wchar_t *v211; // [rsp+460h] [rbp+360h]
  int *v212; // [rsp+468h] [rbp+368h]
  int v213; // [rsp+470h] [rbp+370h]
  int *v214; // [rsp+478h] [rbp+378h]
  int v215; // [rsp+480h] [rbp+380h]
  __int64 v216; // [rsp+488h] [rbp+388h]
  int v217; // [rsp+490h] [rbp+390h]
  const wchar_t *v218; // [rsp+498h] [rbp+398h]
  int *v219; // [rsp+4A0h] [rbp+3A0h]
  int v220; // [rsp+4A8h] [rbp+3A8h]
  int *v221; // [rsp+4B0h] [rbp+3B0h]
  int v222; // [rsp+4B8h] [rbp+3B8h]
  __int64 v223; // [rsp+4C0h] [rbp+3C0h]
  int v224; // [rsp+4C8h] [rbp+3C8h]
  const wchar_t *v225; // [rsp+4D0h] [rbp+3D0h]
  int *v226; // [rsp+4D8h] [rbp+3D8h]
  int v227; // [rsp+4E0h] [rbp+3E0h]
  int *v228; // [rsp+4E8h] [rbp+3E8h]
  int v229; // [rsp+4F0h] [rbp+3F0h]
  __int64 v230; // [rsp+4F8h] [rbp+3F8h]
  int v231; // [rsp+500h] [rbp+400h]
  const wchar_t *v232; // [rsp+508h] [rbp+408h]
  int *v233; // [rsp+510h] [rbp+410h]
  int v234; // [rsp+518h] [rbp+418h]
  int *v235; // [rsp+520h] [rbp+420h]
  int v236; // [rsp+528h] [rbp+428h]
  __int64 v237; // [rsp+530h] [rbp+430h]
  int v238; // [rsp+538h] [rbp+438h]
  const wchar_t *v239; // [rsp+540h] [rbp+440h]
  int *v240; // [rsp+548h] [rbp+448h]
  int v241; // [rsp+550h] [rbp+450h]
  int *v242; // [rsp+558h] [rbp+458h]
  int v243; // [rsp+560h] [rbp+460h]
  __int64 v244; // [rsp+568h] [rbp+468h]
  int v245; // [rsp+570h] [rbp+470h]
  const wchar_t *v246; // [rsp+578h] [rbp+478h]
  int *v247; // [rsp+580h] [rbp+480h]
  int v248; // [rsp+588h] [rbp+488h]
  int *v249; // [rsp+590h] [rbp+490h]
  int v250; // [rsp+598h] [rbp+498h]
  __int64 v251; // [rsp+5A0h] [rbp+4A0h]
  int v252; // [rsp+5A8h] [rbp+4A8h]
  const wchar_t *v253; // [rsp+5B0h] [rbp+4B0h]
  int *v254; // [rsp+5B8h] [rbp+4B8h]
  int v255; // [rsp+5C0h] [rbp+4C0h]
  int *v256; // [rsp+5C8h] [rbp+4C8h]
  int v257; // [rsp+5D0h] [rbp+4D0h]
  __int64 v258; // [rsp+5D8h] [rbp+4D8h]
  int v259; // [rsp+5E0h] [rbp+4E0h]
  const wchar_t *v260; // [rsp+5E8h] [rbp+4E8h]
  int *v261; // [rsp+5F0h] [rbp+4F0h]
  int v262; // [rsp+5F8h] [rbp+4F8h]
  int *v263; // [rsp+600h] [rbp+500h]
  int v264; // [rsp+608h] [rbp+508h]
  __int64 v265; // [rsp+610h] [rbp+510h]
  int v266; // [rsp+618h] [rbp+518h]
  const wchar_t *v267; // [rsp+620h] [rbp+520h]
  int *v268; // [rsp+628h] [rbp+528h]
  int v269; // [rsp+630h] [rbp+530h]
  int *v270; // [rsp+638h] [rbp+538h]
  int v271; // [rsp+640h] [rbp+540h]
  __int64 v272; // [rsp+648h] [rbp+548h]
  int v273; // [rsp+650h] [rbp+550h]
  const wchar_t *v274; // [rsp+658h] [rbp+558h]
  int *v275; // [rsp+660h] [rbp+560h]
  int v276; // [rsp+668h] [rbp+568h]
  int *v277; // [rsp+670h] [rbp+570h]
  int v278; // [rsp+678h] [rbp+578h]
  __int64 v279; // [rsp+680h] [rbp+580h]
  int v280; // [rsp+688h] [rbp+588h]
  const wchar_t *v281; // [rsp+690h] [rbp+590h]
  int *v282; // [rsp+698h] [rbp+598h]
  int v283; // [rsp+6A0h] [rbp+5A0h]
  int *v284; // [rsp+6A8h] [rbp+5A8h]
  int v285; // [rsp+6B0h] [rbp+5B0h]
  __int64 v286; // [rsp+6B8h] [rbp+5B8h]
  int v287; // [rsp+6C0h] [rbp+5C0h]
  const wchar_t *v288; // [rsp+6C8h] [rbp+5C8h]
  int *v289; // [rsp+6D0h] [rbp+5D0h]
  int v290; // [rsp+6D8h] [rbp+5D8h]
  int *v291; // [rsp+6E0h] [rbp+5E0h]
  int v292; // [rsp+6E8h] [rbp+5E8h]
  __int64 v293; // [rsp+6F0h] [rbp+5F0h]
  int v294; // [rsp+6F8h] [rbp+5F8h]
  const wchar_t *v295; // [rsp+700h] [rbp+600h]
  int *v296; // [rsp+708h] [rbp+608h]
  int v297; // [rsp+710h] [rbp+610h]
  int *v298; // [rsp+718h] [rbp+618h]
  int v299; // [rsp+720h] [rbp+620h]
  __int64 v300; // [rsp+728h] [rbp+628h]
  int v301; // [rsp+730h] [rbp+630h]
  const wchar_t *v302; // [rsp+738h] [rbp+638h]
  int *v303; // [rsp+740h] [rbp+640h]
  int v304; // [rsp+748h] [rbp+648h]
  int *v305; // [rsp+750h] [rbp+650h]
  int v306; // [rsp+758h] [rbp+658h]
  __int64 v307; // [rsp+760h] [rbp+660h]
  int v308; // [rsp+768h] [rbp+668h]
  const wchar_t *v309; // [rsp+770h] [rbp+670h]
  int *v310; // [rsp+778h] [rbp+678h]
  int v311; // [rsp+780h] [rbp+680h]
  int *v312; // [rsp+788h] [rbp+688h]
  int v313; // [rsp+790h] [rbp+690h]
  __int64 v314; // [rsp+798h] [rbp+698h]
  int v315; // [rsp+7A0h] [rbp+6A0h]
  const wchar_t *v316; // [rsp+7A8h] [rbp+6A8h]
  int *v317; // [rsp+7B0h] [rbp+6B0h]
  int v318; // [rsp+7B8h] [rbp+6B8h]
  int *v319; // [rsp+7C0h] [rbp+6C0h]
  int v320; // [rsp+7C8h] [rbp+6C8h]
  __int64 v321; // [rsp+7D0h] [rbp+6D0h]
  int v322; // [rsp+7D8h] [rbp+6D8h]
  const wchar_t *v323; // [rsp+7E0h] [rbp+6E0h]
  int *v324; // [rsp+7E8h] [rbp+6E8h]
  int v325; // [rsp+7F0h] [rbp+6F0h]
  int *v326; // [rsp+7F8h] [rbp+6F8h]
  int v327; // [rsp+800h] [rbp+700h]
  __int64 v328; // [rsp+808h] [rbp+708h]
  int v329; // [rsp+810h] [rbp+710h]
  const wchar_t *v330; // [rsp+818h] [rbp+718h]
  int *v331; // [rsp+820h] [rbp+720h]
  int v332; // [rsp+828h] [rbp+728h]
  int *v333; // [rsp+830h] [rbp+730h]
  int v334; // [rsp+838h] [rbp+738h]
  __int64 v335; // [rsp+840h] [rbp+740h]
  int v336; // [rsp+848h] [rbp+748h]
  const wchar_t *v337; // [rsp+850h] [rbp+750h]
  int *v338; // [rsp+858h] [rbp+758h]
  int v339; // [rsp+860h] [rbp+760h]
  int *v340; // [rsp+868h] [rbp+768h]
  int v341; // [rsp+870h] [rbp+770h]
  __int64 v342; // [rsp+878h] [rbp+778h]
  int v343; // [rsp+880h] [rbp+780h]
  const wchar_t *v344; // [rsp+888h] [rbp+788h]
  int *v345; // [rsp+890h] [rbp+790h]
  int v346; // [rsp+898h] [rbp+798h]
  int *v347; // [rsp+8A0h] [rbp+7A0h]
  int v348; // [rsp+8A8h] [rbp+7A8h]
  __int64 v349; // [rsp+8B0h] [rbp+7B0h]
  int v350; // [rsp+8B8h] [rbp+7B8h]
  const wchar_t *v351; // [rsp+8C0h] [rbp+7C0h]
  int *v352; // [rsp+8C8h] [rbp+7C8h]
  int v353; // [rsp+8D0h] [rbp+7D0h]
  int *v354; // [rsp+8D8h] [rbp+7D8h]
  int v355; // [rsp+8E0h] [rbp+7E0h]
  __int64 v356; // [rsp+8E8h] [rbp+7E8h]
  int v357; // [rsp+8F0h] [rbp+7F0h]
  const wchar_t *v358; // [rsp+8F8h] [rbp+7F8h]
  unsigned int *v359; // [rsp+900h] [rbp+800h]
  int v360; // [rsp+908h] [rbp+808h]
  int *v361; // [rsp+910h] [rbp+810h]
  int v362; // [rsp+918h] [rbp+818h]
  __int64 v363; // [rsp+920h] [rbp+820h]
  int v364; // [rsp+928h] [rbp+828h]
  const wchar_t *v365; // [rsp+930h] [rbp+830h]
  int *v366; // [rsp+938h] [rbp+838h]
  int v367; // [rsp+940h] [rbp+840h]
  int *v368; // [rsp+948h] [rbp+848h]
  int v369; // [rsp+950h] [rbp+850h]
  __int64 v370; // [rsp+958h] [rbp+858h]
  int v371; // [rsp+960h] [rbp+860h]
  const wchar_t *v372; // [rsp+968h] [rbp+868h]
  int *v373; // [rsp+970h] [rbp+870h]
  int v374; // [rsp+978h] [rbp+878h]
  int *v375; // [rsp+980h] [rbp+880h]
  int v376; // [rsp+988h] [rbp+888h]
  __int64 v377; // [rsp+990h] [rbp+890h]
  int v378; // [rsp+998h] [rbp+898h]
  const wchar_t *v379; // [rsp+9A0h] [rbp+8A0h]
  int *v380; // [rsp+9A8h] [rbp+8A8h]
  int v381; // [rsp+9B0h] [rbp+8B0h]
  int *v382; // [rsp+9B8h] [rbp+8B8h]
  int v383; // [rsp+9C0h] [rbp+8C0h]
  __int64 v384; // [rsp+9C8h] [rbp+8C8h]
  int v385; // [rsp+9D0h] [rbp+8D0h]
  const wchar_t *v386; // [rsp+9D8h] [rbp+8D8h]
  int *v387; // [rsp+9E0h] [rbp+8E0h]
  int v388; // [rsp+9E8h] [rbp+8E8h]
  int *v389; // [rsp+9F0h] [rbp+8F0h]
  int v390; // [rsp+9F8h] [rbp+8F8h]
  __int64 v391; // [rsp+A00h] [rbp+900h]
  int v392; // [rsp+A08h] [rbp+908h]
  const wchar_t *v393; // [rsp+A10h] [rbp+910h]
  int *v394; // [rsp+A18h] [rbp+918h]
  int v395; // [rsp+A20h] [rbp+920h]
  int *v396; // [rsp+A28h] [rbp+928h]
  int v397; // [rsp+A30h] [rbp+930h]
  __int64 v398; // [rsp+A38h] [rbp+938h]
  int v399; // [rsp+A40h] [rbp+940h]
  const wchar_t *v400; // [rsp+A48h] [rbp+948h]
  int *v401; // [rsp+A50h] [rbp+950h]
  int v402; // [rsp+A58h] [rbp+958h]
  int *v403; // [rsp+A60h] [rbp+960h]
  int v404; // [rsp+A68h] [rbp+968h]
  __int64 v405; // [rsp+A70h] [rbp+970h]
  int v406; // [rsp+A78h] [rbp+978h]
  const wchar_t *v407; // [rsp+A80h] [rbp+980h]
  int *v408; // [rsp+A88h] [rbp+988h]
  int v409; // [rsp+A90h] [rbp+990h]
  int *v410; // [rsp+A98h] [rbp+998h]
  int v411; // [rsp+AA0h] [rbp+9A0h]
  __int64 v412; // [rsp+AA8h] [rbp+9A8h]
  int v413; // [rsp+AB0h] [rbp+9B0h]
  const wchar_t *v414; // [rsp+AB8h] [rbp+9B8h]
  int *v415; // [rsp+AC0h] [rbp+9C0h]
  int v416; // [rsp+AC8h] [rbp+9C8h]
  int *v417; // [rsp+AD0h] [rbp+9D0h]
  int v418; // [rsp+AD8h] [rbp+9D8h]
  __int64 v419; // [rsp+AE0h] [rbp+9E0h]
  int v420; // [rsp+AE8h] [rbp+9E8h]
  const wchar_t *v421; // [rsp+AF0h] [rbp+9F0h]
  int *v422; // [rsp+AF8h] [rbp+9F8h]
  int v423; // [rsp+B00h] [rbp+A00h]
  int *v424; // [rsp+B08h] [rbp+A08h]
  int v425; // [rsp+B10h] [rbp+A10h]
  __int64 v426; // [rsp+B18h] [rbp+A18h]
  int v427; // [rsp+B20h] [rbp+A20h]
  const wchar_t *v428; // [rsp+B28h] [rbp+A28h]
  int *v429; // [rsp+B30h] [rbp+A30h]
  int v430; // [rsp+B38h] [rbp+A38h]
  int *v431; // [rsp+B40h] [rbp+A40h]
  int v432; // [rsp+B48h] [rbp+A48h]
  __int64 v433; // [rsp+B50h] [rbp+A50h]
  int v434; // [rsp+B58h] [rbp+A58h]
  const wchar_t *v435; // [rsp+B60h] [rbp+A60h]
  int *v436; // [rsp+B68h] [rbp+A68h]
  int v437; // [rsp+B70h] [rbp+A70h]
  int *v438; // [rsp+B78h] [rbp+A78h]
  int v439; // [rsp+B80h] [rbp+A80h]
  __int64 v440; // [rsp+B88h] [rbp+A88h]
  int v441; // [rsp+B90h] [rbp+A90h]
  const wchar_t *v442; // [rsp+B98h] [rbp+A98h]
  int *v443; // [rsp+BA0h] [rbp+AA0h]
  int v444; // [rsp+BA8h] [rbp+AA8h]
  int *v445; // [rsp+BB0h] [rbp+AB0h]
  int v446; // [rsp+BB8h] [rbp+AB8h]
  __int64 v447; // [rsp+BC0h] [rbp+AC0h]
  int v448; // [rsp+BC8h] [rbp+AC8h]
  const wchar_t *v449; // [rsp+BD0h] [rbp+AD0h]
  int *v450; // [rsp+BD8h] [rbp+AD8h]
  int v451; // [rsp+BE0h] [rbp+AE0h]
  int *v452; // [rsp+BE8h] [rbp+AE8h]
  int v453; // [rsp+BF0h] [rbp+AF0h]
  __int64 v454; // [rsp+BF8h] [rbp+AF8h]
  int v455; // [rsp+C00h] [rbp+B00h]
  const wchar_t *v456; // [rsp+C08h] [rbp+B08h]
  int *v457; // [rsp+C10h] [rbp+B10h]
  int v458; // [rsp+C18h] [rbp+B18h]
  int *v459; // [rsp+C20h] [rbp+B20h]
  int v460; // [rsp+C28h] [rbp+B28h]
  __int64 v461; // [rsp+C30h] [rbp+B30h]
  int v462; // [rsp+C38h] [rbp+B38h]
  __int64 v463; // [rsp+C40h] [rbp+B40h]
  __int128 v464; // [rsp+C48h] [rbp+B48h]
  __int128 v465; // [rsp+C58h] [rbp+B58h]

  v123 = 16LL;
  v77 = 0;
  v122 = 1395864371LL;
  v124 = 1395864371LL;
  v66 = 0;
  v78 = 0;
  v83 = 7000;
  v45 = 7000;
  v109 = 30000;
  v53 = 30000;
  v110 = 5000;
  v54 = 5000;
  v111 = 500;
  v67 = 0;
  v82 = 0;
  v68 = 0;
  v80 = 0;
  v39 = 0;
  v35 = 0;
  v34 = 0;
  v36 = 0;
  v32 = 0;
  v79 = 1;
  v37 = 1;
  v81 = 0;
  v38 = 0;
  v84 = 0;
  v40 = 0;
  v85 = 0;
  v41 = 0;
  v86 = 0;
  v42 = 0;
  v87 = 0;
  v43 = 0;
  v88 = 0;
  v44 = 0;
  v89 = 1;
  v46 = 1;
  v90 = 0;
  v74 = 0;
  v91 = 0;
  v47 = 0;
  v93 = 0;
  v48 = 0;
  v92 = 0;
  v49 = 0;
  v94 = 0;
  v75 = 0;
  v95 = 1;
  v69 = 1;
  v96 = 0;
  v70 = 0;
  v98 = 0;
  v97 = 0;
  v99 = 0;
  v72 = 0;
  v101 = 0;
  v100 = 0;
  v102 = 0;
  v73 = 0;
  v103 = 0;
  v71 = 0;
  v104 = 0;
  v50 = 0;
  v105 = 0;
  v51 = 0;
  v106 = 0;
  v76 = 0;
  v107 = 1;
  v33 = 1;
  v108 = 0;
  v52 = 0;
  v55 = 500;
  v112 = 0;
  v56 = 0;
  v113 = 0;
  v127 = L"ForceDirectFlip";
  v65 = 0;
  v128 = &v66;
  v130 = &v77;
  v134 = L"DisableOverlays";
  v135 = &v67;
  v137 = &v78;
  v141 = L"EnableOfferReclaimOnDriver";
  v142 = &v37;
  v144 = &v79;
  v148 = L"LeanMemoryLimit";
  v149 = &v123;
  v151 = &v122;
  v155 = L"ForceEnableDxgMms2";
  v156 = &v39;
  v158 = &v80;
  v162 = L"ContextNoPatchMode";
  v163 = &v38;
  v114 = 2;
  v57 = 2;
  v115 = 1;
  v58 = 1;
  v116 = 0;
  v59 = 0;
  v117 = 0;
  v60 = 0;
  v118 = 1;
  v61 = 1;
  v119 = 1;
  v62 = 1;
  v120 = 1;
  v63 = 1;
  v121 = 1;
  v64 = 1;
  v125 = 0LL;
  v126 = 288;
  v129 = 67108868;
  v131 = 4;
  v132 = 0LL;
  v133 = 288;
  v136 = 67108868;
  v138 = 4;
  v139 = 0LL;
  v140 = 288;
  v143 = 67108868;
  v145 = 4;
  v146 = 0LL;
  v147 = 288;
  v150 = 184549387;
  v152 = 8;
  v153 = 0LL;
  v154 = 288;
  v157 = 67108868;
  v159 = 4;
  v160 = 0LL;
  v161 = 288;
  v164 = 67108868;
  v165 = &v81;
  v166 = 4;
  v170 = &v34;
  v167 = 0LL;
  v172 = &v35;
  v177 = &v32;
  v179 = &v36;
  v183 = L"Force32BitFences";
  v184 = &v68;
  v186 = &v82;
  v190 = L"InitialPagingQueueFenceValue";
  v191 = &v45;
  v193 = &v83;
  v197 = L"ForceInitPagingProcessVaSpace";
  v198 = &v40;
  v200 = &v84;
  v204 = L"DisableGdiContextGpuVa";
  v205 = &v41;
  v207 = &v85;
  v211 = L"DisablePagingContextGpuVa";
  v212 = &v42;
  v214 = &v86;
  v218 = L"DisableMonitoredFenceGpuVa";
  v219 = &v43;
  v168 = 288;
  v169 = L"ForceToMapGpuVa";
  v171 = 67108868;
  v173 = 4;
  v174 = 0LL;
  v175 = 288;
  v176 = L"ForceAccessedPhysically";
  v178 = 67108868;
  v180 = 4;
  v181 = 0LL;
  v182 = 288;
  v185 = 67108868;
  v187 = 4;
  v188 = 0LL;
  v189 = 288;
  v192 = 67108868;
  v194 = 4;
  v195 = 0LL;
  v196 = 288;
  v199 = 67108868;
  v201 = 4;
  v202 = 0LL;
  v203 = 288;
  v206 = 67108868;
  v208 = 4;
  v209 = 0LL;
  v210 = 288;
  v213 = 67108868;
  v215 = 4;
  v216 = 0LL;
  v217 = 288;
  v220 = 67108868;
  v222 = 4;
  v221 = &v87;
  v225 = L"ForceExplicitResidencyNotification";
  v226 = &v44;
  v228 = &v88;
  v233 = &v34;
  v235 = &v35;
  v240 = &v32;
  v242 = &v36;
  v246 = L"DriverManagesResidencyOverride";
  v247 = &v46;
  v249 = &v89;
  v253 = L"GdiPhysicalAdapterIndex";
  v254 = &v74;
  v256 = &v90;
  v260 = L"ForceReplicateGdiContent";
  v261 = &v47;
  v263 = &v91;
  v267 = L"EnableTimedCalls";
  v268 = &v49;
  v270 = &v92;
  v274 = L"CreateGdiPrimaryOnSlaveGpu";
  v275 = &v48;
  v277 = &v93;
  v223 = 0LL;
  v224 = 288;
  v227 = 67108868;
  v229 = 4;
  v230 = 0LL;
  v231 = 288;
  v232 = L"ForceToMapGpuVa";
  v234 = 67108868;
  v236 = 4;
  v237 = 0LL;
  v238 = 288;
  v239 = L"ForceAccessedPhysically";
  v241 = 67108868;
  v243 = 4;
  v244 = 0LL;
  v245 = 288;
  v248 = 67108868;
  v250 = 4;
  v251 = 0LL;
  v252 = 288;
  v255 = 67108868;
  v257 = 4;
  v258 = 0LL;
  v259 = 288;
  v262 = 67108868;
  v264 = 4;
  v265 = 0LL;
  v266 = 288;
  v269 = 67108868;
  v271 = 4;
  v272 = 0LL;
  v273 = 288;
  v276 = 67108868;
  v278 = 4;
  v279 = 0LL;
  v281 = L"ForceSurpriseRemovalSupport";
  v282 = &v75;
  v284 = &v94;
  v288 = L"EnableDecodeMPO";
  v289 = &v69;
  v291 = &v95;
  v295 = L"DisableBadDriverCheckForHwProtection";
  v296 = &v70;
  v298 = &v96;
  v302 = L"ForceSecondaryMPOSupport";
  v303 = &v97;
  v305 = &v98;
  v309 = L"ForceSecondaryIFlipSupport";
  v310 = &v72;
  v312 = &v99;
  v316 = L"EnablePanelFitterSupport";
  v317 = &v100;
  v319 = &v101;
  v323 = L"EnableMultiPlaneOverlay3DDIs";
  v324 = &v73;
  v326 = &v102;
  v330 = L"DisableSecondaryIFlipSupport";
  v331 = &v71;
  v333 = &v103;
  v280 = 288;
  v283 = 67108868;
  v285 = 4;
  v286 = 0LL;
  v287 = 288;
  v290 = 67108868;
  v292 = 4;
  v293 = 0LL;
  v294 = 288;
  v297 = 67108868;
  v299 = 4;
  v300 = 0LL;
  v301 = 288;
  v304 = 67108868;
  v306 = 4;
  v307 = 0LL;
  v308 = 288;
  v311 = 67108868;
  v313 = 4;
  v314 = 0LL;
  v315 = 288;
  v318 = 67108868;
  v320 = 4;
  v321 = 0LL;
  v322 = 288;
  v325 = 67108868;
  v327 = 4;
  v328 = 0LL;
  v329 = 288;
  v332 = 67108868;
  v334 = 4;
  v335 = 0LL;
  v336 = 288;
  v337 = L"EnableWDDM23Synchronization";
  v338 = &v50;
  v340 = &v104;
  v344 = L"IoMmuFlags";
  v345 = &v51;
  v347 = &v105;
  v351 = L"DisableMultiSourceMPOCheck";
  v352 = &v76;
  v354 = &v106;
  v358 = L"DriverStoreCopyMode";
  v359 = &v33;
  v361 = &v107;
  v365 = L"ForceVariableRefresh";
  v366 = &v52;
  v368 = &v108;
  v372 = L"DeadlockTimeout";
  v373 = &v53;
  v375 = &v109;
  v379 = L"DeadlockPulse";
  v380 = &v54;
  v382 = &v110;
  v386 = L"DeadlockPulseTolerance";
  v387 = &v55;
  v389 = &v111;
  v339 = 67108868;
  v341 = 4;
  v342 = 0LL;
  v343 = 288;
  v346 = 67108868;
  v348 = 4;
  v349 = 0LL;
  v350 = 288;
  v353 = 67108868;
  v355 = 4;
  v356 = 0LL;
  v357 = 288;
  v360 = 67108868;
  v362 = 4;
  v363 = 0LL;
  v364 = 288;
  v367 = 67108868;
  v369 = 4;
  v370 = 0LL;
  v371 = 288;
  v374 = 67108868;
  v376 = 4;
  v377 = 0LL;
  v378 = 288;
  v381 = 67108868;
  v383 = 4;
  v384 = 0LL;
  v385 = 288;
  v388 = 67108868;
  v390 = 4;
  v391 = 0LL;
  v392 = 288;
  v395 = 67108868;
  v393 = L"DisableIndependentVidPnVSync";
  v394 = &v56;
  v396 = &v112;
  v400 = L"NumVirtualFunctions";
  v401 = &v65;
  v403 = &v113;
  v407 = L"CrtcPhaseFrames";
  v408 = &v57;
  v410 = &v114;
  v414 = L"EnableFbrValidation";
  v415 = &v58;
  v417 = &v115;
  v421 = L"DisableBoostedVSyncVirtualization";
  v422 = &v59;
  v424 = &v116;
  v428 = L"EnableBasicRenderGpuPv";
  v429 = &v60;
  v431 = &v117;
  v435 = L"KnownProcessBoostMode";
  v436 = &v61;
  v438 = &v118;
  v442 = L"SmallQuantumMode";
  v443 = &v62;
  v445 = &v119;
  v397 = 4;
  v398 = 0LL;
  v399 = 288;
  v402 = 67108868;
  v404 = 4;
  v405 = 0LL;
  v406 = 288;
  v409 = 67108868;
  v411 = 4;
  v412 = 0LL;
  v413 = 288;
  v416 = 67108868;
  v418 = 4;
  v419 = 0LL;
  v420 = 288;
  v423 = 67108868;
  v425 = 4;
  v426 = 0LL;
  v427 = 288;
  v430 = 67108868;
  v432 = 4;
  v433 = 0LL;
  v434 = 288;
  v437 = 67108868;
  v439 = 4;
  v440 = 0LL;
  v441 = 288;
  v444 = 67108868;
  v446 = 4;
  v447 = 0LL;
  v448 = 288;
  v449 = L"HighPriorityCompletionMode";
  v451 = 67108868;
  v453 = 4;
  v450 = &v63;
  v458 = 67108868;
  v452 = &v120;
  v460 = 4;
  v456 = L"GpuPriorityChangeMode";
  v454 = 0LL;
  v457 = &v64;
  v459 = &v121;
  v455 = 288;
  v461 = 0LL;
  v462 = 0;
  v463 = 0LL;
  v464 = 0LL;
  v465 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v125);
  v4 = v39 == 0;
  *((_BYTE *)this + 3021) = v37 != 0;
  *((_DWORD *)this + 758) = v38;
  *((_QWORD *)this + 378) = v124;
  v5 = !v4;
  v4 = v32 == 0;
  *((_BYTE *)this + 3036) = v5;
  v6 = !v4;
  v4 = v40 == 0;
  *((_BYTE *)this + 3037) = v6;
  v7 = !v4;
  v4 = v41 == 0;
  *((_BYTE *)this + 3039) = v7;
  v8 = !v4;
  v4 = v42 == 0;
  *((_BYTE *)this + 3040) = v8;
  v9 = !v4;
  v4 = v43 == 0;
  *((_BYTE *)this + 3041) = v9;
  v10 = !v4;
  v4 = v44 == 0;
  *((_BYTE *)this + 3042) = v10;
  v11 = !v4;
  v4 = v46 == 0;
  *((_BYTE *)this + 3038) = v11;
  *((_DWORD *)this + 773) = v45;
  *((_BYTE *)this + 3043) = !v4;
  if ( v47 || (v12 = 0, (*((_DWORD *)this + 617) & 0x100) != 0) )
    v12 = 1;
  v4 = v48 == 0;
  *((_BYTE *)this + 3022) = v12;
  v13 = !v4;
  v4 = v49 == 0;
  *((_BYTE *)this + 3023) = v13;
  DXGADAPTER::Config = !v4 | DXGADAPTER::Config & 0xFE;
  if ( !v50 || (v14 = 1, *((int *)this + 684) < 8704) )
    v14 = 0;
  v15 = v33 < 2;
  *((_BYTE *)this + 3052) = v14;
  *((_DWORD *)this + 765) = v51;
  v16 = 2;
  if ( v15 )
    v16 = v33;
  v4 = v52 == 0;
  *((_DWORD *)this + 766) = v16;
  v17 = !v4;
  v4 = v56 == 0;
  *((_BYTE *)this + 3068) = v17;
  *((_DWORD *)this + 1226) = v53;
  *((_DWORD *)this + 1227) = v54;
  *((_DWORD *)this + 1228) = v55;
  v18 = !v4;
  v4 = v58 == 0;
  *((_BYTE *)this + 3220) = v18;
  *((_DWORD *)this + 1104) = v57;
  v19 = !v4;
  v4 = v59 == 0;
  *((_BYTE *)this + 3069) = v19;
  v20 = !v4;
  v4 = g_OSTestSigningEnabled == 0;
  *((_BYTE *)this + 3070) = v20;
  if ( v4 || (v21 = 1, !v60) )
    v21 = 0;
  *((_BYTE *)this + 3071) = v21;
  *((_DWORD *)this + 769) = v61;
  *((_DWORD *)this + 770) = v62;
  *((_DWORD *)this + 771) = v63;
  *((_DWORD *)this + 772) = v64;
  if ( v65 )
    *((_DWORD *)this + 1200) = v65;
  if ( v66 )
    *((_BYTE *)this + 2939) = 1;
  if ( v67 )
    *((_BYTE *)this + 2940) = 0;
  if ( v68 )
    *((_DWORD *)this + 616) |= 0x20u;
  if ( *((_BYTE *)this + 2940) )
  {
    if ( *((_BYTE *)this + 3018) )
      *((_DWORD *)this + 736) = 2;
  }
  else
  {
    *((_DWORD *)this + 736) = 1;
  }
  if ( *((int *)this + 684) < 4608 )
    *((_BYTE *)this + 3021) = 0;
  if ( !DXGADAPTER::IsDxgmms2(this) )
    *((_BYTE *)this + 3043) = 0;
  if ( !v69 || (v24 = 1, !v22) )
    v24 = 0;
  v4 = v70 == 0;
  *((_BYTE *)this + 3044) = v24;
  *((_BYTE *)this + 3047) = 0;
  *((_BYTE *)this + 3045) = !v4;
  if ( !v71 && (*((_DWORD *)this + 615) & 0x10) != 0 )
  {
    v25 = *((_DWORD *)this + 684);
    if ( v25 < 8448 )
    {
      if ( v25 >= 0x2000 )
        *((_BYTE *)this + 3047) = v72 != 0;
    }
    else
    {
      *((_BYTE *)this + 3047) = 1;
    }
  }
  v4 = *((_QWORD *)this + 80) == 0LL;
  *((_BYTE *)this + 3049) = 0;
  *((_BYTE *)this + 3056) = !v4;
  v26 = 0;
  if ( *((_QWORD *)this + 129) )
  {
    v27 = *((_DWORD *)this + 684);
    v26 = v27 >= v23 || v27 >= 8448 && ((*((_DWORD *)this + 111) & 0x200) != 0 || v73);
    *((_BYTE *)this + 3049) = v26;
  }
  v28 = *((_BYTE *)this + 2940);
  if ( v28 && !v26 && !*((_QWORD *)this + 109) && !*((_QWORD *)this + 125) )
  {
    *((_BYTE *)this + 2940) = 0;
    v28 = 0;
  }
  *((_BYTE *)this + 3050) = 0;
  if ( v26 && *((_DWORD *)this + 684) >= v23 && (*((_QWORD *)this + 153) || *((_QWORD *)this + 154)) )
  {
    v29 = (_DWORD *)((char *)this + 2972);
    *((_BYTE *)this + 3050) = 1;
    v30 = (_DWORD *)((char *)this + 2972);
  }
  else
  {
    v29 = (_DWORD *)((char *)this + 2972);
    v30 = (_DWORD *)((char *)this + 2972);
    if ( !v26 )
    {
      *v29 = 1;
      goto LABEL_59;
    }
  }
  if ( !*v29 )
    *v30 = 1;
LABEL_59:
  *((_BYTE *)this + 3048) = v26;
  if ( !v26 || (v31 = 1, !v28) )
    v31 = 0;
  *((_BYTE *)this + 3046) = v31;
  if ( *((_DWORD *)this + 74) > 1u )
    *((_DWORD *)this + 787) = v74;
  if ( v75 )
    *(_BYTE *)a2 |= 0x10u;
  *((_BYTE *)this + 3051) = v76 != 0;
}

__int64 __fastcall DXGADAPTER::InitializePowerManagement(DXGADAPTER *this, __int64 a2, __int64 a3)
{
  _BYTE *v3; // rbx
  __int64 v5; // rsi
  unsigned int v6; // r13d
  bool v7; // cc
  int v8; // r8d
  __int64 v9; // rcx
  unsigned __int64 v10; // rdx
  bool v11; // zf
  __int64 v12; // rcx
  unsigned int v13; // ebx
  unsigned int NumDifferentPhysicalAdapters; // r12d
  unsigned int v15; // edx
  __int64 v16; // rax
  DXGADAPTER *v17; // rcx
  int v18; // eax
  __int64 v20; // rax
  __int64 v21; // rax
  const wchar_t *v22; // r9
  __int64 v23; // rax
  char *v24; // r14
  unsigned int v25; // esi
  unsigned int v26; // ebx
  __int64 v27; // rax
  unsigned int v28; // edx
  unsigned int v29; // ecx
  __int64 v30; // rax
  int v31; // eax
  __int64 v32; // rcx
  __int16 v33; // dx
  __int64 v34; // rsi
  __int64 v35; // rcx
  unsigned int v36; // eax
  __int64 v37; // r12
  const wchar_t *v38; // r9
  int v39; // eax
  void *v40; // rcx
  char *v41; // r9
  __int64 v42; // rdx
  __int64 v43; // rax
  unsigned int v44; // r8d
  unsigned int v45; // r9d
  unsigned int v46; // ecx
  unsigned __int64 v47; // r9
  unsigned __int64 v48; // rcx
  __int64 v49; // rax
  __int64 v50; // rax
  __int64 v51; // rax
  unsigned int v52; // edx
  unsigned int j; // r8d
  __int64 v54; // r10
  __int64 v55; // r9
  unsigned int v56; // edx
  unsigned int v57; // ecx
  unsigned int v58; // eax
  __int64 v59; // rbx
  int v60; // eax
  __int64 v61; // rcx
  __int64 v62; // rsi
  unsigned int v63; // eax
  __int64 v64; // rax
  unsigned int v65; // ecx
  __int64 v66; // rdx
  __int64 v67; // rax
  void *v68; // rcx
  unsigned int v69; // eax
  unsigned int v70; // edx
  __int64 v71; // r8
  __int64 v72; // r10
  __int64 v73; // rax
  unsigned int v74; // ebx
  __int64 v75; // r9
  unsigned int k; // ecx
  __int64 v77; // r10
  __int64 v78; // rsi
  unsigned int v79; // r11d
  __int64 v80; // r12
  __int64 v81; // rbx
  __int64 v82; // rbx
  __int64 v83; // rbx
  ADAPTER_RENDER *v84; // rcx
  int v85; // eax
  __int64 v86; // r15
  const wchar_t *v87; // r9
  ADAPTER_DISPLAY *v88; // rcx
  int v89; // eax
  int v90; // eax
  ULONG TimeIncrement; // eax
  __int64 v92; // rcx
  unsigned __int64 v93; // r9
  __int64 v94; // rax
  unsigned __int64 v95; // rtt
  __int64 v96; // rdx
  __int64 v97; // rax
  __int64 v98; // rcx
  __int64 v99; // rax
  __int64 v100; // rcx
  __int64 v101; // rax
  __int64 v102; // rcx
  __int64 v103; // rax
  __int64 v104; // rcx
  __int64 v105; // rax
  unsigned __int64 v106; // rtt
  __int64 v107; // rax
  unsigned __int64 v108; // rtt
  __int64 v109; // rax
  __int64 v110; // rcx
  __int64 v111; // rax
  unsigned __int64 v112; // rtt
  __int64 v113; // rax
  __int64 v114; // rcx
  __int64 v115; // rax
  __int64 v116; // rcx
  __int64 v117; // rax
  __int64 v118; // rcx
  __int64 v119; // rax
  __int64 v120; // rcx
  __int64 v121; // rax
  __int64 v122; // rcx
  __int64 v123; // rax
  __int64 v124; // rcx
  __int64 v125; // rax
  __int64 v126; // rcx
  __int64 v127; // rax
  __int64 v128; // rcx
  __int64 v129; // rax
  __int64 v130; // rax
  __int64 v131; // r12
  __int64 v132; // rsi
  __int64 v133; // rbx
  DXGADAPTER *v134; // rdx
  int v135; // ecx
  int v136; // ecx
  int v137; // ecx
  int v138; // ecx
  int v139; // ecx
  int v140; // ecx
  DXGADAPTER *v141; // rcx
  unsigned int v142; // edx
  unsigned __int64 v143; // r8
  DXGADAPTER **v144; // rcx
  __int64 v145; // rax
  DXGADAPTER **v146; // rcx
  unsigned __int64 v147; // rcx
  unsigned int v148; // eax
  unsigned __int64 *v149; // rdx
  __int64 v150; // r8
  unsigned __int64 v151; // rax
  unsigned int i; // edx
  unsigned int v153; // edx
  unsigned __int64 v154; // r8
  __int64 v155; // rcx
  __int64 v156; // rax
  __int64 v157; // r8
  struct _SLIST_ENTRY *v158; // rbx
  __int64 v159; // rsi
  NTSTATUS v160; // eax
  int v161; // eax
  PCLIENT_ID ClientId; // [rsp+28h] [rbp-E0h]
  char v163; // [rsp+58h] [rbp-B0h]
  unsigned int v164; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v165; // [rsp+60h] [rbp-A8h]
  BOOL v166; // [rsp+64h] [rbp-A4h] BYREF
  unsigned int v167; // [rsp+68h] [rbp-A0h] BYREF
  unsigned int v168; // [rsp+6Ch] [rbp-9Ch] BYREF
  unsigned int v169; // [rsp+70h] [rbp-98h] BYREF
  unsigned int v170; // [rsp+74h] [rbp-94h] BYREF
  unsigned int v171; // [rsp+78h] [rbp-90h] BYREF
  unsigned int v172; // [rsp+7Ch] [rbp-8Ch]
  unsigned int v173; // [rsp+80h] [rbp-88h]
  int v174; // [rsp+84h] [rbp-84h]
  unsigned int v175; // [rsp+88h] [rbp-80h] BYREF
  unsigned int v176; // [rsp+8Ch] [rbp-7Ch] BYREF
  int v177; // [rsp+90h] [rbp-78h] BYREF
  int v178; // [rsp+94h] [rbp-74h] BYREF
  int v179; // [rsp+98h] [rbp-70h] BYREF
  int v180; // [rsp+9Ch] [rbp-6Ch] BYREF
  int v181; // [rsp+A0h] [rbp-68h] BYREF
  int v182; // [rsp+A4h] [rbp-64h] BYREF
  int v183; // [rsp+A8h] [rbp-60h] BYREF
  unsigned int v184; // [rsp+ACh] [rbp-5Ch] BYREF
  unsigned int v185; // [rsp+B0h] [rbp-58h] BYREF
  unsigned int v186; // [rsp+B4h] [rbp-54h] BYREF
  unsigned int v187; // [rsp+B8h] [rbp-50h] BYREF
  unsigned int v188; // [rsp+BCh] [rbp-4Ch] BYREF
  unsigned int v189; // [rsp+C0h] [rbp-48h] BYREF
  unsigned int v190; // [rsp+C4h] [rbp-44h] BYREF
  unsigned int v191; // [rsp+C8h] [rbp-40h] BYREF
  unsigned int v192; // [rsp+CCh] [rbp-3Ch] BYREF
  unsigned int v193; // [rsp+D0h] [rbp-38h] BYREF
  unsigned int v194; // [rsp+D4h] [rbp-34h] BYREF
  unsigned int v195; // [rsp+D8h] [rbp-30h] BYREF
  unsigned int v196; // [rsp+DCh] [rbp-2Ch] BYREF
  unsigned int v197; // [rsp+E0h] [rbp-28h] BYREF
  unsigned int v198; // [rsp+E4h] [rbp-24h] BYREF
  unsigned int v199; // [rsp+E8h] [rbp-20h] BYREF
  unsigned int v200; // [rsp+ECh] [rbp-1Ch] BYREF
  unsigned int v201; // [rsp+F0h] [rbp-18h] BYREF
  unsigned int v202; // [rsp+F4h] [rbp-14h] BYREF
  unsigned int v203; // [rsp+F8h] [rbp-10h] BYREF
  unsigned int v204; // [rsp+FCh] [rbp-Ch] BYREF
  unsigned int v205; // [rsp+100h] [rbp-8h] BYREF
  unsigned int v206; // [rsp+104h] [rbp-4h] BYREF
  unsigned int v207; // [rsp+108h] [rbp+0h] BYREF
  unsigned int v208; // [rsp+10Ch] [rbp+4h] BYREF
  unsigned int v209; // [rsp+110h] [rbp+8h] BYREF
  int v210; // [rsp+114h] [rbp+Ch] BYREF
  int v211; // [rsp+118h] [rbp+10h] BYREF
  int v212; // [rsp+11Ch] [rbp+14h] BYREF
  int v213; // [rsp+120h] [rbp+18h] BYREF
  int v214; // [rsp+124h] [rbp+1Ch] BYREF
  int v215; // [rsp+128h] [rbp+20h] BYREF
  int v216; // [rsp+12Ch] [rbp+24h] BYREF
  int v217; // [rsp+130h] [rbp+28h] BYREF
  int v218; // [rsp+134h] [rbp+2Ch] BYREF
  int v219; // [rsp+138h] [rbp+30h] BYREF
  int v220; // [rsp+13Ch] [rbp+34h] BYREF
  int v221; // [rsp+140h] [rbp+38h] BYREF
  int v222; // [rsp+144h] [rbp+3Ch] BYREF
  int v223; // [rsp+148h] [rbp+40h] BYREF
  int v224; // [rsp+14Ch] [rbp+44h] BYREF
  int v225; // [rsp+150h] [rbp+48h] BYREF
  int v226; // [rsp+154h] [rbp+4Ch] BYREF
  int v227; // [rsp+158h] [rbp+50h] BYREF
  int v228; // [rsp+15Ch] [rbp+54h] BYREF
  int v229; // [rsp+160h] [rbp+58h] BYREF
  int v230; // [rsp+164h] [rbp+5Ch] BYREF
  int v231; // [rsp+168h] [rbp+60h] BYREF
  int v232; // [rsp+16Ch] [rbp+64h] BYREF
  int v233; // [rsp+170h] [rbp+68h] BYREF
  int v234; // [rsp+174h] [rbp+6Ch] BYREF
  int v235; // [rsp+178h] [rbp+70h] BYREF
  int v236; // [rsp+17Ch] [rbp+74h] BYREF
  int v237; // [rsp+180h] [rbp+78h] BYREF
  int v238; // [rsp+184h] [rbp+7Ch] BYREF
  int v239; // [rsp+188h] [rbp+80h] BYREF
  int v240; // [rsp+18Ch] [rbp+84h] BYREF
  int v241; // [rsp+190h] [rbp+88h] BYREF
  int v242; // [rsp+194h] [rbp+8Ch] BYREF
  int v243; // [rsp+198h] [rbp+90h] BYREF
  int v244; // [rsp+19Ch] [rbp+94h] BYREF
  int v245; // [rsp+1A0h] [rbp+98h] BYREF
  int v246; // [rsp+1A4h] [rbp+9Ch] BYREF
  int v247; // [rsp+1A8h] [rbp+A0h] BYREF
  int v248; // [rsp+1ACh] [rbp+A4h] BYREF
  int v249; // [rsp+1B0h] [rbp+A8h] BYREF
  unsigned int v250; // [rsp+1B4h] [rbp+ACh] BYREF
  __int64 v251; // [rsp+1B8h] [rbp+B0h]
  void *v252; // [rsp+1C0h] [rbp+B8h]
  __int64 v253; // [rsp+1C8h] [rbp+C0h] BYREF
  __int64 v254; // [rsp+1D0h] [rbp+C8h]
  struct _DXGKARG_QUERYADAPTERINFO v255; // [rsp+1D8h] [rbp+D0h] BYREF
  __int64 v256; // [rsp+208h] [rbp+100h]
  __int64 v257; // [rsp+210h] [rbp+108h]
  struct _DXGKARG_QUERYADAPTERINFO v258; // [rsp+218h] [rbp+110h] BYREF
  struct _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+248h] [rbp+140h] BYREF
  __int64 v260; // [rsp+278h] [rbp+170h] BYREF
  int v261; // [rsp+280h] [rbp+178h]
  const wchar_t *v262; // [rsp+288h] [rbp+180h]
  BOOL *v263; // [rsp+290h] [rbp+188h]
  int v264; // [rsp+298h] [rbp+190h]
  int *v265; // [rsp+2A0h] [rbp+198h]
  int v266; // [rsp+2A8h] [rbp+1A0h]
  __int64 v267; // [rsp+2B0h] [rbp+1A8h]
  int v268; // [rsp+2B8h] [rbp+1B0h]
  __int64 v269; // [rsp+2C0h] [rbp+1B8h]
  __int128 v270; // [rsp+2C8h] [rbp+1C0h]
  __int128 v271; // [rsp+2D8h] [rbp+1D0h]
  __int64 v272; // [rsp+2E8h] [rbp+1E0h] BYREF
  int v273; // [rsp+2F0h] [rbp+1E8h]
  const wchar_t *v274; // [rsp+2F8h] [rbp+1F0h]
  BOOL *v275; // [rsp+300h] [rbp+1F8h]
  int v276; // [rsp+308h] [rbp+200h]
  int *v277; // [rsp+310h] [rbp+208h]
  int v278; // [rsp+318h] [rbp+210h]
  __int64 v279; // [rsp+320h] [rbp+218h]
  int v280; // [rsp+328h] [rbp+220h]
  __int64 v281; // [rsp+330h] [rbp+228h]
  __int128 v282; // [rsp+338h] [rbp+230h]
  __int128 v283; // [rsp+348h] [rbp+240h]
  __int64 v284; // [rsp+358h] [rbp+250h] BYREF
  int v285; // [rsp+360h] [rbp+258h]
  const wchar_t *v286; // [rsp+368h] [rbp+260h]
  int *v287; // [rsp+370h] [rbp+268h]
  int v288; // [rsp+378h] [rbp+270h]
  int *v289; // [rsp+380h] [rbp+278h]
  int v290; // [rsp+388h] [rbp+280h]
  __int64 v291; // [rsp+390h] [rbp+288h]
  int v292; // [rsp+398h] [rbp+290h]
  const wchar_t *v293; // [rsp+3A0h] [rbp+298h]
  int *v294; // [rsp+3A8h] [rbp+2A0h]
  int v295; // [rsp+3B0h] [rbp+2A8h]
  int *v296; // [rsp+3B8h] [rbp+2B0h]
  int v297; // [rsp+3C0h] [rbp+2B8h]
  __int64 v298; // [rsp+3C8h] [rbp+2C0h]
  int v299; // [rsp+3D0h] [rbp+2C8h]
  const wchar_t *v300; // [rsp+3D8h] [rbp+2D0h]
  unsigned int *v301; // [rsp+3E0h] [rbp+2D8h]
  int v302; // [rsp+3E8h] [rbp+2E0h]
  int *v303; // [rsp+3F0h] [rbp+2E8h]
  int v304; // [rsp+3F8h] [rbp+2F0h]
  __int64 v305; // [rsp+400h] [rbp+2F8h]
  int v306; // [rsp+408h] [rbp+300h]
  const wchar_t *v307; // [rsp+410h] [rbp+308h]
  unsigned int *v308; // [rsp+418h] [rbp+310h]
  int v309; // [rsp+420h] [rbp+318h]
  int *v310; // [rsp+428h] [rbp+320h]
  int v311; // [rsp+430h] [rbp+328h]
  __int64 v312; // [rsp+438h] [rbp+330h]
  int v313; // [rsp+440h] [rbp+338h]
  const wchar_t *v314; // [rsp+448h] [rbp+340h]
  unsigned int *v315; // [rsp+450h] [rbp+348h]
  int v316; // [rsp+458h] [rbp+350h]
  int *v317; // [rsp+460h] [rbp+358h]
  int v318; // [rsp+468h] [rbp+360h]
  __int64 v319; // [rsp+470h] [rbp+368h]
  int v320; // [rsp+478h] [rbp+370h]
  const wchar_t *v321; // [rsp+480h] [rbp+378h]
  unsigned int *v322; // [rsp+488h] [rbp+380h]
  int v323; // [rsp+490h] [rbp+388h]
  int *v324; // [rsp+498h] [rbp+390h]
  int v325; // [rsp+4A0h] [rbp+398h]
  __int64 v326; // [rsp+4A8h] [rbp+3A0h]
  int v327; // [rsp+4B0h] [rbp+3A8h]
  const wchar_t *v328; // [rsp+4B8h] [rbp+3B0h]
  unsigned int *v329; // [rsp+4C0h] [rbp+3B8h]
  int v330; // [rsp+4C8h] [rbp+3C0h]
  int *v331; // [rsp+4D0h] [rbp+3C8h]
  int v332; // [rsp+4D8h] [rbp+3D0h]
  __int64 v333; // [rsp+4E0h] [rbp+3D8h]
  int v334; // [rsp+4E8h] [rbp+3E0h]
  const wchar_t *v335; // [rsp+4F0h] [rbp+3E8h]
  unsigned int *v336; // [rsp+4F8h] [rbp+3F0h]
  int v337; // [rsp+500h] [rbp+3F8h]
  int *v338; // [rsp+508h] [rbp+400h]
  int v339; // [rsp+510h] [rbp+408h]
  __int64 v340; // [rsp+518h] [rbp+410h]
  int v341; // [rsp+520h] [rbp+418h]
  const wchar_t *v342; // [rsp+528h] [rbp+420h]
  unsigned int *v343; // [rsp+530h] [rbp+428h]
  int v344; // [rsp+538h] [rbp+430h]
  int *v345; // [rsp+540h] [rbp+438h]
  int v346; // [rsp+548h] [rbp+440h]
  __int64 v347; // [rsp+550h] [rbp+448h]
  int v348; // [rsp+558h] [rbp+450h]
  const wchar_t *v349; // [rsp+560h] [rbp+458h]
  unsigned int *v350; // [rsp+568h] [rbp+460h]
  int v351; // [rsp+570h] [rbp+468h]
  int *v352; // [rsp+578h] [rbp+470h]
  int v353; // [rsp+580h] [rbp+478h]
  __int64 v354; // [rsp+588h] [rbp+480h]
  int v355; // [rsp+590h] [rbp+488h]
  const wchar_t *v356; // [rsp+598h] [rbp+490h]
  int *v357; // [rsp+5A0h] [rbp+498h]
  int v358; // [rsp+5A8h] [rbp+4A0h]
  int *v359; // [rsp+5B0h] [rbp+4A8h]
  int v360; // [rsp+5B8h] [rbp+4B0h]
  __int64 v361; // [rsp+5C0h] [rbp+4B8h]
  int v362; // [rsp+5C8h] [rbp+4C0h]
  const wchar_t *v363; // [rsp+5D0h] [rbp+4C8h]
  unsigned int *v364; // [rsp+5D8h] [rbp+4D0h]
  int v365; // [rsp+5E0h] [rbp+4D8h]
  int *v366; // [rsp+5E8h] [rbp+4E0h]
  int v367; // [rsp+5F0h] [rbp+4E8h]
  __int64 v368; // [rsp+5F8h] [rbp+4F0h]
  int v369; // [rsp+600h] [rbp+4F8h]
  const wchar_t *v370; // [rsp+608h] [rbp+500h]
  int *v371; // [rsp+610h] [rbp+508h]
  int v372; // [rsp+618h] [rbp+510h]
  int *v373; // [rsp+620h] [rbp+518h]
  int v374; // [rsp+628h] [rbp+520h]
  __int64 v375; // [rsp+630h] [rbp+528h]
  int v376; // [rsp+638h] [rbp+530h]
  const wchar_t *v377; // [rsp+640h] [rbp+538h]
  unsigned int *v378; // [rsp+648h] [rbp+540h]
  int v379; // [rsp+650h] [rbp+548h]
  int *v380; // [rsp+658h] [rbp+550h]
  int v381; // [rsp+660h] [rbp+558h]
  __int64 v382; // [rsp+668h] [rbp+560h]
  int v383; // [rsp+670h] [rbp+568h]
  const wchar_t *v384; // [rsp+678h] [rbp+570h]
  unsigned int *v385; // [rsp+680h] [rbp+578h]
  int v386; // [rsp+688h] [rbp+580h]
  int *v387; // [rsp+690h] [rbp+588h]
  int v388; // [rsp+698h] [rbp+590h]
  __int64 v389; // [rsp+6A0h] [rbp+598h]
  int v390; // [rsp+6A8h] [rbp+5A0h]
  const wchar_t *v391; // [rsp+6B0h] [rbp+5A8h]
  unsigned int *v392; // [rsp+6B8h] [rbp+5B0h]
  int v393; // [rsp+6C0h] [rbp+5B8h]
  int *v394; // [rsp+6C8h] [rbp+5C0h]
  int v395; // [rsp+6D0h] [rbp+5C8h]
  __int64 v396; // [rsp+6D8h] [rbp+5D0h]
  int v397; // [rsp+6E0h] [rbp+5D8h]
  const wchar_t *v398; // [rsp+6E8h] [rbp+5E0h]
  unsigned int *v399; // [rsp+6F0h] [rbp+5E8h]
  int v400; // [rsp+6F8h] [rbp+5F0h]
  int *v401; // [rsp+700h] [rbp+5F8h]
  int v402; // [rsp+708h] [rbp+600h]
  __int64 v403; // [rsp+710h] [rbp+608h]
  int v404; // [rsp+718h] [rbp+610h]
  const wchar_t *v405; // [rsp+720h] [rbp+618h]
  unsigned int *v406; // [rsp+728h] [rbp+620h]
  int v407; // [rsp+730h] [rbp+628h]
  int *v408; // [rsp+738h] [rbp+630h]
  int v409; // [rsp+740h] [rbp+638h]
  __int64 v410; // [rsp+748h] [rbp+640h]
  int v411; // [rsp+750h] [rbp+648h]
  const wchar_t *v412; // [rsp+758h] [rbp+650h]
  unsigned int *v413; // [rsp+760h] [rbp+658h]
  int v414; // [rsp+768h] [rbp+660h]
  int *v415; // [rsp+770h] [rbp+668h]
  int v416; // [rsp+778h] [rbp+670h]
  __int64 v417; // [rsp+780h] [rbp+678h]
  int v418; // [rsp+788h] [rbp+680h]
  const wchar_t *v419; // [rsp+790h] [rbp+688h]
  unsigned int *v420; // [rsp+798h] [rbp+690h]
  int v421; // [rsp+7A0h] [rbp+698h]
  int *v422; // [rsp+7A8h] [rbp+6A0h]
  int v423; // [rsp+7B0h] [rbp+6A8h]
  __int64 v424; // [rsp+7B8h] [rbp+6B0h]
  int v425; // [rsp+7C0h] [rbp+6B8h]
  const wchar_t *v426; // [rsp+7C8h] [rbp+6C0h]
  unsigned int *v427; // [rsp+7D0h] [rbp+6C8h]
  int v428; // [rsp+7D8h] [rbp+6D0h]
  int *v429; // [rsp+7E0h] [rbp+6D8h]
  int v430; // [rsp+7E8h] [rbp+6E0h]
  __int64 v431; // [rsp+7F0h] [rbp+6E8h]
  int v432; // [rsp+7F8h] [rbp+6F0h]
  const wchar_t *v433; // [rsp+800h] [rbp+6F8h]
  int *v434; // [rsp+808h] [rbp+700h]
  int v435; // [rsp+810h] [rbp+708h]
  int *v436; // [rsp+818h] [rbp+710h]
  int v437; // [rsp+820h] [rbp+718h]
  __int64 v438; // [rsp+828h] [rbp+720h]
  int v439; // [rsp+830h] [rbp+728h]
  const wchar_t *v440; // [rsp+838h] [rbp+730h]
  int *v441; // [rsp+840h] [rbp+738h]
  int v442; // [rsp+848h] [rbp+740h]
  int *v443; // [rsp+850h] [rbp+748h]
  int v444; // [rsp+858h] [rbp+750h]
  __int64 v445; // [rsp+860h] [rbp+758h]
  int v446; // [rsp+868h] [rbp+760h]
  const wchar_t *v447; // [rsp+870h] [rbp+768h]
  int *v448; // [rsp+878h] [rbp+770h]
  int v449; // [rsp+880h] [rbp+778h]
  int *v450; // [rsp+888h] [rbp+780h]
  int v451; // [rsp+890h] [rbp+788h]
  __int64 v452; // [rsp+898h] [rbp+790h]
  int v453; // [rsp+8A0h] [rbp+798h]
  const wchar_t *v454; // [rsp+8A8h] [rbp+7A0h]
  unsigned int *v455; // [rsp+8B0h] [rbp+7A8h]
  int v456; // [rsp+8B8h] [rbp+7B0h]
  int *v457; // [rsp+8C0h] [rbp+7B8h]
  int v458; // [rsp+8C8h] [rbp+7C0h]
  __int64 v459; // [rsp+8D0h] [rbp+7C8h]
  int v460; // [rsp+8D8h] [rbp+7D0h]
  const wchar_t *v461; // [rsp+8E0h] [rbp+7D8h]
  unsigned int *v462; // [rsp+8E8h] [rbp+7E0h]
  int v463; // [rsp+8F0h] [rbp+7E8h]
  int *v464; // [rsp+8F8h] [rbp+7F0h]
  int v465; // [rsp+900h] [rbp+7F8h]
  __int64 v466; // [rsp+908h] [rbp+800h]
  int v467; // [rsp+910h] [rbp+808h]
  const wchar_t *v468; // [rsp+918h] [rbp+810h]
  unsigned int *v469; // [rsp+920h] [rbp+818h]
  int v470; // [rsp+928h] [rbp+820h]
  int *v471; // [rsp+930h] [rbp+828h]
  int v472; // [rsp+938h] [rbp+830h]
  __int64 v473; // [rsp+940h] [rbp+838h]
  int v474; // [rsp+948h] [rbp+840h]
  const wchar_t *v475; // [rsp+950h] [rbp+848h]
  unsigned int *v476; // [rsp+958h] [rbp+850h]
  int v477; // [rsp+960h] [rbp+858h]
  int *v478; // [rsp+968h] [rbp+860h]
  int v479; // [rsp+970h] [rbp+868h]
  __int64 v480; // [rsp+978h] [rbp+870h]
  int v481; // [rsp+980h] [rbp+878h]
  const wchar_t *v482; // [rsp+988h] [rbp+880h]
  unsigned int *v483; // [rsp+990h] [rbp+888h]
  int v484; // [rsp+998h] [rbp+890h]
  int *v485; // [rsp+9A0h] [rbp+898h]
  int v486; // [rsp+9A8h] [rbp+8A0h]
  __int64 v487; // [rsp+9B0h] [rbp+8A8h]
  int v488; // [rsp+9B8h] [rbp+8B0h]
  const wchar_t *v489; // [rsp+9C0h] [rbp+8B8h]
  unsigned int *v490; // [rsp+9C8h] [rbp+8C0h]
  int v491; // [rsp+9D0h] [rbp+8C8h]
  int *v492; // [rsp+9D8h] [rbp+8D0h]
  int v493; // [rsp+9E0h] [rbp+8D8h]
  __int64 v494; // [rsp+9E8h] [rbp+8E0h]
  int v495; // [rsp+9F0h] [rbp+8E8h]
  const wchar_t *v496; // [rsp+9F8h] [rbp+8F0h]
  unsigned int *v497; // [rsp+A00h] [rbp+8F8h]
  int v498; // [rsp+A08h] [rbp+900h]
  int *v499; // [rsp+A10h] [rbp+908h]
  int v500; // [rsp+A18h] [rbp+910h]
  __int64 v501; // [rsp+A20h] [rbp+918h]
  int v502; // [rsp+A28h] [rbp+920h]
  const wchar_t *v503; // [rsp+A30h] [rbp+928h]
  unsigned int *v504; // [rsp+A38h] [rbp+930h]
  int v505; // [rsp+A40h] [rbp+938h]
  int *v506; // [rsp+A48h] [rbp+940h]
  int v507; // [rsp+A50h] [rbp+948h]
  __int64 v508; // [rsp+A58h] [rbp+950h]
  int v509; // [rsp+A60h] [rbp+958h]
  const wchar_t *v510; // [rsp+A68h] [rbp+960h]
  unsigned int *v511; // [rsp+A70h] [rbp+968h]
  int v512; // [rsp+A78h] [rbp+970h]
  int *v513; // [rsp+A80h] [rbp+978h]
  int v514; // [rsp+A88h] [rbp+980h]
  __int64 v515; // [rsp+A90h] [rbp+988h]
  int v516; // [rsp+A98h] [rbp+990h]
  const wchar_t *v517; // [rsp+AA0h] [rbp+998h]
  unsigned int *v518; // [rsp+AA8h] [rbp+9A0h]
  int v519; // [rsp+AB0h] [rbp+9A8h]
  int *v520; // [rsp+AB8h] [rbp+9B0h]
  int v521; // [rsp+AC0h] [rbp+9B8h]
  __int64 v522; // [rsp+AC8h] [rbp+9C0h]
  int v523; // [rsp+AD0h] [rbp+9C8h]
  const wchar_t *v524; // [rsp+AD8h] [rbp+9D0h]
  unsigned int *v525; // [rsp+AE0h] [rbp+9D8h]
  int v526; // [rsp+AE8h] [rbp+9E0h]
  int *v527; // [rsp+AF0h] [rbp+9E8h]
  int v528; // [rsp+AF8h] [rbp+9F0h]
  __int64 v529; // [rsp+B00h] [rbp+9F8h]
  int v530; // [rsp+B08h] [rbp+A00h]
  const wchar_t *v531; // [rsp+B10h] [rbp+A08h]
  unsigned int *v532; // [rsp+B18h] [rbp+A10h]
  int v533; // [rsp+B20h] [rbp+A18h]
  int *v534; // [rsp+B28h] [rbp+A20h]
  int v535; // [rsp+B30h] [rbp+A28h]
  __int64 v536; // [rsp+B38h] [rbp+A30h]
  int v537; // [rsp+B40h] [rbp+A38h]
  const wchar_t *v538; // [rsp+B48h] [rbp+A40h]
  unsigned int *v539; // [rsp+B50h] [rbp+A48h]
  int v540; // [rsp+B58h] [rbp+A50h]
  int *v541; // [rsp+B60h] [rbp+A58h]
  int v542; // [rsp+B68h] [rbp+A60h]
  __int64 v543; // [rsp+B70h] [rbp+A68h]
  int v544; // [rsp+B78h] [rbp+A70h]
  const wchar_t *v545; // [rsp+B80h] [rbp+A78h]
  unsigned int *v546; // [rsp+B88h] [rbp+A80h]
  int v547; // [rsp+B90h] [rbp+A88h]
  int *v548; // [rsp+B98h] [rbp+A90h]
  int v549; // [rsp+BA0h] [rbp+A98h]
  __int64 v550; // [rsp+BA8h] [rbp+AA0h]
  int v551; // [rsp+BB0h] [rbp+AA8h]
  const wchar_t *v552; // [rsp+BB8h] [rbp+AB0h]
  unsigned int *v553; // [rsp+BC0h] [rbp+AB8h]
  int v554; // [rsp+BC8h] [rbp+AC0h]
  int *v555; // [rsp+BD0h] [rbp+AC8h]
  int v556; // [rsp+BD8h] [rbp+AD0h]
  __int64 v557; // [rsp+BE0h] [rbp+AD8h]
  int v558; // [rsp+BE8h] [rbp+AE0h]
  __int64 v559; // [rsp+BF0h] [rbp+AE8h]
  __int128 v560; // [rsp+BF8h] [rbp+AF0h]
  __int128 v561; // [rsp+C08h] [rbp+B00h]
  _DWORD v562[64]; // [rsp+C18h] [rbp+B10h] BYREF
  unsigned __int16 v563[264]; // [rsp+D18h] [rbp+C10h] BYREF

  v3 = (char *)this + 2941;
  if ( (Microsoft_Windows_DxgKrnlEnableBits & 0x20000) != 0 )
    McTemplateK0pt_EtwWriteTransfer(
      &DxgkControlGuid_Context,
      &Dxgk_PowerManagementSupport,
      a3,
      this,
      (unsigned __int8)*v3);
  if ( !*v3 )
  {
    WdLogSingleEntry0(3LL);
    WdLogGlobalForLineNumber = 4740;
    return 0LL;
  }
  v233 = 0;
  v212 = -1;
  v175 = -1;
  v213 = 2000;
  v216 = 35000;
  v186 = 35000;
  v176 = 2000;
  v218 = 50000;
  LOBYTE(v5) = 0;
  v189 = 50000;
  v6 = 0;
  v217 = 2000;
  v219 = 100000;
  v190 = 100000;
  v224 = 300000;
  v208 = 300000;
  v225 = 17000;
  v207 = 17000;
  v220 = 200;
  v183 = 200;
  v221 = 200;
  v187 = 200;
  v223 = 100;
  v222 = 100;
  v226 = 25000;
  v209 = 25000;
  v228 = 300;
  v170 = 300;
  v229 = 700;
  v169 = 700;
  v230 = 900;
  v168 = 900;
  v231 = 500;
  v167 = 500;
  v237 = 140000;
  v196 = 140000;
  v238 = 200000;
  v198 = 200000;
  v239 = 250000;
  v199 = 250000;
  v240 = 250000;
  v200 = 250000;
  v188 = 2000;
  v227 = 2000;
  v191 = 2000;
  v241 = 10000;
  v214 = 80;
  v184 = 80;
  v215 = 15000;
  v185 = 15000;
  v232 = 3;
  v182 = 3;
  v180 = 0;
  v234 = 0;
  v181 = 0;
  v235 = 80;
  v192 = 80;
  v236 = 80000;
  v194 = 80000;
  v7 = *((_DWORD *)this + 751) < 2400;
  v193 = 10000;
  v242 = 60000;
  v195 = 60000;
  v243 = 60000;
  v197 = 60000;
  v245 = 30000;
  v202 = 30000;
  v248 = 30000;
  v205 = 30000;
  v177 = 1;
  v166 = 1;
  v210 = 1;
  v178 = 1;
  v244 = 15000;
  v201 = 15000;
  v247 = 15000;
  v204 = 15000;
  v249 = 80000;
  v206 = 80000;
  v246 = 80;
  v203 = 80;
  v211 = 0;
  v179 = 0;
  if ( v7 )
  {
    v262 = L"UseSelfRefreshVRAMInS3";
    v261 = 288;
    v264 = 67108868;
    v263 = &v166;
    v260 = 0LL;
    v265 = &v177;
    v266 = 4;
    v267 = 0LL;
    v268 = 0;
    v269 = 0LL;
    v270 = 0LL;
    v271 = 0LL;
    RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\Power", &v260);
  }
  else
  {
    v166 = (*((_DWORD *)this + 617) & 0x1000) == 0;
  }
  v284 = 0LL;
  v286 = L"EnableRuntimePowerManagement";
  v287 = &v178;
  v289 = &v210;
  v293 = L"DisableDevicePowerRequired";
  v294 = &v179;
  v296 = &v211;
  v300 = L"DefaultLatencyToleranceOther";
  v301 = &v175;
  v303 = &v212;
  v307 = L"DefaultExpectedResidency";
  v308 = &v176;
  v310 = &v213;
  v314 = L"DefaultLatencyToleranceIdle0";
  v315 = &v184;
  v317 = &v214;
  v321 = L"DefaultLatencyToleranceIdle1";
  v322 = &v185;
  v324 = &v215;
  v328 = L"DefaultLatencyToleranceNoContext";
  v329 = &v186;
  v331 = &v216;
  v335 = L"DefaultLatencyToleranceIdle0MonitorOff";
  v336 = &v188;
  v338 = &v217;
  v285 = 288;
  v288 = 67108868;
  v290 = 4;
  v291 = 0LL;
  v292 = 288;
  v295 = 67108868;
  v297 = 4;
  v298 = 0LL;
  v299 = 288;
  v302 = 67108868;
  v304 = 4;
  v305 = 0LL;
  v306 = 288;
  v309 = 67108868;
  v311 = 4;
  v312 = 0LL;
  v313 = 288;
  v316 = 67108868;
  v318 = 4;
  v319 = 0LL;
  v320 = 288;
  v323 = 67108868;
  v325 = 4;
  v326 = 0LL;
  v327 = 288;
  v330 = 67108868;
  v332 = 4;
  v333 = 0LL;
  v334 = 288;
  v337 = 67108868;
  v339 = 4;
  v340 = 0LL;
  v341 = 288;
  v342 = L"DefaultLatencyToleranceIdle1MonitorOff";
  v343 = &v189;
  v345 = &v218;
  v349 = L"DefaultLatencyToleranceNoContextMonitorOff";
  v350 = &v190;
  v352 = &v219;
  v356 = L"DefaultLatencyToleranceTimerPeriod";
  v357 = &v183;
  v359 = &v220;
  v363 = L"DefaultIdleThresholdIdle0";
  v364 = &v187;
  v366 = &v221;
  v370 = L"DefaultIdleThresholdIdle0MonitorOff";
  v371 = &v222;
  v373 = &v223;
  v377 = L"MonitorLatencyTolerance";
  v378 = &v208;
  v380 = &v224;
  v384 = L"MonitorRefreshLatencyTolerance";
  v385 = &v207;
  v387 = &v225;
  v391 = L"DefaultPowerNotRequiredTimeout";
  v392 = &v209;
  v394 = &v226;
  v344 = 67108868;
  v346 = 4;
  v347 = 0LL;
  v348 = 288;
  v351 = 67108868;
  v353 = 4;
  v354 = 0LL;
  v355 = 288;
  v358 = 67108868;
  v360 = 4;
  v361 = 0LL;
  v362 = 288;
  v365 = 67108868;
  v367 = 4;
  v368 = 0LL;
  v369 = 288;
  v372 = 67108868;
  v374 = 4;
  v375 = 0LL;
  v376 = 288;
  v379 = 67108868;
  v381 = 4;
  v382 = 0LL;
  v383 = 288;
  v386 = 67108868;
  v388 = 4;
  v389 = 0LL;
  v390 = 288;
  v393 = 67108868;
  v395 = 4;
  v396 = 0LL;
  v397 = 288;
  v400 = 67108868;
  v398 = L"DefaultActiveIdleThreshold";
  v399 = &v191;
  v401 = &v227;
  v405 = L"ulow";
  v406 = &v170;
  v408 = &v228;
  v412 = L"uhigh";
  v413 = &v169;
  v415 = &v229;
  v419 = L"uglitch";
  v420 = &v168;
  v422 = &v230;
  v426 = L"uideal";
  v427 = &v167;
  v429 = &v231;
  v433 = L"lowdebounce";
  v434 = &v182;
  v436 = &v232;
  v440 = L"EnablePODebounce";
  v441 = &v180;
  v443 = &v233;
  v447 = L"DisablePStateManagement";
  v448 = &v181;
  v450 = &v234;
  v402 = 4;
  v403 = 0LL;
  v404 = 288;
  v407 = 67108868;
  v409 = 4;
  v410 = 0LL;
  v411 = 288;
  v414 = 67108868;
  v416 = 4;
  v417 = 0LL;
  v418 = 288;
  v421 = 67108868;
  v423 = 4;
  v424 = 0LL;
  v425 = 288;
  v428 = 67108868;
  v430 = 4;
  v431 = 0LL;
  v432 = 288;
  v435 = 67108868;
  v437 = 4;
  v438 = 0LL;
  v439 = 288;
  v442 = 67108868;
  v444 = 4;
  v445 = 0LL;
  v446 = 288;
  v449 = 67108868;
  v451 = 4;
  v452 = 0LL;
  v453 = 288;
  v454 = L"DefaultD3TransitionLatencyActivelyUsed";
  v455 = &v192;
  v457 = &v235;
  v461 = L"DefaultD3TransitionLatencyIdleShortTime";
  v462 = &v194;
  v464 = &v236;
  v468 = L"DefaultD3TransitionLatencyIdleLongTime";
  v469 = &v196;
  v471 = &v237;
  v475 = L"DefaultD3TransitionLatencyIdleVeryLongTime";
  v476 = &v198;
  v478 = &v238;
  v482 = L"DefaultD3TransitionLatencyIdleNoContext";
  v483 = &v199;
  v485 = &v239;
  v489 = L"DefaultD3TransitionLatencyIdleMonitorOff";
  v490 = &v200;
  v492 = &v240;
  v496 = L"DefaultD3TransitionIdleShortTimeThreshold";
  v497 = &v193;
  v499 = &v241;
  v503 = L"DefaultD3TransitionIdleLongTimeThreshold";
  v504 = &v195;
  v506 = &v242;
  v510 = L"DefaultD3TransitionIdleVeryLongTimeThreshold";
  v456 = 67108868;
  v458 = 4;
  v459 = 0LL;
  v460 = 288;
  v463 = 67108868;
  v465 = 4;
  v466 = 0LL;
  v467 = 288;
  v470 = 67108868;
  v472 = 4;
  v473 = 0LL;
  v474 = 288;
  v477 = 67108868;
  v479 = 4;
  v480 = 0LL;
  v481 = 288;
  v484 = 67108868;
  v486 = 4;
  v487 = 0LL;
  v488 = 288;
  v491 = 67108868;
  v493 = 4;
  v494 = 0LL;
  v495 = 288;
  v498 = 67108868;
  v500 = 4;
  v501 = 0LL;
  v502 = 288;
  v505 = 67108868;
  v507 = 4;
  v508 = 0LL;
  v509 = 288;
  v512 = 67108868;
  v511 = &v197;
  v516 = 288;
  v513 = &v243;
  v519 = 67108868;
  v517 = L"DefaultLatencyToleranceMemory";
  v523 = 288;
  v518 = &v201;
  v520 = &v244;
  v524 = L"DefaultLatencyToleranceMemoryNoContext";
  v525 = &v202;
  v527 = &v245;
  v531 = L"DefaultMemoryRefreshLatencyToleranceActivelyUsed";
  v532 = &v203;
  v534 = &v246;
  v538 = L"DefaultMemoryRefreshLatencyToleranceIdleShortTime";
  v539 = &v204;
  v541 = &v247;
  v545 = L"DefaultMemoryRefreshLatencyToleranceNoContext";
  v546 = &v205;
  v548 = &v248;
  v552 = L"DefaultMemoryRefreshLatencyToleranceMonitorOff";
  v553 = &v206;
  v526 = 67108868;
  v530 = 288;
  v533 = 67108868;
  v537 = 288;
  v540 = 67108868;
  v544 = 288;
  v547 = 67108868;
  v551 = 288;
  v554 = 67108868;
  v555 = &v249;
  v514 = 4;
  v515 = 0LL;
  v521 = 4;
  v522 = 0LL;
  v528 = 4;
  v529 = 0LL;
  v535 = 4;
  v536 = 0LL;
  v542 = 4;
  v543 = 0LL;
  v549 = 4;
  v550 = 0LL;
  v556 = 4;
  v557 = 0LL;
  v558 = 0;
  v559 = 0LL;
  v560 = 0LL;
  ClientId = 0LL;
  v561 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v284);
  if ( *((int *)this + 751) < 2400 )
  {
    v9 = *((_QWORD *)this + 27);
    v253 = 0LL;
    if ( (int)DpiGetPnpRegistryKeyName(v9, 2LL, &v253) >= 0
      && (int)RtlStringCbCopyW(v563, 0x208uLL, *(const unsigned __int16 **)(v253 + 8)) >= 0
      && (int)RtlStringCbCatW(v563, v10, L"\\DxgkSettings") >= 0 )
    {
      v272 = 0LL;
      v273 = 288;
      v274 = L"UseSelfRefreshVRAMInS3";
      v276 = 67108868;
      v275 = &v166;
      v278 = 4;
      v277 = &v177;
      v279 = 0LL;
      v280 = 0;
      v281 = 0LL;
      v282 = 0LL;
      ClientId = 0LL;
      v283 = 0LL;
      RtlQueryRegistryValuesEx(0LL, v563, &v272);
    }
  }
  if ( !v178 )
    return 0LL;
  v11 = !v166;
  *((_BYTE *)this + 204) = v179 != 0;
  *((_BYTE *)this + 207) = !v11;
  v12 = *(_QWORD *)(*((_QWORD *)this + 27) + 64LL);
  v13 = *(_DWORD *)(*(_QWORD *)(v12 + 40) + 28LL);
  if ( v13 < 0x5019 )
    NumDifferentPhysicalAdapters = 1;
  else
    NumDifferentPhysicalAdapters = DXGADAPTER::GetNumDifferentPhysicalAdapters(this);
  v164 = NumDifferentPhysicalAdapters;
  v15 = 0;
  v171 = 0;
  v16 = 0LL;
  while ( v15 < NumDifferentPhysicalAdapters )
  {
    v258.pOutputData = &v562[v16];
    memset(&v258, 0, 24);
    v258.Type = DXGKQAITYPE_NUMPOWERCOMPONENTS;
    *(_OWORD *)&v258.OutputDataSize = 0LL;
    v258.OutputDataSize = 4;
    if ( DXGADAPTER::IsDxgmms2(this) )
    {
      if ( v13 >= 0x5019 )
      {
        v258.InputDataSize = 4;
        v258.pInputData = &v171;
      }
    }
    v18 = DXGADAPTER::DdiQueryAdapterInfo(v17, &v258);
    v5 = v18;
    if ( v18 < 0 )
    {
      WdLogSingleEntry2(2LL, this, v18);
      WdLogGlobalForLineNumber = 4937;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"DdiQueryAdapterInfo failed. Adapter: 0x%p Status: 0x%I64x",
        (__int64)this,
        v5,
        0LL,
        0LL,
        0LL);
      return (unsigned int)v5;
    }
    v6 += v562[v171];
    v15 = v171 + 1;
    v171 = v15;
    v16 = v15;
  }
  if ( (Microsoft_Windows_DxgKrnlEnableBits & 0x20000) != 0 )
    McTemplateK0pqq_EtwWriteTransfer(v12, (unsigned int)&Dxgk_PowerManagementComponents, v8, (_DWORD)this, v5, v6);
  if ( !v6 )
  {
    WdLogSingleEntry0(3LL);
    WdLogGlobalForLineNumber = 4952;
    return 0LL;
  }
  if ( v6 > 0xFFFF )
  {
    WdLogSingleEntry1(2LL, v6);
    WdLogGlobalForLineNumber = 4958;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Miniport returned invalid number of power components:0x%I64x",
      v6,
      0LL,
      0LL,
      0LL,
      0LL);
    LODWORD(v5) = -1073741811;
    goto LABEL_212;
  }
  *((_DWORD *)this + 842) = v6;
  v20 = 520LL * v6;
  if ( !is_mul_ok(v6, 0x208uLL) )
    v20 = -1LL;
  v21 = operator new[](v20, 1265072196LL, 64LL);
  *((_QWORD *)this + 403) = v21;
  if ( !v21 )
  {
    WdLogSingleEntry1(6LL, this);
    v22 = L"Adapter 0x%I64x: Out of memory allocating m_pPowerComponents";
    WdLogGlobalForLineNumber = 4968;
LABEL_36:
    DxgkLogInternalTriageEvent(0, 262145, -1, (_DWORD)v22, (__int64)this, 0LL, 0LL, 0LL, 0LL);
    LODWORD(v5) = -1073741801;
    goto LABEL_212;
  }
  v23 = operator new[](312 * v6 + 160, 1265072196LL, 256LL);
  v24 = (char *)v23;
  if ( !v23 )
  {
    WdLogSingleEntry1(6LL, this);
    v22 = L"Adapter 0x%I64x: Out of memory allocating pRegistrationInfo";
    WdLogGlobalForLineNumber = 4985;
    goto LABEL_36;
  }
  *(_DWORD *)v23 = 3;
  *(_QWORD *)(v23 + 8) = 2LL;
  v25 = 0;
  *(_DWORD *)(v23 + 96) = v6;
  *(_QWORD *)(v23 + 64) = DxgkPowerRuntimeDeviceDirectedPowerUpCallback;
  *(_QWORD *)(v23 + 88) = this;
  v254 = v23 + 56LL * v6 + 104;
  *(_QWORD *)(v23 + 72) = DxgkPowerRuntimeDeviceDirectedPowerDownCallback;
  v26 = 0;
  v172 = 0;
  *(_QWORD *)(v23 + 32) = DxgkPowerRuntimeComponentIdleStateCallback;
  *(_QWORD *)(v23 + 16) = DxgkPowerRuntimeComponentActiveCallback;
  *(_QWORD *)(v23 + 24) = DxgkPowerRuntimeComponentIdleCallback;
  *(_QWORD *)(v23 + 40) = DxgkPowerRuntimeDevicePowerRequiredCallback;
  *(_QWORD *)(v23 + 48) = DxgkPowerRuntimeDevicePowerNotRequiredCallback;
  *(_QWORD *)(v23 + 56) = DxgkPowerRuntimeControlCallback;
  v252 = (void *)(v23 + 56LL * v6 + 104 + 192LL * v6);
  v27 = 0LL;
  memset(&v255, 0, sizeof(v255));
  v28 = 0;
  v255.Type = DXGKQAITYPE_POWERCOMPONENTINFO;
  v255.InputDataSize = 4;
  v255.OutputDataSize = 336;
  while ( 1 )
  {
    v165 = v28;
    v174 = v27;
    if ( (unsigned int)v27 >= NumDifferentPhysicalAdapters )
      break;
    v257 = v27;
    v29 = 0;
    *((_WORD *)this + v27 + 1620) = v26;
    while ( 1 )
    {
      v173 = v29;
      if ( v29 >= v562[v27] )
        break;
      v251 = 56LL * v26;
      v250 = v29 + v28;
      v255.pInputData = &v250;
      v30 = *((_QWORD *)this + 403);
      v256 = 520LL * v26;
      v255.pOutputData = (void *)(v30 + 8 + v256);
      v31 = DXGADAPTER::DdiQueryAdapterInfo(this, &v255);
      v5 = v31;
      if ( v31 < 0 )
      {
        WdLogSingleEntry2(2LL, v26, v31);
        WdLogGlobalForLineNumber = 5052;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Miniport failed QueryAdapterInfo(DXGKQAITYPE_POWERCOMPONENTINFO). Component: 0x%I64x, Status: 0x%I64x",
          v26,
          v5,
          0LL,
          0LL,
          0LL);
        goto LABEL_211;
      }
      v32 = v256;
      v33 = v173;
      *(_DWORD *)(v256 + *((_QWORD *)this + 403)) = v26;
      *(_WORD *)(*((_QWORD *)this + 403) + v32 + 4) = v33;
      *(_WORD *)(*((_QWORD *)this + 403) + v32 + 6) = v174;
      v34 = v32 + *((_QWORD *)this + 403);
      v35 = v251;
      *(_DWORD *)&v24[v251 + 132] = *(_DWORD *)(v34 + 8);
      if ( (unsigned int)(*(_DWORD *)(v34 + 8) - 1) > 7 )
      {
        WdLogSingleEntry3(2LL, v26, *(unsigned int *)(v34 + 8), 0LL, ClientId);
        v49 = *(unsigned int *)(v34 + 8);
        WdLogGlobalForLineNumber = 5066;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Miniport returned invalid number of F states for component:0x%I64x 0x%I64x",
          v26,
          v49,
          0LL,
          0LL,
          0LL);
        goto LABEL_103;
      }
      *(_OWORD *)&v24[v35 + 104] = *(_OWORD *)(v34 + 220);
      *(_BYTE *)(v34 + 275) = 0;
      v36 = *(_DWORD *)(v34 + 216);
      if ( v36 >= 0x20 )
      {
        WdLogSingleEntry2(2LL, v26, 2LL);
        WdLogGlobalForLineNumber = 5080;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Reserved flags are not zero. Component:0x%I64x",
          v26,
          2LL,
          0LL,
          0LL,
          0LL);
        goto LABEL_103;
      }
      v37 = v35;
      if ( (v36 & 4) != 0 )
        *(_QWORD *)&v24[v35 + 120] |= 1uLL;
      if ( !v180 )
        *(_QWORD *)&v24[v35 + 120] |= 2uLL;
      if ( (*(_DWORD *)(v34 + 216) & 0x10) != 0 )
      {
        if ( ((*(_DWORD *)(v34 + 208) - 3) & 0xFFFFFFFB) != 0 )
        {
          WdLogSingleEntry1(2LL, v26);
          v38 = L"Power component ActiveInD3 flag can only be used with DXGK_POWER_COMPONENT_MEMORY and DXGK_POWER_COMPONE"
                 "NT_SHARED. Component:0x%I64x";
          WdLogGlobalForLineNumber = 5099;
          goto LABEL_56;
        }
        if ( *(_DWORD *)(v34 + 8) != 2 )
        {
          WdLogSingleEntry1(2LL, v26);
          v38 = L"F state count must be 2 when the ActiveInD3 flag is set. Component:0x%I64x";
          WdLogGlobalForLineNumber = 5105;
          goto LABEL_56;
        }
        if ( *(_QWORD *)(v34 + 40) )
        {
          WdLogSingleEntry1(2LL, v26);
          v38 = L"TransitionLatency for the F1 state must be 0 when the ActiveInD3 flag is set. Component:0x%I64x";
          WdLogGlobalForLineNumber = 5111;
          goto LABEL_56;
        }
        if ( *(_DWORD *)(v34 + 276) )
        {
          WdLogSingleEntry1(2LL, v26);
          v38 = L"Provider count must be 0 when the ActiveInD3 flag is set. Component:0x%I64x";
          WdLogGlobalForLineNumber = 5117;
LABEL_56:
          DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v38, v26, 0LL, 0LL, 0LL, 0LL);
          goto LABEL_57;
        }
      }
      else if ( *(_DWORD *)(v34 + 276) > 0x10u )
      {
        WdLogSingleEntry2(2LL, v26, 3LL);
        WdLogGlobalForLineNumber = 5125;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Invalid component ProviderCount. Component:0x%I64x",
          v26,
          3LL,
          0LL,
          0LL,
          0LL);
        goto LABEL_57;
      }
      v39 = *(_DWORD *)(v34 + 208);
      if ( v39 == 4 )
      {
        if ( *((_DWORD *)this + 844) != -1 )
        {
          WdLogSingleEntry1(2LL, v26);
          v38 = L"DXGK_POWER_COMPONENT_MEMORY_REFRESH component is defined second time. Component:0x%I64x";
          WdLogGlobalForLineNumber = 5165;
          goto LABEL_56;
        }
        *((_DWORD *)this + 844) = v26;
      }
      else if ( v39 == 6 )
      {
        if ( *((_DWORD *)this + 843) == -1 )
        {
          *((_DWORD *)this + 843) = v26;
          *((_QWORD *)this + 448) = *((_QWORD *)this + 403) + 520LL * v26;
          if ( *(_DWORD *)(v34 + 8) == 2 )
          {
            *((_BYTE *)this + 3664) = 1;
          }
          else if ( *(_DWORD *)(v34 + 8) > 2u )
          {
            WdLogSingleEntry1(2LL, v26);
            v38 = L"F state count for the DXGK_POWER_COMPONENT_D3_TRANSITION component must be 1 or 2. Component:0x%I64x";
            WdLogGlobalForLineNumber = 5155;
            goto LABEL_56;
          }
        }
        else
        {
          WdLogSingleEntry1(3LL, v26);
          WdLogGlobalForLineNumber = 5139;
        }
      }
      v40 = v252;
      *(_DWORD *)&v24[v37 + 144] = *(_DWORD *)(v34 + 276);
      memmove(v40, (const void *)(v34 + 280), 4LL * *(unsigned int *)(v34 + 276));
      v41 = (char *)v252;
      v42 = v254;
      *(_QWORD *)&v24[v37 + 152] = v252;
      v43 = *(unsigned int *)(v34 + 276);
      *(_QWORD *)&v24[v37 + 136] = v42;
      v44 = 0;
      v252 = &v41[4 * v43];
      while ( v44 < *(_DWORD *)(v34 + 8) )
      {
        *(_QWORD *)v42 = *(_QWORD *)(v34 + 24LL * v44 + 16);
        *(_QWORD *)(v42 + 8) = *(_QWORD *)(v34 + 24LL * v44 + 24);
        *(_DWORD *)(v42 + 16) = *(_DWORD *)(v34 + 24LL * v44 + 32);
        if ( *(_QWORD *)(v34 + 24LL * v44 + 16) == -1LL )
          *(_QWORD *)v42 = -1LL;
        if ( *(_QWORD *)(v34 + 24LL * v44 + 24) == -1LL )
          *(_QWORD *)(v42 + 8) = -1LL;
        if ( *(_DWORD *)(v34 + 24LL * v44 + 32) == -1 )
          *(_DWORD *)(v42 + 16) = -1;
        if ( v44 )
        {
          v45 = *(_DWORD *)(v34 + 24LL * v44 + 32);
          if ( v45 != -1 )
          {
            v46 = *(_DWORD *)(v34 + 24 * (v44 - 1 + 1LL) + 8);
            if ( v46 != -1 && v45 > v46 )
            {
              WdLogSingleEntry2(2LL, v26, 5LL);
              WdLogGlobalForLineNumber = 5229;
              DxgkLogInternalTriageEvent(
                0,
                0x40000,
                -1,
                (unsigned int)L"NominalPower must be decreasing for higher F states. Component:0x%I64x",
                v26,
                5LL,
                0LL,
                0LL,
                0LL);
              goto LABEL_57;
            }
          }
          v47 = *(_QWORD *)(v34 + 24LL * v44 + 16);
          if ( v47 != -1LL )
          {
            v48 = *(_QWORD *)(v34 + 24LL * (v44 - 1) + 16);
            if ( v48 != -1LL && v47 < v48 )
            {
              WdLogSingleEntry2(2LL, v26, 6LL);
              WdLogGlobalForLineNumber = 5237;
              DxgkLogInternalTriageEvent(
                0,
                0x40000,
                -1,
                (unsigned int)L"TransitionLatency must be increasing for higher F states. Component:0x%I64x",
                v26,
                6LL,
                0LL,
                0LL,
                0LL);
              goto LABEL_57;
            }
          }
        }
        else
        {
          if ( ((*(_QWORD *)(v34 + 16) + 1LL) & 0xFFFFFFFFFFFFFFFEuLL) != 0
            || ((*(_QWORD *)(v34 + 24) + 1LL) & 0xFFFFFFFFFFFFFFFEuLL) != 0 )
          {
            WdLogSingleEntry2(2LL, v26, 3LL);
            WdLogGlobalForLineNumber = 5212;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"TransitionLatency and ResidencyRequirement must be zero for the F0 state. Component:0x%I64x",
              v26,
              3LL,
              0LL,
              0LL,
              0LL);
            goto LABEL_57;
          }
          if ( !*(_DWORD *)(v34 + 32) )
          {
            WdLogSingleEntry2(2LL, v26, 4LL);
            WdLogGlobalForLineNumber = 5218;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"NominalPower must not be zero for the F0 state. Component:0x%I64x",
              v26,
              4LL,
              0LL,
              0LL,
              0LL);
            goto LABEL_57;
          }
        }
        v42 += 24LL;
        v254 = v42;
        ++v44;
      }
      v11 = *(_DWORD *)(v34 + 208) == 0;
      v25 = v172;
      if ( v11 )
        v25 = ++v172;
      v28 = v165;
      v29 = v173 + 1;
      v27 = v257;
      ++v26;
    }
    NumDifferentPhysicalAdapters = v164;
    v27 = (unsigned int)(v174 + 1);
    v28 += 0x10000;
  }
  if ( *((_DWORD *)this + 844) == -1 && !*((_BYTE *)this + 3664) )
    *((_QWORD *)this + 448) = 0LL;
  if ( *((int *)this + 751) < 1300 || !v25 || v181 )
  {
LABEL_151:
    v84 = (ADAPTER_RENDER *)*((_QWORD *)this + 391);
    *((_DWORD *)this + 914) = v183;
    if ( v84 )
    {
      v85 = ADAPTER_RENDER::InitializePowerManagement(v84);
      v5 = v85;
      if ( v85 < 0 )
      {
        v86 = 7LL;
        WdLogSingleEntry2(2LL, v85, 7LL);
        v87 = L"InitializePowerManagement failed for render adapter:0x%I64x";
        WdLogGlobalForLineNumber = 5446;
LABEL_210:
        DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v87, v5, v86, 0LL, 0LL, 0LL);
        goto LABEL_211;
      }
    }
    v88 = (ADAPTER_DISPLAY *)*((_QWORD *)this + 390);
    if ( v88 )
    {
      v89 = ADAPTER_DISPLAY::InitializePowerManagement(v88);
      v5 = v89;
      if ( v89 < 0 )
      {
        v86 = 8LL;
        WdLogSingleEntry2(2LL, v89, 8LL);
        v87 = L"InitializePowerManagement failed for display adapter:0x%I64x";
        WdLogGlobalForLineNumber = 5456;
        goto LABEL_210;
      }
    }
    v90 = PoFxRegisterDevice(*((_QWORD *)this + 27), v24, (char *)this + 3232);
    v5 = v90;
    if ( v90 < 0 )
    {
      WdLogSingleEntry1(2LL, v90);
      WdLogGlobalForLineNumber = 5464;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"PoFxRegisterDevice failed with status:0x%I64x",
        v5,
        0LL,
        0LL,
        0LL,
        0LL);
      goto LABEL_211;
    }
    KeInitializeEvent((PRKEVENT)((char *)this + 3392), SynchronizationEvent, 0);
    *((_QWORD *)this + 460) = (char *)this + 3672;
    *((_QWORD *)this + 459) = (char *)this + 3672;
    *((_BYTE *)this + 3660) = 0;
    TimeIncrement = KeQueryTimeIncrement();
    v92 = v184;
    v93 = TimeIncrement;
    *((_QWORD *)this + 430) = 0LL;
    *((_QWORD *)this + 432) = 0LL;
    *((_QWORD *)this + 436) = 0LL;
    *((_QWORD *)this + 438) = 0LL;
    *((_QWORD *)this + 427) = 10 * v92;
    v94 = v186;
    *((_QWORD *)this + 429) = 10LL * v185;
    v95 = 10000LL * v187;
    *((_QWORD *)this + 431) = 10 * v94;
    v96 = (unsigned int)(v95 / v93);
    v97 = v188;
    *((_QWORD *)this + 428) = v96;
    *((_QWORD *)this + 434) = v96;
    v98 = 5 * v97;
    v99 = v189;
    *((_QWORD *)this + 433) = 2 * v98;
    v100 = 5 * v99;
    v101 = v190;
    *((_QWORD *)this + 435) = 2 * v100;
    v102 = 5 * v101;
    v103 = v191;
    *((_QWORD *)this + 437) = 2 * v102;
    *((_QWORD *)this + 439) = (char *)this + 3416;
    v104 = 5 * v103;
    v105 = v192;
    *((_QWORD *)this + 471) = 2 * v104;
    v106 = 10000LL * v193;
    *((_QWORD *)this + 440) = 10 * v105;
    v107 = v194;
    *((_QWORD *)this + 441) = (unsigned int)(v106 / v93);
    v108 = 10000LL * v195;
    *((_QWORD *)this + 442) = 10 * v107;
    v109 = v196;
    *((_QWORD *)this + 443) = (unsigned int)(v108 / v93);
    v110 = 5 * v109;
    v111 = 10000LL * v197;
    *((_QWORD *)this + 444) = 2 * v110;
    v112 = v111;
    v113 = v198;
    *((_QWORD *)this + 445) = (unsigned int)(v112 / v93);
    *((_QWORD *)this + 447) = 0LL;
    v163 = 0;
    v114 = 5 * v113;
    v115 = v199;
    *((_QWORD *)this + 446) = 2 * v114;
    v116 = 5 * v115;
    v117 = v200;
    *((_QWORD *)this + 449) = 2 * v116;
    v118 = 5 * v117;
    v119 = v201;
    *((_QWORD *)this + 450) = 2 * v118;
    v120 = 5 * v119;
    v121 = v202;
    *((_QWORD *)this + 451) = 2 * v120;
    v122 = 5 * v121;
    v123 = v203;
    *((_QWORD *)this + 452) = 2 * v122;
    v124 = 5 * v123;
    v125 = v204;
    *((_QWORD *)this + 453) = 2 * v124;
    v126 = 5 * v125;
    v127 = v205;
    *((_QWORD *)this + 454) = 2 * v126;
    v128 = 5 * v127;
    v129 = v206;
    *((_QWORD *)this + 455) = 2 * v128;
    *((_QWORD *)this + 456) = 10 * v129;
    *((_QWORD *)this + 465) = (char *)this + 3712;
    *((_QWORD *)this + 464) = (char *)this + 3712;
    KeInitializeSpinLock((PKSPIN_LOCK)this + 470);
    v130 = 0LL;
    v165 = 0;
    while ( 1 )
    {
      v131 = *((_QWORD *)this + 403);
      v132 = 520 * v130;
      v133 = 520 * v130 + v131;
      *(_BYTE *)(v133 + 356) = 1;
      v134 = (DXGADAPTER *)(v133 + 424);
      *(_OWORD *)(v133 + 424) = 0LL;
      v135 = *(_DWORD *)(v133 + 208);
      if ( !v135 )
      {
        *(_BYTE *)(v133 + 357) = 1;
        v146 = (DXGADAPTER **)*((_QWORD *)this + 469);
        if ( *v146 != (DXGADAPTER *)((char *)this + 3744) )
LABEL_207:
          __fastfail(3u);
        *(_QWORD *)(v133 + 432) = v146;
        *(_QWORD *)v134 = (char *)this + 3744;
        *v146 = v134;
        v147 = 0LL;
        *((_QWORD *)this + 469) = v134;
        v148 = *(_DWORD *)(v133 + 8);
        if ( v148 > 1 )
        {
          v149 = (unsigned __int64 *)(v133 + 40);
          v150 = v148 - 1;
          do
          {
            v151 = *v149;
            v149 += 3;
            if ( v147 >= v151 )
              v151 = v147;
            v147 = v151;
            --v150;
          }
          while ( v150 );
        }
        *(_DWORD *)(v133 + 388) = 1;
        for ( i = 0; ; ++i )
        {
          if ( i >= 2 )
            goto LABEL_190;
          if ( *((_QWORD *)this + 2 * i + 427) >= v147 )
            break;
        }
        *(_DWORD *)(v133 + 388) = i;
LABEL_190:
        v153 = *(_DWORD *)(v133 + 4);
        *(_DWORD *)(v133 + 384) = 2;
        DXGADAPTER::SetPowerComponentLatencyCB(this, v153, *(_QWORD *)(*((_QWORD *)this + 439) + 32LL));
        ++*((_DWORD *)this + 846);
        goto LABEL_191;
      }
      v136 = v135 - 1;
      if ( !v136 )
        break;
      v137 = v136 - 1;
      if ( !v137 )
      {
        v145 = v207;
LABEL_178:
        v142 = *(_DWORD *)(v133 + 4);
        v143 = 10 * v145;
        v141 = this;
LABEL_169:
        DXGADAPTER::SetPowerComponentLatencyCB(v141, v142, v143);
        goto LABEL_191;
      }
      v138 = v137 - 1;
      if ( !v138 )
      {
        v144 = (DXGADAPTER **)*((_QWORD *)this + 467);
        if ( *v144 != (DXGADAPTER *)((char *)this + 3728) )
          goto LABEL_207;
        *(_QWORD *)v134 = (char *)this + 3728;
        *(_QWORD *)(v133 + 432) = v144;
        *v144 = v134;
        *((_QWORD *)this + 467) = v134;
        if ( (*(_DWORD *)(v133 + 216) & 0x10) != 0 )
          *(_BYTE *)(v133 + 360) = 1;
        goto LABEL_191;
      }
      v139 = v138 - 1;
      if ( v139 )
      {
        v140 = v139 - 2;
        if ( v140 )
        {
          if ( v140 == 1 )
          {
            v163 = 1;
            if ( (*(_DWORD *)(v133 + 216) & 0x10) != 0 )
            {
              *(_BYTE *)(v133 + 360) = 1;
              *(_BYTE *)(v133 + 356) = 0;
              *(_DWORD *)(v133 + 344) = 1;
            }
            goto LABEL_191;
          }
          v141 = this;
          v142 = *(_DWORD *)(v133 + 4);
          if ( v175 == -1 )
            v143 = -1LL;
          else
            v143 = 10LL * v175;
          goto LABEL_169;
        }
      }
LABEL_191:
      if ( v176 == -1 )
        v154 = -1LL;
      else
        v154 = 10000LL * v176;
      DXGADAPTER::SetPowerComponentResidencyCB(this, *(_DWORD *)(v133 + 4), v154);
      KeInitializeSpinLock((PKSPIN_LOCK)(v133 + 504));
      if ( *(_DWORD *)(v133 + 8) <= 1u || (v155 = *(_QWORD *)(v133 + 48), v155 == -1) )
      {
        v156 = *((_QWORD *)this + 471);
      }
      else
      {
        v156 = *((_QWORD *)this + 471);
        if ( v155 > v156 )
          v156 = *(_QWORD *)(v133 + 48);
      }
      *(_QWORD *)(v132 + v131 + 496) = v156;
      v130 = v165 + 1;
      v165 = v130;
      if ( (unsigned int)v130 >= v6 )
      {
        DXGADAPTER::UpdateLatencyTolerances(this);
        PoFxSetDeviceIdleTimeout(*((_QWORD *)this + 404), 10LL * v209);
        if ( *((_DWORD *)this + 105) == 1297040209 && *((_DWORD *)this + 684) == 4608 )
        {
          KeInitializeEvent((PRKEVENT)this + 163, SynchronizationEvent, 0);
          KeInitializeEvent((PRKEVENT)this + 164, SynchronizationEvent, 0);
          KeInitializeEvent((PRKEVENT)this + 165, SynchronizationEvent, 0);
          KeInitializeSpinLock((PKSPIN_LOCK)this + 498);
          *((_QWORD *)this + 501) = (char *)this + 4000;
          *((_QWORD *)this + 500) = (char *)this + 4000;
          InitializeSListHead((PSLIST_HEADER)this + 251);
          v86 = 8LL;
          v158 = (struct _SLIST_ENTRY *)((char *)this + 4048);
          v159 = 8LL;
          do
          {
            ExpInterlockedPushEntrySList((PSLIST_HEADER)this + 251, v158);
            v158 += 2;
            --v159;
          }
          while ( v159 );
          *(_QWORD *)&ObjectAttributes.Length = 48LL;
          *(_QWORD *)&ObjectAttributes.Attributes = 512LL;
          ObjectAttributes.RootDirectory = 0LL;
          ObjectAttributes.ObjectName = 0LL;
          *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
          v160 = PsCreateSystemThread(
                   (PHANDLE)this + 504,
                   0x1FFFFFu,
                   &ObjectAttributes,
                   0LL,
                   0LL,
                   DXGADAPTER::PowerRuntimeComponentIdleStateCallbackThread,
                   this);
          v5 = v160;
          if ( v160 < 0 )
          {
            WdLogSingleEntry2(2LL, v160, 8LL);
            v87 = L"InitializePowerManagement failed to create worker thread for display adapter:0x%I64x";
            WdLogGlobalForLineNumber = 5716;
            goto LABEL_210;
          }
        }
        LOBYTE(v157) = v163;
        v161 = DpiEnablePowerManagement(*((_QWORD *)this + 27), *((_QWORD *)this + 404), v157);
        v5 = v161;
        if ( v161 < 0 )
        {
          DXGADAPTER::DestroySerializeFStateTransitWorker(this);
          v86 = 9LL;
          WdLogSingleEntry2(2LL, v5, 9LL);
          v87 = L"Port power management enable failed:0x%I64x";
          WdLogGlobalForLineNumber = 5731;
          goto LABEL_210;
        }
        DXGQUOTAALLOCATOR<256,1835156294>::operator delete(v24);
        return 0LL;
      }
    }
    v145 = v208;
    goto LABEL_178;
  }
  if ( v170 > 0x3E8 || v169 > 0x3E8 || v168 > 0x3E8 || v167 > 0x3E8 || v170 >= v167 || v167 >= v169 || v169 >= v168 )
  {
    WdLogSingleEntry4(2LL, v170, v169, v168, v167);
    WdLogGlobalForLineNumber = 5286;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"P-State engine regkey validation error - low: 0x%I64x high: 0x%I64x glitch: 0x%I64x ideal: 0x%I64x",
      v170,
      v169,
      v168,
      v167,
      0LL);
    goto LABEL_57;
  }
  v50 = 248LL * v25;
  v255.Type = DXGKQAITYPE_POWERCOMPONENTPSTATEINFO;
  v255.OutputDataSize = 136;
  if ( !is_mul_ok(v25, 0xF8uLL) )
    v50 = -1LL;
  v51 = operator new[](v50, 1265072196LL, 64LL);
  *((_QWORD *)this + 553) = v51;
  *((_DWORD *)this + 1108) = v25;
  if ( !v51 )
  {
    WdLogSingleEntry1(6LL, this);
    WdLogGlobalForLineNumber = 5302;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Adapter 0x%I64x: Out of memory allocating m_NodePStateData",
      (__int64)this,
      0LL,
      0LL,
      0LL,
      0LL);
    LODWORD(v5) = -1073741801;
    goto LABEL_211;
  }
  v52 = 0;
  for ( j = 0; v52 < *((_DWORD *)this + 842); ++v52 )
  {
    v54 = *((_QWORD *)this + 403);
    v55 = 520LL * v52;
    if ( !*(_DWORD *)(v55 + v54 + 208) )
      *(_QWORD *)(v55 + v54 + 512) = *((_QWORD *)this + 553) + 248LL * j++;
  }
  v56 = 0;
  *((_DWORD *)this + 1160) = v168;
  v57 = 0;
  *((_DWORD *)this + 1161) = v169;
  *((_DWORD *)this + 1162) = v170;
  *((_DWORD *)this + 1163) = v167;
  *((_DWORD *)this + 1164) = v182;
  v58 = 0;
  while ( 1 )
  {
    v164 = v56;
    if ( v58 >= v6 )
      break;
    v59 = *(_QWORD *)(520LL * v57 + *((_QWORD *)this + 403) + 512);
    if ( v59 )
    {
      v255.pOutputData = *(void **)(520LL * v57 + *((_QWORD *)this + 403) + 512);
      v255.pInputData = &v164;
      v60 = DXGADAPTER::DdiQueryAdapterInfo(this, &v255);
      v62 = v60;
      if ( v60 < 0 )
      {
        v64 = WdLogNewEntry5_WdTrace(v61);
        *(_QWORD *)(v64 + 24) = v164;
        v65 = 0;
        *(_QWORD *)(v64 + 32) = v62;
        for ( WdLogGlobalForLineNumber = 5352; v65 < *((_DWORD *)this + 842); ++v65 )
        {
          v66 = 520LL * v65;
          v67 = *((_QWORD *)this + 403);
          if ( !*(_DWORD *)(v66 + v67 + 208) )
            *(_QWORD *)(v66 + v67 + 512) = 0LL;
        }
        v68 = (void *)*((_QWORD *)this + 553);
        *((_DWORD *)this + 1108) = 0;
        DXGQUOTAALLOCATOR<256,1835156294>::operator delete(v68);
        *((_QWORD *)this + 553) = 0LL;
        break;
      }
      v63 = v164;
      *(_QWORD *)(v59 + 136) = this;
      *(_DWORD *)(v59 + 144) = v63;
      *(_QWORD *)(v59 + 152) = v59;
      KeInitializeSpinLock((PKSPIN_LOCK)(v59 + 160));
      *(_DWORD *)(v59 + 244) = -1;
      *(_BYTE *)(v59 + 240) = 0;
      v56 = v164;
    }
    v58 = ++v56;
    v57 = v56;
  }
  v69 = *((_DWORD *)this + 1108);
  v70 = 0;
  v165 = v69;
LABEL_138:
  if ( v70 >= v69 )
    goto LABEL_151;
  v71 = *((_QWORD *)this + 553);
  v72 = v70;
  v73 = 248LL * v70;
  v74 = *(_DWORD *)(v73 + v71);
  v75 = *(unsigned int *)(v73 + v71 + 144);
  if ( v74 > 0x20 )
  {
    v83 = *(unsigned int *)(v73 + v71 + 144);
    WdLogSingleEntry1(2LL, v83);
    WdLogGlobalForLineNumber = 5407;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"P-State StateCount cannot be larger than DXGK_MAX_P_STATES. Component:0x%I64x",
      v83,
      0LL,
      0LL,
      0LL,
      0LL);
    goto LABEL_57;
  }
  for ( k = 0; ; ++k )
  {
    if ( k >= v74 )
    {
      v69 = v165;
      ++v70;
      goto LABEL_138;
    }
    v77 = 62 * v72;
    v78 = k;
    v79 = *(_DWORD *)(v71 + 4 * (k + v77) + 4);
    if ( !v79 )
    {
      v82 = *(unsigned int *)(v73 + v71 + 144);
      WdLogSingleEntry2(2LL, v75, k);
      WdLogGlobalForLineNumber = 5420;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"P-State cannot specify 0 operating frequency. Component:0x%I64x, P-State:0x%I64x",
        v82,
        v78,
        0LL,
        0LL,
        0LL);
LABEL_57:
      LODWORD(v5) = -1073741811;
      goto LABEL_211;
    }
    if ( k )
    {
      v80 = k - 1;
      if ( v79 > *(_DWORD *)(v71 + 4 * (v77 + v80) + 4) )
        break;
    }
    v72 = v70;
  }
  v81 = *(unsigned int *)(v73 + v71 + 144);
  WdLogSingleEntry3(2LL, v75, k, k - 1, ClientId);
  WdLogGlobalForLineNumber = 5430;
  DxgkLogInternalTriageEvent(
    0,
    0x40000,
    -1,
    (unsigned int)L"P-States must have monotonically decreasing operating frequency. Component:0x%I64x, P-State1:0x%I64x, "
                   "P-State2:0x%I64x",
    v81,
    v78,
    v80,
    0LL,
    0LL);
LABEL_103:
  LODWORD(v5) = -1073741811;
LABEL_211:
  DXGQUOTAALLOCATOR<256,1835156294>::operator delete(v24);
LABEL_212:
  if ( *((_QWORD *)this + 404) )
  {
    PoFxUnregisterDevice();
    *((_QWORD *)this + 404) = 0LL;
  }
  return (unsigned int)v5;
}

__int64 __fastcall DXGGLOBAL::Initialize(DXGGLOBAL *this)
{
  __int64 v1; // rdi
  __int128 v2; // xmm1
  __int128 v3; // xmm0
  __int128 v4; // xmm1
  __int128 v5; // xmm0
  NTSTATUS v6; // eax
  __int64 v7; // r14
  const wchar_t *v8; // r9
  NTSTATUS v10; // eax
  struct _ERESOURCE *v11; // rax
  NTSTATUS v12; // eax
  unsigned int v13; // ebx
  NTSTATUS v14; // eax
  NTSTATUS v15; // eax
  unsigned __int8 v16; // r9
  int v17; // ecx
  int v18; // r8d
  int v19; // eax
  int v20; // eax
  int v21; // eax
  __int128 v22; // xmm1
  __int128 v23; // xmm0
  __int128 v24; // xmm1
  __int128 v25; // xmm0
  __int128 v26; // xmm1
  __int128 v27; // xmm0
  int v28; // eax
  __int64 v29; // rcx
  int DxgkSharedObjectTypes; // eax
  unsigned int v31; // ecx
  unsigned int v32; // edx
  unsigned __int64 v33; // rbx
  __int64 v34; // rax
  __int64 v35; // rax
  __int64 v36; // rax
  __int64 v37; // rax
  DXGSESSIONMGR *v38; // rax
  DXGSESSIONMGR *v39; // rax
  int v40; // ecx
  __int64 v41; // rbx
  __int64 v42; // rax
  ULONG *v43; // rax
  __int64 v44; // rax
  _BYTE *v45; // rbx
  NTSTATUS v46; // eax
  __int64 v47; // rbx
  NTSTATUS v48; // eax
  NTSTATUS v49; // eax
  __int64 v50; // rdi
  ULONG Flags[2]; // [rsp+28h] [rbp-E0h]
  ULONG Flagsa[2]; // [rsp+28h] [rbp-E0h]
  ULONG Flagsb[2]; // [rsp+28h] [rbp-E0h]
  int OutputBuffer; // [rsp+58h] [rbp-B0h] BYREF
  unsigned int v55; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v56; // [rsp+60h] [rbp-A8h] BYREF
  unsigned int v57; // [rsp+64h] [rbp-A4h] BYREF
  unsigned int v58; // [rsp+68h] [rbp-A0h] BYREF
  unsigned int v59; // [rsp+6Ch] [rbp-9Ch] BYREF
  unsigned int v60; // [rsp+70h] [rbp-98h] BYREF
  unsigned int v61; // [rsp+74h] [rbp-94h] BYREF
  unsigned int v62; // [rsp+78h] [rbp-90h] BYREF
  int v63; // [rsp+7Ch] [rbp-8Ch] BYREF
  int v64; // [rsp+80h] [rbp-88h] BYREF
  int v65; // [rsp+84h] [rbp-84h] BYREF
  int v66; // [rsp+88h] [rbp-80h] BYREF
  int v67; // [rsp+8Ch] [rbp-7Ch] BYREF
  int v68; // [rsp+90h] [rbp-78h] BYREF
  int v69; // [rsp+94h] [rbp-74h] BYREF
  int v70; // [rsp+98h] [rbp-70h] BYREF
  int v71; // [rsp+9Ch] [rbp-6Ch] BYREF
  int v72; // [rsp+A0h] [rbp-68h] BYREF
  int v73; // [rsp+A4h] [rbp-64h] BYREF
  unsigned int v74; // [rsp+A8h] [rbp-60h] BYREF
  unsigned int v75; // [rsp+ACh] [rbp-5Ch] BYREF
  int v76; // [rsp+B0h] [rbp-58h] BYREF
  int v77; // [rsp+B4h] [rbp-54h] BYREF
  int v78; // [rsp+B8h] [rbp-50h] BYREF
  int v79; // [rsp+BCh] [rbp-4Ch] BYREF
  int v80; // [rsp+C0h] [rbp-48h] BYREF
  int v81; // [rsp+C4h] [rbp-44h] BYREF
  int v82; // [rsp+C8h] [rbp-40h] BYREF
  int v83; // [rsp+CCh] [rbp-3Ch] BYREF
  int v84; // [rsp+D0h] [rbp-38h] BYREF
  int v85; // [rsp+D4h] [rbp-34h] BYREF
  int v86; // [rsp+D8h] [rbp-30h] BYREF
  int v87; // [rsp+DCh] [rbp-2Ch] BYREF
  int v88; // [rsp+E0h] [rbp-28h] BYREF
  int v89; // [rsp+E4h] [rbp-24h] BYREF
  int v90; // [rsp+E8h] [rbp-20h] BYREF
  struct _UNICODE_STRING v91; // [rsp+F0h] [rbp-18h] BYREF
  struct _UNICODE_STRING v92; // [rsp+100h] [rbp-8h] BYREF
  _QWORD v93[13]; // [rsp+110h] [rbp+8h] BYREF
  __int64 v94; // [rsp+178h] [rbp+70h] BYREF
  int v95; // [rsp+180h] [rbp+78h]
  const wchar_t *v96; // [rsp+188h] [rbp+80h]
  unsigned int *v97; // [rsp+190h] [rbp+88h]
  int v98; // [rsp+198h] [rbp+90h]
  _QWORD *v99; // [rsp+1A0h] [rbp+98h]
  int v100; // [rsp+1A8h] [rbp+A0h]
  __int64 v101; // [rsp+1B0h] [rbp+A8h]
  int v102; // [rsp+1B8h] [rbp+B0h]
  const wchar_t *v103; // [rsp+1C0h] [rbp+B8h]
  int *v104; // [rsp+1C8h] [rbp+C0h]
  int v105; // [rsp+1D0h] [rbp+C8h]
  int *v106; // [rsp+1D8h] [rbp+D0h]
  int v107; // [rsp+1E0h] [rbp+D8h]
  __int64 v108; // [rsp+1E8h] [rbp+E0h]
  int v109; // [rsp+1F0h] [rbp+E8h]
  const wchar_t *v110; // [rsp+1F8h] [rbp+F0h]
  unsigned int *v111; // [rsp+200h] [rbp+F8h]
  int v112; // [rsp+208h] [rbp+100h]
  int *v113; // [rsp+210h] [rbp+108h]
  int v114; // [rsp+218h] [rbp+110h]
  __int64 v115; // [rsp+220h] [rbp+118h]
  int v116; // [rsp+228h] [rbp+120h]
  const wchar_t *v117; // [rsp+230h] [rbp+128h]
  unsigned int *v118; // [rsp+238h] [rbp+130h]
  int v119; // [rsp+240h] [rbp+138h]
  int *v120; // [rsp+248h] [rbp+140h]
  int v121; // [rsp+250h] [rbp+148h]
  __int64 v122; // [rsp+258h] [rbp+150h]
  int v123; // [rsp+260h] [rbp+158h]
  const wchar_t *v124; // [rsp+268h] [rbp+160h]
  int *v125; // [rsp+270h] [rbp+168h]
  int v126; // [rsp+278h] [rbp+170h]
  int *v127; // [rsp+280h] [rbp+178h]
  int v128; // [rsp+288h] [rbp+180h]
  __int64 v129; // [rsp+290h] [rbp+188h]
  int v130; // [rsp+298h] [rbp+190h]
  const wchar_t *v131; // [rsp+2A0h] [rbp+198h]
  int *v132; // [rsp+2A8h] [rbp+1A0h]
  int v133; // [rsp+2B0h] [rbp+1A8h]
  int *v134; // [rsp+2B8h] [rbp+1B0h]
  int v135; // [rsp+2C0h] [rbp+1B8h]
  __int64 v136; // [rsp+2C8h] [rbp+1C0h]
  int v137; // [rsp+2D0h] [rbp+1C8h]
  const wchar_t *v138; // [rsp+2D8h] [rbp+1D0h]
  int *v139; // [rsp+2E0h] [rbp+1D8h]
  int v140; // [rsp+2E8h] [rbp+1E0h]
  int *v141; // [rsp+2F0h] [rbp+1E8h]
  int v142; // [rsp+2F8h] [rbp+1F0h]
  __int64 v143; // [rsp+300h] [rbp+1F8h]
  int v144; // [rsp+308h] [rbp+200h]
  const wchar_t *v145; // [rsp+310h] [rbp+208h]
  int *v146; // [rsp+318h] [rbp+210h]
  int v147; // [rsp+320h] [rbp+218h]
  int *v148; // [rsp+328h] [rbp+220h]
  int v149; // [rsp+330h] [rbp+228h]
  __int64 v150; // [rsp+338h] [rbp+230h]
  int v151; // [rsp+340h] [rbp+238h]
  const wchar_t *v152; // [rsp+348h] [rbp+240h]
  int *v153; // [rsp+350h] [rbp+248h]
  int v154; // [rsp+358h] [rbp+250h]
  int *v155; // [rsp+360h] [rbp+258h]
  int v156; // [rsp+368h] [rbp+260h]
  __int64 v157; // [rsp+370h] [rbp+268h]
  int v158; // [rsp+378h] [rbp+270h]
  const wchar_t *v159; // [rsp+380h] [rbp+278h]
  int *v160; // [rsp+388h] [rbp+280h]
  int v161; // [rsp+390h] [rbp+288h]
  int *v162; // [rsp+398h] [rbp+290h]
  int v163; // [rsp+3A0h] [rbp+298h]
  __int64 v164; // [rsp+3A8h] [rbp+2A0h]
  int v165; // [rsp+3B0h] [rbp+2A8h]
  const wchar_t *v166; // [rsp+3B8h] [rbp+2B0h]
  unsigned int *v167; // [rsp+3C0h] [rbp+2B8h]
  int v168; // [rsp+3C8h] [rbp+2C0h]
  int *v169; // [rsp+3D0h] [rbp+2C8h]
  int v170; // [rsp+3D8h] [rbp+2D0h]
  __int64 v171; // [rsp+3E0h] [rbp+2D8h]
  int v172; // [rsp+3E8h] [rbp+2E0h]
  const wchar_t *v173; // [rsp+3F0h] [rbp+2E8h]
  unsigned int *v174; // [rsp+3F8h] [rbp+2F0h]
  int v175; // [rsp+400h] [rbp+2F8h]
  unsigned int *v176; // [rsp+408h] [rbp+300h]
  int v177; // [rsp+410h] [rbp+308h]
  __int64 v178; // [rsp+418h] [rbp+310h]
  int v179; // [rsp+420h] [rbp+318h]
  const wchar_t *v180; // [rsp+428h] [rbp+320h]
  unsigned int *v181; // [rsp+430h] [rbp+328h]
  int v182; // [rsp+438h] [rbp+330h]
  int *v183; // [rsp+440h] [rbp+338h]
  int v184; // [rsp+448h] [rbp+340h]
  __int64 v185; // [rsp+450h] [rbp+348h]
  int v186; // [rsp+458h] [rbp+350h]
  const wchar_t *v187; // [rsp+460h] [rbp+358h]
  unsigned int *v188; // [rsp+468h] [rbp+360h]
  int v189; // [rsp+470h] [rbp+368h]
  int *v190; // [rsp+478h] [rbp+370h]
  int v191; // [rsp+480h] [rbp+378h]
  __int64 v192; // [rsp+488h] [rbp+380h]
  int v193; // [rsp+490h] [rbp+388h]
  const wchar_t *v194; // [rsp+498h] [rbp+390h]
  unsigned int *v195; // [rsp+4A0h] [rbp+398h]
  int v196; // [rsp+4A8h] [rbp+3A0h]
  int *v197; // [rsp+4B0h] [rbp+3A8h]
  int v198; // [rsp+4B8h] [rbp+3B0h]
  __int64 v199; // [rsp+4C0h] [rbp+3B8h]
  int v200; // [rsp+4C8h] [rbp+3C0h]
  const wchar_t *v201; // [rsp+4D0h] [rbp+3C8h]
  int *v202; // [rsp+4D8h] [rbp+3D0h]
  int v203; // [rsp+4E0h] [rbp+3D8h]
  int *v204; // [rsp+4E8h] [rbp+3E0h]
  int v205; // [rsp+4F0h] [rbp+3E8h]
  __int64 v206; // [rsp+4F8h] [rbp+3F0h]
  int v207; // [rsp+500h] [rbp+3F8h]
  const wchar_t *v208; // [rsp+508h] [rbp+400h]
  int *v209; // [rsp+510h] [rbp+408h]
  int v210; // [rsp+518h] [rbp+410h]
  int *v211; // [rsp+520h] [rbp+418h]
  int v212; // [rsp+528h] [rbp+420h]
  __int64 v213; // [rsp+530h] [rbp+428h]
  int v214; // [rsp+538h] [rbp+430h]
  const wchar_t *v215; // [rsp+540h] [rbp+438h]
  unsigned int *v216; // [rsp+548h] [rbp+440h]
  int v217; // [rsp+550h] [rbp+448h]
  __int64 v218; // [rsp+558h] [rbp+450h]
  int v219; // [rsp+560h] [rbp+458h]
  __int64 v220; // [rsp+568h] [rbp+460h]
  int v221; // [rsp+570h] [rbp+468h]
  const wchar_t *v222; // [rsp+578h] [rbp+470h]
  unsigned int *v223; // [rsp+580h] [rbp+478h]
  int v224; // [rsp+588h] [rbp+480h]
  __int64 v225; // [rsp+590h] [rbp+488h]
  int v226; // [rsp+598h] [rbp+490h]
  __int64 v227; // [rsp+5A0h] [rbp+498h]
  int v228; // [rsp+5A8h] [rbp+4A0h]
  const wchar_t *v229; // [rsp+5B0h] [rbp+4A8h]
  unsigned int *v230; // [rsp+5B8h] [rbp+4B0h]
  int v231; // [rsp+5C0h] [rbp+4B8h]
  __int64 v232; // [rsp+5C8h] [rbp+4C0h]
  int v233; // [rsp+5D0h] [rbp+4C8h]
  __int64 v234; // [rsp+5D8h] [rbp+4D0h]
  int v235; // [rsp+5E0h] [rbp+4D8h]
  const wchar_t *v236; // [rsp+5E8h] [rbp+4E0h]
  unsigned int *v237; // [rsp+5F0h] [rbp+4E8h]
  int v238; // [rsp+5F8h] [rbp+4F0h]
  __int64 v239; // [rsp+600h] [rbp+4F8h]
  int v240; // [rsp+608h] [rbp+500h]
  __int64 v241; // [rsp+610h] [rbp+508h]
  int v242; // [rsp+618h] [rbp+510h]
  const wchar_t *v243; // [rsp+620h] [rbp+518h]
  unsigned int *v244; // [rsp+628h] [rbp+520h]
  int v245; // [rsp+630h] [rbp+528h]
  __int64 v246; // [rsp+638h] [rbp+530h]
  int v247; // [rsp+640h] [rbp+538h]
  __int64 v248; // [rsp+648h] [rbp+540h]
  int v249; // [rsp+650h] [rbp+548h]
  const wchar_t *v250; // [rsp+658h] [rbp+550h]
  unsigned int *v251; // [rsp+660h] [rbp+558h]
  int v252; // [rsp+668h] [rbp+560h]
  __int64 v253; // [rsp+670h] [rbp+568h]
  int v254; // [rsp+678h] [rbp+570h]
  __int64 v255; // [rsp+680h] [rbp+578h]
  int v256; // [rsp+688h] [rbp+580h]
  const wchar_t *v257; // [rsp+690h] [rbp+588h]
  int *v258; // [rsp+698h] [rbp+590h]
  int v259; // [rsp+6A0h] [rbp+598h]
  __int64 v260; // [rsp+6A8h] [rbp+5A0h]
  int v261; // [rsp+6B0h] [rbp+5A8h]
  __int64 v262; // [rsp+6B8h] [rbp+5B0h]
  int v263; // [rsp+6C0h] [rbp+5B8h]
  const wchar_t *v264; // [rsp+6C8h] [rbp+5C0h]
  int *v265; // [rsp+6D0h] [rbp+5C8h]
  int v266; // [rsp+6D8h] [rbp+5D0h]
  __int64 v267; // [rsp+6E0h] [rbp+5D8h]
  int v268; // [rsp+6E8h] [rbp+5E0h]
  __int64 v269; // [rsp+6F0h] [rbp+5E8h]
  int v270; // [rsp+6F8h] [rbp+5F0h]
  __int64 v271; // [rsp+700h] [rbp+5F8h]
  __int128 v272; // [rsp+708h] [rbp+600h]
  __int128 v273; // [rsp+718h] [rbp+610h]
  _OWORD v274[2]; // [rsp+728h] [rbp+620h] BYREF
  wchar_t v275; // [rsp+748h] [rbp+640h]
  _OWORD v276[9]; // [rsp+758h] [rbp+650h] BYREF
  int v277; // [rsp+7E8h] [rbp+6E0h]
  wchar_t v278; // [rsp+7ECh] [rbp+6E4h]

  v1 = *(_QWORD *)&DXGGLOBAL::m_pGlobal;
  memset(&v93[1], 0, 0x58uLL);
  v2 = *(_OWORD *)&v93[3];
  *(_OWORD *)(*(_QWORD *)&DXGGLOBAL::m_pGlobal + 64LL) = *(_OWORD *)&v93[1];
  v3 = *(_OWORD *)&v93[5];
  *(_OWORD *)(v1 + 80) = v2;
  v4 = *(_OWORD *)&v93[7];
  *(_OWORD *)(v1 + 96) = v3;
  v5 = *(_OWORD *)&v93[9];
  *(_OWORD *)(v1 + 112) = v4;
  *(_QWORD *)&v4 = v93[11];
  *(_OWORD *)(v1 + 128) = v5;
  *(_QWORD *)(v1 + 144) = v4;
  g_WindowsSubsystem = ZwAllocateVirtualMemory;
  qword_1C01538D0 = ZwAllocateVirtualMemoryEx;
  qword_1C01538D8 = (__int64)ZwFreeVirtualMemory;
  qword_1C01538E0 = MmMapViewOfSection;
  qword_1C01538E8 = MmUnmapViewOfSection;
  qword_1C01538F0 = (__int64)MmMapLockedPagesSpecifyCache;
  qword_1C01538F8 = (__int64)MmUnmapLockedPages;
  g_WslSubsystem = ZwAllocateVirtualMemory;
  qword_1C0153898 = ZwAllocateVirtualMemoryEx;
  qword_1C01538A0 = (__int64)ZwFreeVirtualMemory;
  qword_1C01538A8 = MmMapViewOfSection;
  qword_1C01538B0 = MmUnmapViewOfSection;
  qword_1C01538B8 = (__int64)MmMapLockedPagesSpecifyCache;
  qword_1C01538C0 = (__int64)MmUnmapLockedPages;
  v6 = ExInitializeLookasideListEx(
         (PLOOKASIDE_LIST_EX)(v1 + 305392),
         0LL,
         0LL,
         (POOL_TYPE)512,
         0,
         0x10uLL,
         0x4B677844u,
         0);
  v7 = v6;
  if ( v6 < 0 )
  {
    WdLogSingleEntry2(2LL, v1, v6);
    v8 = L"DXGGlobal 0x%I64x: Unable to initialize the lookaside list for lock order tracker, returning 0x%I64x";
    WdLogGlobalForLineNumber = 1800;
LABEL_3:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v8, v1, v7, 0LL, 0LL, 0LL);
    return (unsigned int)v7;
  }
  *(_BYTE *)(v1 + 305376) = 1;
  v10 = ExInitializeLookasideListEx(
          (PLOOKASIDE_LIST_EX)(v1 + 160),
          0LL,
          0LL,
          (POOL_TYPE)512,
          0,
          0xA0uLL,
          0x576B7844u,
          0);
  v7 = v10;
  if ( v10 < 0 )
  {
    WdLogSingleEntry2(2LL, v1, v10);
    v8 = L"DXGGlobal 0x%I64x: Unable to initialize m_VmBusPacketWorkItemList, returning 0x%I64x";
    WdLogGlobalForLineNumber = 1812;
    goto LABEL_3;
  }
  *(_BYTE *)(v1 + 1347) = 1;
  if ( !HMGRTABLE::ExpandTable((HMGRTABLE *)(v1 + 336)) )
  {
    WdLogSingleEntry1(6LL, -1073741801LL);
    WdLogGlobalForLineNumber = 1824;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed the initial shared resource handle table expansion, returning 0x%I64x",
      -1073741801LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v11 = (struct _ERESOURCE *)operator new(104LL, 1265072196LL);
  *(_QWORD *)(v1 + 600) = v11;
  if ( !v11 )
  {
    WdLogSingleEntry2(3LL, v1, -1073741801LL);
    WdLogGlobalForLineNumber = 1837;
    return 3221225495LL;
  }
  v12 = ExInitializeResourceLite(v11);
  v13 = v12;
  if ( v12 < 0 )
  {
    WdLogSingleEntry2(3LL, v1, v12);
    WdLogGlobalForLineNumber = 1847;
    return v13;
  }
  v14 = ExInitializeLookasideListEx((PLOOKASIDE_LIST_EX)(v1 + 1136), 0LL, 0LL, PagedPool, 0, 0x5F8uLL, 0x4B677844u, 0);
  v13 = v14;
  if ( v14 < 0 )
  {
    WdLogSingleEntry3(3LL, v1, v14, 0LL, *(_QWORD *)Flags);
    WdLogGlobalForLineNumber = 1856;
    return v13;
  }
  *(_BYTE *)(v1 + 1345) = 1;
  v15 = ExInitializeLookasideListEx((PLOOKASIDE_LIST_EX)(v1 + 1232), 0LL, 0LL, PagedPool, 0, 0x5E0uLL, 0x4B677844u, 0);
  v13 = v15;
  if ( v15 < 0 )
  {
    WdLogSingleEntry3(3LL, v1, v15, 0LL, *(_QWORD *)Flagsa);
    WdLogGlobalForLineNumber = 1866;
    return v13;
  }
  v16 = g_bSkuSupportMultipleUsers;
  *(_BYTE *)(v1 + 1346) = 1;
  v79 = 32;
  v93[0] = 0x4000000LL;
  v62 = 0;
  v77 = 0;
  v63 = 0;
  v78 = 1;
  v61 = 0;
  v60 = 0;
  v65 = 0;
  v80 = 0;
  v81 = 0;
  v66 = 0;
  v67 = 0;
  v82 = 0;
  v83 = 0;
  v68 = 0;
  v84 = 0;
  v69 = 0;
  v85 = 0;
  v64 = 0;
  v72 = 0;
  if ( v16 )
    v17 = g_IsInternalReleaseOrDbg != 0 ? 0x100000 : 0x80000;
  else
    v17 = g_IsInternalReleaseOrDbg != 0 ? 0x80000 : 0x10000;
  v86 = v17;
  if ( v16 )
    v18 = g_IsInternalReleaseOrDbg != 0 ? 8 : 4;
  else
    v18 = 2;
  v75 = v18;
  if ( v16 )
    v19 = g_IsInternalReleaseOrDbg != 0 ? 0x80000 : 0x10000;
  else
    v19 = g_IsInternalReleaseOrDbg != 0 ? 0x80000 : 0x4000;
  v87 = v19;
  v56 = v19;
  v88 = 300;
  v58 = 300;
  v55 = v17;
  v76 = 1;
  v57 = v18;
  v59 = 1;
  v89 = 5000;
  v70 = 0;
  v90 = 15000;
  v71 = 0;
  v73 = *(_DWORD *)(v1 + 305880);
  v96 = L"TerminationListSizeLimit";
  v97 = &v62;
  v99 = v93;
  v103 = L"ValidateWDDMCaps";
  v104 = &v63;
  v106 = &v77;
  v110 = L"WDDM2LockManagement";
  v111 = &v61;
  v113 = &v78;
  v117 = L"MaximumAdapterCount";
  v118 = &v60;
  v120 = &v79;
  v124 = L"InvestigationDebugParameter";
  v125 = &v65;
  v127 = &v80;
  v131 = L"EnableIgnoreWin32ProcessStatus";
  v132 = &v66;
  v134 = &v81;
  v138 = L"EnableHMDTestMode";
  v94 = 0LL;
  v95 = 288;
  v98 = 67108868;
  v100 = 4;
  v101 = 0LL;
  v102 = 288;
  v105 = 67108868;
  v107 = 4;
  v108 = 0LL;
  v109 = 288;
  v112 = 67108868;
  v114 = 4;
  v115 = 0LL;
  v116 = 288;
  v119 = 67108868;
  v121 = 4;
  v122 = 0LL;
  v123 = 288;
  v126 = 67108868;
  v128 = 4;
  v129 = 0LL;
  v130 = 288;
  v133 = 67108868;
  v135 = 4;
  v136 = 0LL;
  v137 = 288;
  v140 = 67108868;
  v139 = &v67;
  v141 = &v82;
  v145 = L"PreserveFirmwareMode";
  v146 = &v68;
  v148 = &v83;
  v152 = L"PreventFullscreenWireFormatChange";
  v153 = &v69;
  v155 = &v84;
  v159 = L"EnableFuzzing";
  v160 = &v64;
  v162 = &v85;
  v166 = L"InternalDiagnosticsBufferSize";
  v167 = &v55;
  v169 = &v86;
  v173 = L"InternalDiagnosticsBufferMultiplier";
  v174 = &v57;
  v176 = &v75;
  v180 = L"ExternalDiagnosticsBufferSize";
  v181 = &v56;
  v183 = &v87;
  v187 = L"ExternalDiagnosticsBufferMultiplier";
  v188 = &v59;
  v190 = &v76;
  v194 = L"DiagnosticsBufferExpansionTime";
  v142 = 4;
  v143 = 0LL;
  v144 = 288;
  v147 = 67108868;
  v149 = 4;
  v150 = 0LL;
  v151 = 288;
  v154 = 67108868;
  v156 = 4;
  v157 = 0LL;
  v158 = 288;
  v161 = 67108868;
  v163 = 4;
  v164 = 0LL;
  v165 = 288;
  v168 = 67108868;
  v170 = 4;
  v171 = 0LL;
  v172 = 288;
  v175 = 67108868;
  v177 = 4;
  v178 = 0LL;
  v179 = 288;
  v182 = 67108868;
  v184 = 4;
  v185 = 0LL;
  v186 = 288;
  v189 = 67108868;
  v191 = 4;
  v192 = 0LL;
  v193 = 288;
  v195 = &v58;
  v197 = &v88;
  v201 = L"RapidHpdTimeoutInMilliseconds";
  v202 = &v70;
  v204 = &v89;
  v208 = L"RapidHpdMaxChainInMilliseconds";
  v209 = &v71;
  v211 = &v90;
  v215 = L"ForceUsb4MonitorSupport";
  v216 = &g_bDbgForceUsb4MonitorSupport;
  v222 = L"Usb4MonitorTargetId";
  v223 = &g_DbgUsb4MonitorTargetId;
  v229 = L"Usb4MonitorDpcdUSB4_Driver_ID";
  v230 = &g_DbgUsb4MonitorDpcdUSB4_Driver_ID;
  v236 = L"Usb4MonitorDpcdDP_IN_Adapter_Number";
  v237 = &g_DbgUsb4MonitorDpcdDP_IN_Adapter_Number;
  v243 = L"Usb4MonitorPowerOnDelayInSeconds";
  v244 = &g_DbgUsb4MonitorPowerOnDelayInSeconds;
  v250 = L"TreatUsb4MonitorAsNormal";
  v251 = &g_bDbgTreatUsb4MonitorAsNormal;
  v196 = 67108868;
  v198 = 4;
  v199 = 0LL;
  v200 = 288;
  v203 = 67108868;
  v205 = 4;
  v206 = 0LL;
  v207 = 288;
  v210 = 67108868;
  v212 = 4;
  v213 = 0LL;
  v214 = 288;
  v217 = 67108868;
  v218 = 0LL;
  v219 = 0;
  v220 = 0LL;
  v221 = 288;
  v224 = 67108868;
  v225 = 0LL;
  v226 = 0;
  v227 = 0LL;
  v228 = 288;
  v231 = 67108868;
  v232 = 0LL;
  v233 = 0;
  v234 = 0LL;
  v235 = 288;
  v238 = 67108868;
  v239 = 0LL;
  v240 = 0;
  v241 = 0LL;
  v242 = 288;
  v245 = 67108868;
  v246 = 0LL;
  v247 = 0;
  v248 = 0LL;
  v249 = 288;
  v252 = 67108868;
  v253 = 0LL;
  v254 = 0;
  v255 = 0LL;
  v256 = 288;
  v259 = 67108868;
  v263 = 288;
  v257 = L"AllowAdvancedEtwLogging";
  v266 = 67108868;
  v258 = &v72;
  v260 = 0LL;
  v261 = 0;
  v264 = L"NodeUsageTelemetryTimerInterval";
  v265 = &v73;
  v262 = 0LL;
  v267 = 0LL;
  v268 = 0;
  v269 = 0LL;
  v270 = 0;
  v271 = 0LL;
  v272 = 0LL;
  Flagsb[1] = 0;
  v273 = 0LL;
  if ( (int)RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v94) < 0 )
  {
    *(_QWORD *)(v1 + 880) = 0x4000000LL;
    *(_DWORD *)(v1 + 1364) = 32;
    *(_BYTE *)(v1 + 888) = 0;
    *(_DWORD *)(v1 + 1360) = 1;
    *(_DWORD *)(v1 + 1648) = 0;
    *(_DWORD *)(v1 + 1664) = 0;
  }
  else
  {
    *(_QWORD *)(v1 + 880) = v62;
    *(_BYTE *)(v1 + 888) = v63 != 0;
    *(_BYTE *)(v1 + 304834) = v64 != 0;
    v20 = 1;
    if ( v61 < 2 )
      v20 = v61;
    *(_DWORD *)(v1 + 1360) = v20;
    v21 = v60;
    if ( v60 >= 4 )
    {
      if ( v60 > 0x400 )
      {
        v21 = 1024;
        v60 = 1024;
      }
    }
    else
    {
      v21 = 4;
      v60 = 4;
    }
    *(_DWORD *)(v1 + 1364) = v21;
    *(_DWORD *)(v1 + 1648) = v65;
    *(_DWORD *)(v1 + 1664) = v66;
    *(_BYTE *)(v1 + 304833) = v67 == 1;
    *(_BYTE *)(v1 + 304888) = v68 != 0;
    *(_BYTE *)(v1 + 304889) = v69 != 0;
    if ( v70 )
      *(_DWORD *)(v1 + 305600) = v70;
    if ( v71 )
      *(_DWORD *)(v1 + 305604) = v71;
    if ( !g_IsInternalRelease && !g_OSTestSigningEnabled )
    {
      g_bDbgForceUsb4MonitorSupport = 0;
      g_bDbgTreatUsb4MonitorAsNormal = 0;
      g_DbgUsb4MonitorPowerOnDelayInSeconds = 0;
    }
    *(_BYTE *)(v1 + 305672) = v72 != 0;
    *(_DWORD *)(v1 + 305880) = v73;
    DXGGLOBAL::SetNodeUsageTelemetryTimer((DXGGLOBAL *)v1);
  }
  *(_DWORD *)(v1 + 872) = 0;
  v22 = *(_OWORD *)L"Y\\MACHINE\\System\\ControlSet001\\Control\\Terminal Server\\WinStations";
  v74 = 0;
  v276[0] = *(_OWORD *)L"\\REGISTRY\\MACHINE\\System\\ControlSet001\\Control\\Terminal Server\\WinStations";
  *(_QWORD *)&v92.Length = 9830548LL;
  v23 = *(_OWORD *)L"E\\System\\ControlSet001\\Control\\Terminal Server\\WinStations";
  *(_QWORD *)&v91.Length = 2228256LL;
  v276[1] = v22;
  v24 = *(_OWORD *)L"\\ControlSet001\\Control\\Terminal Server\\WinStations";
  v276[2] = v23;
  v25 = *(_OWORD *)L"Set001\\Control\\Terminal Server\\WinStations";
  v276[3] = v24;
  v26 = *(_OWORD *)L"ontrol\\Terminal Server\\WinStations";
  v276[4] = v25;
  v27 = *(_OWORD *)L"erminal Server\\WinStations";
  v276[5] = v26;
  v276[6] = v27;
  v276[7] = *(_OWORD *)L"Server\\WinStations";
  v28 = *(_DWORD *)L"ns";
  v276[8] = *(_OWORD *)L"inStations";
  v277 = v28;
  v278 = aRegistryMachin_13[74];
  v92.Buffer = (wchar_t *)v276;
  v275 = aDwmframeinterv[16];
  v91.Buffer = (wchar_t *)v274;
  v274[0] = *(_OWORD *)L"DWMFRAMEINTERVAL";
  v274[1] = *(_OWORD *)L"INTERVAL";
  if ( ReadRegistryDwordKeyValue(&v92, &v91, &v74) >= 0 && v74 )
    *(_DWORD *)(v1 + 305152) = v74;
  DxgkSharedObjectTypes = CreateDxgkSharedObjectTypes(v29);
  v13 = DxgkSharedObjectTypes;
  if ( DxgkSharedObjectTypes < 0 )
  {
    WdLogSingleEntry1(3LL, DxgkSharedObjectTypes);
    WdLogGlobalForLineNumber = 2143;
    return v13;
  }
  v31 = v57;
  if ( !v57 || ((v57 - 1) & v57) != 0 )
  {
    v31 = v75;
    v57 = v75;
  }
  v32 = v59;
  if ( !v59 || ((v59 - 1) & v59) != 0 )
  {
    v32 = v76;
    v59 = v76;
  }
  if ( !g_OSTestSigningEnabled )
  {
    if ( v55 < 0x1000 || v55 * v31 > 0x1000000 )
    {
      v55 = 0x1000000;
      v57 = 1;
    }
    if ( v56 < 0x1000 || v56 * v32 > 0x1000000 )
    {
      v56 = 0x1000000;
      v59 = 1;
    }
  }
  if ( v58 > 0xE10 )
    v58 = 3600;
  v33 = (-(__int64)(g_IsInternalReleaseOrDbg != 0) & 0xFFFFFFFFFFFFFF40uLL) + 256;
  v34 = operator new(112LL, 1265072196LL);
  if ( v34 )
    v35 = DXGDIAGNOSTICS::DXGDIAGNOSTICS(v34, v55, v57, v33, v58);
  else
    v35 = 0LL;
  *(_QWORD *)(v1 + 928) = v35;
  v36 = operator new(112LL, 1265072196LL);
  if ( v36 )
  {
    Flagsb[0] = v58;
    v37 = DXGDIAGNOSTICS::DXGDIAGNOSTICS(v36, v56, v59, v33, *(_QWORD *)Flagsb);
  }
  else
  {
    v37 = 0LL;
  }
  *(_QWORD *)(v1 + 936) = v37;
  if ( !*(_QWORD *)(v1 + 928) )
  {
    WdLogSingleEntry1(6LL, v55);
    WdLogGlobalForLineNumber = 2197;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for internal diagnostics buffers (SmallInternalDiagnosticsSize = 0x%I64x).",
      v55,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  if ( !v37 )
  {
    WdLogSingleEntry1(6LL, v56);
    WdLogGlobalForLineNumber = 2203;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for external diagnostics buffers (SmallExternalDiagnosticsSize = 0x%I64x).",
      v56,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v38 = (DXGSESSIONMGR *)operator new(448LL, 1265072196LL);
  if ( v38 )
    v39 = DXGSESSIONMGR::DXGSESSIONMGR(v38);
  else
    v39 = 0LL;
  *(_QWORD *)(v1 + 944) = v39;
  if ( !v39 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2210;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for dxgkrnl session manager.",
      2210LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v40 = *(_DWORD *)(v1 + 1364);
  v41 = (unsigned int)(v40 + 31) >> 5;
  v42 = 4LL * ((unsigned int)v41 + ((unsigned int)(1055 - v40) >> 5));
  if ( !is_mul_ok((unsigned int)v41 + ((unsigned int)(1055 - v40) >> 5), 4uLL) )
    v42 = -1LL;
  v43 = (ULONG *)operator new[](v42, 1265072196LL, 256LL);
  *(_QWORD *)(v1 + 864) = v43;
  if ( !v43 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2219;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for dxgkrnl adapter ordinal bits.",
      2219LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  RtlInitializeBitMap((PRTL_BITMAP)(v1 + 832), v43, *(_DWORD *)(v1 + 1364));
  RtlInitializeBitMap((PRTL_BITMAP)(v1 + 848), (PULONG)(*(_QWORD *)(v1 + 864) + 4 * v41), 1024 - *(_DWORD *)(v1 + 1364));
  if ( DXGPROCESS::CreateDxgProcess((struct DXGPROCESS **)(v1 + 1368), 0LL, 0LL, 0, 0LL) < 0 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2233;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate memory for system process.",
      2233LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  if ( PsInitialSystemProcess != *(PEPROCESS *)(*(_QWORD *)(v1 + 1368) + 56LL) )
  {
    WdLogSingleEntry0(1LL);
    WdLogGlobalForLineNumber = 2236;
    DxgkLogInternalTriageEvent(
      0,
      262146,
      -1,
      (unsigned int)L"PsInitialSystemProcess == m_pSystemDxgProcess->GetEProcess()",
      2236LL,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  v44 = operator new(640LL, 1265072196LL);
  v45 = (_BYTE *)v44;
  if ( v44 )
  {
    *(_BYTE *)v44 = 1;
    *(_QWORD *)(v44 + 16) = 0LL;
    *(_QWORD *)(v44 + 24) = 0LL;
    *(_QWORD *)(v44 + 32) = 0LL;
    *(_DWORD *)(v44 + 40) = 0;
    *(_DWORD *)(v44 + 44) = 69;
    *(_DWORD *)(v44 + 48) = 1;
    *(_DWORD *)(v44 + 632) = 0;
    memset((void *)(v44 + 56), 0, 0x240uLL);
    *v45 = 0;
  }
  else
  {
    v45 = 0LL;
  }
  *(_QWORD *)(v1 + 1464) = v45;
  if ( !v45 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 2241;
    DxgkLogInternalTriageEvent(0, 262145, -1, (unsigned int)L"Failed to Qdc cache.", 2241LL, 0LL, 0LL, 0LL, 0LL);
    return 3221225495LL;
  }
  KeInitializeSpinLock(&qword_1C01537D0);
  DXGVALIDATION::InitializeBootSettings((DXGVALIDATION *)(v1 + 1652));
  DXGGLOBAL::CsExitInitiatedWnfSubscription((DXGGLOBAL *)v1);
  KeInitializeTimer((PKTIMER)(v1 + 1904));
  KeInitializeDpc((PRKDPC)(v1 + 1968), CsExitInitiatedReleaseComponentReferences, (PVOID)v1);
  LOBYTE(OutputBuffer) = 0;
  v46 = ZwPowerInformation((POWER_INFORMATION_LEVEL)66, 0LL, 0, &OutputBuffer, 1u);
  if ( v46 >= 0 )
  {
    if ( (_BYTE)OutputBuffer )
      DXGGLOBAL::SubscribeWNFForCSAccounting((DXGGLOBAL *)v1);
  }
  else
  {
    v47 = v46;
    WdLogSingleEntry1(2LL, v46);
    WdLogGlobalForLineNumber = 2278;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Failed to get the platformInformation. Status : 0x%I64x",
      v47,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  *(_QWORD *)(v1 + 2056) = v1;
  *(_QWORD *)(v1 + 2048) = CsExitInitiatedReleaseComponentReferencesPassiveLevel;
  *(_QWORD *)(v1 + 2032) = 0LL;
  DXGGLOBAL::InitializeResourceManagerSid((DXGGLOBAL *)v1);
  *(_DWORD *)(v1 + 304820) &= ~1u;
  *(_DWORD *)(v1 + 304808) = 10;
  *(_DWORD *)(v1 + 304812) = 50;
  *(_DWORD *)(v1 + 304816) = 30;
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 1752));
  DisplayDiagnostics::Initialize((PVOID)(v1 + 304960));
  v48 = PoRegisterPowerSettingCallback(
          0LL,
          &GUID_ADVANCED_COLOR_QUALITY_BIAS,
          DXGGLOBAL::AdvancedColorPowerSettingsCallback,
          (PVOID)v1,
          0LL);
  v7 = v48;
  if ( v48 < 0 )
  {
    WdLogSingleEntry1(2LL, v48);
    WdLogGlobalForLineNumber = 2315;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"PoRegisterPowerSettingCallback for GUID_HDR_DISPLAY_QUALITY_BIAS failed with status:0x%I64x",
      v7,
      0LL,
      0LL,
      0LL,
      0LL);
    return (unsigned int)v7;
  }
  v49 = PoRegisterPowerSettingCallback(0LL, &GUID_ACDC_POWER_SOURCE, DXGGLOBAL::AcDcPowerSourceCallback, (PVOID)v1, 0LL);
  v50 = v49;
  if ( v49 < 0 )
  {
    WdLogSingleEntry1(2LL, v49);
    WdLogGlobalForLineNumber = 2325;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"PoRegisterPowerSettingCallback for GUID_ACDC_POWER_SOURCE failed with status:0x%I64x",
      v50,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  return (unsigned int)v50;
}

NTSTATUS NotifyUserMSBDAIfApplicable()
{
  ULONGLONG v0; // rax
  NTSTATUS result; // eax
  _DWORD v2[4]; // [rsp+30h] [rbp-D0h] BYREF
  __int64 v3; // [rsp+40h] [rbp-C0h] BYREF
  int v4; // [rsp+48h] [rbp-B8h]
  const wchar_t *v5; // [rsp+50h] [rbp-B0h]
  _DWORD *v6; // [rsp+58h] [rbp-A8h]
  int v7; // [rsp+60h] [rbp-A0h]
  _DWORD *v8; // [rsp+68h] [rbp-98h]
  int v9; // [rsp+70h] [rbp-90h]
  __int64 v10; // [rsp+78h] [rbp-88h]
  int v11; // [rsp+80h] [rbp-80h]
  __int64 v12; // [rsp+88h] [rbp-78h]
  __int128 v13; // [rsp+90h] [rbp-70h]
  __int128 v14; // [rsp+A0h] [rbp-60h]
  _OSVERSIONINFOEXW VersionInfo; // [rsp+B0h] [rbp-50h] BYREF

  memset(&VersionInfo, 0, sizeof(VersionInfo));
  VersionInfo.wProductType = 1;
  v0 = VerSetConditionMask(0LL, 0x80u, 1u);
  result = RtlVerifyVersionInfo(&VersionInfo, 0x80u, v0);
  if ( result >= 0 )
  {
    v2[0] = 0;
    v3 = 0LL;
    v10 = 0LL;
    v11 = 0;
    v12 = 0LL;
    v5 = L"BasicDisplayUserNotified";
    v6 = v2;
    v8 = v2;
    v4 = 288;
    v7 = 67108868;
    v9 = 4;
    v13 = 0LL;
    v14 = 0LL;
    result = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\BasicDisplay", &v3);
    if ( !v2[0] )
      return WdDiagNotifyUser(0LL, 8LL, 0LL, 0LL, 0LL);
  }
  return result;
}

__int64 __fastcall DpiFdoInitializeFdo(_QWORD *StartContext)
{
  __int64 v1; // rdi
  char v3; // r12
  char v4; // si
  char v5; // r15
  int v6; // eax
  size_t v7; // rbx
  void *Pool2; // rax
  __int64 v9; // rbx
  int DevicePropertyString; // eax
  int v11; // eax
  __int64 v12; // rcx
  struct _DEVICE_OBJECT *v13; // rcx
  int MiniportInterface; // eax
  struct _DEVICE_OBJECT *v15; // rcx
  NTSTATUS v16; // eax
  __int64 v17; // rax
  NTSTATUS v18; // eax
  __int64 v19; // rax
  NTSTATUS v20; // eax
  __int64 v21; // rax
  _WORD *v22; // rsi
  int v23; // eax
  bool v24; // bl
  __int64 v25; // rdx
  __int64 v26; // rax
  int v27; // eax
  size_t v28; // r8
  void *v29; // rcx
  void *v30; // rcx
  void *v31; // rcx
  void *v32; // rcx
  void *v33; // rcx
  void *v34; // rcx
  void *v35; // rcx
  void *v36; // rcx
  void (__fastcall *v37)(_QWORD); // rax
  void (__fastcall *v38)(_QWORD); // rax
  struct SYSMM_ADAPTER *v39; // rcx
  int Size; // [rsp+28h] [rbp-E0h]
  ULONG Sizea[2]; // [rsp+28h] [rbp-E0h]
  int Sizec; // [rsp+28h] [rbp-E0h]
  int Sizeb; // [rsp+28h] [rbp-E0h]
  char v45; // [rsp+48h] [rbp-C0h] BYREF
  char v46; // [rsp+49h] [rbp-BFh] BYREF
  char Data; // [rsp+4Ah] [rbp-BEh] BYREF
  ULONG RequiredSize; // [rsp+4Ch] [rbp-BCh] BYREF
  ULONG Type; // [rsp+50h] [rbp-B8h] BYREF
  unsigned int v50; // [rsp+54h] [rbp-B4h] BYREF
  _QWORD SymbolicLinkName[3]; // [rsp+58h] [rbp-B0h] BYREF
  __int64 v52; // [rsp+70h] [rbp-98h] BYREF
  void *ThreadHandle; // [rsp+78h] [rbp-90h] BYREF
  PVOID Object; // [rsp+80h] [rbp-88h] BYREF
  __int64 v55; // [rsp+88h] [rbp-80h] BYREF
  int v56; // [rsp+90h] [rbp-78h]
  const wchar_t *v57; // [rsp+98h] [rbp-70h]
  unsigned int *v58; // [rsp+A0h] [rbp-68h]
  int v59; // [rsp+A8h] [rbp-60h]
  unsigned int *v60; // [rsp+B0h] [rbp-58h]
  int v61; // [rsp+B8h] [rbp-50h]
  __int64 v62; // [rsp+C0h] [rbp-48h]
  int v63; // [rsp+C8h] [rbp-40h]
  const wchar_t *v64; // [rsp+D0h] [rbp-38h]
  unsigned int *v65; // [rsp+D8h] [rbp-30h]
  int v66; // [rsp+E0h] [rbp-28h]
  unsigned int *v67; // [rsp+E8h] [rbp-20h]
  int v68; // [rsp+F0h] [rbp-18h]
  __int64 v69; // [rsp+F8h] [rbp-10h]
  int v70; // [rsp+100h] [rbp-8h]
  const wchar_t *v71; // [rsp+108h] [rbp+0h]
  int *v72; // [rsp+110h] [rbp+8h]
  int v73; // [rsp+118h] [rbp+10h]
  int *v74; // [rsp+120h] [rbp+18h]
  int v75; // [rsp+128h] [rbp+20h]
  __int64 v76; // [rsp+130h] [rbp+28h]
  int v77; // [rsp+138h] [rbp+30h]
  const wchar_t *v78; // [rsp+140h] [rbp+38h]
  int *v79; // [rsp+148h] [rbp+40h]
  int v80; // [rsp+150h] [rbp+48h]
  int *v81; // [rsp+158h] [rbp+50h]
  int v82; // [rsp+160h] [rbp+58h]
  __int64 v83; // [rsp+168h] [rbp+60h]
  int v84; // [rsp+170h] [rbp+68h]
  const wchar_t *v85; // [rsp+178h] [rbp+70h]
  __int64 *v86; // [rsp+180h] [rbp+78h]
  int v87; // [rsp+188h] [rbp+80h]
  __int64 v88; // [rsp+190h] [rbp+88h]
  int v89; // [rsp+198h] [rbp+90h]
  __int64 v90; // [rsp+1A0h] [rbp+98h]
  int v91; // [rsp+1A8h] [rbp+A0h]
  __int64 v92; // [rsp+1B0h] [rbp+A8h]
  __int128 v93; // [rsp+1B8h] [rbp+B0h]
  __int128 v94; // [rsp+1C8h] [rbp+C0h]

  v1 = StartContext[8];
  RequiredSize = 0;
  Type = 0;
  ThreadHandle = 0LL;
  v3 = 0;
  *(_OWORD *)&SymbolicLinkName[1] = 0LL;
  *(_QWORD *)(v1 + 112) = &DpiFdoDispatchInternalIoctl;
  *(_QWORD *)(v1 + 144) = DpiFdoDispatchSystemControl;
  v4 = 0;
  v5 = 0;
  *(_QWORD *)(v1 + 352) = &DpiFdoHandleQueryInterface;
  *(_QWORD *)(v1 + 344) = &DpiFdoHandleQueryDeviceRelations;
  LODWORD(v52) = 0;
  v55 = 0LL;
  v57 = L"GpuVirtualizationFlags";
  v56 = 288;
  v50 = g_VgpuReplaceWarp != 0 ? 8 : 0;
  v58 = &v50;
  v59 = 67108868;
  v60 = &v50;
  v61 = 4;
  v64 = L"DisableVaBackedVm";
  v65 = &g_VgpuDisableVaBackedVm;
  v67 = &g_VgpuDisableVaBackedVm;
  v71 = L"VirtualGpuOnly";
  v72 = &g_VirtualGpuOnly;
  v74 = &g_VirtualGpuOnly;
  v78 = L"LimitNumberOfVfs";
  v79 = &g_LimitNumberOfVfs;
  v81 = &g_LimitNumberOfVfs;
  v85 = L"DisableVersionMismatchCheck";
  v86 = &v52;
  v62 = 0LL;
  v63 = 288;
  v66 = 67108868;
  v68 = 4;
  v69 = 0LL;
  v70 = 288;
  v73 = 67108868;
  v75 = 4;
  v76 = 0LL;
  v77 = 288;
  v80 = 67108868;
  v82 = 4;
  v83 = 0LL;
  v84 = 288;
  v87 = 67108868;
  v88 = 0LL;
  v89 = 0;
  v90 = 0LL;
  v91 = 0;
  v92 = 0LL;
  v93 = 0LL;
  v94 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v55);
  g_bCreateParavirtualizedGpu = v50 & 1;
  g_VgpuReplaceWarp = (v50 >> 3) & 1;
  v6 = *(_DWORD *)(v1 + 504);
  g_ForceSecureVirtualMachine = (v50 >> 2) & 1;
  if ( v6 )
  {
    v7 = (unsigned int)(8 * v6);
    Pool2 = (void *)ExAllocatePool2(64LL, v7, 1953656900LL);
    *(_QWORD *)(v1 + 2832) = Pool2;
    if ( !Pool2 )
    {
      LODWORD(v9) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 9400;
      goto LABEL_140;
    }
    memset(Pool2, 0, v7);
    **(_QWORD **)(v1 + 2832) = StartContext;
    *(_DWORD *)(v1 + 2840) = 1;
  }
  *(_DWORD *)(v1 + 3604) = -1;
  DevicePropertyString = DpiGetDevicePropertyString(
                           *(PDEVICE_OBJECT *)(v1 + 152),
                           DevicePropertyDeviceDescription,
                           (__int64)&RequiredSize);
  LODWORD(v9) = DevicePropertyString;
  if ( DevicePropertyString < 0 )
  {
    WdLogSingleEntry1(2LL, DevicePropertyString);
    WdLogGlobalForLineNumber = 9427;
LABEL_41:
    v4 = 0;
    goto LABEL_140;
  }
  DpiGetDevicePropertyDataString(
    *(PDEVICE_OBJECT *)(v1 + 152),
    (DEVPROPKEY *)&DEVPKEY_Device_DriverVersion,
    v1 + 4952,
    (__int64)&RequiredSize);
  IoGetDevicePropertyData(
    *(PDEVICE_OBJECT *)(v1 + 152),
    &DEVPKEY_Device_DriverDate,
    0,
    0,
    8u,
    (PVOID)(v1 + 4960),
    &RequiredSize,
    &Type);
  IoGetDevicePropertyData(
    *(PDEVICE_OBJECT *)(v1 + 152),
    &DEVPKEY_Device_DriverRank,
    0,
    0,
    4u,
    (PVOID)(v1 + 4968),
    &RequiredSize,
    &Type);
  if ( !(_DWORD)v52 )
  {
    v11 = DpiFdoValidateKmdAndPnpVersionMatch(v1);
    LODWORD(v9) = v11;
    if ( v11 < 0 )
    {
      WdLogSingleEntry1(2LL, v11);
      WdLogGlobalForLineNumber = 9474;
      goto LABEL_41;
    }
  }
  v12 = *(_QWORD *)(v1 + 152);
  v46 = 0;
  if ( (int)DpiGetDevicePropertyDataBoolean(v12, &DEVPKEY_Device_InstallInProgress, &v46) >= 0 && v46 )
  {
    v13 = *(struct _DEVICE_OBJECT **)(v1 + 152);
    v45 = 0;
    IoSetDevicePropertyData(v13, &DEVPKEY_Device_InstallInProgress, 0, 0, 0x11u, 1u, &v45);
  }
  if ( *(_BYTE *)(v1 + 1153) )
  {
    if ( *(_BYTE *)(v1 + 480) )
    {
      MiniportInterface = DpiQueryMiniportInterface(
                            (_DWORD)StartContext,
                            (unsigned int)&GUID_DEVINTERFACE_MSBDD_FALLBACK,
                            56,
                            1,
                            Size,
                            v1 + 944);
      LODWORD(v9) = MiniportInterface;
      if ( MiniportInterface < 0 || !*(_QWORD *)(v1 + 976) || !*(_QWORD *)(v1 + 984) || !*(_QWORD *)(v1 + 992) )
      {
        WdLogSingleEntry3(0LL, 275LL, 21LL, MiniportInterface, *(_QWORD *)Sizea);
        WdLogGlobalForLineNumber = 9527;
        goto LABEL_137;
      }
    }
  }
  v3 = 1;
  if ( *(_BYTE *)(v1 + 1158) )
  {
    v15 = *(struct _DEVICE_OBJECT **)(v1 + 152);
    Data = 0;
    if ( IoGetDevicePropertyData(v15, &DEVPKEY_Gpu_IddVirtualMonitorDevice, 0, 0, 1u, &Data, &RequiredSize, &Type) >= 0
      && Type == 17
      && RequiredSize == 1
      && Data == -1 )
    {
      *(_BYTE *)(v1 + 1159) = 1;
    }
  }
  v16 = IoRegisterDeviceInterface(
          *(PDEVICE_OBJECT *)(v1 + 152),
          &GUID_COMPUTE_DEVICE_ARRIVAL,
          0LL,
          (PUNICODE_STRING)&SymbolicLinkName[1]);
  LODWORD(v9) = v16;
  if ( v16 < 0 )
  {
    WdLogSingleEntry1(2LL, v16);
    WdLogGlobalForLineNumber = 9566;
    goto LABEL_41;
  }
  v4 = 1;
  v17 = ExAllocatePool2(64LL, WORD1(SymbolicLinkName[1]), 1953656900LL);
  *(_QWORD *)(v1 + 2856) = v17;
  if ( !v17 )
  {
    LODWORD(v9) = -1073741801;
    WdLogSingleEntry1(6LL, -1073741801LL);
    WdLogGlobalForLineNumber = 9587;
    goto LABEL_140;
  }
  *(_DWORD *)(v1 + 2848) = SymbolicLinkName[1];
  RtlCopyUnicodeString((PUNICODE_STRING)(v1 + 2848), (PCUNICODE_STRING)&SymbolicLinkName[1]);
  RtlFreeUnicodeString((PUNICODE_STRING)&SymbolicLinkName[1]);
  if ( !*(_BYTE *)(v1 + 2722) )
  {
    v18 = IoRegisterDeviceInterface(
            *(PDEVICE_OBJECT *)(v1 + 152),
            &GUID_DISPLAY_DEVICE_ARRIVAL,
            0LL,
            (PUNICODE_STRING)&SymbolicLinkName[1]);
    LODWORD(v9) = v18;
    if ( v18 < 0 )
    {
      WdLogSingleEntry1(2LL, v18);
      WdLogGlobalForLineNumber = 9617;
      goto LABEL_41;
    }
    v19 = ExAllocatePool2(64LL, WORD1(SymbolicLinkName[1]), 1953656900LL);
    *(_QWORD *)(v1 + 2872) = v19;
    if ( !v19 )
    {
      LODWORD(v9) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 9638;
      goto LABEL_140;
    }
    *(_DWORD *)(v1 + 2864) = SymbolicLinkName[1];
    RtlCopyUnicodeString((PUNICODE_STRING)(v1 + 2864), (PCUNICODE_STRING)&SymbolicLinkName[1]);
    RtlFreeUnicodeString((PUNICODE_STRING)&SymbolicLinkName[1]);
  }
  *(_BYTE *)(v1 + 482) = 0;
  *(_BYTE *)(v1 + 484) = 0;
  *(_QWORD *)(v1 + 488) = 0LL;
  if ( !*(_BYTE *)(v1 + 480) )
  {
    KeInitializeEvent((PRKEVENT)(v1 + 4056), SynchronizationEvent, 0);
    *(_QWORD *)(v1 + 4096) = v1 + 4088;
    *(_QWORD *)(v1 + 4088) = v1 + 4088;
    KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 4208));
    KeInitializeEvent((PRKEVENT)(v1 + 4224), NotificationEvent, 1u);
    KeInitializeEvent((PRKEVENT)(v1 + 4248), NotificationEvent, 1u);
    *(_BYTE *)(v1 + 484) = 1;
    *(_QWORD *)(v1 + 4272) = 0LL;
    *(_DWORD *)(v1 + 4216) = 0;
    memset((void *)(v1 + 4112), 0, 0x60uLL);
    *(_DWORD *)(v1 + 4128) = 1953656900;
    *(_DWORD *)(v1 + 4132) = 11;
    *(_DWORD *)(v1 + 4152) = 64;
    KeInitializeTimer((PKTIMER)(v1 + 4288));
    KeInitializeDpc((PRKDPC)(v1 + 4352), DpiSuspendAdapterDpc, (PVOID)v1);
    v20 = PsCreateSystemThread(&ThreadHandle, 0x1FFFFFu, 0LL, 0LL, 0LL, DpiPowerArbiterThread, StartContext);
    LODWORD(v9) = v20;
    if ( v20 < 0 )
    {
      WdLogSingleEntry1(2LL, v20);
      WdLogGlobalForLineNumber = 9716;
      goto LABEL_41;
    }
    Object = 0LL;
    v9 = ObReferenceObjectByHandle(ThreadHandle, 0x1FFFFFu, 0LL, 0, &Object, 0LL);
    *(_QWORD *)(v1 + 4048) = Object;
    ZwClose(ThreadHandle);
    if ( (int)v9 < 0 )
    {
      WdLogSingleEntry1(2LL, v9);
      WdLogGlobalForLineNumber = 9738;
      goto LABEL_41;
    }
  }
  KeInitializeEvent((PRKEVENT)(v1 + 3816), NotificationEvent, 1u);
  *(_QWORD *)(v1 + 3592) = v1 + 3584;
  *(_QWORD *)(v1 + 3584) = v1 + 3584;
  ExInitializeResourceLite((PERESOURCE)(v1 + 3424));
  *(_QWORD *)(v1 + 3624) = v1 + 3616;
  *(_QWORD *)(v1 + 3616) = v1 + 3616;
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 3608));
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 3640));
  KeInitializeEvent((PRKEVENT)(v1 + 3648), NotificationEvent, 1u);
  *(_QWORD *)(v1 + 5456) = v1 + 5448;
  *(_QWORD *)(v1 + 5448) = v1 + 5448;
  KeInitializeSpinLock((PKSPIN_LOCK)(v1 + 5464));
  IoCsqInitialize(
    (PIO_CSQ)(v1 + 5384),
    DpiPendingIrpCancelQueueInsert,
    DpiPendingIrpCancelQueueRemove,
    DpiPendingIrpCancelQueuePick,
    DpiPendingIrpCancelQueueAcquireLock,
    DpiPendingIrpCancelQueueReleaseLock,
    DpiPendingIrpCancelQueueComplete);
  *(_QWORD *)(v1 + 5536) = 0LL;
  *(_QWORD *)(v1 + 5544) = 0LL;
  KeInitializeEvent((PRKEVENT)(v1 + 5552), NotificationEvent, 0);
  *(_DWORD *)(v1 + 5528) = 1;
  *(_DWORD *)(v1 + 5496) = 0;
  KeInitializeMutex((PRKMUTEX)(v1 + 3528), 0);
  KeInitializeMutex((PRKMUTEX)(v1 + 3704), 0);
  *(_QWORD *)(v1 + 3776) = v1 + 3768;
  *(_QWORD *)(v1 + 3768) = v1 + 3768;
  *(_QWORD *)(v1 + 3800) = v1 + 3792;
  *(_QWORD *)(v1 + 3792) = v1 + 3792;
  *(_QWORD *)(v1 + 3696) = v1 + 3688;
  *(_QWORD *)(v1 + 3688) = v1 + 3688;
  ExInitializeResourceLite((PERESOURCE)(v1 + 3912));
  LODWORD(v9) = DpiFdoInitializeAdapterUniqueString(StartContext);
  v4 = 0;
  if ( (int)v9 < 0 )
  {
LABEL_139:
    ExDeleteResourceLite((PERESOURCE)(v1 + 3912));
    ExDeleteResourceLite((PERESOURCE)(v1 + 3424));
    goto LABEL_140;
  }
  v5 = 1;
  DpiQueryBusInterface(*(PDEVICE_OBJECT *)(v1 + 152), v1 + 2992);
  DpiQueryBusInterface(*(PDEVICE_OBJECT *)(v1 + 152), v1 + 3040);
  DpiQueryMiniportInterface((_DWORD)StartContext, (unsigned int)&GUID_DEVINTERFACE_I2C, 48, 1, Sizec, v1 + 3088);
  v21 = *(_QWORD *)(v1 + 40);
  *(_DWORD *)(v1 + 6008) = 1;
  if ( !*(_BYTE *)(v21 + 133) && !*(_BYTE *)(v1 + 1158) )
  {
    v22 = (_WORD *)(v1 + 5880);
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_WDDM_INTERFACE_DISPLAYMUX_2,
                128,
                2,
                Sizeb,
                v1 + 5880) >= 0 )
    {
      if ( *v22 != 128
        || *(_WORD *)(v1 + 5882) != 2
        || !*(_QWORD *)(v1 + 5912)
        || !*(_QWORD *)(v1 + 5920)
        || !*(_QWORD *)(v1 + 5928)
        || !*(_QWORD *)(v1 + 5936)
        || !*(_QWORD *)(v1 + 5944)
        || !*(_QWORD *)(v1 + 5952)
        || !*(_QWORD *)(v1 + 5960)
        || !*(_QWORD *)(v1 + 5968)
        || !*(_QWORD *)(v1 + 5976)
        || !*(_QWORD *)(v1 + 5984)
        || !*(_QWORD *)(v1 + 5992)
        || !*(_QWORD *)(v1 + 6000) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        WdLogGlobalForLineNumber = 9887;
        goto LABEL_85;
      }
      LODWORD(SymbolicLinkName[0]) = 0;
      if ( (int)DpiDxgkDdiDisplayMuxGetDriverSupportLevel(v1, SymbolicLinkName) < 0 )
      {
        *(_DWORD *)(v1 + 6008) = 1;
      }
      else
      {
        v23 = SymbolicLinkName[0];
        *(_DWORD *)(v1 + 6008) = SymbolicLinkName[0];
        if ( v23 != 1 )
        {
          v24 = DISPLAY_MUX_MGR::DisplayMuxPresent(qword_1C0154100);
          if ( DISPLAY_MUX_MGR::ShouldHideMuxFromDriver(qword_1C0154100) )
          {
            WdLogSingleEntry0(4LL);
            WdLogGlobalForLineNumber = 9915;
            v24 = 0;
          }
          LOBYTE(v25) = v24;
          DpiDxgkDdiDisplayMuxReportPresence(v1, v25);
          *(_BYTE *)(v1 + 6377) = v24;
        }
      }
    }
  }
  v26 = *(_QWORD *)(v1 + 40);
  *(_DWORD *)(v1 + 3136) = 0;
  if ( !*(_BYTE *)(v26 + 133) || *(_BYTE *)(v1 + 1158) )
  {
    v22 = (_WORD *)(v1 + 3144);
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM_3,
                128,
                4,
                Sizeb,
                v1 + 3144) >= 0 )
    {
      if ( *v22 != 128
        || (v27 = 4, *(_WORD *)(v1 + 3146) != 4)
        || !*(_QWORD *)(v1 + 3176)
        || !*(_QWORD *)(v1 + 3184)
        || !*(_QWORD *)(v1 + 3192)
        || !*(_QWORD *)(v1 + 3200)
        || !*(_QWORD *)(v1 + 3208)
        || !*(_QWORD *)(v1 + 3216)
        || !*(_QWORD *)(v1 + 3224)
        || !*(_QWORD *)(v1 + 3232)
        || !*(_QWORD *)(v1 + 3240)
        || !*(_QWORD *)(v1 + 3248)
        || !*(_QWORD *)(v1 + 3256)
        || !*(_QWORD *)(v1 + 3264) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        WdLogGlobalForLineNumber = 10051;
LABEL_85:
        v28 = 128LL;
LABEL_86:
        v29 = v22;
LABEL_87:
        memset(v29, 0, v28);
        v4 = 0;
        goto LABEL_139;
      }
      goto LABEL_115;
    }
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM_2,
                112,
                3,
                Sizeb,
                v1 + 3144) >= 0 )
    {
      if ( *v22 != 112
        || (v27 = 3, *(_WORD *)(v1 + 3146) != 3)
        || !*(_QWORD *)(v1 + 3176)
        || !*(_QWORD *)(v1 + 3184)
        || !*(_QWORD *)(v1 + 3192)
        || !*(_QWORD *)(v1 + 3200)
        || !*(_QWORD *)(v1 + 3208)
        || !*(_QWORD *)(v1 + 3216)
        || !*(_QWORD *)(v1 + 3224)
        || !*(_QWORD *)(v1 + 3232)
        || !*(_QWORD *)(v1 + 3240)
        || !*(_QWORD *)(v1 + 3248) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        v28 = 112LL;
        WdLogGlobalForLineNumber = 10102;
        goto LABEL_86;
      }
      goto LABEL_115;
    }
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM_2_JTP,
                120,
                2,
                Sizeb,
                v1 + 3144) >= 0 )
    {
      v27 = 2;
      if ( *v22 != 120
        || *(_WORD *)(v1 + 3146) != 2
        || !*(_QWORD *)(v1 + 3176)
        || !*(_QWORD *)(v1 + 3184)
        || !*(_QWORD *)(v1 + 3192)
        || !*(_QWORD *)(v1 + 3200)
        || !*(_QWORD *)(v1 + 3208)
        || !*(_QWORD *)(v1 + 3216)
        || !*(_QWORD *)(v1 + 3224)
        || !*(_QWORD *)(v1 + 3232)
        || !*(_QWORD *)(v1 + 3240)
        || !*(_QWORD *)(v1 + 3256) )
      {
        LODWORD(v9) = -1073741811;
        WdLogSingleEntry1(2LL, -1073741811LL);
        v28 = 120LL;
        WdLogGlobalForLineNumber = 10155;
        goto LABEL_86;
      }
LABEL_115:
      *(_DWORD *)(v1 + 3136) = v27;
      goto LABEL_119;
    }
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_OPM,
                104,
                1,
                Sizeb,
                v1 + 3144) >= 0 )
      *(_DWORD *)(v1 + 3136) = 1;
  }
LABEL_119:
  *(_DWORD *)(v1 + 3344) = -1;
  if ( byte_1C0153A96
    && *(_DWORD *)(*(_QWORD *)(StartContext[8] + 40LL) + 28LL) >= 0x4000u
    && (!*(_BYTE *)(*(_QWORD *)(v1 + 40) + 133LL) || *(_BYTE *)(v1 + 1158)) )
  {
    if ( (int)DpiQueryMiniportInterface(
                (_DWORD)StartContext,
                (unsigned int)&GUID_DEVINTERFACE_MIRACAST_DISPLAY,
                64,
                1,
                Sizeb,
                v1 + 3272) < 0 )
    {
      memset((void *)(v1 + 3272), 0, 0x40uLL);
    }
    else if ( *(_WORD *)(v1 + 3272) < 0x40u
           || *(_WORD *)(v1 + 3274) != 1
           || !*(_QWORD *)(v1 + 3304)
           || !*(_QWORD *)(v1 + 3312)
           || !*(_QWORD *)(v1 + 3320)
           || !*(_QWORD *)(v1 + 3328) )
    {
      LODWORD(v9) = -1073741811;
      WdLogSingleEntry1(2LL, -1073741811LL);
      v29 = (void *)(v1 + 3272);
      WdLogGlobalForLineNumber = 10232;
      v28 = 64LL;
      goto LABEL_87;
    }
  }
  if ( *(_BYTE *)(v1 + 1159) )
    *(_QWORD *)(v1 + 120) = DpiFdoDispatchIoctl;
  if ( *(_BYTE *)(v1 + 1158) )
  {
    *(_QWORD *)(v1 + 104) = &DpiFdoDispatchCreate;
    *(_QWORD *)(v1 + 96) = &DpiFdoDispatchCleanupAndClose;
  }
  v9 = StartContext[8];
  memset((void *)(v9 + 4504), 0, 0x48uLL);
  memset((void *)(v9 + 4576), 0, 0x48uLL);
  memset((void *)(v9 + 4648), 0, 0x58uLL);
  *(_OWORD *)(v9 + 4736) = 0LL;
  *(_OWORD *)(v9 + 4752) = 0LL;
  *(_OWORD *)(v9 + 4768) = 0LL;
  *(_QWORD *)(v9 + 4784) = 0LL;
  memset((void *)(v9 + 4792), 0, 0x58uLL);
  LODWORD(v9) = DpiInitializeBlockList(StartContext);
LABEL_137:
  v5 = v3;
  if ( (int)v9 >= 0 )
    return (unsigned int)v9;
  v4 = 0;
  if ( v3 == 1 )
    goto LABEL_139;
LABEL_140:
  if ( *(_QWORD *)(v1 + 4048) )
    DpiRequestIoPowerState(StartContext, 7LL);
  if ( v4 == 1 )
    RtlFreeUnicodeString((PUNICODE_STRING)&SymbolicLinkName[1]);
  if ( v5 )
  {
    RtlFreeUnicodeString((PUNICODE_STRING)(v1 + 4880));
    RtlFreeUnicodeString((PUNICODE_STRING)(v1 + 4896));
  }
  v30 = *(void **)(v1 + 3416);
  *(_DWORD *)(v1 + 3400) = 0;
  if ( v30 )
  {
    ExFreePoolWithTag(v30, 0);
    *(_QWORD *)(v1 + 3416) = 0LL;
  }
  v31 = *(void **)(v1 + 3408);
  if ( v31 )
  {
    ExFreePoolWithTag(v31, 0);
    *(_QWORD *)(v1 + 3408) = 0LL;
  }
  v32 = *(void **)(v1 + 4944);
  if ( v32 )
  {
    ExFreePoolWithTag(v32, 0);
    *(_QWORD *)(v1 + 4944) = 0LL;
  }
  v33 = *(void **)(v1 + 4952);
  if ( v33 )
  {
    ExFreePoolWithTag(v33, 0);
    *(_QWORD *)(v1 + 4952) = 0LL;
  }
  v34 = *(void **)(v1 + 2832);
  if ( v34 )
  {
    ExFreePoolWithTag(v34, 0);
    *(_QWORD *)(v1 + 2832) = 0LL;
  }
  v35 = *(void **)(v1 + 2856);
  if ( v35 )
  {
    ExFreePoolWithTag(v35, 0);
    *(_QWORD *)(v1 + 2856) = 0LL;
  }
  v36 = *(void **)(v1 + 2872);
  if ( v36 )
  {
    ExFreePoolWithTag(v36, 0);
    *(_QWORD *)(v1 + 2872) = 0LL;
  }
  v37 = *(void (__fastcall **)(_QWORD))(v1 + 3016);
  if ( v37 )
  {
    v37(*(_QWORD *)(v1 + 3000));
    *(_OWORD *)(v1 + 2992) = 0LL;
    *(_OWORD *)(v1 + 3008) = 0LL;
    *(_OWORD *)(v1 + 3024) = 0LL;
  }
  v38 = *(void (__fastcall **)(_QWORD))(v1 + 3064);
  if ( v38 )
  {
    v38(*(_QWORD *)(v1 + 3048));
    *(_OWORD *)(v1 + 3040) = 0LL;
    *(_OWORD *)(v1 + 3056) = 0LL;
    *(_OWORD *)(v1 + 3072) = 0LL;
  }
  v39 = *(struct SYSMM_ADAPTER **)(v1 + 5808);
  if ( v39 )
    SysMmDestroyAdapter(v39);
  return (unsigned int)v9;
}

__int64 SmmQueryRegistry()
{
  int v0; // ebx
  __int64 v1; // rcx
  __int64 result; // rax
  unsigned int v3; // [rsp+30h] [rbp-D0h] BYREF
  unsigned int v4; // [rsp+34h] [rbp-CCh] BYREF
  unsigned int v5; // [rsp+38h] [rbp-C8h] BYREF
  unsigned int v6; // [rsp+3Ch] [rbp-C4h] BYREF
  int v7; // [rsp+40h] [rbp-C0h] BYREF
  int v8; // [rsp+44h] [rbp-BCh] BYREF
  int v9; // [rsp+48h] [rbp-B8h] BYREF
  int v10; // [rsp+4Ch] [rbp-B4h] BYREF
  int v11; // [rsp+50h] [rbp-B0h] BYREF
  int v12; // [rsp+54h] [rbp-ACh] BYREF
  int v13; // [rsp+58h] [rbp-A8h] BYREF
  int v14; // [rsp+5Ch] [rbp-A4h] BYREF
  int v15; // [rsp+60h] [rbp-A0h] BYREF
  int v16; // [rsp+64h] [rbp-9Ch] BYREF
  __int64 v17; // [rsp+70h] [rbp-90h] BYREF
  int v18; // [rsp+78h] [rbp-88h]
  const wchar_t *v19; // [rsp+80h] [rbp-80h]
  unsigned int *v20; // [rsp+88h] [rbp-78h]
  int v21; // [rsp+90h] [rbp-70h]
  unsigned int *v22; // [rsp+98h] [rbp-68h]
  int v23; // [rsp+A0h] [rbp-60h]
  __int64 v24; // [rsp+A8h] [rbp-58h]
  int v25; // [rsp+B0h] [rbp-50h]
  const wchar_t *v26; // [rsp+B8h] [rbp-48h]
  int *v27; // [rsp+C0h] [rbp-40h]
  int v28; // [rsp+C8h] [rbp-38h]
  int *v29; // [rsp+D0h] [rbp-30h]
  int v30; // [rsp+D8h] [rbp-28h]
  __int64 v31; // [rsp+E0h] [rbp-20h]
  int v32; // [rsp+E8h] [rbp-18h]
  const wchar_t *v33; // [rsp+F0h] [rbp-10h]
  unsigned int *v34; // [rsp+F8h] [rbp-8h]
  int v35; // [rsp+100h] [rbp+0h]
  unsigned int *v36; // [rsp+108h] [rbp+8h]
  int v37; // [rsp+110h] [rbp+10h]
  __int64 v38; // [rsp+118h] [rbp+18h]
  int v39; // [rsp+120h] [rbp+20h]
  const wchar_t *v40; // [rsp+128h] [rbp+28h]
  int *v41; // [rsp+130h] [rbp+30h]
  int v42; // [rsp+138h] [rbp+38h]
  int *v43; // [rsp+140h] [rbp+40h]
  int v44; // [rsp+148h] [rbp+48h]
  __int64 v45; // [rsp+150h] [rbp+50h]
  int v46; // [rsp+158h] [rbp+58h]
  const wchar_t *v47; // [rsp+160h] [rbp+60h]
  int *v48; // [rsp+168h] [rbp+68h]
  int v49; // [rsp+170h] [rbp+70h]
  int *v50; // [rsp+178h] [rbp+78h]
  int v51; // [rsp+180h] [rbp+80h]
  __int64 v52; // [rsp+188h] [rbp+88h]
  int v53; // [rsp+190h] [rbp+90h]
  const wchar_t *v54; // [rsp+198h] [rbp+98h]
  int *v55; // [rsp+1A0h] [rbp+A0h]
  int v56; // [rsp+1A8h] [rbp+A8h]
  int *v57; // [rsp+1B0h] [rbp+B0h]
  int v58; // [rsp+1B8h] [rbp+B8h]
  __int64 v59; // [rsp+1C0h] [rbp+C0h]
  int v60; // [rsp+1C8h] [rbp+C8h]
  const wchar_t *v61; // [rsp+1D0h] [rbp+D0h]
  int *v62; // [rsp+1D8h] [rbp+D8h]
  int v63; // [rsp+1E0h] [rbp+E0h]
  int *v64; // [rsp+1E8h] [rbp+E8h]
  int v65; // [rsp+1F0h] [rbp+F0h]
  __int128 v66; // [rsp+1F8h] [rbp+F8h]
  __int128 v67; // [rsp+208h] [rbp+108h]
  __int128 v68; // [rsp+218h] [rbp+118h]
  __int64 v69; // [rsp+228h] [rbp+128h]

  v0 = 0;
  v19 = L"ForceEnableIommu";
  v6 = 0;
  v3 = 0;
  v20 = &v3;
  v12 = 0;
  v22 = &v6;
  v26 = L"EnablePageTracking";
  v27 = &v8;
  v29 = &v12;
  v33 = L"LogicalAddressMode";
  v34 = &v4;
  v36 = &v5;
  v40 = L"PreferHighLogicalAddresses";
  v41 = &v10;
  v43 = &v13;
  v47 = L"DebugMode";
  v48 = &v11;
  v50 = &v14;
  v54 = L"IdentityMappedPassthrough";
  v55 = &v7;
  v57 = &v15;
  v8 = 0;
  v5 = 0;
  v4 = 0;
  v13 = 0;
  v10 = 0;
  v16 = 0;
  v9 = 0;
  v14 = 0;
  v11 = 0;
  v15 = 0;
  v7 = 0;
  v17 = 0LL;
  v18 = 288;
  v21 = 67108868;
  v23 = 4;
  v24 = 0LL;
  v25 = 288;
  v28 = 67108868;
  v30 = 4;
  v31 = 0LL;
  v32 = 288;
  v35 = 67108868;
  v37 = 4;
  v38 = 0LL;
  v39 = 288;
  v42 = 67108868;
  v44 = 4;
  v45 = 0LL;
  v46 = 288;
  v49 = 67108868;
  v51 = 4;
  v52 = 0LL;
  v53 = 288;
  v56 = 67108868;
  v58 = 4;
  v59 = 0LL;
  v60 = 288;
  v61 = L"ForceDmaRemapping";
  v63 = 67108868;
  v65 = 4;
  v62 = &v9;
  v64 = &v16;
  v69 = 0LL;
  v66 = 0LL;
  v67 = 0LL;
  v68 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\Smm", &v17);
  if ( v4 >= 3 )
    v4 = v5;
  if ( v3 >= 3 )
    v3 = v6;
  if ( !(unsigned __int8)HviIsHypervisorMicrosoftCompatible(v1) && SmmGetIommuInterfaceVersion() >= 2 )
    v0 = v7;
  result = v0 != 0 ? 0x400 : 0;
  dword_1C0154380 = result | (v8 != 0 ? 4 : 0) | v3 & 3 | dword_1C0154380 & 0xFFFFFB00 | (unsigned __int8)(8 * (v4 & 3 | (4 * (v11 & 1 | (2 * (v10 & 1 | (2 * (v9 & 1))))))));
  return result;
}

__int64 __fastcall VIDPN_MGR::_ReadConfiguration(VIDPN_MGR *this)
{
  int RegistryValues; // eax
  int v3; // ebx
  unsigned int v4; // ecx
  struct DXGADAPTER *ContainingAdapter; // rax
  struct DXGADAPTER *v6; // rax
  struct DXGADAPTER *v7; // rax
  struct DXGADAPTER *v8; // rax
  bool v9; // al
  _DWORD *v10; // rbx
  __int64 v12; // [rsp+28h] [rbp-E0h]
  __int64 v13; // [rsp+28h] [rbp-E0h]
  __int64 v14; // [rsp+28h] [rbp-E0h]
  __int64 v15; // [rsp+28h] [rbp-E0h]
  unsigned int v16; // [rsp+38h] [rbp-D0h] BYREF
  int v17; // [rsp+3Ch] [rbp-CCh] BYREF
  int v18; // [rsp+40h] [rbp-C8h] BYREF
  int v19; // [rsp+44h] [rbp-C4h] BYREF
  int v20; // [rsp+48h] [rbp-C0h] BYREF
  int v21; // [rsp+4Ch] [rbp-BCh] BYREF
  __int64 v22; // [rsp+50h] [rbp-B8h] BYREF
  __int64 v23; // [rsp+58h] [rbp-B0h] BYREF
  __int64 v24; // [rsp+60h] [rbp-A8h]
  const wchar_t *v25; // [rsp+68h] [rbp-A0h]
  unsigned int *v26; // [rsp+70h] [rbp-98h]
  __int64 v27; // [rsp+78h] [rbp-90h]
  unsigned int *v28; // [rsp+80h] [rbp-88h]
  int v29; // [rsp+88h] [rbp-80h]
  __int64 v30; // [rsp+90h] [rbp-78h]
  int v31; // [rsp+98h] [rbp-70h]
  const wchar_t *v32; // [rsp+A0h] [rbp-68h]
  char *v33; // [rsp+A8h] [rbp-60h]
  int v34; // [rsp+B0h] [rbp-58h]
  char *v35; // [rsp+B8h] [rbp-50h]
  int v36; // [rsp+C0h] [rbp-48h]
  __int64 v37; // [rsp+C8h] [rbp-40h]
  int v38; // [rsp+D0h] [rbp-38h]
  const wchar_t *v39; // [rsp+D8h] [rbp-30h]
  __int64 *v40; // [rsp+E0h] [rbp-28h]
  int v41; // [rsp+E8h] [rbp-20h]
  __int64 *v42; // [rsp+F0h] [rbp-18h]
  int v43; // [rsp+F8h] [rbp-10h]
  __int64 v44; // [rsp+100h] [rbp-8h]
  int v45; // [rsp+108h] [rbp+0h]
  __int64 v46; // [rsp+110h] [rbp+8h]
  __int128 v47; // [rsp+118h] [rbp+10h]
  __int128 v48; // [rsp+128h] [rbp+20h]
  _QWORD v49[22]; // [rsp+138h] [rbp+30h] BYREF

  if ( !VIDPN_MGR::_BadMonitorSourceModeDiagnosibility )
  {
    v17 = 2;
    memset(v49, 0, 0xA8uLL);
    LODWORD(v49[1]) = 288;
    LODWORD(v49[4]) = 0x4000000;
    v49[2] = L"BadMonitorModeDiag";
    LODWORD(v49[11]) = 0x4000000;
    v49[3] = &v17;
    v49[5] = 0LL;
    v49[9] = L"AssertOnDdiViolation";
    LODWORD(v49[6]) = 0;
    v49[10] = &g_DmmAssertOnDdiViolation;
    v49[7] = 0LL;
    LODWORD(v49[8]) = 288;
    v49[12] = 0LL;
    LODWORD(v49[13]) = 0;
    HIDWORD(v12) = 0;
    RegistryValues = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\DMM", v49);
    v3 = RegistryValues;
    if ( RegistryValues >= 0 )
    {
      v4 = v17;
    }
    else
    {
      WdLogSingleEntry1(7LL, RegistryValues);
      WdLogGlobalForLineNumber = 689;
      if ( v3 != -1073741772 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 692;
      }
      v4 = 2;
      v17 = 2;
    }
    if ( v4 == 1 || v4 == 2 )
    {
      VIDPN_MGR::_BadMonitorSourceModeDiagnosibility = v4;
    }
    else
    {
      WdLogSingleEntry1(2LL, v4);
      WdLogGlobalForLineNumber = 717;
    }
  }
  v18 = 0;
  ContainingAdapter = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v12) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)ContainingAdapter + 27), L"AllowUnspecifiedVSync", &v18, 4LL, v12) >= 0 )
  {
    VIDPN_MGR::_bAllowUnspecifiedVSync = v18 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    WdLogGlobalForLineNumber = 738;
  }
  v19 = 0;
  v6 = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v13) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)v6 + 27), L"AllowUnspecifiedHSync", &v19, 4LL, v13) >= 0 )
  {
    VIDPN_MGR::_bAllowUnspecifiedHSync = v19 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    WdLogGlobalForLineNumber = 761;
  }
  v20 = 0;
  v7 = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v14) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)v7 + 27), L"AllowUnspecifiedPixelRate", &v20, 4LL, v14) >= 0 )
  {
    VIDPN_MGR::_bAllowUnspecifiedPixelRate = v20 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    WdLogGlobalForLineNumber = 784;
  }
  v21 = 0;
  v8 = VIDPN_MGR::GetContainingAdapter(this);
  LODWORD(v15) = 2;
  if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)v8 + 27), L"ForceDualViewBehavior", &v21, 4LL, v15) >= 0 )
  {
    v9 = v21 != 0;
  }
  else
  {
    WdLogSingleEntry0(7LL);
    v9 = 0;
    WdLogGlobalForLineNumber = 808;
  }
  *((_BYTE *)this + 520) = v9;
  v10 = (_DWORD *)((char *)this + 544);
  v16 = 1000;
  LODWORD(v27) = 67108868;
  v34 = 67108868;
  v25 = L"RapidHPDTime";
  v41 = 67108868;
  v26 = &v16;
  *((_DWORD *)this + 136) = 5;
  v28 = &v16;
  LODWORD(v22) = 0;
  v32 = L"RapidHPDThresholdCount";
  v23 = 0LL;
  v39 = L"EnableExperimentalRefreshRates";
  v40 = &v22;
  v42 = &v22;
  LODWORD(v24) = 288;
  v29 = 4;
  v30 = 0LL;
  v31 = 288;
  v33 = (char *)this + 544;
  v35 = (char *)this + 544;
  v36 = 4;
  v37 = 0LL;
  v38 = 288;
  v43 = 4;
  v44 = 0LL;
  v45 = 0;
  v46 = 0LL;
  v47 = 0LL;
  v48 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v23);
  if ( v16 > 0xEA60 )
    v16 = 60000;
  *((_DWORD *)this + 135) = 10000 * v16 / KeQueryTimeIncrement();
  if ( *v10 == 1 )
  {
    *v10 = 0;
  }
  else if ( *v10 > 0x20u )
  {
    *v10 = 32;
  }
  return 0LL;
}

void TdrInit(void)
{
  int v0; // eax
  int v1; // eax
  int v2; // eax
  int v3; // eax
  int v4; // eax
  int v5; // eax
  int v6; // eax
  __int64 v7; // rcx
  unsigned int v8; // [rsp+38h] [rbp-D0h] BYREF
  unsigned int v9; // [rsp+3Ch] [rbp-CCh] BYREF
  unsigned int v10; // [rsp+40h] [rbp-C8h] BYREF
  unsigned int v11; // [rsp+44h] [rbp-C4h] BYREF
  unsigned int v12; // [rsp+48h] [rbp-C0h] BYREF
  unsigned int v13; // [rsp+4Ch] [rbp-BCh] BYREF
  unsigned int v14; // [rsp+50h] [rbp-B8h] BYREF
  unsigned int v15; // [rsp+54h] [rbp-B4h] BYREF
  int v16; // [rsp+58h] [rbp-B0h] BYREF
  int v17; // [rsp+5Ch] [rbp-ACh] BYREF
  int v18; // [rsp+60h] [rbp-A8h] BYREF
  int v19; // [rsp+64h] [rbp-A4h] BYREF
  int v20; // [rsp+68h] [rbp-A0h] BYREF
  int v21; // [rsp+6Ch] [rbp-9Ch] BYREF
  __int64 v22; // [rsp+70h] [rbp-98h] BYREF
  __int64 v23; // [rsp+78h] [rbp-90h] BYREF
  int v24; // [rsp+80h] [rbp-88h]
  const wchar_t *v25; // [rsp+88h] [rbp-80h]
  unsigned int *v26; // [rsp+90h] [rbp-78h]
  int v27; // [rsp+98h] [rbp-70h]
  int *v28; // [rsp+A0h] [rbp-68h]
  int v29; // [rsp+A8h] [rbp-60h]
  __int64 v30; // [rsp+B0h] [rbp-58h]
  int v31; // [rsp+B8h] [rbp-50h]
  const wchar_t *v32; // [rsp+C0h] [rbp-48h]
  unsigned int *v33; // [rsp+C8h] [rbp-40h]
  int v34; // [rsp+D0h] [rbp-38h]
  int *v35; // [rsp+D8h] [rbp-30h]
  int v36; // [rsp+E0h] [rbp-28h]
  __int64 v37; // [rsp+E8h] [rbp-20h]
  int v38; // [rsp+F0h] [rbp-18h]
  const wchar_t *v39; // [rsp+F8h] [rbp-10h]
  unsigned int *v40; // [rsp+100h] [rbp-8h]
  int v41; // [rsp+108h] [rbp+0h]
  int *v42; // [rsp+110h] [rbp+8h]
  int v43; // [rsp+118h] [rbp+10h]
  __int64 v44; // [rsp+120h] [rbp+18h]
  int v45; // [rsp+128h] [rbp+20h]
  const wchar_t *v46; // [rsp+130h] [rbp+28h]
  unsigned int *v47; // [rsp+138h] [rbp+30h]
  int v48; // [rsp+140h] [rbp+38h]
  int *v49; // [rsp+148h] [rbp+40h]
  int v50; // [rsp+150h] [rbp+48h]
  __int64 v51; // [rsp+158h] [rbp+50h]
  int v52; // [rsp+160h] [rbp+58h]
  const wchar_t *v53; // [rsp+168h] [rbp+60h]
  unsigned int *v54; // [rsp+170h] [rbp+68h]
  int v55; // [rsp+178h] [rbp+70h]
  int *v56; // [rsp+180h] [rbp+78h]
  int v57; // [rsp+188h] [rbp+80h]
  __int64 v58; // [rsp+190h] [rbp+88h]
  int v59; // [rsp+198h] [rbp+90h]
  const wchar_t *v60; // [rsp+1A0h] [rbp+98h]
  unsigned int *v61; // [rsp+1A8h] [rbp+A0h]
  int v62; // [rsp+1B0h] [rbp+A8h]
  int *v63; // [rsp+1B8h] [rbp+B0h]
  int v64; // [rsp+1C0h] [rbp+B8h]
  __int64 v65; // [rsp+1C8h] [rbp+C0h]
  int v66; // [rsp+1D0h] [rbp+C8h]
  const wchar_t *v67; // [rsp+1D8h] [rbp+D0h]
  unsigned int *v68; // [rsp+1E0h] [rbp+D8h]
  int v69; // [rsp+1E8h] [rbp+E0h]
  __int64 *v70; // [rsp+1F0h] [rbp+E8h]
  int v71; // [rsp+1F8h] [rbp+F0h]
  __int64 v72; // [rsp+200h] [rbp+F8h]
  int v73; // [rsp+208h] [rbp+100h]
  const wchar_t *v74; // [rsp+210h] [rbp+108h]
  unsigned int *v75; // [rsp+218h] [rbp+110h]
  int v76; // [rsp+220h] [rbp+118h]
  char *v77; // [rsp+228h] [rbp+120h]
  int v78; // [rsp+230h] [rbp+128h]
  __int64 v79; // [rsp+238h] [rbp+130h]
  int v80; // [rsp+240h] [rbp+138h]
  __int64 v81; // [rsp+248h] [rbp+140h]
  __int128 v82; // [rsp+250h] [rbp+148h]
  __int128 v83; // [rsp+260h] [rbp+158h]

  v22 = 0x20000003CLL;
  v13 = 0;
  v8 = 0;
  v9 = 0;
  v16 = 3;
  v25 = L"TdrLevel";
  v26 = &v13;
  v28 = &v16;
  v10 = 0;
  v17 = 2;
  v32 = L"TdrDelay";
  v18 = 2;
  v33 = &v8;
  v35 = &v17;
  v39 = L"TdrDodPresentDelay";
  v40 = &v9;
  v42 = &v18;
  v46 = L"TdrDodVSyncDelay";
  v47 = &v10;
  v49 = &v19;
  v53 = L"TdrDdiDelay";
  v54 = &v11;
  v56 = &v20;
  v60 = L"TdrLimitCount";
  v61 = &v14;
  v19 = 2;
  v20 = 5;
  v11 = 0;
  v12 = 0;
  v21 = 5;
  v14 = 0;
  v15 = 0;
  v23 = 0LL;
  v24 = 288;
  v27 = 67108868;
  v29 = 4;
  v30 = 0LL;
  v31 = 288;
  v34 = 67108868;
  v36 = 4;
  v37 = 0LL;
  v38 = 288;
  v41 = 67108868;
  v43 = 4;
  v44 = 0LL;
  v45 = 288;
  v48 = 67108868;
  v50 = 4;
  v51 = 0LL;
  v52 = 288;
  v55 = 67108868;
  v57 = 4;
  v58 = 0LL;
  v59 = 288;
  v62 = 67108868;
  v63 = &v21;
  v64 = 4;
  v67 = L"TdrLimitTime";
  v66 = 288;
  v68 = &v15;
  v70 = &v22;
  v74 = L"TdrDebugMode";
  v75 = &v12;
  v69 = 67108868;
  v71 = 4;
  v73 = 288;
  v76 = 67108868;
  v78 = 4;
  v77 = (char *)&v22 + 4;
  v65 = 0LL;
  v72 = 0LL;
  v79 = 0LL;
  v80 = 0;
  v81 = 0LL;
  v82 = 0LL;
  v83 = 0LL;
  v0 = RtlQueryRegistryValuesEx(0LL, L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\GraphicsDrivers", &v23);
  if ( v0 < 0 )
  {
    v13 = 3;
    v8 = 2;
    v9 = 2;
    v10 = 2;
    v11 = 5;
    v12 = 2;
    WdLogSingleEntry1(3LL, v0);
    WdLogGlobalForLineNumber = 2211;
  }
  if ( v13 < 2 || v13 == 3 )
  {
    g_TdrConfig = v13;
  }
  else
  {
    g_TdrConfig = 3;
    WdLogSingleEntry2(3LL, v13, 3LL);
    WdLogGlobalForLineNumber = 2238;
  }
  v1 = v8;
  if ( v8 )
  {
    if ( v8 > 0x384 )
      v1 = 900;
    dword_1C01537DC = v1;
  }
  else
  {
    dword_1C01537DC = 1;
  }
  if ( dword_1C01537DC != v8 )
  {
    WdLogSingleEntry2(3LL, v8, (unsigned int)dword_1C01537DC);
    WdLogGlobalForLineNumber = 2262;
  }
  v2 = v9;
  if ( v9 )
  {
    if ( v9 > 0x384 )
      v2 = 900;
    dword_1C01537E0 = v2;
  }
  else
  {
    dword_1C01537E0 = 1;
  }
  if ( dword_1C01537E0 != v9 )
  {
    WdLogSingleEntry2(3LL, v9, (unsigned int)dword_1C01537E0);
    WdLogGlobalForLineNumber = 2287;
  }
  v3 = v10;
  if ( v10 )
  {
    if ( v10 > 0x384 )
      v3 = 900;
    dword_1C01537E4 = v3;
  }
  else
  {
    dword_1C01537E4 = 1;
  }
  if ( dword_1C01537E4 != v10 )
  {
    WdLogSingleEntry2(3LL, v10, (unsigned int)dword_1C01537E4);
    WdLogGlobalForLineNumber = 2312;
  }
  v4 = v11;
  if ( v11 )
  {
    if ( v11 > 0x384 )
      v4 = 900;
    dword_1C01537E8 = v4;
  }
  else
  {
    dword_1C01537E8 = 1;
  }
  if ( dword_1C01537E8 != v11 )
  {
    WdLogSingleEntry2(3LL, v11, (unsigned int)dword_1C01537E8);
    WdLogGlobalForLineNumber = 2337;
  }
  v5 = v14;
  if ( v14 <= 0x20 )
  {
    if ( !v14 )
      v5 = 1;
    dword_1C01537F0 = v5;
  }
  else
  {
    dword_1C01537F0 = 32;
  }
  if ( dword_1C01537F0 != v14 )
  {
    WdLogSingleEntry2(3LL, v14, (unsigned int)dword_1C01537F0);
    WdLogGlobalForLineNumber = 2362;
  }
  v6 = v15;
  v7 = 3600LL;
  if ( v15 <= 0xE10 )
  {
    if ( v15 < 5 )
      v6 = 5;
    dword_1C01537F4 = v6;
  }
  else
  {
    dword_1C01537F4 = 3600;
  }
  if ( dword_1C01537F4 != v15 )
  {
    WdLogSingleEntry2(3LL, v15, (unsigned int)dword_1C01537F4);
    WdLogGlobalForLineNumber = 2387;
  }
  LOBYTE(v7) = 1;
  byte_1C01537EC = (unsigned __int8)WdIsDebuggerPresent(v7) != 0;
  if ( v12 < 2 || v12 - 2 < 2 )
    g_TdrDebugMode = v12;
  else
    g_TdrDebugMode = 2;
  if ( g_TdrDebugMode != v12 )
  {
    WdLogSingleEntry2(3LL, v12, g_TdrDebugMode);
    WdLogGlobalForLineNumber = 2418;
  }
  TdrHistoryInit(&g_TdrHistory);
}

__int64 __fastcall CCD_BTL::SetUnsupportedMonitorModesFlag(CCD_BTL *this, unsigned __int8 a2)
{
  bool v3; // zf
  int v5; // esi
  NTSTATUS v6; // eax
  __int64 v7; // rdi
  const wchar_t *v8; // r9
  struct _UNICODE_STRING DestinationString; // [rsp+50h] [rbp+7h] BYREF
  struct _OBJECT_ATTRIBUTES ObjectAttributes; // [rsp+60h] [rbp+17h] BYREF
  void *KeyHandle; // [rsp+B0h] [rbp+67h] BYREF
  int Data; // [rsp+B8h] [rbp+6Fh] BYREF

  v3 = *((_BYTE *)this + 152) == 0;
  KeyHandle = 0LL;
  v5 = a2;
  DxgkLogCodePointPacket(41LL, !v3, a2, 0LL, 0LL);
  *((_BYTE *)this + 152) = a2;
  *(&ObjectAttributes.Length + 1) = 0;
  *(&ObjectAttributes.Attributes + 1) = 0;
  DestinationString = 0LL;
  KeyHandle = 0LL;
  RtlInitUnicodeString(&DestinationString, L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\GraphicsDrivers");
  ObjectAttributes.Length = 48;
  ObjectAttributes.ObjectName = &DestinationString;
  ObjectAttributes.RootDirectory = 0LL;
  ObjectAttributes.Attributes = 576;
  *(_OWORD *)&ObjectAttributes.SecurityDescriptor = 0LL;
  v6 = ZwOpenKey(&KeyHandle, 0x40000000u, &ObjectAttributes);
  v7 = v6;
  if ( v6 < 0 )
  {
    WdLogSingleEntry1(2LL, v6);
    v8 = L"Cannot open handle to GraphicsDrivers register key (NtStatus = 0x%I64x)";
    WdLogGlobalForLineNumber = 1458;
LABEL_5:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v8, v7, 0LL, 0LL, 0LL, 0LL);
    return (unsigned int)v7;
  }
  Data = v5;
  RtlInitUnicodeString(&DestinationString, L"UnsupportedMonitorModesAllowed");
  v7 = ZwSetValueKey(KeyHandle, &DestinationString, 0, 4u, &Data, 4u);
  ZwClose(KeyHandle);
  if ( (int)v7 < 0 )
  {
    WdLogSingleEntry1(2LL, v7);
    v8 = L"Failed to write unsupported monitor modes flag to registry. (NtStatus = 0x%I64x)";
    WdLogGlobalForLineNumber = 1478;
    goto LABEL_5;
  }
  return (unsigned int)v7;
}

__int64 __fastcall ADAPTER_DISPLAY::Initialize(ADAPTER_DISPLAY *this)
{
  int *v1; // rdi
  __int64 v3; // rcx
  __int64 v4; // rdx
  unsigned int v5; // eax
  __int64 v6; // rbx
  __int64 v7; // rax
  unsigned __int64 v8; // kr00_8
  bool v9; // cf
  __int64 v10; // rax
  _QWORD *v11; // rax
  __int64 v12; // rcx
  _QWORD *v13; // rdi
  DISPLAY_SOURCE *i; // r14
  unsigned int j; // ebx
  MONITOR_MGR **v16; // r14
  MONITOR_MGR *v17; // rax
  MONITOR_MGR *v18; // rax
  MONITOR_MGR *v19; // rdi
  unsigned int v20; // ebx
  __int64 result; // rax
  unsigned int *v22; // r15
  int RegistryValues; // eax
  int v24; // r14d
  int v25; // eax
  unsigned int v26; // eax
  int v27; // ecx
  __int64 v28; // rcx
  int v29; // edx
  __int64 v30; // rax
  __int64 v31; // rcx
  __int64 v32; // rdx
  __int64 v33; // rax
  bool v34; // al
  __int64 v35; // rcx
  __int64 v36; // rax
  __int64 v37; // rcx
  bool v38; // sf
  bool v39; // of
  __int64 v40; // rcx
  int v41; // r12d
  int v42; // ebx
  struct DXGGLOBAL *v43; // rax
  int v44; // eax
  __int64 v45; // rcx
  __int64 v46; // rax
  int v47; // ecx
  struct _LUID v48; // rcx
  __int64 v49; // rax
  DXGGLOBAL *Global; // rax
  __int64 v51; // rax
  __int64 v52; // rcx
  bool v53; // zf
  __int64 v54; // rcx
  _DWORD *v55; // rcx
  __int64 v56; // rax
  struct DXGGLOBAL *v57; // rax
  struct DXGDODPRESENT *DodPresent; // rax
  __int64 v59; // rcx
  int (__fastcall *v60)(_QWORD, __int128 *); // rax
  __int64 v61; // rcx
  _DWORD *v62; // rdx
  int v63; // eax
  __int64 v64; // rcx
  unsigned int k; // r10d
  __int64 v66; // rax
  struct _KEVENT *v67; // rax
  struct OUTPUTDUPL_MGR **v68; // [rsp+28h] [rbp-E0h]
  __int64 v69; // [rsp+28h] [rbp-E0h]
  __int64 v70; // [rsp+28h] [rbp-E0h]
  __int64 v71; // [rsp+28h] [rbp-E0h]
  __int64 v72; // [rsp+28h] [rbp-E0h]
  __int64 v73; // [rsp+28h] [rbp-E0h]
  __int64 v74; // [rsp+28h] [rbp-E0h]
  __int64 v75; // [rsp+30h] [rbp-D8h]
  __int64 v76; // [rsp+30h] [rbp-D8h]
  __int64 v77; // [rsp+30h] [rbp-D8h]
  __int64 v78; // [rsp+30h] [rbp-D8h]
  __int64 v79; // [rsp+38h] [rbp-D0h]
  __int64 v80; // [rsp+38h] [rbp-D0h]
  int v81; // [rsp+58h] [rbp-B0h] BYREF
  int v82; // [rsp+5Ch] [rbp-ACh] BYREF
  unsigned int v83; // [rsp+60h] [rbp-A8h] BYREF
  int v84; // [rsp+64h] [rbp-A4h] BYREF
  void *EventHandle; // [rsp+68h] [rbp-A0h] BYREF
  struct _LUID v86; // [rsp+70h] [rbp-98h] BYREF
  struct _LUID v87; // [rsp+78h] [rbp-90h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v88; // [rsp+80h] [rbp-88h] BYREF
  __int128 v89; // [rsp+B0h] [rbp-58h] BYREF
  __int64 v90; // [rsp+C0h] [rbp-48h]
  _QWORD v91[50]; // [rsp+C8h] [rbp-40h] BYREF

  v1 = (int *)((char *)this + 24);
  *((_DWORD *)this + 6) = 0;
  v3 = *((_QWORD *)this + 2);
  v4 = v3;
  if ( *(_DWORD *)(v3 + 2280) >= 0x5010u && !*(_BYTE *)(v3 + 209) && (*(_DWORD *)(v3 + 2976) & 8) == 0 )
  {
    *(_QWORD *)&v88.Type = 16LL;
    *(_QWORD *)&v88.InputDataSize = 0LL;
    *(_QWORD *)&v88.Flags.0 = 0LL;
    HIDWORD(v88.hKmdProcessHandle) = 0;
    v88.pInputData = 0LL;
    v88.pOutputData = v1;
    v88.OutputDataSize = 4;
    v44 = DXGADAPTER::DdiQueryAdapterInfo((DXGADAPTER *)v3, &v88);
    if ( v44 < 0 )
    {
      *(_QWORD *)(WdLogNewEntry5_WdTrace(v45) + 24) = v44;
      *v1 = 0;
      v46 = *((_QWORD *)this + 2);
      WdLogGlobalForLineNumber = 4707;
      if ( *(int *)(v46 + 2736) >= 8704 )
        *v1 |= 2u;
    }
    v4 = *((_QWORD *)this + 2);
    v47 = *v1;
    if ( *(int *)(v4 + 2736) >= 9472 )
    {
      if ( (v47 & 0xC) == 0xC )
      {
        WdLogSingleEntry1(2LL, this);
        WdLogGlobalForLineNumber = 4736;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Adapter 0x%I64x: Both HdrFP16ScanoutSupport and HdrARGB10ScanoutSupport can't be set to 1 at the same time",
          (__int64)this,
          0LL,
          0LL,
          0LL,
          0LL);
        return 3221225485LL;
      }
    }
    else
    {
      v47 &= 0xFFFFFFF3;
      *v1 = v47;
    }
    if ( *(int *)(v4 + 2736) < 9984 )
    {
      v47 &= ~0x10u;
      *v1 = v47;
    }
    if ( *(int *)(v4 + 2736) < 10496 || *(_QWORD *)(v4 + 832) || !*(_DWORD *)(v4 + 1856) || (v47 & 2) == 0 )
    {
      v47 &= ~0x20u;
      *v1 = v47;
    }
    if ( *(int *)(v4 + 2736) < 12288 )
    {
      v47 &= ~0x40u;
      *v1 = v47;
    }
    if ( g_bDbgForceUsb4MonitorSupport )
      *v1 = v47 | 0x40;
  }
  v5 = *(_DWORD *)(v4 + 1856);
  *((_DWORD *)this + 24) = v5;
  v6 = v5;
  v8 = v5;
  v7 = 4000LL * v5;
  if ( !is_mul_ok(v8, 0xFA0uLL) )
    v7 = -1LL;
  v9 = __CFADD__(v7, 8LL);
  v10 = v7 + 8;
  if ( v9 )
    v10 = -1LL;
  v11 = (_QWORD *)operator new[](v10, 1265072196LL, 64LL);
  if ( v11 )
  {
    *v11 = v6;
    v13 = v11 + 1;
    for ( i = (DISPLAY_SOURCE *)(v11 + 1); v6; --v6 )
    {
      DISPLAY_SOURCE::DISPLAY_SOURCE(i);
      i = (DISPLAY_SOURCE *)((char *)i + 4000);
    }
  }
  else
  {
    v13 = 0LL;
  }
  *((_QWORD *)this + 16) = v13;
  if ( !v13 )
  {
    WdLogSingleEntry3(6LL, *((unsigned int *)this + 24), *((_QWORD *)this + 2), -1073741801LL);
    v76 = *((_QWORD *)this + 2);
    v71 = *((unsigned int *)this + 24);
    WdLogGlobalForLineNumber = 4792;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate 0x%I64x of display sources for adapter 0x%I64x, returning 0x%I64x",
      v71,
      v76,
      -1073741801LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  for ( j = 0; j < *((_DWORD *)this + 24); ++j )
  {
    result = DISPLAY_SOURCE::Initialize((DISPLAY_SOURCE *)(*((_QWORD *)this + 16) + 4000LL * j), this, j);
    if ( (int)result < 0 )
      return result;
  }
  v16 = (MONITOR_MGR **)((char *)this + 112);
  *(_QWORD *)(WdLogNewEntry5_WdTrace(v12) + 24) = this;
  WdLogGlobalForLineNumber = 253;
  if ( this == (ADAPTER_DISPLAY *)-112LL )
  {
    WdLogSingleEntry2(2LL, -112LL, 0LL);
    WdLogGlobalForLineNumber = 265;
    return (unsigned int)-1073741811;
  }
  *v16 = 0LL;
  v17 = (MONITOR_MGR *)operator new(688LL, 1298626628LL);
  if ( !v17 || (v18 = MONITOR_MGR::MONITOR_MGR(v17, this), (v19 = v18) == 0LL) )
  {
    WdLogSingleEntry1(2LL, *((_QWORD *)this + 2));
    WdLogGlobalForLineNumber = 285;
    return (unsigned int)-1073741811;
  }
  v20 = MONITOR_MGR::_InitializeMonitorManager(v18);
  if ( (v20 & 0x80000000) != 0 )
  {
    MONITOR_MGR::`vector deleting destructor(v19, 1u);
    return v20;
  }
  *v16 = v19;
  result = VIDPN_MGR_CLASSFACTORY::CreateVidPnMgr(this, (struct VIDPN_MGR **)this + 13);
  if ( (int)result > -1071774937 && (unsigned int)(result + 1071774934) > 0x3FE1FCD5 )
  {
    if ( (*(_DWORD *)(*((_QWORD *)this + 2) + 444LL) & 0x100) != 0 )
    {
      v48 = (struct _LUID)*((_QWORD *)DXGGLOBAL::GetGlobal() + 123);
      v49 = *((_QWORD *)this + 2);
      v87 = v48;
      v86 = *(struct _LUID *)(v49 + 412);
      result = CreateOutputDuplManager(*((_DWORD *)this + 24), 0LL, &v87, &v86, (struct OUTPUTDUPL_MGR **)this + 15);
      if ( (int)result < 0 )
        return result;
      Global = DXGGLOBAL::GetGlobal();
      DXGGLOBAL::AddIndirectOutputDuplMgr(
        Global,
        (struct OUTPUTDUPL_MGR_INDIRECT *)((*((_QWORD *)this + 15) - 24LL) & -(__int64)(*((_QWORD *)this + 15) != 0LL)));
    }
    else
    {
      result = CreateOutputDuplManager(*((_DWORD *)this + 24), this, 0LL, 0LL, (struct OUTPUTDUPL_MGR **)this + 15);
      if ( (int)result < 0 )
        return result;
    }
    v81 = 1;
    *((_QWORD *)this + 76) = (char *)this + 600;
    *((_QWORD *)this + 75) = (char *)this + 600;
    v22 = (unsigned int *)((char *)this + 528);
    *((_DWORD *)this + 130) = 0;
    *((_DWORD *)this + 132) = 1000;
    *((_DWORD *)this + 131) = 200;
    *((_DWORD *)this + 133) = 20000000;
    *((_DWORD *)this + 134) = 0;
    memset(v91, 0, 0x188uLL);
    v91[5] = 0LL;
    LODWORD(v91[4]) = 0x4000000;
    LODWORD(v91[1]) = 288;
    v91[2] = L"ModeListCaching";
    LODWORD(v91[8]) = 288;
    v91[3] = &v81;
    LODWORD(v91[11]) = 0x4000000;
    v91[9] = L"SetTimingsFlags";
    v91[16] = L"ShortLinkTrainingTimeout";
    v91[23] = L"LongLinkTrainingTimeout";
    v91[30] = L"HPDFilterLimit";
    LODWORD(v91[15]) = 288;
    LODWORD(v91[18]) = 0x4000000;
    LODWORD(v91[22]) = 288;
    LODWORD(v91[25]) = 0x4000000;
    LODWORD(v91[29]) = 288;
    LODWORD(v91[32]) = 0x4000000;
    LODWORD(v91[36]) = 288;
    LODWORD(v91[39]) = 0x4000000;
    v91[37] = L"EnableVirtualRefreshRateOnExternalMonitor";
    LODWORD(v91[6]) = 0;
    v91[7] = 0LL;
    v91[10] = (char *)this + 520;
    v91[12] = 0LL;
    LODWORD(v91[13]) = 0;
    v91[14] = 0LL;
    v91[17] = (char *)this + 524;
    v91[19] = 0LL;
    LODWORD(v91[20]) = 0;
    v91[21] = 0LL;
    v91[24] = (char *)this + 528;
    v91[26] = 0LL;
    LODWORD(v91[27]) = 0;
    v91[28] = 0LL;
    v91[31] = (char *)this + 532;
    v91[33] = 0LL;
    LODWORD(v91[34]) = 0;
    v91[35] = 0LL;
    v91[38] = (char *)this + 536;
    v91[40] = 0LL;
    LODWORD(v91[41]) = 0;
    HIDWORD(v68) = 0;
    RegistryValues = RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers\\DMM", v91);
    v24 = RegistryValues;
    if ( RegistryValues < 0 )
    {
      WdLogSingleEntry1(4LL, RegistryValues);
      WdLogGlobalForLineNumber = 4941;
      if ( v24 != -1073741772 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 4944;
        DxgkLogInternalTriageEvent(
          0,
          262146,
          -1,
          (unsigned int)L"Status == STATUS_OBJECT_NAME_NOT_FOUND",
          4944LL,
          0LL,
          0LL,
          0LL,
          0LL);
      }
      *((_DWORD *)this + 131) = 200;
      v25 = 1;
      v81 = 1;
      v24 = 0;
      *((_DWORD *)this + 130) = 0;
      *v22 = 1000;
    }
    else
    {
      v25 = v81;
    }
    *((_BYTE *)this + 292) = v25 == 1;
    v26 = *v22;
    if ( !*v22 || *((_DWORD *)this + 131) >= v26 || v26 >= 0x7530 )
    {
      WdLogSingleEntry3(2LL, *((unsigned int *)this + 131), *((unsigned int *)this + 131), *((_QWORD *)this + 2));
      v79 = *((_QWORD *)this + 2);
      v69 = *((unsigned int *)this + 131);
      WdLogGlobalForLineNumber = 4969;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Invalid link training timeout registry value (0x%I64x, 0x%I64x) on adapter 0x%I64x, fallback to th"
                       "e default value.",
        v69,
        v69,
        v79,
        0LL,
        0LL);
      *((_DWORD *)this + 131) = 200;
      *((_DWORD *)this + 132) = 1000;
    }
    v27 = *((_DWORD *)this + 133);
    if ( (unsigned int)(v27 - 1000000) > 0x5E69EC0 )
    {
      if ( v27 )
      {
        WdLogSingleEntry3(2LL, *((unsigned int *)this + 133), 20000000LL, *((_QWORD *)this + 2));
        v80 = *((_QWORD *)this + 2);
        v72 = *((unsigned int *)this + 133);
        WdLogGlobalForLineNumber = 4984;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Invalid hot-plug filter limit of %#x on adapter 0x%I64x.  Using default of %#x.",
          v72,
          20000000LL,
          v80,
          0LL,
          0LL);
      }
      *((_DWORD *)this + 133) = 20000000;
    }
    if ( (*((_DWORD *)this + 130) & 1) != 0 )
    {
      v51 = *((_QWORD *)this + 2);
      if ( !*(_QWORD *)(v51 + 656) )
      {
        v20 = -1073741735;
        WdLogSingleEntry3(2LL, *(int *)(v51 + 416), *(unsigned int *)(v51 + 412), -1073741735LL);
        v52 = *((_QWORD *)this + 2);
        v77 = *(unsigned int *)(v52 + 412);
        v73 = *(int *)(v52 + 416);
        WdLogGlobalForLineNumber = 5001;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Miniport driver wants t fallback to use DdiCommitVidPn but it does not supply pfnCommitVidPn on "
                         "adapter (0x%I64x%08I64x), returning 0x%I64x.",
          v73,
          v77,
          -1073741735LL,
          0LL,
          0LL);
        return v20;
      }
    }
    v28 = *((_QWORD *)this + 2);
    v29 = *(_DWORD *)(v28 + 420);
    if ( (*(_DWORD *)(v28 + 444) & 0x400) != 0 )
    {
      if ( v29 == 1297040209
        && *(_DWORD *)(*(_QWORD *)(*(_QWORD *)(*(_QWORD *)(v28 + 216) + 64LL) + 40LL) + 28LL) < 0x700Au )
      {
        *((_BYTE *)this + 289) = 1;
        v34 = 1;
      }
      else
      {
        v82 = (*((_DWORD *)this + 6) >> 1) & 1;
        memset(v91, 0, 0x188uLL);
        LODWORD(v91[1]) = 288;
        v91[2] = L"ForceEnableDWMClone";
        LODWORD(v91[4]) = 67108868;
        v91[3] = &v82;
        LODWORD(v91[6]) = 4;
        v91[5] = &v82;
        HIDWORD(v68) = 0;
        RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", v91);
        v53 = v82 == 0;
        *((_BYTE *)this + 289) = v82 != 0;
        v34 = !v53;
      }
    }
    else
    {
      if ( v29 == 1297040209 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 5058;
        DxgkLogInternalTriageEvent(
          0,
          262146,
          -1,
          (unsigned int)L"GetAdapter()->GetAdapterVendorId() != VENDOR_ID_QUALCOMM",
          5058LL,
          0LL,
          0LL,
          0LL,
          0LL);
      }
      v30 = *((_QWORD *)this + 2);
      v31 = *(unsigned int *)(v30 + 412);
      v32 = *(int *)(v30 + 416);
      if ( (*((_DWORD *)this + 6) & 2) != 0 )
      {
        v20 = -1073741735;
        WdLogSingleEntry3(2LL, v32, (unsigned int)v31, -1073741735LL);
        v33 = *((_QWORD *)this + 2);
        v75 = *(unsigned int *)(v33 + 412);
        v70 = *(int *)(v33 + 416);
        WdLogGlobalForLineNumber = 5070;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Force to stop DWM clone supported adapter (0x%I64x%08I64x) due to target ID does not support DWM"
                         " clone, returning 0x%I64x.",
          v70,
          v75,
          -1073741735LL,
          0LL,
          0LL);
        return v20;
      }
      WdLogSingleEntry2(4LL, v32, v31);
      v34 = 0;
      WdLogGlobalForLineNumber = 5078;
      *((_BYTE *)this + 289) = 0;
    }
    *((_BYTE *)this + 290) = v34;
    v35 = *((_QWORD *)this + 2);
    if ( *(int *)(v35 + 3004) < 2000 )
    {
      v54 = *(_QWORD *)(v35 + 216);
      v84 = 0;
      LODWORD(v68) = 2;
      if ( (int)DpiReadPnpRegistryValue(v54, L"EnableVirtualTopologySupport", &v84, 4LL, v68) >= 0 )
      {
        if ( v84 )
        {
          v55 = (_DWORD *)*((_QWORD *)this + 2);
          if ( (v55[111] & 0x800) == 0 )
          {
            v20 = -1073741735;
            WdLogSingleEntry3(2LL, (int)v55[104], (unsigned int)v55[103], -1073741735LL);
            v56 = *((_QWORD *)this + 2);
            v78 = *(unsigned int *)(v56 + 412);
            v74 = *(int *)(v56 + 416);
            WdLogGlobalForLineNumber = 5104;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"Force to stop adapter (0x%I64x%08I64x) due to target ID does not support reduced hash size a"
                             "nd registry requested to use virtual topologies, returning 0x%I64x.",
              v74,
              v78,
              -1073741735LL,
              0LL,
              0LL);
            return v20;
          }
          *((_BYTE *)this + 290) = 1;
          v57 = DXGGLOBAL::GetGlobal();
          DXGADAPTERSOURCEHASH::ForceReducedHashSize((struct DXGGLOBAL *)((char *)v57 + 1384));
        }
      }
    }
    v36 = *((_QWORD *)this + 2);
    if ( !*(_QWORD *)(v36 + 3128) )
    {
      DodPresent = DxgkpCreateDodPresent(this, *(_QWORD *)(v36 + 696) != 0LL);
      v59 = *((_QWORD *)this + 2);
      *((_QWORD *)this + 57) = DodPresent;
      if ( !DodPresent )
        v24 = -1073741801;
      v90 = 0LL;
      v89 = 0LL;
      v60 = *(int (__fastcall **)(_QWORD, __int128 *))(v59 + 2368);
      if ( v60 && v60(*(_QWORD *)(v59 + 2296), &v89) >= 0 )
      {
        v61 = 0LL;
        v62 = (_DWORD *)((char *)this + 432);
        do
        {
          v63 = *((unsigned __int8 *)&v89 + v61++);
          *v62++ = v63;
        }
        while ( v61 < 4 );
        *((_DWORD *)this + 113) = BYTE4(v90);
        *((_DWORD *)this + 112) = BYTE5(v90);
      }
      else
      {
        *((_DWORD *)this + 108) = 1;
      }
      v64 = *(_QWORD *)(*((_QWORD *)this + 2) + 216LL);
      if ( *(_DWORD *)(*(_QWORD *)(*(_QWORD *)(v64 + 64) + 40LL) + 28LL) >= 0x3007u )
        DpiSetSchedulerCallbackState(v64, 3LL);
    }
    if ( *((_QWORD *)this + 57) )
    {
      for ( k = 0;
            k < *((_DWORD *)this + 24);
            *(_QWORD *)(2936 * v66 + *(_QWORD *)(*((_QWORD *)this + 57) + 8LL) + 400) = *(_QWORD *)(4000 * v66
                                                                                                  + *((_QWORD *)this + 16)
                                                                                                  + 928) )
      {
        v66 = k++;
      }
    }
    v37 = *((_QWORD *)this + 2);
    LODWORD(v68) = 2;
    v39 = __OFSUB__(*(_DWORD *)(v37 + 2736), 8704);
    v38 = *(_DWORD *)(v37 + 2736) - 8704 < 0;
    v40 = *(_QWORD *)(v37 + 216);
    v41 = v38 ^ v39;
    v83 = v41;
    if ( (int)DpiReadPnpRegistryValue(v40, L"NeedToSuspendVidSchBeforeSetGammaRamp", &v83, 4LL, v68) >= 0 )
    {
      v42 = v83;
      if ( v83 != v41 )
      {
        WdLogSingleEntry2(3LL, v83, *((_QWORD *)this + 2));
        WdLogGlobalForLineNumber = 5203;
      }
    }
    else
    {
      v42 = v41;
      v83 = v41;
    }
    *((_BYTE *)this + 291) = v42 != 0;
    v43 = DXGGLOBAL::GetGlobal();
    if ( (int)DXGADAPTERSOURCEHASH::AddNewAdapterEntry(
                (struct DXGGLOBAL *)((char *)v43 + 1384),
                (const struct _LUID *)(*((_QWORD *)this + 2) + 412LL),
                *((unsigned __int8 *)this + 290)) < 0 )
    {
      WdLogSingleEntry0(1LL);
      WdLogGlobalForLineNumber = 5216;
      DxgkLogInternalTriageEvent(0, 262146, -1, (unsigned int)L"NT_SUCCESS(TmpStatus)", 5216LL, 0LL, 0LL, 0LL, 0LL);
    }
    if ( v24 >= 0 )
    {
      EventHandle = 0LL;
      v67 = IoCreateNotificationEvent(0LL, &EventHandle);
      *((_QWORD *)this + 83) = v67;
      if ( v67 )
      {
        KeClearEvent(v67);
        ObfReferenceObject(*((PVOID *)this + 83));
        ZwClose(EventHandle);
      }
      else
      {
        WdLogSingleEntry0(6LL);
        WdLogGlobalForLineNumber = 5227;
        DxgkLogInternalTriageEvent(
          0,
          262145,
          -1,
          (unsigned int)L"Failed to create adapter VidPnSourceUsedBySession event object.",
          5227LL,
          0LL,
          0LL,
          0LL,
          0LL);
        return (unsigned int)-1073741801;
      }
    }
    return (unsigned int)v24;
  }
  return result;
}

__int64 __fastcall DpiFdoHandleStartDevice(PDEVICE_OBJECT DeviceObject, PIRP Irp)
{
  int *DeviceExtension; // rdi
  struct _IO_STACK_LOCATION *CurrentStackLocation; // r14
  __int64 v6; // rcx
  int v7; // eax
  int v8; // esi
  ULONG_PTR v9; // r8
  bool v10; // r15
  __int64 Status; // rsi
  PIO_SECURITY_CONTEXT SecurityContext; // rdx
  __int64 v14; // r9
  struct _UNICODE_STRING *v15; // rax
  __int64 v16; // rcx
  int v17; // eax
  PUNICODE_STRING v18; // rcx
  size_t v19; // r12
  void *Pool2; // rax
  int v21; // eax
  bool v22; // zf
  _WORD *StartContext; // r14
  int v24; // eax
  NTSTATUS v25; // eax
  void *v26; // rdx
  PIRP v27; // rax
  void *v28; // rdx
  PIRP v29; // rax
  int v31; // eax
  void *v32; // rcx
  PCLIENT_ID ClientId; // [rsp+20h] [rbp-A9h]
  PCLIENT_ID ClientIda; // [rsp+20h] [rbp-A9h]
  int StartRoutine; // [rsp+28h] [rbp-A1h]
  char v36; // [rsp+40h] [rbp-89h] BYREF
  char v37; // [rsp+41h] [rbp-88h]
  unsigned int v38; // [rsp+44h] [rbp-85h] BYREF
  int v39; // [rsp+48h] [rbp-81h] BYREF
  int v40; // [rsp+4Ch] [rbp-7Dh] BYREF
  size_t Size; // [rsp+50h] [rbp-79h] BYREF
  ULONG_PTR v42; // [rsp+58h] [rbp-71h] BYREF
  struct _UNICODE_STRING *v43; // [rsp+60h] [rbp-69h] BYREF
  ULONG_PTR v44; // [rsp+68h] [rbp-61h]
  struct _UNICODE_STRING *FileName; // [rsp+70h] [rbp-59h]
  void *ThreadHandle; // [rsp+78h] [rbp-51h] BYREF
  __int64 v47; // [rsp+80h] [rbp-49h] BYREF
  int v48; // [rsp+88h] [rbp-41h]
  const wchar_t *v49; // [rsp+90h] [rbp-39h]
  int *v50; // [rsp+98h] [rbp-31h]
  int v51; // [rsp+A0h] [rbp-29h]
  int *v52; // [rsp+A8h] [rbp-21h]
  int v53; // [rsp+B0h] [rbp-19h]
  __int64 v54; // [rsp+B8h] [rbp-11h]
  int v55; // [rsp+C0h] [rbp-9h]
  __int64 v56; // [rsp+C8h] [rbp-1h]
  __int128 v57; // [rsp+D0h] [rbp+7h]
  __int128 v58; // [rsp+E0h] [rbp+17h]
  char v60; // [rsp+140h] [rbp+77h]
  char v61; // [rsp+148h] [rbp+7Fh] BYREF

  DeviceExtension = (int *)DeviceObject->DeviceExtension;
  CurrentStackLocation = Irp->Tail.Overlay.CurrentStackLocation;
  v60 = 0;
  v37 = 0;
  v44 = 0LL;
  FileName = 0LL;
  LODWORD(Size) = 0;
  v38 = 0;
  AcquireMiniportListMutex();
  KeEnterCriticalRegion();
  if ( *((_BYTE *)DeviceExtension + 484) )
    DpiCheckForOutstandingD3Requests(DeviceExtension);
  ExAcquireResourceExclusiveLite(*((PERESOURCE *)DeviceExtension + 21), 1u);
  v7 = WindowsQueryLicenseDWORD(v6, &v38);
  if ( v7 < 0 )
  {
    v8 = 1;
    v38 = 1;
    WdLogSingleEntry1(4LL, v7);
    WdLogGlobalForLineNumber = 7066;
  }
  else
  {
    v8 = v38;
    WdLogSingleEntry1(4LL, v38);
    WdLogGlobalForLineNumber = 7053;
  }
  v39 = 1;
  v47 = 0LL;
  v54 = 0LL;
  v49 = L"MultiMonSupport";
  v55 = 0;
  v56 = 0LL;
  v50 = &v39;
  HIDWORD(ClientId) = 0;
  v52 = &v39;
  v48 = 288;
  v51 = 67108868;
  v53 = 4;
  v57 = 0LL;
  v58 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v47);
  v9 = 0LL;
  if ( !v39 )
    v8 = 0;
  v38 = v8;
  if ( !v8
    && DeviceExtension[4] == 1953656900
    && DeviceExtension[5] == 2
    && !(unsigned __int8)DpiFdoIsPostDevice(DeviceObject)
    && *((_BYTE *)DeviceExtension + 480) == (_BYTE)v9 )
  {
    v10 = 1;
    LODWORD(Status) = -1071774664;
    WdLogSingleEntry1((unsigned int)(v9 + 3), *((_QWORD *)DeviceExtension + 3));
    WdLogGlobalForLineNumber = 7113;
LABEL_96:
    if ( DeviceExtension[59] == 1 )
    {
      v31 = DeviceExtension[60];
      --DeviceExtension[69];
      DeviceExtension[59] = v31;
      DeviceExtension[60] = DeviceExtension[(DeviceExtension[69] & 7) + 61];
    }
    v32 = (void *)*((_QWORD *)DeviceExtension + 164);
    if ( v32 )
    {
      ExFreePoolWithTag(v32, 0);
      *((_QWORD *)DeviceExtension + 164) = 0LL;
    }
    if ( !v10 )
    {
      LOBYTE(StartRoutine) = 0;
      DxgCreateLiveDumpWithWdLogs(403LL, 2050LL, (int)Status, DeviceExtension[59], DeviceExtension[60], StartRoutine);
    }
    goto LABEL_88;
  }
  if ( *((_BYTE *)DeviceExtension + 1158) != (_BYTE)v9 )
  {
    ExEnterCriticalRegionAndAcquireFastMutexUnsafe(&dword_1C01540A0);
    if ( !dword_1C01540D8++ )
      KeClearEvent(&Object);
    ExReleaseFastMutexUnsafeAndLeaveCriticalRegion(&dword_1C01540A0);
    v9 = 0LL;
    v37 = 1;
  }
  DeviceExtension[678] = v8;
  if ( *((_BYTE *)DeviceExtension + 1155) == 1 )
  {
    SecurityContext = CurrentStackLocation->Parameters.Create.SecurityContext;
    v42 = v9;
    v43 = (struct _UNICODE_STRING *)v9;
    if ( SecurityContext )
    {
      FileName = CurrentStackLocation->Parameters.QueryDirectory.FileName;
      v44 = (ULONG_PTR)SecurityContext;
      DpiFilterOutVgaResources(DeviceExtension, SecurityContext, &v42, 0LL, 0LL);
      LOBYTE(v14) = 1;
      DpiFilterOutVgaResources(
        DeviceExtension,
        CurrentStackLocation->Parameters.QueryDirectory.FileName,
        &v43,
        v14,
        ClientIda);
      if ( v42 )
      {
        v15 = v43;
        if ( v43 )
        {
          CurrentStackLocation->Parameters.WMI.ProviderId = v42;
          CurrentStackLocation->Parameters.QueryDirectory.FileName = v15;
          v60 = 1;
        }
      }
    }
    else
    {
      WdLogSingleEntry1(2LL, 0LL);
      WdLogGlobalForLineNumber = 7181;
    }
  }
  if ( !(unsigned __int8)DpiFdoIsPostDevice(DeviceObject) && DeviceExtension[4] == 1953656900 && DeviceExtension[5] == 2 )
  {
    v40 = 0;
    LODWORD(ClientId) = 2;
    v17 = DpiReadPnpRegistryValue(v16, L"DisableNonPOSTDevice", &v40, 4LL, ClientId);
    if ( v17 >= 0 )
    {
      if ( v40 )
      {
        LODWORD(Status) = -1073741823;
        WdLogSingleEntry1(2LL, -1073741823LL);
        WdLogGlobalForLineNumber = 7230;
LABEL_36:
        v10 = 0;
        goto LABEL_96;
      }
    }
    else
    {
      WdLogSingleEntry1(4LL, v17);
      WdLogGlobalForLineNumber = 7217;
    }
  }
  IoForwardIrpSynchronously(*((PDEVICE_OBJECT *)DeviceExtension + 20), Irp);
  Status = Irp->IoStatus.Status;
  if ( (int)Status < 0 )
  {
    WdLogSingleEntry5(
      2LL,
      (unsigned int)DeviceExtension[136],
      Status,
      (unsigned int)DeviceExtension[281],
      (unsigned int)DeviceExtension[282],
      *(_QWORD *)(*((_QWORD *)DeviceExtension + 5) + 152LL));
    WdLogGlobalForLineNumber = 7255;
    v10 = (_DWORD)Status == -1073741810
       && *(_BYTE *)(*((_QWORD *)DeviceExtension + 5) + 134LL)
       && RtlCompareMemory(DeviceExtension + 136, &GUID_BUS_TYPE_USB, 0x10uLL) == 16;
    goto LABEL_96;
  }
  v18 = CurrentStackLocation->Parameters.QueryDirectory.FileName;
  if ( v18 )
  {
    DpiDetermineResourceListSize(v18, &Size);
    v19 = (unsigned int)Size;
    Pool2 = (void *)ExAllocatePool2(256LL, (unsigned int)Size, 1953656900LL);
    *((_QWORD *)DeviceExtension + 164) = Pool2;
    if ( !Pool2 )
    {
      LODWORD(Status) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 7301;
      goto LABEL_36;
    }
    memmove(Pool2, CurrentStackLocation->Parameters.QueryDirectory.FileName, v19);
    if ( v60 == 1 )
    {
      ExFreePoolWithTag(CurrentStackLocation->Parameters.Create.SecurityContext, 0);
      ExFreePoolWithTag(CurrentStackLocation->Parameters.QueryDirectory.FileName, 0);
      CurrentStackLocation->Parameters.WMI.ProviderId = v44;
      CurrentStackLocation->Parameters.QueryDirectory.FileName = FileName;
    }
  }
  DeviceExtension[(DeviceExtension[69] & 7) + 61] = DeviceExtension[60];
  v21 = DeviceExtension[59];
  ++DeviceExtension[69];
  DeviceExtension[60] = v21;
  DeviceExtension[59] = 1;
  if ( DeviceExtension[4] != 1953656900 || DeviceExtension[5] != 2 )
  {
LABEL_47:
    if ( !(_BYTE)word_1C0153AE0 )
      goto LABEL_50;
    goto LABEL_48;
  }
  if ( !*((_BYTE *)DeviceExtension + 2717) )
  {
    HIBYTE(word_1C0153AE0) = 1;
    goto LABEL_47;
  }
  LOBYTE(word_1C0153AE0) = 1;
LABEL_48:
  if ( HIBYTE(word_1C0153AE0) )
    KeSetEvent(&stru_1C0153AE8, 0, 0);
LABEL_50:
  v61 = 0;
  v36 = 0;
  if ( *((_BYTE *)DeviceExtension + 2716) || (int)DpiFdoIsMdmDeviceAndOwnsMux(DeviceObject, &v61, &v36) < 0 || !v61 )
  {
    if ( qword_1C0153DE8 )
      goto LABEL_64;
    if ( DeviceExtension[4] != 1953656900 || DeviceExtension[5] != 2 )
    {
      if ( *(_BYTE *)(*((_QWORD *)DeviceExtension + 21) + 108LL) && *((_QWORD *)DeviceExtension + 354) )
        qword_1C0153DE8 = *((_QWORD *)DeviceExtension + 354);
      goto LABEL_64;
    }
    if ( (unsigned __int8)DpiFdoIsPostDevice(DeviceObject) )
    {
LABEL_60:
      qword_1C0153DE8 = (__int64)DeviceObject;
      goto LABEL_64;
    }
    v22 = *(_BYTE *)(*((_QWORD *)DeviceExtension + 21) + 108LL) == 0;
  }
  else
  {
    v22 = v36 == 0;
  }
  if ( !v22 )
    goto LABEL_60;
LABEL_64:
  if ( !*((_BYTE *)DeviceExtension + 480)
    && *((_BYTE *)DeviceExtension + 1153)
    && !(unsigned __int8)DpiFdoIsMsBddAnchoredDevice(DeviceObject) )
  {
    WdLogSingleEntry1(4LL, DeviceObject);
    WdLogGlobalForLineNumber = 7432;
    v10 = 1;
    LODWORD(Status) = -1071774664;
    goto LABEL_96;
  }
  if ( byte_1C0153AE2 && !*((_BYTE *)DeviceExtension + 1158) )
  {
    ThreadHandle = 0LL;
    StartContext = (_WORD *)ExAllocatePool2(256LL, 1552LL, 1953656900LL);
    if ( !StartContext )
    {
      LODWORD(Status) = -1073741801;
      WdLogSingleEntry1(6LL, -1073741801LL);
      WdLogGlobalForLineNumber = 7476;
LABEL_72:
      v10 = 0;
      goto LABEL_96;
    }
    if ( *(_BYTE *)(*((_QWORD *)DeviceExtension + 21) + 108LL)
      || (v22 = (unsigned __int8)DpiFdoIsMsBddAnchoredDevice(DeviceObject) == 0, v24 = 0, !v22) )
    {
      v24 = 2;
    }
    *(_DWORD *)StartContext = v24;
    StartContext[2] = 0;
    *((_DWORD *)StartContext + 131) = 0;
    v25 = PsCreateSystemThread(&ThreadHandle, 0x1FFFFFu, 0LL, 0LL, 0LL, DpiFdoStartAdapterThread, StartContext);
    LODWORD(Status) = v25;
    if ( v25 < 0 )
    {
      WdLogSingleEntry1(2LL, v25);
      WdLogGlobalForLineNumber = 7507;
      ExFreePoolWithTag(StartContext, 0x74727044u);
      goto LABEL_72;
    }
    ZwClose(ThreadHandle);
  }
  DeviceExtension[71] = 1;
  DeviceExtension[70] = 1;
  PoSetPowerState(DeviceObject, DevicePowerState, (POWER_STATE)1);
  v10 = 0;
  if ( (int)Status < 0 )
    goto LABEL_96;
  if ( DeviceExtension[4] == 1953656900 && DeviceExtension[5] == 2 )
  {
    v26 = (void *)*((_QWORD *)DeviceExtension + 684);
    if ( v26 )
    {
      v27 = IoCsqRemoveNextIrp((PIO_CSQ)(DeviceExtension + 1346), v26);
      *((_QWORD *)DeviceExtension + 684) = 0LL;
      if ( v27 )
      {
        *((_BYTE *)DeviceExtension + 5500) = 1;
        v27->IoStatus.Status = 0;
        v27->IoStatus.Information = 0LL;
        IofCompleteRequest(v27, 0);
        IoInvalidateDeviceState(*((PDEVICE_OBJECT *)DeviceExtension + 19));
      }
    }
    v28 = (void *)*((_QWORD *)DeviceExtension + 688);
    if ( v28 )
    {
      v29 = IoCsqRemoveNextIrp((PIO_CSQ)(DeviceExtension + 1346), v28);
      *((_QWORD *)DeviceExtension + 688) = 0LL;
      if ( v29 )
      {
        *((_BYTE *)DeviceExtension + 5532) = 1;
        v29->IoStatus.Status = 0;
        v29->IoStatus.Information = 0LL;
        IofCompleteRequest(v29, 0);
      }
    }
  }
LABEL_88:
  if ( v37 )
  {
    ExEnterCriticalRegionAndAcquireFastMutexUnsafe(&dword_1C01540A0);
    if ( !--dword_1C01540D8 )
      KeSetEvent(&Object, 0, 0);
    ExReleaseFastMutexUnsafeAndLeaveCriticalRegion(&dword_1C01540A0);
  }
  if ( *((_BYTE *)DeviceExtension + 484) )
    DpiEnableD3Requests(*((_QWORD *)DeviceExtension + 3));
  ExReleaseResourceLite(*((PERESOURCE *)DeviceExtension + 21));
  KeLeaveCriticalRegion();
  ReleaseMiniportListMutex();
  Irp->IoStatus.Status = Status;
  IofCompleteRequest(Irp, 1);
  return (unsigned int)Status;
}

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

bool DpiHybridInternalPanelOverride()
{
  __int64 v1; // [rsp+30h] [rbp-19h] BYREF
  int v2; // [rsp+38h] [rbp-11h]
  const wchar_t *v3; // [rsp+40h] [rbp-9h]
  int *v4; // [rsp+48h] [rbp-1h]
  int v5; // [rsp+50h] [rbp+7h]
  int *v6; // [rsp+58h] [rbp+Fh]
  int v7; // [rsp+60h] [rbp+17h]
  __int64 v8; // [rsp+68h] [rbp+1Fh]
  int v9; // [rsp+70h] [rbp+27h]
  __int64 v10; // [rsp+78h] [rbp+2Fh]
  __int128 v11; // [rsp+80h] [rbp+37h]
  __int128 v12; // [rsp+90h] [rbp+47h]
  int v13; // [rsp+B0h] [rbp+67h] BYREF

  if ( !g_OSTestSigningEnabled )
    return 0;
  v13 = 0;
  v1 = 0LL;
  v8 = 0LL;
  v9 = 0;
  v10 = 0LL;
  v3 = L"HybridInternalPanelOverrideEnable";
  v4 = &v13;
  v6 = &v13;
  v2 = 288;
  v5 = 67108868;
  v7 = 4;
  v11 = 0LL;
  v12 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v1);
  return v13 != 0;
}

void __fastcall OUTPUTDUPL_SESSION_MGR::InitializeMaxActiveOutputDuplApps(OUTPUTDUPL_SESSION_MGR *this)
{
  __int64 v2; // [rsp+30h] [rbp-19h] BYREF
  int v3; // [rsp+38h] [rbp-11h]
  const wchar_t *v4; // [rsp+40h] [rbp-9h]
  OUTPUTDUPL_SESSION_MGR *v5; // [rsp+48h] [rbp-1h]
  int v6; // [rsp+50h] [rbp+7h]
  int *v7; // [rsp+58h] [rbp+Fh]
  int v8; // [rsp+60h] [rbp+17h]
  __int64 v9; // [rsp+68h] [rbp+1Fh]
  int v10; // [rsp+70h] [rbp+27h]
  __int64 v11; // [rsp+78h] [rbp+2Fh]
  __int128 v12; // [rsp+80h] [rbp+37h]
  __int128 v13; // [rsp+90h] [rbp+47h]
  int v14; // [rsp+B0h] [rbp+67h] BYREF

  v3 = 288;
  v2 = 0LL;
  v9 = 0LL;
  v14 = 4;
  v8 = 4;
  v10 = 0;
  v4 = L"OutputDuplicationSessionApplicationLimit";
  v11 = 0LL;
  v5 = this;
  v6 = 67108868;
  v7 = &v14;
  v12 = 0LL;
  v13 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v2);
  if ( (unsigned int)(*(_DWORD *)this - 1) > 0xF )
    *(_DWORD *)this = v14;
}

__int64 __fastcall DXGADAPTER::Initialize(DXGADAPTER *this, PDEVICE_OBJECT DeviceObject, struct _DXGK_ADAPTER_CAPS *a3)
{
  struct _ERESOURCE *v6; // rax
  __int64 result; // rax
  NTSTATUS v8; // eax
  NTSTATUS LocallyUniqueId; // ebx
  PDEVICE_OBJECT DeviceAttachmentBaseRef; // rax
  __int64 v11; // rax
  const wchar_t *v12; // r9
  int v13; // edx
  struct _ERESOURCE *v14; // rax
  NTSTATUS v15; // eax
  int v16; // eax
  __int64 v17; // r15
  int AdapterInfo; // eax
  struct _LUID *v19; // rdx
  int (__fastcall *v20)(_QWORD, __int128 *); // rax
  unsigned __int8 v21; // bl
  DXGSESSIONDATA *SessionDataForSpecifiedSession; // rax
  unsigned int v23; // eax
  __int64 v24; // rax
  const wchar_t *v25; // r9
  unsigned int v26; // eax
  const struct _GUID *v27; // rdx
  unsigned __int16 v28; // r8
  unsigned __int16 v29; // r9
  int v30; // eax
  const wchar_t *v31; // r9
  int v32; // eax
  unsigned __int8 v33; // dl
  __int64 v34; // rcx
  __int64 v35; // r14
  __int64 v36; // rax
  unsigned int v37; // r13d
  unsigned __int8 v38; // r8
  __int64 v39; // rax
  const wchar_t *v40; // r9
  int v41; // eax
  int v42; // ecx
  __int64 v43; // rax
  int v44; // ecx
  __int64 v45; // r15
  int v46; // eax
  int v47; // ecx
  __int64 v48; // rcx
  int v49; // eax
  int v50; // ecx
  char v51; // al
  int v52; // eax
  unsigned int v53; // ebx
  __int64 v54; // rax
  __int64 v55; // rax
  char v56; // r12
  unsigned int v57; // eax
  unsigned int v58; // r8d
  __int64 v59; // r9
  UINT PhysicalAdapterCapsSizeFromDdiVersion; // r15d
  int v61; // eax
  __int64 v62; // rdx
  __int64 v63; // rcx
  __int64 v64; // rax
  DXGGLOBAL *Global; // rax
  __int64 v66; // rcx
  unsigned int v67; // r8d
  __int64 v68; // rdx
  __int64 v69; // rcx
  __int64 v70; // rdi
  __int64 v71; // rbx
  __int64 v72; // rdi
  __int64 v73; // rbx
  __int64 v74; // rax
  __int64 v75; // rax
  int v76; // eax
  __int64 RenderCore; // rdi
  unsigned int v78; // edx
  unsigned __int64 v79; // r8
  __int64 v80; // r13
  __int64 v81; // r12
  int v82; // ecx
  int v83; // edi
  int v84; // ebx
  unsigned __int8 IsGpuVaIoMmuGlobalSupported; // al
  const wchar_t *v86; // r9
  int v87; // eax
  char v88; // al
  int v89; // eax
  __int64 v90; // rax
  int v91; // eax
  int v92; // ecx
  int v93; // eax
  struct _DXGK_ADAPTER_CAPS *v94; // r12
  char v95; // cl
  char v96; // dl
  char v97; // al
  char v98; // r8
  char v99; // cl
  char v100; // dl
  char v101; // cl
  char v102; // al
  char v103; // al
  char v104; // cl
  unsigned int v105; // eax
  int v106; // ecx
  __int64 v107; // rax
  struct DXGGLOBAL *v108; // rax
  struct DXGGLOBAL *v109; // rax
  struct DXGGLOBAL *v110; // rax
  char v111; // r9
  char v112; // r8
  unsigned int v113; // ecx
  unsigned int v114; // edx
  __int64 v115; // rax
  __int64 v116; // rax
  unsigned int v117; // ebx
  DXGGLOBAL *v118; // rax
  int v119; // eax
  int v120; // ecx
  __int64 v121; // rax
  __int64 v122; // rcx
  int v123; // eax
  char v124; // cl
  int v125; // eax
  __int64 v126; // rax
  char *v127; // rbx
  int DisplayCore; // eax
  bool v129; // zf
  char v130; // cl
  char v131; // dl
  int v132; // eax
  char v133; // al
  __int64 v134; // rdx
  DXGADAPTER *v135; // rcx
  int v136; // eax
  __int64 v137; // rax
  bool v138; // cf
  __int64 v139; // rdx
  __int64 v140; // r8
  int v141; // eax
  DXGGLOBAL *v142; // rax
  int v143; // eax
  __int64 v144; // rax
  int v145; // eax
  DXGADAPTER *v146; // rcx
  __int64 v147; // r14
  __int64 v148; // rbx
  struct DXGGLOBAL *v149; // rax
  int v150; // eax
  struct DXGGLOBAL *v151; // rax
  __int64 v152; // rdx
  DXGGLOBAL *v153; // rax
  __int64 v154; // [rsp+20h] [rbp-E0h]
  void *v155; // [rsp+20h] [rbp-E0h]
  void *v156; // [rsp+20h] [rbp-E0h]
  void *v157; // [rsp+20h] [rbp-E0h]
  void *v158; // [rsp+28h] [rbp-D8h]
  __int64 v159; // [rsp+28h] [rbp-D8h]
  __int64 v160; // [rsp+28h] [rbp-D8h]
  __int64 v161; // [rsp+30h] [rbp-D0h]
  unsigned int v162; // [rsp+50h] [rbp-B0h] BYREF
  bool IsAdapterSessionized; // [rsp+54h] [rbp-ACh]
  unsigned int v164; // [rsp+58h] [rbp-A8h] BYREF
  int v165; // [rsp+5Ch] [rbp-A4h] BYREF
  int v166; // [rsp+60h] [rbp-A0h] BYREF
  int v167; // [rsp+64h] [rbp-9Ch] BYREF
  int v168; // [rsp+68h] [rbp-98h] BYREF
  unsigned int v169; // [rsp+6Ch] [rbp-94h]
  __int64 v170; // [rsp+70h] [rbp-90h] BYREF
  unsigned int v171; // [rsp+78h] [rbp-88h]
  __int64 v172; // [rsp+80h] [rbp-80h] BYREF
  _DXGKARG_QUERYADAPTERINFO v173; // [rsp+88h] [rbp-78h] BYREF
  struct _DXGK_ADAPTER_CAPS *v174[2]; // [rsp+B8h] [rbp-48h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v175; // [rsp+C8h] [rbp-38h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v176; // [rsp+F8h] [rbp-8h] BYREF
  struct _DXGKARG_QUERYADAPTERINFO v177; // [rsp+128h] [rbp+28h] BYREF
  __int128 v178; // [rsp+158h] [rbp+58h] BYREF
  unsigned int v179[2]; // [rsp+168h] [rbp+68h] BYREF

  v174[0] = a3;
  if ( KeGetCurrentIrql() )
  {
    WdLogSingleEntry0(1LL);
    WdLogGlobalForLineNumber = 6949;
    DxgkLogInternalTriageEvent(
      0,
      262146,
      -1,
      (unsigned int)L"KeGetCurrentIrql() == PASSIVE_LEVEL",
      6949LL,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  if ( *((_DWORD *)this + 50) )
    return 3221225485LL;
  v6 = (struct _ERESOURCE *)operator new(104LL, 1265072196LL);
  *((_QWORD *)this + 21) = v6;
  if ( !v6 )
  {
    WdLogSingleEntry2(3LL, this, -1073741801LL);
    WdLogGlobalForLineNumber = 6967;
    return 3221225495LL;
  }
  v8 = ExInitializeResourceLite(v6);
  LocallyUniqueId = v8;
  if ( v8 < 0 )
  {
    WdLogSingleEntry2(3LL, this, v8);
    WdLogGlobalForLineNumber = 6978;
    return (unsigned int)LocallyUniqueId;
  }
  *((_QWORD *)this + 27) = DeviceObject;
  *((_QWORD *)this + 28) = DpiGetSysMmAdapterFromDevice(DeviceObject);
  DeviceAttachmentBaseRef = IoGetDeviceAttachmentBaseRef(DeviceObject);
  *((_QWORD *)this + 29) = DeviceAttachmentBaseRef;
  ObfDereferenceObject(DeviceAttachmentBaseRef);
  LocallyUniqueId = ZwAllocateLocallyUniqueId((PLUID)((char *)this + 4756));
  if ( LocallyUniqueId < 0 )
  {
    WdLogSingleEntry0(6LL);
    v11 = 6999LL;
    v12 = L"ZwAllocateLocallyUniqueId failed";
LABEL_12:
    v13 = 262145;
LABEL_13:
    WdLogGlobalForLineNumber = v11;
    DxgkLogInternalTriageEvent(0, v13, -1, (_DWORD)v12, v11, 0LL, 0LL, 0LL, 0LL);
    return (unsigned int)LocallyUniqueId;
  }
  v14 = (struct _ERESOURCE *)operator new(104LL, 1265072196LL);
  *((_QWORD *)this + 35) = v14;
  if ( !v14 )
  {
    WdLogSingleEntry2(3LL, this, -1073741801LL);
    WdLogGlobalForLineNumber = 7012;
    return 3221225495LL;
  }
  v15 = ExInitializeResourceLite(v14);
  LocallyUniqueId = v15;
  if ( v15 < 0 )
  {
    WdLogSingleEntry2(3LL, this, v15);
    WdLogGlobalForLineNumber = 7023;
    return (unsigned int)LocallyUniqueId;
  }
  _InterlockedIncrement64((volatile signed __int64 *)this + 3);
  v170 = 0LL;
  *((_QWORD *)this + 5) = -1LL;
  if ( *((_BYTE *)DeviceObject->DeviceExtension + 481) )
  {
    v16 = DXGADAPTER::InitializeParavirtualizedAdapter(this, (struct DRIVER_WORKAROUNDS *)&v170);
    v17 = v16;
    if ( v16 < 0 )
    {
      WdLogSingleEntry1(2LL, v16);
      WdLogGlobalForLineNumber = 7046;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"InitializeParavirtualizedAdapter failed: 0x%I64x",
        v17,
        0LL,
        0LL,
        0LL,
        0LL);
      return (unsigned int)v17;
    }
  }
  else
  {
    *((_BYTE *)this + 1785) = 0;
    AdapterInfo = DpiGetAdapterInfo((int)DeviceObject, (char *)this + 1744, (char *)this + 288);
    LocallyUniqueId = AdapterInfo;
    if ( AdapterInfo < 0 )
    {
      WdLogSingleEntry2(3LL, this, AdapterInfo);
      WdLogGlobalForLineNumber = 7063;
      return (unsigned int)LocallyUniqueId;
    }
  }
  DpiFdoSetFeatureDatabaseDxgAdapter(*((_QWORD *)this + 27), this);
  *(_QWORD *)v179 = 0LL;
  v20 = (int (__fastcall *)(_QWORD, __int128 *))*((_QWORD *)this + 296);
  v178 = 0LL;
  if ( v20 && v20(*((_QWORD *)this + 287), &v178) >= 0 )
  {
    *(_QWORD *)((char *)this + 4828) = *((_QWORD *)&v178 + 1);
    *((_DWORD *)this + 1209) = v179[0];
  }
  IsAdapterSessionized = DXGADAPTER::IsAdapterSessionized(this, v19, v179, 0LL);
  v21 = IsAdapterSessionized;
  if ( IsAdapterSessionized )
  {
    SessionDataForSpecifiedSession = DXGSESSIONMGR::GetSessionDataForSpecifiedSession(
                                       *(DXGSESSIONMGR **)(*((_QWORD *)this + 2) + 944LL),
                                       v179[0]);
    if ( !SessionDataForSpecifiedSession
      || (v23 = DXGSESSIONDATA::AcquireSessionAdapterOrdinal(SessionDataForSpecifiedSession),
          *((_DWORD *)this + 61) = v23,
          v23 == -1) )
    {
      WdLogSingleEntry2(2LL, v179[0], -1073741801LL);
      v24 = v179[0];
      v25 = L"Exceeded the maximum number of sessionized adapter in session 0x%I64x, returning 0x%I64x.";
      v159 = -1073741801LL;
      WdLogGlobalForLineNumber = 7096;
LABEL_31:
      DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v25, v24, v159, 0LL, 0LL, 0LL);
      return 3221225495LL;
    }
  }
  v26 = DXGGLOBAL::AcquireAdapterOrdinal(*((DXGGLOBAL **)this + 2), v21);
  *((_DWORD *)this + 60) = v26;
  if ( v26 == -1 )
    return 3221225495LL;
  if ( (*((_DWORD *)this + 111) & 0x200) != 0 )
    *((_BYTE *)DXGGLOBAL::GetGlobal() + 304832) = 1;
  v30 = *((_DWORD *)this + 111);
  if ( (v30 & 8) != 0 && (v30 & 0x10) != 0 )
  {
    WdLogSingleEntry0(1LL);
    WdLogGlobalForLineNumber = 7120;
    DxgkLogInternalTriageEvent(
      0,
      262146,
      -1,
      (unsigned int)L"!(IsSoftGPU() && IsWarpAdapter())",
      7120LL,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  if ( !*((_QWORD *)this + 57) )
  {
    WdLogSingleEntry0(2LL);
    v31 = L"Miniport did not provide required DDIs";
    v160 = 0LL;
    v154 = 7127LL;
    WdLogGlobalForLineNumber = 7127;
LABEL_40:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v31, v154, v160, 0LL, 0LL, 0LL);
    return 3221225561LL;
  }
  if ( !*((_QWORD *)this + 74) )
    *((_QWORD *)this + 74) = DXGADAPTER::DefaultDdiEscape;
  if ( !*((_QWORD *)this + 135) )
    *((_QWORD *)this + 135) = W32kStub_GreSfmOpenTokenEvent;
  v32 = DXGADAPTER::CallDriverQueryInterface(this, v27, v28, v29, (char *)this + 2096, v158);
  v35 = v32;
  if ( v32 >= 0 )
  {
    if ( *((_WORD *)this + 1049) >= 4u )
      goto LABEL_49;
  }
  else
  {
    v36 = WdLogNewEntry5_WdTrace(v34);
    *(_QWORD *)(v36 + 24) = this;
    *(_QWORD *)(v36 + 32) = v35;
    WdLogGlobalForLineNumber = 7158;
  }
  memset((char *)this + 2096, 0, 0xB8uLL);
LABEL_49:
  v37 = *(_DWORD *)(*(_QWORD *)(*(_QWORD *)(*((_QWORD *)this + 27) + 64LL) + 40LL) + 28LL);
  v171 = v37;
  *((_DWORD *)this + 570) = v37;
  if ( v37 < 0x7000 )
  {
    if ( v37 < 0x6002 )
      goto LABEL_58;
  }
  else
  {
    if ( !*((_DWORD *)this + 464) )
      goto LABEL_58;
    if ( *((_DWORD *)this + 465) )
    {
      v38 = 0;
LABEL_57:
      DXGADAPTER::SetModeBehavior(this, v33, v38);
      goto LABEL_58;
    }
  }
  if ( *((_DWORD *)this + 464) && *((_DWORD *)this + 465) )
  {
    v38 = 1;
    goto LABEL_57;
  }
LABEL_58:
  if ( v37 - 20480 <= 5 )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7202LL;
    v40 = L"Cannot load an M1 threshold driver on later builds.";
LABEL_60:
    WdLogGlobalForLineNumber = v39;
LABEL_61:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v40, v39, 0LL, 0LL, 0LL, 0LL);
    return 3221225485LL;
  }
  *(_QWORD *)&v173.InputDataSize = 0LL;
  v173.pOutputData = (char *)this + 2400;
  *(_QWORD *)&v173.Type = 1LL;
  *(_QWORD *)&v173.Flags.0 = 0LL;
  HIDWORD(v173.hKmdProcessHandle) = 0;
  v173.pInputData = 0LL;
  v173.OutputDataSize = GetDriverCapsSizeFromDdiVersion(v37);
  if ( !v173.OutputDataSize )
    return 3221225485LL;
  v41 = DXGADAPTER::DdiQueryAdapterInfo(this, &v173);
  v17 = v41;
  if ( v41 < 0 )
  {
    WdLogSingleEntry1(2LL, v41);
    WdLogGlobalForLineNumber = 7225;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Miniport failed DdiQueryAdapterInfo(DXGKQAITYPE_DRIVERCAPS) with status 0x%I64x",
      v17,
      0LL,
      0LL,
      0LL,
      0LL);
    return (unsigned int)v17;
  }
  v42 = *((_DWORD *)this + 684);
  if ( v42 <= 9472 )
  {
    if ( v42 < 4864 )
    {
      v45 = 0LL;
      if ( *((_QWORD *)this + 104) )
      {
        v44 = 1300;
      }
      else if ( v42 == 4608 )
      {
        v44 = 1200;
      }
      else if ( !*((_QWORD *)this + 100) || (v44 = 1105, (*((_DWORD *)this + 613) & 4) == 0) )
      {
        v44 = 1000;
      }
      *((_DWORD *)this + 751) = v44;
      goto LABEL_79;
    }
  }
  else if ( *((_DWORD *)DeviceObject->DeviceExtension + 687) <= 0xA00Bu )
  {
    WdLogSingleEntry1(2LL, *((int *)this + 684));
    v43 = *((int *)this + 684);
    WdLogGlobalForLineNumber = 7231;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Miniport returned incorrect WDDMVersion: 0x%I64x",
      v43,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225485LL;
  }
  v44 = DxgkConvertWddmVersionToD3DKMTDriverVersion();
  *((_DWORD *)this + 751) = v44;
  v45 = 0LL;
LABEL_79:
  v46 = *((_DWORD *)this + 744);
  if ( v44 >= 2600 )
  {
    v47 = *((_DWORD *)this + 111);
    if ( (v46 & 8) != 0 )
    {
      *((_DWORD *)this + 111) = v47 | 0x80000;
    }
    else if ( (v47 & 0x80000) != 0 && v37 >= 0x11002 )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7285LL;
      v40 = L"MiscCaps.ComputeOnly is not set, but the device belongs to the ComputeAccelerator class";
      goto LABEL_60;
    }
  }
  else
  {
    v46 &= ~8u;
    *((_DWORD *)this + 744) = v46;
  }
  if ( *((_BYTE *)this + 1784) && (v46 & 0xC) == 0 )
  {
    WdLogSingleEntry0(2LL);
    WdLogGlobalForLineNumber = 7292;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"UMD name is missing and device is not compute only",
      7292LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225524LL;
  }
  v48 = *((_QWORD *)this + 27);
  v165 = 0;
  LODWORD(v155) = 2;
  v49 = DpiReadPnpRegistryValue(v48, L"ACGSupported", &v165, 4LL, v155);
  v50 = v165;
  if ( v49 < 0 )
    v50 = 0;
  v165 = v50;
  if ( v50 || (v51 = 0, *((int *)this + 751) >= 2200) )
    v51 = 1;
  *((_BYTE *)this + 212) = v51;
  if ( *((_BYTE *)this + 209) )
  {
    *((_BYTE *)a3 + 1) &= ~1u;
    *(_BYTE *)a3 &= 0x7Bu;
    *((_DWORD *)this + 744) &= 0xFFFFFFEB;
    *((_DWORD *)this + 617) &= 0xFFFFD2FF;
    *((_BYTE *)this + 2940) = 0;
    *((_BYTE *)this + 2968) = 1;
    *((_BYTE *)this + 2942) = 1;
    if ( *((_BYTE *)this + 210) )
      *((_DWORD *)this + 613) &= ~0x100000u;
  }
  else if ( v37 >= 0x5023 )
  {
    if ( g_bCreateParavirtualizedGpu )
    {
      v52 = *((_DWORD *)this + 111);
      if ( (v52 & 4) == 0 && (v52 & 0x10) == 0 && !*(_BYTE *)(*((_QWORD *)DeviceObject->DeviceExtension + 5) + 133LL) )
        *((_DWORD *)this + 617) |= 0x400u;
    }
  }
  v169 = *((_DWORD *)this + 74);
  v53 = v169;
  v54 = 344LL * v169;
  if ( !is_mul_ok(v169, 0x158uLL) )
    v54 = -1LL;
  v55 = operator new[](v54, 1265072196LL, 64LL);
  *((_QWORD *)this + 374) = v55;
  if ( !v55 )
  {
    WdLogSingleEntry0(6LL);
    WdLogGlobalForLineNumber = 7351;
    DxgkLogInternalTriageEvent(
      0,
      262145,
      -1,
      (unsigned int)L"Failed to allocate DXGK_PHYSICALADAPTERINFO",
      7351LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return 3221225495LL;
  }
  v56 = 0;
  if ( *((int *)this + 684) < 0x2000 || v37 < 0x5005 )
    goto LABEL_147;
  *((_DWORD *)this + 750) = 0;
  v57 = 0;
  v162 = 0;
  if ( v53 )
  {
    PhysicalAdapterCapsSizeFromDdiVersion = GetPhysicalAdapterCapsSizeFromDdiVersion(v37);
    while ( 1 )
    {
      v175.pInputData = &v162;
      *(_QWORD *)&v175.Type = 15LL;
      *(_QWORD *)&v175.InputDataSize = 4LL;
      v175.pOutputData = (void *)(v59 + 344LL * v58);
      *(_QWORD *)&v175.Flags.0 = 0LL;
      HIDWORD(v175.hKmdProcessHandle) = 0;
      v175.OutputDataSize = PhysicalAdapterCapsSizeFromDdiVersion;
      v61 = DXGADAPTER::DdiQueryAdapterInfo(this, &v175);
      if ( v61 < 0 )
        break;
      if ( v37 >= 0xC003 )
      {
        v62 = *((_QWORD *)this + 374);
        v63 = 344LL * v162;
        if ( (*(_DWORD *)(v63 + v62 + 16) & 0x20) != 0 )
        {
          v64 = *(unsigned int *)(v63 + v62 + 24);
          if ( (unsigned int)v64 >= *(unsigned __int16 *)(v63 + v62) )
          {
            WdLogSingleEntry3(2LL, this, v64, *(unsigned __int16 *)(v63 + v62));
            v75 = *((_QWORD *)this + 374);
            WdLogGlobalForLineNumber = 7396;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"Adapter 0x%I64x: VirtualCopyEngineSupported but node index is invalid (VirtualCopyIndex:%u, "
                             "NumExecutionNodes:%u)",
              (__int64)this,
              *(unsigned int *)(344LL * v162 + v75 + 24),
              *(unsigned __int16 *)(344LL * v162 + v75),
              0LL,
              0LL);
            return 3221225485LL;
          }
          if ( (*((_DWORD *)this + 617) & 0x2000) == 0 )
          {
            WdLogSingleEntry1(2LL, this);
            WdLogGlobalForLineNumber = 7403;
            DxgkLogInternalTriageEvent(
              0,
              0x40000,
              -1,
              (unsigned int)L"Adapter 0x%I64x: IoMmuSecureModeRequired must be set for a device exposing a virtual copy engine",
              (__int64)this,
              0LL,
              0LL,
              0LL,
              0LL);
            return 3221225485LL;
          }
        }
      }
      Global = DXGGLOBAL::GetGlobal();
      if ( DXGGLOBAL::GpuVaIoMmuEnabled(Global) )
      {
        v66 = *((_QWORD *)this + 27);
        v166 = 0;
        v167 = 0;
        LODWORD(v156) = 2;
        if ( (int)DpiReadPnpRegistryValue(v66, L"DxgkGpuVaIommuRequired", &v166, 4LL, v156) >= 0 )
          *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) = (v166 != 0 ? 0x40 : 0) | *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) & 0xFFFFFFBF;
        LODWORD(v157) = 2;
        if ( (int)DpiReadPnpRegistryValue(*((_QWORD *)this + 27), L"DxgkGpuVaIommuGlobalSupported", &v167, 4LL, v157) >= 0 )
          *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) = (v167 != 0 ? 0x80 : 0) | *(_DWORD *)(344LL * v162 + *((_QWORD *)this + 374) + 16) & 0xFFFFFF7F;
      }
      v67 = v162;
      v68 = *((_QWORD *)this + 374);
      v69 = 344LL * v162;
      if ( (*(_DWORD *)(v69 + v68 + 16) & 2) != 0 )
      {
        *(_BYTE *)(v69 + v68 + 49) = 1;
        v67 = v162;
      }
      v70 = *((_QWORD *)this + 374);
      v71 = 344LL * v67;
      if ( (*(_DWORD *)(v71 + v70 + 16) & 0x40) != 0 )
      {
        if ( !DXGADAPTER::IsGpuVaIoMmuSupported(this) )
        {
          WdLogSingleEntry1(2LL, this);
          WdLogGlobalForLineNumber = 7434;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Adapter 0x%I64x: GpuVaIommuRequired is set for a physical adapter, but not in IOMMU_CAPS",
            (__int64)this,
            0LL,
            0LL,
            0LL,
            0LL);
          return 3221225485LL;
        }
        *(_BYTE *)(v71 + v70 + 49) = 1;
        *(_BYTE *)(344LL * v162 + *((_QWORD *)this + 374) + 48) = 1;
        v67 = v162;
      }
      v72 = *((_QWORD *)this + 374);
      v73 = 344LL * v67;
      if ( (*(_DWORD *)(v73 + v72 + 16) & 0x80u) != 0 )
      {
        if ( !DXGADAPTER::IsGpuVaIoMmuGlobalSupported(this) )
        {
          WdLogSingleEntry1(2LL, this);
          WdLogGlobalForLineNumber = 7445;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Adapter 0x%I64x: GpuVaIommuGlobalRequired is set for a physical adapter, but not in IOMMU_CAPS",
            (__int64)this,
            0LL,
            0LL,
            0LL,
            0LL);
          return 3221225485LL;
        }
        *(_BYTE *)(v73 + v72 + 49) = 1;
        *(_BYTE *)(344LL * v162 + *((_QWORD *)this + 374) + 48) = 1;
        v67 = v162;
      }
      v59 = *((_QWORD *)this + 374);
      v53 = v169;
      v74 = v67;
      v58 = v67 + 1;
      v57 = *(unsigned __int16 *)(344 * v74 + v59) + *((_DWORD *)this + 750);
      v162 = v58;
      *((_DWORD *)this + 750) = v57;
      if ( v58 >= v53 )
        goto LABEL_130;
    }
    WdLogSingleEntry1(4LL, v61);
    WdLogGlobalForLineNumber = 7378;
    v56 = 1;
  }
  else
  {
LABEL_130:
    if ( *((int *)this + 751) <= 2400 && v57 > 0x40 )
    {
      WdLogSingleEntry3(2LL, this, 64LL, v57);
      v161 = *((unsigned int *)this + 750);
      WdLogGlobalForLineNumber = 7463;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Adapter 0x%I64x: Exceeded maximum number of %I64d nodes on pre-WDDM 2.5 adapter. Total node count: %I64d",
        (__int64)this,
        64LL,
        v161,
        0LL,
        0LL);
      return 3221225485LL;
    }
    if ( (*((_DWORD *)this + 616) & 1) == 0 )
    {
      WdLogSingleEntry1(2LL, this);
      WdLogGlobalForLineNumber = 7468;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Adapter 0x%I64x: SchedulingCaps.MultiEngineAware is not set by WDDMv2 driver",
        (__int64)this,
        0LL,
        0LL,
        0LL,
        0LL);
      return 3221225485LL;
    }
  }
  v45 = 0LL;
  if ( (*((_DWORD *)this + 617) & 0x800) != 0 )
  {
    v164 = 0;
    if ( v53 )
    {
      while ( 1 )
      {
        v172 = 0LL;
        v173.pInputData = &v164;
        v173.Type = DXGKQAITYPE_FRAMEBUFFERSAVESIZE;
        v173.pOutputData = &v172;
        v173.InputDataSize = 4;
        v173.OutputDataSize = 8;
        v76 = DXGADAPTER::DdiQueryAdapterInfo(this, &v173);
        RenderCore = v76;
        if ( v76 < 0 )
          break;
        if ( (v172 & 0xFFF) != 0 )
        {
          WdLogSingleEntry1(2LL, v172);
          v39 = v172;
          v40 = L"Frame buffer reserve size must be a multiple of PAGE_SIZE. Size=%I64u";
          WdLogGlobalForLineNumber = 7493;
          goto LABEL_61;
        }
        *(_QWORD *)(344LL * v164 + *((_QWORD *)this + 374) + 56) = v172;
        v78 = v164;
        v79 = *(_QWORD *)(344LL * v164 + *((_QWORD *)this + 374) + 56);
        if ( v79 )
        {
          result = DXGADAPTER::CreateFrameBufferSaveAreaSection(this, v164, v79);
          if ( (int)result < 0 )
            return result;
          v78 = v164;
        }
        v164 = v78 + 1;
        if ( v78 + 1 >= v53 )
          goto LABEL_146;
      }
      WdLogSingleEntry1(2LL, v76);
      v86 = L"Failed to query frame buffer save area size. Status 0x%I64x";
      WdLogGlobalForLineNumber = 7487;
      goto LABEL_160;
    }
  }
LABEL_146:
  if ( v56 )
  {
LABEL_147:
    if ( v53 )
    {
      v80 = v53;
      do
      {
        v81 = *((_QWORD *)this + 374);
        *(_WORD *)(v45 + v81) = *((_WORD *)this + 1238);
        v82 = *(_DWORD *)(v45 + v81 + 16) ^ ((unsigned __int8)*(_DWORD *)(v45 + v81 + 16) ^ (unsigned __int8)(*((_DWORD *)this + 617) >> 7)) & 1;
        *(_DWORD *)(v45 + v81 + 16) = v82;
        v83 = v82 ^ (v82 ^ (*((_DWORD *)this + 617) >> 5)) & 2;
        *(_DWORD *)(v45 + v81 + 16) = v83;
        v84 = v83 ^ ((unsigned __int8)v83 ^ (unsigned __int8)(DXGADAPTER::IsGpuVaIoMmuSupported(this) << 6)) & 0x40;
        *(_DWORD *)(v45 + v81 + 16) = v84;
        IsGpuVaIoMmuGlobalSupported = DXGADAPTER::IsGpuVaIoMmuGlobalSupported(this);
        *(_DWORD *)(v45 + v81 + 16) = v84 ^ ((unsigned __int8)v84 ^ (unsigned __int8)(IsGpuVaIoMmuGlobalSupported << 7)) & 0x80;
        *(_WORD *)(v45 + v81 + 2) = *((_WORD *)this + 1236);
        *(_QWORD *)(v45 + v81 + 8) = *((_QWORD *)this + 27);
        if ( (((unsigned __int8)v84 ^ ((unsigned __int8)v84 ^ (unsigned __int8)(IsGpuVaIoMmuGlobalSupported << 7)) & 0x80) & 0xC2) != 0 )
          *(_WORD *)(v45 + v81 + 48) = 257;
        v45 += 344LL;
        --v80;
      }
      while ( v80 );
      v37 = v171;
    }
  }
  if ( *((int *)this + 751) >= 2400 )
  {
    if ( *((_DWORD *)this + 744) >= 0x200u )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7547LL;
      v40 = L"Driver should not set MiscCaps.Reserved bits";
      goto LABEL_60;
    }
    *((_BYTE *)this + 3057) = *((_BYTE *)this + 2976) & 1;
  }
  v87 = *((_DWORD *)this + 744);
  if ( (v87 & 0x10) != 0 && !*((_QWORD *)this + 175) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7557LL;
    v40 = L"Driver sets IndependentVidPnVSyncControl cap but does not support DxgkDdiControlInterrupt3, returning failure";
    goto LABEL_60;
  }
  if ( *((_BYTE *)this + 3220) )
    *((_DWORD *)this + 744) = v87 & 0xFFFFFFEF;
  if ( v37 >= 0x3001 )
  {
    v89 = *((_DWORD *)this + 684);
    if ( v89 != 4096
      && v89 != 4608
      && v89 != 4864
      && v89 != 0x2000
      && v89 != 8448
      && v89 != 8704
      && v89 != 8960
      && v89 != 9216
      && v89 != 9472
      && v89 != 9728
      && v89 != 9984
      && v89 != 10240
      && v89 != 10496
      && v89 != 12288
      && v89 != 12544
      && v89 != 12800 )
    {
      WdLogSingleEntry1(2LL, *((int *)this + 684));
      v90 = *((int *)this + 684);
      v31 = L"Miniport returned unknown WDDM version 0x%I64x";
      v160 = 0LL;
      WdLogGlobalForLineNumber = 7615;
LABEL_203:
      v154 = v90;
      goto LABEL_40;
    }
  }
  else
  {
    *((_DWORD *)this + 684) = 4096;
  }
  if ( !*((_BYTE *)DXGGLOBAL::GetGlobal() + 888) || (v88 = 1, (*((_DWORD *)this + 111) & 8) != 0) )
    v88 = 0;
  *((_BYTE *)this + 3016) = v88;
  if ( v88 )
  {
    if ( *((int *)this + 684) < 4608
      && (*((_DWORD *)this + 732)
       || *((_DWORD *)this + 733)
       || *((_BYTE *)this + 2936)
       || *((_BYTE *)this + 2937)
       || *((_BYTE *)this + 2938)
       || (*((_DWORD *)this + 613) & 0x10000000) != 0
       || (*((_DWORD *)this + 616) & 0x14) != 0
       || *((_BYTE *)this + 2939)
       || *((_BYTE *)this + 2941)
       || *((_BYTE *)this + 2942)) )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7641LL;
      v40 = L"Driver reports WDDM version less than 1.2 but implements some WDDM 1.2 features.";
      goto LABEL_60;
    }
    v91 = *((_DWORD *)this + 684);
    if ( v91 >= 4864 )
    {
      if ( v91 >= 0x2000 )
        goto LABEL_213;
    }
    else if ( (*((_DWORD *)this + 615) & 0x10) != 0
           || (*((_DWORD *)this + 617) & 0x10) != 0
           || *((_BYTE *)this + 2943)
           || *((_DWORD *)this + 736) )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7656LL;
      v40 = L"Driver reports WDDM version less than 1.3 but implements some WDDM 1.3 features.";
      goto LABEL_60;
    }
    if ( *((_BYTE *)this + 2948) )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7684LL;
      v40 = L"Pre-WDDM 2.0 driver should not set the HybridIntegrated cap.";
      goto LABEL_60;
    }
  }
LABEL_213:
  v92 = *((_DWORD *)this + 617);
  if ( (v92 & 0x10000) != 0 )
  {
    if ( (*((_DWORD *)this + 617) & 0x8010) != 0x8010 )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7698LL;
      v40 = L"Driver reports CrossAdapterResourceScanout but does not report lower tier support.";
      goto LABEL_60;
    }
  }
  else if ( (v92 & 0x8000) != 0 && (v92 & 0x10) == 0 )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7706LL;
    v40 = L"Driver reports CrossAdapterResourceTexture but does not report lower tier support.";
    goto LABEL_60;
  }
  if ( v37 >= 0x4000 )
  {
    if ( v37 >= 0x5011 )
      goto LABEL_226;
  }
  else
  {
    v92 &= ~0x10u;
    *((_BYTE *)this + 2943) = 0;
    *((_DWORD *)this + 617) = v92;
  }
  v93 = *((_DWORD *)this + 111);
  if ( (v93 & 1) != 0 && (v92 & 0x10) != 0 && (v93 & 0x1000) != 0 )
    *((_BYTE *)this + 2948) = 1;
LABEL_226:
  v94 = v174[0];
  v95 = *(_BYTE *)v174[0] ^ (*(_BYTE *)v174[0] ^ (4 * *((_BYTE *)this + 2936))) & 4;
  *(_BYTE *)v174[0] = v95;
  v96 = v95 & 0xF7 | (*((_BYTE *)this + 2942) != 0 ? 8 : 0);
  *(_BYTE *)v94 = v96;
  v97 = v96 ^ (v96 ^ (32 * (*((_DWORD *)this + 617) >> 4))) & 0x20;
  *(_BYTE *)v94 = v97;
  v98 = v97 ^ (v97 ^ (*((_BYTE *)this + 2943) << 6)) & 0x40;
  *(_BYTE *)v94 = v98;
  *((_DWORD *)v94 + 1) = *((_DWORD *)this + 609);
  v99 = *((_BYTE *)v94 + 1) ^ (*((_BYTE *)this + 2948) ^ *((_BYTE *)v94 + 1)) & 1;
  *((_BYTE *)v94 + 1) = v99;
  *((_DWORD *)v94 + 2) = *((_DWORD *)this + 684);
  v100 = v99 ^ (v99 ^ (32 * (*((_DWORD *)this + 744) >> 3))) & 0x20;
  v101 = v98 & 0xEF;
  *((_BYTE *)v94 + 1) = v100;
  *(_BYTE *)v94 = v98 & 0xEF;
  if ( v37 >= 0x5021 )
  {
    v101 = v98 ^ (v98 ^ (16 * *((_BYTE *)this + 2968))) & 0x10;
    *(_BYTE *)v94 = v101;
  }
  if ( *((_BYTE *)this + 209) )
    goto LABEL_259;
  if ( (v101 & 0x40) != 0 )
  {
    if ( v37 < 0x5005 && (*((_DWORD *)this + 464) || *((_DWORD *)this + 465)) )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports device 0x%I64x is hybrid discrete device but it has VidPn source and target.";
      WdLogGlobalForLineNumber = 7769;
      goto LABEL_61;
    }
    v102 = v100 ^ (v100 ^ (2 * *((_BYTE *)this + 2971))) & 2;
    *((_BYTE *)v94 + 1) = v102;
    v103 = v102 & 1;
    goto LABEL_236;
  }
  v103 = v100 & 1;
  if ( (v100 & 1) != 0 )
  {
LABEL_236:
    if ( (v101 & 0x20) == 0 )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports device 0x%I64x as hybrid device but does not support cross adapter resource.";
      WdLogGlobalForLineNumber = 7783;
      goto LABEL_61;
    }
  }
  v104 = v101 & 0x40;
  if ( v103 )
  {
    if ( v104 )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports both HybridIntegrated and HybridDiscrete caps 0x%I64x";
      WdLogGlobalForLineNumber = 7790;
      goto LABEL_61;
    }
    if ( !*((_DWORD *)this + 465) )
    {
      WdLogSingleEntry1(2LL, *((_QWORD *)this + 27));
      v39 = *((_QWORD *)this + 27);
      v40 = L"Driver reports the HybridIntegrated cap, but the adapter has no outputs. 0x%I64x";
      WdLogGlobalForLineNumber = 7798;
      goto LABEL_61;
    }
  }
  if ( *((_BYTE *)this + 2938) && (!*((_QWORD *)this + 101) || !*((_QWORD *)this + 102) || !*((_QWORD *)this + 103)) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7812LL;
    v40 = L"Driver reports SupportPerEngineTDR cap but does not fill in all of the required DDIs.";
    goto LABEL_60;
  }
  if ( (*((_DWORD *)this + 613) & 4) != 0 && !*((_QWORD *)this + 100) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7819LL;
    v40 = L"Driver reports SupportKernelModeCommandBuffer cap but does not fill in the pfnRenderKm DDI.";
    goto LABEL_60;
  }
  if ( *((_BYTE *)this + 2941) && (!*((_QWORD *)this + 105) || !*((_QWORD *)this + 106)) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7827LL;
    v40 = L"Driver reports SupportRuntimePowerManagement cap but does not fill in the pfnSetPowerComponentFState or pfnPow"
           "erRuntimeControlRequest DDI.";
    goto LABEL_60;
  }
  if ( v37 < 0x300C && *((_QWORD *)this + 105) && *((_QWORD *)this + 106) )
    *((_BYTE *)this + 2941) = 1;
LABEL_259:
  *((_WORD *)this + 1509) = 0;
  *((_BYTE *)this + 3020) = 0;
  if ( !*((_BYTE *)this + 2940) )
    goto LABEL_297;
  if ( v37 < 0x300B )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7849LL;
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but it is not compiled with expected header files.";
    goto LABEL_60;
  }
  if ( v37 < 0x4000 )
  {
    *((_BYTE *)this + 3018) = 1;
    goto LABEL_279;
  }
  if ( v37 == 0x4000 )
  {
    *((_BYTE *)this + 3019) = 1;
    goto LABEL_279;
  }
  v105 = *((_DWORD *)this + 736);
  if ( !v105 )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7862LL;
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but doesn't report any overlay planes or panel fitter.";
    goto LABEL_60;
  }
  if ( v105 <= 8 )
  {
    if ( v37 > 0x5000 )
      *((_BYTE *)this + 3020) = 1;
    goto LABEL_279;
  }
  v106 = *((_DWORD *)this + 684);
  if ( v106 < 8704 )
  {
    if ( v106 < 0x2000 || v105 != 10 )
    {
      WdLogSingleEntry0(2LL);
      v39 = 7885LL;
      goto LABEL_272;
    }
    *((_DWORD *)this + 736) = 8;
  }
  else if ( v105 > 0xA )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7872LL;
LABEL_272:
    v40 = L"Driver reports more than the supported number of overlay planes.";
    goto LABEL_60;
  }
LABEL_279:
  v107 = *((_QWORD *)this + 109);
  if ( !v107 && !*((_QWORD *)this + 125) && !*((_QWORD *)this + 129) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7901LL;
LABEL_283:
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but does not fill in all of the required DDIs.";
    goto LABEL_60;
  }
  if ( v37 > 0x4002 && !*((_QWORD *)this + 113) && !*((_QWORD *)this + 124) && !*((_QWORD *)this + 128) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7913LL;
    goto LABEL_283;
  }
  if ( !*((_BYTE *)this + 2939) )
  {
    WdLogSingleEntry0(2LL);
    v39 = 7923LL;
    v40 = L"Driver reports SupportMultiPlaneOverlay cap but DirectFlip is not supported.";
    goto LABEL_60;
  }
  if ( v107 )
  {
    v108 = DXGGLOBAL::GetGlobal();
    DXGGLOBAL::RecordFeatureUsage(v108, 1LL, 1LL);
  }
  if ( *((_QWORD *)this + 125) )
  {
    v109 = DXGGLOBAL::GetGlobal();
    DXGGLOBAL::RecordFeatureUsage(v109, 2LL, 1LL);
  }
  if ( *((_QWORD *)this + 129) )
  {
    v110 = DXGGLOBAL::GetGlobal();
    DXGGLOBAL::RecordFeatureUsage(v110, 3LL, 1LL);
  }
LABEL_297:
  v111 = *((_BYTE *)this + 209);
  *((_BYTE *)this + 3055) = 0;
  if ( v111 )
    goto LABEL_310;
  v112 = 0;
  if ( v37 >= 0x700A && *((int *)this + 684) >= 8704 && (!*((_QWORD *)this + 83) || *((_QWORD *)this + 146)) )
  {
    *((_BYTE *)this + 3055) = 1;
    v112 = 1;
  }
  if ( *((int *)this + 684) < 8960 )
  {
LABEL_310:
    *((_DWORD *)this + 612) &= 0xFFFFFFE3;
  }
  else
  {
    v113 = (*((_DWORD *)this + 612) >> 3) & 1;
    v114 = (*((_DWORD *)this + 612) >> 2) & 1;
    if ( v114 < v113 || v113 < ((*((_DWORD *)this + 612) >> 4) & 1u) )
    {
      WdLogSingleEntry2(2LL, *((_QWORD *)this + 27), -1073741811LL);
      v116 = *((_QWORD *)this + 27);
      WdLogGlobalForLineNumber = 7973;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Driver reports support higher level of colorSpaceTransform but not lower levels on device 0x%I64x,"
                       " returning 0x%I64x.",
        v116,
        -1073741811LL,
        0LL,
        0LL,
        0LL);
      return 3221225485LL;
    }
    if ( !v112 && v114 )
    {
      WdLogSingleEntry2(2LL, *((_QWORD *)this + 27), -1073741811LL);
      v115 = *((_QWORD *)this + 27);
      WdLogGlobalForLineNumber = 7981;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"ColorSpaceTransform is supported on the device 0x%I64x which does not have pfnSetTargetGamma, returning 0x%I64x.",
        v115,
        -1073741811LL,
        0LL,
        0LL,
        0LL);
      return 3221225485LL;
    }
  }
  if ( !*(_BYTE *)(*(_QWORD *)(*(_QWORD *)(*((_QWORD *)this + 27) + 64LL) + 40LL) + 133LL) && !v111 )
  {
    v117 = *((_DWORD *)this + 684) >= 0x2000;
    v118 = DXGGLOBAL::GetGlobal();
    v119 = DXGGLOBAL::DeferredInitialize(v118, v117);
    RenderCore = v119;
    if ( v119 < 0 )
    {
      WdLogSingleEntry1(2LL, v119);
      v86 = L"DXGGLOBAL::DeferredInitialize failed (Status = 0x%I64x).";
      WdLogGlobalForLineNumber = 8008;
LABEL_160:
      DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v86, RenderCore, 0LL, 0LL, 0LL, 0LL);
      return (unsigned int)RenderCore;
    }
  }
  DXGADAPTER::Config = 0;
  DXGADAPTER::ReadConfig(this, v94);
  DXGADAPTER::InitializeDriverWorkarounds(this);
  if ( *((_BYTE *)this + 209) )
  {
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFFDFFFF | v170 & 0x20000;
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFFE7FFF | v170 & 0x18000;
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFEFFFFF | v170 & 0x100000;
    **((_DWORD **)this + 376) = **((_DWORD **)this + 376) & 0xFFF3FFFF | v170 & 0xC0000;
    *((_BYTE *)this + 3021) = 0;
  }
  else if ( (*((_DWORD *)this + 111) & 0x10) != 0 && *((_BYTE *)this + 3071) )
  {
    *((_DWORD *)this + 617) |= 0x400u;
  }
  v120 = *((_DWORD *)this + 684);
  if ( v120 < 9216 )
    goto LABEL_323;
  v121 = *((_QWORD *)this + 167);
  if ( *((_QWORD *)this + 166) )
  {
    if ( v121 )
      goto LABEL_324;
LABEL_335:
    WdLogSingleEntry0(2LL);
    v39 = 8062LL;
    v40 = L"Driver cannot support only one of DdiQueryDiagnosticTypesSupport and DdiControlDiagnosticReporting.";
    goto LABEL_60;
  }
  if ( v121 )
    goto LABEL_335;
LABEL_323:
  *((_QWORD *)this + 166) = W32kStub_UserRemoveWindowedSwapChain;
  *((_QWORD *)this + 167) = DXGADAPTER::DefaultDdiControlDiagnosticReporting;
LABEL_324:
  if ( v120 >= 12800 && v37 >= 0x11001 )
  {
    memset(&v176, 0, 24);
    v176.Type = DXGKQAITYPE_POWERCOMPONENTINFO|0x20;
    *(_OWORD *)&v176.OutputDataSize = 0LL;
    v176.pOutputData = (char *)this + 5088;
    v176.OutputDataSize = 4;
    if ( (int)DXGADAPTER::DdiQueryAdapterInfo(this, &v176) < 0 )
    {
      *(_QWORD *)(WdLogNewEntry5_WdTrace(v122) + 24) = this;
      WdLogGlobalForLineNumber = 8077;
    }
  }
  v168 = 0;
  memset(&v177, 0, 24);
  v177.Type = DXGKQAITYPE_PHYSICALADAPTERCAPS|0x20;
  v177.pOutputData = &v168;
  *(_OWORD *)&v177.OutputDataSize = 0LL;
  v177.OutputDataSize = 4;
  v123 = DXGADAPTER::DdiQueryAdapterInfo(this, &v177);
  v124 = *((_BYTE *)this + 3072) & 0xFD;
  if ( v123 >= 0 )
    v124 |= 2 * (v168 & 1);
  *((_BYTE *)this + 3072) = v124;
  result = DXGADAPTER::CheckMcdmDdiOverall(this);
  if ( (int)result >= 0 )
  {
    DXGADAPTER::InitializeDriverDiagnosticReporting(this);
    DXGADAPTER::QueryFeatureEnablement(this);
    if ( (*((_DWORD *)this + 616) & 0x800) != 0 )
    {
      if ( (*((_DWORD *)this + 1257) & 0x40) == 0 )
      {
        WdLogSingleEntry0(2LL);
        v39 = 8117LL;
        v40 = L"Driver reports NativeGpuFence cap when NativeFence feature is disabled, returning failure";
        goto LABEL_60;
      }
      v173.Type = DXGKQAITYPE_QUERYSEGMENT3|0x20;
      v173.pOutputData = (char *)this + 5032;
      v173.OutputDataSize = 56;
      v125 = DXGADAPTER::DdiQueryAdapterInfo(this, &v173);
      RenderCore = v125;
      if ( v125 < 0 )
      {
        WdLogSingleEntry1(2LL, v125);
        v86 = L"Failed to get DXGK_NATIVE_FENCE_CAPS. Status 0x%I64x";
        WdLogGlobalForLineNumber = 8128;
        goto LABEL_160;
      }
    }
    RenderCore = (int)ADAPTER_RENDER::CreateRenderCore(this, (struct ADAPTER_RENDER **)this + 391);
    v126 = *((_QWORD *)this + 391);
    if ( (int)RenderCore < 0 )
    {
      if ( v126 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 8140;
        DxgkLogInternalTriageEvent(0, 262146, -1, (unsigned int)L"m_pRenderCore == NULL", 8140LL, 0LL, 0LL, 0LL, 0LL);
      }
      WdLogSingleEntry2(2LL, this, RenderCore);
      WdLogGlobalForLineNumber = 8143;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Failed to create ADAPTER_RENDER on adapter 0x%I64x (Status = 0x%I64x).",
        (__int64)this,
        RenderCore,
        0LL,
        0LL,
        0LL);
      return (unsigned int)RenderCore;
    }
    if ( v126 )
    {
      if ( IsAdapterSessionized )
      {
        WdLogSingleEntry0(2LL);
        v31 = L"Render capable adapter should NOT be sessionized!";
        v90 = 8159LL;
        WdLogGlobalForLineNumber = 8159;
        v160 = 0LL;
        goto LABEL_203;
      }
      if ( (*((_DWORD *)this + 744) & 0xC) == 0 )
        *((_BYTE *)this + 3072) |= 1u;
    }
    v127 = (char *)this + 3120;
    DisplayCore = ADAPTER_DISPLAY::CreateDisplayCore(this, (struct ADAPTER_DISPLAY **)this + 390);
    RenderCore = DisplayCore;
    if ( DisplayCore < 0 )
    {
      if ( *(_QWORD *)v127 )
      {
        WdLogSingleEntry0(1LL);
        WdLogGlobalForLineNumber = 8174;
        DxgkLogInternalTriageEvent(0, 262146, -1, (unsigned int)L"m_pDisplayCore == NULL", 8174LL, 0LL, 0LL, 0LL, 0LL);
      }
      WdLogSingleEntry2(2LL, this, RenderCore);
      WdLogGlobalForLineNumber = 8177;
      DxgkLogInternalTriageEvent(
        0,
        0x40000,
        -1,
        (unsigned int)L"Failed to create ADAPTER_DISPLAY on adapter 0x%I64x (Status = 0x%I64x).",
        (__int64)this,
        RenderCore,
        0LL,
        0LL,
        0LL);
      return (unsigned int)RenderCore;
    }
    if ( *((_QWORD *)this + 391) )
    {
      v129 = *(_QWORD *)v127 == 0LL;
    }
    else
    {
      v129 = *(_QWORD *)v127 == 0LL;
      if ( !*(_QWORD *)v127 )
      {
        WdLogSingleEntry2(2LL, this, -1073741735LL);
        v31 = L"Current adapter 0x%I64x does not have display or render capabilities (Status = 0x%I64x).";
        v160 = -1073741735LL;
        v154 = (__int64)this;
        WdLogGlobalForLineNumber = 8190;
        goto LABEL_40;
      }
    }
    v130 = *(_BYTE *)v94 & 0xFE | !v129;
    *(_BYTE *)v94 = v130;
    v131 = v130 & 0xFD | (*((_QWORD *)this + 391) != 0LL ? 2 : 0);
    *(_BYTE *)v94 = v131;
    if ( *(_QWORD *)v127 )
      v132 = *(_DWORD *)(*(_QWORD *)v127 + 24LL);
    else
      LOBYTE(v132) = 0;
    v133 = v131 & 0x7F | ((_BYTE)v132 << 7);
    *(_BYTE *)v94 = v133;
    if ( (v133 & 1) != 0 )
      *((_BYTE *)v94 + 1) = *((_BYTE *)v94 + 1) & 0xFB | (DXGADAPTER::DriverSupportSetTimingsFromVidPn(this) != 0 ? 4 : 0);
    else
      *((_BYTE *)v94 + 1) &= ~4u;
    if ( !*((_QWORD *)this + 391) )
      *((_DWORD *)this + 613) |= 1u;
    if ( DXGADAPTER::IsDxgmms2(this) )
    {
      v136 = *((_DWORD *)this + 111);
      if ( (v136 & 4) == 0
        && (v136 & 8) == 0
        && v134
        && v37 >= 0x5008
        && (!*((_QWORD *)this + 114) || !*((_QWORD *)this + 126)) )
      {
        WdLogSingleEntry0(2LL);
        v39 = 8231LL;
        v40 = L"Driver is compiled against DXGKDDI_INTERFACE_VERSION_WDDM2_0_M2_2_1 or greater, but does not fill in the p"
               "fnCalibrateGpuClock or pfnSetStablePowerState DDI.";
        goto LABEL_60;
      }
    }
    if ( *((_BYTE *)this + 3016) && DXGADAPTER::IsFullWDDMAdapter(v135) && *((int *)this + 684) >= 4608 )
    {
      if ( !*((_BYTE *)this + 2939) )
      {
        WdLogSingleEntry0(2LL);
        v39 = 8246LL;
        v40 = L"Driver reports WDDM version 1.2 but does not implement all mandatory WDDM 1.2 full adapter features.";
        goto LABEL_60;
      }
    }
    else if ( !*((_BYTE *)this + 2939) )
    {
      goto LABEL_381;
    }
    if ( *((_BYTE *)this + 209) )
      goto LABEL_382;
    v137 = *((_QWORD *)this + 391);
    if ( !v137
      || !(*(unsigned __int8 (__fastcall **)(_QWORD))(*(_QWORD *)(*(_QWORD *)(v137 + 760) + 8LL) + 656LL))(*(_QWORD *)(v137 + 768)) )
    {
      *(_WORD *)((char *)this + 2939) = 0;
    }
LABEL_381:
    if ( !*((_BYTE *)this + 209) )
    {
LABEL_383:
      v138 = DXGADAPTER::IsBddFallbackDriver(this) != 0;
      v141 = *((_DWORD *)this + 111);
      *((_DWORD *)this + 50) = v138 ? 3 : 1;
      if ( (v141 & 0x10) != 0 && !*((_QWORD *)this + 390) )
      {
        DXGGLOBALSHAREMUTEX::DXGGLOBALSHAREMUTEX((DXGGLOBALSHAREMUTEX *)v174);
        DXGAUTOMUTEX::Acquire((DXGAUTOMUTEX *)v174);
        if ( *((_QWORD *)DXGGLOBAL::GetGlobal() + 119) )
        {
          WdLogSingleEntry2(2LL, this, -1073741735LL);
          WdLogGlobalForLineNumber = 8296;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Current adapter 0x%I64x does not have display or render capabilities (Status = 0x%I64x).",
            (__int64)this,
            -1073741735LL,
            0LL,
            0LL,
            0LL);
        }
        else
        {
          _InterlockedIncrement64((volatile signed __int64 *)this + 3);
          *((_QWORD *)this + 4) = -1LL;
          v142 = DXGGLOBAL::GetGlobal();
          DXGGLOBAL::SetWarpAdapter(v142, this);
        }
        DXGPROCESSCOPYPROTECTIONMUTEX::~DXGPROCESSCOPYPROTECTIONMUTEX((DXGPROCESSCOPYPROTECTIONMUTEX *)v174);
      }
      if ( *((_BYTE *)this + 209)
        || (v143 = DXGADAPTER::InitializePowerManagement(this, v139, v140), RenderCore = v143, v143 >= 0) )
      {
        if ( *((_BYTE *)this + 3016) )
        {
          if ( *((int *)this + 684) >= 4864 )
          {
            if ( DXGADAPTER::IsFullWDDMAdapter(this) )
            {
              v145 = *((_DWORD *)this + 111);
              if ( (v145 & 4) == 0 && (v145 & 0x20) == 0 && (*((_DWORD *)this + 615) & 0x10) == 0 )
              {
                WdLogSingleEntry0(2LL);
                v39 = 8327LL;
                v40 = L"WDDM 1.3 driver must support independent flip.";
                goto LABEL_60;
              }
            }
          }
        }
      }
      else
      {
        WdLogSingleEntry2(2LL, this, v143);
        WdLogGlobalForLineNumber = 8314;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Failed to initialize power management for the adapter 0x%I64x (Status = 0x%I64x).",
          (__int64)this,
          RenderCore,
          0LL,
          0LL,
          0LL);
      }
      if ( (*((_DWORD *)this + 111) & 0x10) != 0 )
        *((_BYTE *)this + 3058) = 1;
      if ( v37 >= 0xA008 )
        *((_BYTE *)this + 3058) = 1;
      v144 = operator new(40LL, 1265072196LL);
      if ( v144 )
      {
        *(_OWORD *)v144 = 0LL;
        *(_OWORD *)(v144 + 16) = 0LL;
        *(_QWORD *)(v144 + 32) = 0LL;
      }
      else
      {
        v144 = 0LL;
      }
      *((_QWORD *)this + 621) = v144;
      if ( !v144 )
      {
        WdLogSingleEntry0(2LL);
        v25 = L"Failed to allocate MockDriverState object";
        v24 = 8365LL;
        WdLogGlobalForLineNumber = 8365;
        v159 = 0LL;
        goto LABEL_31;
      }
      LocallyUniqueId = MOCKDRIVERSTATE::Initialize((MOCKDRIVERSTATE *)v144, this);
      if ( LocallyUniqueId < 0 )
      {
        WdLogSingleEntry0(2LL);
        v11 = 8372LL;
        v12 = L"Failed to initialize MockDriverState object";
        v13 = 0x40000;
        goto LABEL_13;
      }
      *((_BYTE *)this + 4976) = 0;
      LocallyUniqueId = DXGADAPTER::InitializeVSyncPhaseState(this);
      if ( LocallyUniqueId < 0 )
      {
        WdLogSingleEntry0(6LL);
        v11 = 8385LL;
        v12 = L"Failed to allocate VSync Phase Timer state";
        goto LABEL_12;
      }
      if ( (int)DXGADAPTER::InitializeCABCStateV2(v146) < 0 )
      {
        WdLogSingleEntry0(2LL);
        WdLogGlobalForLineNumber = 8400;
        DxgkLogInternalTriageEvent(
          0,
          0x40000,
          -1,
          (unsigned int)L"Failed to initialize CABC State",
          8400LL,
          0LL,
          0LL,
          0LL,
          0LL);
      }
      v147 = *((_QWORD *)this + 391);
      if ( v147 && !*((_BYTE *)this + 209) )
      {
        v148 = *(_QWORD *)(v147 + 736);
        v149 = DXGGLOBAL::GetGlobal();
        (*(void (__fastcall **)(_QWORD, __int64))(*(_QWORD *)(v148 + 8) + 920LL))(
          *(_QWORD *)(v147 + 744),
          (__int64)v149 + 1328);
      }
      if ( (*((_DWORD *)this + 111) & 1) != 0 )
        *((_QWORD *)DXGGLOBAL::GetGlobal() + 123) = *(_QWORD *)((char *)this + 412);
      if ( (int)RenderCore < 0 )
        return (unsigned int)RenderCore;
      if ( v169 <= 1 )
        goto LABEL_426;
      v150 = *((_DWORD *)this + 105);
      if ( v150 == 4318 )
      {
        v151 = DXGGLOBAL::GetGlobal();
        v152 = 7LL;
      }
      else
      {
        if ( v150 != 4098 )
        {
LABEL_426:
          v153 = DXGGLOBAL::GetGlobal();
          DXGGLOBAL::RecordFeatureUsageWddmVersion(v153, this);
          return (unsigned int)RenderCore;
        }
        v151 = DXGGLOBAL::GetGlobal();
        v152 = 8LL;
      }
      DXGGLOBAL::RecordFeatureUsage(v151, v152, 1LL);
      goto LABEL_426;
    }
LABEL_382:
    *((_QWORD *)this + 114) = 0LL;
    goto LABEL_383;
  }
  return result;
}

bool _TdrIsTestMode(void)
{
  int v0; // eax
  __int64 v2; // [rsp+30h] [rbp-19h] BYREF
  int v3; // [rsp+38h] [rbp-11h]
  const wchar_t *v4; // [rsp+40h] [rbp-9h]
  int *v5; // [rsp+48h] [rbp-1h]
  int v6; // [rsp+50h] [rbp+7h]
  int *v7; // [rsp+58h] [rbp+Fh]
  int v8; // [rsp+60h] [rbp+17h]
  __int64 v9; // [rsp+68h] [rbp+1Fh]
  int v10; // [rsp+70h] [rbp+27h]
  __int64 v11; // [rsp+78h] [rbp+2Fh]
  __int128 v12; // [rsp+80h] [rbp+37h]
  __int128 v13; // [rsp+90h] [rbp+47h]
  int v14; // [rsp+B0h] [rbp+67h] BYREF
  int v15; // [rsp+B8h] [rbp+6Fh] BYREF

  v15 = 0;
  v14 = 0;
  v2 = 0LL;
  v9 = 0LL;
  v10 = 0;
  v11 = 0LL;
  v4 = L"TdrTestMode";
  v5 = &v14;
  v7 = &v15;
  v3 = 288;
  v6 = 67108868;
  v8 = 4;
  v12 = 0LL;
  v13 = 0LL;
  if ( (int)RtlQueryRegistryValuesEx(
              0LL,
              L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\GraphicsDrivers",
              &v2) >= 0 )
    v0 = v14;
  else
    v0 = 0;
  return v0 != 0;
}

NTSTATUS __stdcall DriverEntry(_DRIVER_OBJECT *DriverObject, PUNICODE_STRING RegistryPath)
{
  NTSTATUS result; // eax
  int v5; // eax
  __int64 v6; // rdi
  const wchar_t *v7; // r9
  NTSTATUS v8; // eax
  int ProcessNotifyRoutineEx2; // eax
  __int64 v10; // rbx
  unsigned __int8 v11; // al
  __int64 v12; // rcx
  __int64 v13; // rcx
  __int64 v14; // rcx
  int v15; // eax
  __int64 v16; // rdx
  __int64 v17; // rbx
  __int64 v18; // rcx
  __int64 v19; // r8
  NTSTATUS v20; // eax
  int v21; // eax
  int v22; // eax
  int v23; // eax
  NTSTATUS v24; // eax
  __int64 v25; // rcx
  __int64 v26; // r8
  __int64 v27; // rcx
  __int64 v28; // r8
  BOOLEAN Size; // [rsp+28h] [rbp-D8h]
  unsigned int v30; // [rsp+50h] [rbp-B0h] BYREF
  __int64 v31; // [rsp+58h] [rbp-A8h]
  char v32; // [rsp+60h] [rbp-A0h]
  _QWORD v33[2]; // [rsp+68h] [rbp-98h] BYREF
  UNICODE_STRING DefaultSDDLString; // [rsp+78h] [rbp-88h] BYREF
  struct _UNICODE_STRING DestinationString; // [rsp+88h] [rbp-78h] BYREF
  __int64 v36; // [rsp+A0h] [rbp-60h] BYREF
  int v37; // [rsp+A8h] [rbp-58h]
  const wchar_t *v38; // [rsp+B0h] [rbp-50h]
  unsigned __int8 *v39; // [rsp+B8h] [rbp-48h]
  int v40; // [rsp+C0h] [rbp-40h]
  unsigned __int8 *v41; // [rsp+C8h] [rbp-38h]
  int v42; // [rsp+D0h] [rbp-30h]
  __int64 v43; // [rsp+D8h] [rbp-28h]
  int v44; // [rsp+E0h] [rbp-20h]
  __int64 v45; // [rsp+E8h] [rbp-18h]
  __int128 v46; // [rsp+F0h] [rbp-10h]
  __int128 v47; // [rsp+100h] [rbp+0h]
  __int64 SystemInformation; // [rsp+130h] [rbp+30h] BYREF

  g_pDriverObject = (PDEVICE_OBJECT)DriverObject;
  g_RegistryPath.Buffer = (wchar_t *)operator new[](RegistryPath->MaximumLength, 1265072196LL, 256LL);
  if ( !g_RegistryPath.Buffer )
  {
    WdLogSingleEntry0(2LL);
    WdLogGlobalForLineNumber = 298;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"Failed to allocate registry path buffer.",
      298LL,
      0LL,
      0LL,
      0LL,
      0LL);
    return -1073741801;
  }
  g_RegistryPath.MaximumLength = RegistryPath->MaximumLength;
  RtlCopyUnicodeString(&g_RegistryPath, RegistryPath);
  v5 = PsTlsAlloc(DxgkThreadPsTslCallback, 0LL, &g_DxgkThreadTlsId);
  v6 = v5;
  if ( v5 < 0 )
  {
    WdLogSingleEntry1(2LL, v5);
    v7 = L"Failed to allocate a PsTls slot for DxgkThread, returning 0x%I64x.";
    WdLogGlobalForLineNumber = 311;
LABEL_5:
    DxgkLogInternalTriageEvent(0, 0x40000, -1, (_DWORD)v7, v6, 0LL, 0LL, 0LL, 0LL);
    return v6;
  }
  v8 = ExInitializeLookasideListEx(&g_DxgkThreadLookasideList, 0LL, 0LL, (POOL_TYPE)512, 0, 0x40uLL, 0x54677844u, 0);
  v6 = v8;
  if ( v8 < 0 )
  {
    PsTlsFree(g_DxgkThreadTlsId);
    WdLogSingleEntry1(2LL, v6);
    v7 = L"Failed to initialize the lookaside list for DXGTHREAD, returning 0x%I64x";
    WdLogGlobalForLineNumber = 326;
    goto LABEL_5;
  }
  ProcessNotifyRoutineEx2 = PsSetCreateProcessNotifyRoutineEx2(0LL, DxgkProcessNotify, 0LL);
  if ( ProcessNotifyRoutineEx2 < 0 )
  {
    v10 = ProcessNotifyRoutineEx2;
    WdLogSingleEntry1(2LL, ProcessNotifyRoutineEx2);
    WdLogGlobalForLineNumber = 337;
    DxgkLogInternalTriageEvent(
      0,
      0x40000,
      -1,
      (unsigned int)L"PsSetCreateProcessNotifyRoutineEx failed 0x%I64x",
      v10,
      0LL,
      0LL,
      0LL,
      0LL);
  }
  SystemInformation = 8LL;
  if ( ZwQuerySystemInformation(MaxSystemInfoClass|SystemProcessInformation, &SystemInformation, 8u, 0LL) < 0
    || (v11 = 1, (SystemInformation & 0x200000000LL) == 0) )
  {
    v11 = 0;
  }
  g_OSTestSigningEnabled = v11;
  v36 = 0LL;
  v37 = 288;
  v40 = 67108868;
  v38 = L"IsInternalRelease";
  v42 = 4;
  v39 = &g_IsInternalRelease;
  v41 = &g_IsInternalRelease;
  v43 = 0LL;
  v44 = 0;
  v45 = 0LL;
  v46 = 0LL;
  v47 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v36);
  g_IsInternalRelease = g_IsInternalRelease != 0;
  g_IsInternalReleaseOrDbg = g_IsInternalRelease;
  g_bSkuSupportMultipleUsers = (RtlGetSuiteMask(v12) & 0x110) == 16;
  wil_InitializeFeatureStaging(v13);
  InitializeTelemetryAssertsKMByDriverObject(DriverObject);
  WdInitialize(v14);
  result = DXGGLOBAL::CreateGlobal();
  if ( result >= 0 )
  {
    result = DpiInitializeGlobalState();
    if ( result >= 0 )
    {
      result = CCD_BTL::CreateGlobal();
      if ( result >= 0 )
      {
        DxgkInitializeTelemetry();
        Size = 0;
        v15 = ExSubscribeWnfStateChange(&gScreenStudyEventSubscription, &WNF_SRUM_SCREENONSTUDY_SESSION, 1LL);
        if ( v15 < 0 )
        {
          v17 = v15;
          WdLogSingleEntry1(2LL, v15);
          WdLogGlobalForLineNumber = 447;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"ExSubscribeWnfStateChange failed, returing 0x%I64x",
            v17,
            0LL,
            0LL,
            0LL,
            0LL);
          gScreenStudyEventSubscription = 0LL;
        }
        bTracingEnabled = 0;
        McGenEventRegister_EtwRegister(&DxgkControlGuid, v16, &DxgkControlGuid_Context, &DxgkControlGuid_Context);
        v30 = -1;
        v31 = 0LL;
        if ( (qword_1C0151480 & 2) != 0 )
        {
          v32 = 1;
          v30 = 0;
          if ( (Microsoft_Windows_DxgKrnlEnableBits & 0x10000) != 0 )
            McTemplateK0q_EtwWriteTransfer(v18, &EventProfilerEnter, v19, 0LL);
        }
        else
        {
          v32 = 0;
        }
        DXGETWPROFILER_BASE_PushProfilerEntry(&v30, 0LL);
        v33[0] = &DxgkControlGuid;
        v33[1] = &Dxgk_WDI_NotifyUser;
        WdDiagInit(v33);
        DestinationString = 0LL;
        RtlInitUnicodeString(&DestinationString, L"\\Device\\DxgKrnl");
        DriverObject->MajorFunction[0] = (PDRIVER_DISPATCH)&DxgkCreateClose;
        DriverObject->MajorFunction[2] = (PDRIVER_DISPATCH)&DxgkCreateClose;
        DriverObject->MajorFunction[14] = (PDRIVER_DISPATCH)&DxgkDeviceIoctl;
        DriverObject->MajorFunction[15] = (PDRIVER_DISPATCH)&DxgkInternalDeviceIoctl;
        DriverObject->MajorFunction[16] = (PDRIVER_DISPATCH)&DxgkShutdown;
        DriverObject->DriverUnload = (PDRIVER_UNLOAD)DxgkUnload;
        DefaultSDDLString = 0LL;
        RtlInitUnicodeString(&DefaultSDDLString, L"D:P(A;;GRGW;;;S-1-5-83-0)");
        v20 = WdmlibIoCreateDeviceSecure(
                DriverObject,
                0,
                &DestinationString,
                0x22u,
                0x100u,
                Size,
                &DefaultSDDLString,
                &GUID_SD_DXGKRNL_DRIVER_OBJECT,
                &g_pDeviceObject);
        LODWORD(v6) = v20;
        if ( v20 < 0 )
        {
          WdLogSingleEntry1(3LL, v20);
          WdLogGlobalForLineNumber = 500;
LABEL_33:
          DxgkCleanupPower();
          MonitorCleanupGlobal();
          if ( g_pDeviceObject )
          {
            IoDeleteDevice(g_pDeviceObject);
            g_pDeviceObject = 0LL;
          }
          if ( g_RegistryPath.Buffer )
          {
            DXGQUOTAALLOCATOR<256,1835156294>::operator delete(g_RegistryPath.Buffer);
            g_RegistryPath = 0LL;
          }
          DXGGLOBAL::DestroyGlobal();
          PsTlsFree(g_DxgkThreadTlsId);
          ExDeleteLookasideListEx(&g_DxgkThreadLookasideList);
          DXGETWPROFILER_BASE::PopProfilerEntry((DXGETWPROFILER_BASE *)&v30);
          if ( v32 )
          {
            if ( (Microsoft_Windows_DxgKrnlEnableBits & 0x10000) != 0 )
              McTemplateK0q_EtwWriteTransfer(v25, &EventProfilerExit, v26, v30);
          }
          return v6;
        }
        v21 = DxgkInitialPower();
        LODWORD(v6) = v21;
        if ( v21 < 0 )
        {
          WdLogSingleEntry1(3LL, v21);
          WdLogGlobalForLineNumber = 513;
          goto LABEL_33;
        }
        v22 = MonitorInitializeGlobal();
        LODWORD(v6) = v22;
        if ( v22 < 0 )
        {
          WdLogSingleEntry1(3LL, v22);
          WdLogGlobalForLineNumber = 526;
          goto LABEL_33;
        }
        SysMmInitializeGlobal();
        DxgkInitTest();
        DxgDbgInit();
        TdrInit();
        v23 = SMgrRegisterSessionChangeCallout(DxgkNotifySessionStateChange);
        v6 = v23;
        if ( v23 < 0 )
        {
          WdLogSingleEntry1(2LL, v23);
          WdLogGlobalForLineNumber = 559;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Could not register session change callout with session manager, returning 0x%I64x.",
            v6,
            0LL,
            0LL,
            0LL,
            0LL);
          goto LABEL_33;
        }
        v24 = IoRegisterShutdownNotification(g_pDeviceObject);
        v6 = v24;
        if ( v24 < 0 )
        {
          WdLogSingleEntry1(2LL, v24);
          WdLogGlobalForLineNumber = 569;
          DxgkLogInternalTriageEvent(
            0,
            0x40000,
            -1,
            (unsigned int)L"Could not register for shutdown notification, returning 0x%I64x.",
            v6,
            0LL,
            0LL,
            0LL,
            0LL);
          SMgrUnregisterSessionChangeCallout(DxgkNotifySessionStateChange);
          goto LABEL_33;
        }
        DXGETWPROFILER_BASE::PopProfilerEntry((DXGETWPROFILER_BASE *)&v30);
        if ( v32 && (Microsoft_Windows_DxgKrnlEnableBits & 0x10000) != 0 )
          McTemplateK0q_EtwWriteTransfer(v27, &EventProfilerExit, v28, v30);
        return 0;
      }
    }
  }
  return result;
}

char DxgkpIsDrtEnabled()
{
  struct DXGPROCESS *Current; // rax
  char result; // al
  __int64 v2; // [rsp+30h] [rbp-19h] BYREF
  int v3; // [rsp+38h] [rbp-11h]
  const wchar_t *v4; // [rsp+40h] [rbp-9h]
  int *v5; // [rsp+48h] [rbp-1h]
  int v6; // [rsp+50h] [rbp+7h]
  int *v7; // [rsp+58h] [rbp+Fh]
  int v8; // [rsp+60h] [rbp+17h]
  __int64 v9; // [rsp+68h] [rbp+1Fh]
  int v10; // [rsp+70h] [rbp+27h]
  __int64 v11; // [rsp+78h] [rbp+2Fh]
  __int128 v12; // [rsp+80h] [rbp+37h]
  __int128 v13; // [rsp+90h] [rbp+47h]
  int v14; // [rsp+B0h] [rbp+67h] BYREF

  Current = DXGPROCESS::GetCurrent();
  if ( Current && (*((_DWORD *)Current + 102) & 0x1000) != 0 )
    return 1;
  v14 = 0;
  v2 = 0LL;
  v4 = L"DRTTestEnable";
  v9 = 0LL;
  v10 = 0;
  v5 = &v14;
  v11 = 0LL;
  v7 = &v14;
  v3 = 288;
  v6 = 67108868;
  v8 = 4;
  v12 = 0LL;
  v13 = 0LL;
  RtlQueryRegistryValuesEx(2LL, L"GraphicsDrivers", &v2);
  if ( v14 == 1484026436 )
    return 1;
  WdLogSingleEntry0(4LL);
  result = 0;
  WdLogGlobalForLineNumber = 50;
  return result;
}

void __fastcall DxgMonitor::MonitorColorState::OnFunctionDriverArrival(
        DxgMonitor::MonitorColorState *this,
        struct _DXGK_DISPLAY_SCENARIO_CONTEXT *a2)
{
  _QWORD *v4; // rcx
  bool v5; // r12
  bool v6; // cf
  int v7; // edi
  bool v8; // si
  __int64 v9; // rcx
  char v10; // al
  char v11; // di
  __int64 v12; // rcx
  __int64 v13; // rax
  __int64 v14; // rax
  __int64 v15; // rdx
  int v16; // ecx
  int v17; // r8d
  int v18; // r9d
  char v19; // si
  char v20; // r15
  _QWORD *v21; // rdi
  __int64 v22; // rax
  __int64 v23; // rax
  __int64 v24; // rcx
  __int64 v25; // rax
  int v26; // [rsp+40h] [rbp-10h] BYREF
  __int64 v27; // [rsp+48h] [rbp-8h] BYREF
  __int64 v28; // [rsp+90h] [rbp+40h] BYREF
  bool v29; // [rsp+A0h] [rbp+50h] BYREF
  int v30; // [rsp+A8h] [rbp+58h] BYREF

  v5 = DxgMonitor::MonitorColorState::EdidSupportsHDR(this);
  if ( !v5 || (*(unsigned __int8 (__fastcall **)(_QWORD))(*(_QWORD *)*v4 + 72LL))(*v4) )
    goto LABEL_29;
  v6 = *((_BYTE *)this + 361) != 0;
  LOBYTE(v28) = 0;
  v30 = 0;
  v7 = v6 ? 0x40000 : 0;
  DxgMonitor::MonitorColorState::_ReadDisplayHdrSupportFromPnpRegistry(
    this,
    (enum _DISPLAYCONFIG_HDR_CERTIFICATIONS *)&v30,
    (bool *)&v28);
  v8 = v30 >= 0 && ((v30 & 0x40000000) != 0 || (v30 & 0x20000000) != 0);
  v9 = *((_QWORD *)this + 1);
  v29 = 0;
  *((_DWORD *)this + 104) = v30 | v7;
  v10 = (*(__int64 (__fastcall **)(__int64, __int64, const wchar_t *, bool *))(*(_QWORD *)v9 + 104LL))(
          v9,
          2LL,
          L"HDREnabled",
          &v29);
  v11 = v28;
  if ( v10
    || (*(unsigned __int8 (__fastcall **)(_QWORD, __int64, const wchar_t *, bool *))(**((_QWORD **)this + 1) + 104LL))(
         *((_QWORD *)this + 1),
         2LL,
         L"AdvancedColorEnabled",
         &v29) )
  {
    DxgMonitor::MonitorColorState::SetHdrEnabled(this, v29);
    goto LABEL_24;
  }
  if ( (*(unsigned __int8 (__fastcall **)(_QWORD))(**(_QWORD **)this + 56LL))(*(_QWORD *)this) )
  {
    v12 = *((_QWORD *)this + 1);
    LOBYTE(v28) = 0;
    if ( !(*(unsigned __int8 (__fastcall **)(__int64, __int64, const wchar_t *, __int64 *))(*(_QWORD *)v12 + 104LL))(
            v12,
            1LL,
            L"EnableIntegratedPanelHdrByDefault",
            &v28) )
      (*(void (__fastcall **)(_QWORD, __int64, const wchar_t *, __int64 *))(**((_QWORD **)this + 1) + 104LL))(
        *((_QWORD *)this + 1),
        8LL,
        L"EnableIntegratedPanelHdrByDefault",
        &v28);
    if ( !*((_BYTE *)this + 404) && (_BYTE)v28 )
    {
      DxgMonitor::MonitorColorState::SetHdrEnabled(this, 1);
      v13 = (*(__int64 (__fastcall **)(_QWORD))(**(_QWORD **)this + 32LL))(*(_QWORD *)this);
      (*(void (__fastcall **)(__int64, _QWORD, struct _DXGK_DISPLAY_SCENARIO_CONTEXT *))(*(_QWORD *)v13 + 112LL))(
        v13,
        0LL,
        a2);
    }
    goto LABEL_24;
  }
  if ( v8 || *((_BYTE *)this + 361) )
  {
    DxgMonitor::MonitorColorState::SetHdrEnabled(this, 1);
    v14 = (*(__int64 (__fastcall **)(_QWORD))(**(_QWORD **)this + 32LL))(*(_QWORD *)this);
    (*(void (__fastcall **)(__int64, _QWORD, struct _DXGK_DISPLAY_SCENARIO_CONTEXT *))(*(_QWORD *)v14 + 112LL))(
      v14,
      0LL,
      a2);
    v15 = 10LL;
LABEL_19:
    WdDiagNotifyUser(0LL, v15, 0LL, 0LL);
    goto LABEL_24;
  }
  if ( *((_DWORD *)this + 104) && v11 )
  {
    v15 = 11LL;
    goto LABEL_19;
  }
LABEL_24:
  if ( *((_DWORD *)this + 104)
    && v11
    && (unsigned int)dword_1C0151620 > 5
    && (unsigned __int8)tlgKeywordOn(&dword_1C0151620, 0x400000200000LL) )
  {
    v30 = v8;
    v26 = v18;
    LOWORD(v28) = 2;
    v27 = 0x2000000LL;
    _tlgWriteTemplate<long (_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),&long _tlgWriteTransfer_EtwWriteTransfer(_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),_GUID const *,_GUID const *>::Write<_tlgWrapperByVal<8>,_tlgWrapperByVal<2>,_tlgWrapperByVal<4>,_tlgWrapperByVal<4>>(
      v16,
      (unsigned int)&unk_1C0134562,
      v17,
      v18,
      (__int64)&v27,
      (__int64)&v28,
      (__int64)&v26,
      (__int64)&v30);
  }
LABEL_29:
  if ( !(*(unsigned __int8 (__fastcall **)(_QWORD))(**(_QWORD **)this + 72LL))(*(_QWORD *)this) )
  {
    v19 = 0;
    v20 = 0;
    v21 = (_QWORD *)((char *)this + 8);
    if ( (*(unsigned __int8 (__fastcall **)(_QWORD))(**(_QWORD **)this + 56LL))(*(_QWORD *)this)
      && (v19 = (*(__int64 (__fastcall **)(_QWORD, __int64, const wchar_t *))(*(_QWORD *)*v21 + 56LL))(
                  *v21,
                  1LL,
                  L"MicrosoftApprovedAcmSupport")) != 0 )
    {
      v20 = (*(__int64 (__fastcall **)(_QWORD, __int64, const wchar_t *))(*(_QWORD *)*v21 + 56LL))(
              *v21,
              1LL,
              L"EnableIntegratedPanelAcmByDefault");
    }
    else
    {
      v22 = (*(__int64 (__fastcall **)(_QWORD))(**(_QWORD **)this + 32LL))(*(_QWORD *)this);
      if ( (*(int (__fastcall **)(__int64))(*(_QWORD *)v22 + 8LL))(v22) >= 3000 )
        v19 = (*(__int64 (__fastcall **)(_QWORD, __int64, const wchar_t *))(*(_QWORD *)*v21 + 56LL))(
                *v21,
                8LL,
                L"EnableAcmSupportDeveloperPreview");
      if ( !v19 )
      {
        v28 = (unsigned int)Feature_AutoColorManagement_WideRollout__private_featureState;
        if ( (Feature_AutoColorManagement_WideRollout__private_featureState & 0x10) == 0 )
        {
          LODWORD(v28) = Feature_AutoColorManagement_WideRollout__private_featureState | 1;
          wil_details_FeatureReporting_ReportUsageToService(
            &Feature_AutoColorManagement_WideRollout__private_descriptor,
            v28,
            3LL);
          wil_details_FeatureStateCache_TryEnableDeviceUsageFastPath(
            v28,
            3LL,
            &Feature_AutoColorManagement_WideRollout__private_descriptor);
          v21 = (_QWORD *)((char *)this + 8);
        }
        v23 = (*(__int64 (__fastcall **)(_QWORD))(**(_QWORD **)this + 32LL))(*(_QWORD *)this);
        if ( (*(int (__fastcall **)(__int64))(*(_QWORD *)v23 + 8LL))(v23) >= 3000
          && !(*(unsigned __int8 (__fastcall **)(_QWORD))(**(_QWORD **)this + 80LL))(*(_QWORD *)this) )
        {
          v19 = 1;
        }
      }
    }
    DxgMonitor::MonitorColorState::SetWcgPolicySupported(this, v19);
    if ( v19 )
    {
      v24 = *v21;
      LOBYTE(v28) = 0;
      if ( (*(unsigned __int8 (__fastcall **)(__int64, __int64, const wchar_t *, __int64 *))(*(_QWORD *)v24 + 104LL))(
             v24,
             2LL,
             L"AutoColorManagementEnabled",
             &v28)
        || !v5
        && (*(unsigned __int8 (__fastcall **)(_QWORD, __int64, const wchar_t *, __int64 *))(*(_QWORD *)*v21 + 104LL))(
             *v21,
             2LL,
             L"AdvancedColorEnabled",
             &v28)
        || (*(unsigned __int8 (__fastcall **)(_QWORD, __int64, const wchar_t *, __int64 *))(*(_QWORD *)*v21 + 104LL))(
             *v21,
             1LL,
             L"EnableIntegratedPanelAcmByDefault",
             &v28)
        || (*(unsigned __int8 (__fastcall **)(_QWORD, __int64, const wchar_t *, __int64 *))(*(_QWORD *)*v21 + 104LL))(
             *v21,
             8LL,
             L"EnableIntegratedPanelAcmByDefault",
             &v28) )
      {
        DxgMonitor::MonitorColorState::SetWcgEnabled(this, v28);
      }
      else if ( v20 )
      {
        DxgMonitor::MonitorColorState::SetWcgEnabled(this, 1);
        if ( !*((_BYTE *)this + 404) )
        {
          v25 = (*(__int64 (__fastcall **)(_QWORD))(**(_QWORD **)this + 32LL))(*(_QWORD *)this);
          (*(void (__fastcall **)(__int64, _QWORD, struct _DXGK_DISPLAY_SCENARIO_CONTEXT *))(*(_QWORD *)v25 + 112LL))(
            v25,
            0LL,
            a2);
        }
      }
    }
    else
    {
      *((_BYTE *)this + 405) = 0;
    }
  }
}