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
