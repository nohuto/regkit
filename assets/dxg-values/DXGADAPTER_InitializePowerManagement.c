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
