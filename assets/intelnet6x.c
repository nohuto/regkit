void __fastcall POWER::PwrReadRegistryParameters(POWER *this, REGISTRY **a2, void *a3)
{
  REGISTRY *v5; // rcx
  int v6; // [rsp+50h] [rbp-B0h] BYREF
  const wchar_t *v7; // [rsp+58h] [rbp-A8h]
  __int64 v8; // [rsp+60h] [rbp-A0h]
  int v9; // [rsp+68h] [rbp-98h]
  __int64 v10; // [rsp+6Ch] [rbp-94h]
  __int64 v11; // [rsp+74h] [rbp-8Ch]
  __int16 v12; // [rsp+7Ch] [rbp-84h]
  int v13; // [rsp+80h] [rbp-80h]
  const wchar_t *v14; // [rsp+88h] [rbp-78h]
  __int64 v15; // [rsp+90h] [rbp-70h]
  int v16; // [rsp+98h] [rbp-68h]
  __int64 v17; // [rsp+9Ch] [rbp-64h]
  __int64 v18; // [rsp+A4h] [rbp-5Ch]
  __int16 v19; // [rsp+ACh] [rbp-54h]
  int v20; // [rsp+B0h] [rbp-50h]
  const wchar_t *v21; // [rsp+B8h] [rbp-48h]
  __int64 v22; // [rsp+C0h] [rbp-40h]
  int v23; // [rsp+C8h] [rbp-38h]
  __int64 v24; // [rsp+CCh] [rbp-34h]
  __int64 v25; // [rsp+D4h] [rbp-2Ch]
  __int16 v26; // [rsp+DCh] [rbp-24h]
  int v27; // [rsp+E0h] [rbp-20h]
  const wchar_t *v28; // [rsp+E8h] [rbp-18h]
  __int64 v29; // [rsp+F0h] [rbp-10h]
  int v30; // [rsp+F8h] [rbp-8h]
  __int64 v31; // [rsp+FCh] [rbp-4h]
  int v32; // [rsp+104h] [rbp+4h]
  int v33; // [rsp+108h] [rbp+8h]
  __int16 v34; // [rsp+10Ch] [rbp+Ch]
  int v35; // [rsp+110h] [rbp+10h]
  const wchar_t *v36; // [rsp+118h] [rbp+18h]
  __int64 v37; // [rsp+120h] [rbp+20h]
  int v38; // [rsp+128h] [rbp+28h]
  __int64 v39; // [rsp+12Ch] [rbp+2Ch]
  int v40; // [rsp+134h] [rbp+34h]
  int v41; // [rsp+138h] [rbp+38h]
  __int16 v42; // [rsp+13Ch] [rbp+3Ch]
  int v43; // [rsp+140h] [rbp+40h]
  const wchar_t *v44; // [rsp+148h] [rbp+48h]
  __int64 v45; // [rsp+150h] [rbp+50h]
  int v46; // [rsp+158h] [rbp+58h]
  __int64 v47; // [rsp+15Ch] [rbp+5Ch]
  __int64 v48; // [rsp+164h] [rbp+64h]
  __int16 v49; // [rsp+16Ch] [rbp+6Ch]
  int v50; // [rsp+170h] [rbp+70h]
  const wchar_t *v51; // [rsp+178h] [rbp+78h]
  __int64 v52; // [rsp+180h] [rbp+80h]
  int v53; // [rsp+188h] [rbp+88h]
  __int64 v54; // [rsp+18Ch] [rbp+8Ch]
  __int64 v55; // [rsp+194h] [rbp+94h]
  __int16 v56; // [rsp+19Ch] [rbp+9Ch]
  int v57; // [rsp+1A0h] [rbp+A0h]
  const wchar_t *v58; // [rsp+1A8h] [rbp+A8h]
  __int64 v59; // [rsp+1B0h] [rbp+B0h]
  int v60; // [rsp+1B8h] [rbp+B8h]
  __int64 v61; // [rsp+1BCh] [rbp+BCh]
  int v62; // [rsp+1C4h] [rbp+C4h]
  int v63; // [rsp+1C8h] [rbp+C8h]
  __int16 v64; // [rsp+1CCh] [rbp+CCh]
  int v65; // [rsp+1D0h] [rbp+D0h]
  const wchar_t *v66; // [rsp+1D8h] [rbp+D8h]
  __int64 v67; // [rsp+1E0h] [rbp+E0h]
  int v68; // [rsp+1E8h] [rbp+E8h]
  __int64 v69; // [rsp+1ECh] [rbp+ECh]
  __int64 v70; // [rsp+1F4h] [rbp+F4h]
  __int16 v71; // [rsp+1FCh] [rbp+FCh]
  int v72; // [rsp+200h] [rbp+100h]
  const wchar_t *v73; // [rsp+208h] [rbp+108h]
  __int64 v74; // [rsp+210h] [rbp+110h]
  int v75; // [rsp+218h] [rbp+118h]
  __int64 v76; // [rsp+21Ch] [rbp+11Ch]
  __int64 v77; // [rsp+224h] [rbp+124h]
  __int16 v78; // [rsp+22Ch] [rbp+12Ch]
  int v79; // [rsp+230h] [rbp+130h]
  const wchar_t *v80; // [rsp+238h] [rbp+138h]
  __int64 v81; // [rsp+240h] [rbp+140h]
  int v82; // [rsp+248h] [rbp+148h]
  __int64 v83; // [rsp+24Ch] [rbp+14Ch]
  __int64 v84; // [rsp+254h] [rbp+154h]
  __int16 v85; // [rsp+25Ch] [rbp+15Ch]
  int v86; // [rsp+260h] [rbp+160h]
  const wchar_t *v87; // [rsp+268h] [rbp+168h]
  __int64 v88; // [rsp+270h] [rbp+170h]
  int v89; // [rsp+278h] [rbp+178h]
  __int64 v90; // [rsp+27Ch] [rbp+17Ch]
  __int64 v91; // [rsp+284h] [rbp+184h]
  __int16 v92; // [rsp+28Ch] [rbp+18Ch]
  int v93; // [rsp+290h] [rbp+190h]
  const wchar_t *v94; // [rsp+298h] [rbp+198h]
  __int64 v95; // [rsp+2A0h] [rbp+1A0h]
  int v96; // [rsp+2A8h] [rbp+1A8h]
  __int64 v97; // [rsp+2ACh] [rbp+1ACh]
  __int64 v98; // [rsp+2B4h] [rbp+1B4h]
  __int16 v99; // [rsp+2BCh] [rbp+1BCh]
  int v100; // [rsp+2C0h] [rbp+1C0h]
  const wchar_t *v101; // [rsp+2C8h] [rbp+1C8h]
  __int64 v102; // [rsp+2D0h] [rbp+1D0h]
  int v103; // [rsp+2D8h] [rbp+1D8h]
  __int64 v104; // [rsp+2DCh] [rbp+1DCh]
  int v105; // [rsp+2E4h] [rbp+1E4h]
  int v106; // [rsp+2E8h] [rbp+1E8h]
  __int16 v107; // [rsp+2ECh] [rbp+1ECh]

  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16306),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"EnablePowerManagement",
    0,
    1u,
    1u,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16307),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"ULPMode",
    0,
    1u,
    1u,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16308),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"SidebandUngateOverride",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16309),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"I218DisablePLLShut",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16310),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"I218DisablePLLShutGiga",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16311),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"I219DisableK1Off",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16312),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"DisableIntelRST",
    0,
    1u,
    1u,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16313),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"ForceHostExitUlp",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<unsigned int>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16315),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"ForceLtrValue",
    0,
    0xFFFFu,
    0xFFFFu,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16318),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"WakeOnLink",
    0,
    2u,
    0,
    0,
    1);
  REGKEY<unsigned int>::Initialize(
    (enum _REGKEY_STATE *)((char *)a2 + 130564),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"WakeFromS5",
    0,
    0xFFFFu,
    2u,
    0,
    1);
  REGKEY<unsigned int>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16319),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"WakeOn",
    0,
    4u,
    0,
    1,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16294),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"*WakeOnPattern",
    0,
    1u,
    1u,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16295),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"*WakeOnMagicPacket",
    0,
    1u,
    1u,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16296),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"EnableDisconnectedStandby",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16297),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"*EnableDynamicPowerGating",
    0,
    1u,
    1u,
    0,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)(a2 + 16298),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"EnableHWAutonomous",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<short>::Initialize(
    (enum _REGKEY_STATE *)((char *)a2 + 130532),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"DMACoalescing",
    0,
    0x2800u,
    0,
    1,
    1);
  REGKEY<unsigned char>::Initialize(
    (enum _REGKEY_STATE *)((char *)a2 + 130436),
    (struct ADAPTER_CONTEXT *)a2,
    a3,
    (PUCHAR)"EnablePME",
    0,
    1u,
    0,
    0,
    1);
  v6 = 3538996;
  v8 = 0LL;
  v7 = L"EnableWakeOnManagmentOnTCO";
  v9 = 130444;
  v14 = L"EnablePHYWakeUp";
  v21 = L"EnableD0PHYFlexibleSpeed";
  v28 = L"EnablePHYFlexibleSpeed";
  v36 = L"EnableSavePowerNow";
  v44 = L"AutoPowerSaveModeEnabled";
  v10 = 1LL;
  v11 = 1LL;
  v12 = 256;
  v13 = 2097182;
  v15 = 0LL;
  v16 = 129784;
  v17 = 1LL;
  v18 = 1LL;
  v19 = 256;
  v20 = 3276848;
  v22 = 0LL;
  v23 = 130576;
  v24 = 4LL;
  v25 = 2LL;
  v26 = 257;
  v27 = 3014700;
  v29 = 0LL;
  v30 = 130580;
  v31 = 4LL;
  v32 = 2;
  v33 = 1;
  v34 = 257;
  v35 = 2490404;
  v37 = 0LL;
  v38 = 129789;
  v39 = 1LL;
  v40 = 2;
  v41 = 1;
  v42 = 256;
  v43 = 3276848;
  v45 = 0LL;
  v5 = a2[14912];
  v51 = L"SipsEnabled";
  v47 = 1LL;
  v58 = L"SipsThreshold";
  v66 = L"DisableSMBusMode";
  v73 = L"EnableDeviceBusPowerStateDependency";
  v80 = L"WakeOnFastStartup";
  v87 = L"*PMARPOffload";
  v94 = L"*PMNSOffload";
  v48 = 1LL;
  v54 = 1LL;
  v55 = 1LL;
  v63 = 50;
  v69 = 1LL;
  v70 = 1LL;
  v76 = 1LL;
  v77 = 1LL;
  v83 = 1LL;
  v84 = 1LL;
  v90 = 1LL;
  v91 = 1LL;
  v97 = 1LL;
  v98 = 1LL;
  v101 = L"ProtocolOffloadLinkDownTimer";
  v46 = 129788;
  v49 = 256;
  v50 = 1572886;
  v52 = 0LL;
  v53 = 130592;
  v56 = 256;
  v57 = 1835034;
  v59 = 0LL;
  v60 = 130600;
  v61 = 4LL;
  v62 = 0xFFFF;
  v64 = 257;
  v65 = 2228256;
  v67 = 0LL;
  v68 = 130800;
  v71 = 256;
  v72 = 4718662;
  v74 = 0LL;
  v75 = 130803;
  v78 = 256;
  v79 = 2359330;
  v81 = 0LL;
  v82 = 130804;
  v85 = 256;
  v86 = 1835034;
  v88 = 0LL;
  v89 = 130952;
  v92 = 256;
  v93 = 1703960;
  v95 = 0LL;
  v96 = 130953;
  v99 = 256;
  v100 = 3801144;
  v102 = 0LL;
  v103 = 132352;
  v104 = 4LL;
  v105 = 0xFFFF;
  v106 = 16;
  v107 = 257;
  REGISTRY::RegReadRegTable(v5, (struct ADAPTER_CONTEXT *)a2, a3, (struct REGTABLE_ENTRY *)&v6, 0xEu);
}

void __fastcall REGISTRY::RegReadGeneralParameters(
    REGISTRY *this,
    NDIS_HANDLE *DstBuf,
    NDIS_HANDLE ConfigurationHandle)
{
REGISTRY *v6; // rcx
int v7; // [rsp+50h] [rbp-B0h] BYREF
const wchar_t *v8; // [rsp+58h] [rbp-A8h]
__int64 v9; // [rsp+60h] [rbp-A0h]
int v10; // [rsp+68h] [rbp-98h]
__int64 v11; // [rsp+6Ch] [rbp-94h]
__int64 v12; // [rsp+74h] [rbp-8Ch]
__int16 v13; // [rsp+7Ch] [rbp-84h]
int v14; // [rsp+80h] [rbp-80h]
const wchar_t *v15; // [rsp+88h] [rbp-78h]
__int64 v16; // [rsp+90h] [rbp-70h]
int v17; // [rsp+98h] [rbp-68h]
__int64 v18; // [rsp+9Ch] [rbp-64h]
int v19; // [rsp+A4h] [rbp-5Ch]
int v20; // [rsp+A8h] [rbp-58h]
__int16 v21; // [rsp+ACh] [rbp-54h]
int v22; // [rsp+B0h] [rbp-50h]
const wchar_t *v23; // [rsp+B8h] [rbp-48h]
__int64 v24; // [rsp+C0h] [rbp-40h]
int v25; // [rsp+C8h] [rbp-38h]
__int64 v26; // [rsp+CCh] [rbp-34h]
int v27; // [rsp+D4h] [rbp-2Ch]
int v28; // [rsp+D8h] [rbp-28h]
__int16 v29; // [rsp+DCh] [rbp-24h]
int v30; // [rsp+E0h] [rbp-20h]
const wchar_t *v31; // [rsp+E8h] [rbp-18h]
__int64 v32; // [rsp+F0h] [rbp-10h]
int v33; // [rsp+F8h] [rbp-8h]
__int64 v34; // [rsp+FCh] [rbp-4h]
__int64 v35; // [rsp+104h] [rbp+4h]
__int16 v36; // [rsp+10Ch] [rbp+Ch]
int v37; // [rsp+110h] [rbp+10h]
const wchar_t *v38; // [rsp+118h] [rbp+18h]
__int64 v39; // [rsp+120h] [rbp+20h]
int v40; // [rsp+128h] [rbp+28h]
int v41; // [rsp+12Ch] [rbp+2Ch]
int v42; // [rsp+130h] [rbp+30h]
int v43; // [rsp+134h] [rbp+34h]
int v44; // [rsp+138h] [rbp+38h]
__int16 v45; // [rsp+13Ch] [rbp+3Ch]
int v46; // [rsp+140h] [rbp+40h]
const wchar_t *v47; // [rsp+148h] [rbp+48h]
__int64 v48; // [rsp+150h] [rbp+50h]
int v49; // [rsp+158h] [rbp+58h]
int v50; // [rsp+15Ch] [rbp+5Ch]
int v51; // [rsp+160h] [rbp+60h]
int v52; // [rsp+164h] [rbp+64h]
int v53; // [rsp+168h] [rbp+68h]
__int16 v54; // [rsp+16Ch] [rbp+6Ch]
int v55; // [rsp+170h] [rbp+70h]
const wchar_t *v56; // [rsp+178h] [rbp+78h]
__int64 v57; // [rsp+180h] [rbp+80h]
int v58; // [rsp+188h] [rbp+88h]
int v59; // [rsp+18Ch] [rbp+8Ch]
int v60; // [rsp+190h] [rbp+90h]
int v61; // [rsp+194h] [rbp+94h]
int v62; // [rsp+198h] [rbp+98h]
__int16 v63; // [rsp+19Ch] [rbp+9Ch]
int v64; // [rsp+1A0h] [rbp+A0h]
const wchar_t *v65; // [rsp+1A8h] [rbp+A8h]
__int64 v66; // [rsp+1B0h] [rbp+B0h]
int v67; // [rsp+1B8h] [rbp+B8h]
__int64 v68; // [rsp+1BCh] [rbp+BCh]
int v69; // [rsp+1C4h] [rbp+C4h]
int v70; // [rsp+1C8h] [rbp+C8h]
__int16 v71; // [rsp+1CCh] [rbp+CCh]
int v72; // [rsp+1D0h] [rbp+D0h]
const wchar_t *v73; // [rsp+1D8h] [rbp+D8h]
__int64 v74; // [rsp+1E0h] [rbp+E0h]
int v75; // [rsp+1E8h] [rbp+E8h]
__int64 v76; // [rsp+1ECh] [rbp+ECh]
int v77; // [rsp+1F4h] [rbp+F4h]
int v78; // [rsp+1F8h] [rbp+F8h]
__int16 v79; // [rsp+1FCh] [rbp+FCh]
int v80; // [rsp+200h] [rbp+100h]
const wchar_t *v81; // [rsp+208h] [rbp+108h]
__int64 v82; // [rsp+210h] [rbp+110h]
int v83; // [rsp+218h] [rbp+118h]
__int64 v84; // [rsp+21Ch] [rbp+11Ch]
__int64 v85; // [rsp+224h] [rbp+124h]
__int16 v86; // [rsp+22Ch] [rbp+12Ch]
int v87; // [rsp+230h] [rbp+130h]
const wchar_t *v88; // [rsp+238h] [rbp+138h]
__int64 v89; // [rsp+240h] [rbp+140h]
int v90; // [rsp+248h] [rbp+148h]
__int64 v91; // [rsp+24Ch] [rbp+14Ch]
__int64 v92; // [rsp+254h] [rbp+154h]
__int16 v93; // [rsp+25Ch] [rbp+15Ch]
int v94; // [rsp+260h] [rbp+160h]
const wchar_t *v95; // [rsp+268h] [rbp+168h]
__int64 v96; // [rsp+270h] [rbp+170h]
int v97; // [rsp+278h] [rbp+178h]
__int64 v98; // [rsp+27Ch] [rbp+17Ch]
int v99; // [rsp+284h] [rbp+184h]
int v100; // [rsp+288h] [rbp+188h]
__int16 v101; // [rsp+28Ch] [rbp+18Ch]
int v102; // [rsp+290h] [rbp+190h]
const wchar_t *v103; // [rsp+298h] [rbp+198h]
__int64 v104; // [rsp+2A0h] [rbp+1A0h]
int v105; // [rsp+2A8h] [rbp+1A8h]
__int64 v106; // [rsp+2ACh] [rbp+1ACh]
__int64 v107; // [rsp+2B4h] [rbp+1B4h]
__int16 v108; // [rsp+2BCh] [rbp+1BCh]
int v109; // [rsp+2C0h] [rbp+1C0h]
const wchar_t *v110; // [rsp+2C8h] [rbp+1C8h]
__int64 v111; // [rsp+2D0h] [rbp+1D0h]
int v112; // [rsp+2D8h] [rbp+1D8h]
__int64 v113; // [rsp+2DCh] [rbp+1DCh]
__int64 v114; // [rsp+2E4h] [rbp+1E4h]
__int16 v115; // [rsp+2ECh] [rbp+1ECh]
int v116; // [rsp+2F0h] [rbp+1F0h]
const wchar_t *v117; // [rsp+2F8h] [rbp+1F8h]
__int64 v118; // [rsp+300h] [rbp+200h]
int v119; // [rsp+308h] [rbp+208h]
__int64 v120; // [rsp+30Ch] [rbp+20Ch]
__int64 v121; // [rsp+314h] [rbp+214h]
__int16 v122; // [rsp+31Ch] [rbp+21Ch]
int v123; // [rsp+320h] [rbp+220h]
const wchar_t *v124; // [rsp+328h] [rbp+228h]
__int64 v125; // [rsp+330h] [rbp+230h]
int v126; // [rsp+338h] [rbp+238h]
__int64 v127; // [rsp+33Ch] [rbp+23Ch]
__int64 v128; // [rsp+344h] [rbp+244h]
__int16 v129; // [rsp+34Ch] [rbp+24Ch]
int v130; // [rsp+350h] [rbp+250h]
const wchar_t *v131; // [rsp+358h] [rbp+258h]
__int64 v132; // [rsp+360h] [rbp+260h]
int v133; // [rsp+368h] [rbp+268h]
__int64 v134; // [rsp+36Ch] [rbp+26Ch]
__int64 v135; // [rsp+374h] [rbp+274h]
__int16 v136; // [rsp+37Ch] [rbp+27Ch]

REGKEY<unsigned char>::Initialize(
(enum _REGKEY_STATE *)(DstBuf + 17006),
(struct ADAPTER_CONTEXT *)DstBuf,
DstBuf[16971],
(PUCHAR)"DisableReset",
0,
1u,
0,
0,
0);
REGKEY<unsigned int>::Initialize(
(enum _REGKEY_STATE *)(DstBuf + 17007),
(struct ADAPTER_CONTEXT *)DstBuf,
DstBuf[16971],
(PUCHAR)"CheckForHangTime",
0,
0x3Cu,
2u,
1,
0);
REGKEY<unsigned int>::Initialize(
(enum _REGKEY_STATE *)(DstBuf + 17009),
(struct ADAPTER_CONTEXT *)DstBuf,
DstBuf[16971],
(PUCHAR)"ResetTest",
0,
1u,
0,
0,
0);
REGKEY<unsigned int>::Initialize(
(enum _REGKEY_STATE *)((char *)DstBuf + 136084),
(struct ADAPTER_CONTEXT *)DstBuf,
DstBuf[16971],
(PUCHAR)"ResetTestTime",
0x14u,
0x93A80u,
0x12Cu,
1,
0);
v7 = 2097182;
v12 = 3LL;
v27 = 3;
v8 = L"*IPsecOffloadV2";
v28 = 3;
v15 = L"IPSecurity";
v9 = 0LL;
v23 = L"*PriorityVLANTag";
v31 = L"VlanFiltering";
v10 = 119416;
v11 = 4LL;
v13 = 256;
v14 = 1441812;
v16 = 0LL;
v17 = 119420;
v18 = 1LL;
v19 = 1;
v20 = 1;
v21 = 256;
v22 = 2228256;
v24 = 0LL;
v25 = 132648;
v26 = 4LL;
v29 = 256;
v30 = 1835034;
v32 = 0LL;
v33 = 132652;
v34 = 4LL;
v48 = 0LL;
v51 = 0;
v53 = 0;
v57 = 0LL;
v60 = 0;
v62 = 0;
v38 = L"*JumboPacket";
v47 = L"PadReceiveBuffer";
v50 = 1;
v52 = 1;
v56 = L"StoreBadPackets";
v59 = 1;
v61 = 1;
v65 = L"LogLinkStateEvent";
v66 = 0LL;
v73 = L"CEMDriverVer";
v74 = 0LL;
v39 = 0LL;
v81 = L"EEPROMValidation";
v41 = 4;
v68 = 4LL;
v88 = L"EnableTxHangWA";
v35 = 2LL;
v95 = L"SVOFFMode";
v36 = 256;
v37 = 1703960;
v40 = 136044;
v42 = 1510;
v43 = 9014;
v44 = 1514;
v45 = 257;
v46 = 2228256;
v49 = 136109;
v54 = 256;
v55 = 2097182;
v58 = 136108;
v63 = 256;
v64 = 2359330;
v67 = 134024;
v69 = -1;
v70 = -1;
v71 = 256;
v72 = 1703960;
v75 = 155700;
v76 = 4LL;
v77 = -1;
v78 = -1;
v79 = 256;
v80 = 2228256;
v82 = 0LL;
v83 = 136112;
v84 = 4LL;
v85 = 2LL;
v86 = 256;
v87 = 1966108;
v89 = 0LL;
v90 = 155634;
v91 = 1LL;
v92 = 1LL;
v93 = 256;
v94 = 1310738;
v96 = 0LL;
v97 = 129675;
v98 = 1LL;
v99 = 1;
v6 = (REGISTRY *)DstBuf[14912];
v103 = L"SVOFFModeHWM";
v100 = 1;
v110 = L"SVOFFModeTimer";
v117 = L"Enable9KJFTpt";
v124 = L"OBFFEnabled";
v104 = 0LL;
v106 = 1LL;
v111 = 0LL;
v118 = 0LL;
v120 = 1LL;
v121 = 1LL;
v125 = 0LL;
v127 = 1LL;
v128 = 1LL;
v132 = 0LL;
v131 = L"TxDelay";
v101 = 256;
v102 = 1703960;
v105 = 129676;
v107 = 255LL;
v108 = 257;
v109 = 1966108;
v112 = 129678;
v113 = 2LL;
v114 = 0xFFFFLL;
v115 = 257;
v116 = 1835034;
v119 = 155633;
v122 = 256;
v123 = 1572886;
v126 = 129696;
v129 = 256;
v130 = 1048590;
v133 = 155672;
v134 = 4LL;
v135 = 1000000LL;
v136 = 257;
REGISTRY::RegReadRegTable(
v6,
(struct ADAPTER_CONTEXT *)DstBuf,
ConfigurationHandle,
(struct REGTABLE_ENTRY *)&v7,
0x11u);
REGISTRY::RegReadVendorDescription(this, (struct ADAPTER_CONTEXT *)DstBuf, ConfigurationHandle);
REGISTRY::RegReadVlanIds(this, (struct ADAPTER_CONTEXT *)DstBuf, ConfigurationHandle);
MsgReadAdapterInstanceName((struct ADAPTER_CONTEXT *)DstBuf, ConfigurationHandle);
RECEIVE_FILTER::ParseMMA((RECEIVE_FILTER *)(DstBuf + 409), (struct ADAPTER_CONTEXT *)DstBuf, ConfigurationHandle);
}

void __fastcall RECEIVE_FILTER::RxFilterReadRegistryParameters(
    RECEIVE_FILTER *this,
    struct ADAPTER_CONTEXT *a2,
    void *a3)
{
REGKEY<unsigned int>::Initialize(
(struct ADAPTER_CONTEXT *)((char *)a2 + 4148),
a2,
a3,
(PUCHAR)"*VMQ",
0,
1u,
0,
0,
1);
REGKEY<unsigned int>::Initialize(
(struct ADAPTER_CONTEXT *)((char *)a2 + 4160),
a2,
a3,
(PUCHAR)"*VMQLookaheadSplit",
0,
1u,
0,
0,
1);
REGKEY<unsigned int>::Initialize(
(struct ADAPTER_CONTEXT *)((char *)a2 + 4172),
a2,
a3,
(PUCHAR)"*VMQVlanFiltering",
0,
1u,
1u,
0,
1);
REGKEY<unsigned int>::Initialize(
(struct ADAPTER_CONTEXT *)((char *)a2 + 4184),
a2,
a3,
(PUCHAR)"*RssOrVmqPreference",
0,
1u,
0,
0,
1);
REGKEY<unsigned int>::Initialize(
(struct ADAPTER_CONTEXT *)((char *)a2 + 4136),
a2,
a3,
(PUCHAR)"VMQSupported",
0,
1u,
0,
0,
1);
if ( *((int *)a2 + 29935) < 18 )
*((_DWORD *)a2 + 1044) = 0;
*((_DWORD *)a2 + 1041) = 0;
}

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

void __fastcall RtAdapterCheckSetupAspmAndClkReq(struct _RT_ADAPTER *a1)
{
  unsigned int v2; // esi
  char v3; // bp
  unsigned int (__fastcall *v4)(__int64, _QWORD, __int64 *, __int64, int); // rax
  __int64 v5; // rcx
  char v6; // bl
  int v7; // ebx
  __int64 v8; // [rsp+50h] [rbp+8h] BYREF

  if ( **(_DWORD **)&RealtekTraceProvider > 5u )
  {
    v8 = (__int64)L"RtAdapterCheckSetupAspmAndClkReq";
    _tlgWriteTemplate<long (_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),&long _tlgWriteTransfer_EtwWriteTransfer(_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),_GUID const *,_GUID const *>::Write<_tlgWrapSz<unsigned short>>(
      RealtekTraceProvider,
      (int)&dword_14007994A,
      (__int64)&v8);
  }
  WppTraceEntry("RtAdapterCheckSetupAspmAndClkReq");
  v2 = 0;
  if ( *((_BYTE *)a1 + 10716) )
  {
    v3 = 1;
    *((_BYTE *)a1 + 10188) = 0;
    *((_BYTE *)a1 + 10228) = 0;
    if ( *((_BYTE *)a1 + 10187) )
    {
      v4 = (unsigned int (__fastcall *)(__int64, _QWORD, __int64 *, __int64, int))*((_QWORD *)a1 + 291);
      v5 = *((_QWORD *)a1 + 285);
      LOBYTE(v8) = 0;
      v6 = 0;
      if ( v4(v5, 0LL, &v8, 128LL, 1) == 1 )
        v6 = v8 & 3;
      *((_BYTE *)a1 + 10188) = v6;
    }
    else
    {
      DisableNicAspm(a1);
    }
    LOBYTE(v8) = -1;
    if ( *((_BYTE *)a1 + 10716)
      && (v7 = (*((__int64 (__fastcall **)(_QWORD, _QWORD, __int64 *, __int64, int))a1 + 291))(
                 *((_QWORD *)a1 + 285),
                 0LL,
                 &v8,
                 129LL,
                 1),
          KeStallExecutionProcessor(1u),
          v7 == 1)
      && (v8 & 1) != 0 )
    {
      *((_BYTE *)a1 + 10228) = 1;
    }
    else
    {
      v3 = *((_BYTE *)a1 + 10228);
    }
    RtWriteIntegerDataToRegistry(a1, L"ASPM", *((unsigned __int8 *)a1 + 10188));
    LOBYTE(v2) = v3 != 0;
    RtWriteIntegerDataToRegistry(a1, L"CLKREQ", v2);
  }
  if ( **(_DWORD **)&RealtekTraceProvider > 5u )
  {
    v8 = (__int64)L"RtAdapterCheckSetupAspmAndClkReq";
    _tlgWriteTemplate<long (_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),&long _tlgWriteTransfer_EtwWriteTransfer(_tlgProvider_t const *,void const *,_GUID const *,_GUID const *,unsigned int,_EVENT_DATA_DESCRIPTOR *),_GUID const *,_GUID const *>::Write<_tlgWrapSz<unsigned short>>(
      RealtekTraceProvider,
      (int)&dword_140079971,
      (__int64)&v8);
  }
  WppTraceExit("RtAdapterCheckSetupAspmAndClkReq");
}

NdisReadConfiguration(&Status, &ParameterValue, &ConfigurationHandle, &AoAcTestStr, NdisParameterHexInteger);
if (!Status && ParameterValue->ParameterData.IntegerData) {
    a1->FilterPnPFlags |= 0x200;
    ndisAoAcTest = 1;
}

void __fastcall Link::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<enum HdSplitLocation>::Initialize(
    (struct ADAPTER_CONTEXT *)((char *)*this + 992),
    *this,
    *((NDIS_HANDLE *)*this + 383),
    (PUCHAR)"*FlowControl",
    0,
    4u,
    4u,
    0,
    1);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (struct ADAPTER_CONTEXT *)((char *)*this + 980),
    *this,
    *((NDIS_HANDLE *)*this + 383),
    (PUCHAR)"*SpeedDuplex",
    0,
    0xC350u,
    0,
    0,
    1);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (struct ADAPTER_CONTEXT *)((char *)*this + 1004),
    *this,
    *((NDIS_HANDLE *)*this + 383),
    (PUCHAR)"FecMode",
    0,
    3u,
    0,
    0,
    1);
  (**(void (__fastcall ***)(struct ADAPTER_CONTEXT *))this[1])(this[1]);
}

void __fastcall ReceiveSideCoalescing::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)((char *)this + 36),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*RSCIPv4",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)((char *)this + 44),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*RSCIPv6",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 2),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"ForceRscEnabled",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 3),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RscMode",
    0,
    2u,
    1u,
    0,
    0);
}

void __fastcall HeaderSplitConfiguration::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 7),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*HeaderDataSplit",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 68),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitSize",
    0x80u,
    0x3C0u,
    0x80u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 10),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitAlways",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 11),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitLocation",
    0,
    3u,
    2u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 100),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"HDSplitBufferPad",
    0,
    2u,
    2u,
    0,
    0);
}

void __fastcall ReceiveConfiguration::ReadRegistryParameters(struct ADAPTER_CONTEXT **this)
{
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 11),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"*ReceiveBuffers",
    0x80u,
    0x1000u,
    0x200u,
    1,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 3),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"ReceiveBuffersOverride",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 8),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"MaxPacketCountPerDPC",
    8u,
    0xFFFFu,
    0x100u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 76),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"MaxPacketCountPerIndicate",
    1u,
    0xFFFFu,
    0x40u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 100),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RxDescriptorCountPerTailWrite",
    4u,
    0x1000u,
    8u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 52),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"MinHardwareOwnedPacketCount",
    8u,
    0x1000u,
    0x20u,
    1,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 5),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RxBufferPad",
    0,
    0x3Fu,
    0xAu,
    1,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 4),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"RegForceRxPathSerialization",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 2),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"EnableRxDescriptorChaining",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 14),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"EnableAdaptiveQueuing",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 124),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"AdaptiveQSize",
    0x40u,
    0x2000u,
    0x80u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(this + 17),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"AdaptiveQWorkSet",
    0x20u,
    0x2000u,
    0x60u,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)((char *)this + 148),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"AdaptiveQHysteresis",
    0x10u,
    0x400u,
    0x40u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 21),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"PadReceiveBuffer",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(this + 20),
    this[1],
    *((NDIS_HANDLE *)this[1] + 383),
    (PUCHAR)"StoreBadPackets",
    0,
    1u,
    0,
    0,
    0);
}

// Misc
__int64 __fastcall CheckRssSetting(struct _MP_PORT *a1)
{
  __int64 v1; // rbx
  char CurrentIrql; // bl
  _DWORD *DfltActivityPtr; // rax
  int v5; // esi
  char v6; // bl
  _DWORD *v7; // rax
  char v9; // bl
  _DWORD *v10; // rax
  char v11; // bl
  _DWORD *v12; // rax
  _QWORD StringsList[2]; // [rsp+40h] [rbp-28h] BYREF
  int v14; // [rsp+70h] [rbp+8h] BYREF

  v1 = *((_QWORD *)a1 + 147) + 21720LL;
  if ( *((_DWORD *)a1 + 1030) != 1 )
    return 0LL;
  if ( (*((unsigned __int8 (__fastcall **)(_QWORD))a1 + 235))(*((_QWORD *)a1 + 148)) )
  {
    StringsList[0] = v1;
    StringsList[1] = v1;
    NdisWriteEventLogEntry_0(g_pDriverObject, -2147024862, 0, 2u, StringsList, 0, 0LL);
    if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
      && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
      && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
    {
      CurrentIrql = KeGetCurrentIrql();
      DfltActivityPtr = (_DWORD *)TraceGetDfltActivityPtr();
      WPP_SF_sDD(
        WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
        80,
        (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
        (unsigned int)"IPOIB",
        CurrentIrql,
        *DfltActivityPtr);
    }
  }
  v14 = 0;
  v5 = ReadRegistryDword(
         L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Services\\Tcpip",
         L"\\Parameters",
         L"EnableRSS",
         1u,
         &v14);
  if ( v5 < 0 )
  {
    _mm_lfence();
    if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
      && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
      && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
    {
      v6 = KeGetCurrentIrql();
      v7 = (_DWORD *)TraceGetDfltActivityPtr();
      WPP_SF_sDDd(
        WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
        81,
        (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
        (unsigned int)"IPOIB",
        v6,
        *v7,
        v5);
    }
    return (unsigned int)v5;
  }
  if ( v14 == 1 )
    return 0LL;
  if ( !*((_BYTE *)a1 + 3529) )
  {
    if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
      && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
      && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
    {
      v11 = KeGetCurrentIrql();
      v12 = (_DWORD *)TraceGetDfltActivityPtr();
      WPP_SF_sDD(
        WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
        83,
        (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
        (unsigned int)"IPOIB",
        v11,
        *v12);
    }
    *((_DWORD *)a1 + 1030) = 0;
    StringsList[0] = (char *)a1 + 272;
    NdisWriteEventLogEntry_0(g_pDriverObject, -2147024881, 0, 1u, StringsList, 0, 0LL);
    return 0LL;
  }
  StringsList[0] = (char *)a1 + 272;
  NdisWriteEventLogEntry_0(g_pDriverObject, -1073283068, 0, 1u, StringsList, 0, 0LL);
  if ( WPP_GLOBAL_Control != (PDEVICE_OBJECT)&WPP_GLOBAL_Control
    && BYTE1(WPP_GLOBAL_Control[2].ActiveThreadCount) >= 2u
    && (*(&WPP_GLOBAL_Control[2].ActiveThreadCount + 1) & 1) != 0 )
  {
    v9 = KeGetCurrentIrql();
    v10 = (_DWORD *)TraceGetDfltActivityPtr();
    WPP_SF_sDD(
      WPP_GLOBAL_Control[2].Dpc.SystemArgument2,
      82,
      (unsigned int)&WPP_b8eb93e0936e34a0f000fad759de8ad2_Traceguids,
      (unsigned int)"IPOIB",
      v9,
      *v10);
  }
  return 3221225473LL;
}

void __fastcall Power::ReadRegistryParameters(Power *this)
{
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1196LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"EnablePowerManagement",
    0,
    1u,
    1u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1188LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"EnablePME",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<enum HdSplitLocation>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1208LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"WakeFromS5",
    0,
    0xFFFFu,
    2u,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1180LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"*DeviceSleepOnDisconnect",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1220LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"EnableModernStandby",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1232LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"*PMARPOffload",
    0,
    1u,
    0,
    0,
    0);
  RegistryKey<unsigned char>::Initialize(
    (enum RegKeyState *)(*((_QWORD *)this + 1) + 1240LL),
    *((struct ADAPTER_CONTEXT **)this + 1),
    *(NDIS_HANDLE *)(*((_QWORD *)this + 1) + 3064LL),
    (PUCHAR)"*PMNSOffload",
    0,
    1u,
    0,
    0,
    0);
}

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

void __fastcall SRIOV_CTRL::IovReadRegistryParameters(
        SRIOV_CTRL *this,
        REGISTRY **DstBuf,
        NDIS_HANDLE ConfigurationHandle)
{
  bool v6; // cc
  UNICODE_STRING SubKeyName; // [rsp+50h] [rbp+7h] BYREF
  int v8; // [rsp+60h] [rbp+17h] BYREF
  const wchar_t *v9; // [rsp+68h] [rbp+1Fh]
  __int64 v10; // [rsp+70h] [rbp+27h]
  int v11; // [rsp+78h] [rbp+2Fh]
  int v12; // [rsp+7Ch] [rbp+33h]
  int v13; // [rsp+80h] [rbp+37h]
  int v14; // [rsp+84h] [rbp+3Bh]
  int v15; // [rsp+88h] [rbp+3Fh]
  __int16 v16; // [rsp+8Ch] [rbp+43h]
  int Status; // [rsp+B0h] [rbp+67h] BYREF
  PVOID SubKeyHandle; // [rsp+B8h] [rbp+6Fh] BYREF

  SubKeyHandle = 0LL;
  *(_DWORD *)&SubKeyName.Length = 1835034;
  SubKeyName.Buffer = L"NicSwitches\\0";
  REGKEY<unsigned int>::Initialize(
    (enum _REGKEY_STATE *)(DstBuf + 14656),
    (struct ADAPTER_CONTEXT *)DstBuf,
    ConfigurationHandle,
    (PUCHAR)"*SriovPreferred",
    0,
    1u,
    0,
    0,
    1);
  REGKEY<unsigned int>::Initialize(
    (enum _REGKEY_STATE *)((char *)DstBuf + 117260),
    (struct ADAPTER_CONTEXT *)DstBuf,
    ConfigurationHandle,
    (PUCHAR)"*Sriov",
    0,
    1u,
    0,
    0,
    1);
  v10 = 0LL;
  v13 = 0;
  v6 = *((_DWORD *)DstBuf + 29935) < 18;
  v9 = L"*NumVFs";
  v15 = *((_DWORD *)this + 141);
  v8 = 1048590;
  v11 = 117284;
  v12 = 4;
  v14 = 7;
  v16 = 256;
  if ( v6 )
    *((_DWORD *)DstBuf + 29316) = 0;
  NdisOpenConfigurationKeyByName_0(&Status, ConfigurationHandle, &SubKeyName, &SubKeyHandle);
  if ( SubKeyHandle )
  {
    REGISTRY::RegReadRegTable(
      DstBuf[14912],
      (struct ADAPTER_CONTEXT *)DstBuf,
      SubKeyHandle,
      (struct REGTABLE_ENTRY *)&v8,
      1u);
    NdisCloseConfiguration_0(SubKeyHandle);
  }
  SRIOV_CTRL::CalculateNumVfs(this, (struct ADAPTER_CONTEXT *)DstBuf);
  *(_WORD *)(*((_QWORD *)this + 73) + 16LL) = *((_WORD *)this + 282);
}

void __fastcall DcaReadRegistryParameters(struct ADAPTER_CONTEXT *a1, void *a2)
{
  REGISTRY *v4; // rcx
  int v5; // [rsp+30h] [rbp-49h] BYREF
  const wchar_t *v6; // [rsp+38h] [rbp-41h]
  __int64 v7; // [rsp+40h] [rbp-39h]
  int v8; // [rsp+48h] [rbp-31h]
  int v9; // [rsp+4Ch] [rbp-2Dh]
  int v10; // [rsp+50h] [rbp-29h]
  int v11; // [rsp+54h] [rbp-25h]
  int v12; // [rsp+58h] [rbp-21h]
  __int16 v13; // [rsp+5Ch] [rbp-1Dh]
  int v14; // [rsp+60h] [rbp-19h]
  const wchar_t *v15; // [rsp+68h] [rbp-11h]
  __int64 v16; // [rsp+70h] [rbp-9h]
  int v17; // [rsp+78h] [rbp-1h]
  int v18; // [rsp+7Ch] [rbp+3h]
  int v19; // [rsp+80h] [rbp+7h]
  int v20; // [rsp+84h] [rbp+Bh]
  int v21; // [rsp+88h] [rbp+Fh]
  __int16 v22; // [rsp+8Ch] [rbp+13h]
  int v23; // [rsp+90h] [rbp+17h]
  const wchar_t *v24; // [rsp+98h] [rbp+1Fh]
  __int64 v25; // [rsp+A0h] [rbp+27h]
  int v26; // [rsp+A8h] [rbp+2Fh]
  int v27; // [rsp+ACh] [rbp+33h]
  int v28; // [rsp+B0h] [rbp+37h]
  int v29; // [rsp+B4h] [rbp+3Bh]
  int v30; // [rsp+B8h] [rbp+3Fh]
  __int16 v31; // [rsp+BCh] [rbp+43h]

  v7 = 0LL;
  v10 = 0;
  v16 = 0LL;
  v19 = 0;
  v25 = 0LL;
  v28 = 0;
  v6 = L"EnableDCA";
  v18 = 4;
  v27 = 4;
  v15 = L"DcaRxSettings";
  v4 = (REGISTRY *)*((_QWORD *)a1 + 14912);
  v21 = 3;
  v5 = 1310738;
  v8 = 2840;
  v9 = 1;
  v11 = 1;
  v12 = 1;
  v13 = 256;
  v14 = 1835034;
  v17 = 2912;
  v20 = 7;
  v22 = 256;
  v23 = 1835034;
  v24 = L"DcaTxSettings";
  v26 = 2916;
  v29 = 1;
  v30 = 1;
  v31 = 256;
  REGISTRY::RegReadRegTable(v4, a1, a2, (struct REGTABLE_ENTRY *)&v5, 3u);
}

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