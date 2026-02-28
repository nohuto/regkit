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
