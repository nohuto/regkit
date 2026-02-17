char __fastcall GetRegistrySettings(__int64 a1)
{
  __int64 v1; // rax
  unsigned int v3; // eax
  __int64 v4; // r10
  char *v5; // rax
  unsigned int v6; // ecx
  int v7; // edi
  int v8; // r8d
  int v9; // r8d
  int v10; // r8d
  int v11; // r8d
  int v12; // r8d
  int v13; // r8d
  int v14; // r8d
  int v15; // r8d
  int v16; // r8d
  int v17; // r8d
  int v18; // r8d
  int v19; // r8d
  int v20; // r8d
  __int64 v21; // rsi
  int v22; // r8d
  int v23; // r8d
  int v24; // r8d
  __int64 v25; // rcx
  int v26; // r8d
  int v27; // r8d
  int v28; // r8d
  int v29; // r8d
  int v30; // r8d
  int v31; // r8d
  int v32; // r8d
  int v33; // r8d
  int v34; // r8d
  int v35; // r8d
  int v36; // r8d
  unsigned int v37; // r12d
  int v38; // r8d
  int v39; // r8d
  int v40; // r8d
  int v41; // r8d
  int v42; // r8d
  int v43; // esi
  int v44; // r8d
  int v45; // r8d
  int v46; // r8d
  int v47; // r8d
  int v48; // r8d
  int v49; // r8d
  int v50; // r8d
  int v51; // r8d
  int v52; // edx
  int v53; // r8d
  unsigned int v54; // r14d
  int v55; // r8d
  int v56; // r8d
  int v57; // r8d
  int v58; // r8d
  int v59; // r8d
  int v60; // r8d
  int v61; // r8d
  int v62; // r8d
  int v63; // r8d
  int v64; // r8d
  int v65; // r8d
  int v66; // r8d
  _BYTE v68[4]; // [rsp+50h] [rbp-9h] BYREF
  unsigned int v69; // [rsp+54h] [rbp-5h] BYREF
  int v70; // [rsp+58h] [rbp-1h] BYREF
  _DWORD Size[3]; // [rsp+5Ch] [rbp+3h] BYREF
  char v72[32]; // [rsp+68h] [rbp+Fh] BYREF

  LOBYTE(v1) = aVenVvvvDevDddd[24];
  strcpy(v72, "VEN_vvvv&DEV_dddd&REV_rr");
  v70 = 512;
  if ( !*(_BYTE *)(a1 + 20) )
  {
    *(_QWORD *)(a1 + 40) = 0LL;
    *(_DWORD *)(a1 + 92) = -1;
    *(_DWORD *)(a1 + 108) = -1;
    *(_DWORD *)(a1 + 112) = -1;
    *(_DWORD *)(a1 + 124) = -1;
    *(_DWORD *)(a1 + 140) = -694967296;
    v3 = *(_DWORD *)(a1 + 64) & 0xFFF3E428;
    *(_DWORD *)(a1 + 48) = 0;
    *(_DWORD *)(a1 + 52) = 16711680;
    *(_DWORD *)(a1 + 64) = v3 | 8;
    *(_QWORD *)(a1 + 56) = 0LL;
    *(_QWORD *)(a1 + 68) = 1823LL;
    *(_QWORD *)(a1 + 84) = 0x4000LL;
    *(_QWORD *)(a1 + 116) = 0LL;
    *(_DWORD *)(a1 + 132) = 0;
    *(_DWORD *)(a1 + 136) = 1;
    *(_DWORD *)(a1 + 96) = 50;
    *(_DWORD *)(a1 + 100) = 3000;
    *(_DWORD *)(a1 + 104) = 1000;
    *(_DWORD *)(a1 + 76) = 0;
    *(_DWORD *)(a1 + 80) = 100;
    *(_DWORD *)(a1 + 144) = 600000000;
    *(_WORD *)(a1 + 148) = 0;
    *(_BYTE *)(a1 + 150) = 0;
    *(_DWORD *)(a1 + 160) = 0;
    *(_DWORD *)(a1 + 168) = 20000;
    v1 = StorPortAllocateRegistryBuffer(a1, &v70);
    *(_QWORD *)&Size[1] = v1;
    if ( v1 )
    {
      UlongToHex(&v72[4], *(unsigned __int16 *)(a1 + 4), 4LL);
      UlongToHex(&v72[13], *(unsigned __int16 *)(a1 + 6), 4LL);
      UlongToHex(&v72[22], *(unsigned __int8 *)(a1 + 8), 2LL);
      v5 = v72;
      v6 = 0;
      v7 = 29;
      while ( *v5 )
      {
        ++v6;
        ++v5;
        if ( v6 >= 0x1D )
          goto LABEL_8;
      }
      v7 = v6;
LABEL_8:
      Size[0] = v70;
      NVMeZeroMemory(v4);
      v69 = 0;
      LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                     a1,
                     (unsigned int)"MaxTransferSize",
                     v8,
                     (unsigned int)&Size[1],
                     (__int64)&v70,
                     (__int64)Size,
                     (__int64)v72,
                     v7,
                     (__int64)&v69);
      v68[0] = v1;
      if ( (_BYTE)v1 == 1 )
      {
        LODWORD(v1) = v69;
        if ( v69 )
        {
          if ( v69 > 0x800 )
            LODWORD(v1) = 2048;
          LODWORD(v1) = (_DWORD)v1 << 10;
          *(_DWORD *)(a1 + 40) = v1;
        }
      }
      if ( *(_QWORD *)&Size[1] )
      {
        Size[0] = v70;
        NVMeZeroMemory(*(_QWORD *)&Size[1]);
        v69 = 0;
        LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                       a1,
                       (unsigned int)"IoQueueDepth",
                       v9,
                       (unsigned int)&Size[1],
                       (__int64)&v70,
                       (__int64)Size,
                       (__int64)v72,
                       v7,
                       (__int64)&v69);
        v68[0] = v1;
        if ( (_BYTE)v1 == 1 )
        {
          LOBYTE(v1) = v69;
          if ( v69 )
            *(_DWORD *)(a1 + 44) = v69;
        }
        if ( *(_QWORD *)&Size[1] )
        {
          Size[0] = v70;
          NVMeZeroMemory(*(_QWORD *)&Size[1]);
          v69 = 0;
          LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                         a1,
                         (unsigned int)"IoSubmissionQueueCount",
                         v10,
                         (unsigned int)&Size[1],
                         (__int64)&v70,
                         (__int64)Size,
                         (__int64)v72,
                         v7,
                         (__int64)&v69);
          v68[0] = v1;
          if ( (_BYTE)v1 == 1 )
          {
            LOBYTE(v1) = v69;
            if ( v69 )
              *(_WORD *)(a1 + 48) = v69;
          }
          if ( *(_QWORD *)&Size[1] )
          {
            Size[0] = v70;
            NVMeZeroMemory(*(_QWORD *)&Size[1]);
            v69 = 0;
            LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                           a1,
                           (unsigned int)"IoCompletionQueueCount",
                           v11,
                           (unsigned int)&Size[1],
                           (__int64)&v70,
                           (__int64)Size,
                           (__int64)v72,
                           v7,
                           (__int64)&v69);
            v68[0] = v1;
            if ( (_BYTE)v1 == 1 )
            {
              LOBYTE(v1) = v69;
              if ( v69 )
                *(_WORD *)(a1 + 50) = v69;
            }
            if ( *(_QWORD *)&Size[1] )
            {
              Size[0] = v70;
              NVMeZeroMemory(*(_QWORD *)&Size[1]);
              v69 = 0;
              LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                             a1,
                             (unsigned int)"InterruptCoalescingTime",
                             v12,
                             (unsigned int)&Size[1],
                             (__int64)&v70,
                             (__int64)Size,
                             (__int64)v72,
                             v7,
                             (__int64)&v69);
              v68[0] = v1;
              if ( (_BYTE)v1 == 1 )
              {
                LOBYTE(v1) = v69;
                if ( v69 )
                  *(_BYTE *)(a1 + 52) = v69;
              }
              if ( *(_QWORD *)&Size[1] )
              {
                Size[0] = v70;
                NVMeZeroMemory(*(_QWORD *)&Size[1]);
                v69 = 0;
                LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                               a1,
                               (unsigned int)"InterruptCoalescingEntry",
                               v13,
                               (unsigned int)&Size[1],
                               (__int64)&v70,
                               (__int64)Size,
                               (__int64)v72,
                               v7,
                               (__int64)&v69);
                v68[0] = v1;
                if ( (_BYTE)v1 == 1 )
                {
                  LOBYTE(v1) = v69;
                  if ( v69 )
                    *(_BYTE *)(a1 + 53) = v69;
                }
                if ( *(_QWORD *)&Size[1] )
                {
                  Size[0] = v70;
                  NVMeZeroMemory(*(_QWORD *)&Size[1]);
                  v69 = 0;
                  LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                 a1,
                                 (unsigned int)"ArbitrationBurst",
                                 v14,
                                 (unsigned int)&Size[1],
                                 (__int64)&v70,
                                 (__int64)Size,
                                 (__int64)v72,
                                 v7,
                                 (__int64)&v69);
                  v68[0] = v1;
                  if ( (_BYTE)v1 == 1 )
                  {
                    LOBYTE(v1) = v69;
                    if ( v69 )
                      *(_BYTE *)(a1 + 54) = v69;
                  }
                  if ( *(_QWORD *)&Size[1] )
                  {
                    Size[0] = v70;
                    NVMeZeroMemory(*(_QWORD *)&Size[1]);
                    v69 = 0;
                    LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                   a1,
                                   (unsigned int)"ContiguousMemoryFromAnyNode",
                                   v15,
                                   (unsigned int)&Size[1],
                                   (__int64)&v70,
                                   (__int64)Size,
                                   (__int64)v72,
                                   v7,
                                   (__int64)&v69);
                    v68[0] = v1;
                    if ( (_BYTE)v1 == 1 && v69 )
                      *(_DWORD *)(a1 + 64) |= 2u;
                    if ( *(_QWORD *)&Size[1] )
                    {
                      Size[0] = v70;
                      NVMeZeroMemory(*(_QWORD *)&Size[1]);
                      v69 = 0;
                      LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                     a1,
                                     (unsigned int)"ShutdownTimeout",
                                     v16,
                                     (unsigned int)&Size[1],
                                     (__int64)&v70,
                                     (__int64)Size,
                                     (__int64)v72,
                                     v7,
                                     (__int64)&v69);
                      v68[0] = v1;
                      if ( (_BYTE)v1 == 1 )
                      {
                        LOBYTE(v1) = v69;
                        if ( v69 )
                        {
                          if ( v69 > 0xFF )
                            LOBYTE(v1) = -1;
                          *(_BYTE *)(a1 + 55) = v1;
                        }
                      }
                      if ( *(_QWORD *)&Size[1] )
                      {
                        Size[0] = v70;
                        NVMeZeroMemory(*(_QWORD *)&Size[1]);
                        v69 = 0;
                        LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                       a1,
                                       (unsigned int)"DeallocateMaxLbaCount",
                                       v17,
                                       (unsigned int)&Size[1],
                                       (__int64)&v70,
                                       (__int64)Size,
                                       (__int64)v72,
                                       v7,
                                       (__int64)&v69);
                        v68[0] = v1;
                        if ( (_BYTE)v1 == 1 )
                        {
                          LOBYTE(v1) = v69;
                          if ( v69 )
                            *(_DWORD *)(a1 + 56) = v69;
                        }
                        if ( *(_QWORD *)&Size[1] )
                        {
                          Size[0] = v70;
                          NVMeZeroMemory(*(_QWORD *)&Size[1]);
                          v69 = 0;
                          LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                         a1,
                                         (unsigned int)"DisableDeallocate",
                                         v18,
                                         (unsigned int)&Size[1],
                                         (__int64)&v70,
                                         (__int64)Size,
                                         (__int64)v72,
                                         v7,
                                         (__int64)&v69);
                          v68[0] = v1;
                          if ( (_BYTE)v1 == 1 && v69 )
                            *(_DWORD *)(a1 + 64) |= 1u;
                          if ( *(_QWORD *)&Size[1] )
                          {
                            Size[0] = v70;
                            NVMeZeroMemory(*(_QWORD *)&Size[1]);
                            v69 = 0;
                            LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                           a1,
                                           (unsigned int)"ControllerBasicInit",
                                           v19,
                                           (unsigned int)&Size[1],
                                           (__int64)&v70,
                                           (__int64)Size,
                                           (__int64)v72,
                                           v7,
                                           (__int64)&v69);
                            v68[0] = v1;
                            if ( (_BYTE)v1 == 1 && v69 )
                              *(_DWORD *)(a1 + 64) |= 4u;
                            if ( *(_QWORD *)&Size[1] )
                            {
                              Size[0] = v70;
                              NVMeZeroMemory(*(_QWORD *)&Size[1]);
                              v69 = 0;
                              LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                             a1,
                                             (unsigned int)"AsyncEventMask",
                                             v20,
                                             (unsigned int)&Size[1],
                                             (__int64)&v70,
                                             (__int64)Size,
                                             (__int64)v72,
                                             v7,
                                             (__int64)&v69);
                              if ( (_BYTE)v1 == 1 )
                              {
                                LOBYTE(v1) = v69;
                                if ( v69 )
                                {
                                  LODWORD(v1) = v69 & 0x1F;
                                  *(_DWORD *)(a1 + 68) = v1;
                                }
                              }
                              v21 = *(_QWORD *)&Size[1];
                              if ( *(_QWORD *)&Size[1] )
                              {
                                v68[0] = 0;
                                if ( !(unsigned int)StorPortExtendedFunction(97LL, a1, 0LL, v68) && v68[0] )
                                {
                                  *(_DWORD *)(a1 + 88) = 2;
LABEL_71:
                                  Size[0] = v70;
                                  NVMeZeroMemory(v21);
                                  v69 = 0;
                                  v68[0] = ReadMultiSzRegistryValueAndCompareId(
                                             a1,
                                             (unsigned int)"DiagnosticFlags",
                                             v23,
                                             (unsigned int)&Size[1],
                                             (__int64)&v70,
                                             (__int64)Size,
                                             (__int64)v72,
                                             v7,
                                             (__int64)&v69);
                                  if ( v68[0] == 1 && v69 )
                                    *(_DWORD *)(a1 + 116) = v69;
                                  LODWORD(v1) = *(_DWORD *)(a1 + 116);
                                  if ( (v1 & 2) != 0 )
                                    *(_DWORD *)(a1 + 120) = 0x100000;
                                  if ( *(_QWORD *)&Size[1] )
                                  {
                                    Size[0] = v70;
                                    NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                    v69 = 0;
                                    v68[0] = ReadMultiSzRegistryValueAndCompareId(
                                               a1,
                                               (unsigned int)"LogSize",
                                               v24,
                                               (unsigned int)&Size[1],
                                               (__int64)&v70,
                                               (__int64)Size,
                                               (__int64)v72,
                                               v7,
                                               (__int64)&v69);
                                    if ( v68[0] == 1 && v69 )
                                      *(_DWORD *)(a1 + 120) = v69 << 10;
                                    LODWORD(v1) = -1431655765 * *(_DWORD *)(a1 + 120);
                                    v25 = *(_QWORD *)&Size[1];
                                    *(_DWORD *)(a1 + 4076) = *(_DWORD *)(a1 + 120) / 0x18u;
                                    if ( v25 )
                                    {
                                      Size[0] = v70;
                                      NVMeZeroMemory(v25);
                                      v69 = 0;
                                      LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                     a1,
                                                     (unsigned int)"IoStripeAlignment",
                                                     v26,
                                                     (unsigned int)&Size[1],
                                                     (__int64)&v70,
                                                     (__int64)Size,
                                                     (__int64)v72,
                                                     v7,
                                                     (__int64)&v69);
                                      v68[0] = v1;
                                      if ( (_BYTE)v1 == 1 )
                                      {
                                        LOBYTE(v1) = v69;
                                        if ( v69 )
                                        {
                                          LODWORD(v1) = v69 << 10;
                                          if ( ((v69 << 10) & 0xFFF) == 0 )
                                            *(_DWORD *)(a1 + 60) = v1;
                                        }
                                      }
                                      if ( *(_QWORD *)&Size[1] )
                                      {
                                        Size[0] = v70;
                                        NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                        v69 = 0;
                                        LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                       a1,
                                                       (unsigned int)"MedPowerFxIdleTimeout",
                                                       v27,
                                                       (unsigned int)&Size[1],
                                                       (__int64)&v70,
                                                       (__int64)Size,
                                                       (__int64)v72,
                                                       v7,
                                                       (__int64)&v69);
                                        v68[0] = v1;
                                        if ( (_BYTE)v1 == 1 )
                                        {
                                          LOBYTE(v1) = v69;
                                          *(_DWORD *)(a1 + 92) = v69;
                                        }
                                        if ( *(_QWORD *)&Size[1] )
                                        {
                                          Size[0] = v70;
                                          NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                          v69 = 0;
                                          LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                         a1,
                                                         (unsigned int)"LowestPowerFxIdleTimeout",
                                                         v28,
                                                         (unsigned int)&Size[1],
                                                         (__int64)&v70,
                                                         (__int64)Size,
                                                         (__int64)v72,
                                                         v7,
                                                         (__int64)&v69);
                                          v68[0] = v1;
                                          if ( (_BYTE)v1 == 1 )
                                          {
                                            LOBYTE(v1) = v69;
                                            *(_DWORD *)(a1 + 96) = v69;
                                          }
                                          if ( *(_QWORD *)&Size[1] )
                                          {
                                            Size[0] = v70;
                                            NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                            v69 = 0;
                                            LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                           a1,
                                                           (unsigned int)"MedPowerD3IdleTimeout",
                                                           v29,
                                                           (unsigned int)&Size[1],
                                                           (__int64)&v70,
                                                           (__int64)Size,
                                                           (__int64)v72,
                                                           v7,
                                                           (__int64)&v69);
                                            v68[0] = v1;
                                            if ( (_BYTE)v1 == 1 )
                                            {
                                              LOBYTE(v1) = v69;
                                              *(_DWORD *)(a1 + 100) = v69;
                                            }
                                            if ( *(_QWORD *)&Size[1] )
                                            {
                                              Size[0] = v70;
                                              NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                              v69 = 0;
                                              LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                             a1,
                                                             (unsigned int)"LowestPowerD3IdleTimeout",
                                                             v30,
                                                             (unsigned int)&Size[1],
                                                             (__int64)&v70,
                                                             (__int64)Size,
                                                             (__int64)v72,
                                                             v7,
                                                             (__int64)&v69);
                                              v68[0] = v1;
                                              if ( (_BYTE)v1 == 1 )
                                              {
                                                LOBYTE(v1) = v69;
                                                *(_DWORD *)(a1 + 104) = v69;
                                              }
                                              if ( *(_QWORD *)&Size[1] )
                                              {
                                                Size[0] = v70;
                                                NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                v69 = 0;
                                                LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                               a1,
                                                               (unsigned int)"MedPowerResumeLatency",
                                                               v31,
                                                               (unsigned int)&Size[1],
                                                               (__int64)&v70,
                                                               (__int64)Size,
                                                               (__int64)v72,
                                                               v7,
                                                               (__int64)&v69);
                                                v68[0] = v1;
                                                if ( (_BYTE)v1 == 1 )
                                                {
                                                  LOBYTE(v1) = v69;
                                                  *(_DWORD *)(a1 + 108) = v69;
                                                }
                                                if ( *(_QWORD *)&Size[1] )
                                                {
                                                  Size[0] = v70;
                                                  NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                  v69 = 0;
                                                  LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                 a1,
                                                                 (unsigned int)"LowestPowerResumeLatency",
                                                                 v32,
                                                                 (unsigned int)&Size[1],
                                                                 (__int64)&v70,
                                                                 (__int64)Size,
                                                                 (__int64)v72,
                                                                 v7,
                                                                 (__int64)&v69);
                                                  v68[0] = v1;
                                                  if ( (_BYTE)v1 == 1 )
                                                  {
                                                    LOBYTE(v1) = v69;
                                                    *(_DWORD *)(a1 + 112) = v69;
                                                  }
                                                  if ( *(_QWORD *)&Size[1] )
                                                  {
                                                    Size[0] = v70;
                                                    NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                    v69 = 0;
                                                    LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                   a1,
                                                                   (unsigned int)"HostMemoryBufferBytes",
                                                                   v33,
                                                                   (unsigned int)&Size[1],
                                                                   (__int64)&v70,
                                                                   (__int64)Size,
                                                                   (__int64)v72,
                                                                   v7,
                                                                   (__int64)&v69);
                                                    v68[0] = v1;
                                                    if ( (_BYTE)v1 == 1 )
                                                    {
                                                      LOBYTE(v1) = v69;
                                                      *(_DWORD *)(a1 + 124) = v69;
                                                    }
                                                    if ( *(_QWORD *)&Size[1] )
                                                    {
                                                      Size[0] = v70;
                                                      NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                      v69 = 0;
                                                      LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                     a1,
                                                                     (unsigned int)"BypassSgl",
                                                                     v34,
                                                                     (unsigned int)&Size[1],
                                                                     (__int64)&v70,
                                                                     (__int64)Size,
                                                                     (__int64)v72,
                                                                     v7,
                                                                     (__int64)&v69);
                                                      v68[0] = v1;
                                                      if ( (_BYTE)v1 == 1 )
                                                      {
                                                        LODWORD(v1) = (*(_DWORD *)(a1 + 64) ^ (8 * v69)) & 8;
                                                        *(_DWORD *)(a1 + 64) ^= v1;
                                                      }
                                                      if ( *(_QWORD *)&Size[1] )
                                                      {
                                                        Size[0] = v70;
                                                        NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                        v69 = 0;
                                                        LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                       a1,
                                                                       (unsigned int)"TestMdlDataBufferOffsetInBytes",
                                                                       v35,
                                                                       (unsigned int)&Size[1],
                                                                       (__int64)&v70,
                                                                       (__int64)Size,
                                                                       (__int64)v72,
                                                                       v7,
                                                                       (__int64)&v69);
                                                        v68[0] = v1;
                                                        if ( (_BYTE)v1 == 1 )
                                                        {
                                                          LOBYTE(v1) = v69;
                                                          *(_DWORD *)(a1 + 128) = v69;
                                                        }
                                                        if ( *(_QWORD *)&Size[1] )
                                                        {
                                                          Size[0] = v70;
                                                          NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                          LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                         a1,
                                                                         (unsigned int)"UseDumpPointers",
                                                                         v36,
                                                                         (unsigned int)&Size[1],
                                                                         (__int64)&v70,
                                                                         (__int64)Size,
                                                                         (__int64)v72,
                                                                         v7,
                                                                         0LL);
                                                          v68[0] = v1;
                                                          v37 = 16;
                                                          if ( (_BYTE)v1 == 1 )
                                                            *(_DWORD *)(a1 + 64) |= 0x10u;
                                                          if ( *(_QWORD *)&Size[1] )
                                                          {
                                                            Size[0] = v70;
                                                            NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                            v69 = 0;
                                                            LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                           a1,
                                                                           (unsigned int)"ReservedQueuePairCount",
                                                                           v38,
                                                                           (unsigned int)&Size[1],
                                                                           (__int64)&v70,
                                                                           (__int64)Size,
                                                                           (__int64)v72,
                                                                           v7,
                                                                           (__int64)&v69);
                                                            v68[0] = v1;
                                                            if ( (_BYTE)v1 == 1 )
                                                            {
                                                              LOBYTE(v1) = v69 - 1;
                                                              if ( v69 - 1 <= 0xFFFE )
                                                                *(_DWORD *)(a1 + 132) = v69;
                                                            }
                                                            if ( *(_QWORD *)&Size[1] )
                                                            {
                                                              Size[0] = v70;
                                                              NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                              v69 = 0;
                                                              LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                             a1,
                                                                             (unsigned int)"NvmeTestSwitch",
                                                                             v39,
                                                                             (unsigned int)&Size[1],
                                                                             (__int64)&v70,
                                                                             (__int64)Size,
                                                                             (__int64)v72,
                                                                             v7,
                                                                             (__int64)&v69);
                                                              v68[0] = v1;
                                                              if ( (_BYTE)v1 == 1 )
                                                              {
                                                                LOBYTE(v1) = v69;
                                                                if ( v69 )
                                                                  *(_DWORD *)(a1 + 136) = v69;
                                                              }
                                                              if ( *(_QWORD *)&Size[1] )
                                                              {
                                                                Size[0] = v70;
                                                                NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                v69 = 0;
                                                                LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                               a1,
                                                                               (unsigned int)"IoQueuePercentageInPollingMode",
                                                                               v40,
                                                                               (unsigned int)&Size[1],
                                                                               (__int64)&v70,
                                                                               (__int64)Size,
                                                                               (__int64)v72,
                                                                               v7,
                                                                               (__int64)&v69);
                                                                v68[0] = v1;
                                                                if ( (_BYTE)v1 == 1 )
                                                                {
                                                                  if ( v69 <= 0x64 )
                                                                  {
                                                                    LOBYTE(v1) = v69;
                                                                    *(_DWORD *)(a1 + 72) = (unsigned __int16)v69;
                                                                  }
                                                                  else
                                                                  {
                                                                    *(_DWORD *)(a1 + 72) = 100;
                                                                  }
                                                                }
                                                                if ( *(_QWORD *)&Size[1] )
                                                                {
                                                                  Size[0] = v70;
                                                                  NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                  v69 = 0;
                                                                  LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                 a1,
                                                                                 (unsigned int)"IoPollingInterval",
                                                                                 v41,
                                                                                 (unsigned int)&Size[1],
                                                                                 (__int64)&v70,
                                                                                 (__int64)Size,
                                                                                 (__int64)v72,
                                                                                 v7,
                                                                                 (__int64)&v69);
                                                                  v68[0] = v1;
                                                                  if ( (_BYTE)v1 == 1 )
                                                                  {
                                                                    LODWORD(v1) = 100000;
                                                                    if ( v69 < 0x186A0 )
                                                                      LODWORD(v1) = v69;
                                                                    *(_DWORD *)(a1 + 76) = v1;
                                                                  }
                                                                  if ( *(_QWORD *)&Size[1] )
                                                                  {
                                                                    Size[0] = v70;
                                                                    NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                    v69 = 0;
                                                                    LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                   a1,
                                                                                   (unsigned int)"IoCompletionCapInDPC",
                                                                                   v42,
                                                                                   (unsigned int)&Size[1],
                                                                                   (__int64)&v70,
                                                                                   (__int64)Size,
                                                                                   (__int64)v72,
                                                                                   v7,
                                                                                   (__int64)&v69);
                                                                    v68[0] = v1;
                                                                    v43 = 128;
                                                                    if ( (_BYTE)v1 == 1 )
                                                                    {
                                                                      LODWORD(v1) = v69;
                                                                      if ( v69 )
                                                                      {
                                                                        if ( v69 > 0x80 )
                                                                          LODWORD(v1) = 128;
                                                                        *(_DWORD *)(a1 + 80) = v1;
                                                                      }
                                                                    }
                                                                    if ( *(_QWORD *)&Size[1] )
                                                                    {
                                                                      Size[0] = v70;
                                                                      NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                      v69 = 0;
                                                                      LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                     a1,
                                                                                     (unsigned int)"IoPollingSize",
                                                                                     v44,
                                                                                     (unsigned int)&Size[1],
                                                                                     (__int64)&v70,
                                                                                     (__int64)Size,
                                                                                     (__int64)v72,
                                                                                     v7,
                                                                                     (__int64)&v69);
                                                                      v68[0] = v1;
                                                                      if ( (_BYTE)v1 == 1 )
                                                                      {
                                                                        LOBYTE(v1) = v69;
                                                                        if ( v69 )
                                                                          *(_DWORD *)(a1 + 84) = v69;
                                                                      }
                                                                      if ( *(_QWORD *)&Size[1] )
                                                                      {
                                                                        Size[0] = v70;
                                                                        NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                        v69 = 0;
                                                                        LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                       a1,
                                                                                       (unsigned int)"ErrorEtwThrottleInterval",
                                                                                       v45,
                                                                                       (unsigned int)&Size[1],
                                                                                       (__int64)&v70,
                                                                                       (__int64)Size,
                                                                                       (__int64)v72,
                                                                                       v7,
                                                                                       (__int64)&v69);
                                                                        v68[0] = v1;
                                                                        if ( (_BYTE)v1 == 1 )
                                                                        {
                                                                          LODWORD(v1) = v69;
                                                                          if ( v69 )
                                                                          {
                                                                            if ( v69 > 0xD693A400 )
                                                                              LODWORD(v1) = -694967296;
                                                                            *(_DWORD *)(a1 + 140) = v1;
                                                                          }
                                                                        }
                                                                        if ( *(_QWORD *)&Size[1] )
                                                                        {
                                                                          Size[0] = v70;
                                                                          NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                          v69 = 0;
                                                                          LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                         a1,
                                                                                         (unsigned int)"ResetEnableMask",
                                                                                         v46,
                                                                                         (unsigned int)&Size[1],
                                                                                         (__int64)&v70,
                                                                                         (__int64)Size,
                                                                                         (__int64)v72,
                                                                                         v7,
                                                                                         (__int64)&v69);
                                                                          v68[0] = v1;
                                                                          if ( (_BYTE)v1 == 1 )
                                                                          {
                                                                            LOBYTE(v1) = v69;
                                                                            if ( v69 )
                                                                            {
                                                                              if ( (v69 & 1) != 0 )
                                                                                *(_DWORD *)(a1 + 64) |= 0x40u;
                                                                              if ( (v1 & 2) != 0 )
                                                                                *(_DWORD *)(a1 + 64) |= 0x800u;
                                                                              if ( (v1 & 4) != 0 )
                                                                                *(_DWORD *)(a1 + 64) |= 0x1000u;
                                                                            }
                                                                          }
                                                                          if ( *(_QWORD *)&Size[1] )
                                                                          {
                                                                            Size[0] = v70;
                                                                            NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                            v69 = 0;
                                                                            LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                           a1,
                                                                                           (unsigned int)"ReliabilityDegraded",
                                                                                           v47,
                                                                                           (unsigned int)&Size[1],
                                                                                           (__int64)&v70,
                                                                                           (__int64)Size,
                                                                                           (__int64)v72,
                                                                                           v7,
                                                                                           (__int64)&v69);
                                                                            v68[0] = v1;
                                                                            if ( (_BYTE)v1 == 1 && v69 )
                                                                              *(_DWORD *)(a1 + 64) |= 0x80u;
                                                                            if ( *(_QWORD *)&Size[1] )
                                                                            {
                                                                              Size[0] = v70;
                                                                              NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                              v69 = 0;
                                                                              LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                             a1,
                                                                                             (unsigned int)"ReadOnly",
                                                                                             v48,
                                                                                             (unsigned int)&Size[1],
                                                                                             (__int64)&v70,
                                                                                             (__int64)Size,
                                                                                             (__int64)v72,
                                                                                             v7,
                                                                                             (__int64)&v69);
                                                                              v68[0] = v1;
                                                                              if ( (_BYTE)v1 == 1 && v69 )
                                                                                *(_DWORD *)(a1 + 64) |= 0x100u;
                                                                              if ( *(_QWORD *)&Size[1] )
                                                                              {
                                                                                Size[0] = v70;
                                                                                NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                v69 = 0;
                                                                                v68[0] = ReadMultiSzRegistryValueAndCompareId(
                                                                                           a1,
                                                                                           (unsigned int)"VolatileMemoryBackupDeviceFailed",
                                                                                           v49,
                                                                                           (unsigned int)&Size[1],
                                                                                           (__int64)&v70,
                                                                                           (__int64)Size,
                                                                                           (__int64)v72,
                                                                                           v7,
                                                                                           (__int64)&v69);
                                                                                if ( v68[0] == 1 && v69 )
                                                                                  *(_DWORD *)(a1 + 64) |= 0x200u;
                                                                                Size[0] = v70;
                                                                                NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                v69 = 0;
                                                                                v68[0] = ReadMultiSzRegistryValueAndCompareId(
                                                                                           a1,
                                                                                           (unsigned int)"AvailableSpare",
                                                                                           v50,
                                                                                           (unsigned int)&Size[1],
                                                                                           (__int64)&v70,
                                                                                           (__int64)Size,
                                                                                           (__int64)v72,
                                                                                           v7,
                                                                                           (__int64)&v69);
                                                                                if ( v68[0] == 1 && v69 )
                                                                                  *(_BYTE *)(a1 + 148) = v69;
                                                                                Size[0] = v70;
                                                                                NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                v69 = 0;
                                                                                LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                               a1,
                                                                                               (unsigned int)"AvailableSpareThreshold",
                                                                                               v51,
                                                                                               (unsigned int)&Size[1],
                                                                                               (__int64)&v70,
                                                                                               (__int64)Size,
                                                                                               (__int64)v72,
                                                                                               v7,
                                                                                               (__int64)&v69);
                                                                                v68[0] = v1;
                                                                                if ( (_BYTE)v1 == 1 )
                                                                                {
                                                                                  LOBYTE(v1) = v69;
                                                                                  if ( v69 )
                                                                                    *(_BYTE *)(a1 + 149) = v69;
                                                                                }
                                                                                if ( *(_QWORD *)&Size[1] )
                                                                                {
                                                                                  Size[0] = v70;
                                                                                  NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                  LOBYTE(v1) = ReadBinaryRegistryValue(
                                                                                                 a1,
                                                                                                 v52,
                                                                                                 v53,
                                                                                                 (unsigned int)&Size[1],
                                                                                                 (__int64)&v70,
                                                                                                 (__int64)Size);
                                                                                  v68[0] = v1;
                                                                                  if ( (_BYTE)v1 == 1 )
                                                                                  {
                                                                                    v54 = Size[0];
                                                                                    if ( Size[0] )
                                                                                    {
                                                                                      if ( Size[0] < 0x10u )
                                                                                        v37 = Size[0];
                                                                                      memmove(
                                                                                        (void *)(a1 + 4297),
                                                                                        *(const void **)&Size[1],
                                                                                        v37);
                                                                                      LOBYTE(v1) = *(_BYTE *)(a1 + 4296) | 1;
                                                                                      *(_BYTE *)(a1 + 4296) = v1;
                                                                                      if ( v54 > 8 )
                                                                                      {
                                                                                        LOBYTE(v1) = v1 | 2;
                                                                                        *(_BYTE *)(a1 + 4296) = v1;
                                                                                      }
                                                                                    }
                                                                                  }
                                                                                  if ( *(_QWORD *)&Size[1] )
                                                                                  {
                                                                                    Size[0] = v70;
                                                                                    NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                    v69 = 0;
                                                                                    LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                                   a1,
                                                                                                   (unsigned int)"ForcedPhysicalSectorSizeInBytes",
                                                                                                   v55,
                                                                                                   (unsigned int)&Size[1],
                                                                                                   (__int64)&v70,
                                                                                                   (__int64)Size,
                                                                                                   (__int64)v72,
                                                                                                   v7,
                                                                                                   (__int64)&v69);
                                                                                    v68[0] = v1;
                                                                                    if ( (_BYTE)v1 == 1 )
                                                                                    {
                                                                                      LOBYTE(v1) = v69;
                                                                                      if ( v69 )
                                                                                        *(_DWORD *)(a1 + 152) = v69;
                                                                                    }
                                                                                    if ( *(_QWORD *)&Size[1] )
                                                                                    {
                                                                                      Size[0] = v70;
                                                                                      NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                      v69 = 0;
                                                                                      LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                                                                                     a1,
                                                                                                     (unsigned int)"RetainAsyncEventControlMask",
                                                                                                     v56,
                                                                                                     (unsigned int)&Size[1],
                                                                                                     (__int64)&v70,
                                                                                                     (__int64)Size,
                                                                                                     (__int64)v72,
                                                                                                     v7,
                                                                                                     (__int64)&v69);
                                                                                      v68[0] = v1;
                                                                                      if ( (_BYTE)v1 == 1 )
                                                                                      {
                                                                                        LOBYTE(v1) = v69;
                                                                                        *(_DWORD *)(a1 + 156) = v69;
                                                                                      }
                                                                                      if ( *(_QWORD *)&Size[1] )
                                                                                      {
                                                                                        Size[0] = v70;
                                                                                        NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                        v69 = 0;
                                                                                        LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"ShutdownTimeoutForSurpriseRemove", v57, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                        v68[0] = v1;
                                                                                        if ( (_BYTE)v1 == 1 )
                                                                                        {
                                                                                          LOBYTE(v1) = v69;
                                                                                          if ( v69 )
                                                                                          {
                                                                                            if ( v69 > 0xFF )
                                                                                              LOBYTE(v1) = -1;
                                                                                            *(_BYTE *)(a1 + 150) = v1;
                                                                                          }
                                                                                        }
                                                                                        if ( *(_QWORD *)&Size[1] )
                                                                                        {
                                                                                          Size[0] = v70;
                                                                                          NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                          v69 = 0;
                                                                                          LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"MaxIoCountLimit", v58, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                          v68[0] = v1;
                                                                                          if ( (_BYTE)v1 )
                                                                                          {
                                                                                            LOBYTE(v1) = v69;
                                                                                            if ( v69 )
                                                                                            {
                                                                                              if ( v69 > 0x80 )
                                                                                                v43 = v69;
                                                                                              *(_DWORD *)(a1 + 160) = v43;
                                                                                            }
                                                                                          }
                                                                                          if ( *(_QWORD *)&Size[1] )
                                                                                          {
                                                                                            Size[0] = v70;
                                                                                            NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                            v69 = 0;
                                                                                            LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"SubmissionQueueAssignmentPolicy", v59, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                            v68[0] = v1;
                                                                                            if ( (_BYTE)v1 == 1 )
                                                                                            {
                                                                                              LOBYTE(v1) = v69;
                                                                                              if ( v69 )
                                                                                                *(_DWORD *)(a1 + 164) = v69;
                                                                                            }
                                                                                            if ( *(_QWORD *)&Size[1] )
                                                                                            {
                                                                                              Size[0] = v70;
                                                                                              NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                              v69 = 0;
                                                                                              LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"DisableMFNDCCDuringRemoval", v60, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                              v68[0] = v1;
                                                                                              if ( (_BYTE)v1 == 1 && v69 )
                                                                                                *(_DWORD *)(a1 + 64) |= 0x2000u;
                                                                                              if ( *(_QWORD *)&Size[1] )
                                                                                              {
                                                                                                Size[0] = v70;
                                                                                                NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                                v69 = 0;
                                                                                                LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"EnableSingleDpcForIoCompletion", v61, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                                v68[0] = v1;
                                                                                                if ( (_BYTE)v1 == 1 && v69 )
                                                                                                  *(_DWORD *)(a1 + 64) |= 0x4000u;
                                                                                                if ( *(_QWORD *)&Size[1] )
                                                                                                {
                                                                                                  Size[0] = v70;
                                                                                                  NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                                  v69 = 0;
                                                                                                  LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"DisableNamespacePreferredValueCheck", v62, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                                  v68[0] = v1;
                                                                                                  if ( (_BYTE)v1 == 1 && v69 )
                                                                                                    *(_DWORD *)(a1 + 64) |= 0x8000u;
                                                                                                  if ( *(_QWORD *)&Size[1] )
                                                                                                  {
                                                                                                    Size[0] = v70;
                                                                                                    NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                                    v69 = 0;
                                                                                                    LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"IgnoreNamespacePreferredValues", v63, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                                    v68[0] = v1;
                                                                                                    if ( (_BYTE)v1 == 1 && v69 )
                                                                                                      *(_DWORD *)(a1 + 64) |= 0x10000u;
                                                                                                    if ( *(_QWORD *)&Size[1] )
                                                                                                    {
                                                                                                      Size[0] = v70;
                                                                                                      NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                                      v69 = 0;
                                                                                                      LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"DisableBypassIO", v64, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                                      v68[0] = v1;
                                                                                                      if ( (_BYTE)v1 == 1 && v69 )
                                                                                                        *(_DWORD *)(a1 + 64) |= 0x20000u;
                                                                                                      if ( *(_QWORD *)&Size[1] )
                                                                                                      {
                                                                                                        Size[0] = v70;
                                                                                                        NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                                        v69 = 0;
                                                                                                        LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"DisableGetActiveNSIDList", v65, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                                        v68[0] = v1;
                                                                                                        if ( (_BYTE)v1 == 1 && v69 )
                                                                                                          *(_DWORD *)(a1 + 64) |= 0x40000u;
                                                                                                        if ( *(_QWORD *)&Size[1] )
                                                                                                        {
                                                                                                          Size[0] = v70;
                                                                                                          NVMeZeroMemory(*(_QWORD *)&Size[1]);
                                                                                                          v69 = 0;
                                                                                                          LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(a1, (unsigned int)"ForceCryptoEraseToUseFormatNVM", v66, (unsigned int)&Size[1], (__int64)&v70, (__int64)Size, (__int64)v72, v7, (__int64)&v69);
                                                                                                          v68[0] = v1;
                                                                                                          if ( (_BYTE)v1 == 1 && v69 )
                                                                                                            *(_DWORD *)(a1 + 64) |= 0x80000u;
                                                                                                          if ( *(_QWORD *)&Size[1] )
                                                                                                          {
                                                                                                            StorPortFreeRegistryBuffer(a1);
                                                                                                            LOBYTE(v1) = GetDynamicRegistrySettings(a1);
                                                                                                          }
                                                                                                        }
                                                                                                      }
                                                                                                    }
                                                                                                  }
                                                                                                }
                                                                                              }
                                                                                            }
                                                                                          }
                                                                                        }
                                                                                      }
                                                                                    }
                                                                                  }
                                                                                }
                                                                              }
                                                                            }
                                                                          }
                                                                        }
                                                                      }
                                                                    }
                                                                  }
                                                                }
                                                              }
                                                            }
                                                          }
                                                        }
                                                      }
                                                    }
                                                  }
                                                }
                                              }
                                            }
                                          }
                                        }
                                      }
                                    }
                                  }
                                  return v1;
                                }
                                Size[0] = v70;
                                NVMeZeroMemory(v21);
                                v69 = 0;
                                LOBYTE(v1) = ReadMultiSzRegistryValueAndCompareId(
                                               a1,
                                               (unsigned int)"IdlePowerMode",
                                               v22,
                                               (unsigned int)&Size[1],
                                               (__int64)&v70,
                                               (__int64)Size,
                                               (__int64)v72,
                                               v7,
                                               (__int64)&v69);
                                v68[0] = v1;
                                if ( (_BYTE)v1 == 1 )
                                {
                                  LOBYTE(v1) = v69;
                                  if ( v69 < 6 )
                                    *(_DWORD *)(a1 + 88) = v69;
                                }
                                v21 = *(_QWORD *)&Size[1];
                                if ( *(_QWORD *)&Size[1] )
                                  goto LABEL_71;
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  return v1;
}