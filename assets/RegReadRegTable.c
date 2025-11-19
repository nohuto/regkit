void __fastcall REGISTRY::RegReadRegTable(
        REGISTRY *this,
        struct ADAPTER_CONTEXT *a2,
        void *a3,
        struct REGTABLE_ENTRY *a4,
        unsigned int a5)
{
  __int64 v5; // r12
  struct REGTABLE_ENTRY *v6; // r10
  char *v8; // r14
  ULONG IntegerData; // esi
  ULONG v10; // ebp
  int v11; // edi
  char *v12; // rbx
  char *v13; // rax
  _BYTE *v14; // rcx
  int v15; // edi
  int v16; // ebx
  char *v17; // rax
  int v18; // edi
  char *v19; // rbx
  char *v20; // rax
  int LockArray_high; // edi
  char *LevelString; // rbx
  char *v23; // rax
  int v24; // r9d
  int v25; // edi
  char *v26; // rbx
  char *v27; // rax
  int Status; // [rsp+60h] [rbp-38h] BYREF
  PNDIS_CONFIGURATION_PARAMETER ParameterValue; // [rsp+68h] [rbp-30h] BYREF
  void *v30; // [rsp+B0h] [rbp+18h]

  v30 = a3;
  ParameterValue = 0LL;
  v5 = 0LL;
  v6 = a4;
  Status = -1073741823;
  if ( a5 )
  {
    v8 = (char *)a4 + 32;
    do
    {
      if ( a3 )
        NdisReadConfiguration_0(&Status, &ParameterValue, a3, (PNDIS_STRING)v6 + 3 * v5, NdisParameterInteger);
      if ( Status || !ParameterValue )
      {
        IntegerData = *((_DWORD *)v8 + 2);
        if ( !v8[13] || WPP_RECORDER_INITIALIZED == (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          goto LABEL_12;
        LockArray_high = HIDWORD(KeGetPcr()[1].LockArray);
        LevelString = WppGetLevelString(4u);
        v23 = TrcHdr(a2);
        v24 = 23;
        goto LABEL_29;
      }
      IntegerData = ParameterValue->ParameterData.IntegerData;
      v10 = *(_DWORD *)v8;
      if ( v8[12] )
      {
        if ( IntegerData < v10 )
        {
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          {
            v11 = HIDWORD(KeGetPcr()[1].LockArray);
            v12 = WppGetLevelString(4u);
            v13 = TrcHdr(a2);
            WPP_RECORDER_SF_DssSDD(
              WPP_GLOBAL_Control->DeviceExtension,
              0,
              2,
              18,
              (__int64)&WPP_16994a84f63c31d8ae8ff2b658e92d93_Traceguids,
              v11,
              (__int64)v13,
              (__int64)v12,
              *((_QWORD *)v8 - 3),
              IntegerData,
              v10);
            v10 = *(_DWORD *)v8;
          }
LABEL_11:
          IntegerData = v10;
          goto LABEL_12;
        }
        v10 = *((_DWORD *)v8 + 1);
        if ( IntegerData > v10 )
        {
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          {
            v18 = HIDWORD(KeGetPcr()[1].LockArray);
            v19 = WppGetLevelString(4u);
            v20 = TrcHdr(a2);
            WPP_RECORDER_SF_DssSDD(
              WPP_GLOBAL_Control->DeviceExtension,
              0,
              2,
              19,
              (__int64)&WPP_16994a84f63c31d8ae8ff2b658e92d93_Traceguids,
              v18,
              (__int64)v20,
              (__int64)v19,
              *((_QWORD *)v8 - 3),
              IntegerData,
              v10);
            v10 = *((_DWORD *)v8 + 1);
          }
          goto LABEL_11;
        }
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          LockArray_high = HIDWORD(KeGetPcr()[1].LockArray);
          LevelString = WppGetLevelString(4u);
          v23 = TrcHdr(a2);
          v24 = 20;
LABEL_29:
          WPP_RECORDER_SF_DssSD(
            WPP_GLOBAL_Control->DeviceExtension,
            0,
            2,
            v24,
            (__int64)&WPP_16994a84f63c31d8ae8ff2b658e92d93_Traceguids,
            LockArray_high,
            (__int64)v23,
            (__int64)LevelString,
            *((_QWORD *)v8 - 3),
            IntegerData);
        }
      }
      else
      {
        if ( IntegerData < v10 || IntegerData > *((_DWORD *)v8 + 1) )
        {
          if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
          {
            v25 = HIDWORD(KeGetPcr()[1].LockArray);
            v26 = WppGetLevelString(4u);
            v27 = TrcHdr(a2);
            WPP_RECORDER_SF_DssSDD(
              WPP_GLOBAL_Control->DeviceExtension,
              0,
              2,
              22,
              (__int64)&WPP_16994a84f63c31d8ae8ff2b658e92d93_Traceguids,
              v25,
              (__int64)v27,
              (__int64)v26,
              *((_QWORD *)v8 - 3),
              IntegerData,
              *((_DWORD *)v8 + 2));
          }
          IntegerData = *((_DWORD *)v8 + 2);
          goto LABEL_12;
        }
        if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
        {
          LockArray_high = HIDWORD(KeGetPcr()[1].LockArray);
          LevelString = WppGetLevelString(4u);
          v23 = TrcHdr(a2);
          v24 = 21;
          goto LABEL_29;
        }
      }
LABEL_12:
      v14 = (_BYTE *)*((_QWORD *)v8 - 2);
      if ( !v14 )
        v14 = (char *)a2 + *((unsigned int *)v8 - 2);
      v15 = *((_DWORD *)v8 - 1);
      if ( v15 == 1 )
      {
        *v14 = IntegerData;
      }
      else
      {
        switch ( *((_DWORD *)v8 - 1) )
        {
          case 2:
            *(_WORD *)v14 = IntegerData;
            break;
          case 4:
            *(_DWORD *)v14 = IntegerData;
            break;
          case 8:
            *(_QWORD *)v14 = IntegerData;
            break;
          default:
            if ( WPP_RECORDER_INITIALIZED != (_UNKNOWN *)&WPP_RECORDER_INITIALIZED )
            {
              v16 = HIDWORD(KeGetPcr()[1].LockArray);
              v17 = TrcHdr(a2);
              WPP_RECORDER_SF_DsL(
                WPP_GLOBAL_Control->DeviceExtension,
                0,
                2,
                24,
                (__int64)&WPP_16994a84f63c31d8ae8ff2b658e92d93_Traceguids,
                v16,
                (__int64)v17,
                v15);
            }
            break;
        }
      }
      v6 = a4;
      v5 = (unsigned int)(v5 + 1);
      a3 = v30;
      v8 += 48;
    }
    while ( (unsigned int)v5 < a5 );
  }
}