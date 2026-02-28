__int64 __fastcall PopApplyPolicy(char a1, char a2, _OWORD *a3, unsigned int a4)
{
  __int64 v8; // r8
  __int128 v9; // xmm1
  __int128 v10; // xmm0
  __int128 v11; // xmm1
  __int128 v12; // xmm0
  __int128 v13; // xmm1
  __int128 v14; // xmm0
  __int128 v15; // xmm0
  _OWORD *v16; // rbx
  __int64 v17; // rax
  __int128 v18; // xmm0
  __int128 v19; // xmm1
  __int128 v20; // xmm0
  __int128 v21; // xmm1
  __int128 v22; // xmm0
  __int64 result; // rax
  int v24; // ebx
  _QWORD *v25; // rdi
  char v26; // r14
  __int64 i; // r8
  __int64 v28; // rcx
  _OWORD *v29; // rcx
  __int128 v30; // xmm1
  __int128 v31; // xmm0
  __int128 v32; // xmm1
  __int128 v33; // xmm0
  __int128 v34; // xmm1
  __int128 v35; // xmm0
  __int128 v36; // xmm1
  __int128 v37; // xmm0
  __int128 v38; // xmm1
  __int128 v39; // xmm0
  __int128 v40; // xmm1
  __int128 v41; // xmm0
  __int128 v42; // xmm1
  __int64 v43; // rax
  __int64 v44; // rcx
  HANDLE DestinationString; // [rsp+38h] [rbp-D0h] BYREF
  UNICODE_STRING DestinationString_8; // [rsp+40h] [rbp-C8h] BYREF
  _OWORD Buf1[6]; // [rsp+58h] [rbp-B0h] BYREF
  __int128 v48; // [rsp+B8h] [rbp-50h]
  _OWORD v49[7]; // [rsp+C8h] [rbp-40h]
  __int64 v50; // [rsp+138h] [rbp+30h]
  _OWORD Data[14]; // [rsp+148h] [rbp+40h] BYREF
  __int64 v52; // [rsp+228h] [rbp+120h]

  memset_0(Buf1, 0, 0xE8uLL);
  DestinationString = 0LL;
  DestinationString_8 = 0LL;
  if ( a4 < 0xE8 )
    return 3221225507LL;
  if ( a4 > 0xE8 )
    return 2147483653LL;
  v9 = a3[1];
  Data[0] = *a3;
  v10 = a3[2];
  Data[1] = v9;
  v11 = a3[3];
  Data[2] = v10;
  v12 = a3[4];
  Data[3] = v11;
  v13 = a3[5];
  Data[4] = v12;
  v14 = a3[6];
  Data[5] = v13;
  Data[6] = v14;
  v15 = a3[7];
  v16 = a3 + 8;
  Data[7] = v15;
  v17 = *((_QWORD *)v16 + 12);
  v18 = v16[1];
  Data[8] = *v16;
  v19 = v16[2];
  Data[9] = v18;
  v20 = v16[3];
  Data[10] = v19;
  v21 = v16[4];
  Data[11] = v20;
  v22 = v16[5];
  Data[12] = v21;
  Data[13] = v22;
  v52 = v17;
  result = PopVerifySystemPowerPolicy(Data, Buf1, v8);
  v24 = result;
  if ( (int)result >= 0 )
  {
    v25 = PopPolicy;
    if ( !memcmp(Buf1, PopPolicy, 0xE8uLL) && !a1 )
    {
      return 0LL;
    }
    else
    {
      v26 = 0;
      for ( i = 0LL; (unsigned int)i < 4; i = (unsigned int)(i + 1) )
      {
        v28 = *((_QWORD *)&v49[-1] + 3 * i) - v25[3 * i + 12];
        if ( !v28 )
        {
          v28 = *((_QWORD *)&v48 + 3 * i + 1) - v25[3 * i + 13];
          if ( !v28 )
            v28 = *((_QWORD *)v49 + 3 * i) - v25[3 * i + 14];
        }
        if ( v28 )
        {
          v26 = 1;
          break;
        }
      }
      v29 = PopPolicy;
      v30 = Buf1[1];
      *(_OWORD *)PopPolicy = Buf1[0];
      v31 = Buf1[2];
      v29[1] = v30;
      v32 = Buf1[3];
      v29[2] = v31;
      v33 = Buf1[4];
      v29[3] = v32;
      v34 = Buf1[5];
      v29[4] = v33;
      v35 = v48;
      v29[5] = v34;
      v36 = v49[0];
      v29[6] = v35;
      v29 += 8;
      v37 = v49[1];
      *(v29 - 1) = v36;
      v38 = v49[2];
      *v29 = v37;
      v39 = v49[3];
      v29[1] = v38;
      v40 = v49[4];
      v29[2] = v39;
      v41 = v49[5];
      v29[3] = v40;
      v42 = v49[6];
      v43 = v50;
      v29[4] = v41;
      v29[5] = v42;
      *((_QWORD *)v29 + 12) = v43;
      PopSetNotificationWork(2LL);
      if ( v26 && !a2 )
      {
        LOBYTE(v44) = -125;
        PopResetCBTriggers(v44);
      }
      PopUpdateSystemIdleContext(3LL);
      if ( a1 )
      {
        v24 = PopOpenPowerKey((__int64)&DestinationString);
        if ( v24 >= 0 )
        {
          RtlInitUnicodeString(&DestinationString_8, L"SystemPowerPolicy");
          v24 = ZwSetValueKey(DestinationString, &DestinationString_8, 0, 3u, Data, 0xE8u);
          ZwClose(DestinationString);
        }
      }
      return (unsigned int)v24;
    }
  }
  return result;
}
