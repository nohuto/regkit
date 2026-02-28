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
