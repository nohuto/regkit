__int64 __fastcall GetRegistrySettingsForSpecificKey(__int64 a1)
{
  bool v1; // zf
  unsigned int v3; // r11d
  int v4; // r11d
  int v5; // r8d
  void *v6; // r10
  unsigned int v7; // ecx
  char *v8; // rax
  int v9; // edi
  unsigned int Size; // [rsp+50h] [rbp+17h] BYREF
  unsigned int Size_4; // [rsp+54h] [rbp+1Bh] BYREF
  unsigned int v13; // [rsp+58h] [rbp+1Fh] BYREF
  __int64 v14; // [rsp+60h] [rbp+27h] BYREF
  char v15[32]; // [rsp+68h] [rbp+2Fh] BYREF

  v1 = *(_BYTE *)(a1 + 20) == 0;
  strcpy(v15, "VEN_vvvv&DEV_dddd&REV_rr");
  Size = 512;
  if ( !v1 )
    return 0LL;
  v14 = StorPortAllocateRegistryBuffer(a1, &Size);
  if ( !v14 )
    return 0LL;
  UlongToHex(&v15[4], *(unsigned __int16 *)(a1 + 4), 4LL);
  UlongToHex(&v15[13], *(unsigned __int16 *)(a1 + 6), v3);
  UlongToHex(&v15[22], *(unsigned __int8 *)(a1 + 8), (unsigned int)(v4 - 2));
  v7 = 0;
  v8 = v15;
  v9 = 29;
  while ( *v8 )
  {
    ++v7;
    ++v8;
    if ( v7 >= 0x1D )
      goto LABEL_8;
  }
  v9 = v7;
LABEL_8:
  v13 = Size;
  if ( (Size & 3) != 0 )
  {
    if ( Size )
      memset(v6, 0, Size);
  }
  else if ( Size >> 2 )
  {
    memset(v6, 0, 4LL * (Size >> 2));
  }
  Size_4 = 0;
  ReadMultiSzRegistryValueAndCompareId(
    a1,
    (unsigned int)"DisableActivateFWWithoutReset",
    v5,
    (unsigned int)&v14,
    (__int64)&Size,
    (__int64)&v13,
    (__int64)v15,
    v9,
    (__int64)&Size_4);
  if ( v14 )
    StorPortFreeRegistryBuffer(a1);
  return Size_4;
}