__int64 __fastcall GetDynamicRegistrySettings(__int64 a1)
{
  __int64 result; // rax
  unsigned int v3; // r11d
  int v4; // r11d
  int v5; // r8d
  void *v6; // r10
  unsigned int v7; // ecx
  __int128 *v8; // rax
  int v9; // edi
  int v10; // r8d
  unsigned int Size; // [rsp+50h] [rbp+17h] BYREF
  unsigned int Size_4; // [rsp+54h] [rbp+1Bh] BYREF
  unsigned int v13; // [rsp+58h] [rbp+1Fh] BYREF
  void *v14; // [rsp+60h] [rbp+27h] BYREF
  __int128 v15; // [rsp+68h] [rbp+2Fh] BYREF
  __int64 v16; // [rsp+78h] [rbp+3Fh] BYREF
  char v17; // [rsp+80h] [rbp+47h]

  v17 = aVenVvvvDevDddd[24];
  v15 = *(_OWORD *)"VEN_vvvv&DEV_dddd&REV_rr";
  Size = 512;
  *(_DWORD *)(a1 + 56) &= ~0x200000u;
  v16 = *(_QWORD *)"d&REV_rr";
  result = StorPortAllocateRegistryBuffer(a1, &Size);
  v14 = (void *)result;
  if ( result )
  {
    UlongToHex((char *)&v15 + 4, *(unsigned __int16 *)(a1 + 4), 4LL);
    UlongToHex((char *)&v15 + 13, *(unsigned __int16 *)(a1 + 6), v3);
    UlongToHex((char *)&v16 + 6, *(unsigned __int8 *)(a1 + 8), (unsigned int)(v4 - 2));
    v7 = 0;
    v8 = &v15;
    v9 = 29;
    while ( *(_BYTE *)v8 )
    {
      ++v7;
      v8 = (__int128 *)((char *)v8 + 1);
      if ( v7 >= 0x1D )
        goto LABEL_7;
    }
    v9 = v7;
LABEL_7:
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
    result = ReadMultiSzRegistryValueAndCompareId(
               a1,
               (unsigned int)"ControllerResetWaitTimeCushion",
               v5,
               (unsigned int)&v14,
               (__int64)&Size,
               (__int64)&v13,
               (__int64)&v15,
               v9,
               (__int64)&Size_4);
    if ( (_BYTE)result == 1 )
    {
      result = Size_4;
      *(_DWORD *)(a1 + 156) = Size_4;
    }
    if ( v14 )
    {
      v13 = Size;
      if ( (Size & 3) != 0 )
      {
        if ( Size )
          memset(v14, 0, Size);
      }
      else if ( Size >> 2 )
      {
        memset(v14, 0, 4LL * (Size >> 2));
      }
      Size_4 = 0;
      result = ReadMultiSzRegistryValueAndCompareId(
                 a1,
                 (unsigned int)"DisableDSTThrottle",
                 v10,
                 (unsigned int)&v14,
                 (__int64)&Size,
                 (__int64)&v13,
                 (__int64)&v15,
                 v9,
                 (__int64)&Size_4);
      if ( (_BYTE)result == 1 && Size_4 )
        *(_DWORD *)(a1 + 56) |= 0x200000u;
      if ( v14 )
        return StorPortFreeRegistryBuffer(a1);
    }
  }
  return result;
}