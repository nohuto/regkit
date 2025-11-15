NTSTATUS RtlpHpApplySegmentHeapConfigurations()
{
  NTSTATUS result; // eax
  int v1; // [rsp+30h] [rbp-9h] BYREF
  HANDLE Handle; // [rsp+38h] [rbp-1h] BYREF
  _QWORD v3[3]; // [rsp+40h] [rbp+7h] BYREF
  int v4; // [rsp+58h] [rbp+1Fh]
  int v5; // [rsp+5Ch] [rbp+23h]
  __int128 v6; // [rsp+60h] [rbp+27h]
  __int128 v7; // [rsp+70h] [rbp+37h] BYREF
  int v8; // [rsp+80h] [rbp+47h]

  v5 = 0;
  v1 = 0;
  v3[0] = 48LL;
  Handle = 0LL;
  v3[1] = 0LL;
  v8 = 0;
  v3[2] = &unk_180171818;
  v7 = 0LL;
  v4 = 64;
  v6 = 0LL;
  result = NtOpenKey(&Handle, 1LL, v3);
  if ( result >= 0 )
  {
    result = NtQueryValueKey(Handle, &unk_180172A40, 2LL, &v7, 20, &v1);
    if ( result >= 0 && DWORD2(v7) == 4 )
    {
      if ( HIDWORD(v7) )
      {
        result = RtlpLowFragHeapGlobalFlags | 0x10;
        RtlpLowFragHeapGlobalFlags |= 0x10u;
        if ( (BYTE12(v7) & 2) != 0 )
        {
          result |= 0x20u;
          RtlpLowFragHeapGlobalFlags = result;
        }
      }
      else
      {
        RtlpLowFragHeapGlobalFlags |= 8u;
      }
    }
  }
  if ( Handle )
    return NtClose(Handle);
  return result;
}