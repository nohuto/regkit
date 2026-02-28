__int64 __fastcall VidSchiReadNodeConfiguration(__int64 a1, __int64 a2, __int64 a3)
{
  __int64 v4; // rcx
  NTSTATUS v6; // ebx
  __int64 v7; // rdx
  const wchar_t *v8; // rcx
  NTSTATUS v9; // eax
  __int64 v11; // rax
  _DWORD *v12; // rdi
  unsigned __int64 v13; // r8
  unsigned int v14; // edx
  __int64 v15; // rcx
  struct _UNICODE_STRING ValueName; // [rsp+30h] [rbp-10h] BYREF
  ULONG ResultLength; // [rsp+70h] [rbp+30h] BYREF
  HANDLE KeyHandle; // [rsp+80h] [rbp+40h] BYREF

  v4 = *(_QWORD *)(a1 + 16);
  KeyHandle = 0LL;
  ResultLength = 0;
  ValueName = 0LL;
  v6 = DpiOpenPnpRegistryKey(*(_QWORD *)(v4 + 216), a2, a3, &KeyHandle);
  if ( v6 >= 0 )
  {
    v7 = 0x7FFFLL;
    v8 = L"HwQueuedRenderPacketGroupLimitPerNode";
    while ( *v8 )
    {
      ++v8;
      if ( !--v7 )
        goto LABEL_7;
    }
    ValueName.Buffer = L"HwQueuedRenderPacketGroupLimitPerNode";
    ValueName.Length = 2 * (0x7FFF - v7);
    ValueName.MaximumLength = ValueName.Length + 2;
LABEL_7:
    v9 = ZwQueryValueKey(KeyHandle, &ValueName, KeyValuePartialInformation, 0LL, 0, &ResultLength);
    if ( v9 == -2147483643 || v9 == -1073741789 )
    {
      v11 = 4LL * ResultLength;
      if ( !is_mul_ok(ResultLength, 4uLL) )
        v11 = -1LL;
      v12 = (_DWORD *)operator new[](v11, 828467542LL, 256LL);
      if ( v12 )
      {
        v6 = ZwQueryValueKey(KeyHandle, &ValueName, KeyValuePartialInformation, v12, ResultLength, &ResultLength);
        if ( v6 >= 0 )
        {
          if ( v12[1] != 3
            || (v13 = (unsigned int)v12[2], (v13 & 3) != 0)
            || v13 > 4 * (unsigned __int64)*(unsigned int *)(a1 + 80) )
          {
            v6 = -1073741811;
          }
          else
          {
            v14 = 0;
            if ( (v13 & 0xFFFFFFFC) != 0 )
            {
              do
              {
                v15 = v14++;
                *(_DWORD *)(a2 + 4 * v15) = _byteswap_ulong(v12[v15 + 3]);
              }
              while ( v14 < v12[2] >> 2 );
            }
          }
        }
        operator delete(v12);
      }
      else
      {
        v6 = -1073741801;
      }
    }
    else
    {
      v6 = -1073741275;
    }
  }
  if ( KeyHandle )
    ZwClose(KeyHandle);
  return (unsigned int)v6;
}