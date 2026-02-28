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
