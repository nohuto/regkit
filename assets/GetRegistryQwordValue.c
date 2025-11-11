__int64 __fastcall GetRegistryQwordValue(__int64 a1, __int64 a2, __int64 *a3)
{
  unsigned int v3; // ebx
  __int64 result; // rax
  _QWORD v6[16]; // [rsp+38h] [rbp-29h] BYREF
  __int64 v7; // [rsp+C8h] [rbp+67h] BYREF
  unsigned int v8; // [rsp+D0h] [rbp+6Fh] BYREF
  int v9; // [rsp+D4h] [rbp+73h]

  v9 = HIDWORD(a2);
  v3 = 0;
  v8 = 0;
  v7 = 4294967288LL;
  memset(v6, 0, 0x70uLL);
  LODWORD(v6[1]) = 292;
  v6[3] = &v7;
  v6[2] = L"Capabilities";
  LODWORD(v6[4]) = 184549376;
  if ( (int)RtlQueryRegistryValuesEx(
              0LL,
              L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Processor",
              v6,
              0LL,
              0LL) < 0 )
  {
    if ( (int)GetRegistryDwordValueNoDefault(
                L"\\Registry\\Machine\\System\\CurrentControlSet\\Control\\Processor",
                L"Capabilities",
                &v8) >= 0 )
      v3 = v8;
    result = v3;
  }
  else
  {
    result = v7;
  }
  *a3 = result;
  return result;
}