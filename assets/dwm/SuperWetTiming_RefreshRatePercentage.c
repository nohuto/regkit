__int64 dynamic_initializer_for__CCommonRegistryData::SuperWetTiming::RefreshRatePercentage__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"SuperWetTiming",
             L"Software\\Microsoft\\Windows\\Dwm\\GpuAccelInkTiming",
             L"RefreshRatePercentage",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 10;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::SuperWetTiming::RefreshRatePercentage = v1;
  return result;
}
