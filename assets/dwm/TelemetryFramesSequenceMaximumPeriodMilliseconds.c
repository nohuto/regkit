__int64 dynamic_initializer_for__CCommonRegistryData::TelemetryFramesSequenceMaximumPeriodMilliseconds__()
{
  __int64 result; // rax
  int v1; // ecx
  int v2; // [rsp+50h] [rbp+8h] BYREF

  v2 = 0;
  result = GetPersistedRegistryValueW(
             L"DWMSwitches",
             L"Software\\Microsoft\\Windows\\Dwm",
             L"TelemetryFramesSequenceMaximumPeriodMilliseconds",
             16LL,
             0LL,
             &v2,
             4,
             0LL);
  v1 = 1000;
  if ( !(_DWORD)result )
    v1 = v2;
  CCommonRegistryData::TelemetryFramesSequenceMaximumPeriodMilliseconds = v1;
  return result;
}
