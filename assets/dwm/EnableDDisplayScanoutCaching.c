bool dynamic_initializer_for__CCommonRegistryData::EnableDDisplayScanoutCaching__()
{
  bool result; // al
  int v1; // [rsp+50h] [rbp+8h] BYREF

  v1 = 0;
  if ( (unsigned int)GetPersistedRegistryValueW(
                       L"DWMSwitches",
                       L"Software\\Microsoft\\Windows\\Dwm",
                       L"EnableDDisplayScanoutCaching",
                       16LL,
                       0LL,
                       &v1,
                       4,
                       0LL) )
    result = 1;
  else
    result = v1 != 0;
  CCommonRegistryData::EnableDDisplayScanoutCaching = result;
  return result;
}
